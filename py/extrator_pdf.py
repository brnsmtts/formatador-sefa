"""Importação conservadora de atos em PDF com tabelas vetoriais editáveis.

A extração nunca calcula valores. Cada número é lido de suas coordenadas e
comparado com o conjunto de números presente nas colunas da tabela do PDF.
PDF digitalizado ou grade não reconhecida requer revisão manual.
"""

from __future__ import annotations

import difflib
import html
import io
import os
import re
import statistics
import tempfile
import unicodedata
from collections import Counter
from copy import deepcopy
from dataclasses import dataclass, field
from pathlib import Path
from types import SimpleNamespace
from typing import Any

import pdfplumber
import fitz
from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Emu, Inches, Pt

from formatador_sefa import (
    ANNEX, ARTICLE, CARGO, CHANGE_MARK, COMMAND, EMENTA_START, ENDS_WITH_ABBREVIATION, EPIGRAPH, HEADING, NAME, NUMBER_SIGN, PREAMBLE,
    SIGNER_CARGO, Item, continues_epigraph, join_line, looks_like_signer_name, looks_uppercase, vocabulary,
    add_formatted_paragraph, apply_styles, classify, clear_model_contents,
    ementa_spacer_before, looks_like_person_name, paragraph_style_names,
)


MONEY = re.compile(r"(?<!\d)-?(?:\d{1,3}(?:\.\d{3})+|\d+),\d{2}(?![\d%])")
PREFIXED_CURRENCY = re.compile(r"R\$\s*(-?(?:\d{1,3}(?:\.\d{3})+|\d+)(?:,\d{2})?)")
PERCENT = re.compile(r"-?\d+(?:\.\d{3})*,\d+(?:\s*%)")
NUMERIC_CELL = re.compile(r"\s*(?:R\$\s*)?-?\d[\d.,]*(?:\s*%)?\s*")
ROW_START = re.compile(r"^(?:\d+(?:\.\d+)*\s*(?:[-–]|\.(?!\d))|[IVXLCDM]+\s*[-–]|TOTAIS\b)", re.I)
FOOTNOTE = re.compile(r"^[¹²³⁴⁵⁶⁷⁸⁹]")
PROTOCOL = re.compile(r"^PROTOCOLO\s*:\s*\d+", re.I)
SECTION_HEAD = re.compile(r"^ATOS\s+(?:DO|DA|DOS|DAS)\b", re.I)
ANNEX_HEAD = re.compile(r"^(?:GOVERNO DO ESTADO|DEMONSTRATIVO\b|ANEXO\b|R\$\s*1,00)", re.I)
PDF_DATE = re.compile(r"(\d{4})[._-](\d{2})[._-](\d{2})")
PORTUGUESE_DATE = re.compile(
    r"\b(\d{1,2})\s+de\s+(janeiro|fevereiro|março|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)\s+de\s+(\d{4})\b",
    re.I,
)
MONTH_NUMBERS = {name: f"{index:02d}" for index, name in enumerate(
    ("janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto",
     "setembro", "outubro", "novembro", "dezembro"), 1)}


@dataclass
class ParagraphBlock:
    text: str
    role: str = "texto"
    page: int = 0
    kind: str = "norma"  # norma, annex_head, unit, footnote, date, signer
    item: Item | None = None
    attention: list[str] = field(default_factory=list)
    # Trechos do diário com marcas (1 = negrito, 2 = itálico). Só nos atos cuja formatação é copiada do
    # PDF (Solução de Consulta); nos demais fica None e valem as regras de negrito dos exemplos SEFA.
    source_runs: list[tuple[str, int]] | None = None


@dataclass
class ImageBlock:
    data: bytes
    page: int
    kind: str = "image"
    width: float = 0.0        # em pontos; 0 = brasão/logotipo (largura fixa)


@dataclass
class TableBlock:
    matrix: list[list[str]]
    column_widths: list[float]
    page: int
    bbox: tuple[float, float, float, float]
    source_numbers: Counter
    source_signature: Counter = field(default_factory=Counter)
    original_matrix: list[list[str]] = field(default_factory=list)
    strategy: str = "linhas"
    merges: list[tuple[int, int, int, int]] = field(default_factory=list)
    issues: list[str] = field(default_factory=list)
    reviewed: bool = False
    kind: str = "table"
    source_parts: list[TablePagePart] = field(default_factory=list)


@dataclass
class TablePagePart:
    page: int
    bbox: tuple[float, float, float, float]
    first_row: int
    last_row: int  # exclusivo
    column: str = ""  # esquerda, direita ou vazio em páginas sem colunas


def table_page_parts(table: TableBlock) -> list[TablePagePart]:
    return table.source_parts or [TablePagePart(table.page, table.bbox, 0, len(table.matrix))]


@dataclass
class Extraction:
    pdf_path: Path
    query: str
    blocks: list[ParagraphBlock | ImageBlock | TableBlock]
    pages: list[int]
    issues: list[str] = field(default_factory=list)
    blocking_issues: list[str] = field(default_factory=list)
    limits_reviewed: bool = False
    protocol: str = ""
    intertable_segments_checked: int = 0
    # Citação aberta no fim de uma coluna ou página continua na região seguinte (3.24).
    quote_state: dict = field(default_factory=dict)
    source_url: str = ""          # importação pelo link do DOU (sem PDF)

    @property
    def tables(self) -> list[TableBlock]:
        return [b for b in self.blocks if isinstance(b, TableBlock)]

    @property
    def paragraphs(self) -> list[ParagraphBlock]:
        return [b for b in self.blocks if isinstance(b, ParagraphBlock)]


def normalized(text: str) -> str:
    text = unicodedata.normalize("NFKD", text).upper()
    return re.sub(r"\s+", " ", "".join(c for c in text if not unicodedata.combining(c))).strip()


def character_signature(text: str) -> Counter:
    """Compara todo o texto da tabela, independentemente das quebras e espaços."""
    return Counter(c for c in unicodedata.normalize("NFKC", text) if not c.isspace())


def matrix_signature(matrix: list[list[str]]) -> Counter:
    return character_signature("".join(value for row in matrix for value in row))


def monetary_values(matrix: list[list[str]]) -> Counter:
    numbers: Counter = Counter()
    for row in matrix:
        for value in row:
            for line in value.splitlines():
                if PERCENT.fullmatch(line.strip()):
                    continue
                prefixed = list(PREFIXED_CURRENCY.finditer(line))
                numbers.update(match.group(1) for match in prefixed)
                numbers.update(match.group() for match in MONEY.finditer(line)
                               if not any(m.start() <= match.start() < m.end() for m in prefixed))
    return numbers


def span_marks(span: dict) -> int:
    """1 = negrito, 2 = itálico, pelo indicador do PDF ou pelo nome da fonte ("Times-BoldItalic")."""
    font, flags = span.get("font", ""), span.get("flags", 0)
    bold = bool(flags & 16) or bool(re.search(r"Bold|Black|Heavy|Semibold|Demi", font, re.I))
    italic = bool(flags & 2) or bool(re.search(r"Italic|Oblique", font, re.I))
    return (1 if bold else 0) | (2 if italic else 0)


def pdf_lines(pdf_path: Path, page_number: int, document: Any | None = None) -> list[dict]:
    """Lê linhas e coordenadas, preservando a caixa das letras sem programas externos. Cada linha
    leva também os trechos com negrito/itálico ("spans"), usados só quando a formatação é copiada."""
    if document is None:
        with fitz.open(str(pdf_path)) as pdf:
            return pdf_lines(pdf_path, page_number, pdf)
    page = document[page_number - 1]
    raw = page.get_text("rawdict")["blocks"]
    widths = _glyph_widths(raw)
    lines = []
    for block in raw:
        if block["type"] != 0:
            continue
        for line in block["lines"]:
            spans = [(_span_text(span, widths), span) for span in line["spans"]
                     if span["size"] >= HIDDEN_SIZE and not _white_dots(span)]
            value = re.sub(r"\s+", " ", "".join(text for text, _ in spans)).strip()
            if value:
                x0, top, x1, bottom = line["bbox"]
                lines.append({"text": value, "x0": x0, "top": top,
                              "x1": x1, "bottom": bottom,
                              "spans": [(text, span_marks(span)) for text, span in spans if text]})
    return sorted(lines, key=lambda l: (l["top"], l["x0"]))


# Texto invisível: no DOE de 21.10.2026 há um "GABINETE DO GOVERNADOR" branco, em corpo 0,1, que entrava no
# começo da Lei nº 11.672. Nada abaixo de 1,5 pt é legível.
HIDDEN_SIZE = 1.5


def _white_dots(span: dict) -> bool:
    """Pontos impressos em branco, invisíveis na página: no DOU de 26.10.2026 cercam cada célula das tabelas do
    Ato Técnico Conjunto RFB/CGIBS nº 9 (".José Guilherme ... .", ".DIREX"). Texto branco com letras fica (títulos
    brancos sobre faixa preta, como no DOE)."""
    text = "".join(char["c"] for char in span.get("chars", ())) or span.get("text", "")
    return span.get("color") == 0xFFFFFF and bool(text.strip()) and not text.strip(". ")


def _invisible_char(obj: dict) -> bool:
    """O mesmo para as letras lidas pelo pdfplumber (células das tabelas)."""
    if obj.get("object_type") != "char":
        return False
    if obj.get("size", 9) < HIDDEN_SIZE:
        return True
    color = obj.get("non_stroking_color")
    values = tuple(color) if isinstance(color, (list, tuple)) else (color,)
    white = values in ((1,), (1.0,), (1, 1, 1), (1.0, 1.0, 1.0), (0, 0, 0, 0), (0.0, 0.0, 0.0, 0.0))
    return white and obj.get("text") == "."


def _phantom_space(previous: Any, char: Any, following: Any,
                   box: Any = lambda c: (c["x0"], c["top"], c["x1"])) -> bool:
    """Espaço sem largura visível: a letra seguinte começa onde a anterior termina, por cima do espaço. Nas
    tabelas do Anexo I do Ato Técnico Conjunto RFB/CGIBS nº 9 (DOU de 26.10.2026, p. 36-37), o PDF guarda
    "R E P R ES E N T AÇ ÃO", "Ko u r y" e "Fo c a l", mas a página mostra "REPRESENTAÇÃO", "Koury" e "Focal"; o
    mesmo diário tem "D EC I S Õ ES" e "Fe d e r a l". Espaço estreito mas visível (texto justificado do DOE, a
    partir de 0,4 pt entre as letras) continua."""
    if previous is None or following is None:
        return False
    x0, top, x1 = box(char)
    _, previous_top, previous_x1 = box(previous)
    next_x0, next_top, _ = box(following)
    return (abs(top - next_top) < .5 and abs(top - previous_top) < .5
            and x0 - 1 <= next_x0 < x0 + .5 * (x1 - x0) and next_x0 - previous_x1 < .2)


def visible_page(page: Any) -> Any:
    chars = page.chars
    phantom = {id(char) for previous, char, following in zip(chars, chars[1:], chars[2:])
               if char["text"] == " " and previous["text"] != " " and following["text"] != " "
               and _phantom_space(previous, char, following)}
    return page.filter(lambda obj: id(obj) not in phantom and not _invisible_char(obj))


def open_top_settings(page: Any) -> dict:
    """Borda de cima que falta numa tabela: a continuação da tabela do GTI 7 no alto da coluna direita (DOU de
    26.10.2026, p. 36) começa sem linha horizontal, e a primeira linha ("José Guilherme de Souza Moitta Koury |
    | DIREX") ficava fora. Quando duas ou mais verticais começam na mesma altura, sem horizontal ali, e há uma
    horizontal logo abaixo cobrindo as mesmas verticais, acrescenta-se a borda que falta."""
    edges = page.edges
    horizontals = [e for e in edges if e["orientation"] == "h"]
    groups: dict[int, list[dict]] = {}
    for edge in edges:
        if edge["orientation"] == "v":
            groups.setdefault(round(edge["top"]), []).append(edge)
    lines = []
    for top, verticals in groups.items():
        if len(verticals) < 2:
            continue
        left, right = min(v["x0"] for v in verticals), max(v["x0"] for v in verticals)
        if right - left < 50:
            continue
        def covers(h: dict) -> bool:
            return h["x0"] <= left + 2 and h["x1"] >= right - 2
        # já há horizontal nessa altura, ainda que só em parte (o "Ad valorem (%)" que cobre "Inicial | Fator de
        # conversão | Final" no cabeçalho do Anexo IV da MP nº 1.399): é divisão do cabeçalho, não borda que falta
        if any(abs(h["top"] - top) <= 2 and h["x1"] > left + 2 and h["x0"] < right - 2 for h in horizontals):
            continue
        # cada vertical desce até uma horizontal da tabela (não são traços do logotipo do cabeçalho do DOU); a
        # primeira linha pode ser alta: "Bebidas alcoólicas das seguintes posições ..." tem 70 pt (Anexo IV da
        # MP nº 1.399, DOU de 04.11.2026, edição extra B, p. 3)
        below = [h for h in horizontals if 0 < h["top"] - top <= 200 and covers(h)]
        if not below or not all(v["bottom"] - v["top"] >= 6 and any(abs(v["bottom"] - h["top"]) <= 2 for h in below)
                                for v in verticals):
            continue
        top = min(v["top"] for v in verticals)
        lines.append({"object_type": "line", "x0": left, "x1": right, "top": top, "bottom": top,
                      "width": right - left, "height": 0, "orientation": "h",
                      "pts": [(left, top), (right, top)], "doctop": top + getattr(page, "initial_doctop", 0)})
    return {"explicit_horizontal_lines": lines} if lines else {}


def _glyph_widths(blocks: list[dict]) -> dict[tuple[str, float, str], float]:
    """Largura mediana de cada letra na página, por fonte e corpo."""
    samples: dict[tuple[str, float, str], list[float]] = {}
    for block in blocks:
        for line in block.get("lines", []):
            for span in line["spans"]:
                key = (span["font"], round(span["size"], 1))
                for char in span["chars"]:
                    if char["c"].isalpha():
                        samples.setdefault(key + (char["c"],), []).append(char["bbox"][2] - char["bbox"][0])
    return {key: statistics.median(values) for key, values in samples.items() if len(values) >= 2}


def _span_text(span: dict, widths: dict[tuple[str, float, str], float]) -> str:
    """Texto do trecho com a caixa que aparece na página. Em versalete (DOE de 21.10.2026), o PDF desenha
    maiúsculas mas guarda o que foi digitado: "A AssemblEia Legislativa do Estado do Pará estatui ..." aparece
    como "A ASSEMBLEIA LEGISLATIVA DO ESTADO DO PARÁ estatui ...". A palavra só passa a maiúsculas quando todas
    as suas minúsculas têm a largura da maiúscula correspondente e pelo menos duas delas têm larguras bem
    diferentes nas duas caixas ("s": 3,65 x 4,79 pt) - espaçamento de justificação não produz isso."""
    fitz_box = lambda c: (c["bbox"][0], c["bbox"][1], c["bbox"][2])  # noqa: E731
    raw = span["chars"]
    chars = [char for index, char in enumerate(raw)
             if not (char["c"] == " " and 0 < index < len(raw) - 1 and raw[index - 1]["c"] != " "
                     and raw[index + 1]["c"] != " " and _phantom_space(raw[index - 1], char, raw[index + 1], fitz_box))]
    text = "".join(char["c"] for char in chars)
    key = (span["font"], round(span["size"], 1))
    pieces, start = [], 0
    for word in re.finditer(r"[^\W\d_]+", text):
        letters = chars[word.start():word.end()]
        lower = [c for c in letters if c["c"].islower() and c["c"].upper() != c["c"]]
        if lower:
            matches, strong = True, 0
            for char in lower:
                small, capital = widths.get(key + (char["c"],)), widths.get(key + (char["c"].upper(),))
                width = char["bbox"][2] - char["bbox"][0]
                if small is not None and abs(width - small) <= .2:
                    matches = False              # largura de minúscula: a letra aparece minúscula
                    break
                if capital is not None:
                    if abs(width - capital) > .2:
                        matches = False
                        break
                    strong += small is not None and abs(capital - small) >= .5
                # sem maiúscula dessa letra na página ("T", "Á"): basta não ter a largura da minúscula
            if matches and strong >= 2:
                pieces.append(text[start:word.start()] + word.group(0).upper())
                start = word.end()
    return "".join(pieces) + text[start:]


def fix_char_case(page: Any, fitz_page: Any, positions: list[tuple[float, float]] | None = None) -> int:
    """As células das tabelas são lidas pelo pdfplumber; o texto corrido, pelo PyMuPDF. Em PDFs
    reprocessados (o DOE de 09.10.2026 passou pelo iLovePDF), o pdfplumber lê parte das maiúsculas como
    minúsculas ("aBaETETUBa", "ToTal"), enquanto o PyMuPDF lê o que aparece na página ("ABAETETUBA").
    Cada letra do pdfplumber recebe a caixa da letra do PyMuPDF na mesma posição, só quando as duas forem a
    mesma letra; nenhum outro caractere muda. O mesmo ocorria no DOE de 05.10.2026 ("NoME", "MAtRÍcuLA",
    "servidor" onde a página mostra "Servidor"). Devolve quantas letras foram corrigidas; em positions,
    acrescenta o centro de cada uma."""
    reference: dict[int, list[tuple[float, float, str]]] = {}
    for block in fitz_page.get_text("rawdict")["blocks"]:
        for line in block.get("lines", []):
            for span in line["spans"]:
                for char in span["chars"]:
                    x0, _, _, bottom = char["bbox"]
                    reference.setdefault(int(x0), []).append((x0, bottom, char["c"]))
    fixed = 0
    for char in page.chars:
        text = char["text"]
        if len(text) != 1 or not text.isalpha():
            continue
        found = next((c for key in (int(char["x0"]), int(char["x0"]) - 1, int(char["x0"]) + 1)
                      for x0, bottom, c in reference.get(key, ())
                      if abs(x0 - char["x0"]) < .5 and abs(bottom - char["bottom"]) < 1), None)
        if found and found != text and found.casefold() == text.casefold():
            char["text"] = found
            fixed += 1
            if positions is not None:
                positions.append(((char["x0"] + char["x1"]) / 2 if "x1" in char else char["x0"],
                                  (char["top"] + char["bottom"]) / 2 if "top" in char else char["bottom"]))
    return fixed


def grouped_words(words: list[dict], tolerance: float = 1.8) -> list[tuple[float, list[dict]]]:
    groups: list[tuple[float, list[dict]]] = []
    for word in sorted(words, key=lambda w: (w["top"], w["x0"])):
        if groups and abs(groups[-1][0] - word["top"]) <= tolerance:
            groups[-1][1].append(word)
        else:
            groups.append((word["top"], [word]))
    return groups


def words_inside(words: list[dict], box: tuple[float, float, float, float]) -> list[dict]:
    x0, top, x1, bottom = box
    return [w for w in words if top - .15 <= w["top"] < bottom + .15 and x0 - .15 <= (w["x0"] + w["x1"]) / 2 < x1 + .15]


def cell_lines(words: list[dict], box: tuple[float, float, float, float], concatenate: bool = False) -> list[tuple[float, str]]:
    lines = []
    for top, members in grouped_words(words_inside(words, box)):
        ordered = sorted(members, key=lambda w: w["x0"])
        if not concatenate:
            lines.append((top, " ".join(w["text"] for w in ordered)))
            continue
        # junta só os pedaços de um mesmo número ("1 .234,56"); um espaço de verdade continua espaço, ou
        # "24.01, 24.02, 24.03" vira "24.01,24.02" e a conferência acha o valor "01,24" (MP nº 1.399, Anexo IV)
        text = ordered[0]["text"] if ordered else ""
        for previous, word in zip(ordered, ordered[1:]):
            gap = word["x0"] - previous["x1"]
            text += ("" if gap < .25 * (previous["bottom"] - previous["top"]) else " ") + word["text"]
        lines.append((top, text))
    return lines


def source_numbers_for_table(words: list[dict], table: Any) -> Counter:
    numbers: Counter = Counter()
    for row in table.rows:
        for cell in row.cells:
            if cell is None:
                continue
            for _, value in cell_lines(words, cell, concatenate=True):
                if not PERCENT.fullmatch(value):
                    numbers.update(monetary_values([[value]]))
    return numbers


def extract_table(page: Any, table: Any, words: list[dict], page_number: int) -> TableBlock:
    rows: list[list[str]] = []
    issues: list[str] = []
    output_map: list[list[int]] = []
    extracted = table.extract()
    for row_index, geometry in enumerate(table.rows):
        boxes = geometry.cells
        first = boxes[0]
        if first is None:
            # segunda linha do cabeçalho, embaixo de "Descrição" mesclada em duas linhas e sem números acima
            # ("Descrição | Ad valorem (%) | Específica (R$)" / "Inicial | Fator de conversão | Final", MP nº 1.399)
            header_row = row_index <= 2 and not monetary_values(
                [[cell or "" for cell in extracted[index]] for index in range(row_index)])
            if any(extracted[row_index]):
                if not header_row:
                    issues.append(f"Linha {row_index + 1} com célula mesclada: confirme cabeçalho e colunas no PDF.")
                output_map.append([len(rows)])
                rows.append([cell or "" for cell in extracted[row_index]])
            else:
                output_map.append([])
            continue
        labels: list[list] = []
        for top, text in cell_lines(words, first):
            if ROW_START.match(text):
                labels.append([top, text])
            elif labels:
                labels[-1][1] += " " + text
            elif row_index > 0 and text and not (extracted[row_index][0] or "").startswith("ITEM"):
                # Pode ser um cabeçalho de coluna, tratado como linha única.
                pass
        if len(labels) <= 1:
            raw = [cell or "" for cell in extracted[row_index]]
            if not any(raw):
                output_map.append([])
                continue
            # pdfplumber.Table.extract pode cortar o primeiro dígito junto à
            # borda de uma célula. Para linhas simples, reconstituímos o valor
            # a partir das palavras da coluna inteira, incluindo fragmentos.
            for ci in range(1, len(boxes)):
                cell = boxes[ci]
                if cell is None:
                    continue
                values = [text for _, text in cell_lines(words, cell, concatenate=True) if MONEY.fullmatch(text)]
                if len(values) == 1:
                    raw[ci] = values[0]
                elif len(values) > 1:
                    issues.append(f"Linha visual {row_index + 1}, coluna {ci + 1}: vários valores sem rótulo individual.")
            output_map.append([len(rows)])
            rows.append(raw)
            continue
        split = [[label] + [""] * (len(boxes) - 1) for _, label in labels]
        for ci in range(1, len(boxes)):
            cell = boxes[ci]
            if cell is None:
                issues.append(f"Linha visual {row_index + 1}, coluna {ci + 1}: célula irregular.")
                continue
            for top, value in cell_lines(words, cell):
                compact = "".join(value.split())
                if MONEY.fullmatch(compact) or PERCENT.fullmatch(compact):
                    value = compact
                candidates = [i for i, (label_top, _) in enumerate(labels) if label_top <= top + 1.5]
                if not candidates:
                    issues.append(f"Conteúdo {value!r} sem rótulo correspondente na coluna {ci + 1}.")
                target = candidates[-1] if candidates else 0
                if split[target][ci]:
                    issues.append(f"Várias linhas na coluna {ci + 1} do mesmo rótulo: confira a associação.")
                    split[target][ci] += "\n" + value
                else:
                    split[target][ci] = value
        output_map.append(list(range(len(rows), len(rows) + len(split))))
        rows.extend(split)

    merges: list[tuple[int, int, int, int]] = []
    for ri, visual in enumerate(table.rows):
        if len(output_map[ri]) != 1:
            continue
        for ci, cell in enumerate(visual.cells):
            if cell is None:
                continue
            last_col = ci
            for next_col in range(ci + 1, len(visual.cells)):
                if visual.cells[next_col] is not None:
                    break
                if any(other.cells[next_col] and other.cells[next_col][0] < cell[2] - .5
                       for other in table.rows):
                    last_col = next_col
            last_row = ri
            for next_row in range(ri + 1, len(table.rows)):
                top = min((b[1] for b in table.rows[next_row].cells if b), default=cell[3])
                if top >= cell[3] - .5 or table.rows[next_row].cells[ci] is not None:
                    break
                last_row = next_row
            if (last_col > ci or last_row > ri) and all(
                    len(output_map[index]) == 1 for index in range(ri, last_row + 1)):
                merges.append((output_map[ri][0], ci, output_map[last_row][0], last_col))

    expected = source_numbers_for_table(words, table)
    actual = monetary_values(rows)
    if expected != actual:
        issues.append(
            f"A conferência dos números falhou: {sum(expected.values())} lidos no PDF, "
            f"{sum(actual.values())} nas células exportáveis. Revise a tabela antes de usar."
        )
    column_geometry = table.rows[0].cells
    widths = [min((row.cells[ci][2] - row.cells[ci][0]
                   for row in table.rows if row.cells[ci]), default=30.0)
              for ci in range(len(column_geometry))]
    if isinstance(table, TextTable) or not hasattr(table, "page"):
        source = character_signature("".join(word["text"] for word in words_inside(words, table.bbox)))
    else:
        # letras com o centro dentro da tabela, como em Table.extract: "CCE/FC" + "E" que passa da borda da
        # célula (Portaria MME nº 943, DOU de 26.10.2026) não conta como texto perdido
        x0, top, x1, bottom = table.bbox
        source = character_signature("".join(
            char["text"] for char in table.page.chars
            if x0 <= (char["x0"] + char["x1"]) / 2 < x1 and top <= (char["top"] + char["bottom"]) / 2 < bottom))
    if matrix_signature(rows) != source:
        issues.append("Parte do texto da tabela no PDF não corresponde às células extraídas; confira todas as células.")
    strategy = getattr(table, "strategy", "linhas")
    if strategy == "alinhamento":
        issues.append("Tabela sem grade identificada por alinhamento: confira linhas e colunas no PDF.")
    return TableBlock(rows, widths, page_number, table.bbox, expected, source,
                      deepcopy(rows), strategy, merges, issues)


# Linha que termina no meio da frase: a seguinte é continuação, ainda que comece por
# "Portaria ME nº 284" (parece epígrafe), "DIRETOR-EXECUTIVO" (parece cargo) ou "o DIRETOR ...,
# no uso" (parece preâmbulo). Ato Técnico Conjunto nº 6/2026, DOU de 01.10.2026.
MID_SENTENCE = re.compile(r"(?:\b(?:de|da|do|das|dos|e|em|o|a|os|as|ao|aos|à|às|no|na|nos|nas|pelo|pela|"
                          r"pelos|pelas|com|para|por|que|sobre|entre|um|uma))\s*$", re.I)   # sem vírgula: "...," / "R E S O L V E:"
ITEM_NUMBER = re.compile(r"^\d{1,2}\.\s+\S")      # "1. Anexo VI ..." / "2. Anexo VII ..."
# Mobília do DOU: cabeçalho, expediente e rodapé de autenticação
DOU_HEADER = re.compile(r"^(?:ISSN\s+\d{4}-\d{3}[\dX]\b|REP[ÚU]BLICA FEDERATIVA DO BRASIL\b|"
                        r"Ano\s+[IVXLCDM]+\s+N[º°o]|Bras[íi]lia\s*-\s*DF\s*,|Di[áa]rio Oficial da Uni[ãa]o\b|"
                        r"EDI[ÇC][ÃA]O EXTRA\b|SE[ÇC][ÃA]O\s+\d\b)", re.I)
DOU_FOOTER = re.compile(r"PRESID[ÊE]NCIA DA REP[ÚU]BLICA\s*[•·]|Este documento pode ser verificado no endere[çc]o|"
                        r"Documento assinado digitalmente conforme MP\b", re.I)   # busca em qualquer ponto da linha


def dou_margins(lines: list[dict], page_height: float) -> tuple[float, float]:
    """Faixa útil de uma página do DOU: abaixo do cabeçalho (ISSN, "Ano CLXIV Nº 186-A", "Brasília - DF, ...",
    número da página) e acima do expediente ("PRESIDÊNCIA DA REPÚBLICA • CASA CIVIL ...") e do rodapé de
    autenticação. Sem essas linhas, devolve a página inteira."""
    def page_number(line: dict) -> bool:
        # O número da página fica sozinho na sua altura; "81" na primeira coluna de uma linha de tabela, não
        # (Portaria MPA nº 778, DOU de 26.10.2026, p. 65: a tabela perdia a última linha e saía do Word)
        return bool(re.fullmatch(r"\d{1,4}", line["text"].strip())) and not any(
            other is not line and abs(other["top"] - line["top"]) < 1 for other in lines)
    header = [l["bottom"] for l in lines if l["top"] < page_height * .15 and DOU_HEADER.match(l["text"].strip())]
    header += [l["bottom"] for l in lines if l["top"] < page_height * .085 and page_number(l)]
    footer = [l["top"] for l in lines if l["top"] > page_height * .5 and DOU_FOOTER.search(l["text"])]
    footer += [l["top"] for l in lines if l["top"] > page_height * .92 and page_number(l)]
    return (max(header) + .5 if header else 0.0), (min(footer) - 1 if footer else page_height)


PAGE_HEADER = re.compile(r"DI[ÁA]RIO OFICIAL N[º°]\s*[\d.]+|^(?:Segunda|Ter[çc]a|Quarta|Quinta|Sexta)-feira,\s*\d"
                         r"|^(?:S[áa]bado|Domingo),\s*\d|DO-e/SEFA N[º°]:", re.I)
PAGE_HEADER_TOP = re.compile(r"^\d{1,2} de [a-zç]+ de \d{4}$|^\d{1,4}$", re.I)   # só bem no alto (DO-e/SEFA)


def page_top_margin(lines: list[dict], page_height: float) -> float | None:
    """Fim do cabeçalho de uma página do DOE ou do DO-e/SEFA ("2 ▪ DIÁRIO OFICIAL Nº 36.789",
    "Sexta-feira, 02 de OUTUBRO de 2026"; "2 de outubro de 2026 / DO-e/SEFA Nº: 646 / 2"). No DOE o texto
    começa a 40 pt; com a faixa fixa de 52 pt, as primeiras linhas de cada continuação se perdiam."""
    # Só a faixa de cima (até 36 pt): o texto dos atos começa a 40 pt no DOE e pode citar "DIÁRIO OFICIAL Nº ...".
    header = [l["bottom"] for l in lines if l["top"] < 36 and PAGE_HEADER.search(l["text"].strip())]
    if not header:
        return None
    header += [l["bottom"] for l in lines if l["top"] < 36 and PAGE_HEADER_TOP.match(l["text"].strip())]
    return max(header) + .5


def three_columns(lines: list[dict], width: float) -> bool:
    """Página com três colunas de texto (cada uma com ao menos 8 linhas cheias dentro de um terço)."""
    third = width / 3
    counts = [0, 0, 0]
    for line in lines:
        size = line["x1"] - line["x0"]
        if .2 * width <= size <= .34 * width:
            column = int(line["x0"] // third)
            if column < 3 and line["x1"] <= (column + 1) * third + 10:
                counts[column] += 1
    return min(counts) >= 8


# "Belém, 1º de outubro de 2026." e, no DOU, "Brasília, 4 de novembro de 2026; 205º da Independência e 138º da
# República."
PLACE_DATE = re.compile(r"^[^,;:]+,\s*\d{1,2}[º°]?\s+de\s+\w+\s+de\s+\d{4}"
                        r"(?:;\s*\d+[º°]\s+da\s+Independ[eê]ncia\s+e\s+\d+[º°]\s+da\s+Rep[uú]blica)?\.?$", re.I)
OPEN_ENDING = re.compile(r"(?:\b(?:de|da|do|das|dos|e|em)|[-,])\s*$", re.I)


def signature_break(paragraphs: list[str], text: str) -> bool:
    """Nome / cargo / nome / cargo em linhas próprias, sem maiúsculas nem pontuação que
    separem os parágrafos: "Ana Carolina Lobo Gluck Paúl" / "Procuradora-Geral do Estado do Pará"."""
    last = paragraphs[-1]
    if looks_like_signer_name(last) and SIGNER_CARGO.match(text):
        return True                                   # cargo logo abaixo do nome
    if (len(paragraphs) >= 2 and looks_like_signer_name(paragraphs[-2]) and SIGNER_CARGO.match(last)
            and len(last) <= 90 and not OPEN_ENDING.search(last) and looks_like_signer_name(text)):
        return True                                   # próximo signatário, depois de nome + cargo
    if (len(paragraphs) >= 2 and looks_like_signer_name(last) and looks_like_signer_name(text)
            and not text.isupper() and len(text) <= 60
            and (PLACE_DATE.match(paragraphs[-2]) or looks_like_signer_name(paragraphs[-2]))):
        return True       # DOU: "LUIZ INÁCIO LULA DA SILVA" / "Dario Carnevalli Durigan" (ministros, em itálico)
    return bool(PLACE_DATE.match(last) and looks_like_signer_name(text))   # "Belém, 1º de ... de 2026"


_VOCABULARY = None   # palavras do diário em processamento (montado em extract_act)


OMISSION = re.compile(r"\[\s*(?:\.\s*){3}\]|\(\s*(?:\.\s*){3}\)|…|\[…\]|\(…\)")


def _line_marks(line: dict) -> set[int]:
    return {marks for text, marks in line.get("spans") or [] if text.strip()}


def font_turn(previous: dict, line: dict) -> bool:
    """Linha inteira em negrito seguida de linha sem negrito, ou o contrário, ambas sem itálico: o cargo
    "Coordenador da CCOT" e o despacho "Com base no art. 9º ... aprovo"; o título "DA CONSULTA" e o
    texto. Itálico fica de fora: a transcrição alterna negrito-itálico e itálico no mesmo parágrafo."""
    before, after = _line_marks(previous), _line_marks(line)
    if not before or not after or any(m & 2 for m in before | after):
        return False
    return (before == {1} and after == {0}) or (before == {0} and after == {1})


def side_by_side_signatures(lines: list[dict]) -> list[dict]:
    """Assinaturas em duas colunas lado a lado, depois da linha de local e data (Emenda Constitucional nº 140,
    DOU de 16.10.2026: "Mesa da Câmara dos Deputados" | "Mesa do Senado Federal", "Deputado Hugo Motta" |
    "Senador Davi Alcolumbre", "Presidente" | "Presidente" ...). Lidas linha a linha, as duas colunas se
    misturavam. Devolve a coluna da esquerda inteira e depois a da direita, cada linha marcada para ser um
    parágrafo próprio; sem esse arranjo, devolve as linhas na ordem de sempre."""
    ordered = sorted(lines, key=lambda l: (l["top"], l["x0"]))
    place = max((i for i, l in enumerate(ordered) if PLACE_DATE.match(l["text"].strip())), default=None)
    if place is None or len(ordered) - place - 1 < 4:
        return ordered
    tail = ordered[place + 1:]
    middle = (min(l["x0"] for l in tail) + max(l["x1"] for l in tail)) / 2
    left = [l for l in tail if l["x1"] <= middle + 2]
    right = [l for l in tail if l["x0"] >= middle - 2]
    if len(left) + len(right) != len(tail) or len(left) < 2 or len(right) < 2:
        return ordered
    paired = sum(1 for l in right if any(abs(l["top"] - other["top"]) < 1.5 for other in left))
    if paired < .6 * len(right):
        return ordered
    return ordered[:place + 1] + [dict(l, standalone=True) for l in sorted(left, key=lambda l: l["top"])
                                  + sorted(right, key=lambda l: l["top"])]


def merge_split_line(lines: list[dict]) -> list[dict]:
    """O PDF do DOU parte as linhas muito espaçadas pela justificação em pedaços na mesma altura ("agentes;" |
    "e" | "dá" | "outras"; "Art. 11." | "A União fornecerá ..."). Os pedaços voltam a ser uma linha; antes, o
    pedaço terminado em ";" ou "." fechava o parágrafo no meio da linha. Não junta células de tabela não
    reconhecida: pedaços alinhados em coluna com os da linha de cima ou de baixo ficam como estão."""
    rows: list[list[dict]] = []
    for line in lines:
        if rows and abs(line["top"] - rows[-1][-1]["top"]) < 1 and not line.get("standalone"):
            rows[-1].append(line)
        else:
            rows.append([line])
    starts = [[l["x0"] for l in row[1:]] for row in rows]

    def columnar(index: int) -> bool:
        for other in (index - 1, index + 1):
            if 0 <= other < len(rows) and abs(rows[other][0]["top"] - rows[index][0]["top"]) < 20:
                if sum(any(abs(x - y) < 2 for y in starts[other]) for x in starts[index]) >= 2:
                    return True
        return False

    tabular = {id(line) for index, row in enumerate(rows) if len(row) > 1 and columnar(index) for line in row}
    result: list[dict] = []
    for line in lines:
        previous = result[-1] if result else None
        if (previous and not previous.get("standalone") and not line.get("standalone") and id(line) not in tabular
                and abs(line["top"] - previous["top"]) < 1 and 0 <= line["x0"] - previous["x1"] < 40
                # dois textos longos lado a lado são colunas paralelas (duas portarias na mesma altura, DOE de
                # 09.10.2026, p. 147), não pedaços de uma linha justificada ("Art. 11." | "A União ...")
                and min(len(previous["text"]), len(line["text"])) < 25):
            merged = dict(previous, text=previous["text"] + " " + line["text"], x1=line["x1"],
                          bottom=max(previous["bottom"], line["bottom"]))
            if "spans" in previous or "spans" in line:
                merged["spans"] = (previous.get("spans") or [(previous["text"], 0)]) + [(" ", 0)] + (
                    line.get("spans") or [(line["text"], 0)])
            result[-1] = merged
        else:
            result.append(line)
    return result


# Texto alterado entre aspas: '"Art. 198. ......' depois de '.........." (NR)' abre parágrafo.
QUOTED_LABEL = re.compile(r'^[“"]\s*(?:Art\.|ART\.|§|[IVXLCDM]+\s*[-–]|[a-z]\s*\)|\.{3,})')

# Nome e cargo na mesma linha, separados por hífen: "DEPUTADO ADRIANO COELHO-  3º Secretário", "DEPUTADA CILENE
# COUTO -  1ª Secretária" (Emenda Constitucional nº 102 e Decretos Legislativos, DOE de 21.10.2026). Saíam
# grudados no cargo da assinatura anterior.
DOUBLE_ROMAN = re.compile(r"^[IVXLCDM]+\s+[IVXLCDM]+\s*[-–]\s")
NAME_DASH_CARGO = re.compile(r"^(?P<nome>[^\d,;:–-]{5,80}?)\s*[-–]\s+(?P<cargo>\S.{2,80})$")
# Local e data com palavra partida no fim da linha: "PALÁCIO CABANAGEM, PLENÁRIO NEWTON MIRANDA, EM 15 DE SETEM-" /
# "BRO DE 2026." (Decretos Legislativos, DOE de 21.10.2026). Sem palavra partida, as duas linhas ficam como no diário
# (revisão da Emenda Constitucional nº 102/2026).
PALACE = re.compile(r"^(?:PAL[AÁ]CIO|GABINETE|PLEN[AÁ]RIO)\b[^,]*,", re.I)


def split_name_and_cargo(lines: list[dict]) -> list[dict]:
    result = []
    for line in lines:
        found = NAME_DASH_CARGO.match(line["text"].strip())
        if (found and looks_uppercase(found.group("nome")) and looks_like_person_name(found.group("nome").strip())
                and SIGNER_CARGO.match(found.group("cargo"))
                # "3º Secretário", "Secretária de Estado ...": o cargo tem minúsculas; "SERVIDORES REMOVIDOS – AUDITOR
                # FISCAL DE RECEITAS ESTADUAIS – AFRE" (subtítulo da Portaria nº 522/2026-SEFA.GS) não é assinatura
                and not looks_uppercase(found.group("cargo"))):
            name, cargo = found.group("nome").strip(), found.group("cargo").strip()
            middle = line["x0"] + (line["x1"] - line["x0"]) * len(name) / max(len(line["text"]), 1)
            result.append({**line, "text": name, "x1": middle, "spans": [(name, _first_marks(line))], "standalone": True})
            result.append({**line, "text": cargo, "x0": middle, "spans": [(cargo, _first_marks(line))],
                           "standalone": True})
        else:
            result.append(line)
    return result


def _first_marks(line: dict) -> int:
    return next((marks for text, marks in line.get("spans") or [] if text.strip()), 0)


def reflow_lines(lines: list[dict], opinion: bool = False, groups: list[list[dict]] | None = None) -> list[str]:
    """Une somente quebras físicas curtas dentro do mesmo parágrafo. Com opinion (Solução de Consulta),
    a troca de negrito também separa parágrafos e "?" fecha a frase ("... distinta?" / "Resposta. ...").
    Em groups, devolve as linhas de origem de cada parágrafo."""
    paragraphs: list[str] = []
    members: list[list[dict]] = []
    endings = ('.', ':', ';', '?', '!') if opinion else ('.', ':', ';')
    previous_bottom = 0.0
    for line in split_name_and_cargo(merge_split_line(side_by_side_signatures(lines))):
        text = line["text"].strip()
        if not text:
            continue
        if (len(paragraphs) == 1 and line["top"] - previous_bottom <= 12
                and continues_epigraph(paragraphs[0], text)):
            paragraphs[0] = join_line(paragraphs[0], text)   # "... Nº 024 -" + "SEFA.GS, DE ... 2026"; ".../" + "PRODEPA"
            members[0].append(line)
            previous_bottom = line["bottom"]
            continue
        new_start = bool(EPIGRAPH.match(text) or (EMENTA_START.match(text) and text[:1].isupper())
                         or PREAMBLE.match(text) or COMMAND.match(text) or ARTICLE.match(text))
        if paragraphs and EPIGRAPH.match(paragraphs[-1]) and re.search(r"\b\d{4}\s*$", paragraphs[-1]):
            new_start = True
        # "... referente ao Processo Administrativo" / "63546/1996-ITERPA." (Decretos Legislativos, DOE de 21.10.2026):
        # a linha que começa por algarismo só tem maiúsculas nas letras da sigla e continua a frase.
        number_tail = bool(text[:1].isdigit() and paragraphs and not paragraphs[-1].rstrip().endswith(endings)
                           and line["top"] - previous_bottom <= 6
                           and not looks_uppercase(paragraphs[-1])    # título ou cabeçalho de tabela: não
                           and bool(re.search(r"[^\W\d_]\s*$", paragraphs[-1]))    # "... Fonte: 0.1.500" / "339030_": não
                           and len(re.findall(r"\d,\d{2}\b", text + " " + paragraphs[-1])) < 2)   # tabela de valores: não
        if (CARGO.match(text) or NAME.fullmatch(text) or looks_uppercase(text)) and not number_tail:
            # Solução de Consulta: a ementa vem toda em maiúsculas, sem negrito; títulos e assinaturas
            # são os que estão em negrito. Ali a linha em maiúsculas sem negrito continua o parágrafo.
            marks = _line_marks(line)
            if (not opinion or not paragraphs or not marks or all(m & 1 for m in marks)
                    or paragraphs[-1].rstrip().endswith(endings)):
                new_start = True
        if paragraphs and signature_break(paragraphs, text):
            new_start = True
        close = bool(paragraphs) and line["top"] - previous_bottom <= 6
        if (new_start and close and PREAMBLE.match(paragraphs[-1]) and not paragraphs[-1].rstrip().endswith(endings)
                and not re.match(r"CONSIDERANDO\b", paragraphs[-1], re.I)   # "Considerando LAUDO PERICIAL" / "PROCESSO: ..."
                and not any(rx.match(text) for rx in (EPIGRAPH, COMMAND, ARTICLE, PREAMBLE))
                and not line.get("standalone")):
            new_start = False     # "... estatui e sua Mesa" / "Diretora promulga ...": o cargo não abre parágrafo
        if (new_start and close and PALACE.match(paragraphs[-1]) and paragraphs[-1].rstrip().endswith(("-", "\xad"))
                and not looks_like_person_name(text) and not EPIGRAPH.match(text) and not line.get("standalone")):
            new_start = False     # "PALÁCIO ..., EM 15 DE SETEM-" / "BRO DE 2026." (palavra partida)
        if ITEM_NUMBER.match(text):
            new_start = True
        # "DOS TRABALHOS CONJUNTOS" / "Seção I" / "Da organização" (ATC RFB/CGIBS nº 9): a seção e o seu nome
        # vêm em linhas próprias, sem caixa alta, e grudavam num parágrafo só.
        if HEADING.match(text) or (paragraphs and HEADING.match(paragraphs[-1].strip())):
            new_start = True
        if QUOTED_LABEL.match(text) or (paragraphs and re.search(r"\((?:NR|AC)\)\s*[.;]?\s*$", paragraphs[-1])):
            new_start = True      # '"Art. 198. ...' / depois de '..." (NR)' ou '..." (AC)'
        if opinion and members and (font_turn(members[-1][-1], line) or OMISSION.fullmatch(text)
                                    or OMISSION.fullmatch(members[-1][-1]["text"].strip())):
            new_start = True      # "[...]" das transcrições fica em parágrafo próprio
        if paragraphs and ANNEX.match(paragraphs[-1]) and not re.search(r"\s[-–]\s", paragraphs[-1]):
            new_start = True      # "ANEXO ÚNICO" sozinho: a linha seguinte é o conteúdo do anexo
        if (new_start and paragraphs and line["top"] - previous_bottom <= 6
                and MID_SENTENCE.search(paragraphs[-1])
                and not any(rx.match(text) for rx in (ARTICLE, COMMAND, ANNEX, HEADING, EMENTA_START))
                and not ITEM_NUMBER.match(text)
                and not (PREAMBLE.match(text) and text[:1].isupper())):   # "..., e" / "Considerando ..." (IN 024)
            new_start = False
        if (new_start and paragraphs and line["top"] - previous_bottom <= 6 and MID_SENTENCE.search(paragraphs[-1])
                and not re.search(r";\s*(?:e|ou)\s*$", paragraphs[-1])
                and re.match(r"(?:Art|ART|§|Par[aá]grafo|PAR[AÁ]GRAFO)", text)):
            new_start = False     # "...conforme previsto no" / "§ 1º O ..." nunca abre dispositivo no meio da frase
        if paragraphs and re.search(r"\b(?:do|da|dos|das)\s*$", paragraphs[-1], re.I) and EPIGRAPH.match(text):
            # "... e do / Decreto Estadual nº ..." continua a frase; o nome
            # de uma lei citada não é a epígrafe de um segundo ato.
            new_start = False
        if (paragraphs and re.search(r"\b(?:desta|deste)\s*$", paragraphs[-1], re.I)
                and re.fullmatch(r"(?:LEI|DECRETO|PORTARIA|RESOLUÇÃO)\.", text, re.I)):
            new_start = False
        if paragraphs and re.search(r";\s*(?:e|ou)\s*$", paragraphs[-1]) and DOUBLE_ROMAN.match(text):
            new_start = True      # "...; e" / "III II - Anexo II, ..." (numeração repetida no diário)
        if (new_start and len(paragraphs) >= 2 and HEADING.match(paragraphs[-2].strip()) and members
                and line["top"] - previous_bottom <= 6
                and looks_uppercase(paragraphs[-1]) and looks_uppercase(text)
                and not paragraphs[-1].rstrip().endswith((".", ";", ":"))
                and abs((line["x0"] + line["x1"]) / 2 - (members[-1][-1]["x0"] + members[-1][-1]["x1"]) / 2) <= 4
                and not any(rx.match(text) for rx in (HEADING, ARTICLE, ANNEX, EPIGRAPH))):
            # nome do capítulo em várias linhas centralizadas: "DOS VEÍCULOS, DAS AERONAVES, ..." / "DAS BEBIDAS
            # ALCOÓLICAS, ..." / "EXTRAÍDOS" (CAPÍTULO II da MP nº 1.399/2026) é um título só
            new_start = False
        if line.get("standalone"):
            new_start = True      # assinaturas lado a lado: cada linha é um parágrafo
        if paragraphs and CHANGE_MARK.match(text) and re.search(r'[”"]\s*$', paragraphs[-1]) and not line.get("standalone"):
            paragraphs[-1] = paragraphs[-1].rstrip() + " " + text   # '... agendamento prévio.”' / "(NR)" (3.24)
            members[-1].append(line)
            previous_bottom = line["bottom"]
            continue
        if (paragraphs and re.fullmatch(r"(?:e|ou|e/ou)", text) and paragraphs[-1].rstrip().endswith(";")
                and not line.get("standalone")):
            paragraphs[-1] += " " + text          # "... atividade profissional;" / "e" (inciso III)
            members[-1].append(line)
            previous_bottom = line["bottom"]
            continue
        too_far = bool(paragraphs and line["top"] - previous_bottom > 6)
        if (paragraphs and not new_start and not too_far
                and paragraphs[-1].endswith("\xad")
                and (text[:1].islower() or (text[:1].isupper() and paragraphs[-1][-2:-1].isupper()))):
            # hífen invisível do diário: "SO\xad" + "CIOECONÔMICO" também se une, sem espaço
            paragraphs[-1] = paragraphs[-1][:-1] + text
            members[-1].append(line)
        elif paragraphs and not new_start and not too_far and (
                not paragraphs[-1].rstrip().endswith(endings)
                or (ENDS_WITH_ABBREVIATION.search(paragraphs[-1].rstrip()) and text[:1].isdigit())):
            paragraphs[-1] = join_line(paragraphs[-1], text, _VOCABULARY)
            members[-1].append(line)
        else:
            paragraphs.append(text)
            members.append([line])
        previous_bottom = line["bottom"]
    if groups is not None:
        groups[:] = members
    return paragraphs


def pdf_issue_date(pdf_path: Path, document: Any) -> tuple[str, str, str] | None:
    """Usa apenas a data explicitamente impressa no cabeçalho da primeira página."""
    for line in pdf_lines(pdf_path, 1, document)[:8]:
        match = PORTUGUESE_DATE.search(line["text"])
        if match and line["top"] < 140:
            day, month, year = match.groups()
            return year, MONTH_NUMBERS[month.lower()], f"{int(day):02d}"
    return None


def publication_note(pdf_path: Path, epigraph: str, source_name: str | None,
                     issue_date: tuple[str, str, str] | None = None, extra_edition: bool = False) -> str | None:
    found = PDF_DATE.search(pdf_path.name)
    if not source_name or (not found and not issue_date):
        return None
    year, month, day = found.groups() if found else issue_date
    feminine = normalized(epigraph).startswith(("PORTARIA", "RESOLUCAO", "INSTRUCAO", "LEI", "SOLUCAO", "MEDIDA",
                                                "EMENDA"))
    extra = " Ed. Extra." if extra_edition or "EXTRA" in pdf_path.name.upper() else ""
    date = f"{day}.{month}.{year[-2:]}" + ("." if extra else "")   # "de 02.10.26. Ed. Extra." (todos os diários)
    return f"{'Publicada' if feminine else 'Publicado'} no {source_name} de {date}{extra or '.'}"


def is_extra_edition(pdf_path: Path, document: Any) -> bool:
    """Edição extra indicada no próprio cabeçalho: "EDIÇÃO EXTRA" ou, no DOU, número com letra
    ("Ano CLXIV Nº 186-A")."""
    for line in pdf_lines(pdf_path, 1, document)[:12]:
        text = line["text"].strip()
        if re.search(r"\bEDI[ÇC][ÃA]O\s+EXTRA\b", text, re.I) or re.match(r"Ano\s+[IVXLCDM]+\s+N[º°o]\s*\d+-[A-Z]\b", text):
            return True
    return False


def detect_publication(pdf: Any, document: Any | None = None) -> str | None:
    """Diário de origem, lido primeiro no cabeçalho da 1ª página (quarto superior: título, número da
    edição, ISSN). Assim, um ato do DOE que cita "publicado no Diário Oficial da União" não faz o
    diário inteiro ser tratado como DOU. Sem marca no cabeçalho, procura na página inteira,
    começando pelos diários do Pará. Com document (PyMuPDF), lê também por ele: no DOE reprocessado de
    09.10.2026, o pdfplumber lia o título grande como "D O IARIO FICIAL"."""
    first = pdf.pages[0]
    x0, top, x1, bottom = first.bbox            # pode não começar exatamente em zero (DOU)
    header = normalized(first.crop((x0, top, x1, top + (bottom - top) * .25)).extract_text() or "")
    full = ""
    if document is not None and len(document):
        area = document[0].rect
        header += " " + normalized(document[0].get_text(clip=fitz.Rect(area.x0, area.y0, area.x1,
                                                                       area.y0 + area.height * .25)))
        full = " " + normalized(document[0].get_text())
    if "ISSN 1677-7042" in header or "IMPRENSA NACIONAL" in header or "DIARIO OFICIAL DA UNIAO" in header:
        return "DOU"
    if "DO-E/SEFA" in header or "DIARIO ELETRONICO DA SEFA" in header or "DIARIO OFICIAL ELETRONICO DA SECRETARIA" in header:
        return "DO-e/SEFA"
    if "DIARIO OFICIAL" in header and ("ESTADO DO PARA" in header or " IOE" in header):
        return "DOE (PA)"
    page = normalized(first.extract_text() or "") + full
    if "DO-E/SEFA" in page or "DIARIO ELETRONICO DA SEFA" in page:
        return "DO-e/SEFA"
    if "DIARIO OFICIAL" in page and "ESTADO DO PARA" in page:
        return "DOE (PA)"
    if "DIARIO OFICIAL DA UNIAO" in page or "ISSN 1677-7042" in page:
        return "DOU"
    if "CGIBS" in page:
        return "CGIBS"
    return None


# Solução de Consulta (DO-e/SEFA): não tem "RESOLVE:" nem artigos próprios; os "Art." são transcrições.
OPINION = re.compile(r"^SOLU[CÇ][AÃ]O\s+DE\s+CONSULTA\b", re.I)
EMENTA_LABEL = re.compile(r"^(EMENTA)\s*[:.–-]", re.I)


def source_runs_of(lines: list[dict]) -> list[tuple[str, int]]:
    runs: list[tuple[str, int]] = []
    for line in lines:
        if runs:
            runs.append((" ", 0))
        runs.extend(line.get("spans") or [(line["text"], 0)])
    return runs


def _all_bold(runs: list[tuple[str, int]] | None) -> bool:
    marks = [m for text, m in runs or [] if text.strip()]
    return bool(marks) and all(m & 1 for m in marks)


def opinion_blocks(texts: list[str], groups: list[list[dict]], page: int, publication: str | None,
                   first_region: bool) -> list[ParagraphBlock]:
    """Papéis da Solução de Consulta, com a formatação do diário: epígrafe, publicação, "EMENTA: ...",
    títulos de seção em negrito ("DA CONSULTA", "DA MANIFESTAÇÃO"), texto, e as duas assinaturas
    (parecerista e quem aprova), cada uma com nome e cargo em negrito."""
    runs: list[list[tuple[str, int]] | None] = [source_runs_of(group) for group in groups]
    if publication and texts:
        texts.insert(1, publication)
        runs.insert(1, None)

    def bold(i: int) -> bool:
        return i < len(texts) and _all_bold(runs[i])

    def short_bold(i: int) -> bool:
        return bold(i) and len(texts[i]) <= 90 and not texts[i].rstrip().endswith((".", ";", ","))

    blocks: list[ParagraphBlock] = []
    for index, text in enumerate(texts):
        attention: list[str] = []
        if index == 0 and first_region and EPIGRAPH.match(text):
            role, reason = "epigrafe", "Tipo, número e data do ato"
        elif publication and index == 1 and text == publication:
            role, reason = "publicacao", "Linha de publicação"
            attention.append("Linha criada a partir da identificação e da data da edição do PDF; confirme no diário.")
        elif EMENTA_LABEL.match(text):
            role, reason = "ementa", "Ementa da solução de consulta"
        elif (bold(index) and looks_like_signer_name(text) and index + 1 < len(texts)
              and (short_bold(index + 1) or SIGNER_CARGO.match(texts[index + 1]))):
            role, reason = "autoridade", "Nome em negrito seguido de cargo"
        elif blocks and blocks[-1].role == "autoridade" and (short_bold(index) or SIGNER_CARGO.match(text)):
            role, reason = "cargo", "Cargo após a assinatura"
        elif (bold(index) and looks_uppercase(text) and len(text) <= 100 and not ARTICLE.match(text)
              and not text.rstrip().endswith((".", ";", ",", ":"))):
            role, reason = "titulo", "Título de seção (negrito no diário)"
        else:
            role, reason = "texto", "Texto da solução de consulta (negrito e itálico copiados do diário)"
        item = Item(text, 0, 0, role, reason, list(attention))
        blocks.append(ParagraphBlock(text, role, page, "norma", item, list(attention), runs[index]))
    return blocks


def aligned_runs(text: str, source_runs: list[tuple[str, int]]) -> list[tuple[str, int]]:
    """Formatação do diário aplicada ao texto final (que pode ter perdido hífens de quebra, ganhado espaços
    ou sido corrigido na prévia): os caracteres são alinhados ignorando os espaços. Espaço entre trechos
    só leva a marca comum aos dois lados ("Art. 5º" em negrito + " O imposto" sem negrito)."""
    source = [(c, m) for piece, m in source_runs for c in piece if not c.isspace()]
    positions = [i for i, c in enumerate(text) if not c.isspace()]
    target = "".join(text[i] for i in positions)
    marks: list[int | None] = [None] * len(text)
    matcher = difflib.SequenceMatcher(a="".join(c for c, _ in source), b=target, autojunk=False)
    matched = 0
    for a0, b0, size in matcher.get_matching_blocks():
        for k in range(size):
            marks[positions[b0 + k]] = source[a0 + k][1]
        matched += size
    if not target or matched < .9 * len(target):
        return [(text, 0)]
    previous = None
    for i in positions:                       # caractere sem par: marca do vizinho
        if marks[i] is None:
            marks[i] = previous
        previous = marks[i]
    following = 0
    for i in reversed(positions):
        if marks[i] is None:
            marks[i] = following
        following = marks[i]
    before = [0] * len(text)
    last = 0
    for i, c in enumerate(text):
        before[i] = last
        if not c.isspace():
            last = marks[i]
    last = 0
    for i in range(len(text) - 1, -1, -1):
        if text[i].isspace():
            marks[i] = before[i] & last
        else:
            last = marks[i]
    runs: list[list] = []
    for c, m in zip(text, marks):
        if runs and runs[-1][1] == m:
            runs[-1][0] += c
        else:
            runs.append([c, m])
    return [(t, m) for t, m in runs]


def block_runs(block: ParagraphBlock) -> list[tuple[str, int]] | None:
    """Trechos para o Word quando a formatação vem do diário; None mantém as regras de negrito."""
    if block.source_runs is None:
        return None
    if block.role == "ementa":
        label = EMENTA_LABEL.match(block.text)          # "EMENTA" em negrito, como no exemplo corrigido
        return [(block.text[:label.end(1)], 1), (block.text[label.end(1):], 0)] if label else [(block.text, 0)]
    if block.role != "texto":
        return None
    return aligned_runs(block.text, block.source_runs)


def normative_blocks(lines: list[dict], page: int, publication: str | None = None,
                     opinion: bool = False, first_region: bool = False,
                     state: dict | None = None) -> list[ParagraphBlock]:
    if opinion:
        groups: list[list[dict]] = []
        texts = reflow_lines(lines, True, groups)
        return opinion_blocks(texts, groups, page, publication, first_region)
    texts = reflow_lines(lines)
    if publication and texts:
        texts.insert(1, publication)
    items = classify("\n".join(texts), state=state)
    blocks = []
    for index, item in enumerate(items):
        if publication and item.text == publication:
            item.attention.append("Linha criada a partir da identificação e da data da edição do PDF; confirme no diário.")
        if index + 1 < len(items) and looks_like_person_name(item.text) and SIGNER_CARGO.match(items[index + 1].text):
            item.role, item.reason = "autoridade", "Assinatura seguida de cargo"
        if index and SIGNER_CARGO.match(item.text) and items[index - 1].role == "autoridade":
            item.role, item.reason = "cargo", "Cargo após assinatura"
        blocks.append(ParagraphBlock(item.text, item.role, page, "norma", item, list(item.attention)))
    return blocks


def logo_block(page: Any, upper: float, lower: float, page_number: int) -> ImageBlock | None:
    candidates = [im for im in page.images if upper < im["top"] < im["bottom"] < lower and im["width"] < 120]
    if not candidates:
        return None
    im = min(candidates, key=lambda x: abs((x["x0"] + x["x1"]) / 2 - page.width / 2))
    box = (max(0, im["x0"] - 2), im["top"] - 2, min(page.width, im["x1"] + 2), im["bottom"] + 2)
    raster = page.crop(box).to_image(resolution=300).original
    stream = io.BytesIO()
    raster.save(stream, format="PNG")
    return ImageBlock(stream.getvalue(), page_number)


# Observação logo abaixo de uma tabela: "OBS *: BASE COM DEDUÇÃO DE 20% SOBRE A RECEITA,TOTAL REF. CON-" /
# "TB.AO FUNDEB" (Portaria nº 659/2026-SEFA). Antes, cada linha virava uma assinatura.
TABLE_NOTE = re.compile(r"^(?:OBS(?:ERVA[ÇC](?:[ÃA]O|[ÕO]ES))?\b|NOTAS?\s*[:.\d]|FONTE\s*:|\*)", re.I)


LEGEND_HEAD = re.compile(r"^(?:Siglas|Legenda|Abreviaturas)\s*:?$", re.I)
LEGEND_ITEM = re.compile(r"\s[-–]\s*[A-Z]{2,12}$")
NOTE_LABEL = re.compile(r"(?:OBS(?:ERVA[ÇC](?:[ÃA]O|[ÕO]ES))?|NOTAS?)\s*:", re.I)
NOTE_ITEM = re.compile(r"^(?:\d{1,2}\s*[).]|[a-z]\))\s*\S")


def after_table_blocks(lines: list[dict], top: float, bottom: float, page_number: int) -> list[ParagraphBlock]:
    lines = [line for line in lines if top < line["top"] < bottom]
    first_article = next((i for i, line in enumerate(lines) if ARTICLE.match(line["text"].strip())), None)
    if first_article is not None:
        before = after_table_blocks(lines[:first_article], top, bottom, page_number)
        signer = next((i for i in range(first_article + 1, len(lines) - 1)
                       if looks_like_person_name(lines[i]["text"].strip())
                       and CARGO.match(lines[i + 1]["text"].strip())), len(lines))
        return (before + normative_blocks(lines[first_article:signer], page_number) +
                after_table_blocks(lines[signer:], top, bottom, page_number))
    result: list[ParagraphBlock] = []
    previous_bottom = top
    legend = notes = False
    for line in lines:
        text = line["text"].strip()
        if not text or PROTOCOL.match(text):
            continue
        # "Siglas:" / "Secretaria Especial da Receita Federal do Brasil - RFB" (ATC RFB/CGIBS nº 9): legenda, não
        # assinatura ("Secretaria" parecia cargo e a linha seguinte, nome).
        if LEGEND_HEAD.match(text):
            legend = True
            result.append(ParagraphBlock(text, "texto", page_number, "footnote"))
            previous_bottom = line["bottom"]
            continue
        if legend and LEGEND_ITEM.search(text):
            result.append(ParagraphBlock(text, "texto", page_number, "footnote"))
            previous_bottom = line["bottom"]
            continue
        legend = False
        # "OBSERVAÇÕES:" sozinho e os itens "1) ...", "2) ..." em várias linhas (pautas do CARF/COSIT, DOU de
        # 26.10.2026): cada item é um parágrafo; antes, com o rótulo "O B S E R V AÇÕ ES :" gravado com espaços,
        # as linhas saíam juntas, e sem os espaços cada linha virava nome de assinatura
        if NOTE_LABEL.fullmatch(text):
            notes = True
            result.append(ParagraphBlock(text, "texto", page_number, "footnote"))
            previous_bottom = line["bottom"]
            continue
        if (notes and result and line["top"] - previous_bottom < 8 and not CARGO.match(text)
                and not (looks_like_person_name(text) and text.isupper())):
            if NOTE_ITEM.match(text) or NOTE_LABEL.fullmatch(result[-1].text.strip()):
                result.append(ParagraphBlock(text, "texto", page_number, "footnote"))
            else:
                result[-1].text = join_line(result[-1].text, text, _VOCABULARY)
            previous_bottom = line["bottom"]
            continue
        after_notes, notes = notes, False
        note_ended = after_notes or bool(result and TABLE_NOTE.match(result[-1].text)
                          and (result[-1].text.rstrip().endswith((".", ";", ":"))
                               or (looks_like_person_name(text) and text.isupper())))
        # "*Republicada por conter incorreção ... de 30/09/2026." / "MARILDA ..." (assinatura, não nota)
        if (result and result[-1].kind == "footnote" and line["top"] - previous_bottom < 8
                and not FOOTNOTE.match(text) and not CARGO.match(text) and not TABLE_NOTE.match(text)
                and not note_ended
                and not re.match(r"^(?:Belém|PAL[AÁ]CIO|GABINETE)\b", text, re.I)):
            if TABLE_NOTE.match(result[-1].text):     # observação: as linhas formam um só parágrafo
                result[-1].text = join_line(result[-1].text, text, _VOCABULARY)
            else:
                result[-1].text += "\n" + text
            previous_bottom = line["bottom"]
            continue
        if FOOTNOTE.match(text) or TABLE_NOTE.match(text):
            result.append(ParagraphBlock(text, "texto", page_number, "footnote"))
        elif re.match(r"^(?:Belém|PAL[AÁ]CIO|GABINETE)\b", text, re.I):
            result.append(ParagraphBlock(text, "texto", page_number, "date"))
        elif CARGO.match(text):
            result.append(ParagraphBlock(text, "cargo", page_number, "signer"))
        else:
            if len(text.split()) >= 2 and not re.search(r"[;:]$", text):
                result.append(ParagraphBlock(text, "autoridade", page_number, "signer"))
            else:
                result.append(ParagraphBlock(text, "texto", page_number, "footnote", attention=["Trecho após a tabela requer conferência."]))
        previous_bottom = line["bottom"]
    # Nas assinaturas em duas colunas, organiza nome e cargo da esquerda antes
    # do nome e cargo da direita, preservando a associação entre cada par.
    for offset in range(len(result) - 3):
        subset = result[offset:offset + 4]
        if [b.role for b in subset] == ["autoridade", "autoridade", "cargo", "cargo"]:
            result[offset + 1], result[offset + 2] = result[offset + 2], result[offset + 1]
    return result


def has_two_columns(page: Any, lines: list[dict], tables: list[Any]) -> bool:
    """Identifica colunas, inclusive antes de uma tabela larga de outro ato."""
    middle = page.width / 2
    wide = [t for t in tables if t.bbox[0] < middle - 15 and t.bbox[2] > middle + 15]
    if wide:
        # Uma tabela de outra publicação, abaixo dos protocolos das duas
        # colunas, não muda a leitura do trecho anterior da página.
        first_wide_top = min(t.bbox[1] for t in wide)
        earlier_tables = [t for t in tables if t.bbox[3] < first_wide_top]
        earlier_lines = [line for line in lines if line["top"] < first_wide_top]
        ended_sides = {(line["x0"] + line["x1"]) / 2 < middle
                       for line in earlier_lines if PROTOCOL.match(line["text"].strip())}
        if not earlier_tables or ended_sides != {True, False}:
            return False
        lines, tables = earlier_lines, earlier_tables
    left = [line for line in lines if line["x1"] < middle + 12]
    right = [line for line in lines if line["x0"] > middle - 12]
    # Uma tabela curta em cada margem pode não produzir as 12 linhas de texto
    # exigidas pelo detector geral, mas ainda indicar uma virada entre colunas.
    left_end = any(t.bbox[2] < middle + 12 and
                   t.bbox[3] > page.height - 75 for t in tables)
    right_start = any(t.bbox[0] > middle - 12 and
                      t.bbox[1] < 150 for t in tables)
    if left_end and right_start and len(left) >= 2 and len(right) >= 2:
        return True
    return len(left) >= 12 and len(right) >= 12 and \
        min(max(l["top"] for l in left), max(l["top"] for l in right)) - \
        max(min(l["top"] for l in left), min(l["top"] for l in right)) > 150


class TextTable:
    """Recorte de uma tabela encontrada por alinhamento, sem inventar células."""

    strategy = "alinhamento"

    def __init__(self, rows: list[Any], extracted: list[list[str | None]]):
        used = [i for i in range(len(extracted[0])) if any(
            line[i] and line[i].strip() for line in extracted)]
        self.rows = [SimpleNamespace(cells=[row.cells[i] for i in used]) for row in rows]
        self._extracted = [[line[i] for i in used] for line in extracted]
        boxes = [cell for row in self.rows for cell in row.cells if cell]
        self.bbox = (min(b[0] for b in boxes), min(b[1] for b in boxes),
                     max(b[2] for b in boxes), max(b[3] for b in boxes))

    def extract(self) -> list[list[str | None]]:
        return self._extracted


def _numeric_cells(row: list[str | None]) -> int:
    return sum(bool(NUMERIC_CELL.fullmatch(line.strip()))
               for cell in row[1:] if cell for line in cell.splitlines() if line.strip())


def find_text_tables(page: Any, start_y: float, stop_y: float, grid_tables: list[Any]) -> list[TextTable]:
    """Tenta uma segunda estratégia apenas em grupos tabulares reconhecíveis."""
    if stop_y - start_y < 35:
        return []
    x0, top, x1, bottom = page.bbox            # no DOU a página pode não começar exatamente em zero
    region = page.crop((x0, max(top, start_y), x1, min(bottom, stop_y)))
    settings = {"vertical_strategy": "text", "horizontal_strategy": "text",
                "min_words_vertical": 2, "min_words_horizontal": 1}
    result = []
    clean_region = None
    for candidate in region.find_tables(table_settings=settings):
        raw = candidate.extract()
        if not raw or len(raw[0]) < 2:
            continue
        data = [i for i, row in enumerate(raw) if row[0] and _numeric_cells(row)]
        if len(data) < 2:
            continue
        header = next((i for i in range(data[0] - 1, -1, -1)
                       if sum(bool(v and v.strip()) for v in raw[i]) >= 2), data[0])
        indices = [i for i in range(header, data[-1] + 1) if any(v and v.strip() for v in raw[i])]
        # o texto das células sem pontos brancos e sem espaços cobertos (o reconhecimento usa a página como está)
        clean_region = clean_region or visible_page(region)
        candidate.page = clean_region
        cleaned = candidate.extract()
        table = TextTable([candidate.rows[i] for i in indices], [cleaned[i] for i in indices])
        if table.bbox[2] - table.bbox[0] < 180:
            continue
        if any(max(0, min(table.bbox[2], other.bbox[2]) - max(table.bbox[0], other.bbox[0])) *
               max(0, min(table.bbox[3], other.bbox[3]) - max(table.bbox[1], other.bbox[1])) >
               (table.bbox[2] - table.bbox[0]) * (table.bbox[3] - table.bbox[1]) * .15
               for other in grid_tables):
            continue
        result.append(table)
    return result


def unrecognized_table_lines(words: list[dict], tables: list[Any], start_y: float, stop_y: float) -> int:
    """Conta linhas de valores alinhados que ficaram fora de qualquer tabela."""
    outside = [w for w in words if start_y <= w["top"] < stop_y and not any(
        t.bbox[0] - 2 <= (w["x0"] + w["x1"]) / 2 <= t.bbox[2] + 2 and
        t.bbox[1] - 2 <= w["top"] <= t.bbox[3] + 2 for t in tables)]
    likely = []
    for top, row in grouped_words(outside):
        fields = [w for w in row if MONEY.fullmatch(w["text"]) or PERCENT.fullmatch(w["text"])]
        if fields and any(w["x0"] + 65 < fields[0]["x0"] for w in row):
            likely.append((top, fields[0]["x0"], len(fields)))
    return sum(1 for i, (top, x, _) in enumerate(likely[1:], 1)
               if top - likely[i - 1][0] <= 38 and abs(x - likely[i - 1][1]) <= 12)


def line_midpoint(line: dict) -> float:
    """PyMuPDF e pdfplumber podem discordar na borda de uma tabela."""
    return (line["top"] + line["bottom"]) / 2


def line_inside_table(line: dict, table: Any) -> bool:
    mid_x = (line["x0"] + line["x1"]) / 2
    x0, top, x1, bottom = table.bbox
    return x0 - 1 <= mid_x <= x1 + 1 and top - 1 <= line_midpoint(line) <= bottom + 1


def table_title_after_signature(lines: list[dict], table_top: float) -> list[dict]:
    """Título em negrito logo acima de uma tabela publicada depois da assinatura, sem "ANEXO":
    "René de Oliveira e Sousa Júnior" / "Secretário de Estado da Fazenda" / "QUOTA PARTE DO IPI – Período
    de 01 a 30.09.2026" / tabela (Portaria nº 659/2026, DOE de 09.10.2026). Antes, o título grudava no cargo.
    Sai como título de tabela, como "DEMONSTRATIVO DA APURAÇÃO ..." na Portaria nº 494/2026."""
    ordered = sorted(lines, key=lambda l: l["top"])
    run: list[dict] = []
    index, limit = len(ordered) - 1, table_top
    while index >= 0 and len(run) < 3:
        line, text = ordered[index], ordered[index]["text"].strip()
        marks = _line_marks(line)
        if (not marks or not all(m & 1 for m in marks) or len(text) > 120
                or text.endswith((".", ";", ":", ",")) or limit - line["bottom"] > 14):
            break
        run.insert(0, line)
        limit, index = line["top"], index - 1
    if not run or index < 1:
        return []
    cargo, name = ordered[index]["text"].strip(), ordered[index - 1]["text"].strip()
    return run if SIGNER_CARGO.match(cargo) and looks_like_person_name(name) else []


DOU_AD = re.compile(r"in\.gov\.br|imprensa\s+nacional|baixe\s+o\s+app|ao\s+alcance\s+de\s+todos", re.I)


def column_images(page: Any, text_lines: list[dict], start_y: float, stop_y: float) -> list[dict]:
    """Imagens da coluna em que estão as linhas (a página do DOU tem duas colunas)."""
    if not text_lines:
        return []
    # lado da página: o expediente do DOU, centralizado, entra na lista da coluna direita e não serve de limite
    middle = page.width / 2
    left_share = sum((l["x0"] + l["x1"]) / 2 < middle for l in text_lines) / len(text_lines)
    side = "esquerda" if left_share >= .9 else "direita" if left_share <= .1 else None
    return [im for im in page.images if start_y < im["top"] < im["bottom"] < stop_y
            and (side is None or ((im["x0"] + im["x1"]) / 2 < middle) == (side == "esquerda"))]


def content_images(page: Any, text_lines: list[dict], start_y: float, stop_y: float) -> list[dict]:
    """Trecho do texto publicado como imagem: o inciso I do § 4º do art. 8º da MP nº 1.399/2026 ("I - o fator de
    ajuste será determinado mediante a aplicação da fórmula FA = 1/(1-MB), em que MB ...") é uma figura no DOU
    (edição extra B de 04.11.2026, p. 2). O texto não existe no PDF e o inciso sumia do Word."""
    found = []
    for im in sorted(column_images(page, text_lines, start_y, stop_y), key=lambda im: im["top"]):
        if im["width"] < 150 or im["height"] > 300:
            continue          # brasão / logotipo, ou página digitalizada
        if any(im["top"] < line_midpoint(line) < im["bottom"] and im["x0"] < (line["x0"] + line["x1"]) / 2 < im["x1"]
               for line in text_lines):
            continue          # fundo atrás do texto
        if any(im["top"] - 150 < line_midpoint(line) < im["bottom"] + 150 and DOU_AD.search(line["text"])
               for line in text_lines):
            continue          # propaganda da Imprensa Nacional ("Baixe o app do DOU", DOU de 26.10.2026, p. 64)
        found.append(im)
    return found


def content_image_block(page: Any, image: dict, page_number: int) -> ImageBlock:
    box = (max(0, image["x0"] - 1), image["top"] - 1, min(page.width, image["x1"] + 1), image["bottom"] + 1)
    raster = page.crop(box).to_image(resolution=300).original
    stream = io.BytesIO()
    raster.save(stream, format="PNG")
    return ImageBlock(stream.getvalue(), page_number, "content", image["width"])


def normative_with_images(page: Any, extraction: Extraction, lines: list[dict], images: list[dict],
                          page_number: int, publication: str | None, opinion: bool,
                          first_region: bool) -> list[Any]:
    """Os parágrafos da região, com cada imagem de texto no lugar em que aparece."""
    blocks: list[Any] = []
    rest = lines
    for image in images:
        above = [line for line in rest if line_midpoint(line) < image["top"]]
        rest = [line for line in rest if line_midpoint(line) >= image["top"]]
        if above:
            blocks.extend(normative_blocks(above, page_number, publication, opinion, first_region,
                                           extraction.quote_state))
            publication, first_region = None, False
        blocks.append(content_image_block(page, image, page_number))
        extraction.issues.append(
            f"Página {page_number}: um trecho do ato está publicado como imagem no PDF (fórmula ou figura) e foi "
            "incluído como imagem no Word, no mesmo lugar. Se quiser o texto, digite-o a partir da imagem.")
    if rest or not blocks:
        blocks.extend(normative_blocks(rest, page_number, publication, opinion, first_region,
                                       extraction.quote_state))
    return blocks


def append_region(page: Any, pdf_path: Path, extraction: Extraction, text_lines: list[dict],
                  page_tables: list[Any], words: list[dict], start_y: float, stop_y: float,
                  page_number: int, publication: str | None, split_columns: bool,
                  opinion: bool = False, first_region: bool = False) -> None:
    tables = [table for table in page_tables if table.bbox[1] >= start_y
              and table.bbox[3] < stop_y + 1 and table.bbox[2] - table.bbox[0] > 200]
    if not split_columns:
        tables.extend(find_text_tables(page, start_y, stop_y, tables))
    tables.sort(key=lambda t: t.bbox[1])
    if unrecognized_table_lines(words, tables, start_y, stop_y):
        extraction.blocking_issues.append(
            f"Página {page_number}: há valores alinhados fora das tabelas reconhecidas. "
            "Use o PDF original para reconstruir este anexo antes de exportar.")
    if not tables:
        region = [line for line in text_lines if start_y <= line["top"] < stop_y]
        images = content_images(page, text_lines, start_y, stop_y)
        extraction.blocks.extend(normative_with_images(page, extraction, region, images, page_number, publication,
                                                       opinion, first_region))
        return
    outside_lines = [line for line in text_lines if start_y <= line["top"] < stop_y
                     and not any(line_inside_table(line, table) for table in tables)]
    first_top = tables[0].bbox[1]
    # brasão acima da tabela, na mesma coluna (a fórmula em imagem da coluna esquerda cortava o texto da direita
    # ao meio na MP nº 1.399: do inciso II do art. 14 em diante, cada linha virava um parágrafo)
    images = [im for im in column_images(page, text_lines, start_y, first_top) if im["width"] < 120]
    logo_top = min((im["top"] for im in images), default=first_top)
    all_lines = [line for line in outside_lines if line_midpoint(line) < first_top]
    annex_heading = min((line["top"] for line in all_lines
                         if ANNEX_HEAD.match(line["text"].strip())), default=first_top)
    normative_end = min(logo_top, annex_heading)
    before = [line for line in all_lines if line["top"] + 2 < normative_end]
    title = table_title_after_signature(before, first_top)
    if title:          # "QUOTA PARTE DO IPI – Período de 01 a 30.09.2026" (Portaria nº 659/2026-SEFA)
        normative_end = min(normative_end, title[0]["top"] - .1)
        before = [line for line in all_lines if line["top"] + 2 < normative_end]
    images = content_images(page, text_lines, start_y, normative_end)
    extraction.blocks.extend(normative_with_images(page, extraction, before, images, page_number, publication,
                                                   opinion, first_region))
    logo = logo_block(page, normative_end - 2, first_top, page_number)
    if logo:
        extraction.blocks.append(logo)
    # Entre o título do anexo e a primeira tabela, todas as linhas entram (antes só as do tipo "ANEXO ...",
    # "GOVERNO DO ESTADO", "DEMONSTRATIVO", "R$ 1,00"; o nome do anexo, como "SERVIDORES REMOVIDOS –
    # AUDITOR", era descartado). A conferência abaixo bloqueia se algum caractere não chegar ao Word.
    heading_lines = [line for line in all_lines if line["top"] >= normative_end and line["text"].strip()]
    heading_blocks = []
    for line in heading_lines:
        text = line["text"].strip()
        kind = "unit" if text.startswith("R$") else "annex_head"
        heading_blocks.append(ParagraphBlock(text, "anexo" if ANNEX.match(text) else "texto", page_number, kind))
    extraction.blocks.extend(heading_blocks)
    if logo is None and character_signature("".join(l["text"] for l in heading_lines)) != character_signature(
            "".join(b.text for b in heading_blocks)):
        extraction.blocking_issues.append(f"Página {page_number}: há texto antes da tabela que não chegou ao Word. "
                                          "Confira o anexo no PDF.")
    for table_index, table in enumerate(tables):
        if table_index:
            previous_bottom = tables[table_index - 1].bbox[3]
            gap_lines = [line for line in outside_lines
                         if previous_bottom < line_midpoint(line) < table.bbox[1]]
            if any(ARTICLE.match(line["text"].strip()) for line in gap_lines):
                gap_blocks = normative_blocks(gap_lines, page_number)
            else:
                gap_blocks, legend = [], False
                for line in gap_lines:
                    text = line["text"].strip()
                    if not text:
                        continue
                    # "Siglas:" / "... - RFB" entre o Anexo I e o Anexo II (ATC RFB/CGIBS nº 9): legenda da tabela
                    # acima, como a do fim do ato, e não título do anexo seguinte
                    legend = bool(LEGEND_HEAD.match(text) or (legend and LEGEND_ITEM.search(text)))
                    gap_blocks.append(ParagraphBlock(text, "anexo" if ANNEX.match(text) else "texto", page_number,
                                                     "footnote" if legend else "annex_head"))
            if gap_lines:
                if character_signature("".join(line["text"] for line in gap_lines)) != character_signature(
                        "".join(block.text for block in gap_blocks)):
                    extraction.blocking_issues.append(
                        f"Página {page_number}: há texto entre tabelas que não chegou ao Word. "
                        "Confira os artigos e títulos dos anexos antes de exportar.")
                else:
                    extraction.intertable_segments_checked += 1
            extraction.blocks.extend(gap_blocks)
        block = extract_table(page, table, words, page_number)
        if split_columns:
            column = "esquerda" if (table.bbox[0] + table.bbox[2]) / 2 < page.width / 2 else "direita"
            block.source_parts = [TablePagePart(page_number, table.bbox, 0,
                                                len(block.matrix), column)]
        extraction.blocks.append(block)
    after = [line for line in outside_lines if line_midpoint(line) > tables[-1].bbox[3]]
    if after:
        extraction.blocks.extend(after_table_blocks(after, min(line["top"] for line in after) - 1,
                                                  stop_y, page_number))


def _query_key(value: str) -> str:
    # "N°" (grau) e "Nº" (ordinal) valem o mesmo: o normalizado de "º" é "O"
    return re.sub(r"\s*([/.-])\s*", r"\1", normalized(value).replace("°", "O"))


_ACT_TYPES = (("INSTRUCAO NORMATIVA", ("INSTRUCAO NORMATIVA", "IN")), ("LEI COMPLEMENTAR", ("LEI COMPLEMENTAR", "LC")),
              ("LEI", ("LEI",)), ("DECRETO", ("DECRETO",)), ("PORTARIA", ("PORTARIA",)),
              ("RESOLUCAO", ("RESOLUCAO",)), ("ATO COTEPE", ("ATO COTEPE",)), ("CONVENIO", ("CONVENIO",)),
              ("AJUSTE", ("AJUSTE",)), ("PROTOCOLO", ("PROTOCOLO",)),
              ("ATO TECNICO", ("ATO TECNICO",)), ("ATO DECLARATORIO", ("ATO DECLARATORIO", "ADE")),
              ("MEDIDA PROVISORIA", ("MEDIDA PROVISORIA", "MP")), ("DESPACHO", ("DESPACHO",)),
              ("SOLUCAO DE CONSULTA", ("SOLUCAO DE CONSULTA",)),
              ("EMENDA CONSTITUCIONAL", ("EMENDA CONSTITUCIONAL", "EC")))


def _act_parts(text: str, epigraph: bool) -> tuple[str | None, str | None, str | None]:
    """(tipo, número sem pontos nem zeros à esquerda, ano) de uma epígrafe ou do que foi digitado."""
    key = normalized(text).replace("°", "O")
    kind = None
    for name, spellings in _ACT_TYPES:
        if any(re.search(rf"(?<![A-Z]){re.escape(word)}(?![A-Z])", key[:40] if epigraph else key)
               for word in spellings):
            kind = name
            break
    found = (re.search(r"\bN\s*(?:O\.?|\.\s*O?)?\s*(\d[\d.]*)", key) if epigraph else None) or re.search(r"(\d[\d.]*)", key)
    if not found:
        return kind, None, None
    number = found.group(1).replace(".", "").lstrip("0") or "0"
    rest = key[found.end():]
    year = re.match(r"\s*/\s*(\d{4}|\d{2})(?!\d)", rest)
    if year:
        year_value = year.group(1)
    else:
        years = re.findall(r"(?<!\d)((?:19|20)\d{2})(?!\d)", rest)
        year_value = years[-1] if years else None
    return kind, number, year_value


def _same_act(query: str, epigraph: str) -> bool:
    q_kind, q_number, q_year = _act_parts(query, epigraph=False)
    e_kind, e_number, e_year = _act_parts(epigraph, epigraph=True)
    if not q_number or q_number != e_number:
        return False
    if q_kind and e_kind and not e_kind.startswith(q_kind):
        return False
    return not q_year or not e_year or q_year[-2:] == e_year[-2:]


def _query_hit(query: str, epigraph: str) -> bool:
    """O identificador digitado aparece na epígrafe sem estar colado a outros algarismos:
    "4/2026" casa com "Nº 04/2026" (zeros à esquerda), mas não com "Nº 1.884/2026"."""
    q, e = _query_key(query), _query_key(epigraph)
    start = e.find(q) if q else -1
    while start >= 0:
        before = re.search(r"[\d.]*$", e[:start]).group(0).replace(".", "")
        after = e[start + len(q):start + len(q) + 1]
        if ((not q[0].isdigit() or not before.strip("0"))
                and (not q[-1].isdigit() or not after.isdigit())):
            return True
        start = e.find(q, start + 1)
    return False


CGIBS_SITE = re.compile(r"S[IÍ]TIO ELETR[OÔ]NICO DO (?:COMIT[EÊ] GESTOR DO IMPOSTO SOBRE BENS E SERVI[CÇ]OS|CGIBS)", re.I)


EC_WITHOUT_DATE = re.compile(r"^EMENDA\s+CONSTITUCIONAL\s+" + NUMBER_SIGN + r"\s*[\d.]+\s*$", re.I)
PLACE_DATE_PARTS = re.compile(r"^[^,;:]+,\s*(\d{1,2})[º°]?\s+de\s+([a-zç]+)\s+de\s+(\d{4})\.?$", re.I)


def date_constitutional_amendment(extraction: Extraction) -> None:
    """No DOU, a Emenda Constitucional vem só com o número ("EMENDA CONSTITUCIONAL Nº 140"); a data está no
    fim ("Brasília, 16 de outubro de 2026."). A epígrafe recebe a data, como na forma oficial de citação
    ("EMENDA CONSTITUCIONAL Nº 140, DE 16 DE OUTUBRO DE 2026"), com aviso para conferir."""
    paragraphs = extraction.paragraphs
    if not paragraphs or paragraphs[0].role != "epigrafe" or not EC_WITHOUT_DATE.match(paragraphs[0].text):
        return
    place = next((PLACE_DATE_PARTS.match(b.text.strip()) for b in paragraphs[1:] if PLACE_DATE_PARTS.match(b.text.strip())), None)
    if not place:
        return
    day, month, year = place.groups()
    epigraph = paragraphs[0]
    epigraph.text = f"{epigraph.text.strip()}, DE {int(day)} DE {month.upper()} DE {year}"
    note = (f"Data acrescentada à epígrafe a partir de “{place.group(0).strip()}”, como na forma oficial de citação; "
            "confira.")
    epigraph.attention.append(note)
    if epigraph.item:
        epigraph.item.text = epigraph.text
        epigraph.item.attention.append(note)


def add_cgibs_availability(extraction: Extraction) -> None:
    """DOU: quando o ato diz que entra em vigor com a publicação no DOU "e no sítio eletrônico do Comitê
    Gestor do IBS", acrescenta "Disponibilizado no site do CGIBS em dd.mm.aa." logo depois da linha de
    publicação, com a data do DOU (ATC nº 6/2026, corrigido à mão). A data deve ser confirmada no site."""
    blocks = extraction.blocks
    publication = next((i for i, b in enumerate(blocks) if isinstance(b, ParagraphBlock) and b.role == "publicacao"), None)
    if publication is None or not any(isinstance(b, ParagraphBlock) and CGIBS_SITE.search(b.text) for b in blocks):
        return
    found = re.search(r"\bde (\d{2}\.\d{2}\.\d{2})\b", blocks[publication].text)
    if not found:
        return
    verb = "Disponibilizada" if blocks[publication].text.startswith("Publicada") else "Disponibilizado"
    text = f"{verb} no site do CGIBS em {found.group(1)}."
    attention = ["Linha criada porque o ato também é publicado no site do CGIBS; a data é a do DOU: confirme no site."]
    item = Item(text, 0, 0, "publicacao", "Disponibilização no site do CGIBS", list(attention))
    blocks.insert(publication + 1, ParagraphBlock(text, "publicacao", blocks[publication].page, "norma", item, attention))


# Citação no corpo do ato, e não epígrafe do ato seguinte: tipo com inicial minúscula ("Lei n. 5.530",
# "Decreto nº 428") e a frase continua depois da data ("..., de 13 de janeiro de 1989, que disciplina ...";
# "..., de 27 de dezembro de 2006 e Instrução Normativa ..."). Antes, essas linhas encerravam o ato
# (Solução de Consulta nº 010/2026, "Da legislação aplicável"). Epígrafe em maiúsculas não muda.
CITATION_AFTER_DATE = re.compile(
    r"\b\d{1,2}[º°]?\s+de\s+(?:janeiro|fevereiro|mar[çc]o|abril|maio|junho|julho|agosto|setembro|outubro|"
    r"novembro|dezembro)\s+de\s+\d{4}\b\W*[A-Za-zÀ-ÿ]", re.I)


def cited_in_text(text: str) -> bool:
    return not text.split()[0].isupper() and bool(CITATION_AFTER_DATE.search(text))


def line_above(lines: list[dict], index: int) -> dict | None:
    """Linha imediatamente acima, na mesma coluna (faixas horizontais que se sobrepõem)."""
    current = lines[index]
    candidates = [l for l in lines[:index] if l["top"] < current["top"] - 1
                  and l["x0"] < current["x1"] and current["x0"] < l["x1"]]
    return max(candidates, key=lambda l: l["top"], default=None)


def _heading_like(text: str) -> bool:
    text = text.strip()
    return bool(text) and len(text) <= 120 and not text.endswith((".", ";", ":", ",")) and text[:1].isupper()


def trim_trailing_headings(extraction: Extraction, dou: bool = True) -> None:
    """DOU: entre o fim de um ato e a epígrafe do seguinte vêm os títulos do órgão ("Ministério da
    Fazenda", "SECRETARIA EXECUTIVA"). São descartados só quando vêm logo após uma assinatura
    reconhecida ou ao fim de um ANEXO que segue a assinatura; uma assinatura não reconhecida
    (nome e cargo como texto) nunca é apagada.
    DOE e DO-e/SEFA (3.24, dou=False): só o título de seção em maiúsculas logo depois do cargo de quem
    assina ("ASSEMBLEIA LEGISLATIVA - LEIS SANCIONADAS", antes da epígrafe da lei seguinte)."""
    blocks = extraction.blocks
    start = len(blocks)
    # Só na página onde está o último título: as siglas no pé do ANEXO, na página anterior, ficam (ATC RFB/CGIBS
    # nº 9, DOU de 26.10.2026: "Siglas:" na p. 38, "SECRETARIA ADJUNTA" ... na p. 39).
    last_page = getattr(blocks[-1], "page", None) if blocks else None
    while (start > 1 and isinstance(blocks[start - 1], ParagraphBlock) and blocks[start - 1].role == "texto"
           and _heading_like(blocks[start - 1].text) and blocks[start - 1].page == last_page
           and (dou or looks_uppercase(blocks[start - 1].text))):
        start -= 1
    if start == len(blocks):
        return
    before = blocks[start - 1]
    roles = [b.role if isinstance(b, ParagraphBlock) else "tabela" for b in blocks[:start]]
    after_signature = isinstance(before, ParagraphBlock) and before.role == "cargo"
    new_page = (last_page is not None and getattr(before, "page", None) is not None and before.page < last_page)
    annex_after_signature = dou and ("cargo" in roles and "anexo" in roles[len(roles) - 1 - roles[::-1].index("cargo"):]
                             and (not isinstance(before, ParagraphBlock) or new_page
                                  or before.text.rstrip().endswith((".", ";", ":"))))
    if after_signature or annex_after_signature:
        tail = blocks[start:]
        del blocks[start:]
        used = {getattr(b, "page", 0) for b in blocks} | {part.page for b in blocks if isinstance(b, TableBlock)
                                                          for part in table_page_parts(b)}
        extraction.pages = [p for p in extraction.pages if p <= max(used)]
        extraction.issues.append("Títulos de órgão após o fim do ato foram deixados de fora ("
                                 + "; ".join(b.text for b in tail)[:120] + "); confira.")


HOUSE_AD_TEXT = re.compile(r"in\.gov\.br|imprensa\s+nacional", re.I)


def _image_coverage(page: Any, bbox: tuple[float, float, float, float]) -> float:
    """Fração da caixa coberta por imagens da página (limitada a 1)."""
    x0, top, x1, bottom = bbox
    area = max((x1 - x0) * (bottom - top), 1.0)
    covered = 0.0
    for image in getattr(page, "images", []):
        width = min(x1, image["x1"]) - max(x0, image["x0"])
        height = min(bottom, image["bottom"]) - max(top, image["top"])
        if width > 0 and height > 0:
            covered += width * height
    return min(covered / area, 1.0)


def drop_dou_house_ads(extraction: Extraction, pdf: Any) -> None:
    """DOU: quadros de divulgação da Imprensa Nacional ("Biblioteca Machado de Assis", "Baixe o app do DOU")
    ocupam o espaço que sobra depois do último ato da página e eram lidos como tabela do ato, o que impedia
    a gravação (EC 141, DOU de 16.10.2026). São descartados só quando vêm depois de uma assinatura
    reconhecida, sem ANEXO, e o quadro é sobretudo imagem ou traz o endereço in.gov.br / "Imprensa
    Nacional". Uma tabela comum depois da assinatura continua no ato."""
    blocks = extraction.blocks
    last_cargo = max((i for i, b in enumerate(blocks) if isinstance(b, ParagraphBlock) and b.role == "cargo"),
                     default=None)
    if last_cargo is None or last_cargo == len(blocks) - 1:
        return
    tail = blocks[last_cargo + 1:]
    if not all(isinstance(b, (TableBlock, ImageBlock)) for b in tail):
        return
    tables = [b for b in tail if isinstance(b, TableBlock)]
    if not tables:
        return
    for table in tables:
        text = " ".join(cell for row in table.matrix for cell in row)
        covered = max(_image_coverage(pdf.pages[part.page - 1], part.bbox) for part in table_page_parts(table))
        if not (covered >= .3 or HOUSE_AD_TEXT.search(text)):
            return
    sample = " ".join(" ".join(cell for row in tables[0].matrix for cell in row).split())[:70]
    del blocks[last_cargo + 1:]
    pages = sorted({t.page for t in tables})
    extraction.issues.append(
        f"Quadro de divulgação do DOU depois das assinaturas (p. {', '.join(map(str, pages))}: “{sample}…”) "
        "foi deixado de fora; confira.")


def epigraph_text(lines: list[dict], index: int) -> str:
    """Epígrafe completa a partir da linha index, reunindo a continuação na linha seguinte."""
    text = lines[index]["text"].strip()
    for following in lines[index + 1:index + 3]:
        if not continues_epigraph(text, following["text"].strip()):
            break
        text += " " + following["text"].strip()
    return text


def _mostly_uppercase(text: str) -> bool:
    """"8704.31 E NOS CÓDIGOS 8704.41.00, 8704.51.00, 8704.60.00 e 8704.90.00 DA NCM/SH": o "e" minúsculo
    do diário não tira a linha do nome do anexo."""
    letters = [c for c in text if c.isalpha()]
    return len(letters) >= 8 and sum(c == c.upper() for c in letters) >= .85 * len(letters)


def _annex_name_with_number(parts: list[Any], text: str) -> bool:
    if not parts:
        return len(text) >= 30 and len(re.findall(r"[^\W\d_]{2,}", text)) >= 5
    previous = parts[-1].text.rstrip()
    return bool(previous.endswith(",") or MID_SENTENCE.search(previous)
                or re.match(r"(?:NA|NAS|NO|NOS|DA|DAS|DO|DOS|DE|E)\s", text))


def merge_annex_subtitles(extraction: Extraction) -> None:
    """Depois de "ANEXO I", as linhas em maiúsculas do nome do anexo ("SERVIDORES REMOVIDOS – AUDITOR" /
    "FISCAL DE RECEITAS ESTADUAIS – AFRE") formam um só parágrafo, com o estilo de subtítulo. Antes saíam
    partidas e, entre duas tabelas, sem estilo da SEFA (Portaria nº 522/2026-SEFA.GS)."""
    blocks, merged, index = extraction.blocks, [], 0
    while index < len(blocks):
        block = blocks[index]
        merged.append(block)
        index += 1
        if not (isinstance(block, ParagraphBlock) and block.role == "anexo"):
            continue
        # "(ANEXO I DA LEI Nº 99.100, DE 15 DE MARÇO DE 2024)": a indicação do anexo da lei alterada fica num
        # parágrafo próprio, antes do nome; colada ao nome, o Consolidador perdia o título do anexo (3.28.1)
        while (index < len(blocks) and isinstance(blocks[index], ParagraphBlock) and blocks[index].role == "texto"
               and blocks[index].kind in ("norma", "annex_head")
               and re.fullmatch(r"\(\s*ANEXO\b[^()]*\)", blocks[index].text.strip(), re.I)):
            merged.append(blocks[index])
            index += 1
        parts = []
        while (index < len(blocks) and len(parts) < 4 and isinstance(blocks[index], ParagraphBlock)
               and blocks[index].role == "texto" and blocks[index].kind in ("norma", "annex_head")
               and (looks_uppercase(blocks[index].text) or (parts and _mostly_uppercase(blocks[index].text)))
               and len(blocks[index].text) <= 120
               and not blocks[index].text.rstrip().endswith((".", ";", ":"))
               and not ANNEX.match(blocks[index].text)
               # nome de anexo é frase sem algarismos ("SERVIDORES REMOVIDOS – AUDITOR"); cabeçalho de coluna de
               # tabela não lida ("ÁREA/UNIDADE", "3º QUADRIMESTRE - 2026") não entra. Algarismos só na frase longa
               # ou na continuação dela: "AJUSTES ... CLASSIFICADOS" / "NA POSIÇÃO 87.03 DA NCM/SH" (MP nº 1.399)
               and (not re.search(r"\d", blocks[index].text) or _annex_name_with_number(parts, blocks[index].text))
               and (len(blocks[index].text) >= 8 if parts
                    else len(blocks[index].text) >= 15 and len(blocks[index].text.split()) >= 2)):
            parts.append(blocks[index])
            index += 1
        if parts:
            text = parts[0].text
            for part in parts[1:]:
                text = join_line(text, part.text, _VOCABULARY)
            item = Item(text, 0, 0, "subtitulo", "Nome do anexo")
            merged.append(ParagraphBlock(text, "subtitulo", parts[0].page, "norma", item, []))
    extraction.blocks[:] = merged


def merge_across_regions(extraction: Extraction) -> None:
    """Parágrafo que vira de coluna ou de página: termina sem pontuação e o seguinte começa com minúscula.
    Ex.: "... e a empre-" / "sa AVANÇO FACILITIES ..." (Portaria nº 0628/2026-CRG) e o cargo "Presidente do
    Instituto ... do Es-" / "tado do Pará" (Portaria nº 649/2026-IGEPPS), ambos no DOE de 02.10.2026.
    Dentro de uma mesma coluna isso já é feito ao juntar as linhas."""
    blocks, merged = extraction.blocks, []
    for block in blocks:
        previous = merged[-1] if merged else None
        if (isinstance(previous, ParagraphBlock) and isinstance(block, ParagraphBlock)
                and previous.kind == block.kind and block.kind in ("norma", "signer")
                and previous.role in ("texto", "cargo", "ementa")
                and block.role not in ("epigrafe", "publicacao", "anexo", "ordem")    # "resolvem:" fica separado
                and block.text[:1].islower() and not re.match(r"[a-z]\s*\)", block.text)
                and not previous.text.rstrip("\xad ").endswith((".", ":", ";", "!", "?"))):
            head = previous.text.rstrip()
            text = head[:-1] + block.text if head.endswith("\xad") else join_line(head, block.text, _VOCABULARY)
            previous.text = text
            if previous.source_runs is not None and block.source_runs is not None:
                previous.source_runs = previous.source_runs + [(" ", 0)] + block.source_runs
            previous.attention = previous.attention + [a for a in block.attention if a not in previous.attention]
            if previous.item:
                previous.item.text = text
            continue
        merged.append(block)
    extraction.blocks[:] = merged


def link_split_signatures(extraction: Extraction) -> None:
    """Nome no pé de uma coluna (ou página) e cargo no topo da seguinte: cada região é
    classificada em separado, então a regra "nome seguido de cargo" é refeita aqui."""
    paragraphs = [b for b in extraction.blocks]
    for name, cargo in zip(paragraphs, paragraphs[1:]):
        if not (isinstance(name, ParagraphBlock) and isinstance(cargo, ParagraphBlock)
                and name.kind == cargo.kind == "norma" and name.role != "autoridade"
                and looks_like_person_name(name.text) and SIGNER_CARGO.match(cargo.text)):
            continue
        for block, role, reason in ((name, "autoridade", "Assinatura seguida de cargo"),
                                    (cargo, "cargo", "Cargo após assinatura")):
            block.role = role
            block.attention = []   # avisos de quando nome e cargo foram classificados separados
            if block.item:
                block.item.role, block.item.reason = role, reason
                block.item.attention = list(block.attention)


def _row_key(row: list[str]) -> tuple[str, ...]:
    return tuple(" ".join(value.split()).upper() for value in row)


def _title_row(table: TableBlock) -> bool:
    """Primeira linha com texto só na primeira célula: título do quadro ("GTI 1: Assessoria Normativa")."""
    if not table.matrix or len(table.matrix[0]) < 2:
        return False
    first = table.matrix[0]
    return bool(first[0].strip()) and not any(cell.strip() for cell in first[1:]) and len(first[0]) <= 120


def merge_spanning_tables(extraction: Extraction, pdf: Any) -> None:
    """Une fragmentos contíguos ao virar coluna ou página, preservando sua origem."""
    index = 0
    while index + 1 < len(extraction.blocks):
        first, second = extraction.blocks[index:index + 2]
        if not isinstance(first, TableBlock) or not isinstance(second, TableBlock):
            index += 1
            continue
        previous_part = table_page_parts(first)[-1]
        next_part = table_page_parts(second)[0]
        widths_a, widths_b = first.column_widths, second.column_widths
        same_columns = len(widths_a) == len(widths_b) and len(widths_a) > 1
        same_widths = same_columns and all(a > 0 and b > 0 and
            .75 <= (a / sum(widths_a)) / (b / sum(widths_b)) <= 1.25
            for a, b in zip(widths_a, widths_b))
        closing_row = "".join(first.matrix[-1]).strip().endswith(('”', '"'))
        page_continuation = (next_part.page == previous_part.page + 1
                             and previous_part.bbox[3] > pdf.pages[previous_part.page - 1].height - 50
                             and next_part.bbox[1] < 100)
        page = pdf.pages[previous_part.page - 1]
        middle = page.width / 2 if hasattr(page, "width") else 0
        column_continuation = (next_part.page == previous_part.page
                               and previous_part.column == "esquerda"
                               and next_part.column == "direita"
                               and previous_part.bbox[2] < middle + 12
                               and next_part.bbox[0] > middle - 12
                               and previous_part.bbox[3] > page.height - 90   # DOU: corpo até ~1.120 pt
                               and next_part.bbox[1] < 150)
        same_column_continuation = (next_part.page == previous_part.page
                                    and bool(previous_part.column)
                                    and next_part.column == previous_part.column
                                    and 0 <= next_part.bbox[1] - previous_part.bbox[3] <= 7
                                    and abs(next_part.bbox[0] - previous_part.bbox[0]) < 3
                                    and abs(next_part.bbox[2] - previous_part.bbox[2]) < 3)
        if _title_row(first) and _title_row(second):
            index += 1            # dois quadros, cada um com seu título ("GTI 8: Economia Digital" / "GTI 9: ...")
            continue
        last_label = first.matrix[-1][0].strip() if first.matrix[-1] else ""
        finished_total = bool(re.match(r"^TOTA(?:L|IS)\b", last_label, re.I))
        if not ((page_continuation or column_continuation or same_column_continuation)
                and same_widths and not closing_row and not finished_total
                and first.strategy == second.strategy):
            index += 1
            continue
        if (len(second.matrix) > 1 and first.matrix
                and _row_key(second.matrix[0]) == _row_key(first.matrix[0])
                and not any(r0 == 0 for r0, _, _, _ in second.merges)):
            # O diário repete o cabeçalho ao virar a página; o Word já o repete sozinho
            # (Portaria nº 522/2026-SEFA.GS, Anexo II, p. 3 -> 4).
            repeated = second.matrix.pop(0)
            if second.original_matrix:
                second.original_matrix.pop(0)
            second.source_signature = second.source_signature - matrix_signature([repeated])
            second.merges = [(r0 - 1, c0, r1 - 1, c1) for r0, c0, r1, c1 in second.merges]
            second.source_parts = [TablePagePart(p.page, p.bbox, max(0, p.first_row - 1), p.last_row - 1, p.column)
                                   for p in second.source_parts]
            extraction.issues.append(f"Página {next_part.page}: o cabeçalho repetido da tabela foi retirado "
                                     "(no Word, a linha de cabeçalho se repete sozinha a cada página).")
        offset = len(first.matrix)
        old_parts = table_page_parts(first)
        new_parts = table_page_parts(second)
        first.matrix.extend(second.matrix)
        first.original_matrix.extend(second.original_matrix)
        first.source_numbers.update(second.source_numbers)
        first.source_signature.update(second.source_signature)
        first.issues.extend(second.issues)
        first.merges.extend((r0 + offset, c0, r1 + offset, c1)
                            for r0, c0, r1, c1 in second.merges)
        first.source_parts = old_parts + [TablePagePart(part.page, part.bbox,
                                                       part.first_row + offset, part.last_row + offset,
                                                       part.column)
                                          for part in new_parts]
        first.column_widths = [(a * offset + b * len(second.matrix)) / len(first.matrix)
                               for a, b in zip(widths_a, widths_b)]
        first.reviewed = False
        del extraction.blocks[index + 1]


def edition_without_number(pdf_path: Path, document: Any) -> bool:
    """DO-e/SEFA em prévia: "ANO III - Nº - - 5 PÁGINAS" e "DO-e/SEFA Nº: -"."""
    return any(re.search(r"\bN[º°]\s*:?\s*-\s*(?:-|$)", line["text"])
               for line in pdf_lines(pdf_path, 1, document)[:12])


def not_found_message(pdf_path: Path, document: Any, query: str) -> str:
    """Quando nenhuma epígrafe conhecida tem o identificador, diz onde ele aparece, se aparecer num título
    (tipo de ato que o programa ainda não reconhece), em vez de sugerir PDF digitalizado."""
    general = ("Não encontrei uma epígrafe com esse identificador no texto do PDF. "
               "PDF digitalizado ou título diferente requer seleção manual.")
    if not re.search(r"\d", query):
        return general
    for page_number in range(1, len(document) + 1):
        for line in pdf_lines(pdf_path, page_number, document):
            text = line["text"].strip()
            words = text.split()
            if (len(text) <= 160 and words and words[0].isupper() and len(words[0]) > 2
                    and re.search(r"\b" + NUMBER_SIGN + r"\s*\d", text, re.I) and _query_hit(query, text)):
                return (f"O identificador aparece na página {page_number}, no título “{text}”, mas esse tipo de ato "
                        "ainda não é reconhecido pelo programa. Envie o PDF para ajuste.")
    return general


@dataclass
class Candidate:
    """Epígrafe que corresponde ao identificador digitado. Um diário pode ter dois atos com o mesmo
    título (Portaria nº 659, de 07.10.2026, do IGEPPS na p. 61 e da SEFA na p. 85 do DOE de 09.10.2026)."""
    page: int            # 1 = primeira página
    column: str          # "esquerda", "direita" ou "" (página sem colunas)
    epigraph: str
    preview: str         # começo do texto logo abaixo da epígrafe ("O Secretário de Estado da Fazenda, ...")


def _search(pdf_path: Path, document: Any, query: str) -> tuple[list[tuple[int, dict, str]], list[str]]:
    """Epígrafes que correspondem ao identificador, na ordem do diário, e o texto de todas as páginas."""
    hits: list[tuple[int, dict, str]] = []
    by_number: list[tuple[int, dict, str]] = []
    all_text: list[str] = []
    for page_index in range(len(document)):
        page_lines = pdf_lines(pdf_path, page_index + 1, document)
        all_text += [line["text"] for line in page_lines]
        for line_index, line in enumerate(page_lines):
            if not EPIGRAPH.match(line["text"].strip()):
                continue
            above = line_above(page_lines, line_index)
            if above and MID_SENTENCE.search(above["text"]):
                continue
            complete = epigraph_text(page_lines, line_index)
            if _query_hit(query, complete):
                hits.append((page_index, line, complete))
            elif _same_act(query, complete):   # "024/2026" acha "... Nº 024 - / SEFA.GS, DE ... 2026"
                by_number.append((page_index, line, complete))
    return hits or by_number, all_text


def _preview(lines: list[dict], heading: dict, epigraph: str, size: int = 170) -> str:
    """Primeiras linhas abaixo da epígrafe, na mesma coluna."""
    below = sorted((l for l in lines if l["top"] > heading["top"] + 1
                    and l["x0"] < heading["x1"] and heading["x0"] < l["x1"]), key=lambda l: l["top"])
    text = ""
    for line in below[:6]:
        value = line["text"].strip()
        if not text and value in epigraph:          # continuação da epígrafe ("SEFA.GS, DE ... 2026")
            continue
        text = join_line(text, value) if text else value
        if len(text) >= size:
            break
    return text if len(text) <= size else text[:size].rsplit(" ", 1)[0] + "…"


def find_candidates(pdf_path: Path, query: str) -> list[Candidate]:
    """Todos os atos do PDF que correspondem ao identificador, para escolher quando houver mais de um."""
    if not query.strip():
        raise ValueError("Informe um trecho único do título do ato, como 494/2026-SEFA.GS.")
    with fitz.open(str(pdf_path)) as document:
        chosen, _ = _search(pdf_path, document, query)
        result = []
        for page_index, line, complete in chosen:
            page_lines = pdf_lines(pdf_path, page_index + 1, document)
            width = document[page_index].rect.width
            middle = (line["x0"] + line["x1"]) / 2
            column = "" if line["x1"] - line["x0"] > width * .6 else ("esquerda" if middle < width / 2 else "direita")
            result.append(Candidate(page_index + 1, column, complete, _preview(page_lines, line, complete)))
    return result


SUGGESTED = re.compile(r"\bFAZENDA\b|\bSEFA\b")


def suggested_candidate(candidates: list[Candidate]) -> int:
    """Posição (a partir de 0) a deixar marcada na lista: o único ato da Secretaria da Fazenda, se houver
    exatamente um entre os candidatos; senão, o primeiro."""
    sefa = [i for i, c in enumerate(candidates) if SUGGESTED.search(normalized(c.epigraph + " " + c.preview))]
    return sefa[0] if len(sefa) == 1 else 0


def extract_act(pdf_path: Path, query: str, max_pages: int = 12, occurrence: int | None = None) -> Extraction:
    """occurrence: qual dos atos com esse identificador importar (1 = o primeiro do diário), quando houver
    mais de um; sem ele, vale o primeiro, com aviso."""
    if not query.strip():
        raise ValueError("Informe um trecho único do título do ato, como 494/2026-SEFA.GS.")
    if not 1 <= max_pages <= 100:
        raise ValueError("O limite deve estar entre 1 e 100 páginas.")
    with pdfplumber.open(pdf_path) as pdf, fitz.open(str(pdf_path)) as fitz_pdf:
        source_name = detect_publication(pdf, fitz_pdf)
        issue_date = pdf_issue_date(pdf_path, fitz_pdf) if source_name else None
        extra_edition = bool(source_name) and is_extra_edition(pdf_path, fitz_pdf)
        global _VOCABULARY
        chosen, all_text = _search(pdf_path, fitz_pdf, query)
        _VOCABULARY = vocabulary(all_text)
        if not chosen:
            raise ValueError(not_found_message(pdf_path, fitz_pdf, query))
        if occurrence is not None and not 1 <= occurrence <= len(chosen):
            raise ValueError(f"O identificador corresponde a {len(chosen)} ato(s) no PDF; "
                             f"escolha um número de 1 a {len(chosen)}.")
        picked = chosen[(occurrence or 1) - 1]
        start_index, heading = picked[:2]
        opinion = bool(OPINION.match(picked[2]))
        extraction = Extraction(pdf_path, query, [], [])
        others = [c for c in chosen if c is not picked]
        if others:
            # Também quando os títulos são idênticos (antes, nesse caso, não havia aviso).
            def describe(page_index: int, line: dict, epigraph: str) -> str:
                if normalized(epigraph) != normalized(picked[2]):
                    return f"p. {page_index + 1}: {epigraph[:80]}"
                start = _preview(pdf_lines(pdf_path, page_index + 1, fitz_pdf), line, epigraph, 70)
                return f"p. {page_index + 1}, que começa por “{start}”"
            lista = "; ".join(describe(*other) for other in others[:3])
            position = chosen.index(picked) + 1
            extraction.issues.append(
                f"O identificador corresponde a {len(chosen)} atos no PDF; foi usado o {position}º, da "
                f"p. {start_index + 1} (outros: {lista}{'; ...' if len(others) > 3 else ''}). "
                + ("" if occurrence else "Se não for o desejado, digite o identificador completo, com a sigla do "
                   "órgão (ex.: 04/2026-PGE), ou escolha o ato na lista da janela."))
        if re.search(r"PREVIS(?:TA|AO)", normalized(pdf_path.name)):     # "Prevista", "2026-10-09. Previsão.pdf"
            extraction.issues.append("O nome do PDF indica edição prevista; confirme a publicação final antes de usar o ato.")
        elif source_name == "DO-e/SEFA" and edition_without_number(pdf_path, fitz_pdf):
            extraction.issues.append("A edição não tem número (\"Nº -\"), o que indica prévia; confirme a publicação "
                                     "final antes de usar o ato.")
        found_end = signed_off = ended_by_next = False
        case_fixed: dict[int, list[tuple[float, float]]] = {}
        for page_index in range(start_index, min(start_index + max_pages, len(pdf.pages))):
            page = pdf.pages[page_index]
            page_number = page_index + 1
            blocks_before_page = len(extraction.blocks)
            all_lines = pdf_lines(pdf_path, page_number, fitz_pdf)
            fix_char_case(page, fitz_pdf[page_index], case_fixed.setdefault(page_number, []))
            # Tabelas e colunas são reconhecidas na página como está; o texto das células e as palavras vêm da
            # página sem os pontos brancos e sem os espaços cobertos (ATC RFB/CGIBS nº 9). Reconhecer as tabelas
            # na página filtrada mudava tabelas de outros atos do mesmo diário.
            clean = visible_page(page)
            all_tables = page.find_tables(table_settings=open_top_settings(page))
            for table in all_tables:
                table.page = clean
            all_words = clean.extract_words(x_tolerance=.5, y_tolerance=1.5, keep_blank_chars=False)
            split_columns = has_two_columns(page, all_lines, all_tables)
            if three_columns(all_lines, page.width):
                extraction.blocking_issues.append(
                    f"A página {page_number} tem três colunas de texto, formato que o programa ainda não lê "
                    "(só uma ou duas colunas). Envie este PDF para ajuste.")
            dou = source_name == "DOU"
            top_margin, bottom_margin = dou_margins(all_lines, page.height) if dou else (0.0, page.height)
            page_header = None if dou else page_top_margin(all_lines, page.height)
            middle = page.width / 2
            heading_left = (heading["x0"] + heading["x1"]) / 2 < middle
            columns = ([heading_left] if page_index == start_index and not heading_left else [True, False])
            if not split_columns:
                columns = [None]
            for column in columns:
                first_region = page_index == start_index and (column is None or column == heading_left)
                if column is None:
                    text_lines, page_tables, words = all_lines, all_tables, all_words
                else:
                    text_lines = [l for l in all_lines if
                                  (((l["x0"] + l["x1"]) / 2 < middle) == column)]
                    page_tables = [t for t in all_tables if
                                   (((t.bbox[0] + t.bbox[2]) / 2 < middle) == column)]
                    words = [w for w in all_words if
                             (((w["x0"] + w["x1"]) / 2 < middle) == column)]
                start_y = (heading["top"] - .5 if first_region else top_margin if dou and top_margin
                           else page_header if page_header is not None else 52.0)
                if not first_region and column is not None:
                    # Grades logo abaixo do cabeçalho do diário podem começar
                    # acima de 52 pt (p. ex. 42,6 pt no DOE de 01.10.2026).
                    early = [t.bbox[1] for t in page_tables
                             if 35 <= t.bbox[1] < start_y and
                             (t.bbox[2] < middle + 12 if column else t.bbox[0] > middle - 12)]
                    if early:
                        start_y = min(start_y, max(35.0, min(early) - 1))
                stop_y = min(page.height - 25, bottom_margin)
                if dou and not first_region and signed_off:
                    # DOU: depois da assinatura, o ato só continua se a região seguinte começar por ANEXO
                    first_line = next((l["text"].strip() for l in text_lines
                                       if start_y < l["top"] < stop_y and l["text"].strip()), "")
                    if not ANNEX.match(first_line):
                        found_end = True
                        extraction.issues.append("O fim foi identificado pela assinatura; confirme a delimitação.")
                        break
                previous_text = ""
                for line in text_lines:
                    if line["top"] <= start_y + (8 if first_region else 0):
                        continue
                    text = line["text"].strip()
                    before_text = previous_text
                    after_mid_sentence, previous_text = bool(MID_SENTENCE.search(previous_text)), text
                    if PROTOCOL.match(text):
                        stop_y = min(stop_y, line["top"] - 1)
                        extraction.protocol = text
                        found_end = True
                        break
                    # Nos 45 pt logo abaixo da epígrafe, uma linha que começa por "Decreto n. 2.703, ..." ou
                    # "PORTARIA nº 107/2026, ..." costuma ser citação; o ato seguinte só começa ali se o tipo estiver
                    # em maiúsculas e a linha anterior fechar a frase (Resoluções SENAI nº 09 a 20, DOE de 05.10.2026).
                    next_act_close = (first_region and EPIGRAPH.match(text) and line["top"] > heading["top"] + 2
                                      and text.split()[0].isupper() and len(text.split()[0]) > 2
                                      and before_text.rstrip().endswith((".", ";", ":")))
                    if (SECTION_HEAD.match(text) or next_act_close or
                            (EPIGRAPH.match(text) and line["top"] > start_y + (45 if first_region else 8)
                             and not after_mid_sentence and not cited_in_text(text))):
                        stop_y = min(stop_y, line["top"] - 1)
                        found_end = ended_by_next = True
                        extraction.issues.append("O fim foi identificado pelo início do ato ou seção seguinte; confira a delimitação.")
                        break
                start_block = len(extraction.blocks)
                publication = (publication_note(pdf_path, heading["text"], source_name, issue_date, extra_edition)
                               if first_region else None)
                append_region(page, pdf_path, extraction, text_lines, page_tables, words,
                              start_y, stop_y, page_number, publication, split_columns, opinion, first_region)
                if column is False and start_block and len(extraction.blocks) > start_block:
                    previous = extraction.blocks[start_block - 1]
                    current = extraction.blocks[start_block]
                    if (isinstance(previous, ParagraphBlock) and isinstance(current, ParagraphBlock)
                            and previous.kind == current.kind == "norma"
                            and previous.role == "ementa" and current.text[:1].islower()
                            and not previous.text.rstrip().endswith((".", ";", ":"))):
                        previous.text += " " + current.text
                        if previous.source_runs is not None and current.source_runs is not None:
                            previous.source_runs = previous.source_runs + [(" ", 0)] + current.source_runs
                        if previous.item:
                            previous.item.text = previous.text
                        del extraction.blocks[start_block]
                if dou:
                    # O quadro de divulgação depois da assinatura não pode impedir que o ato termine ali.
                    drop_dou_house_ads(extraction, pdf)
                last_block = extraction.blocks[-1] if extraction.blocks else None
                signed_off = isinstance(last_block, ParagraphBlock) and last_block.role == "cargo"
                if found_end:
                    break
                if column is True and split_columns and extraction.blocks and not dou:
                    last = extraction.blocks[-1]
                    signers = sum(isinstance(b, ParagraphBlock) and b.role == "autoridade" for b in extraction.blocks)
                    if (isinstance(last, ParagraphBlock) and last.role == "cargo"
                            and not last.text.rstrip().endswith(("-", "\xad"))     # "... do Es-" / "tado do Pará"
                            and not MID_SENTENCE.search(last.text)
                            # Solução de Consulta: depois do parecerista vêm o despacho e quem aprova
                            and (not opinion or signers >= 2)):
                        found_end = True
                        extraction.issues.append("O fim foi identificado pela assinatura na coluna esquerda; confirme a delimitação.")
                        break
            if page_index == start_index or len(extraction.blocks) > blocks_before_page:
                extraction.pages.append(page_number)
            if found_end:
                break
        if not found_end:
            extraction.issues.append("Não foi encontrado protocolo ou início do próximo ato. Confira onde o ato termina.")
            if start_index + max_pages < len(pdf.pages):
                extraction.blocking_issues.append(
                    f"A busca parou após {max_pages} páginas sem encontrar o fim do ato. "
                    "Aumente o limite antes de exportar.")
        merge_spanning_tables(extraction, pdf)
        case_pages = sorted({part.page for table in extraction.tables for part in table_page_parts(table)
                             if any(part.bbox[0] <= x <= part.bbox[2] and part.bbox[1] <= y <= part.bbox[3]
                                    for x, y in case_fixed.get(part.page, ()))})
        if case_pages:
            extraction.issues.append(
                "Na leitura das tabelas, algumas letras vinham com maiúscula/minúscula trocada por causa da fonte do "
                "PDF; a grafia das células foi acertada pelo texto da página "
                f"(p. {', '.join(map(str, case_pages))}). Confira as tabelas.")
        merge_across_regions(extraction)
        merge_annex_subtitles(extraction)
        link_split_signatures(extraction)
        if ended_by_next:
            trim_trailing_headings(extraction, source_name == "DOU")
        if extraction.quote_state.get("in_quote"):
            last = next((b for b in reversed(extraction.blocks)
                         if isinstance(b, ParagraphBlock) and b.role == "redacao_nova"), None)
            if last is not None:
                last.attention.append("Aspas de redação nova não foram fechadas no texto fornecido.")
        if source_name == "DOU":
            add_cgibs_availability(extraction)
        date_constitutional_amendment(extraction)
        return extraction


def set_cell_border(cell: Any, edge: str, color: str = "595959", size: int = 4) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    borders = tcPr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tcPr.append(borders)
    name = qn("w:" + edge)
    rule = borders.find(name)
    if rule is None:
        rule = OxmlElement("w:" + edge)
        borders.append(rule)
    rule.set(qn("w:val"), "single")
    rule.set(qn("w:sz"), str(size))
    rule.set(qn("w:color"), color)


def set_cell_margins(cell: Any) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    mar = tcPr.first_child_found_in("w:tcMar")
    if mar is None:
        mar = OxmlElement("w:tcMar")
        tcPr.append(mar)
    for side in ("top", "left", "bottom", "right"):
        el = mar.find(qn("w:" + side))
        if el is None:
            el = OxmlElement("w:" + side)
            mar.append(el)
        el.set(qn("w:w"), str(4 if side in ("top", "bottom") else 22))
        el.set(qn("w:type"), "dxa")


CELL_LIST_START = re.compile(r"^(?:[a-zA-Z]\)|[IVXLCDM]+\s*[-–]|\d+\s*[.)–-]|[-–•·])\s")


def unwrap_cell(value: str) -> str:
    """Desfaz as quebras de linha do diário dentro de uma célula de texto ("CHARLES JOHNSON DA" /
    "SILVA ALCANTARA", "DIRETORIA DE" / "FISCALIZAÇÃO – DFI"): a coluna do Word tem outra largura.
    Linhas com algarismos, itens de lista e linhas depois de ":" ou ";" continuam separadas."""
    lines = value.split("\n")
    joined = [lines[0]]
    for line in lines[1:]:
        previous = joined[-1]
        if (re.search(r"[A-Za-zÀ-ÿ]", previous) and re.search(r"[A-Za-zÀ-ÿ]", line)
                and not re.search(r"\d", previous + line) and not CELL_LIST_START.match(line.strip())
                and not previous.rstrip().endswith((":", ";"))):
            joined[-1] = join_line(previous, line.strip())
        else:
            joined.append(line)
    return "\n".join(joined)


def formatted_table_matrix(block: TableBlock, normalize_municipalities: bool = True
                           ) -> tuple[list[list[str]], list[tuple[int, int, str, str, str]]]:
    """Separa a grafia de apresentação da matriz auditada do PDF.

    Somente colunas explicitamente identificadas como Município recebem iniciais
    maiúsculas. Códigos, produtos, siglas e demais colunas permanecem intactos.
    Tabelas com cabeçalhos mesclados exigem escolha manual do campo apropriado.
    """
    matrix = [[unwrap_cell(value) for value in row] for row in block.matrix]
    changes = []
    if not normalize_municipalities or not matrix or block.merges:
        return matrix, changes
    columns = [ci for ci, value in enumerate(matrix[0])
               if re.fullmatch(r"(?:NOME(?: D[OE]S?)? )?MUNICIPIOS?", normalized(value))]
    if not columns:
        return matrix, changes
    for ri, row in enumerate(matrix):
        for ci, before in enumerate(row):
            if ri == 0:
                # Não muda símbolos ou unidades de outros campos do cabeçalho.
                label = normalized(before)
                if ci in columns or label.startswith(("INDICE", "INDICADOR DE QUALIDADE")):
                    after = " ".join(before.split()).upper()
                    rule = "Cabeçalho em maiúsculas e sem quebra manual de linha"
                else:
                    continue
            elif ci in columns and re.fullmatch(r"[^\W\d_]+(?:(?:\.\s*|[\s'’\-])[^\W\d_]+)*\.?", before):
                # abreviaturas: "CONC. ARAGUAIA", "BREJO GRAN.ARAGUAIA" (Portaria nº 659/2026-SEFA)
                if normalized(before) in {"TOTAL", "TOTAIS"}:
                    continue
                after = re.sub(r"(^|[\s.-])([^\W\d_])",
                               lambda m: m[1] + m[2].upper(), before.lower())
                rule = "Iniciais maiúsculas no nome do município, conforme o exemplo revisado"
            else:
                continue
            if before != after:
                if re.sub(r"\s", "", before).casefold() != re.sub(r"\s", "", after).casefold():
                    raise ValueError("Um ajuste de grafia tentou alterar o conteúdo da célula.")
                row[ci] = after
                changes.append((ri, ci, before, after, rule))
    return matrix, changes


def add_table(document: Document, block: TableBlock, normalize_municipalities: bool = True) -> None:
    matrix, _ = formatted_table_matrix(block, normalize_municipalities)
    table = document.add_table(rows=len(matrix), cols=len(matrix[0]))
    table.autofit = False
    section = document.sections[-1]
    usable_cm = Emu(section.page_width - section.left_margin - section.right_margin).cm - .1
    source_widths = block.column_widths[:]
    dense = len(source_widths) >= 7
    if dense:
        # O diário pode atribuir nove colunas de largura idêntica a nomes,
        # códigos longos e valores de quatro dígitos. Distribua o espaço
        # conforme a palavra mais longa de cada coluna, sem alterar células.
        for ci, original in enumerate(source_widths):
            longest = max((len(part) for row in block.matrix
                           for part in re.findall(r"[^\s]+", row[ci])
                           if re.search(r"\w", part)), default=7)
            source_widths[ci] = original * max(.68, min(1.55, (longest + 2) / 9))
    width_sum = sum(source_widths)
    widths = [Cm(usable_cm * w / width_sum) for w in source_widths]
    for col, width in zip(table.columns, widths):
        col.width = width
    header_rows = max((r1 + 1 for r0, _, r1, _ in block.merges if r0 == 0), default=1)
    dotted_header = bool(dense and all(re.fullmatch(r'[“”".\s…]+', value)
                                       for value in block.matrix[0]))
    if dotted_header:
        header_rows = 0
    if (len(block.matrix) > 1 and any(r0 == 0 and c0 == 0 and c1 == len(widths) - 1
                                      for r0, c0, _, c1 in block.merges)
            and not monetary_values([block.matrix[1]])):
        header_rows = max(header_rows, 2)
    # Dois títulos em linha inteira ("GTI 7: Serviços Financeiros ..." / "Natureza Mista ...") e depois os
    # nomes das colunas ("TITULAR | SUPLENTE | REPRESENTAÇÃO", ATC RFB/CGIBS nº 9): os nomes das colunas também
    # são cabeçalho, como nas tabelas com um só título ("Coordenação Normativa" / "TITULAR | ...").
    titles = 0
    while titles < len(block.matrix) and any(r0 == titles == r1 and c0 == 0 and c1 == len(widths) - 1
                                             for r0, c0, r1, c1 in block.merges):
        titles += 1
    if (titles >= 2 and titles < len(block.matrix) - 1
            and all(value.strip() and value == value.upper() and not re.search(r"\d", value)
                    for value in block.matrix[titles])):
        header_rows = max(header_rows, titles + 1)
    for heading_row in table.rows[:header_rows]:
        header = OxmlElement("w:tblHeader")
        header.set(qn("w:val"), "true")
        heading_row._tr.get_or_add_trPr().append(header)
    available = paragraph_style_names(document)
    font_size = 8.0 if len(widths) <= 3 else (6.2 if len(widths) == 4 else
                                              (5.9 if len(widths) <= 6 else 6.0))
    for row_index, (row, values) in enumerate(zip(table.rows, matrix)):
        is_heading = row_index < header_rows or (not any(MONEY.fullmatch(v) for v in values[1:]) and re.match(r"^\d+\s*[-–]", values[0]))
        is_total = values[0].strip().upper() == "TOTAIS"
        is_major = bool(re.match(r"^\d+\s*[-–]", values[0]))
        cant_split = OxmlElement("w:cantSplit")
        row._tr.get_or_add_trPr().append(cant_split)
        for col_index, (cell, value) in enumerate(zip(row.cells, values)):
            cell.width = widths[col_index]
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_margins(cell)
            for edge in ("left", "right", "top", "bottom"):
                set_cell_border(cell, edge, color="595959" if row_index == 0 and edge != "bottom" else "000000")
            cell.text = ""
            p = cell.paragraphs[0]
            compact = "".join(value.split())
            alignment = (WD_ALIGN_PARAGRAPH.CENTER if len(widths) <= 3 or is_heading else
                         WD_ALIGN_PARAGRAPH.LEFT if col_index == 0 else
                           WD_ALIGN_PARAGRAPH.CENTER if is_heading or
                           (dense and not re.fullmatch(r"(?:R\$)?[\d.,/%-]+", compact)) else
                           WD_ALIGN_PARAGRAPH.RIGHT)
            suffix = {WD_ALIGN_PARAGRAPH.LEFT: "esquerda", WD_ALIGN_PARAGRAPH.CENTER: "centro",
                      WD_ALIGN_PARAGRAPH.RIGHT: "direita"}[alignment]
            style = f"a50_tab_{suffix}_8" if len(widths) <= 3 else f"a60_tab_{suffix}_6"
            p.style = style if style in available else "Normal"
            if style not in available:
                p.alignment = alignment
                p.paragraph_format.space_before = Pt(1)
                p.paragraph_format.space_after = Pt(1)
            if len(widths) > 3:
                p.paragraph_format.line_spacing = .95
            if is_heading or (len(matrix) <= 6 and row_index < len(matrix) - 1):
                p.paragraph_format.keep_with_next = True
            for line_index, line in enumerate(value.split("\n")):
                if line_index:
                    p.add_run().add_break()
                run = p.add_run(line)
                if style not in available:
                    run.font.name = "Arial"
                    run.font.size = Pt(font_size)
                elif len(widths) > 3 and font_size != 6.0:
                    run.font.size = Pt(font_size)
                if is_heading or is_total or is_major:
                    run.bold = True
    for first_row, first_col, last_row, last_col in block.merges:
        merged = table.cell(first_row, first_col).merge(table.cell(last_row, last_col))
        if last_row > first_row:
            merged.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
        for paragraph in list(merged.paragraphs)[1:]:
            if not paragraph.text.strip():
                paragraph._element.getparent().remove(paragraph._element)
    document.add_paragraph().paragraph_format.space_after = Pt(0)


def build_word(extraction: Extraction, model_path: Path, output_path: Path, overwrite: bool = False,
               normalize_municipalities: bool = True) -> None:
    if output_path.resolve() == model_path.resolve():
        raise ValueError("Escolha um nome de saída diferente do documento modelo.")
    if output_path.exists() and not overwrite:
        raise FileExistsError(f"O arquivo já existe: {output_path}")
    if any(table.issues and not table.reviewed for table in extraction.tables):
        raise ValueError("Há tabela com associação incerta. Confira-a com o PDF e marque a revisão antes de exportar.")
    model = Document(str(model_path))
    for block in extraction.paragraphs:
        if block.kind == "annex_head" and ANNEX.match(block.text):
            block.role = "anexo"
        if block.kind in {"norma", "signer", "date"} or block.role == "anexo":
            item = block.item or Item(block.text, 0, 0, block.role, "Importado do PDF")
            item.text, item.role = block.text, block.role
            apply_styles([item], model)
            block.item = item
    items = [b.item for b in extraction.paragraphs if b.item]
    spacer = ementa_spacer_before(items) if "a03_ementa" in paragraph_style_names(model) else None
    clear_model_contents(model)
    expected: list[str] = []          # texto de cada parágrafo do corpo, para conferir o arquivo salvo
    for block_index, block in enumerate(extraction.blocks):
        if isinstance(block, TableBlock):
            add_table(model, block, normalize_municipalities)
            expected.append("")       # add_table deixa um parágrafo vazio depois da tabela
        elif isinstance(block, ImageBlock):
            expected.append("")
            p = model.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if block.width else WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.space_before = Pt(0)
            # trecho do ato em imagem: ampliado para a letra ficar perto da do texto do Word (o DOU usa corpo
            # menor), sem passar da largura da página
            section = model.sections[-1]
            available = section.page_width - section.left_margin - section.right_margin
            width = min(Pt(block.width * 1.4), available) if block.width else Inches(.46)
            p.add_run().add_picture(io.BytesIO(block.data), width=width)
        elif block.kind in {"norma", "signer", "date"} or block.role == "anexo":
            if block.item is spacer:
                model.add_paragraph(style="a03_ementa")
                expected.append("")
            p = add_formatted_paragraph(model, block.item, block_runs(block))
            expected.append(block.item.text)
            if (block_index + 1 < len(extraction.blocks)
                    and isinstance(extraction.blocks[block_index + 1], TableBlock)):
                p.paragraph_format.keep_with_next = True
            if block.kind == "date":
                p.paragraph_format.space_before = Pt(7)
            if block.item.reason == "Mesa que assina (Emenda Constitucional)":
                p.paragraph_format.keep_with_next = True     # "Mesa da Câmara dos Deputados" junto do 1º nome
        else:
            p = model.add_paragraph()
            p.style = "Normal"
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if block.kind == "unit" else (
                WD_ALIGN_PARAGRAPH.CENTER if block.kind == "annex_head" else WD_ALIGN_PARAGRAPH.LEFT
            )
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.space_before = Pt(0)
            run = p.add_run(block.text)
            expected.append(block.text)
            run.font.name = "Arial"
            run.font.size = Pt(7 if block.kind == "annex_head" else 6.4)
            run.bold = block.kind == "annex_head"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(prefix="sefa_pdf_", suffix=".docx", dir=output_path.parent, delete=False) as temp:
        temporary = Path(temp.name)
    try:
        model.save(temporary)
        check = Document(temporary)
        produced = [p.text for p in check.paragraphs]
        if produced != expected:
            where = next((i for i, (a, b) in enumerate(zip(expected, produced)) if a != b), min(len(expected), len(produced)))
            raise RuntimeError("A conferência do Word produzido encontrou um parágrafo diferente do texto extraído "
                               f"(parágrafo {where + 1}: “{(expected[where] if where < len(expected) else '')[:80]}”); "
                               "o arquivo de destino não foi alterado.")
        if len(check.tables) != len(extraction.tables):
            raise RuntimeError("A saída não preservou todas as tabelas.")
        for expected, actual in zip(extraction.tables, check.tables):
            if len(expected.matrix) != len(actual.rows):
                raise RuntimeError("A saída perdeu linhas da tabela.")
            covered = {(ri, ci) for r0, c0, r1, c1 in expected.merges
                       for ri in range(r0, r1 + 1) for ci in range(c0, c1 + 1)
                       if (ri, ci) != (r0, c0)}
            presented, _ = formatted_table_matrix(expected, normalize_municipalities)
            for ri, row in enumerate(presented):
                for ci, value in enumerate(row):
                    if (ri, ci) not in covered and actual.cell(ri, ci).text != value:
                        raise RuntimeError("A saída divergiu de uma célula da tabela.")
        os.replace(temporary, output_path)
    finally:
        temporary.unlink(missing_ok=True)


def write_pdf_report(extraction: Extraction, report_path: Path, normalize_municipalities: bool = True) -> None:
    sections = []
    for block in extraction.blocks:
        if isinstance(block, ParagraphBlock):
            attention = "<br>".join(html.escape(s) for s in block.attention)
            sections.append(f'<p><small>Página {block.page} · {html.escape(block.kind)} · {html.escape(block.role)}</small><br>{html.escape(block.text)}'
                            + (f'<br><strong>Conferir: {attention}</strong>' if attention else '') + '</p>')
        elif isinstance(block, ImageBlock):
            what = ("trecho do ato publicado como imagem, incluído como imagem no Word; confira"
                    if block.width else "imagem do anexo incorporada")
            sections.append(f'<p><small>Página {block.page} · {what}</small></p>')
        else:
            presented, adjustments = formatted_table_matrix(block, normalize_municipalities)
            values = sum(monetary_values(block.matrix).values())
            differences = list(block.issues)
            if matrix_signature(block.matrix) != block.source_signature:
                differences.append("O texto do Word difere dos caracteres presentes nesta região do PDF.")
            if monetary_values(block.matrix) != block.source_numbers:
                differences.append("Há divergência entre os valores monetários do PDF e do Word.")
            if block.matrix != block.original_matrix:
                differences.append("Algumas células foram corrigidas na interface.")
            if differences:
                status = ("Conferência manual declarada. " if block.reviewed else "Revisão necessária. ") + "; ".join(differences)
            else:
                status = "Texto e valores da extração preservados; confira a associação entre as linhas no PDF"
            table_rows = []
            for row in presented:
                table_rows.append('<tr>' + ''.join(f'<td>{html.escape(value).replace(chr(10), "<br>")}</td>' for value in row) + '</tr>')
            parts = table_page_parts(block)
            source_pages = list(dict.fromkeys(part.page for part in parts))
            page_label = f"página {source_pages[0]}" if len(source_pages) == 1 else (
                "páginas " + " e ".join(map(str, source_pages)))
            locations = "; ".join(
                f"página {part.page}" + (f", coluna {part.column}" if part.column else "") +
                f", linhas {part.first_row + 1} a {part.last_row}"
                for part in parts)
            adjustment_rows = ''.join(
                '<tr>' + ''.join(f'<td>{html.escape(str(v))}</td>' for v in
                                (ri + 1, ci + 1, before, after, rule)) + '</tr>'
                for ri, ci, before, after, rule in adjustments)
            adjustment_report = (f'<details><summary>{len(adjustments)} ajustes automáticos de grafia '
                                 '(maiúsculas/minúsculas e quebras de cabeçalho)</summary>'
                                 '<p>Os números, acentos e a ordem das linhas foram preservados. '
                                 'Cada alteração abaixo compara a matriz extraída com o texto gravado no Word.</p>'
                                 '<table><tr><th>Linha</th><th>Coluna</th><th>Extraído do PDF</th>'
                                 '<th>No Word</th><th>Regra</th></tr>' + adjustment_rows + '</table></details>') if adjustments else ''
            sections.append(f'<h2>Tabela em {html.escape(page_label)} · detecção por {html.escape(block.strategy)}</h2><p>{len(block.matrix)} linhas, {len(block.matrix[0])} colunas, '
                            f'{values} valores monetários. Origem: {html.escape(locations)}. '
                            f'{html.escape(status)}.</p>{adjustment_report}<table>{"".join(table_rows)}</table>')
    issues = "".join(f'<li>{html.escape(issue)}</li>' for issue in extraction.issues + extraction.blocking_issues)
    pelo_link = bool(extraction.source_url) or not str(extraction.pdf_path).strip(". ")
    origem_titulo = "Conferência da importação pelo link do DOU" if pelo_link else "Conferência da importação do PDF"
    origem_link = (f'Página do ato: <a href="{html.escape(extraction.source_url)}">{html.escape(extraction.source_url)}</a>. '
                   if extraction.source_url else "")
    page = f'''<!doctype html><html lang="pt-BR"><meta charset="utf-8"><title>Conferência do PDF SEFA</title>
<style>body{{font:14px/1.45 Arial,sans-serif;max-width:1400px;margin:32px auto;padding:0 22px;color:#222}}
table{{border-collapse:collapse;width:100%;font-size:12px;margin-bottom:25px}}
td{{border:1px solid #aaa;padding:4px;vertical-align:top;overflow-wrap:anywhere}}
td:not(:first-child){{text-align:right}}small{{color:#555}}h1{{font-size:22px}}h2{{font-size:17px}}</style>
<h1>{origem_titulo}</h1>
<p>{origem_link}Identificador: {html.escape(extraction.query)} · páginas {', '.join(map(str, extraction.pages))}.
Trechos de texto entre tabelas verificados: {extraction.intertable_segments_checked}.
O conteúdo das tabelas é editável no Word. Confira a correspondência entre rótulos, números e o PDF antes do uso.</p>
<p>Padronização da grafia de municípios: {'ativada' if normalize_municipalities else 'desativada'}.
Somente campos identificados pelo cabeçalho Município recebem iniciais maiúsculas.
Os ajustes de apresentação aparecem em uma lista própria, sem substituir a matriz usada na auditoria.</p>
<ul>{issues}</ul>{''.join(sections)}
<p>Marcador de término identificado: {html.escape(extraction.protocol or 'não encontrado')}.
Conferência manual dos limites: {'sim' if extraction.limits_reviewed else 'não'}.</p></html>'''
    report_path.write_text(page, encoding="utf-8")
