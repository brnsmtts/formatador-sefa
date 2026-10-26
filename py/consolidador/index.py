from __future__ import annotations
from dataclasses import dataclass, field
from collections import Counter, defaultdict
import re
from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from .model import Target, norm


NUM = r"(\d+)(?:[º°o])?(?:[-–]([A-Z]{1,3})(?![A-Za-zÀ-ÿ]))?"
ARTICLE = re.compile(r"^Art(?:igo)?\.?\s*" + NUM + r"\s*\.?\s*(?:[-–]\s*)?", re.I)
PARAGRAPH = re.compile(r"^§\s*" + NUM + r"\s*\.?\s*(?:[-–]\s*)?", re.I)
UNIQUE = re.compile(r"^Par[aá]grafo\s+[uú]nico\s*[.:-]?\s*", re.I)
INCISO = re.compile(r"^([IVXLCDM]+(?:-[A-Z])?)\s*[-–—]\s*")
ALINEA = re.compile(r"^([a-z])\s*\)\s*")
ITEM = re.compile(r"^(\d+(?:\.\d+)*)\s*[.)]\s+")
HEADING = re.compile(r"^(?:T[IÍ]TULO|CAP[IÍ]TULO|SE[CÇ][AÃ]O|SUBSE[CÇ][AÃ]O|LIVRO|PARTE)\b", re.I)
HIST_TEXT = re.compile(r"^Reda[cç][aã]o\s+(?:original|anterior)\b", re.I)
NOTE_TEXT = re.compile(r"^(?:Reda[cç][aã]o\s+dada|Acrescid[oa]s?|Revogad[oa]s?|Renumerad[oa]s?|Vide\b)", re.I)


_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_R, _HL, _T, _TAB, _BR, _CR, _NBH, _PTAB = (_W + x for x in ("r", "hyperlink", "t", "tab", "br", "cr", "noBreakHyphen", "ptab"))


def _run_text(run):
    out = []
    for e in run:
        tag = e.tag
        if tag == _T:
            out.append(e.text or "")
        elif tag in (_TAB, _PTAB):
            out.append("\t")
        elif tag == _BR:
            out.append("\n" if e.get(_W + "type", "textWrapping") == "textWrapping" else "")
        elif tag == _CR:
            out.append("\n")
        elif tag == _NBH:
            out.append("-")
    return "".join(out)


def paragraph_text(el):
    """Mesmo resultado de Paragraph.text, sem consultas XPath por parágrafo."""
    out = []
    for child in el:
        if child.tag == _R:
            out.append(_run_text(child))
        elif child.tag == _HL:
            out.extend(_run_text(r) for r in child if r.tag == _R)
    return "".join(out)


def label(text):
    text = text.strip().lstrip('“”"').strip()
    for level, pattern in ((0, ARTICLE), (1, PARAGRAPH), (1, UNIQUE), (2, INCISO), (3, ALINEA), (4, ITEM)):
        m = pattern.match(text)
        if m:
            if pattern is UNIQUE:
                number = "UNICO"
            elif level < 2:
                number = str(int(m[1])) + ("-" + m[2].upper() if m[2] else "")
            else:
                number = m[1]
            return level, number, m.end(), text[:m.end()].strip()
    return None


def role(paragraph, style_name=None, color=None, text=None):
    if style_name is None:
        style_name = paragraph.style.name
        color = paragraph.style.font.color.rgb
    style = norm(style_name).replace(" ", "")
    text = (paragraph.text if text is None else text).strip()
    if any(s in style for s in ("rdant_remiss", "remiss_red_ant")):
        return "old_note"
    if any(s in style for s in ("a12_rdant", "a23_redacao_ant", "tabela_centroant", "tabela_esquerdant", "tabela_direitant")):
        return "old"
    if any(s in style for s in ("nr_remiss", "remiss_nova_red")):
        return "new_note"
    if HIST_TEXT.match(text):
        return "old_note"
    if NOTE_TEXT.match(text):
        return "new_note"
    # Há modelos que aplicam a01_epigrafe a capítulos e aos seus títulos.
    # O conteúdo estrutural prevalece sobre esse estilo inconsistente.
    if HEADING.match(text):
        return "heading"
    if "epigrafe" in style and not identify(text):
        return "heading"
    if style in ("a05_nr", "a05nr", "a21_nova_redacao"):
        return "quotation"
    if any(s in style for s in ("epigrafe", "publicacao", "ementa", "autoridade", "cargo", "ordem_preamb")):
        return "outside"
    if "titulo" in style or HEADING.match(text):
        return "heading"
    # Modelos sem nomes SEFA podem usar a mesma cor verde para o histórico.
    if label(text):
        if color is not None and str(color).upper() in ("008000", "00B050"):
            return "old"
    return "current"


@dataclass
class Block:
    index: int
    text: str
    role: str
    style: str = ""
    element: object = None


@dataclass
class Node:
    target: Target
    start: int
    head_end: int = 0
    end: int = 0
    lead: int = 0
    heading_end: int = 0
    text: str = ""
    label_text: str = ""


class Legislation:
    def __init__(self, document):
        self.doc = Document(document) if not hasattr(document, "element") else document
        if self.doc.element.xpath("//w:ins|//w:del"):
            raise ValueError("A legislação base contém alterações controladas. Aceite ou rejeite essas alterações em uma cópia antes de consolidar.")
        self.blocks = []
        self.nodes = []
        self.by_target = defaultdict(list)
        self.warnings = []
        self.corpora = ["PRINCIPAL"]
        self._index()
        top = " ".join(b.text for b in self.blocks[:15])
        self.identity = identify(top)
        year=re.search(r"\b(?:19|20)\d{2}\b",self.blocks[0].text if self.blocks else "")
        self.identity_year=year[0] if year else ""
        self.profile = self.styles()

    def styles(self):
        counts = Counter(b.style for b in self.blocks if b.role == "current")
        note_counts = Counter(b.style for b in self.blocks if b.role == "new_note")
        old_counts = Counter(b.style for b in self.blocks if b.role == "old")
        old_note_counts = Counter(b.style for b in self.blocks if b.role == "old_note")
        def choose(counter, candidates):
            for value, count in counter.most_common():
                if value in candidates:
                    return value
            for value in candidates:
                if value in self.doc.styles:
                    return value
            return candidates[0]
        return {"current": choose(counts, ["a06_texto", "a20_texto"]),
                "new_note": choose(note_counts, ["a10_nr_remiss", "a30_remiss_nova_red"]),
                "old_note": choose(old_note_counts, ["a11_rdant_remiss", "a31_remiss_red_ant"]),
                "old": choose(old_counts, ["a12_rdant", "a23_redacao_ant"])}

    def _style_table(self):
        names, colors, default = {}, {}, None
        for st in self.doc.styles.element.findall(qn("w:style")):
            sid = st.get(qn("w:styleId"))
            name = st.find(qn("w:name"))
            names[sid] = name.get(qn("w:val")) if name is not None else sid
            color = st.find(qn("w:rPr") + "/" + qn("w:color"))
            colors[sid] = color.get(qn("w:val")) if color is not None else None
            if st.get(qn("w:type")) == "paragraph" and st.get(qn("w:default")) in ("1", "true"):
                default = sid
        builtin = {"heading1": "Heading 1", "heading2": "Heading 2", "title": "Title"}
        for sid, value in list(names.items()):
            if value and value.lower() in builtin:
                names[sid] = builtin[value.lower()]
        return names, colors, default

    def _paragraph_style(self, el, table):
        names, colors, default = table
        ppr = el.find(qn("w:pPr"))
        sid = None
        if ppr is not None:
            ps = ppr.find(qn("w:pStyle"))
            if ps is not None:
                sid = ps.get(qn("w:val"))
        if sid not in names:
            sid = default
        return names.get(sid, "Normal"), colors.get(sid)

    def _index(self):
        table = self._style_table()
        corpus = "PRINCIPAL"
        context = [""] * 5
        boundaries = []
        outside = False
        for i, el in enumerate(self.doc.element.body):
            if el.tag == qn("w:p"):
                p = None
                text = paragraph_text(el).strip()
                style_name, color = self._paragraph_style(el, table)
                r = role(p, style_name, color, text)
                self.blocks.append(Block(i, text, r, style_name, el))
                nv = norm(text)
                if r not in ("old", "old_note", "new_note"):
                    is_adct = re.match(r"^(?:atos?\s+das?\s+)?disposicoes\s+(?:constitucionais\s+)?transitorias\b", nv)
                    is_annex = re.match(r"^(anexo|apendice)\s+([ivxlcdm]+|\d+|unico|[a-z])\s*(?:$|[-–:])", nv)
                    if is_adct:
                        corpus = "ADCT"; context = [""] * 5; boundaries.append(i); outside = False
                    elif is_annex:
                        corpus = (is_annex[1]+' '+is_annex[2]).upper(); context = [""] * 5; boundaries.append(i); outside = False
                    if corpus not in self.corpora:
                        self.corpora.append(corpus)
                if r in ("heading", "outside"):
                    boundaries.append(i)
                    # Só a assinatura encerra o articulado. Um estilo de
                    # ementa/epígrafe usado no meio da lei não pode truncá-la.
                    if r == "outside" and self.nodes and any(s in norm(style_name) for s in ("autoridade", "cargo")):
                        outside = True
                    continue
                if r != "current" or outside:
                    continue
                info = label(text)
                if not info:
                    continue
                level, number, length, prefix = info
                if level > 0 and not context[0]:
                    continue
                context[level:] = [number] + [""] * (4-level)
                target = Target(corpus, *context)
                lead = i
                while lead > 0 and self.blocks[lead-1].role == "new_note":
                    lead -= 1
                node = Node(target, i, lead=lead, text=text, label_text=prefix)
                self.nodes.append(node); self.by_target[target].append(node)
            elif el.tag == qn("w:tbl"):
                self.blocks.append(Block(i, "[TABELA]", "table", "", el))
            else:
                self.blocks.append(Block(i, "", "other", "", el))
        for k, n in enumerate(self.nodes):
            after = self.nodes[k+1:]  # sem indexar as redações históricas
            end = len(self.blocks)
            for other in after:
                if other.target.corpus != n.target.corpus or other.target.level <= n.target.level:
                    end = other.lead
                    break
            limits = [b for b in boundaries if b > n.start]
            if limits:
                end = min(end, min(limits))
            if self.blocks and self.blocks[-1].element.tag == qn("w:sectPr"):
                end = min(end, len(self.blocks)-1)
            n.end = end
            n.head_end = n.start + 1
            # Continuação em outro parágrafo: somente texto corrente sem outro rótulo.
            while n.head_end < end:
                b = self.blocks[n.head_end]
                if b.role != "current" or label(b.text) or not b.text:
                    break
                n.head_end += 1
        duplicates=[t for t,matches in self.by_target.items() if len(matches)>1]
        if duplicates:
            details="\n".join(f"• {t.label()} ({t.corpus})" for t in duplicates)
            self.warnings.append(
                f"Há {len(duplicates)} destinos com mais de uma ocorrência da mesma numeração na base. "
                "Isso pode ocorrer quando o arquivo apresenta versões diferentes de uma redação. "
                "Essas ocorrências são preservadas. Uma proposta cujo destino apareça mais de uma vez "
                "fica com situação ‘revisar’; os demais destinos podem ser atualizados.\n"+details)
        if self.doc.element.xpath("//w:sdt"):
            self.warnings.append("Há controles de conteúdo no documento. Dispositivos dentro deles podem não ser indexados.")
        indexed={n.start for n in self.nodes}
        skipped = [b for b in self.blocks if b.role == "current" and label(b.text) and b.index not in indexed]
        if skipped:
            self.warnings.append(f"{len(skipped)} rótulos ficaram fora do articulado indexado. Confira as assinaturas e os anexos da base.")

    def find(self, target):
        matches = self.by_target.get(target, [])
        return matches[0] if len(matches) == 1 else None

    def current_lines(self, node, subtree=False):
        stop = node.end if subtree else node.head_end
        return [b.text for b in self.blocks[node.start:stop] if b.role in ("current", "table")]

    def location(self, node):
        return f"{node.target.corpus}, bloco {node.start+1} do Word"


def identify(text):
    n = norm(text)
    if "constituicao" in n[:160] and ("estadual" in n[:160] or "para" in n[:160]):
        return "CONSTITUICAO_ESTADUAL"
    if "constituicao" in n[:160] and ("federativa" in n[:160] or "federal" in n[:160]):
        return "CONSTITUICAO_FEDERAL"
    m = re.search(r"\b(lei\s+complementar|lei|decreto(?:\s+legislativo)?|portaria(?:\s+conjunta)?|instrucao\s+normativa|resolucao|emenda\s+constitucional)\s*(?:n\s*[o.]?\s*)?([\d.]+)", n)
    if m:
        return m[1].upper().replace(" ", "_") + ":" + m[2].replace(".", "")
    return ""
