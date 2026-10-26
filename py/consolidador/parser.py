"""Extrai propostas por regras. Não deduz revogação tácita nem vigência jurídica.

Formas de comando reconhecidas (as mais comuns nos atos que alteram as leis
consolidadas pela SEFA):
  • "A Lei nº X passa a vigorar com as seguintes alterações:" seguida de artigos
    citados entre aspas, com linhas pontilhadas e "(NR)";
  • "O inciso II do art. 2º ... passa a vigorar com a seguinte redação:";
  • "Fica acrescido o § 3º ao art. 5º ..." / "acrescida do seguinte art. 10-A";
  • estilo enumerado dos decretos e leis estaduais:
    "Os dispositivos a seguir indicados ... passam a vigorar ...:
     I - o inciso I do art. 2º: “I - ...”;";
  • "Ficam revogados: I - o art. 3º; II - ..." e
    "Revogam-se o inciso III do art. 2º e o § 4º do art. 5º ...";
  • "renumerado o parágrafo único para § 1º".
"""
from __future__ import annotations
import re
from dataclasses import replace
from .index import label, identify, ARTICLE
from .model import Target, Operation, SourceLine, norm
from . import refs as refparser

VERBS = re.compile(r"\b(?:passa[m]?\s+a\s+vigorar|fica[m]?\s+(?:acrescid|acrescentad|revogad|alterad|renumerad|substitu|suprimid|inclu[ií]d|inserid)|acrescenta|acrescem|acrescid|revoga|renumera|substitui|suprime|nova\s+reda[cç][aã]o|seguinte[s]?\s+(?:altera[cç][oõ]es|reda[cç][aã]o|modifica[cç][oõ]es))", re.I)
NUMREF = r"\d+(?:[º°o])?(?:-[A-Za-z]{1,3})?"
QUOTES = ("“", '"', "”", "‘‘", "''")
END_QUOTE = re.compile(r'[”"]\s*(?:\(\s*(?:N\.?\s?R\.?|A\.?\s?C\.?)\s*\)\s*)?[.;,]?\s*(?:e|ou)?\s*$', re.I)
REVOKED_BODY = re.compile(r"^\(?\s*(?:revogad[oa]s?|suprimid[oa]s?)\s*\)?\s*[.;:]?\s*(?:e|ou)?$", re.I)
SPACING_NOTE = "Foi normalizado o espaçamento de uma referência a artigo existente na base. Compare a nova redação com a origem."
GENERIC_REVOCATION = re.compile(r"revoga(?:m-se|-se|das?|dos?)?\s+(?:as\s+)?disposi[cç][oõ]es\s+em\s+contr[aá]rio|disposi[cç][oõ]es\s+em\s+contr[aá]rio\s+(?:ficam\s+)?revogad", re.I)


def roman(n):
    return refparser.roman(n)


def number_value(s):
    s = s.upper()
    try:
        return int(re.match(r"\d+", s)[0])
    except (TypeError, ValueError):
        v = refparser.roman_value(re.sub(r"[^IVXLCDM]", "", s))
        return v or 0


def clean_num(s):
    return re.sub(r"(?<=\d)[º°o]", "", s).upper()


def reference_lists(text):
    """Compatibilidade: listas por nível da primeira referência do texto."""
    found = refparser.parse(text.replace("–", "-"))
    if not found:
        return {}, ""
    first = found[0]
    result = {k: list(v) for k, v in first.parts.items() if v}
    first.targets()
    warning = first.warning
    if len(found) > 1 and not warning:
        others = [r for r in found[1:] if not refparser.is_self(r.law)]
        if others:
            warning = ""
    return result, warning


def _law_matches_base(law, base_identity):
    if not law or not base_identity:
        return True
    if refparser.is_self(law):
        return False
    n = norm(law)
    if base_identity.startswith("CONSTITUICAO"):
        if re.match(r"(?:ato das|adct)", n):
            return True
        if "constituicao" in n:
            federal = "federal" in n or "republica" in n
            return federal == (base_identity == "CONSTITUICAO_FEDERAL")
        return False
    found = identify(law)
    return not found or found == base_identity


def targets_from_command(text, corpus="PRINCIPAL", base_identity=""):
    # O artigo que abre o comando pertence ao alterador, e não à legislação base.
    text = ARTICLE.sub("", text, count=1)
    result, warnings = [], []
    for ref in refparser.parse(text.replace("–", "-")):
        if refparser.is_self(ref.law) or not _law_matches_base(ref.law, base_identity):
            continue
        ts = ref.targets(corpus)
        if ref.warning:
            warnings.append(ref.warning)
        result.extend(t for t in ts if t not in result)
    return result, " ".join(dict.fromkeys(warnings))


def source_citation(lines):
    for line in lines[:20]:
        n = norm(line.text)
        m = re.search(r"\b(emenda\s+constitucional|lei\s+complementar|lei|decreto|portaria|instrucao\s+normativa|resolucao)\s*(?:n\s*[o.]?\s*)?([\d.]+)(?:/(\d{2,4}))?", n)
        if m:
            names = {"emenda constitucional": "EC", "lei complementar": "LC", "lei": "Lei", "decreto": "Decreto", "portaria": "Portaria", "instrucao normativa": "Instrução Normativa", "resolucao": "Resolução"}
            year = m[3]
            if not year:
                y = re.search(r"\b(?:19|20)\d{2}\b", n[m.end():])
                year = y[0] if y else ""
            return names[m[1]] + " " + m[2] + ("/" + year if year else "")
    return ""


def corpus_for(text):
    n = norm(text)
    if "adct" in n or "disposicoes constitucionais transitorias" in n or "ato das disposicoes transitorias" in n:
        return "ADCT"
    m = re.search(r"\b(anexo|apendice)\s+([ivxlcdm]+|\d+|unico)\b", n)
    return (m[1] + " " + m[2]).upper() if m else "PRINCIPAL"


def clean_payload(text):
    text = text.strip().lstrip('“"”').strip()
    text = re.sub(r'[”"]\s*(?:\(\s*(?:N\.?\s?R\.?|A\.?\s?C\.?)\s*\))?\s*[.;,]?\s*(?:e|ou)?\s*$', '', text, flags=re.I)
    text = re.sub(r"\s*\(\s*(?:N\.?\s?R\.?|A\.?\s?C\.?)\s*\)\s*[.;]?$", "", text, flags=re.I)
    return text.rstrip().rstrip('”"').rstrip()


def is_omission(text):
    stripped = re.sub(r"(?:\.{3,}|…+|[_·.]{3,})", "", text).strip(" .;:-()[]\t")
    return not stripped


def normalize_reference_spacing(text, base):
    # Corrige apenas uma referência a artigo que existe na base, dentro de
    # uma expressão "art./arts.". Não reescreve valores nem palavras soltas.
    def repair(m):
        candidate = re.sub(r"\s+", "", m[1]).upper()
        prefix = text[:m.start()]
        if (re.search(r"\barts?\.[^.;:]{0,70}$", prefix, re.I)
                and any(t.article == candidate for t in base.by_target)):
            return candidate
        return m[0]
    return re.sub(r"\b(\d(?:\s+\d)+\s*-\s*[A-Z]{1,3})\b", repair, text)


def line_location(lines):
    first, last = lines[0], lines[-1]
    if first.page:
        return f"PDF, pág. {first.page}" + (f" a {last.page}" if last.page != first.page else "")
    return f"origem, bloco/linha {first.index}" + (f" a {last.index}" if last.index != first.index else "")


def foreign_target(command, base_identity, base_year=""):
    n = norm(command)
    if not base_identity:
        return False, "A identidade da legislação base não foi reconhecida. Confirme o destino."
    ids = []
    date_warning = ""
    for m in re.finditer(r"\b(lei\s+complementar|lei|decreto(?:\s+legislativo)?|portaria(?:\s+conjunta)?|instrucao\s+normativa|resolucao)\s*(?:n\s*[o.]?\s*)?([\d.]+)", n):
        found_id = m[1].upper().replace(" ", "_") + ":" + m[2].replace(".", "")
        ids.append(found_id)
        if found_id == base_identity and base_year:
            tail = n[m.end():m.end() + 100]
            y = re.match(r"\s*/\s*(\d{2,4})\b", tail) or re.match(r"\s*,?\s*de\s+\d{1,2}\s+de\s+\w+\s+de\s+(\d{4})\b", tail)
            if y:
                supplied = y[1]
                matches = supplied == base_year or (len(supplied) == 2 and supplied == base_year[-2:])
                if not matches:
                    return True, "O número coincide, mas o ano da legislação citada difere do ano da base."
            elif found_id.startswith(("PORTARIA:", "PORTARIA_CONJUNTA:", "RESOLUCAO:", "INSTRUCAO_NORMATIVA:")):
                date_warning = "Confirme o ano e o órgão emissor do ato, pois essa numeração pode se repetir."
    if "constituicao" in n:
        if "federal" in n or "republica" in n:
            ids.append("CONSTITUICAO_FEDERAL")
        elif "estadual" in n or "estado do para" in n or "do estado" in n:
            ids.append("CONSTITUICAO_ESTADUAL")
        elif base_identity.startswith("CONSTITUICAO"):
            return False, "O comando menciona Constituição sem identificar federal ou estadual. Confirme o destino."
    elif base_identity.startswith("CONSTITUICAO") and re.search(r"\bato das disposicoes (?:constitucionais )?transitorias\b|\badct\b", n):
        # Uma emenda que altera o ADCT altera a própria Constituição da base.
        ids.append(base_identity)
    if ids and base_identity not in ids:
        return True, "O comando menciona outra legislação: " + ", ".join(dict.fromkeys(ids)) + "."
    if len(set(ids)) > 1:
        return False, "O comando menciona mais de uma legislação. Confirme a qual pertence cada mudança."
    if not ids:
        return False, "O comando não identifica expressamente a legislação destino. Confirme o vínculo com a base."
    return False, date_warning


# --------------------------------------------------------------------------
# Separação do ato em comandos

REFERENCE_ITEM = re.compile(r"^(?:o|a|os|as|ao|à|aos|às|no|na|nos|nas)\s+(?:caput\b|arts?\.|artigos?\b|§|par[aá]grafos?\b|incisos?\b|al[ií]neas?\b|ite(?:m|ns)\b|anexos?\b)", re.I)


def _is_reference_item(text, info):
    """'I - o inciso I do art. 2º:' — item que só indica o dispositivo alterado."""
    body = text[info[2]:].strip()
    if not re.search(r":\s*$", body) or not REFERENCE_ITEM.match(body):
        return False
    return bool(refparser.parse(body))


def group_commands(lines):
    groups = []
    command, payload = [], []
    in_quote = False
    closed = False
    inherited = ""
    meta = {"article": "", "item": ""}

    def flush():
        nonlocal command, payload, in_quote, closed
        if command:
            groups.append((command, payload, dict(meta)))
        command, payload, in_quote, closed = [], [], False, False

    for line in lines:
        text = line.text.strip()
        info = label(text)
        quote_start = text.startswith(QUOTES) and bool(info)
        own_article = bool(ARTICLE.match(text)) and not quote_start and not in_quote
        verb = bool(VERBS.search(text))
        # Comandos enumerados em incisos/alíneas herdam a identificação da lei
        # do artigo introdutório, mas a numeração do inciso é do alterador.
        subcommand = bool(info and info[0] in (2, 3) and not in_quote and not quote_start and command
                          and (verb or _is_reference_item(text, info)))
        if subcommand:
            if not payload:
                if not meta["item"]:
                    inherited = " ".join(c.text for c in command)
            else:
                flush()
            meta["item"] = info[1]
            command = [SourceLine((inherited + " " + text).strip(), line.index, line.page, line.style, line.kind)]
            continue
        if own_article and command and not payload and not verb and any(c.text.rstrip().endswith(':') for c in command) and not re.search(r"entra\s+em\s+vigor|produz\w*\s+efeitos", norm(text)):
            own_article = False
        if own_article:
            flush()
            inherited = ""
            meta = {"article": label(text)[1] if label(text) else "", "item": ""}
            command = [line]
            continue
        if command:
            if quote_start:
                in_quote = True
                closed = False
            if closed and not quote_start:
                continue
            if payload or quote_start or (info and (verb or any(VERBS.search(c.text) for c in command))) or line.kind == "table":
                payload.append(line)
            else:
                command.append(line)
            if in_quote and END_QUOTE.search(text):
                in_quote = False
                closed = True
    flush()
    return groups


def split_blocks(data):
    """Cada citação entre aspas forma um bloco; linhas sem aspas formam um bloco implícito."""
    blocks, current, inside = [], [], False
    for line in data:
        text = line.text.strip()
        if text.startswith(QUOTES) and not inside:
            if current:
                blocks.append(current)
            current = []
            inside = True
        current.append(line)
        if inside and END_QUOTE.search(text):
            blocks.append(current)
            current = []
            inside = False
    if current:
        blocks.append(current)
    return blocks


# --------------------------------------------------------------------------
# Leitura das citações

class Item:
    def __init__(self, target, level, text, stub, revoked, gap_before):
        self.target = target
        self.level = level
        self.lines = [text]
        self.stub = stub
        self.revoked = revoked
        self.gap_before = gap_before
        self.inner_ellipsis = False
        self.normalized = False


def _common_prefix(targets):
    if not targets:
        return [""] * 5
    parts = list(targets[0].parts)
    for t in targets[1:]:
        for k in range(5):
            if parts[k] != t.parts[k]:
                parts[k:] = [""] * (5 - k)
                break
    return parts


def read_items(block, targets, corpus, base):
    parts = list(targets[0].parts) if len(targets) == 1 else _common_prefix(targets)
    items, pending_gap, warning = [], False, ""
    normalized = set()
    for line in block:
        raw = clean_payload(line.text)
        text = normalize_reference_spacing(raw, base)
        if not text:
            continue
        if text != raw:
            normalized.add(id(line))
        info = label(text)
        if info:
            level, num, length, prefix = info
            parts[level:] = [num] + [""] * (4 - level)
            body = text[length:]
            stub = is_omission(body)
            revoked = bool(REVOKED_BODY.match(body.strip())) or bool(re.fullmatch(r"\s*REVOGAD[OA]S?\s*[.;]?\s*", body))
            item = Item(Target(corpus, *parts), level, text, stub, revoked, pending_gap)
            item.inner_ellipsis = not stub and bool(re.search(r"\.{3,}|…", body))
            item.normalized = id(line) in normalized
            items.append(item)
            pending_gap = False
        elif is_omission(text):
            pending_gap = True
        elif items:
            items[-1].lines[-1] += " " + text
            items[-1].normalized = items[-1].normalized or id(line) in normalized
            if re.search(r"\.{3,}|…", text):
                items[-1].inner_ellipsis = True
        else:
            warning += " Há texto sem rótulo antes dos dispositivos. Confira o comando."
    return items, pending_gap, warning


class BaseView:
    """Consultas à base considerando a renumeração de parágrafo único para § 1º."""

    def __init__(self, base):
        self.base = base
        self.renumbered = set()

    def source(self, t):
        if (t.corpus, t.article) in self.renumbered and t.paragraph == "1":
            return replace(t, paragraph="UNICO")
        return t

    def exists(self, t):
        return bool(self.base.by_target.get(self.source(t)))

    def find(self, t):
        return self.base.find(self.source(t))

    def descendants(self, t):
        node = self.find(t)
        if not node:
            return []
        out = []
        for n in self.base.nodes:
            if node.start < n.start < node.end and node.target.contains(n.target):
                tt = n.target
                if tt.paragraph == "UNICO" and (tt.corpus, tt.article) in self.renumbered:
                    tt = replace(tt, paragraph="1")
                out.append(tt)
        return out


def _body(text):
    info = label(text)
    return " ".join(text[info[2]:].split()) if info else " ".join(text.split())


def block_operations(block, targets, action, corpus, base, cmd):
    """Propostas de uma citação. Devolve (lista de dicts, aviso)."""
    items, trailing_gap, warning = read_items(block, targets, corpus, base)
    view = BaseView(base)
    ncmd = norm(cmd)
    explicit = set(targets)
    # Parágrafo único que passa a § 1º quando a citação traz § 1º.
    for it in items:
        t = it.target
        if (it.level == 1 and t.paragraph == "1" and not base.by_target.get(t)
                and base.by_target.get(replace(t, paragraph="UNICO"))
                and not any(x.target.paragraph == "UNICO" and x.target.article == t.article for x in items)):
            view.renumbered.add((t.corpus, t.article))
    result = []
    if not items:
        return result, warning, items
    caput_only = "caput" in ncmd
    # Artigo transcrito por inteiro, sem linhas pontilhadas.
    if (items[0].level == 0 and not items[0].stub and len(items) > 1 and not trailing_gap and not caput_only
            and not any(x.stub or x.gap_before or x.revoked for x in items)):
        root = items[0].target
        lines = [s for x in items for s in x.lines]
        if view.exists(root):
            missing = [t for t in view.descendants(root) if t not in {x.target for x in items}]
            result.append(dict(action="replace", target=root, lines=lines, scope="subtree", full_text=True,
                               warning=(" A redação integral do artigo não traz " + ", ".join(t.description(corpus_suffix=False) for t in missing)
                                        + ". Esses dispositivos sairão do texto vigente e irão para o histórico. Confira e clique em Validar correção.") if missing else ""))
        else:
            result.append(dict(action="add", target=root, lines=lines, scope="subtree"))
        return result, warning, items
    i = 0
    while i < len(items):
        it = items[i]
        span_end = i + 1
        while span_end < len(items) and items[span_end].level > it.level:
            span_end += 1
        desc = items[i + 1:span_end]
        t = it.target
        if it.stub:
            if view.source(t) != t and it.level == 1 and view.exists(t):
                # "§ 1º ......": só renumera o parágrafo único.
                result.append(dict(action="rename", target=view.source(t), rename_to="1", lines=[]))
            i += 1
            continue
        if it.revoked:
            prev = result[-1] if result else None
            if (prev and prev["action"] == "revoke" and prev.get("in_quote") and not it.gap_before
                    and prev["target"].parent() == t.parent() and prev["target"].level == t.level):
                prev.setdefault("extra", []).append(t)
            else:
                result.append(dict(action="revoke", target=t, lines=[], scope="subtree", in_quote=True))
            i = span_end
            continue
        exists = view.exists(t)
        if not exists and not (action == "replace" and t in explicit):
            lines = [s for x in [it] + desc for s in x.lines]
            warn = ""
            if any(x.stub or x.revoked for x in desc):
                warn = " O acréscimo contém linhas pontilhadas ou revogações internas; confira a redação."
            prev = result[-1] if result else None
            if (prev and prev["action"] == "add" and prev.get("open") and not it.gap_before
                    and prev["target"].parent() == t.parent() and prev["target"].level == t.level):
                prev.setdefault("extra", []).append(t)
                prev["lines"].extend(lines)
                prev["warning"] = prev.get("warning", "") + warn
            else:
                result.append(dict(action="add", target=t, lines=lines, scope="subtree", open=True, warning=warn))
            i = span_end
            continue
        if action == "add" and t in explicit and exists:
            # O ato manda acrescentar um dispositivo que já existe na base.
            lines = [s for x in [it] + desc for s in x.lines]
            result.append(dict(action="add", target=t, lines=lines, scope="subtree"))
            i = span_end
            continue
        renumber = "UNICO" if view.source(t) != t and it.level == 1 else ""
        inner_gap = any(x.gap_before or x.stub for x in desc)
        after_gap = (span_end < len(items) and items[span_end].gap_before) or (span_end == len(items) and trailing_gap)
        if desc and not inner_gap and not any(x.revoked for x in desc) and not (caput_only and it.level == 0):
            new_set = {x.target for x in desc}
            missing = [d for d in view.descendants(t) if d not in new_set]
            if not missing or not after_gap:
                lines = [s for x in [it] + desc for s in x.lines]
                result.append(dict(action="replace", target=t, lines=lines, scope="subtree", renumber=renumber,
                                   full_text=bool(missing),
                                   warning=(" A nova redação não traz linha pontilhada nem " + ", ".join(m.description(corpus_suffix=False) for m in missing)
                                            + ". Esses dispositivos sairão do texto vigente e irão para o histórico. Confira e clique em Validar correção.") if missing else ""))
                i = span_end
                continue
        # Nova redação só do texto do dispositivo; novos filhos imediatos podem
        # acompanhar o caput quando o dispositivo ainda não tinha filhos.
        fold = []
        if desc and not desc[0].gap_before and desc[0].level == (2 if it.level == 0 else it.level + 1):
            child_level = desc[0].level
            existing_children = [d for d in view.descendants(t) if d.level == child_level and d.parent() == t]
            if not existing_children:
                k = i + 1
                while (k < span_end and items[k].level == child_level and not items[k].gap_before and not items[k].stub
                       and not items[k].revoked and not view.exists(items[k].target)):
                    k += 1
                    while k < span_end and items[k].level > child_level and not items[k].gap_before and not items[k].stub:
                        k += 1
                fold = items[i + 1:k]
        lines = [s for x in [it] + fold for s in x.lines]
        if renumber and not fold and _body(it.lines[0]) == _body(" ".join(view.base.current_lines(view.find(t)))):
            result.append(dict(action="rename", target=view.source(t), rename_to="1", lines=[]))
        else:
            result.append(dict(action="replace", target=t, lines=lines, scope="head", renumber=renumber, fold=bool(fold)))
        i += 1 + len(fold)
    changed_lines = {l for x in items if x.normalized for l in x.lines}
    for r in result:
        r.pop("open", None)
        if any(l in changed_lines for l in r.get("lines", [])):
            r["normalized"] = True
    if any(x.inner_ellipsis for x in items):
        warning += " O texto contém reticências misturadas à redação; complete o trecho antes de aplicar."
    return result, warning, items


# --------------------------------------------------------------------------

def _revocation_groups(entries):
    """Une revogações de irmãos consecutivos ("incisos IV e V") numa só nota."""
    out = []
    for target, foreign, warning in entries:
        prev = out[-1] if out else None
        if (prev and not warning and not prev[2] and prev[1] == foreign and target.level > 0
                and prev[0][-1].parent() == target.parent() and prev[0][-1].level == target.level
                and _next_number(prev[0][-1], target)):
            prev[0].append(target)
        else:
            out.append([[target], foreign, warning])
    return out


def _next_number(a, b):
    level = a.level
    x, y = a.parts[level], b.parts[level]
    if "-" in x or "-" in y:
        hx, _, sx = x.partition("-")
        hy, _, sy = y.partition("-")
        return hx == hy and len(sx) <= 1 and len(sy) == 1 and (ord(sy) - ord(sx or "@") == 1)
    if level == 3:
        return ord(y[0]) - ord(x[0]) == 1
    return number_value(y) - number_value(x) == 1


def _merge_adjacent(operations):
    """Acréscimos (ou revogações) de irmãos consecutivos, do mesmo artigo do ato, numa só nota."""
    out = []
    for op in operations:
        prev = out[-1] if out else None
        if (prev and prev.action == op.action and op.action in ("add", "revoke") and not prev.extraction_warning
                and not op.extraction_warning and prev.foreign == op.foreign and not prev.manual
                and prev.source_article == op.source_article and prev.scope == op.scope == "subtree"
                and prev.targets[-1].level == op.target.level and prev.targets[-1].parent() == op.target.parent()
                and _next_number(prev.targets[-1], op.target)):
            prev.extra_targets = prev.extra_targets + op.targets
            prev.new_lines = prev.new_lines + op.new_lines
            if op.source not in prev.source:
                prev.source = prev.source + "\n" + op.source
            continue
        out.append(op)
    return out


def parse_operations(lines, base):
    operations = []
    for cmds, data, meta in group_commands(lines):
        cmd = " ".join(l.text for l in cmds)
        if not VERBS.search(cmd):
            continue
        ncmd = norm(cmd)
        corpus = corpus_for(cmd)
        targets, ref_warning = targets_from_command(cmd, corpus, base.identity)
        if GENERIC_REVOCATION.search(cmd) and not targets:
            continue
        foreign, target_warning = foreign_target(cmd, base.identity, base.identity_year)
        extra = " ".join(x for x in (ref_warning, target_warning) if x)
        src = "\n".join(l.text for l in cmds + data)
        location = line_location(cmds + data)
        common = dict(source=src, source_location=location, foreign=foreign,
                      source_article=meta.get("article", ""), source_item=meta.get("item", ""))
        action = "revoke" if re.search(r"\brevog", ncmd) else "add" if re.search(r"\bacrescid|\bacrescent|\bacresce|\binclu[ií]d|\binserid", ncmd) else "replace"
        blocks = split_blocks(data) if data else []
        has_quotes = any(any(l.text.strip().startswith(QUOTES) for l in b) for b in blocks)
        if re.search(r"renumer", ncmd) and not has_quotes:
            dest = re.search(r"\b(?:para|como)\s+(?:o\s+)?§\s*(\d+)", cmd, re.I)
            if dest:
                sources, _ = targets_from_command(cmd[:dest.start()], corpus, base.identity)
                if len(sources) == 1 and sources[0].paragraph == "UNICO":
                    operations.append(Operation("", "rename", sources[0], scope="head", rename_to=dest[1],
                                                extraction_warning=extra, **common))
                    continue
            if action not in ("revoke",) and not data:
                operations.append(Operation("", "rename", targets[0] if targets else Target(corpus), rename_to="", extraction_warning=(extra + " Renumeração exige indicação manual do dispositivo de origem e do novo rótulo.").strip(), **common))
                continue
        if action == "revoke":
            if re.search(r"a partir da data em que|quando for|apos a (?:implantacao|implementacao|regulamentacao)|mediante aviso|condicionad[oa]", ncmd):
                extra += " A revogação depende de evento ou condição sem data certa; confirme a ocorrência e a data antes de aplicar."
            entries = []
            # Lista hierárquica repartida em incisos: a numeração é do ato revogador.
            if any(label(l.text) for l in data):
                header, header_level = "", None
                for item in data:
                    info = label(item.text)
                    if not info:
                        continue
                    item_text = item.text[info[2]:]
                    if header_level is not None and info[0] <= header_level:
                        header, header_level = "", None
                    item_corpus = corpus_for(header + " " + item_text) if corpus_for(header + " " + item_text) != "PRINCIPAL" else corpus
                    subtargets, subwarning = targets_from_command(item_text, item_corpus, base.identity)
                    if not subtargets and item_text.rstrip().endswith(":"):
                        # "I - da Lei nº 5.530, de 1989:" apenas introduz os itens seguintes.
                        header, header_level = item_text, info[0]
                        continue
                    item_foreign, item_warning = foreign_target(cmd + " " + header + " " + item_text, base.identity, base.identity_year)
                    if not subtargets:
                        if GENERIC_REVOCATION.search(item_text):
                            continue
                        subtargets = [Target(corpus)]
                        subwarning += " Lista de revogação sem destino reconhecido."
                    warning = " ".join(x for x in (ref_warning, subwarning, item_warning) if x).strip()
                    if "condição" in extra or "evento" in extra:
                        warning = (warning + " " + extra).strip()
                    entries.extend((t, item_foreign, warning) for t in subtargets)
            else:
                if not targets:
                    extra += " Não foi encontrado um destino explícito. Revogação genérica não é aplicada."
                    targets = [Target(corpus)]
                entries = [(t, foreign, extra.strip()) for t in targets]
            for group, item_foreign, warning in _revocation_groups(entries):
                operations.append(Operation("", "revoke", group[0], scope="subtree", extra_targets=group[1:],
                                            extraction_warning=warning, **dict(common, foreign=item_foreign)))
            continue
        if any(x.kind == "table" for x in data):
            operations.append(Operation("", action, targets[0] if len(targets) == 1 else Target(corpus),
                                        extraction_warning="O comando inclui tabela. A revisão textual não aplica tabelas; faça essa alteração no Word. " + extra, **common))
            continue
        produced = False
        for block in blocks:
            proposals, warning, items = block_operations(block, targets, action, corpus, base, cmd)
            # Cada proposta mostra o comando e só a citação que lhe corresponde.
            block_common = dict(common, source="\n".join(l.text for l in cmds + block),
                                source_location=line_location(block))
            for prop in proposals:
                w = (extra + warning + prop.get("warning", "")).strip()
                if not prop["target"].article:
                    w += " O artigo de destino não foi identificado."
                op = Operation("", prop["action"], prop["target"], new_lines=prop.get("lines", []), scope=prop.get("scope", "head"),
                               extraction_warning=w.strip(), extra_targets=prop.get("extra", []),
                               renumber_from=prop.get("renumber", ""), rename_to=prop.get("rename_to", ""),
                               note=prop.get("note", ""), **block_common)
                op.full_text = prop.get("full_text", False)
                op.fold = prop.get("fold", False)
                if prop.get("normalized"):
                    op.note = (op.note + " " if op.note else "") + SPACING_NOTE
                if prop.get("in_quote") and op.action == "revoke" and re.search(r"a partir da data em que|quando for|condicionad", ncmd):
                    op.extraction_warning += " A revogação depende de evento ou condição; confirme."
                operations.append(op)
                produced = True
        if not produced:
            operations.append(Operation("", action, targets[0] if len(targets) == 1 else Target(corpus),
                                        extraction_warning=(extra + " Não foi extraída uma redação completa identificada por rótulo.").strip(), **common))
    operations = _merge_adjacent(operations)
    if not operations:
        text = "\n".join(l.text for l in lines)
        operations = [Operation("", "replace", Target(), source=text, source_location=line_location(lines) if lines else "",
                                extraction_warning="Nenhum comando de alteração suportado foi reconhecido. Cadastre os destinos e textos na revisão.")]
    for i, op in enumerate(operations, 1):
        op.id = f"A{i:03d}"
    return operations
