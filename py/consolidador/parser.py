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
from . import headings as headmod
from . import annexes as annexmod

VERBS = re.compile(r"\b(?:substitu[aí]\w*|suprima[m]?-se|onde\s+se\s+l[eê]|passa[m]?\s+a\s+vigorar|passa[m]?\s+a\s+(?:ser|constituir)\s+(?:o|a|os|as)\s+(?:§|par[aá]grafo|incisos?|al[ií]neas?|ite[mn]s?)|passa[m]?\s+a\s+denominar|seguinte\s+denomina[cç][aã]o|denomina[cç][aã]o\s+d[ao]\s+(?:Parte|Livro|T[ií]tulo|Cap[ií]tulo|Se[cç][aã]o|Subse[cç][aã]o)\b.{0,160}?\bpassa[m]?\s+a\s+ser|fica[m]?\s+(?:acrescid|acrescentad|revogad|alterad|renumerad|substitu|suprimid|inclu[ií]d|inserid)|acrescenta|acrescem|acrescid|revoga|renumera|substitui|suprime|nova\s+reda[cç][aã]o|seguinte[s]?\s+(?:altera[cç][oõ]es|reda[cç][aã]o|modifica[cç][oõ]es))", re.I)
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
        m = re.search(r"\b(emenda\s+constitucional|lei\s+complementar|lei|decreto|portaria|instrucao\s+normativa|resolucao)\s*(?:n\s*[o.]?\s*)?(\d[\d.]*)(?:/(\d{2,4}))?", n)
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
    for m in re.finditer(r"\b(lei\s+complementar|lei|decreto(?:\s+legislativo)?|portaria(?:\s+conjunta)?|instrucao\s+normativa|resolucao)\s*(?:n\s*[o.]?\s*)?(\d[\d.]*)", n):
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
        quote_start = text.startswith(QUOTES) and (bool(info) or bool(headmod.parse_line(text)))
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
        # Linhas de título ("SEÇÃO V", "Da Justiça Militar") logo antes deste artigo.
        self.heading = []


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


def read_items(block, targets, corpus, base, orphans=None):
    parts = list(targets[0].parts) if len(targets) == 1 else _common_prefix(targets)
    items, pending_gap, warning = [], False, ""
    normalized = set()
    heading = []
    for line in block:
        raw = clean_payload(line.text)
        text = normalize_reference_spacing(raw, base)
        if not text:
            continue
        if text != raw:
            normalized.add(id(line))
        info = label(text)
        if not info and headmod.parse_line(text):
            if heading and orphans is not None:
                orphans.append(heading)
            heading = [text]
            continue
        if heading and not info and not is_omission(text):
            heading.append(text)
            continue
        if heading and (is_omission(text) or (info and info[0] != 0)):
            # Título sem artigo logo em seguida: só a denominação muda.
            if orphans is not None:
                orphans.append(heading)
            heading = []
        if info:
            level, num, length, prefix = info
            parts[level:] = [num] + [""] * (4 - level)
            body = text[length:]
            stub = is_omission(body)
            revoked = bool(REVOKED_BODY.match(body.strip())) or bool(re.fullmatch(r"\s*REVOGAD[OA]S?\s*[.;]?\s*", body))
            item = Item(Target(corpus, *parts), level, text, stub, revoked, pending_gap)
            item.inner_ellipsis = not stub and bool(re.search(r"\.{3,}|…", body))
            item.normalized = id(line) in normalized
            item.heading, heading = heading, []
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
    if heading and orphans is not None:
        orphans.append(heading)
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
    """Propostas de uma citação. Devolve (lista de dicts, aviso, itens).

    Uma citação com vários artigos (“Art. 168. ... Art. 169. ...” (NR)) é tratada
    artigo por artigo."""
    orphans = []
    items, trailing_gap, warning = read_items(block, targets, corpus, base, orphans)
    heads = [heading_proposal(lines, corpus, base) for lines in orphans]
    heads += [heading_proposal(x.heading, corpus, base, x.target) for x in items if x.level == 0 and x.heading]
    heads = [h for h in heads if h]
    if heads and not items:
        return heads, warning, items
    tops = [k for k, x in enumerate(items) if x.level == 0]
    if len(tops) > 1 and tops[0] == 0:
        chunks = [items[a:b] for a, b in zip(tops, tops[1:] + [len(items)])]
        result = []
        for n, chunk in enumerate(chunks):
            gap = chunks[n + 1][0].gap_before if n + 1 < len(chunks) else trailing_gap
            r, w = _items_operations(chunk, gap, targets, action, corpus, base, cmd)
            result.extend(r)
            warning += w
        return _with_headings(heads, result), warning, items
    result, w = _items_operations(items, trailing_gap, targets, action, corpus, base, cmd)
    return _with_headings(heads, result), warning + w, items


def heading_proposal(lines, corpus, base, article=None):
    """Título citado no ato. Nova denominação de um título existente, ou título novo
    que acompanha um artigo acrescido."""
    info = headmod.parse_line(lines[0])
    if not info:
        return None
    kind, number, rest = info
    key = f"{kind}:{number}"
    group = headmod.find(base, kind, number, corpus, article=article.article if article else "")
    if group is None and article is None:
        group = headmod.find(base, kind, number, corpus)
    if group is not None:
        new = [" ".join(l.split()) for l in lines]
        old = [" ".join(l.split()) for l in group.lines]
        if norm(" ".join(new)) == norm(" ".join(old)):
            return None
        return dict(action="heading", target=Target(corpus, group.anchor.article if group.anchor else ""), lines=list(lines),
                    heading_key=key, heading_parents="/".join(f"{k}:{n}" for k, n in group.parents))
    if article is not None and not base.find(article):
        # Título novo junto do artigo novo: a nota do acréscimo é a do título.
        return dict(action="heading_add", target=article, lines=list(lines), heading_key=key)
    return dict(action="heading", target=Target(corpus, article.article if article else ""), lines=list(lines), heading_key=key,
                warning=" O título citado não foi localizado na base de forma única. Confira o capítulo ou a seção e o artigo seguinte.")


def _with_headings(heads, result):
    """Títulos novos passam ao acréscimo do artigo que os segue; as novas
    denominações vêm antes das mudanças dos artigos."""
    out = [h for h in heads if h["action"] == "heading"]
    for h in heads:
        if h["action"] != "heading_add":
            continue
        for r in result:
            if r["action"] == "add" and (r["target"] == h["target"] or h["target"] in r.get("extra", [])) and r["target"].level == 0:
                r["heading"] = h["lines"]
                r["heading_key"] = h["heading_key"]
                break
        else:
            out.append(dict(h, action="heading", warning=" Título novo sem artigo acrescido logo em seguida. Confira o ato."))
    return out + result


def _items_operations(items, trailing_gap, targets, action, corpus, base, cmd):
    warning = ""
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
        return result, warning
    caput_only = "caput" in ncmd
    # Artigo transcrito por inteiro, sem linhas pontilhadas.
    if (items[0].level == 0 and not items[0].stub and len(items) > 1 and not trailing_gap and not caput_only
            and not any(x.stub or x.gap_before or x.revoked for x in items)):
        root = items[0].target
        lines = [s for x in items for s in x.lines]
        if view.exists(root) and not view.descendants(root):
            # Artigo sem filhos que passa a tê-los (padrão SEFA 2026, EC 102/26 na
            # Constituição do Pará): nova redação do caput (com os seus incisos) e,
            # à parte, o acréscimo dos parágrafos.
            k = 1
            while k < len(items) and not (items[k].level == 1):
                k += 1
            head, rest = items[:k], items[k:]
            result.append(dict(action="replace", target=root, lines=[s for x in head for s in x.lines], scope="head",
                               fold=len(head) > 1, caput_forced=bool(rest)))
            j = 0
            while j < len(rest):
                group = [rest[j]]
                j += 1
                while j < len(rest) and rest[j].level > 1:
                    group.append(rest[j])
                    j += 1
                prev = result[-1]
                t = group[0].target
                glines = [s for x in group for s in x.lines]
                if prev["action"] == "add" and prev["target"].parent() == t.parent() and prev["target"].level == t.level:
                    prev.setdefault("extra", []).append(t)
                    prev["lines"].extend(glines)
                else:
                    result.append(dict(action="add", target=t, lines=glines, scope="subtree"))
            return result, warning
        if view.exists(root):
            missing = [t for t in view.descendants(root) if t not in {x.target for x in items}]
            result.append(dict(action="replace", target=root, lines=lines, scope="subtree", full_text=True,
                               warning=(" A redação integral do artigo não traz " + ", ".join(t.description(corpus_suffix=False) for t in missing)
                                        + ". Esses dispositivos sairão do texto vigente e irão para o histórico. Confira e clique em Validar correção.") if missing else ""))
        else:
            result.append(dict(action="add", target=root, lines=lines, scope="subtree"))
        return result, warning
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
                result.append(dict(action="revoke", target=t, lines=[], scope="subtree", in_quote=True,
                                   suppressed=bool(re.search(r"suprimid", it.lines[0], re.I))))
            i = span_end
            continue
        exists = view.exists(t)
        shift = action == "add" and bool(RENUMBER_REST.search(cmd)) and it.level >= 1
        if (not exists or shift) and not (action == "replace" and t in explicit):
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
                result.append(dict(action="add", target=t, lines=lines, scope="subtree", open=True, warning=warn, shift=shift and exists))
            if shift and exists:
                result[-1]["shift"] = True
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
    return result, warning


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
        if (prev and prev.action == op.action and op.action in ("add", "revoke") and not prev.extraction_warning and not op.heading
                and prev.via == op.via and prev.suppressed == op.suppressed
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


RENUMBER_REST = re.compile(r"renumerand[oa]?-?se\s+(?:os|as|o|a)\s+(?:demais|seguintes|subsequentes|posteriores)"
                           r"|com\s+(?:a\s+)?(?:consequente\s+|devida\s+)?renumera[cç][aã]o\s+d[oa]s?\s+(?:demais|seguintes|subsequentes)", re.I)
RENUMBER_VERB = re.compile(r"\b(?:renumerad[oa]s?|renumerand[oa]-?se|passa[m]?\s+a\s+(?:ser|constituir))\b", re.I)
UNIT_WORD = {1: "§", 2: "inciso", 3: "alínea", 4: "item"}


def _ordinal_value(value, level):
    head = value.split("-")[0]
    if level == 3:
        return ord(head[0].lower()) - 96
    return number_value(head)


def _from_value(n, level):
    if level == 2:
        return roman(n)
    if level == 3:
        return chr(96 + n)
    return str(n)


def renumber_pairs(text, corpus, base, context_targets=()):
    """'Fica renumerado o § 2º do art. 177 para § 3º', 'renumerando-se o atual § 2º como
    § 3º', 'os incisos VI a XII passam a ser os incisos VIII a XIV'. Devolve [(origem, novo número)]."""
    out = []
    for verb in RENUMBER_VERB.finditer(text):
        bounds = [m.end() for m in re.finditer(r"[;:]|\.(?=\s+[A-ZÁÉÍÓÚÂÊÔÃÕÇ“\"])", text[:verb.start()])]
        clause_start = bounds[-1] if bounds else 0
        after = text[verb.end():]
        end = re.search(r"[;:.](?!\s*\d)|,\s*(?:e\s+)?(?:acrescid|revogad|renumerand)", after)
        after = after[:end.start()] if end else after
        key = re.search(r"\b(?:para|como|em)\b", after, re.I)
        if norm(verb[0]).startswith("passa"):
            src_text, dst_text = text[clause_start:verb.start()], after
        elif key and refparser.parse(after[:key.start()]):
            src_text, dst_text = after[:key.start()], after[key.end():]
        elif key:
            src_text, dst_text = text[clause_start:verb.start()], after[key.end():]
        else:
            continue
        sources, _ = targets_from_command(src_text, corpus, base.identity)
        articles = [t for t in context_targets if t.article]
        if not sources and articles and refparser.parse(src_text):
            # "renumerando-se o atual § 2º como § 3º": o artigo é o do comando.
            sources, _ = targets_from_command(src_text.rstrip(" ,") + f" do art. {articles[0].article}", corpus, base.identity)
        sources = [t for t in sources if t.level >= 1]
        if not sources:
            continue
        level = sources[0].level
        if any(t.level != level for t in sources):
            continue
        dst = dst_text.strip()
        if not re.search(r"§|par[aá]grafo|inciso|al[ií]nea|item", dst, re.I):
            plural = len(sources) > 1
            word = {1: "§§" if plural else "§", 2: "incisos" if plural else "inciso", 3: "alíneas" if plural else "alínea", 4: "itens" if plural else "item"}[level]
            dst = word + " " + re.sub(r"^\s*(?:o|a|os|as)\s+", "", dst)
        parent = sources[0].parent()
        dests, _ = targets_from_command(dst + " do " + parent.description(corpus_suffix=False), corpus, base.identity)
        numbers = [t.parts[level] for t in dests if t.level == level]
        if len(numbers) != len(sources):
            continue
        out.extend(zip(sources, numbers))
    return out


def shift_renames(added, base):
    """Acréscimo com "renumerando-se os demais": os irmãos existentes a partir da posição
    dos novos passam a ocupar os números seguintes (incisos VI e VII novos -> o antigo VI
    passa a VIII). Devolve [(origem, novo número)], do maior para o menor."""
    out = []
    groups = {}
    for t in added:
        groups.setdefault((t.parent(), t.level), []).append(t)
    for (parent, level), ts in groups.items():
        if level < 1 or any("-" in t.parts[level] for t in ts):
            continue
        siblings = sorted((n.target for n in base.nodes if n.target.level == level and n.target.parent() == parent),
                          key=lambda t: _ordinal_value(t.parts[level], level))
        if any("-" in t.parts[level] or t.parts[level] == "UNICO" for t in siblings):
            continue
        taken = {_ordinal_value(t.parts[level], level) for t in ts}
        if not taken or min(taken) > len(siblings) + len(taken):
            continue
        position, pending = 1, list(siblings)
        mapping = []
        while pending:
            if position in taken:
                position += 1
                continue
            old = pending.pop(0)
            if _ordinal_value(old.parts[level], level) != position:
                mapping.append((old, _from_value(position, level)))
            position += 1
        out.extend(sorted(mapping, key=lambda x: -_ordinal_value(x[0].parts[level], level)))
    return out


def rename_operations(pairs, common, extra=""):
    ops = []
    level_sorted = sorted(pairs, key=lambda x: -_ordinal_value(x[0].parts[x[0].level], x[0].level)
                          if _ordinal_value(x[1], x[0].level) > _ordinal_value(x[0].parts[x[0].level], x[0].level)
                          else _ordinal_value(x[0].parts[x[0].level], x[0].level))
    for source, number in level_sorted:
        ops.append(Operation("", "rename", source, scope="head", rename_to=number, extraction_warning=extra.strip(), **common))
    return ops


QUOTED = re.compile(r"[“\"]([^”\"]+)[”\"]")


def expression_change(cmd):
    """'Fica substituída, no art. 5º, a expressão “A” por “B”', 'onde se lê “A”, leia-se “B”',
    'Fica suprimida a expressão “A” do § 2º'. Devolve [(antiga, nova)] ou []."""
    n = norm(cmd)
    if not re.search(r"\bexpress|\btermos?\b|\bpalavras?\b|onde se le", n):
        return []
    quoted = [q.strip() for q in QUOTED.findall(cmd)]
    if not quoted:
        return []
    if re.search(r"\bsuprim", n) and not re.search(r"\bpor\b|leia-se", n):
        return [(q, "") for q in quoted]
    if re.search(r"substitu|leia-se|passa\w*\s+a\s+ser|passa\w*\s+a\s+vigorar\s+como", n) and len(quoted) >= 2 and len(quoted) % 2 == 0:
        half = len(quoted) // 2
        return list(zip(quoted[:half], quoted[half:]))
    return []


def expression_operations(pairs, cmd, corpus, base, targets, common, extra):
    """Uma nova redação para cada dispositivo que contém a expressão."""
    whole = re.search(r"em\s+tod[oa]s?\s+(?:o|a|os|as)\s+(?:texto|lei|decreto|dispositivos)|onde\s+couber", norm(cmd))
    nodes = []
    if whole:
        nodes = [n for n in base.nodes if n.target.corpus == corpus]
    for t in targets:
        node = base.find(t)
        if node is None:
            continue
        caput_only = bool(re.search(r"\bcaput\b", norm(cmd))) and t.level == 0
        if t.level == 0 and not caput_only:
            nodes.extend(n for n in base.nodes if node.start <= n.start < node.end and t.contains(n.target))
        else:
            nodes.append(node)
    ops, seen = [], set()
    for node in nodes:
        if node.start in seen:
            continue
        seen.add(node.start)
        lines = base.current_lines(node)
        info = label(lines[0]) if lines else None
        if not info:
            continue
        head, body = lines[0][:info[2]], [lines[0][info[2]:]] + lines[1:]
        changed = []
        for line in body:
            for old, new in pairs:
                line = re.sub(r"\s{2,}", " ", line.replace(old, new)) if old in line else line
            changed.append(line)
        if changed == body:
            continue
        new_lines = [(head + changed[0]).replace("  ", " ")] + changed[1:]
        what = "; ".join(f"“{a}” por “{b}”" if b else f"supressão de “{a}”" for a, b in pairs)
        ops.append(Operation("", "replace", node.target, new_lines=new_lines, scope="head", extraction_warning=extra.strip(),
                             note=f"Substituição de expressão: {what}.", **common))
    if not ops:
        ops.append(Operation("", "replace", targets[0] if targets else Target(corpus), extraction_warning=(extra + " A expressão citada não foi encontrada no texto vigente dos dispositivos indicados.").strip(), **common))
    return ops


def heading_revocation(cmd, corpus, base):
    """'Fica revogado o Capítulo IV do Título II': os artigos desse título."""
    m = re.search(headmod.NAMED + r"((?:\s*,?\s*d[ao]\s+" + headmod.NAMED_NC + r")*)", cmd, re.I)
    if not m:
        return None
    kind, number = headmod.kind_key(m[1]), headmod.number_key(m[2])
    parents = [(headmod.kind_key(k), headmod.number_key(n)) for k, n in re.findall(headmod.NAMED, m[3] or "", re.I)]
    group = headmod.find(base, kind, number, corpus, parents)
    if group is None:
        return None
    rank = headmod.KINDS[kind][2]
    later = [g for g in headmod.groups(base) if g.start > group.start and g.corpus == group.corpus and headmod.KINDS[g.kind][2] <= rank]
    stop = later[0].start if later else len(base.blocks)
    articles = [n.target for n in sorted(base.nodes, key=lambda n: n.start)
                if n.target.level == 0 and n.target.corpus == group.corpus and group.end <= n.start < stop]
    if not articles:
        return None
    return articles[0], articles[1:], headmod.name(kind, number)


def caput_path(text, target):
    """O ato cita o dispositivo como "inciso XVI do caput do art. 37" (ou "ao caput do art. 22")?"""
    if target.level < 2 or target.paragraph or not target.article:
        return False
    article = norm(target.article)
    return bool(re.search(r"\bcaput\s+do\s+art(?:igo)?\.?\s*" + re.escape(article) + r"o?(?![\d-])", norm(text)))


def implicit_renumbering(operations, base):
    """Acréscimo de § 2º (ou seguinte) a artigo que tem parágrafo único: o
    parágrafo único passa a § 1º ("Renumerado o parágrafo único para § 1º do
    art. 7º pela EC 141/26"), ainda que o ato não o diga expressamente."""
    done = {(o.target.corpus, o.target.article) for o in operations if o.action == "rename" or o.renumber_from}
    out = []
    for op in operations:
        t = op.target
        key = (t.corpus, t.article)
        if (op.action == "add" and t.level == 1 and t.paragraph not in ("", "UNICO", "1") and not op.foreign
                and key not in done and base.find(replace(t, paragraph="UNICO"))
                and not base.by_target.get(replace(t, paragraph="1"))
                and not any(x.paragraph == "1" for x in op.targets)):
            out.append(Operation("", "rename", replace(t, paragraph="UNICO"), scope="head", rename_to="1",
                                 source=op.source, source_location=op.source_location, foreign=op.foreign,
                                 source_article=op.source_article, source_item=op.source_item,
                                 note="Renumeração decorrente do acréscimo de parágrafo numerado a artigo que tinha parágrafo único. Confira no ato."))
            done.add(key)
        out.append(op)
    return out


ACT_NAME = re.compile(r"\b(Lei\s+Complementar|Emenda\s+Constitucional|Lei|Decreto(?:-Lei)?)\s*(?:n[º°o.]?\s*)?([\d.]+)\s*(?:,\s*de\s+\d{1,2}[º°o]?\s+de\s+[a-zç]+\s+de\s+(\d{4})|/(\d{2,4}))", re.I)
ABBR = {"lei complementar": "LC", "emenda constitucional": "EC", "lei": "Lei", "decreto": "Decreto", "decreto-lei": "Decreto-Lei"}


def nested_amendment(cmd, data, base):
    """'A LC 214 passa a vigorar com as seguintes alterações: “Art. 517. A LC 123 passa a
    vigorar com as seguintes alterações: ‘Art. 13. ...’ (NR)” (NR)': o ato altera a redação
    que outro ato alterador deu à base. Devolve (citação do ato alterado, linhas internas)."""
    if not data:
        return None
    inner_cmd = None
    for line in data:
        text = line.text.strip().lstrip("“\"‘'").strip()
        if ARTICLE.match(text) and VERBS.search(text):
            f, w = foreign_target(text, base.identity, base.identity_year)
            if not f and "não identifica" not in w and "mais de uma" not in w:
                inner_cmd = line
                break
    if inner_cmd is None:
        return None
    m = ACT_NAME.search(cmd)
    if not m:
        return None
    kind = ABBR.get(norm(m[1]).replace(" ", " "), m[1])
    year = m[3] or m[4] or ""
    via = f"{kind} {m[2]}/{year}" if year else f"{kind} {m[2]}"
    out = []
    lines = list(data)
    for k, line in enumerate(lines):
        text = line.text
        if k == 0:
            text = re.sub(r"^\s*[“\"]", "", text, count=1)
        if k == len(lines) - 1:
            text = re.sub(r'[”"]\s*(?:\(\s*N\.?\s?R\.?\s*\))?\s*[.;]?\s*$', "", text, count=1)
        text = re.sub(r"^\s*[‘']", "“", text)
        text = re.sub(r"[’']\s*(\(\s*N\.?\s?R\.?\s*\))?\s*([.;]?)\s*$", lambda mm: "”" + (" " + mm[1] if mm[1] else "") + mm[2], text)
        out.append(SourceLine(text.strip(), line.index, line.page, line.style, line.kind))
    return via, out


def parse_operations(lines, base, nested=False):
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
        if foreign and not nested:
            inner = nested_amendment(cmd, data, base)
            if inner:
                via, inner_lines = inner
                for op in parse_operations(inner_lines, base, nested=True):
                    op.via = via
                    op.source = "\n".join(l.text for l in cmds) + "\n" + op.source
                    op.source_article = meta.get("article", "") or op.source_article
                    op.source_item = meta.get("item", "")
                    operations.append(op)
                continue
        extra = " ".join(x for x in (ref_warning, target_warning) if x)
        src = "\n".join(l.text for l in cmds + data)
        location = line_location(cmds + data)
        common = dict(source=src, source_location=location, foreign=foreign,
                      source_article=meta.get("article", ""), source_item=meta.get("item", ""))
        suppressed = bool(re.search(r"\bsuprim", ncmd)) and not re.search(r"\brevog", ncmd)
        action = "revoke" if re.search(r"\brevog", ncmd) or suppressed else "add" if re.search(r"\bacrescid|\bacrescent|\bacresce|\binclu[ií]d|\binserid", ncmd) else "replace"
        blocks = split_blocks(data) if data else []
        has_quotes = any(any(l.text.strip().startswith(QUOTES) for l in b) for b in blocks)
        expressions = expression_change(cmd) if not foreign else []
        if expressions and (targets or re.search(r"em\s+tod[oa]s?\s+(?:o|a)|onde\s+couber", ncmd)):
            operations.extend(expression_operations(expressions, cmd, corpus, base, targets, common, extra))
            continue
        pairs = renumber_pairs(cmd, corpus, base, targets) if RENUMBER_VERB.search(cmd) else []
        if pairs:
            operations.extend(rename_operations(pairs, common, extra))
            if not has_quotes:
                continue
        if re.search(r"renumer", ncmd) and not has_quotes and not pairs:
            if action not in ("revoke",) and not data:
                operations.append(Operation("", "rename", targets[0] if targets else Target(corpus), rename_to="", extraction_warning=(extra + " Renumeração exige indicação manual do dispositivo de origem e do novo rótulo.").strip(), **common))
                continue
        if re.search(r"\banexos?\b", cmd, re.I) and not [t for t in targets if t.article]:
            base_refs, own_refs = annexmod.command_annexes(cmd)
            mode = "revoke" if action == "revoke" else "add" if action == "add" else "replace"
            if base_refs and (mode == "revoke" or len(own_refs) == len(base_refs)):
                for k, bk in enumerate(base_refs):
                    op = Operation("", "annex", Target(corpus=f"ANEXO {bk}"), extraction_warning=extra.strip(), **common)
                    op.annex_mode, op.annex_key, op.annex_source = mode, bk, (own_refs[k] if own_refs else "")
                    operations.append(op)
                continue
            if base_refs or own_refs:
                operations.append(Operation("", action, Target(), extraction_warning=(extra + " Alteração de anexo sem correspondência clara entre o anexo da lei e o anexo do ato. Faça no Word.").strip(), **common))
                continue
        ref = headmod.command_reference(cmd)
        if ref and not targets and action == "replace" and re.search(r"denomin|t[ií]tulo\s+d[ao]|ep[ií]grafe|reda[cç][aã]o", cmd, re.I):
            kind, number, parents = ref
            group = headmod.find(base, kind, number, corpus, parents)
            quoted = re.findall(r"[“\"]([^”\"]+)[”\"]", cmd)
            words = quoted[-1:] if quoted else [clean_payload(l.text) for l in data if clean_payload(l.text) and not is_omission(clean_payload(l.text))]
            if words and headmod.parse_line(words[0]):
                lines_new = words
            elif words and group:
                lines_new = [group.lines[0]] + words
            else:
                lines_new = words
            op = Operation("", "heading", Target(corpus, group.anchor.article if group and group.anchor else ""), new_lines=lines_new,
                           extraction_warning=(extra if group else extra + " O título citado não foi localizado na base de forma única.").strip(),
                           **common)
            op.heading_key = f"{kind}:{number}"
            op.heading_parents = "/".join(f"{k}:{n}" for k, n in (group.parents if group else parents))
            operations.append(op)
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
                        refs_here, _ = annexmod.command_annexes(item_text)
                        if refs_here:
                            for bk in refs_here:
                                op = Operation("", "annex", Target(corpus=f"ANEXO {bk}"), extraction_warning=item_warning, **dict(common, foreign=item_foreign))
                                op.annex_mode, op.annex_key = "revoke", bk
                                operations.append(op)
                            continue
                        subtargets = [Target(corpus)]
                        subwarning += " Lista de revogação sem destino reconhecido."
                    warning = " ".join(x for x in (ref_warning, subwarning, item_warning) if x).strip()
                    if "condição" in extra or "evento" in extra:
                        warning = (warning + " " + extra).strip()
                    entries.extend((t, item_foreign, warning) for t in subtargets)
            else:
                chapter = heading_revocation(cmd, corpus, base) if not [t for t in targets if t.article] else None
                if chapter:
                    first, rest, name = chapter
                    op = Operation("", "revoke", first, scope="subtree", extra_targets=rest, extraction_warning=(extra + (
                        f" Revogação de {name} inteiro: o programa revoga os artigos dele"
                        f" ({first.label()}{' a ' + rest[-1].label() if rest else ''}); o título continua no texto."
                        " Confirme e clique em Validar correção.")).strip(), **common)
                    op.suppressed = suppressed
                    operations.append(op)
                    continue
                if not targets:
                    extra += " Não foi encontrado um destino explícito. Revogação genérica não é aplicada."
                    targets = [Target(corpus)]
                entries = [(t, foreign, extra.strip()) for t in targets]
            for group, item_foreign, warning in _revocation_groups(entries):
                op = Operation("", "revoke", group[0], scope="subtree", extra_targets=group[1:],
                               extraction_warning=warning, **dict(common, foreign=item_foreign))
                op.caput_path = caput_path(src, group[0])
                op.suppressed = suppressed
                operations.append(op)
            continue
        if any(x.kind == "table" for x in data):
            operations.append(Operation("", action, targets[0] if len(targets) == 1 else Target(corpus),
                                        extraction_warning="O comando inclui tabela. A revisão textual não aplica tabelas; faça essa alteração no Word. " + extra, **common))
            continue
        produced = False
        shifted = []
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
                op.caput_forced = prop.get("caput_forced", False)
                op.caput_path = caput_path(cmd, op.target)
                op.heading_key = prop.get("heading_key", "")
                op.heading_parents = prop.get("heading_parents", "")
                op.heading = list(prop.get("heading", []))
                op.suppressed = prop.get("suppressed", False)
                if prop.get("normalized"):
                    op.note = (op.note + " " if op.note else "") + SPACING_NOTE
                if prop.get("in_quote") and op.action == "revoke" and re.search(r"a partir da data em que|quando for|condicionad", ncmd):
                    op.extraction_warning += " A revogação depende de evento ou condição; confirme."
                if prop.get("shift"):
                    shifted.append(op)
                operations.append(op)
                produced = True
        if shifted and not pairs:
            # "renumerando-se os demais": renumera antes os irmãos que mudam de número.
            moves = shift_renames([t for o in shifted for t in o.targets], base)
            if moves:
                first = operations.index(shifted[0])
                operations[first:first] = rename_operations(moves, dict(common, source=shifted[0].source), extra)
        if not produced:
            operations.append(Operation("", action, targets[0] if len(targets) == 1 else Target(corpus),
                                        extraction_warning=(extra + " Não foi extraída uma redação completa identificada por rótulo.").strip(), **common))
    operations = _merge_adjacent(operations)
    operations = implicit_renumbering(operations, base)
    if not operations and not nested:
        text = "\n".join(l.text for l in lines)
        operations = [Operation("", "replace", Target(), source=text, source_location=line_location(lines) if lines else "",
                                extraction_warning="Nenhum comando de alteração suportado foi reconhecido. Cadastre os destinos e textos na revisão.")]
    for i, op in enumerate(operations, 1):
        op.id = f"A{i:03d}"
    return operations
