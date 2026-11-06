"""Redação das notas no padrão das consolidações SEFA.

Padrões observados nas legislações consolidadas (CF, Constituição do Pará,
Lei 5.530/89 e LC 123/06):
  Redação dada ao § 1º do art. 100 pela EC 136/25, efeitos a partir de 10.09.25.
  Redação dada ao caput do art. 18 pela LC 155/2016, efeitos a partir de 01.01.2018.
  Acrescido o § 19-A ao art. 100 pela EC 136/25, ...
  Acrescidos os §§ 23 a 30 ao art. 100 pela EC 136/25, ...
  Acrescida a alínea “g” ao inciso II do art. 78 pela Lei 11.233/25, ...
  Acrescido o art. 116-A ao ADCT pela EC 136/25, ...
  Revogados os incisos IV e V do art. 13 pela LC 214/25, ...
  Renumerado o parágrafo único para § 1º do art. 12 pela LC 214/2025, ...
  Redação original, efeitos até 31.12.26.
  Acrescidos os incisos IX e X ao art. 13 pela LC 214/2025, sem efeitos.
    (redação futura substituída ou revogada antes de produzir efeitos)
  Redação dada às alíneas “b” e “c” do art. 41 pela EC 100/26, ...
    (irmãos consecutivos com a mesma origem: uma nota e um histórico)
"""
from __future__ import annotations
from collections import Counter
from datetime import date, timedelta
import re
from .model import Target, by_act, corpus_name, norm, ordinal
from .refs import enumerate_values

NOTE_CITATION = re.compile(r"\bpel[oa]s?\s+(?:art\.\s*\S+\s+d[ao]\s+)?(?:EC|ECR|LC|Lei|Decreto|Emenda|Instru[cç][aã]o Normativa|IN|Portaria|Resolu[cç][aã]o)\s*(?:n[º°o.]?\s*)?(?P<number>[\d.]+)/(\d{2,4})\b", re.I)
NOTE_DATE = re.compile(r"\befeitos\s+(?:a\s+partir\s+)?de\s+\d{1,2}\.\d{1,2}\.(\d{2,4})\b", re.I)


def detect_date_style(blocks):
    """'2' ou '4' dígitos para o ano, conforme as notas dos atos alteradores mais recentes da base.

    As notas são ordenadas pelo ato citado (ano e número), e não pela data de
    efeitos, que pode estar no futuro. Valem as cinco notas mais recentes: a
    prática muda com o tempo (a Constituição do Pará usava 2025 e passou a 26)."""
    notes = []
    for b in blocks:
        if b.role not in ("new_note", "old_note"):
            continue
        cite = NOTE_CITATION.search(b.text)
        if not cite:
            continue
        digits = cite[2]
        year = int(digits) if len(digits) == 4 else 2000 + int(digits) if int(digits) < 70 else 1900 + int(digits)
        number = int(re.sub(r"\D", "", cite["number"]) or 0)
        votes = [len(digits)] + [len(m[1]) for m in NOTE_DATE.finditer(b.text)]
        notes.append((year, number, b.index, votes))
    if not notes:
        # Sem notas: segue as referências do cabeçalho ("Alterada pelas Leis 6.012/96, ...").
        votes = Counter()
        for b in blocks[:18]:
            if re.search(r"alterad[ao]|atualizad[ao]|consolidad[ao]", norm(b.text)):
                votes.update(len(m[1]) for m in re.finditer(r"\b\d[\d.]*/(\d{2}|\d{4})\b", b.text))
        return "2" if votes[2] > votes[4] else "4"
    notes.sort(reverse=True)
    recent = notes[:5]
    count = Counter(v for *_, votes in recent for v in votes)
    return "2" if count[2] > count[4] else "4"


def _citation_year(text):
    cite = NOTE_CITATION.search(text)
    if not cite:
        return None
    digits = cite[2]
    return int(digits) if len(digits) == 4 else 2000 + int(digits) if int(digits) < 70 else 1900 + int(digits)


def detect_adct_suffix(base):
    """As notas recentes do ADCT repetem 'do ADCT'? (Algumas bases não repetem.)"""
    notes = []
    adct_start = None
    for b in base.blocks:
        if b.role in ("heading", "current") and re.match(r"^(?:atos?\s+das?\s+)?disposicoes\s+(?:constitucionais\s+)?transitorias\b", norm(b.text)):
            adct_start = b.index
        elif adct_start is not None and b.role == "new_note" and norm(b.text).startswith(("redacao dada", "acrescid", "revogad")):
            year = _citation_year(b.text)
            if year:
                notes.append((year, b.index, b.text))
    if not notes:
        return True
    notes.sort(reverse=True)
    recent = notes[:10]
    marked = sum(bool(re.search(r"adct|transit", norm(t))) for _, _, t in recent)
    return marked * 2 >= len(recent)


def format_date(iso, style="4"):
    d = date.fromisoformat(iso)
    return d.strftime("%d.%m.%y" if style == "2" else "%d.%m.%Y")


def format_citation(citation, style="4"):
    """'LC 238/2026' -> 'LC 238/26' quando a base usa ano com 2 dígitos."""
    def fix(m):
        year = m[2]
        if style == "2" and len(year) == 4:
            year = year[2:]
        elif style == "4" and len(year) == 2:
            year = ("20" if int(year) < 70 else "19") + year
        return m[1] + year
    return re.sub(r"(\b\d[\d.]*/)(\d{4}|\d{2})\b", fix, citation.strip(), count=1)


def effects_text(effects, style="4", until=""):
    if effects and until:
        return ", efeitos de " + format_date(effects, style) + " a " + format_date(until, style) + "."
    return ", efeitos a partir de " + format_date(effects, style) + "." if effects else "."


def until_text(effects, style="4"):
    return format_date((date.fromisoformat(effects) - timedelta(days=1)).isoformat(), style)


def _group_description(targets, verb_kind, corpus_suffix=True, caput_path=False):
    first = targets[0]
    level = first.level
    plural = len(targets) > 1
    if not plural:
        desc = first.description(corpus_suffix=corpus_suffix, caput_path=caput_path)
    else:
        names = {0: "arts.", 1: "§§", 2: "incisos", 3: "alíneas", 4: "itens"}
        desc = names[level] + " " + enumerate_values([t.parts[level] for t in targets], level)
        if level > 0:
            parent = first.parent()
            parent_desc = parent.description(caput=caput_path and level == 2 and not first.paragraph,
                                             corpus_suffix=corpus_suffix, caput_path=caput_path)
            desc += (" da " if parent.level == 3 else " do ") + parent_desc
        elif first.corpus != "PRINCIPAL" and corpus_suffix:
            desc += " do " + corpus_name(first.corpus)
    if verb_kind == "add":
        if level == 0:
            if first.corpus != "PRINCIPAL" and corpus_suffix:
                desc = re.sub(r" do " + re.escape(corpus_name(first.corpus)) + "$", " ao " + corpus_name(first.corpus), desc)
        else:
            # "Acrescido o § 5º ao art. 2º": a primeira ligação passa a "ao"/"à".
            desc = re.sub(r" (do|da) ", lambda m: " ao " if m[1] == "do" else " à ", desc, count=1)
    return desc


def article_for(level, plural, verb):
    feminine = level == 3
    words = {
        ("add", False, False): "Acrescido o ", ("add", False, True): "Acrescida a ",
        ("add", True, False): "Acrescidos os ", ("add", True, True): "Acrescidas as ",
        ("revoke", False, False): "Revogado o ", ("revoke", False, True): "Revogada a ",
        ("revoke", True, False): "Revogados os ", ("revoke", True, True): "Revogadas as ",
        ("replace", False, False): "Redação dada ao ", ("replace", False, True): "Redação dada à ",
        ("replace", True, False): "Redação dada aos ", ("replace", True, True): "Redação dada às ",
    }
    return words[(verb, plural, feminine)]


def new_notes(op, citation, style="4", corpus_suffix=True):
    """Notas que antecedem o texto vigente. Uma renumeração com nova redação gera duas notas."""
    cite = format_citation(citation, style)
    if getattr(op, "via", ""):
        # Ato que altera a redação dada por outro ato alterador (LC 227/26 sobre a LC 214/25):
        # "pela LC 214/25, alterada pela LC 227/26"; "pelo Decreto 1/2026, alterado pelo Decreto 2/2026".
        changed = "alterado" if by_act(op.via) == "pelo" else "alterada"
        tail = (f" {by_act(op.via)} " + format_citation(op.via, style) + f", {changed} {by_act(cite)} " + cite
                + effects_text(op.effects, style, getattr(op, "effects_until", "")))
    else:
        tail = f" {by_act(cite)} " + cite + effects_text(op.effects, style, getattr(op, "effects_until", ""))
    targets = op.targets
    level = op.target.level
    plural = len(targets) > 1
    notes = []
    if op.action == "rename" and op.target.level >= 2:
        # "Renumerado o inciso II para inciso III do art. 12" (Lei 5.530); alínea no feminino.
        level = op.target.level
        unit = {2: "inciso", 3: "alínea", 4: "item"}[level]
        show = (lambda v: f"“{v}”") if level == 3 else (lambda v: v)
        where = op.target.parent().description(corpus_suffix=corpus_suffix)
        verb = "Renumerada a" if level == 3 else "Renumerado o"
        notes.append(f"{verb} {unit} {show(op.target.parts[level])} para {unit} {show(op.rename_to)} {'da' if op.target.parent().level == 3 else 'do'} {where}{tail}")
        return notes
    if op.renumber_from or op.action == "rename":
        source = Target(op.target.corpus, op.target.article, op.renumber_from or op.target.paragraph)
        new_number = op.target.paragraph if op.renumber_from else op.rename_to
        where = Target(op.target.corpus, op.target.article).description(corpus_suffix=corpus_suffix)
        old_name = "parágrafo único" if source.paragraph == "UNICO" else f"§ {ordinal(source.paragraph)}"
        notes.append(f"Renumerado o {old_name} para § {ordinal(new_number)} do {where}{tail}")
        if op.action == "rename":
            return notes
    path = getattr(op, "caput_path", False)
    if op.action == "annex":
        from .annexes import note
        notes.append(note(op.annex_mode, op.annex_key) + tail)
        return notes
    if op.action == "heading":
        from .headings import replace_note, name, KINDS, parse_line
        kind, _, number = op.heading_key.partition(":")
        if getattr(op, "heading_rename", False) and op.new_lines:
            fem = KINDS[kind][1] == "f"
            new = parse_line(op.new_lines[0])
            notes.append(f"{'Renumerada a' if fem else 'Renumerado o'} {name(kind, number).replace('Única', 'única').replace('Único', 'único')} para {name(kind, new[1])}" + tail)
            return notes
        notes.append(replace_note(kind, number) + tail)
        return notes
    if op.action == "add" and getattr(op, "heading", None) and op.heading_key:
        # Título novo com os seus artigos: a nota é a do título (CF, Seção V-A, EC 132/23).
        from .headings import add_note
        kind, _, number = op.heading_key.partition(":")
        parent = tuple(op.heading_parents.split(":", 1)) if op.heading_parents else None
        notes.append(add_note(kind, number, parent) + tail)
        return notes
    if op.action == "replace":
        desc = op.target.description(caput=op.caput, corpus_suffix=corpus_suffix, caput_path=path)
        if plural:
            desc = _group_description(targets, "replace", corpus_suffix, path)
        notes.append(article_for(level, plural, "replace") + desc + tail)
    elif op.action == "add":
        notes.append(article_for(level, plural, "add") + _group_description(targets, "add", corpus_suffix, path) + tail)
    elif op.action == "revoke":
        verb = article_for(level, plural, "revoke")
        if getattr(op, "suppressed", False):
            verb = verb.replace("Revogad", "Suprimid")
        notes.append(verb + _group_description(targets, "revoke", corpus_suffix, path) + tail)
    return notes


def version_blocks(op, citation, style="4", corpus_suffix=True):
    """Redações do mesmo dispositivo por período, da mais recente para a mais antiga
    (padrão da LC 123 consolidada com a LC 214/25):
      Redação dada ao § 3º do art. 31 pela LC 214/25, efeitos a partir de 01.01.33.
      § 3º (texto a partir de 2033)
      Redação dada ao § 3º do art. 31 pela LC 214/25, efeitos de 01.01.27 a 31.12.32.
      § 3º (texto de 2027 a 2032)
    Devolve [(notas, linhas)], na ordem do documento."""
    from copy import copy
    versions = [dict(action=op.action, effects=op.effects, effects_until=op.effects_until, lines=list(op.new_lines))]
    versions += [dict(action="replace", effects=v.get("effects", ""), effects_until=v.get("effects_until", ""), lines=list(v.get("lines", [])))
                 for v in op.later_versions]
    for i, v in enumerate(versions[:-1]):
        nxt = versions[i + 1]["effects"]
        if nxt:
            v["effects_until"] = (date.fromisoformat(nxt) - timedelta(days=1)).isoformat()
    out = []
    for k, v in enumerate(versions):
        tmp = copy(op)
        tmp.action, tmp.effects, tmp.effects_until, tmp.later_versions = v["action"], v["effects"], v["effects_until"], []
        if k:
            tmp.renumber_from = ""
        out.append((new_notes(tmp, citation, style, corpus_suffix), v["lines"]))
    return list(reversed(out))


def prior_note(text, effects, citation="", publication="", style="4"):
    """Fecha a vigência de uma nota que vai para o histórico.

    Redação futura que é substituída ou revogada antes de começar a produzir
    efeitos (publicação do novo ato antes do início, e novos efeitos até essa
    data) recebe "sem efeitos", como nas bases SEFA desde 2025."""
    if not effects or re.search(r"\bsem\s+efeitos\b|\befeitos\s+de\s+\d{1,2}[./]\d{1,2}[./]\d{2,4}\s+(?:a|até)\s+\d", text, re.I):
        return text
    until = until_text(effects, style)
    m = re.search(r"efeitos\s+a\s+partir\s+de\s+(\d{1,2}[./]\d{1,2}[./]\s?\d{2,4})\.?\s*$", text, re.I)
    if m:
        from .model import date_iso
        try:
            start = date_iso(m[1].replace(" ", ""))
        except ValueError:
            start = ""
        if start and effects <= start and publication and publication < start:
            head = re.sub(r"[,\s]*(?:com\s*)?$", "", text[:m.start()], flags=re.I)
            return head + ", sem efeitos."
        if start and effects <= start:
            return text
        return text[:m.start()] + f"efeitos de {m[1]} a {until}."
    return re.sub(r"\.?\s*$", "", text) + f", efeitos até {until}."


def fallback_note(original, effects, style="4"):
    head = "Redação original" if original else "Redação anterior"
    return head + (f", efeitos até {until_text(effects, style)}." if effects else ".")
