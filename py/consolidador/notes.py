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
"""
from __future__ import annotations
from collections import Counter
from datetime import date, timedelta
import re
from .model import Target, by_act, corpus_name, norm, ordinal
from .refs import enumerate_values

NOTE_CITATION = re.compile(r"\bpel[oa]s?\s+(?:art\.\s*\S+\s+d[ao]\s+)?(?:EC|ECR|LC|Lei|Decreto|Emenda|Instru[cç][aã]o Normativa|IN|Portaria|Resolu[cç][aã]o)\s*(?:n[º°o.]?\s*)?[\d.]+/(\d{2,4})\b", re.I)
NOTE_DATE = re.compile(r"\befeitos\s+(?:a\s+partir\s+)?de\s+\d{1,2}\.\d{1,2}\.(\d{2,4})\b", re.I)


def detect_date_style(blocks):
    """'2' ou '4' dígitos para o ano, conforme as notas dos atos alteradores mais recentes da base.

    As notas são ordenadas pelo ano do ato citado (e não pela data de efeitos, que
    pode estar no futuro); vence o formato mais usado nessas notas."""
    notes = []
    for b in blocks:
        if b.role not in ("new_note", "old_note"):
            continue
        cite = NOTE_CITATION.search(b.text)
        if not cite:
            continue
        digits = cite[1]
        year = int(digits) if len(digits) == 4 else 2000 + int(digits) if int(digits) < 70 else 1900 + int(digits)
        votes = [len(digits)] + [len(m[1]) for m in NOTE_DATE.finditer(b.text)]
        notes.append((year, b.index, votes))
    if not notes:
        # Sem notas: segue as referências do cabeçalho ("Alterada pelas Leis 6.012/96, ...").
        votes = Counter()
        for b in blocks[:18]:
            if re.search(r"alterad[ao]|atualizad[ao]|consolidad[ao]", norm(b.text)):
                votes.update(len(m[1]) for m in re.finditer(r"\b\d[\d.]*/(\d{2}|\d{4})\b", b.text))
        return "2" if votes[2] > votes[4] else "4"
    notes.sort(reverse=True)
    # Usa os atos do ano mais recente (ou dos anos seguintes, até reunir 5 notas):
    # a prática editorial muda com o tempo e o padrão atual é o que vale.
    chosen, years = [], []
    for year, index, votes in notes:
        if len(chosen) >= 5 and year not in years:
            break
        chosen.append(votes)
        if year not in years:
            years.append(year)
    count = Counter(v for votes in chosen for v in votes)
    return "2" if count[2] > count[4] else "4"


def _citation_year(text):
    cite = NOTE_CITATION.search(text)
    if not cite:
        return None
    digits = cite[1]
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


def effects_text(effects, style="4"):
    return ", efeitos a partir de " + format_date(effects, style) + "." if effects else "."


def until_text(effects, style="4"):
    return format_date((date.fromisoformat(effects) - timedelta(days=1)).isoformat(), style)


def _group_description(targets, verb_kind, corpus_suffix=True, caput_parent=False):
    first = targets[0]
    level = first.level
    plural = len(targets) > 1
    if not plural:
        desc = first.description(corpus_suffix=corpus_suffix)
    else:
        names = {0: "arts.", 1: "§§", 2: "incisos", 3: "alíneas", 4: "itens"}
        desc = names[level] + " " + enumerate_values([t.parts[level] for t in targets], level)
        if level > 0:
            parent = first.parent()
            desc += (" da " if parent.level == 3 else " do ") + parent.description(corpus_suffix=corpus_suffix)
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
    tail = f" {by_act(cite)} " + cite + effects_text(op.effects, style)
    targets = op.targets
    level = op.target.level
    plural = len(targets) > 1
    notes = []
    if op.renumber_from or op.action == "rename":
        source = Target(op.target.corpus, op.target.article, op.renumber_from or op.target.paragraph)
        new_number = op.target.paragraph if op.renumber_from else op.rename_to
        where = Target(op.target.corpus, op.target.article).description(corpus_suffix=corpus_suffix)
        old_name = "parágrafo único" if source.paragraph == "UNICO" else f"§ {ordinal(source.paragraph)}"
        notes.append(f"Renumerado o {old_name} para § {ordinal(new_number)} do {where}{tail}")
        if op.action == "rename":
            return notes
    if op.action == "replace":
        desc = op.target.description(caput=op.caput, corpus_suffix=corpus_suffix)
        if plural:
            desc = _group_description(targets, "replace", corpus_suffix)
        notes.append(article_for(level, plural, "replace") + desc + tail)
    elif op.action == "add":
        notes.append(article_for(level, plural, "add") + _group_description(targets, "add", corpus_suffix) + tail)
    elif op.action == "revoke":
        notes.append(article_for(level, plural, "revoke") + _group_description(targets, "revoke", corpus_suffix) + tail)
    return notes


def prior_note(text, effects, citation="", publication="", style="4"):
    """Fecha a vigência de uma nota que vai para o histórico."""
    if not effects or re.search(r"\bsem\s+efeitos\b|\befeitos\s+de\s+\d{1,2}[./]\d{1,2}[./]\d{2,4}\s+(?:a|até)\s+\d", text, re.I):
        return text
    until = until_text(effects, style)
    m = re.search(r"efeitos\s+a\s+partir\s+de\s+(\d{1,2}[./]\d{1,2}[./]\d{2,4})\.?\s*$", text, re.I)
    if m:
        from .model import date_iso
        try:
            start = date_iso(m[1])
        except ValueError:
            start = ""
        if start == effects and publication and publication < effects:
            return text[:m.start()] + f"com efeitos previstos para {m[1]}, substituída {by_act(citation)} {format_citation(citation, style)}, com efeitos a partir da mesma data."
        return text[:m.start()] + f"efeitos de {m[1]} a {until}."
    return re.sub(r"\.?\s*$", "", text) + f", efeitos até {until}."


def fallback_note(original, effects, style="4"):
    head = "Redação original" if original else "Redação anterior"
    return head + (f", efeitos até {until_text(effects, style)}." if effects else ".")
