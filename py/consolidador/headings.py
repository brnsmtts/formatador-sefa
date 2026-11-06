"""Títulos de Partes, Livros, Títulos, Capítulos, Seções e Subseções.

Padrões das bases SEFA:
  Redação dada ao título da Seção V pela EC 101/26, efeitos a partir de 30.04.26.
  SEÇÃO V / Da Justiça Militar                      (título vigente)
  Redação original, efeitos até 29.04.26.
  SEÇÃO V / Dos Conselhos de Justiça Militar        (título anterior, em verde)

  Acrescida a Seção V-A ao Capítulo I pela EC 132/23, efeitos a partir de 21.12.23.
  Seção V-A / Do Imposto de Competência Compartilhada ...
  Art. 156-A. ...                                   (sem nota própria)
"""
from __future__ import annotations
from dataclasses import dataclass, field
import re
from .model import Target, norm

# chave normalizada -> (nome, gênero, hierarquia)
KINDS = {"parte": ("Parte", "f", 0), "livro": ("Livro", "m", 1), "titulo": ("Título", "m", 2),
         "capitulo": ("Capítulo", "m", 3), "secao": ("Seção", "f", 4), "subsecao": ("Subseção", "f", 5)}
LINE = re.compile(r"^\s*[“\"]?\s*(PARTE|LIVRO|T[IÍ]TULO|CAP[IÍ]TULO|SE[CÇ][AÃ]O|SUBSE[CÇ][AÃ]O)\s+"
                  r"([IVXLCDM]+(?:\s*-\s*[A-Z])?|[UÚ]NIC[AO])(?![A-Za-zÀ-ÿ])\s*[-–—.:]?\s*(.*)$", re.I)
NAMED = r"(Parte|Livro|T[ií]tulo|Cap[ií]tulo|Se[cç][aã]o|Subse[cç][aã]o)\s+([IVXLCDM]+(?:-[A-Z])?|[uú]nic[ao])\b"
NAMED_NC = r"(?:Parte|Livro|T[ií]tulo|Cap[ií]tulo|Se[cç][aã]o|Subse[cç][aã]o)\s+(?:[IVXLCDM]+(?:-[A-Z])?|[uú]nic[ao])\b"
# Referência a título no comando: "o título da Seção V do Capítulo III do Título V",
# "a denominação do Capítulo IX", "A Seção V do Capítulo III ... passa a denominar-se".
COMMAND = re.compile(r"(?:\bo\s+t[ií]tulo\s+d[ao]\s+|\ba\s+denomina[cç][aã]o\s+d[ao]\s+|\ba\s+ep[ií]grafe\s+d[ao]\s+|^\s*(?:art\.?\s*\d+[º°o]?\.?\s*[-–]?\s*)?(?:o|a)\s+)"
                     + NAMED + r"((?:\s*,?\s*d[ao]\s+" + NAMED_NC + r")*)", re.I)
HEADING_NOTE = re.compile(r"^(?:Reda[cç][aã]o\s+dada|Acrescid[oa]s?|Renumerad[oa]s?|Revogad[oa]s?)\s+(?:[aà]o?s?\s+)?"
                          r"(?:t[ií]tulo\s+d[ao]\s+|denomina[cç][aã]o\s+d[ao]\s+)?(?:Parte|Livro|T[ií]tulo|Cap[ií]tulo|Se[cç][aã]o|Subse[cç][aã]o)\s+[IVXLCDM]", re.I)


def kind_key(word):
    return norm(word).replace(" ", "")


def number_key(value):
    value = re.sub(r"\s", "", value).upper()
    return "ÚNICA" if norm(value) in ("unica", "unico") else value


def parse_line(text):
    """'SEÇÃO V' / 'Capítulo IX - Dos Povos Indígenas' -> ('secao', 'V', resto)."""
    m = LINE.match(text or "")
    if not m:
        return None
    return kind_key(m[1]), number_key(m[2]), m[3].strip()


def name(kind, number):
    if number in ("ÚNICA", "ÚNICO"):
        number = "única" if KINDS[kind][1] == "f" else "único"
    return f"{KINDS[kind][0]} {number}"


def of(kind):
    return "da" if KINDS[kind][1] == "f" else "do"


def to(kind):
    return "à" if KINDS[kind][1] == "f" else "ao"


def replace_note(kind, number):
    return f"Redação dada ao título {of(kind)} {name(kind, number)}"


def add_note(kind, number, parent=None):
    fem = KINDS[kind][1] == "f"
    text = ("Acrescida a " if fem else "Acrescido o ") + name(kind, number)
    if parent:
        text += f" {to(parent[0])} {name(*parent)}"
    return text


@dataclass
class Group:
    kind: str
    number: str
    start: int
    end: int
    lines: list = field(default_factory=list)
    parents: list = field(default_factory=list)
    corpus: str = "PRINCIPAL"
    anchor: Target | None = None

    @property
    def label(self):
        return name(self.kind, self.number)


def groups(base):
    """Títulos da base, com a hierarquia de cada um e o primeiro artigo que os segue."""
    out, stack, corpus = [], {}, None
    blocks = base.blocks
    articles = sorted((n.start, n) for n in base.nodes if n.target.level == 0)
    i = 0
    while i < len(blocks):
        b = blocks[i]
        info = parse_line(b.text) if b.role == "heading" else None
        if not info:
            i += 1
            continue
        kind, number, rest = info
        j = i + 1
        lines = [b.text]
        while j < len(blocks) and blocks[j].role == "heading" and blocks[j].text and not parse_line(blocks[j].text):
            lines.append(blocks[j].text)
            j += 1
        anchor = next((n for s, n in articles if s >= j), None)
        here = anchor.target.corpus if anchor else "PRINCIPAL"
        if here != corpus:
            stack, corpus = {}, here
        rank = KINDS[kind][2]
        for r in [r for r in stack if r >= rank]:
            del stack[r]
        parents = [stack[r] for r in sorted(stack)]
        stack[rank] = (kind, number)
        out.append(Group(kind, number, i, j, lines, parents, here, anchor.target if anchor else None))
        i = j
    return out


def find(base, kind, number, corpus="PRINCIPAL", parents=(), article=""):
    """Um único título com esse nome (e hierarquia, ou artigo seguinte), ou None."""
    found = [g for g in groups(base) if g.kind == kind and g.number == number and g.corpus == corpus]
    for parent in parents:
        found = [g for g in found if parent in g.parents]
    if article:
        found = [g for g in found if g.anchor and g.anchor.article == article]
    return found[0] if len(found) == 1 else None


def parent_at(base, index, kind):
    """Título superior vigente na posição `index` (para "Acrescida a Seção V-A ao Capítulo I")."""
    rank = KINDS[kind][2]
    best = None
    for g in groups(base):
        if g.start >= index:
            break
        if KINDS[g.kind][2] < rank:
            best = (g.kind, g.number)
        elif KINDS[g.kind][2] == rank and g.parents:
            best = g.parents[-1]
    return best


def command_reference(text):
    """(tipo, número, [superiores]) de 'o título da Seção V do Capítulo III do Título V'."""
    m = COMMAND.search(text)
    if not m:
        return None
    parents = [(kind_key(k), number_key(n)) for k, n in re.findall(NAMED, m[3] or "", re.I)]
    return kind_key(m[1]), number_key(m[2]), parents
