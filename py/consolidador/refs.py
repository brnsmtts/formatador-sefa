"""Leitura de referências a dispositivos ("o inciso III do art. 2º e o § 4º do art. 5º").

Reconhece cadeias do menor para o maior ("alínea “a” do inciso II do § 1º do art. 5º"),
do maior para o menor ("art. 18-A, §§ 1º e 2º"), listas ("§§ 1º e 2º"), intervalos
("incisos I a III", "arts. 39-A a 39-C"), "caput" e "parágrafo único". Cada referência
fica separada, sem combinar números de dispositivos diferentes.
"""
from __future__ import annotations
from dataclasses import dataclass, field
import re
from .model import Target

LEVEL = {"article": 0, "caput": 0.5, "paragraph": 1, "inciso": 2, "alinea": 3, "item": 4}
NUMV = r"\d+(?:\s?[º°o](?![a-zà-ú]))?(?:(?:\s[-–]|[-–]\s?)[A-Z]{1,3}(?![A-Za-zÀ-ú]))?"
ROMV = r"[IVXLCDM]+(?:[-–][A-Z])?(?![A-Za-zÀ-ú])"
LETV = r"[\"“”'‘’]?[a-z][\"“”'‘’]?(?![A-Za-zÀ-ú])"
ITMV = r"\d+(?:\.\d+)*"
SEP = r"(?:\s*,\s*(?:e\s+)?|\s+e\s+|\s+(?:a|ao|até)\s+)"
UNITS = [
    ("paragraph_unique", r"\bpar[aá]grafos?\s+[uú]nicos?\b", None),
    ("article", r"\b(?i:art(?:igo)?s?)\.?\s*", NUMV),
    ("paragraph", r"(?:§§?\s*|\bpar[aá]grafos?\s+)", NUMV),
    ("inciso", r"\b[Ii]ncisos?\s+", ROMV),
    ("alinea", r"\b[Aa]l[ií]neas?\s+", LETV),
    ("item", r"\b[Ii]te(?:m|ns)\s+", ITMV),
    ("caput", r"\bcaput\b", None),
]
SMALL_TO_LARGE = re.compile(r"^\s*,?\s*(?:d[oa]s?|a[os]?|à|às|n[oa]s?|em)\s+(?:mesmo\s+|referido\s+)?$", re.I)
LARGE_TO_SMALL = re.compile(r"^\s*,?\s*$")
LAW = re.compile(r"^\s*,?\s*(?:(?:d[oa]s?|n[oa]s?|a[os]?|à|às)\s+(?:mesm[ao]\s+)?"
                 r"((?:Lei(?:\s+Complementar)?|Decreto(?:-Lei)?|Constitui[cç][aã]o|Emenda|EC|LC|Instru[cç][aã]o\s+Normativa|IN|Portaria|Resolu[cç][aã]o|Conv[eê]nio|Ato\s+das|ADCT|Regulamento|RICMS|C[oó]digo)\b[^;]{0,120})"
                 r"|((?:est[ae]|dest[ae]|nest[ae]|dess[ae])\s+(?:Lei|Decreto|Emenda|Instru[cç][aã]o|Portaria|Resolu[cç][aã]o|Ato|artigo|par[aá]grafo|inciso|al[ií]nea|Constitui[cç][aã]o|Regulamento|Anexo)\w*))", re.I)


def roman(n):
    out = ""
    for v, s in ((1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")):
        while n >= v:
            out += s
            n -= v
    return out


def roman_value(s):
    d = {"I": 1, "V": 5, "X": 10, "L": 50, "C": 100, "D": 500, "M": 1000}
    s = s.upper()
    if not s or any(c not in d for c in s):
        return None
    return sum(-d[c] if i + 1 < len(s) and d[c] < d[s[i + 1]] else d[c] for i, c in enumerate(s))


def clean_value(kind, raw):
    v = raw.strip().strip("\"“”'‘’")
    v = re.sub(r"\s+", "", v).replace("–", "-")
    if kind in ("article", "paragraph"):
        v = re.sub(r"(?<=\d)[º°o]", "", v).upper()
    elif kind == "inciso":
        v = v.upper()
    elif kind == "alinea":
        v = v.lower()
    return v


@dataclass
class Unit:
    kind: str
    values: list
    start: int
    end: int
    range_warning: str = ""

    @property
    def level(self):
        return LEVEL["paragraph" if self.kind == "paragraph_unique" else self.kind]


@dataclass
class Reference:
    parts: dict = field(default_factory=dict)
    caput: bool = False
    start: int = 0
    end: int = 0
    law: str = ""
    warning: str = ""

    def targets(self, corpus="PRINCIPAL"):
        lists = {k: v for k, v in self.parts.items() if v}
        multi = [k for k, v in lists.items() if len(v) > 1]
        if not lists.get("article"):
            return []
        result = [{}]
        for k in ("article", "paragraph", "inciso", "alinea", "item"):
            vals = lists.get(k, [""])
            result = [dict(r, **{k: v}) for r in result for v in vals]
        if len(multi) > 1:
            self.warning = "Há mais de uma lista de dispositivos; confira a correspondência entre os níveis."
        return [Target(corpus, r["article"], r["paragraph"], r["inciso"], r["alinea"], r["item"]) for r in result]

    @property
    def top(self):
        levels = [LEVEL[k] for k, v in self.parts.items() if v]
        if self.caput:
            levels.append(0.5)
        return min(levels) if levels else 9

    @property
    def bottom(self):
        levels = [LEVEL[k] for k, v in self.parts.items() if v]
        if self.caput:
            levels.append(0.5)
        return max(levels) if levels else -1


def _expand(kind, values, raw):
    """Expande intervalos ("I a III", "39-A a 39-C"); devolve (valores, aviso)."""
    pieces = re.split(f"({SEP})", raw)
    tokens = []
    for k, p in enumerate(pieces):
        if k % 2:
            tokens.append(("RANGE", p) if re.fullmatch(r"\s+(?:a|ao|até)\s+", p) else ("SEP", p))
        elif p.strip():
            tokens.append(("V", clean_value(kind, p)))
    out = []
    i = 0
    while i < len(tokens):
        kind_t, v = tokens[i]
        if kind_t == "V" and i + 2 < len(tokens) and tokens[i + 1][0] == "RANGE" and tokens[i + 2][0] == "V":
            seq, warning = _range(kind, v, tokens[i + 2][1])
            if warning:
                return [x[1] for x in tokens if x[0] == "V"], warning
            out.extend(seq)
            i += 3
            continue
        if kind_t == "V":
            out.append(v)
        i += 1
    return out, ""


def _range(kind, a, b):
    if kind == "alinea":
        if len(a) == 1 and len(b) == 1 and a < b:
            return [chr(c) for c in range(ord(a), ord(b) + 1)], ""
        return [], "Intervalo de alíneas precisa ser discriminado na revisão."
    if kind == "item":
        if a.isdigit() and b.isdigit() and 0 < int(b) - int(a) < 100:
            return [str(x) for x in range(int(a), int(b) + 1)], ""
        return [], "Intervalo de itens precisa ser discriminado na revisão."
    ma = re.fullmatch(r"([0-9IVXLCDM]+)(?:-([A-Z]{1,3}))?", a)
    mb = re.fullmatch(r"([0-9IVXLCDM]+)(?:-([A-Z]{1,3}))?", b)
    if not ma or not mb:
        return [], "Intervalo de dispositivos precisa ser discriminado na revisão."
    if ma[1] == mb[1] and ma[2] and mb[2] and len(ma[2]) == len(mb[2]) == 1 and ma[2] < mb[2]:
        # arts. 39-A a 39-C
        return [f"{ma[1]}-{chr(c)}" for c in range(ord(ma[2]), ord(mb[2]) + 1)], ""
    if ma[2] or mb[2]:
        return [], "Intervalo com dispositivos de letra precisa ser discriminado na revisão."
    if kind == "inciso":
        va, vb = roman_value(ma[1]), roman_value(mb[1])
    else:
        va, vb = (int(ma[1]), int(mb[1])) if ma[1].isdigit() and mb[1].isdigit() else (None, None)
    if va is None or vb is None or not 0 < vb - va < 100:
        return [], "Intervalo de dispositivos muito amplo ou inválido."
    return [roman(x) if kind == "inciso" else str(x) for x in range(va, vb + 1)], ""


def _value_pattern(kind):
    return {"article": NUMV, "paragraph": NUMV, "inciso": ROMV, "alinea": LETV, "item": ITMV}[kind]


def units(text):
    text = text.replace(" ", " ")
    found = []
    pos = 0
    combined = "|".join(f"(?P<{k}>{p})" for k, p, _ in UNITS)
    for m in re.finditer(combined, text):
        if m.start() < pos:
            continue
        kind = m.lastgroup
        if kind == "paragraph_unique":
            found.append(Unit("paragraph_unique", ["UNICO"], m.start(), m.end()))
            pos = m.end()
            continue
        if kind == "caput":
            found.append(Unit("caput", [], m.start(), m.end()))
            pos = m.end()
            continue
        vp = _value_pattern(kind)
        vm = re.match(rf"(?:{vp})(?:{SEP}(?:{vp}))*", text[m.end():])
        if not vm:
            continue
        raw = vm[0]
        # Um "a" final seguido de outra palavra não é intervalo.
        values, warning = _expand(kind, [], raw)
        if not values:
            continue
        found.append(Unit(kind, values, m.start(), m.end() + vm.end(), warning))
        pos = m.end() + vm.end()
    return found


def parse(text):
    """Lista de Reference na ordem do texto."""
    us = units(text)
    refs = []
    current = None
    for i, u in enumerate(us):
        joined = False
        if current is not None:
            between = text[us[i - 1].end:u.start]
            prev = us[i - 1]
            if SMALL_TO_LARGE.match(between) and (u.level < prev.level or prev.kind == "caput"):
                # "caput do § 1º", "caput do inciso V do § 3º": o caput é do dispositivo seguinte.
                joined = True
            elif LARGE_TO_SMALL.match(between) and u.level > prev.level:
                joined = True
            if joined:
                key = "paragraph" if u.kind == "paragraph_unique" else u.kind
                if key == "caput":
                    current.caput = True
                elif current.parts.get(key):
                    joined = False
                else:
                    current.parts[key] = u.values
                    current.end = u.end
                    if u.range_warning:
                        current.warning = u.range_warning
        if not joined:
            current = Reference(start=u.start, end=u.end, warning=u.range_warning)
            key = "paragraph" if u.kind == "paragraph_unique" else u.kind
            if key == "caput":
                current.caput = True
            else:
                current.parts[key] = u.values
            refs.append(current)
    for r in refs:
        m = LAW.match(text[r.end:])
        if m:
            r.law = (m[1] or m[2]).strip()
    # Referências sem artigo herdam só o artigo da referência seguinte
    # ("o inciso III do caput e o § 2º do art. 5º") ou, se não houver, da anterior.
    for i, r in enumerate(refs):
        if r.parts.get("article"):
            continue
        donor = next((x for x in refs[i + 1:] if x.parts.get("article")), None)
        if donor is None:
            donor = next((x for x in reversed(refs[:i]) if x.parts.get("article")), None)
        if donor is None:
            continue
        r.parts["article"] = donor.parts["article"][:1]
        if not r.law:
            r.law = donor.law
    return refs


def targets(text, corpus="PRINCIPAL", self_words=False):
    """Destinos explícitos e aviso. Ignora referências ao próprio ato ("desta Lei")."""
    result, warnings = [], []
    for r in parse(text):
        if not self_words and is_self(r.law):
            continue
        ts = r.targets(corpus)
        if r.warning:
            warnings.append(r.warning)
        result.extend(t for t in ts if t not in result)
    return result, " ".join(dict.fromkeys(warnings))


def is_self(law):
    return bool(re.match(r"(?:est[ae]|dest[ae]|nest[ae]|dess[ae])\s", law or "", re.I))


def ordinal(value):
    head, _, tail = value.partition("-")
    text = head + ("º" if head.isdigit() and int(head) <= 9 else "")
    return text + ("-" + tail if tail else "")


def enumerate_values(values, level):
    """'§§ 23 a 30' para sequências contínuas de 3+; 'X e Y' para dois; lista com vírgulas."""
    if level == 3:
        shown = [f"“{v}”" for v in values]
    elif level < 2:
        shown = [ordinal(v) for v in values]
    else:
        shown = list(values)
    if len(values) >= 3 and _consecutive(values, level):
        return f"{shown[0]} a {shown[-1]}"
    if len(shown) == 1:
        return shown[0]
    return ", ".join(shown[:-1]) + " e " + shown[-1]


def _consecutive(values, level):
    def key(v):
        head, _, tail = v.partition("-")
        if level == 2:
            n = roman_value(head)
        elif level == 3:
            n = ord(v[0])
        else:
            n = int(head) if head.isdigit() else None
        return n, tail
    keys = [key(v) for v in values]
    if any(k[0] is None for k in keys):
        return False
    if all(not k[1] for k in keys):
        return all(b[0] - a[0] == 1 for a, b in zip(keys, keys[1:]))
    if all(k[0] == keys[0][0] and len(k[1]) == 1 for k in keys):
        return all(ord(b[1]) - ord(a[1]) == 1 for a, b in zip(keys, keys[1:]))
    return False
