"""Anexos inteiros (tabelas e textos): nova redação, acréscimo e revogação.

Padrões da LC 123 consolidada pela SEFA:
  Redação dada ao Anexo I pela LC 214/25, efeitos a partir de 01.01.27.
  ANEXO I / título / tabelas ...                     (anexo vigente)
  Redação dada ao Anexo I pela LC 155/16, efeitos de 01.01.18 a 31.12.26.
  ANEXO I / título / tabelas (estilos "ant")         (anexo anterior)

  Acrescido o Anexo VII pela LC 214/25, efeitos a partir de 01.01.27.
  Revogado o Anexo VI pela LC 155/2016, efeitos a partir de 01.01.2018.
  ANEXO VI - REVOGADO
"""
from __future__ import annotations
from copy import deepcopy
import re
from docx.oxml.ns import qn
from .model import norm

NUMBER = r"([IVXLCDM]+|\d+|[UÚ]NICO)"
LABEL = re.compile(r"^\s*ANEXO\s+" + NUMBER + r"(?![\w])\s*(?:[-–—.:]\s*(.*))?$", re.I)
# "(Anexo I da Lei Complementar nº 123, de 14 de dezembro de 2006)": indica o destino; não entra no texto.
TARGET_LINE = re.compile(r"^\s*\(?\s*Anexo\s+\S+\s+(?:à|a|ao|d[ao]s?)\s+(?:Lei|Decreto|LC|Emenda|Instru|Resolu|Portaria|Constitui)", re.I)
REF = re.compile(r"\banexos?\s+" + NUMBER + r"(?:\s+(?:a|e)\s+" + NUMBER + r")?", re.I)
OWN = re.compile(r"^[\s,]*(?:\(|,)?\s*(?:dest[ae]|a\s+est[ae]|ao\s+present[ae]|do\s+present[ae]|anexos?\s+(?:a|à)\s+est[ae])\b", re.I)
ROMAN = {"I": 1, "V": 5, "X": 10, "L": 50, "C": 100, "D": 500, "M": 1000}
ANT = {"a15_tabela_centro": "a19_tabela_centroant", "a16_tabela_justificado": "a22_tabela_justificadoant",
       "a17_tabela_esquerda": "a21_tabela_esquerdant", "a18_tabela_direita": "a20_tabela_direitant"}


def key(value):
    value = value.strip().upper()
    return "ÚNICO" if norm(value) == "unico" else value


def value(k):
    if k.isdigit():
        return int(k)
    if k == "ÚNICO":
        return 1
    total = 0
    for i, c in enumerate(k):
        v = ROMAN.get(c, 0)
        total += -v if i + 1 < len(k) and v < ROMAN.get(k[i + 1], 0) else v
    return total


def label_of(text):
    text = (text or "").strip()
    if len(text) > 90:
        return None
    m = LABEL.match(text)
    return key(m[1]) if m else None


def expand(a, b):
    """'I a V' -> I, II, III, IV, V (romanos ou arábicos)."""
    if not b:
        return [key(a)]
    lo, hi = value(key(a)), value(key(b))
    if key(a).isdigit():
        return [str(n) for n in range(lo, hi + 1)]
    out = []
    for n in range(lo, hi + 1):
        r, rest = "", n
        for v, s in ((1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
                     (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")):
            while rest >= v:
                r += s
                rest -= v
        out.append(r)
    return out


def command_annexes(cmd):
    """Anexos citados num comando: (da base, deste ato). 'O Anexo I da LC 123 passa a
    vigorar na forma do Anexo I desta Lei Complementar.'"""
    base, own = [], []
    for m in REF.finditer(cmd):
        items = [key(m[1])] if " e " in m[0].lower() else expand(m[1], m[2])
        if " e " in m[0].lower() and m[2]:
            items.append(key(m[2]))
        (own if OWN.match(cmd[m.end():m.end() + 40]) else base).extend(items)
    return base, own


class Region:
    def __init__(self, k, lead, start, current_end, end):
        self.key, self.lead, self.start, self.current_end, self.end = k, lead, start, current_end, end


def regions(base):
    """Anexos da base: notas acima, rótulo, parte vigente e fim (antes do anexo seguinte)."""
    blocks = base.blocks
    labels = [i for i, b in enumerate(blocks) if b.role not in ("old", "old_note", "new_note") and label_of(b.text)]
    out = {}
    for n, i in enumerate(labels):
        lead = i
        while lead > 0 and blocks[lead - 1].role == "new_note":
            lead -= 1
        nxt = labels[n + 1] if n + 1 < len(labels) else len(blocks)
        if nxt < len(blocks):
            while nxt > i + 1 and blocks[nxt - 1].role == "new_note":
                nxt -= 1
        if blocks and blocks[-1].element.tag == qn("w:sectPr"):
            nxt = min(nxt, len(blocks) - 1)
        current_end = i + 1
        while current_end < nxt and blocks[current_end].role not in ("old", "old_note"):
            current_end += 1
        k = label_of(blocks[i].text)
        if k in out:
            out[k] = None  # rótulo repetido: não é seguro escolher
        else:
            out[k] = Region(k, lead, i, current_end, nxt)
    return {k: r for k, r in out.items() if r is not None}


def act_annexes(doc):
    """Anexos do ato alterador (DOCX): rótulo -> elementos do conteúdo."""
    from .index import paragraph_text
    body = [el for el in doc.element.body if el.tag in (qn("w:p"), qn("w:tbl"))]
    marks = []
    for i, el in enumerate(body):
        if el.tag == qn("w:p"):
            k = label_of(paragraph_text(el))
            if k:
                marks.append((i, k))
    out = {}
    for n, (i, k) in enumerate(marks):
        end = marks[n + 1][0] if n + 1 < len(marks) else len(body)
        content = []
        for el in body[i + 1:end]:
            if el.tag == qn("w:p") and TARGET_LINE.match(paragraph_text(el)):
                continue
            content.append(el)
        while content and content[-1].tag == qn("w:p") and not paragraph_text(content[-1]).strip():
            content.pop()
        out[k] = content
    return out


def has_external_parts(elements):
    """Imagens, objetos ou links dependem de relações do pacote do ato: não são copiados."""
    for el in elements:
        for node in el.iter():
            for attr in node.attrib:
                if attr in (qn("r:embed"), qn("r:id"), qn("r:link"), qn("r:pict")):
                    return True
    return False


def note(mode, k):
    return {"replace": f"Redação dada ao Anexo {k}", "add": f"Acrescido o Anexo {k}", "revoke": f"Revogado o Anexo {k}"}[mode]
