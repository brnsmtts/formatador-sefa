from __future__ import annotations
from dataclasses import dataclass, field, asdict
from datetime import datetime
import hashlib
import re
import unicodedata


def norm(value: str) -> str:
    return " ".join(unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode().lower().split())


def digest(path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def ordinal(value: str) -> str:
    bits = value.split("-", 1)
    return bits[0] + ("º" if bits[0].isdigit() and int(bits[0]) <= 9 else "") + ("-" + bits[1] if len(bits) > 1 else "")


# Gênero do ato citado nas notas e remissões: "pelo Decreto 9.850/2026", "do Decreto ...";
# "pela Lei 99.500/26", "da EC 136/25", "pela IN 12/2026".
MASCULINE_ACT = re.compile(
    r"^(?:Decreto|Dec\.|DL\b|Ato\b|Ajuste|Conv[eê]nio|Protocolo|Parecer|Despacho|Regulamento|Regimento|"
    r"Provimento|Comunicado|Ac[oó]rd[aã]o|C[oó]digo|Edital|Termo|Aviso|Of[ií]cio|ADCT\b)", re.I)


def by_act(citation: str) -> str:
    """'pelo Decreto 9.850/2026', 'pela Lei 99.500/26'."""
    return "pelo" if MASCULINE_ACT.match(citation.strip()) else "pela"


def of_act(citation: str) -> str:
    """'Vide art. 4º do Decreto 9.850/2026', 'Vide art. 4º da Lei 99.500/26'."""
    return "do" if MASCULINE_ACT.match(citation.strip()) else "da"


def corpus_name(corpus: str) -> str:
    if corpus == "ADCT":
        return "ADCT"
    return corpus.title().replace("Unico", "Único")


def date_iso(value: str) -> str:
    value = value.strip()
    if not value:
        return ""
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d.%m.%Y", "%d/%m/%y", "%d.%m.%y"):
        try:
            return datetime.strptime(value, fmt).date().isoformat()
        except ValueError:
            pass
    raise ValueError("Use uma data válida, como 22/10/2026, ou deixe em branco.")


@dataclass(frozen=True)
class Target:
    corpus: str = "PRINCIPAL"
    article: str = ""
    paragraph: str = ""
    inciso: str = ""
    alinea: str = ""
    item: str = ""

    @property
    def level(self):
        return max((i for i, x in enumerate(self.parts) if x), default=-1)

    @property
    def parts(self):
        return (self.article, self.paragraph, self.inciso, self.alinea, self.item)

    def parent(self):
        parts = list(self.parts)
        if self.level >= 0:
            parts[self.level] = ""
        return Target(self.corpus, *parts)

    def contains(self, other):
        return self.corpus == other.corpus and all(not x or x == y for x, y in zip(self.parts, other.parts))

    def label(self):
        labels = [f"art. {self.article}"] if self.article else []
        if self.paragraph:
            labels.append("parágrafo único" if self.paragraph == "UNICO" else f"§ {self.paragraph}")
        for name, val in (("inciso", self.inciso), ("alínea", self.alinea), ("item", self.item)):
            if val:
                labels.append(f"{name} {val}")
        return " / ".join(labels)

    def description(self, caput=False, corpus_suffix=True, caput_path=False):
        text = f"art. {ordinal(self.article)}"
        if (caput and self.level == 0) or (caput_path and self.level >= 2 and not self.paragraph):
            # "inciso XVI do caput do art. 37": só quando o próprio ato o diz assim.
            text = "caput do " + text
        if self.paragraph:
            text = ("parágrafo único" if self.paragraph == "UNICO" else f"§ {ordinal(self.paragraph)}") + " do " + text
            if caput and self.level == 1:
                text = "caput do " + text
        if self.inciso:
            text = f"inciso {self.inciso} do {text}"
            if caput and self.level == 2:
                text = "caput do " + text
        if self.alinea:
            text = f"alínea “{self.alinea}” do {text}"
        if self.item:
            text = f"item {self.item} da {text}"
        if self.corpus != "PRINCIPAL" and corpus_suffix:
            text += f" do {corpus_name(self.corpus)}"
        return text

    @classmethod
    def from_dict(cls, value):
        return cls(**value)

    @classmethod
    def parse(cls, value, corpus="PRINCIPAL"):
        v = norm(value.replace('º','').replace('°','').replace('–','-'))
        a = re.search(r"\barts?\.?\s*(\d+(?:\s*[-–]\s*[a-z]+)?)", v)
        if not a:
            raise ValueError("Informe o artigo, por exemplo: art. 2 / § 1 / inciso II.")
        p = re.search(r"§§?\s*(\d+(?:\s*-\s*[a-z]+)?)", value.replace('º','').replace('°',''), re.I)
        i = re.search(r"\bincisos?\s+([ivxlcdm]+(?:-[a-z]+)?)\b", v)
        l = re.search(r"\balineas?\s+[\"'“”‘’]?([a-z])\b", v)
        n = re.search(r"\bite(?:m|ns)\s+(\d+(?:\.\d+)*)", v)
        return cls(corpus, re.sub(r"\s", "", a[1]).upper(),
                   "UNICO" if "paragrafo unico" in v else re.sub(r"\s", "", p[1]).upper() if p else "",
                   i[1].upper() if i else "", l[1] if l else "", n[1] if n else "")


@dataclass
class SourceLine:
    text: str
    index: int
    page: int | None = None
    style: str = ""
    kind: str = "paragraph"


@dataclass
class Operation:
    id: str
    action: str
    target: Target
    new_lines: list[str] = field(default_factory=list)
    scope: str = "head"
    source: str = ""
    source_location: str = ""
    effects: str = ""
    # Fim dos efeitos, quando o próprio ato limita o período ("efeitos de X a Y").
    effects_until: str = ""
    status: str = "revisar"
    reason: str = ""
    selected: bool = False
    before: list[str] = field(default_factory=list)
    base_location: str = ""
    foreign: bool = False
    manual: bool = False
    note: str = ""
    rename_to: str = ""
    applied: bool = False
    extraction_warning: str = ""
    publication: str = ""
    effects_source: str = ""
    caput: bool = False
    remissions: list[str] = field(default_factory=list)
    related_source: list[str] = field(default_factory=list)
    # Dispositivos irmãos tratados na mesma nota ("Acrescidos os §§ 23 a 30").
    extra_targets: list[Target] = field(default_factory=list)
    # Parágrafo único renumerado para § 1º na mesma mudança.
    renumber_from: str = ""
    # O ato transcreve a redação integral (sem linhas pontilhadas).
    full_text: bool = False
    # Novos filhos imediatos que acompanham a nova redação do caput.
    fold: bool = False
    # O dispositivo ganha filhos no mesmo ato (a nota fala em "caput").
    caput_forced: bool = False
    # O ato cita o inciso "do caput do art. X"; a nota repete essa forma.
    caput_path: bool = False
    # Título (Capítulo, Seção...): "secao:V" e superiores "titulo:V/capitulo:III".
    # Na ação "heading" é o título que muda; num acréscimo, o título novo que
    # antecede o artigo acrescido ("Acrescida a Seção V-A ao Capítulo I").
    heading_key: str = ""
    heading_parents: str = ""
    heading: list[str] = field(default_factory=list)
    # "Renumerada a Seção única para Seção I": só o rótulo do título muda.
    heading_rename: bool = False
    # Redações seguintes do mesmo dispositivo no mesmo ato, por período
    # (EC 132/23, LC 214/25): [{"effects": "2033-01-01", "effects_until": "", "lines": [...], "source": "..."}].
    later_versions: list[dict] = field(default_factory=list)
    # Ato alterador anterior cuja redação este ato alterou ("pela LC 214/25,
    # alterada pela LC 227/26"). Vazio na mudança direta.
    via: str = ""
    # Anexo inteiro: "replace", "add" ou "revoke"; rótulo na base e no ato; ato em DOCX.
    annex_mode: str = ""
    annex_key: str = ""
    annex_source: str = ""
    annex_path: str = ""
    # O ato diz "suprimido" em vez de "revogado" (Constituição do Pará, EC 37/2007).
    suppressed: bool = False
    source_article: str = ""
    source_item: str = ""

    @property
    def targets(self):
        return [self.target] + list(self.extra_targets)

    def display_label(self):
        if self.action == "annex":
            return f"Anexo {self.annex_key}" + (f" (Anexo {self.annex_source} do ato)" if self.annex_source and self.annex_source != self.annex_key else "")
        if self.heading_key and self.action == "heading":
            from .headings import name
            kind, _, number = self.heading_key.partition(":")
            return name(kind, number) + (f" (antes do art. {self.target.article})" if self.target.article else "")
        if self.later_versions:
            base = self.target.label() if not self.extra_targets else ""
            if base:
                return f"{base} ({len(self.later_versions) + 1} redações)"
        if not self.extra_targets:
            return self.target.label()
        from .refs import enumerate_values
        level = self.target.level
        values = [t.parts[level] for t in self.targets]
        names = {0: "arts.", 1: "§§", 2: "incisos", 3: "alíneas", 4: "itens"}
        parent = self.target.parent().label()
        group = f"{names[level]} {enumerate_values(values, level)}"
        return (parent + " / " if parent else "") + group

    def to_dict(self):
        return asdict(self)

    @classmethod
    def from_dict(cls, value):
        value = dict(value)
        value["target"] = Target.from_dict(value["target"])
        value["extra_targets"] = [Target.from_dict(t) if isinstance(t, dict) else t for t in value.get("extra_targets", [])]
        known = set(cls.__dataclass_fields__)
        return cls(**{k: v for k, v in value.items() if k in known})


@dataclass
class Plan:
    base_path: str
    amender_path: str
    converted_base: str
    citation: str
    publication: str = ""
    effects: str = ""
    base_hash: str = ""
    amender_hash: str = ""
    operations: list[Operation] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    style_profile: dict = field(default_factory=dict)
    identity: str = ""
    related_provisions: list[dict] = field(default_factory=list)
    # "auto" segue a base; "2" ou "4" fixam os dígitos do ano nas notas.
    date_style: str = "auto"
    resolved_date_style: str = "4"
    adct_suffix: bool = True
    converted_amender: str = ""

    def to_dict(self):
        return asdict(self)

    @classmethod
    def from_dict(cls, data):
        data = dict(data)
        data["operations"] = [Operation.from_dict(x) for x in data["operations"]]
        known = set(cls.__dataclass_fields__)
        return cls(**{k: v for k, v in data.items() if k in known})
