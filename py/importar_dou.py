"""Importa um ato pelo link da página do DOU (in.gov.br), sem baixar o PDF do diário.

A página do ato no DOU ("https://www.in.gov.br/web/dou/-/medida-provisoria-n-1.399-de-4-de-novembro-de-2026-741754451")
traz o texto já em parágrafos, as tabelas com as células mescladas e as figuras, além de "Publicado em", "Edição",
"Seção", "Página" e "Órgão". O programa lê essa página e monta o mesmo resultado da importação do PDF (prévia,
relatório, Word no modelo da SEFA e hyperlinks): os parágrafos passam pelo mesmo classificador; as tabelas e as
figuras entram no lugar em que aparecem.

Também aceita a página salva no computador (no navegador, Ctrl+S, "Página da Web, arquivo único (.mhtml)", que
traz as figuras, ou "Página da Web, somente HTML"), para quando a rede não deixar o programa abrir o in.gov.br.
"""

from __future__ import annotations

import gzip
import io
import re
import urllib.error
import urllib.parse
import urllib.request
import zlib
from copy import deepcopy
from dataclasses import dataclass, field
from html.parser import HTMLParser
from pathlib import Path
from typing import Callable

from extrator_pdf import (LEGEND_HEAD, LEGEND_ITEM, TABLE_NOTE, Extraction, ImageBlock, ParagraphBlock, TableBlock,
                          TablePagePart, _mostly_uppercase, add_cgibs_availability, date_constitutional_amendment,
                          matrix_signature, monetary_values, normalized)
from formatador_sefa import ANNEX, ARTICLE, HEADING, Item, classify, looks_uppercase

try:      # rede da SEFA: certificados do Windows (como no programa dos hyperlinks)
    import truststore
    truststore.inject_into_ssl()
except Exception:
    pass

DOU_LINK = re.compile(r"^https?://(?:www\.)?in\.gov\.br/(?:en/)?web/dou/-/\S+$", re.I)
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
      "Chrome/140.0.0.0 Safari/537.36 Edg/140.0.0.0")
AVISO_FINAL = re.compile(r"^Este conteúdo não substitui o publicado na versão certificada", re.I)
BLOCOS = {"p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "li", "blockquote", "section", "article", "center"}
ESTRUTURA = BLOCOS | {"table", "img", "ul", "ol"}
VAZIOS = {"br", "img", "hr", "meta", "link", "input", "col", "source", "wbr"}
IGNORAR = {"script", "style", "noscript", "button", "form", "nav", "header", "footer", "template", "iframe", "svg"}
# Na página do DOU, o espaço antes e depois de um trecho em itálico ou negrito às vezes fica de fora do HTML
# ("no<i>caput</i>,", "alíquotas<i>ad valorem</i>do", "<b>O</b><b>PRESIDENTE</b>"); nessas emendas o espaço é reposto.
FORMATO = {"i", "em", "b", "strong", "u", "sub", "sup", "span", "font", "a", "s", "strike", "small", "big", "mark",
           "cite", "abbr", "ins", "del", "q", "tt", "code", "label"}


def _precisa_espaco(esquerda: str, direita: str, tag_direita: str | None) -> bool:
    """Entre dois trechos colados numa emenda de itálico/negrito, falta espaço?"""
    if not esquerda or not direita:
        return False
    a, b = esquerda[-1], direita[0]
    if a.isspace() or b.isspace() or a == "\n" or b == "\n":
        return False
    if tag_direita in ("sub", "sup") or b in "ºª°":      # CO2, m³, 1º: sem espaço antes
        return False
    if a.isalnum():
        return b.isalnum() or b == "("
    if a in ",;:)":
        return b.isalpha()
    if a == "." and b.isalpha() and b.isupper():
        return True
    if a in "-–—+=×÷<>≤≥" and len(esquerda) >= 2 and esquerda[-2] == " ":   # "operação - split", "6,3 + Ajustes"
        return b.isalnum()
    return False


def _texto_de(filhos: list) -> str:
    """Texto de uma sequência de nós: espaços do código HTML viram um espaço, <br> e blocos viram quebra de linha,
    e as emendas de itálico/negrito sem espaço recebem o espaço que falta."""
    partes: list[str] = []
    cauda = ""                      # últimos caracteres já escritos
    anterior_formato = False
    for filho in filhos:
        tag = None
        if isinstance(filho, str):
            pedaco = re.sub(r"\s+", " ", filho)
        elif filho.tag == "br":
            partes.append("\n")
            cauda, anterior_formato = "\n", False
            continue
        elif filho.tag in IGNORAR:
            continue
        elif filho.tag in BLOCOS or filho.tag in ("table", "tr", "ul", "ol"):
            if partes and not cauda.endswith("\n"):
                partes.append("\n")
            partes.append(filho.texto())
            partes.append("\n")
            cauda, anterior_formato = "\n", False
            continue
        else:
            tag = filho.tag
            pedaco = filho.texto()
        if not pedaco:
            continue
        formato = tag in FORMATO
        if (formato or anterior_formato) and _precisa_espaco(cauda, pedaco, tag):
            pedaco = " " + pedaco
        partes.append(pedaco)
        cauda = (cauda + pedaco)[-3:]
        anterior_formato = formato
    return "".join(partes)


def eh_link_dou(texto: str) -> bool:
    return bool(DOU_LINK.match(texto.strip()))


# ----------------------------------------------------------------------------------------------------------------
# HTML -> árvore simples
# ----------------------------------------------------------------------------------------------------------------

@dataclass
class No:
    tag: str
    attrs: dict[str, str] = field(default_factory=dict)
    filhos: list = field(default_factory=list)     # No ou str
    pai: "No | None" = None

    def classes(self) -> set[str]:
        return set(self.attrs.get("class", "").split())

    def texto(self) -> str:
        return _texto_de(self.filhos)

    def tem_estrutura(self) -> bool:
        """Há parágrafo, tabela ou figura dentro deste nó (em qualquer nível)?"""
        if "_estrutura" not in self.__dict__:
            self.__dict__["_estrutura"] = any(
                isinstance(f, No) and f.tag not in IGNORAR and (f.tag in ESTRUTURA or f.tem_estrutura())
                for f in self.filhos)
        return self.__dict__["_estrutura"]

    def todos(self):
        for filho in self.filhos:
            if isinstance(filho, No):
                yield filho
                yield from filho.todos()

    def procurar(self, condicao: Callable[["No"], bool]) -> "No | None":
        return next((no for no in self.todos() if condicao(no)), None)


class _Montador(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.raiz = No("documento")
        self.atual = self.raiz

    def handle_starttag(self, tag, attrs):
        no = No(tag, {k: (v or "") for k, v in attrs}, [], self.atual)
        self.atual.filhos.append(no)
        if tag not in VAZIOS:
            self.atual = no

    def handle_startendtag(self, tag, attrs):
        self.atual.filhos.append(No(tag, {k: (v or "") for k, v in attrs}, [], self.atual))

    def handle_endtag(self, tag):
        no = self.atual
        while no is not None and no.tag != tag:
            no = no.pai
        if no is not None and no.pai is not None:
            self.atual = no.pai

    def handle_data(self, data):
        self.atual.filhos.append(data)


def _de_lxml(elemento, pai: "No | None") -> No:
    no = No(elemento.tag.lower(), {str(k): (v or "") for k, v in elemento.attrib.items()}, [], pai)
    if elemento.text:
        no.filhos.append(elemento.text)
    for filho in elemento:
        if isinstance(filho.tag, str):         # comentários e instruções ficam de fora (o texto depois deles, não)
            no.filhos.append(_de_lxml(filho, no))
        if filho.tail:
            no.filhos.append(filho.tail)
    return no


def arvore(html: str) -> No:
    """Monta a árvore como o navegador: o lxml (que já vem com o python-docx) fecha sozinho o <p> antes de uma tabela
    e corrige as marcas mal fechadas. Sem ele, um montador simples."""
    try:
        import lxml.html
        texto = re.sub(r"^\s*<\?xml[^>]*\?>", "", html)
        documento = lxml.html.document_fromstring(texto.encode("utf-8"),
                                                  parser=lxml.html.HTMLParser(encoding="utf-8", remove_comments=True))
        raiz = No("documento")
        raiz.filhos.append(_de_lxml(documento, raiz))
        return raiz
    except Exception:                  # noqa: BLE001  (lxml ausente, página vazia ou aninhamento extremo)
        montador = _Montador()
        montador.feed(html)
        montador.close()
        return montador.raiz


def limpar(texto: str) -> str:
    texto = texto.replace("\xa0", " ").replace("\u200b", "").replace("\xad", "")
    return re.sub(r"[ \t\r\f\v]+", " ", texto).strip()


# ----------------------------------------------------------------------------------------------------------------
# Página do DOU
# ----------------------------------------------------------------------------------------------------------------

@dataclass
class Elemento:
    tipo: str                       # "p", "tabela" ou "imagem"
    texto: str = ""
    classes: set[str] = field(default_factory=set)
    linhas: list[list[str]] = field(default_factory=list)
    mesclas: list[tuple[int, int, int, int]] = field(default_factory=list)
    larguras: list[float] = field(default_factory=list)
    src: str = ""
    largura_px: float = 0.0


@dataclass
class PaginaDOU:
    url: str
    titulo: str
    publicado_em: str = ""          # dd/mm/aaaa
    edicao: str = ""
    secao: str = ""
    pagina: str = ""
    orgao: str = ""
    elementos: list[Elemento] = field(default_factory=list)


def _dado(raiz: No, classe: str) -> str:
    no = raiz.procurar(lambda n: classe in n.classes())
    return limpar(no.texto()).rstrip(" |") if no else ""


def _cabecalho(raiz: No, texto_pagina: str) -> dict[str, str]:
    """Publicado em / Edição / Seção / Página / Órgão: pelas classes da página e, na falta delas, pelo texto."""
    dados = {"publicado_em": _dado(raiz, "publicado-dou-data"), "edicao": _dado(raiz, "edicao-dou-data"),
             "secao": _dado(raiz, "secao-dou"), "pagina": _dado(raiz, "secao-dou-data"),
             "orgao": _dado(raiz, "orgao-dou-data")}
    padroes = {"publicado_em": r"Publicado em:\s*(\d{1,2}/\d{1,2}/\d{4})", "edicao": r"Edi[çc][ãa]o:\s*([\w-]+)",
               "secao": r"Se[çc][ãa]o:\s*([^|\n]+?)\s*(?:\||P[áa]gina|\n|$)",
               "pagina": r"P[áa]gina:\s*(\d+)", "orgao": r"[ÓO]rg[ãa]o:\s*([^\n]+)"}
    for chave, padrao in padroes.items():
        achado = re.search(padrao, texto_pagina, re.I)
        valor = dados[chave]
        valor = re.sub(r"^(?:Se[çc][ãa]o|P[áa]gina|Edi[çc][ãa]o|[ÓO]rg[ãa]o|Publicado em)\s*:\s*", "", valor, flags=re.I)
        if (not valor or (chave == "pagina" and not re.search(r"\d", valor))) and achado:
            valor = achado.group(1)
        dados[chave] = limpar(valor)
    if dados["pagina"] and not dados["pagina"].isdigit():
        achado = re.search(r"\d+", dados["pagina"])
        dados["pagina"] = achado.group(0) if achado else ""
    return dados


def _largura_css(no: No) -> float:
    estilo = no.attrs.get("style", "")
    achado = re.search(r"width\s*:\s*([\d.]+)\s*(%|px|pt|cm)?", estilo, re.I)
    valor = achado.group(1) if achado else no.attrs.get("width", "").rstrip("%px ")
    try:
        return float(valor)
    except ValueError:
        return 0.0


def _tabela(no: No) -> Elemento:
    """<table> -> matriz com as posições cobertas por colspan/rowspan vazias e a lista de mesclas."""
    linhas_html = []
    for tr in (n for n in no.todos() if n.tag == "tr"):
        dono = tr.pai
        while dono is not None and dono.tag != "table":
            dono = dono.pai
        if dono is no:                  # não pega linhas de tabela dentro de célula
            linhas_html.append(tr)
    grade: dict[tuple[int, int], str] = {}
    ocupado: set[tuple[int, int]] = set()
    mesclas, larguras = [], {}
    colunas = 0
    for ri, tr in enumerate(linhas_html):
        ci = 0
        for celula in (f for f in tr.filhos if isinstance(f, No) and f.tag in ("td", "th")):
            while (ri, ci) in ocupado:
                ci += 1
            try:
                junta = max(1, int(celula.attrs.get("colspan", "1") or 1))
                desce = max(1, int(celula.attrs.get("rowspan", "1") or 1))
            except ValueError:
                junta = desce = 1
            texto = "\n".join(l for l in (limpar(x) for x in celula.texto().split("\n")) if l)
            grade[(ri, ci)] = texto
            for r in range(ri, ri + desce):
                for c in range(ci, ci + junta):
                    ocupado.add((r, c))
            if junta > 1 or desce > 1:
                mesclas.append((ri, ci, ri + desce - 1, ci + junta - 1))
            if junta == 1 and ci not in larguras and _largura_css(celula):
                larguras[ci] = _largura_css(celula)
            ci += junta
            colunas = max(colunas, ci)
    total_linhas = max([r for r, _ in ocupado] + [len(linhas_html) - 1]) + 1 if ocupado else 0
    matriz = [[grade.get((r, c), "") for c in range(colunas)] for r in range(total_linhas)]
    # linhas que ficaram só com posições cobertas por rowspan (tabela malformada) saem
    usadas = [r for r in range(total_linhas) if any((r, c) in grade for c in range(colunas))]
    if len(usadas) != total_linhas:
        novo = {r: i for i, r in enumerate(usadas)}
        matriz = [matriz[r] for r in usadas]
        mesclas = [(novo[r0], c0, max(novo[r] for r in usadas if r0 <= r <= r1), c1) for r0, c0, r1, c1 in mesclas]
    # Largura das colunas sem a da página: pelo texto mais longo de cada coluna. No cabeçalho (a primeira linha e as
    # que ela ocupa com rowspan) o título pode quebrar, então conta a palavra mais longa ("Fator de conversão" não
    # deixa a coluna dos números larga).
    fim_cabecalho = max([r1 for r0, _, r1, _ in mesclas if r0 == 0] + [0])
    pesos = []
    for c in range(colunas):
        maior = 4
        for r, row in enumerate(matriz):
            for linha in row[c].split("\n"):
                medida = max((len(p) for p in linha.split()), default=0) if r <= fim_cabecalho else len(linha)
                maior = max(maior, medida)
        pesos.append(larguras.get(c) or max(4.0, min(60.0, float(maior))))
    return Elemento("tabela", linhas=matriz, mesclas=mesclas, larguras=pesos)


def _elementos(conteudo: No) -> list[Elemento]:
    saida: list[Elemento] = []

    def paragrafo(texto: str, classes: set[str]) -> None:
        for linha in texto.split("\n"):           # <br> separa parágrafos (assinaturas, linhas de data)
            linha = limpar(linha)
            if linha and not AVISO_FINAL.match(linha):
                saida.append(Elemento("p", linha, set(classes)))

    def visitar(no: No, classes_herdadas: set[str]) -> None:
        if no.tag in IGNORAR:
            return
        if no.tag == "table":
            tabela = _tabela(no)
            if tabela.linhas and any(any(c for c in linha) for linha in tabela.linhas):
                saida.append(tabela)
            return
        if no.tag == "img":
            src = _endereco_imagem(no)
            if src:
                saida.append(Elemento("imagem", src=src, largura_px=_largura_css(no)))
            return
        classes = no.classes() or classes_herdadas
        if not no.tem_estrutura():
            paragrafo(no.texto(), classes)
            return
        # Há parágrafo, tabela ou figura dentro (mesmo embrulhados num <span> ou <font>): o texto solto entre eles
        # vira parágrafo, e cada um desses elementos é lido no seu lugar.
        texto_solto: list = []
        for filho in no.filhos:
            if isinstance(filho, str) or (filho.tag not in ESTRUTURA and not filho.tem_estrutura()):
                texto_solto.append(filho)
                continue
            if texto_solto:
                paragrafo(_texto_de(texto_solto), classes)
                texto_solto = []
            visitar(filho, classes if filho.tag not in BLOCOS else set())
        if texto_solto:
            paragrafo(_texto_de(texto_solto), classes)

    visitar(conteudo, set())
    return saida


def _endereco_imagem(no: No) -> str:
    """Endereço da figura; páginas com carregamento adiado guardam o verdadeiro em data-src."""
    src = no.attrs.get("src", "").strip()
    for chave in ("data-src", "data-original", "data-lazy-src", "data-full-src"):
        alternativo = no.attrs.get(chave, "").strip()
        if alternativo and (not src or src.startswith("data:image/gif") or "placeholder" in src.lower()):
            return alternativo
    return src


def _area_do_ato(raiz: No) -> "No | None":
    """A área com o texto do ato. Se houver mais de uma candidata (cópia escondida, versão de impressão), fica a
    que tem o título do ato (classe "identifica"), mais tabelas e figuras e mais texto."""
    candidatas = [n for n in raiz.todos() if "texto-dou" in n.classes()]
    if not candidatas:
        candidatas = [n for n in raiz.todos() if any("texto-dou" in c for c in n.classes())]
    if not candidatas:
        candidatas = [n for n in raiz.todos()
                      if "single-full-dou-content" in n.classes() or n.attrs.get("id") == "materia"]
    if not candidatas:
        return None

    def nota(no: No) -> tuple:
        internos = list(no.todos())
        return (any("identifica" in n.classes() for n in internos), sum(n.tag == "table" for n in internos),
                sum(n.tag == "img" for n in internos), len(no.texto()))
    return max(candidatas, key=nota)


def ler_pagina(html: str, url: str = "") -> PaginaDOU:
    raiz = arvore(html)
    conteudo = _area_do_ato(raiz)
    texto_pagina = raiz.texto()
    if conteudo is None:
        if re.search(r"captcha|access denied|forbidden|bloquead|unusual traffic", texto_pagina, re.I):
            raise ConnectionError("o site do DOU recusou o acesso do programa")
        raise ValueError("a página não tem o texto de um ato do DOU (área “texto-dou” não encontrada). "
                         "Confira o link: deve ser a página do ato, como "
                         "https://www.in.gov.br/web/dou/-/medida-provisoria-n-1.399-de-4-de-novembro-de-2026-741754451")
    elementos = _elementos(conteudo)
    textos = [e for e in elementos if e.tipo == "p"]
    if not textos:
        raise ValueError("a página do DOU não trouxe o texto do ato")
    titulo = next((e.texto for e in textos if "identifica" in e.classes), textos[0].texto)
    return PaginaDOU(url, titulo, elementos=elementos, **_cabecalho(raiz, texto_pagina))


# ----------------------------------------------------------------------------------------------------------------
# Rede
# ----------------------------------------------------------------------------------------------------------------

def _get(url: str, timeout: int = 40) -> tuple[int, bytes, str]:
    pedido = urllib.request.Request(url, headers={
        "User-Agent": UA, "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/*,*/*;q=0.8",
        "Accept-Language": "pt-BR,pt;q=0.9", "Accept-Encoding": "gzip, deflate"})
    try:
        with urllib.request.urlopen(pedido, timeout=timeout) as resposta:
            corpo = resposta.read()
            codificacao = (resposta.headers.get("Content-Encoding") or "").lower()
            if codificacao == "gzip":
                corpo = gzip.decompress(corpo)
            elif codificacao == "deflate":
                corpo = zlib.decompress(corpo)
            return resposta.status, corpo, resposta.headers.get_content_charset() or ""
    except urllib.error.HTTPError as erro:
        return erro.code, b"", ""


def _decodificar(corpo: bytes, charset: str) -> str:
    if not charset:
        achado = re.search(rb"charset=[\"']?([\w-]+)", corpo[:4000], re.I)
        charset = achado.group(1).decode() if achado else "utf-8"
    try:
        return corpo.decode(charset, errors="replace")
    except LookupError:
        return corpo.decode("utf-8", errors="replace")


def _pelo_navegador(url: str, log: Callable[[str], None]) -> str:
    """Abre a página no Edge/Chrome (Playwright, já usado nos hyperlinks), como uma pessoa."""
    from playwright.sync_api import sync_playwright
    with sync_playwright() as pw:
        navegador, erro = None, None
        for canal in ("msedge", "chrome", None):
            try:
                navegador = pw.chromium.launch(channel=canal, headless=True)
                break
            except Exception as e:      # noqa: BLE001
                erro = e
        if navegador is None:
            raise ConnectionError(f"não consegui abrir o Edge/Chrome ({str(erro)[:120]})")
        try:
            pagina = navegador.new_context(user_agent=UA, locale="pt-BR").new_page()
            log("Abrindo a página do DOU no navegador...")
            pagina.goto(url, wait_until="domcontentloaded", timeout=60000)
            try:
                pagina.wait_for_selector(".texto-dou", timeout=20000)
            except Exception:           # noqa: BLE001
                pass
            return pagina.content()
        finally:
            navegador.close()


# Figuras de uma página salva como "Página da Web, arquivo único (*.mhtml)" (Edge/Chrome), pelo endereço de
# cada uma. Com elas, a figura do ato vem da própria página salva, sem nova busca no DOU (3.28.1; é o que a
# versão web usa, porque o navegador não deixa o site baixar nada do in.gov.br).
FIGURAS_SALVAS: dict[str, bytes] = {}


def ler_mhtml(bruto: bytes) -> tuple[str, str, dict[str, bytes]]:
    """(html, endereço da página, {endereço: figura}) de uma página salva como arquivo único (.mhtml)."""
    from email import policy
    from email.parser import BytesParser
    mensagem = BytesParser(policy=policy.compat32).parsebytes(bruto)
    html, url, figuras = "", "", {}
    for parte in mensagem.walk():
        if parte.is_multipart():
            continue
        tipo = parte.get_content_type()
        local = (parte.get("Content-Location") or "").strip()
        dados = parte.get_payload(decode=True) or b""
        if tipo == "text/html" and not html:
            html, url = _decodificar(dados, parte.get_content_charset() or ""), local
        elif tipo.startswith("image/") and local:
            figuras[local] = dados
    if not html:
        raise ValueError("o arquivo .mhtml não tem a página do ato")
    return html, url, figuras


def figura_salva(src: str, url_pagina: str) -> bytes | None:
    """A figura guardada na página salva (.mhtml), procurada pelo endereço completo e pelo que está na página."""
    endereco = urllib.parse.urljoin(url_pagina or "https://www.in.gov.br/", src)
    for chave in (endereco, src, endereco.replace("http://", "https://")):
        if chave in FIGURAS_SALVAS:
            return FIGURAS_SALVAS[chave]
    return None


def baixar_pagina(endereco: str, log: Callable[[str], None] = print) -> tuple[str, str]:
    """Devolve (html, url). Aceita o link do DOU ou o caminho de uma página salva (.html ou .mhtml)."""
    endereco = endereco.strip().strip('"')
    arquivo = Path(endereco)
    FIGURAS_SALVAS.clear()
    if not re.match(r"^https?://", endereco, re.I):
        if arquivo.suffix.lower() in (".html", ".htm", ".mhtml", ".mht") and arquivo.is_file():
            bruto = arquivo.read_bytes()
            padrao = r"https?://(?:www\.)?in\.gov\.br/(?:en/)?web/dou/-/[^\s\"'<>]+"
            if arquivo.suffix.lower() in (".mhtml", ".mht") or bruto.lstrip()[:300].lower().startswith(
                    (b"from:", b"mime-version", b"snapshot-content-location")):
                html, url, figuras = ler_mhtml(bruto)
                FIGURAS_SALVAS.update(figuras)
                if not eh_link_dou(url):
                    achado = re.search(padrao, html)
                    url = achado.group(0) if achado else ""
                return html, url
            achado = re.search(padrao.encode(), bruto)
            url = achado.group(0).decode("ascii", "ignore") if achado else ""
            return _decodificar(bruto, ""), url
        raise ValueError("Cole o link da página do ato no DOU (https://www.in.gov.br/web/dou/-/...) "
                         "ou escolha a página salva (.html ou .mhtml).")
    if not eh_link_dou(endereco):
        raise ValueError("O link não é de uma página de ato do DOU. Ele deve começar por "
                         "https://www.in.gov.br/web/dou/-/ (abra o ato no DOU e copie o endereço da barra).")
    log("Buscando a página no DOU...")
    falhas = []
    try:
        status, corpo, charset = _get(endereco)
        if status == 200 and corpo:
            html = _decodificar(corpo, charset)
            if "texto-dou" in html:
                return html, endereco
            falhas.append("a página veio sem o texto do ato")
        else:
            falhas.append(f"o DOU respondeu {status}")
    except Exception as erro:          # noqa: BLE001
        falhas.append(f"sem conexão direta ({str(erro)[:100]})")
    try:
        return _pelo_navegador(endereco, log), endereco
    except Exception as erro:          # noqa: BLE001
        falhas.append(f"pelo navegador: {str(erro)[:120]}")
    raise ConnectionError("Não consegui abrir a página do DOU (" + "; ".join(falhas) + "). Abra o link no navegador, "
                          "salve a página (Ctrl+S, “Página da Web, arquivo único”, que traz as figuras, ou “somente HTML”) e "
                          "importe o arquivo salvo.")


def baixar_imagem(src: str, url_pagina: str) -> bytes:
    if src.startswith("data:"):
        import base64
        cabeca, _, dados = src.partition(",")
        return base64.b64decode(dados) if ";base64" in cabeca else urllib.parse.unquote_to_bytes(dados)
    salva = figura_salva(src, url_pagina)
    if salva:
        return salva
    endereco = urllib.parse.urljoin(url_pagina or "https://www.in.gov.br/", src)
    try:
        status, corpo, _ = _get(endereco, timeout=40)
    except Exception as erro:          # noqa: BLE001
        status, corpo = str(erro)[:60], b""
    if status == 200 and corpo:
        return corpo
    try:                               # o DOU recusou o programa: a figura vem pelo Edge/Chrome, como a página
        return _imagem_pelo_navegador(endereco)
    except Exception as erro:          # noqa: BLE001
        raise ConnectionError(f"status {status}; pelo navegador: {str(erro)[:80]}") from erro


def _imagem_pelo_navegador(endereco: str) -> bytes:
    from playwright.sync_api import sync_playwright
    with sync_playwright() as pw:
        navegador = None
        for canal in ("msedge", "chrome", None):
            try:
                navegador = pw.chromium.launch(channel=canal, headless=True)
                break
            except Exception:           # noqa: BLE001
                continue
        if navegador is None:
            raise ConnectionError("não consegui abrir o Edge/Chrome")
        try:
            pagina = navegador.new_context(user_agent=UA, locale="pt-BR").new_page()
            resposta = pagina.goto(endereco, wait_until="load", timeout=60000)
            if resposta is None or not resposta.ok:
                raise ConnectionError(f"status {resposta.status if resposta else '?'}")
            return resposta.body()
        finally:
            navegador.close()


# ----------------------------------------------------------------------------------------------------------------
# Página -> mesma estrutura da importação do PDF
# ----------------------------------------------------------------------------------------------------------------

def nota_publicacao(pagina: PaginaDOU) -> str | None:
    achado = re.match(r"(\d{1,2})/(\d{1,2})/(\d{4})$", pagina.publicado_em)
    if not achado:
        return None
    dia, mes, ano = achado.groups()
    feminino = normalized(pagina.titulo).startswith(("PORTARIA", "RESOLUCAO", "INSTRUCAO", "LEI", "SOLUCAO",
                                                       "MEDIDA", "EMENDA"))
    extra = bool(re.search(r"\d+\s*-\s*[A-Z]\b", pagina.edicao) or re.search(r"extra", pagina.secao, re.I))
    data = f"{int(dia):02d}.{int(mes):02d}.{ano[-2:]}"
    return (f"{'Publicada' if feminino else 'Publicado'} no DOU de {data}" + (". Ed. Extra." if extra else "."))


def _png(dados: bytes) -> tuple[bytes, float]:
    """A figura em PNG (o Word não aceita alguns formatos da web) e a largura em pontos (96 pontos por polegada)."""
    from PIL import Image
    with Image.open(io.BytesIO(dados)) as imagem:
        largura = imagem.width * 72 / 96
        saida = io.BytesIO()
        imagem.convert("RGBA" if imagem.mode in ("RGBA", "LA", "P") else "RGB").save(saida, format="PNG")
        return saida.getvalue(), largura


def extrair_pagina(pagina: PaginaDOU, obter_imagem: Callable[[str, str], bytes] = baixar_imagem) -> Extraction:
    numero = int(pagina.pagina) if pagina.pagina.isdigit() else 0
    extraction = Extraction(Path(""), pagina.titulo, [], [numero] if numero else [])
    extraction.source_url = pagina.url
    extraction.protocol = "página do ato no DOU"
    textos = [e for e in pagina.elementos if e.tipo == "p"]
    publicacao = nota_publicacao(pagina)
    sequencia = [e.texto for e in textos]
    epigrafe = next((i for i, e in enumerate(textos) if "identifica" in e.classes), 0)
    if publicacao:
        sequencia.insert(epigrafe + 1, publicacao)
    else:
        extraction.issues.append("A página do DOU não trouxe a data de publicação; acrescente a linha "
                                 "“Publicado no DOU de ...” no Word.")
    itens = classify("\n".join(sequencia))
    if len(itens) != len(sequencia):        # o classificador devolve um item por linha; se não, por segurança:
        itens = [Item(t, 0, 0, "texto", "Importado do DOU") for t in sequencia]
    item_publicacao = None
    if publicacao:
        item_publicacao = itens.pop(epigrafe + 1)
        item_publicacao.role = "publicacao"
        item_publicacao.attention.append("Linha criada a partir de “Publicado em” e da edição informados na página "
                                         "do DOU; confirme.")
    blocos: list = []
    proximo = iter(itens)
    indice_texto = -1
    anterior_tabela = legenda = False
    for elemento in pagina.elementos:
        if elemento.tipo == "p":
            item = next(proximo)
            indice_texto += 1
            classes = elemento.classes
            if "ementa" in classes and item.role in ("texto", "titulo"):
                item.role, item.reason = "ementa", "Ementa na página do DOU"
            elif "assina" in classes and item.role in ("texto", "titulo"):
                item.role, item.reason = "autoridade", "Assinatura na página do DOU"
            elif "cargo" in classes and item.role in ("texto", "titulo", "autoridade"):
                item.role, item.reason = "cargo", "Cargo na página do DOU"
            tipo = "norma"
            if anterior_tabela or legenda:
                # "Siglas:" / "Secretaria Especial ... - RFB" / "OBSERVAÇÕES:" depois da tabela: legenda
                legenda = bool(LEGEND_HEAD.match(item.text) or TABLE_NOTE.match(item.text)
                               or (legenda and (LEGEND_ITEM.search(item.text)
                                                or re.match(r"^\d{1,2}\s*[).]", item.text))))
                if legenda and item.role == "texto":
                    tipo = "footnote"
            anterior_tabela = False
            blocos.append(ParagraphBlock(item.text, item.role, numero, tipo, item, list(item.attention)))
            if item_publicacao is not None and indice_texto == epigrafe:
                blocos.append(ParagraphBlock(item_publicacao.text, "publicacao", numero, "norma", item_publicacao,
                                             list(item_publicacao.attention)))
        elif elemento.tipo == "tabela":
            matriz = elemento.linhas
            largura = max(len(linha) for linha in matriz)
            matriz = [linha + [""] * (largura - len(linha)) for linha in matriz]
            blocos.append(TableBlock(
                matriz, elemento.larguras or [1.0] * largura, numero, (0.0, 0.0, 0.0, 0.0),
                monetary_values(matriz), matrix_signature(matriz), deepcopy(matriz), "página do DOU",
                elemento.mesclas, [], False, "table",
                [TablePagePart(numero, (0.0, 0.0, 0.0, 0.0), 0, len(matriz), "")]))
            anterior_tabela, legenda = True, False
        else:
            try:
                dados, largura = _png(obter_imagem(elemento.src, pagina.url))
            except Exception as erro:          # noqa: BLE001
                if not blocos:
                    continue                    # brasão ou logotipo antes do título
                # o Word sai com um aviso no lugar da figura, para colá-la à mão a partir do DOU
                aviso = "[FIGURA DO ATO: copie-a da página do DOU e cole aqui]"
                marca = Item(aviso, 0, 0, "texto", "Figura que não foi possível copiar da página do DOU")
                marca.attention.append("Cole aqui a figura do ato, copiada da página do DOU.")
                blocos.append(ParagraphBlock(aviso, "texto", numero, "norma", marca, list(marca.attention)))
                extraction.issues.append(
                    f"Não consegui copiar uma figura do ato ({elemento.src[:80]}: {str(erro)[:60]}). No Word, "
                    "o lugar dela está marcado com “[FIGURA DO ATO: ...]”: copie a figura da página do DOU "
                    "(ou do PDF) e cole ali.")
                continue
            if largura < 40 and not blocos:
                continue                        # brasão pequeno antes do título
            blocos.append(ImageBlock(dados, numero, "content", min(largura, 450.0)))
            extraction.issues.append("A página do DOU traz uma figura no texto do ato; ela foi incluída como imagem "
                                     "no Word, no mesmo lugar. Se quiser o texto, digite-o a partir da imagem.")
    _subtitulos_de_anexo(blocos)
    extraction.blocks = blocos
    add_cgibs_availability(extraction)
    date_constitutional_amendment(extraction)
    return extraction


def _subtitulos_de_anexo(blocos: list) -> None:
    """ANEXO I / "ALÍQUOTAS-BASE PARA VEÍCULOS": o nome do anexo (em maiúsculas, logo depois) é subtítulo."""
    for indice, bloco in enumerate(blocos[:-1]):
        seguinte = blocos[indice + 1]
        if (isinstance(bloco, ParagraphBlock) and bloco.role == "anexo" and isinstance(seguinte, ParagraphBlock)
                and seguinte.role == "texto" and len(seguinte.text) <= 300
                and (looks_uppercase(seguinte.text) or _mostly_uppercase(seguinte.text))
                and not any(rx.match(seguinte.text) for rx in (ANNEX, ARTICLE, HEADING))):
            seguinte.role = "subtitulo"
            if seguinte.item:
                seguinte.item.role = "subtitulo"


def extrair_do_link(endereco: str, log: Callable[[str], None] = print) -> Extraction:
    html, url = baixar_pagina(endereco, log)
    pagina = ler_pagina(html, url)
    log(f"Lido: {pagina.titulo}")
    return extrair_pagina(pagina)


def nome_sugerido(extraction: Extraction) -> str:
    return re.sub(r"[^\w.-]+", "_", extraction.query).strip("_")[:120] or "ato_DOU"


def descrever(pagina: PaginaDOU) -> str:
    return (f"{pagina.titulo} · Publicado em {pagina.publicado_em or '?'} · Edição {pagina.edicao or '?'} · "
            f"Seção {pagina.secao or '?'} · Página {pagina.pagina or '?'}")


__all__ = ["FIGURAS_SALVAS", "PaginaDOU", "baixar_pagina", "eh_link_dou", "extrair_do_link", "extrair_pagina",
           "figura_salva", "ler_mhtml", "ler_pagina", "nome_sugerido", "nota_publicacao"]
