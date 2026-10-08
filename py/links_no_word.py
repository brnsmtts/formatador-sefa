#!/usr/bin/env python3
"""Etapa de hyperlinks de legislação do Formatador SEFA.

Depois que o Word é gerado e conferido, esta etapa abre o arquivo e cria os links
das leis, decretos, INs, portarias etc. citados, usando links_legislacao.py
(tabela local -> site de busca de legislação da SEFA/PA -> Planalto).

O arquivo só é substituído se o texto, os estilos e os negritos de todos os
parágrafos continuarem idênticos depois dos links. Se algo falhar, o Word sem
links que já estava salvo permanece intacto e o motivo vai para o relatório.
"""

from __future__ import annotations

import html
import os
import queue
import tempfile
import threading
import traceback
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

from docx import Document
from docx.oxml.ns import qn

_W_P, _W_R, _W_RPR, _W_B = qn("w:p"), qn("w:r"), qn("w:rPr"), qn("w:b")
_W_PPR, _W_PSTYLE, _W_VAL = qn("w:pPr"), qn("w:pStyle"), qn("w:val")
_CONTEUDO = {qn("w:t"): None, qn("w:tab"): "\t", qn("w:br"): "\n", qn("w:cr"): "\n",
             qn("w:drawing"): "\ufffc", qn("w:pict"): "\ufffc", qn("w:object"): "\ufffc"}


@dataclass
class ResumoLinks:
    concluido: bool            # True quando o Word final passou pela etapa de links
    linkados: int = 0          # citações que receberam link
    pendentes: int = 0         # citações sem link
    destacadas: int = 0        # citações sem link destacadas em amarelo no Word
    novos: int = 0             # atos novos guardados em tabela_links.tsv
    atos: list[tuple[str, int, str, str]] = field(default_factory=list)  # (ato, ocorrências, situação, link)
    aviso: str = ""
    enderecos: int = 0         # endereços de internet que receberam link

    def frase(self) -> str:
        if not self.concluido:
            return self.aviso or "O Word foi mantido sem links."
        enderecos = f" Endereços de internet com link: {self.enderecos}." if self.enderecos else ""
        if not self.atos:
            return "Hyperlinks: nenhuma citação de lei, decreto, instrução normativa etc. foi encontrada." + enderecos
        texto = f"Hyperlinks: {self.linkados} citação(ões) com link, {self.pendentes} sem link"
        if self.destacadas:
            texto += f" ({self.destacadas} destacada(s) em amarelo no Word)"
        return texto + "." + enderecos

    def pendencias(self, limite: int = 6) -> str:
        """Atos sem link e o motivo, para a mensagem final quando não há relatório."""
        partes = []
        sem_link = [(ato, situacao) for ato, _, situacao, url in self.atos if not url]
        conferir = [(ato, situacao) for ato, _, situacao, url in self.atos if url and "conferir" in situacao.lower()]
        for titulo, lista in (("Sem link", sem_link), ("Links a conferir (abra e confira o ato)", conferir)):
            if lista:
                linhas = [f"- {ato}: {situacao}" for ato, situacao in lista[:limite]]
                if len(lista) > limite:
                    linhas.append(f"- e mais {len(lista) - limite}")
                partes.append(f"{titulo}:\n" + "\n".join(linhas))
        return "\n\n".join(partes)


def assinatura(doc) -> list[tuple]:
    """Texto, estilo e negrito direto de cada parágrafo (inclusive em tabelas), mais a
    estrutura das tabelas. Os links não podem alterar nada disso."""
    corpo = doc.element.body
    paragrafos = []
    for p in corpo.iter(_W_P):
        ppr = p.find(_W_PPR)
        ps = ppr.find(_W_PSTYLE) if ppr is not None else None
        texto, negrito = [], []
        for r in p.iter(_W_R):
            rpr = r.find(_W_RPR)
            b = rpr.find(_W_B) if rpr is not None else None
            ligado = b is not None and b.get(_W_VAL, "1") not in ("0", "false", "off")
            for filho in r:
                if filho.tag not in _CONTEUDO:
                    continue
                s = _CONTEUDO[filho.tag]
                s = (filho.text or "") if s is None else s
                texto.append(s)
                negrito.append(("1" if ligado else "0") * len(s))
        paragrafos.append((ps.get(_W_VAL) if ps is not None else "", "".join(texto), "".join(negrito)))
    estrutura = (sum(1 for _ in corpo.iter(qn("w:tbl"))), sum(1 for _ in corpo.iter(qn("w:tr"))),
                 sum(1 for _ in corpo.iter(qn("w:tc"))))
    return [estrutura] + paragrafos


def aplicar_links(docx_path: Path, log: Callable[[str], None] = print,
                  parar: threading.Event | None = None) -> ResumoLinks:
    """Cria os links no próprio arquivo (substitui-o só depois da conferência)."""
    docx_path = Path(docx_path)
    try:
        import links_legislacao as ll
    except Exception as erro:  # componente ausente ou arquivo removido da pasta
        return ResumoLinks(False, aviso=f"O módulo de links não pôde ser carregado ({erro}). "
                                        "O Word foi salvo sem links; rode o INSTALAR.bat de novo.")
    log("Procurando citações de atos normativos no Word gerado...")
    doc = Document(str(docx_path))
    antes = assinatura(doc)
    res = ll.criar_links(doc, log=log, parar=parar)
    resumo = ResumoLinks(True, res.linkados, res.pendentes, res.destacadas, res.novos, res.linhas,
                         enderecos=res.enderecos)
    if not (res.linkados or res.destacadas or res.enderecos):
        return resumo          # nada visível mudou; o arquivo salvo já é o final
    with tempfile.NamedTemporaryFile(prefix="sefa_links_", suffix=".docx", dir=docx_path.parent,
                                     delete=False) as tmp:
        temporario = Path(tmp.name)
    try:
        doc.save(str(temporario))
        if assinatura(Document(str(temporario))) != antes:
            resumo.concluido = False
            resumo.aviso = ("A conferência feita depois de criar os links encontrou diferença no texto, "
                            "no estilo ou no negrito; por segurança, o Word foi mantido sem links.")
            log("! " + resumo.aviso)
            return resumo
        os.replace(temporario, docx_path)
    finally:
        temporario.unlink(missing_ok=True)
    return resumo


_CSS = ("<style>#links-legislacao{margin-top:34px}"
        "#links-legislacao table{table-layout:auto!important;width:100%;border-collapse:collapse;font-size:13px}"
        "#links-legislacao th,#links-legislacao td{width:auto!important;text-align:left!important;"
        "border:1px solid #bbb;padding:6px;vertical-align:top;overflow-wrap:anywhere}"
        "#links-legislacao th{background:#eee}#links-legislacao .pendente{background:#fff4dc}"
        "#links-legislacao .aviso{background:#fde2e2;padding:8px}</style>")


def secao_html(resumo: ResumoLinks) -> str:
    partes = [_CSS, '<section id="links-legislacao"><h2>Hyperlinks de legislação</h2>']
    if not resumo.concluido:
        partes.append(f'<p class="aviso"><strong>{html.escape(resumo.aviso)}</strong></p>')
    elif not resumo.atos:
        partes.append("<p>Nenhuma citação de lei, decreto, instrução normativa, portaria etc. foi encontrada.</p>")
    else:
        partes.append(f"<p>{resumo.linkados} citação(ões) com link e {resumo.pendentes} sem link"
                      + (f"; {resumo.destacadas} destacada(s) em amarelo no Word (retire o destaque "
                         "depois de resolver)" if resumo.destacadas else "")
                      + (f". Atos novos guardados na tabela local: {resumo.novos}" if resumo.novos else "")
                      + ".</p>")
    if resumo.atos:
        linhas = []
        for exemplo, qtd, situacao, url in resumo.atos:
            conferir = not url or "conferir" in situacao.lower()
            link = f'<a href="{html.escape(url)}">{html.escape(url)}</a>' if url else "—"
            linhas.append(f'<tr{" class=pendente" if conferir else ""}><td>{html.escape(exemplo)}</td>'
                          f"<td>{qtd}</td><td>{html.escape(situacao)}</td><td>{link}</td></tr>")
        partes.append("<table><tr><th>Ato (1ª citação)</th><th>Ocorrências</th><th>Situação</th>"
                      "<th>Link</th></tr>" + "".join(linhas) + "</table>")
    partes.append("<p><small>Fontes, nesta ordem: <code>tabela_links.tsv</code> (pasta do programa; tem "
                  "prioridade e pode ser editada à mão), site de busca de legislação da SEFA/PA e, para atos "
                  "federais ausentes nele, o Planalto. A epígrafe e a ementa do próprio ato não recebem links. "
                  "Abra alguns links para conferir o ato de destino.</small></p></section>")
    return "".join(partes)


def anexar_ao_relatorio(report_path: Path, resumo: ResumoLinks) -> None:
    report_path = Path(report_path)
    texto = report_path.read_text(encoding="utf-8")
    fim = texto.rfind("</html>")
    secao = secao_html(resumo)
    texto = texto[:fim] + secao + texto[fim:] if fim >= 0 else texto + secao
    report_path.write_text(texto, encoding="utf-8")


def aplicar_e_registrar(docx_path: Path, report_path: Path | None, log: Callable[[str], None] = print,
                        parar: threading.Event | None = None) -> ResumoLinks:
    """Etapa completa usada pelas janelas e pela linha de comando. Não levanta exceção:
    qualquer falha deixa o Word sem links e fica registrada no resumo (e no relatório,
    quando ele foi gerado; report_path=None significa que o usuário não pediu relatório)."""
    try:
        resumo = aplicar_links(docx_path, log, parar)
    except Exception as erro:
        log(traceback.format_exc())
        resumo = ResumoLinks(False, aviso=f"Erro ao criar os links ({erro}). O Word foi mantido sem links.")
    if report_path is None:
        return resumo
    try:
        anexar_ao_relatorio(report_path, resumo)
    except Exception as erro:
        log(f"! Não foi possível acrescentar os links ao relatório: {erro}")
    return resumo


def executar_com_progresso(root, titulo: str,
                           tarefa: Callable[[Callable[[str], None], threading.Event], object],
                           ao_concluir: Callable[[object, BaseException | None], None]) -> None:
    """Roda tarefa(log, parar) em segundo plano, com uma janela de andamento; depois chama
    ao_concluir(resultado, erro) na linha principal do Tkinter. A busca no site pode levar
    alguns minutos, e assim a janela do programa não congela."""
    import tkinter as tk
    from tkinter import scrolledtext, ttk

    janela = tk.Toplevel(root)
    janela.title(titulo)
    janela.geometry("760x400")
    janela.transient(root)
    janela.protocol("WM_DELETE_WINDOW", lambda: None)   # não fecha no meio da gravação
    rodape = ttk.Frame(janela, padding=(10, 0, 10, 10))
    rodape.pack(side="bottom", fill="x")          # antes da caixa, para nunca ser cortado
    caixa = scrolledtext.ScrolledText(janela, font=("Consolas", 9), state="disabled", wrap="word", height=12)
    caixa.pack(fill="both", expand=True, padx=10, pady=(10, 6))
    ttk.Label(rodape, text="O Word já está salvo; esta etapa acrescenta os links das normas citadas.").pack(side="left")
    parar = threading.Event()

    def interromper() -> None:
        parar.set()
        botao.configure(state="disabled", text="Interrompendo...")

    botao = ttk.Button(rodape, text="Parar buscas no site", command=interromper)
    botao.pack(side="right")
    try:                       # modal: evita um segundo "Salvar" durante a busca
        janela.update_idletasks()
        janela.grab_set()
    except tk.TclError:
        pass
    fila: queue.Queue[str] = queue.Queue()
    estado: dict = {}

    def trabalho() -> None:
        try:
            estado["resultado"] = tarefa(fila.put, parar)
        except BaseException as erro:  # noqa: BLE001 - devolvido à janela principal
            estado["erro"] = erro
            fila.put(traceback.format_exc())
        finally:
            estado["fim"] = True

    def bombear() -> None:
        while not fila.empty():
            caixa.configure(state="normal")
            caixa.insert("end", fila.get() + "\n")
            caixa.see("end")
            caixa.configure(state="disabled")
        if estado.get("fim") and fila.empty():
            janela.grab_release()
            janela.destroy()
            ao_concluir(estado.get("resultado"), estado.get("erro"))
        else:
            janela.after(120, bombear)

    threading.Thread(target=trabalho, daemon=True).start()
    bombear()


def texto_final(resumo: object, erro: BaseException | None, com_relatorio: bool = True) -> str:
    """Resumo para a mensagem final. Sem relatório, inclui os atos que ficaram sem link."""
    if erro is not None:
        return f"Erro na etapa de links ({erro}). O Word foi mantido sem links."
    if not isinstance(resumo, ResumoLinks):
        return ""
    texto = resumo.frase()
    if not com_relatorio and resumo.concluido and resumo.pendencias():
        texto += "\n\n" + resumo.pendencias()
    return texto
