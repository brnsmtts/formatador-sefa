"""Ponte entre a página (aba Aplicar alterações) e o Consolidador SEFA, rodando no navegador.

Os módulos da pasta consolidador/ vêm do programa do computador sem alteração. Aqui ficam só as
partes da janela do programa (gui.py, que usa o tkinter) refeitas para a página: a seleção das
propostas, a validação de correções, a mudança manual, a aprovação da extração de PDF/HTML, a
sincronização dos campos antes de gerar e a abertura de um plano salvo.

Cada função devolve um JSON: {"ok": true, ...} ou {"ok": false, "erro": "..."}."""

from __future__ import annotations

import json
import traceback
from pathlib import Path, PureWindowsPath

from consolidador import __version__
from consolidador.engine import analyze, generate, save_plan, validate_plan
from consolidador.index import Legislation
from consolidador.io import prepare_base
from consolidador.model import Operation, Plan, Target, date_iso, digest
from consolidador.observations import analysis_observations, analysis_summary
from consolidador.report import report

PASTA = Path("/work/consolidar")
CACHE = PASTA / "cache"

# Os mesmos rótulos da janela do programa (consolidador/gui.py)
ACOES = {"Substituir redação": "replace", "Acrescentar dispositivo": "add", "Revogar dispositivo": "revoke",
         "Renumerar parágrafo único": "rename"}
ROTULO_ACAO = {v: k for k, v in ACOES.items()}
ALCANCES = {"Somente este texto": "head", "Dispositivo completo": "subtree"}
ANO_NOTAS = ("auto", "2", "4")
PDF_REVIEW = "Confira a extração de PDF/HTML"
CURTO = {"replace": "Substituir", "add": "Acrescentar", "revoke": "Revogar", "rename": "Renumerar"}

ESTADO: dict = {"plan": None, "base": None}


def _json(**dados) -> str:
    return json.dumps({"ok": True, **dados}, ensure_ascii=False)


def _erro(exc: Exception) -> str:
    print(traceback.format_exc())
    return json.dumps({"ok": False, "erro": str(exc) or exc.__class__.__name__}, ensure_ascii=False)


def br_date(value: str) -> str:
    """Datas ISO como dd/mm/aaaa nos campos (como na janela do programa)."""
    parts = value.split("-") if value else []
    return "/".join(reversed(parts)) if len(parts) == 3 else value


def _nome(caminho: str) -> str:
    return PureWindowsPath(caminho or "").name     # vale para caminhos do Windows e daqui


def _sem_doc(caminho: str, papel: str) -> None:
    if str(caminho).lower().endswith(".doc"):
        raise ValueError(f"{papel} está em .doc, formato antigo do Word que o navegador não consegue converter. "
                         "Abra o arquivo no Word, use Salvar como > Documento do Word (.docx) e escolha a cópia. "
                         "O original fica como está.")


def _plano():
    if ESTADO["plan"] is None:
        raise ValueError("Faça a análise primeiro.")
    return ESTADO["plan"], ESTADO["base"]


def _proposta(op_id: str) -> Operation:
    plan, _ = _plano()
    for op in plan.operations:
        if op.id == op_id:
            return op
    raise ValueError("A proposta não existe mais. Selecione outra na lista.")


def _texto_antes(op: Operation) -> str:
    before = "\n\n".join(op.before)
    if before.strip():
        return before
    if op.action == "add":
        return "Este é um acréscimo. Não há redação anterior deste dispositivo na base. Confira o texto na aba Nova redação."
    return ("Não foi localizado um texto atual para este destino. Confira o destino e a situação da proposta em "
            "Corrigir ação, destino e datas.")


def _notas(op: Operation) -> str:
    notes = ["EFEITOS DESTA MUDANÇA: " + (op.effects or "data não identificada")]
    if op.effects_source:
        notes.extend(("CLÁUSULA EXTRAÍDA DO ATO", op.effects_source))
    if op.remissions:
        notes.extend(("REMISSÕES QUE ENTRARÃO NO WORD", "\n\n".join(op.remissions)))
    if op.related_source:
        notes.extend(("DISPOSIÇÕES RELACIONADAS NA ORIGEM", "\n\n".join(op.related_source)))
    if op.note:
        notes.append(op.note)
    return "\n\n".join(notes)


def _dica(op: Operation) -> str:
    if op.action == "revoke":
        return "Revogação: não exige nova redação digitada. O programa insere a marca REVOGADO e mantém o histórico."
    if op.action == "rename":
        return "Renumeração: confira o novo número em Corrigir ação, destino e datas. Não exige nova redação digitada."
    if op.new_lines:
        return "Confira as três abas. Se editar a nova redação, clique em Validar correção antes de gerar o Word."
    return "A nova redação está vazia. Informe o texto com o rótulo do dispositivo e valide a correção."


def _op_json(op: Operation) -> dict:
    return {
        "id": op.id, "situacao": op.status, "acao": CURTO.get(op.action, op.action),
        "acao_rotulo": ROTULO_ACAO.get(op.action, "Substituir redação"),
        "destino": op.display_label(), "corpo": op.target.corpus,
        "alcance": "Dispositivo completo" if op.scope == "subtree" else "Somente este texto",
        "efeitos": "/".join(reversed(op.effects.split("-"))) if op.effects else "não informados",
        "efeitos_campo": br_date(op.effects), "renumerar": op.rename_to,
        "origem": op.source_location, "selecionada": bool(op.selected), "aplicada": bool(op.applied),
        "motivo": op.reason, "antes": _texto_antes(op), "nova": "\n".join(op.new_lines),
        "trecho": op.source.strip() or ("Nenhum trecho de origem foi extraído para esta proposta. "
                                        "Confira o comando diretamente no ato alterador."),
        "notas": _notas(op), "dica": _dica(op),
        "pdf_pendente": op.status == "revisar" and op.extraction_warning.startswith(PDF_REVIEW),
    }


def _estado(incluir_base: bool = False, **extra) -> str:
    plan, base = _plano()
    ops = plan.operations
    prontas = sum(o.status == "pronto" for o in ops)
    selecionadas = sum(bool(o.selected) for o in ops)
    pendentes = sum(o.status == "revisar" for o in ops)
    dados = {
        "versao": __version__,
        "operacoes": [_op_json(o) for o in ops],
        "resumo": (f"{len(ops)} propostas · {prontas} prontas · {selecionadas} selecionadas · "
                   f"{pendentes} pendentes de revisão · detalhes na aba Observações"),
        "observacoes": analysis_observations(plan),
        "prontas": prontas, "selecionadas": selecionadas, "pendentes": pendentes,
        "aplicadas": sum(bool(o.applied) for o in ops),
        "pendentes_pdf": sum(o.status == "revisar" and o.extraction_warning.startswith(PDF_REVIEW) for o in ops),
        "ato": plan.citation, "publicacao": br_date(plan.publication), "efeitos": br_date(plan.effects),
        "ano_notas": plan.date_style if plan.date_style in ANO_NOTAS else "auto",
        "base_nome": _nome(plan.base_path), "alterador_nome": _nome(plan.amender_path),
        "nome_saida": Path(_nome(plan.base_path)).stem + "_CONSOLIDADO.docx",
    }
    if incluir_base:
        dados["corpos"] = list(base.corpora)
        dados["destinos"] = [n.target.label() for n in base.nodes]
    dados.update(extra)
    return _json(**dados)


def versao() -> str:
    return _json(versao=__version__)


def estado() -> str:
    try:
        return _estado()
    except Exception as exc:
        return _erro(exc)


def analisar(base: str, alterador: str, ato: str = "", publicacao: str = "", efeitos: str = "",
             ano_notas: str = "auto") -> str:
    """3  Analisar alterações."""
    try:
        _sem_doc(base, "A legislação base")
        _sem_doc(alterador, "O ato alterador")
        plan, leg = analyze(base, alterador, CACHE, ato or "", publicacao or "", efeitos or "",
                            ano_notas if ano_notas in ANO_NOTAS else "auto")
        ESTADO.update(plan=plan, base=leg)
        return _estado(incluir_base=True, resumo_analise=analysis_summary(plan),
                       ativo=plan.operations[0].id if plan.operations else "")
    except Exception as exc:
        return _erro(exc)


def alternar(op_id: str) -> str:
    """Coluna Aplicar: seleciona ou retira uma proposta pronta."""
    try:
        op = _proposta(op_id)
        if op.status != "pronto":
            return _estado(mensagem="Essa proposta precisa de correção. Confira os campos e clique em Validar correção.")
        op.selected = not op.selected
        return _estado()
    except Exception as exc:
        return _erro(exc)


def corrigir(op_id: str, acao: str, corpo: str, destino: str, alcance: str, efeitos: str, renumerar: str,
             nova: str) -> str:
    """Validar correção (Application.correct da janela do programa)."""
    try:
        plan, base = _plano()
        op = _proposta(op_id)
        if acao not in ACOES:
            raise ValueError("Escolha a ação.")
        if alcance not in ALCANCES:
            raise ValueError("Escolha o alcance.")
        target = Target.parse(destino or "", (corpo or "").strip().upper() or "PRINCIPAL")
        effects = date_iso(efeitos or "")
        if target != op.target or ACOES[acao] not in ("replace", "add"):
            op.remissions = []
            op.related_source = []
            op.note = "Destino alterado manualmente. As remissões automáticas do destino anterior foram retiradas."
        if target != op.target or ACOES[acao] != op.action:
            op.extra_targets = []
            op.renumber_from = ""
        op.target = target
        op.action = ACOES[acao]
        op.scope = ALCANCES[alcance]
        op.effects = effects
        op.rename_to = (renumerar or "").strip()
        op.new_lines = [s.strip() for s in (nova or "").splitlines() if s.strip()]
        op.manual = True
        op.foreign = False
        op.extraction_warning = ""
        validate_plan(plan, base)
        op.selected = op.status == "pronto"
        return _estado(mensagem=op.id + ": " + op.reason, ativo=op.id)
    except Exception as exc:
        return _erro(exc)


def adicionar() -> str:
    """Adicionar mudança manual."""
    try:
        plan, _ = _plano()
        num = max([int(o.id[1:]) for o in plan.operations if o.id[1:].isdigit()] + [0]) + 1
        op = Operation(f"A{num:03d}", "replace", Target(), source="Mudança cadastrada manualmente. Confira o ato alterador.",
                       source_location="cadastro manual", manual=True)
        plan.operations.append(op)
        return _estado(ativo=op.id, mensagem=f"{op.id}: informe a ação, o destino e a nova redação e clique em Validar correção.")
    except Exception as exc:
        return _erro(exc)


def remover(op_id: str) -> str:
    """Remover proposta."""
    try:
        plan, base = _plano()
        plan.operations.remove(_proposta(op_id))
        validate_plan(plan, base)
        return _estado(mensagem=f"Proposta {op_id} removida do plano.")
    except Exception as exc:
        return _erro(exc)


def aprovar_extracao() -> str:
    """Aprovar extração do PDF/HTML: confirma de uma vez as propostas lidas de PDF/HTML já conferidas."""
    try:
        plan, base = _plano()
        waiting = [o for o in plan.operations if o.status == "revisar" and o.extraction_warning.startswith(PDF_REVIEW)]
        if not waiting:
            return _estado(mensagem="Não há propostas de PDF/HTML aguardando conferência.")
        for op in waiting:
            op.extraction_warning = ""
            op.selected = True
        validate_plan(plan, base)
        for op in waiting:
            op.selected = op.status == "pronto"
        aprovadas = sum(o.status == "pronto" for o in waiting)
        return _estado(mensagem=f"{aprovadas} propostas de PDF/HTML aprovadas. As demais continuam pendentes.")
    except Exception as exc:
        return _erro(exc)


def _sincronizar(ato: str, publicacao: str, efeitos: str, ano_notas: str) -> None:
    """Application.sync: os campos do ato, publicação, efeitos e ano nas notas valem na geração."""
    plan, _ = _plano()
    citation = (ato or "").strip()
    if citation != plan.citation:
        for op in plan.operations:
            op.remissions = [s.replace("da " + plan.citation, "da " + citation) for s in op.remissions]
    plan.citation = citation
    plan.date_style = ano_notas if ano_notas in ANO_NOTAS else "auto"
    plan.publication = date_iso(publicacao or "")
    for op in plan.operations:
        op.publication = plan.publication
    if date_iso(efeitos or "") != plan.effects:
        raise ValueError("A data geral de efeitos mudou depois da análise. Analise novamente ou ajuste a data de cada proposta.")


def gerar(destino: str, realcar: bool, ato: str, publicacao: str, efeitos: str, ano_notas: str) -> str:
    """4  Gerar Word atualizado."""
    try:
        plan, _ = _plano()
        _sincronizar(ato, publicacao, efeitos, ano_notas)
        saida = Path(destino)
        saida.parent.mkdir(parents=True, exist_ok=True)
        if saida.exists():
            saida.unlink()
        out = generate(plan, saida, highlight_changes=bool(realcar))
        aplicadas = sum(bool(o.applied) for o in plan.operations)
        pendentes = sum(o.status == "revisar" for o in plan.operations)
        return _estado(word=str(out), aplicadas_agora=aplicadas, pendentes_fora=pendentes)
    except Exception as exc:
        return _erro(exc)


def relatorio(destino: str, ato: str, publicacao: str, efeitos: str, ano_notas: str) -> str:
    """Ver relatório (o plano em HTML)."""
    try:
        _sincronizar(ato, publicacao, efeitos, ano_notas)
        plan, _ = _plano()
        Path(destino).parent.mkdir(parents=True, exist_ok=True)
        report(plan, destino)
        return _json(relatorio=destino)
    except Exception as exc:
        return _erro(exc)


def salvar_plano(destino: str, ato: str, publicacao: str, efeitos: str, ano_notas: str) -> str:
    """Salvar plano (JSON) para continuar depois."""
    try:
        _sincronizar(ato, publicacao, efeitos, ano_notas)
        plan, _ = _plano()
        Path(destino).parent.mkdir(parents=True, exist_ok=True)
        save_plan(plan, destino)
        return _json(plano=destino)
    except Exception as exc:
        return _erro(exc)


def abrir_plano(caminho_json: str, base: str, alterador: str) -> str:
    """Abrir plano salvo. Como no programa (engine.load_plan), os dois arquivos precisam ser os mesmos da
    análise; aqui eles são reconhecidos pelo conteúdo, e não pela pasta, porque a página não guarda pastas.
    Assim também abre o plano salvo no programa do computador."""
    try:
        data = json.loads(Path(caminho_json).read_text(encoding="utf-8-sig"))
        if not isinstance(data, dict) or data.get("format") != "consolidador-sefa" or data.get("version") != 1:
            raise ValueError("Esse arquivo não é um plano desta versão do Consolidador.")
        plan = Plan.from_dict(data["plan"])
        nome_base, nome_alterador = _nome(plan.base_path), _nome(plan.amender_path)
        if not base or not alterador:
            raise ValueError(f"Para abrir este plano, escolha antes a legislação ({nome_base}) e o ato alterador "
                             f"({nome_alterador}) usados nele, sem alterações.")
        if digest(base) != plan.base_hash or digest(alterador) != plan.amender_hash:
            raise ValueError(f"Os arquivos escolhidos não são os deste plano, ou foram alterados depois dele. Escolha "
                             f"{nome_base} e {nome_alterador} como estavam na análise, ou faça uma nova análise.")
        _sem_doc(base, "A legislação base")
        plan.base_path = str(Path(base).resolve())
        plan.amender_path = str(Path(alterador).resolve())
        plan.converted_base = str(prepare_base(plan.base_path, CACHE))
        leg = Legislation(plan.converted_base)
        validate_plan(plan, leg)
        ESTADO.update(plan=plan, base=leg)
        return _estado(incluir_base=True, resumo_analise=analysis_summary(plan),
                       ativo=plan.operations[0].id if plan.operations else "")
    except Exception as exc:
        return _erro(exc)
