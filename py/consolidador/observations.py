"""Resumos da análise que distinguem pendências de avisos sobre a base."""
from collections import Counter


def analysis_summary(plan):
    counts=Counter(op.status for op in plan.operations)
    selected=sum(op.selected and op.status=="pronto" for op in plan.operations)
    lines=[f"Propostas identificadas: {len(plan.operations)}",
           f"Prontas: {counts['pronto']} · Selecionadas: {selected}",
           f"Pendentes de revisão: {counts['revisar']}"]
    if counts['já consta']:
        lines.append(f"Redações que já constam na base: {counts['já consta']}")
    if counts['outro ato']:
        lines.append(f"Propostas de outra legislação: {counts['outro ato']}")
    if counts['revisar']:
        lines.extend(("", "Confira as linhas com situação ‘revisar’. Essas propostas não serão aplicadas enquanto estiverem pendentes."))
    elif counts['pronto']:
        lines.extend(("", "Nenhuma proposta ficou pendente. Os avisos sobre outros dispositivos da base não impedem estas alterações."))
    if selected:
        lines.extend(("", "Depois de conferir as propostas, use ‘4  Gerar Word atualizado’ para salvar o documento."))
    elif counts['pronto']:
        lines.extend(("", "Confira as propostas e marque na coluna Aplicar as mudanças que deseja gerar."))
    lines.extend(("", "Os detalhes da análise ficam na aba Observações."))
    return "\n".join(lines)


def analysis_observations(plan):
    lines=[analysis_summary(plan), "", "PENDÊNCIAS DAS PROPOSTAS"]
    pending=[op for op in plan.operations if op.status=="revisar"]
    if pending:
        lines.extend(f"{op.id} — {op.display_label()}: {op.reason}" for op in pending)
    else:
        lines.append("Nenhuma proposta pendente de revisão.")
    lines.extend(("", "OBSERVAÇÕES SOBRE OS ARQUIVOS E A ANÁLISE"))
    lines.extend(plan.warnings or ["Nenhuma observação adicional."])
    return "\n\n".join(lines)
