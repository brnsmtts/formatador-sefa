from __future__ import annotations
from datetime import datetime
from pathlib import Path
import difflib
import html
from .model import norm
from . import __version__


def report(plan,path,output=""):
    from .engine import protect_outputs
    protect_outputs(plan,[path])
    esc=html.escape
    ready=sum(op.status=="pronto" for op in plan.operations)
    applied=sum(op.applied for op in plan.operations)
    pending=sum(op.status=="revisar" for op in plan.operations)
    actions={"replace":"Substituir", "add":"Acrescentar", "revoke":"Revogar", "rename":"Renumerar"}
    rows=[]
    for op in plan.operations:
        status="APLICADA" if op.applied else op.status.upper()
        before="\n".join(op.before)
        after="\n".join(op.new_lines) if op.action not in ("revoke","rename") else "REVOGADO" if op.action=="revoke" else "Renumerar para § "+op.rename_to
        diff="\n".join(difflib.unified_diff(before.splitlines(),after.splitlines(),fromfile="Redação na base",tofile="Proposta",lineterm=""))
        rows.append(f'''<section class="op"><h2>{esc(op.id)} · {esc(actions.get(op.action,op.action))} · {esc(op.display_label() or 'Destino pendente')}</h2>
<p><strong>{esc(status)}</strong> — {esc(op.reason)}</p><p>{esc(op.target.corpus)} · {esc(op.base_location or 'Base sem destino localizado')}<br>{esc(op.source_location)}<br>Efeitos: {esc(op.effects or 'data não informada')} · Alcance: {'dispositivo completo' if op.scope=='subtree' else 'somente o texto indicado'}</p>
<div class="compare"><div><h3>Redação na base</h3><pre>{esc(before or '(acréscimo ou destino pendente)')}</pre></div><div><h3>Proposta</h3><pre>{esc(after or '(não extraída)')}</pre></div></div>
<details><summary>Trecho do ato alterador</summary><pre>{esc(op.source)}</pre></details><details><summary>Comparação textual</summary><pre>{esc(diff)}</pre></details>
<details><summary>Efeitos e remissões</summary><pre>{esc(op.effects_source or 'Data não extraída de cláusula expressa.')}\n\n{esc(chr(10).join(op.remissions))}\n\n{esc((chr(10)+chr(10)).join(op.related_source))}</pre></details>
{('<p>Observação: '+esc(op.note)+'</p>') if op.note else ''}</section>''')
    warnings="".join("<li>"+esc(s)+"</li>" for s in plan.warnings)
    title="Relatório de consolidação" if output else "Plano de alterações"
    contents=f'''<!doctype html><html lang="pt-BR"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{title}</title>
<style>body{{font:16px/1.55 Segoe UI,Arial,sans-serif;background:#f3f5f8;color:#172431;max-width:1200px;margin:36px auto;padding:0 24px}}h1{{font-size:32px;margin-bottom:8px}}h2{{font-size:20px}}h3{{font-size:16px}}header,.op{{background:white;padding:26px 30px;border:1px solid #d8dfe8;border-radius:10px;margin-bottom:20px}}.stats{{display:flex;gap:30px;flex-wrap:wrap}}.stats b{{display:block;font-size:28px;color:#124e66}}.compare{{display:grid;grid-template-columns:1fr 1fr;gap:24px}}pre{{white-space:pre-wrap;overflow-wrap:anywhere;font:14px/1.6 Segoe UI,Arial,sans-serif;margin:0 0 18px}}summary{{cursor:pointer;font-weight:600;margin:12px 0}}li{{margin:7px 0}}.small{{font-size:13px;color:#546879}}@media(max-width:700px){{.compare{{grid-template-columns:1fr}}}}@media print{{body{{background:white}}.op{{break-inside:avoid}}details{{display:block}}}}</style>
<header><p class="small">CONSOLIDADOR SEFA {__version__} · {datetime.now():%d/%m/%Y %H:%M}</p><h1>{title}</h1><p>Ato alterador: <strong>{esc(plan.citation or 'não informado')}</strong><br>Legislação base: {esc(Path(plan.base_path).name)}<br>Origem: {esc(Path(plan.amender_path).name)}{('<br>Word gerado: '+esc(Path(output).name)) if output else ''}</p>
<div class="stats"><span><b>{len(plan.operations)}</b>propostas</span><span><b>{ready}</b>validadas pela estrutura</span><span><b>{applied}</b>aplicadas</span><span><b>{pending}</b>pendentes de revisão</span></div>
<p>O resultado reúne as mudanças selecionadas. Pendências e mudanças não selecionadas permanecem fora do Word. A revisão jurídica de vigência, alcance e efeitos cabe ao responsável pela consolidação.</p>
<p class="small">As páginas do PDF são páginas físicas verificáveis na origem. Em Word, os locais são blocos do corpo do documento, incluindo parágrafos e tabelas. Não são páginas calculadas; a paginação varia com o Word e a impressora. Os locais da base se referem ao arquivo de entrada.</p>
{('<ul>'+warnings+'</ul>') if warnings else ''}</header>{''.join(rows)}<footer class="small">Os originais não são sobrescritos. O texto que não foi alterado mantém seus elementos e estilos do documento base.</footer></html>'''
    Path(path).write_text(contents,encoding="utf-8")
    return Path(path)
