from __future__ import annotations
from copy import deepcopy
from dataclasses import replace
from datetime import datetime
from pathlib import Path
from io import BytesIO
import json
import os
import re
import tempfile
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from lxml import etree
from .model import Operation, Plan, Target, date_iso, digest, norm
from .index import Legislation, label
from .io import prepare_base, read_amender
from .parser import parse_operations, source_citation, number_value
from .temporal import publication_date, assign_effects, collect_remissions
from .editorial import origin_note, note_for_targets, replace_text, links_signature, update_initial_reference, save_preserving_package
from . import notes as notes_mod
from . import refs as refparser
from . import headings as headmod
from . import annexes as annexmod

ROOT = Path(__file__).resolve().parent.parent


def comparable_lines(lines):
    # Acentos e maiúsculas fazem parte da redação. Só uniformiza espaços.
    return [" ".join(s.split()) for s in lines]


def sort_key(target):
    result = []
    for level, part in enumerate(target.parts):
        if not part:
            result.append((-1, ""))
            continue
        if part == "UNICO":
            result.append((1, ""))
            continue
        pieces = part.split("-", 1)
        try:
            n = number_value(pieces[0]) if level == 2 else int(pieces[0]) if level != 3 else ord(part[0])
        except (ValueError, KeyError, TypeError):
            n = 0
        result.append((n, pieces[1] if len(pieces) > 1 else ""))
    return tuple(result)


def _head_history_end(base, node, begin):
    """Fim do histórico que só guarda versões anteriores do próprio texto do
    dispositivo (nota + "Art. 204-A ..." antigo), logo depois desse texto."""
    level, number = node.target.level, node.target.parts[node.target.level]
    i = begin
    while i < node.end:
        j = i
        while j < node.end and base.blocks[j].role == "old_note":
            j += 1
        k = j
        while k < node.end and base.blocks[k].role == "old":
            info = label(base.blocks[k].text)
            if info and (info[0] != level or info[1] != number):
                return i
            k += 1
        if j == i or k == j:
            return i
        i = k
    return i


def tail_anchor(base, node, level=None):
    """Insere antes do histórico de um ancestral, depois do último filho atual.

    Primeiro filho de um dispositivo (sem irmãos): fica depois do histórico do
    próprio caput, como em "Redação dada ao caput do art. 204-A ... / histórico /
    Acrescidos os §§ 1º a 8º" (Constituição do Pará, EC 102/26)."""
    own = node.target.level if level is None else None
    level = node.target.level if level is None else level
    number = node.target.parts[node.target.level]
    children = [n for n in base.nodes if node.start < n.start < node.end and node.target.contains(n.target)]
    begin = max([node.head_end] + [n.head_end for n in children])
    if own is None and not children:
        begin = _head_history_end(base, node, begin)
    for i in range(begin, node.end):
        b = base.blocks[i]
        info = label(b.text) if b.role == 'old' else None
        # Histórico de um ancestral: rótulo superior, ou do mesmo nível mas de outro irmão.
        # O "Parágrafo único" arquivado logo abaixo do § 1º é a versão anterior do
        # próprio § 1º (renumerado), e não de outro irmão.
        renumbered = own == 1 and number == "1" and info and info[1] == "UNICO"
        if info and (info[0] < level or (own is not None and info[0] == own and info[1] != number and not renumbered)):
            anchor = i
            while anchor > begin and (base.blocks[anchor - 1].role == 'old_note' or not base.blocks[anchor - 1].text):
                anchor -= 1
            return anchor
    return node.end


def insertion_anchor(base, target):
    if not target.article or target.corpus not in base.corpora:
        return None, "O artigo ou o corpo da legislação não foi identificado."
    if target.level > 0:
        parent = base.find(target.parent())
        if not parent:
            return None, "O dispositivo superior não foi encontrado de forma única."
        siblings = [n for n in base.nodes if n.target.level == target.level and n.target.parent() == target.parent()]
        if target.level == 1:
            if any(n.target.paragraph == "UNICO" for n in siblings) and target.paragraph != "UNICO":
                return None, "Existe parágrafo único. Cadastre primeiro sua renumeração se o ato assim determinar."
            if target.paragraph == "UNICO" and siblings:
                return None, "Já existem parágrafos numerados neste artigo."
        before = [n for n in siblings if sort_key(n.target) > sort_key(target)]
        after = [n for n in siblings if sort_key(n.target) < sort_key(target)]
        if after:
            # Logo depois do irmão anterior (e do seu histórico), antes do histórico do superior.
            anchor = tail_anchor(base, max(after, key=lambda n: sort_key(n.target)))
            if before:
                anchor = min(anchor, min(before, key=lambda n: sort_key(n.target)).lead)
            return anchor, ""
        if before:
            return min(before, key=lambda n: sort_key(n.target)).lead, ""
        # Incisos do caput são inseridos antes dos parágrafos do artigo.
        children = [n for n in base.nodes if parent.start < n.start < parent.end and not target.contains(n.target)]
        if children and target.level == 2 and not target.paragraph:
            pars = [n for n in children if n.target.level == 1]
            if pars:
                return min(n.lead for n in pars), ""
        # Sem irmãos, insere após o texto atual e antes do histórico do pai.
        return tail_anchor(base, parent, target.level), ""
    siblings = [n for n in base.nodes if n.target.corpus == target.corpus and n.target.level == 0]
    if not siblings:
        return None, "Nenhum artigo foi localizado neste corpo da legislação."
    before = [n for n in siblings if sort_key(n.target) > sort_key(target)]
    after = [n for n in siblings if sort_key(n.target) < sort_key(target)]
    if before:
        nxt = min(before, key=lambda n: sort_key(n.target))
        if after:
            # "Art. 169-A" fica com o art. 169: antes do título da seção seguinte, se houver.
            prev = max(after, key=lambda n: sort_key(n.target))
            if prev.start < nxt.start and prev.end < nxt.lead:
                return prev.end, ""
        return nxt.lead, ""
    last = max(siblings, key=lambda n: sort_key(n.target))
    anchor = last.end
    # Novo último artigo: antes das linhas em branco que antecedem o fecho do ato.
    while anchor - 1 > last.start and base.blocks[anchor - 1].role == "current" and not base.blocks[anchor - 1].text:
        anchor -= 1
    return anchor, ""


def has_children(base, node):
    return bool(node) and any(node.start < n.start < node.end and node.target.contains(n.target) for n in base.nodes)


def source_target(op):
    """Destino que existe hoje na base (o parágrafo único, numa renumeração)."""
    if op.renumber_from:
        return replace(op.target, paragraph=op.renumber_from)
    return op.target


def _fail(op, reason, status="revisar"):
    op.status = status
    op.reason = reason
    op.selected = False


def validate_operation(op, base, default_effects=""):
    op.applied = False
    op.before = []
    op.base_location = ""
    op.effects = date_iso(op.effects or default_effects)
    op.effects_until = date_iso(op.effects_until or "")
    op.status = "revisar"
    op.reason = ""
    if op.foreign and not op.manual:
        return _fail(op, op.extraction_warning or "O comando pertence a outra legislação.", "outro ato")
    if op.effects_until and (not op.effects or op.effects_until < op.effects):
        return _fail(op, "O fim do período de efeitos é anterior ao início. Confira as datas.")
    if op.action == "heading":
        return _validate_heading(op, base)
    if op.action == "annex":
        return _validate_annex(op, base)
    if op.action == "note":
        return _validate_note(op, base)
    if op.action not in ("replace", "add", "revoke", "rename"):
        return _fail(op, "Escolha substituir, acrescentar, revogar ou renumerar.")
    if op.action in ("add", "replace") and op.new_lines and op.target.article:
        # Os irmãos acrescidos (ou com nova redação) na mesma citação formam uma só
        # nota: "Acrescidos os §§ 23 a 30", "Redação dada às alíneas “b” e “c”".
        context, roots = list(op.target.parts), []
        for text in op.new_lines:
            info = label(text)
            if info:
                context[info[0]:] = [info[1]] + [""] * (4 - info[0])
                if info[0] == op.target.level:
                    roots.append(Target(op.target.corpus, *context))
        if roots and roots[0] == op.target and all(r.parent() == op.target.parent() for r in roots):
            op.extra_targets = roots[1:]
        elif op.action == "replace":
            op.extra_targets = []
    elif op.action not in ("add", "revoke"):
        op.extra_targets = []
    targets = op.targets
    if len(targets) > 1:
        if op.action not in ("add", "revoke", "replace"):
            return _fail(op, "Só acréscimos, revogações e novas redações podem reunir vários dispositivos numa mesma nota.")
        if any(t.level != op.target.level or t.parent() != op.target.parent() for t in targets) or len(set(targets)) != len(targets):
            return _fail(op, "Os dispositivos reunidos precisam ser irmãos do mesmo dispositivo superior.")
        if op.action == "replace":
            if op.scope != "head" or op.renumber_from or op.fold or op.caput_forced:
                return _fail(op, "A nova redação reunida aceita só o texto de irmãos sem filhos. Separe as propostas.")
            nodes = [base.find(t) for t in targets]
            if not all(nodes):
                missing = [t.label() for t, n in zip(targets, nodes) if not n]
                return _fail(op, "Destino não encontrado no texto atual: " + ", ".join(missing) + ".")
            if shared_history(base, nodes, "head") is None:
                return _fail(op, "Os dispositivos reunidos têm notas, filhos ou históricos próprios. Separe a nova redação em propostas próprias.")
    src = source_target(op)
    node = base.find(src)
    # "caput" quando o dispositivo tem (ou passa a ter, na mesma mudança) filhos.
    op.caput = bool(op.action == "replace" and op.scope == "head"
                    and (has_children(base, node) or op.fold or op.caput_forced))
    if node:
        op.before = base.current_lines(node, op.scope == "subtree")
        op.base_location = base.location(node)
    if op.action in ("revoke", "replace") and len(targets) > 1:
        nodes = [base.find(t) for t in targets]
        if all(nodes):
            op.before = [s for n in nodes for s in base.current_lines(n, op.action == "revoke")]
    for t in targets + ([src] if src != op.target else []):
        if len(base.by_target.get(t, [])) > 1:
            return _fail(op, "Numeração duplicada no texto atual. Resolva a duplicidade no Word antes de aplicar.")
    if not op.target.article:
        return _fail(op, "Informe o artigo e a hierarquia completa do destino.")
    if op.action in ("replace", "revoke", "rename") and not node:
        return _fail(op, "Destino não encontrado no texto atual. O histórico não é usado como destino.")
    if op.renumber_from:
        if op.action != "replace" or op.target.level != 1 or op.renumber_from != "UNICO":
            return _fail(op, "A renumeração automática aceita parágrafo único passando a § 1º.")
        if base.by_target.get(op.target):
            return _fail(op, "O parágrafo de destino da renumeração já existe.")
    if op.action == "revoke" and len(targets) > 1:
        nodes = [base.find(t) for t in targets]
        if not all(nodes):
            missing = [t.label() for t, n in zip(targets, nodes) if not n]
            return _fail(op, "Destino não encontrado no texto atual: " + ", ".join(missing) + ".")
        siblings = sorted([n for n in base.nodes if n.target.level == op.target.level and n.target.parent() == op.target.parent()], key=lambda n: n.start)
        positions = [siblings.index(n) for n in nodes]
        if positions != list(range(positions[0], positions[0] + len(nodes))):
            return _fail(op, "Os dispositivos revogados não estão em sequência na base. Separe a revogação em propostas próprias.")
    if op.action == "add":
        for t in targets:
            existing = base.find(t)
            if existing:
                current = base.current_lines(existing, True)
                if len(targets) == 1 and comparable_lines(current) == comparable_lines(op.new_lines):
                    return _fail(op, "A mesma redação já consta no destino.", "já consta")
                return _fail(op, f"O destino do acréscimo ({t.label()}) já existe. Confira a ação e o rótulo.")
        anchor, reason = insertion_anchor(base, op.target)
        if reason:
            return _fail(op, reason)
        if op.heading:
            info = headmod.parse_line(op.heading[0])
            if not info or op.target.level != 0:
                return _fail(op, "O título novo deve começar por Capítulo, Seção etc. e anteceder um artigo.")
            kind, number, _ = info
            op.heading_key = f"{kind}:{number}"
            parent = headmod.parent_at(base, anchor, kind)
            op.heading_parents = f"{parent[0]}:{parent[1]}" if parent else ""
            if headmod.find(base, kind, number, op.target.corpus, [parent] if parent else ()):
                return _fail(op, f"O título {headmod.name(kind, number)} já existe na base. Confira se é nova denominação.")
        if len(targets) > 1:
            siblings = [n for n in base.nodes if n.target.level == op.target.level and n.target.parent() == op.target.parent()]
            low, high = sort_key(targets[0]), sort_key(targets[-1])
            if any(low < sort_key(n.target) < high for n in siblings):
                return _fail(op, "Há dispositivos existentes entre os acréscimos reunidos. Separe o acréscimo em propostas próprias.")
        op.base_location = f"{op.target.corpus}, inserir antes do bloco {anchor + 1} do Word"
    if op.action in ("replace", "add"):
        if not op.new_lines and op.extraction_warning and not op.manual:
            return _fail(op, op.extraction_warning.strip())
        if not op.new_lines or any(not s.strip() for s in op.new_lines):
            return _fail(op, "Informe a nova redação completa, com os rótulos dos dispositivos.")
        first = label(op.new_lines[0])
        if not first or first[0] != op.target.level or first[1] != op.target.parts[op.target.level]:
            return _fail(op, "O rótulo da nova redação não corresponde ao destino selecionado.")
        if any(re.search(r"\.{3,}|…", s) for s in op.new_lines):
            return _fail(op, "Há reticências no novo texto. Informe somente a redação completa que será inserida.")
        # Impede que uma citação de outro artigo entre no interior deste.
        context = list(op.target.parts)
        new_targets = []
        roots = set(targets)
        for text in op.new_lines:
            info = label(text)
            if not info:
                continue
            level, num, *_ = info
            context[level:] = [num] + [""] * (4 - level)
            t = Target(op.target.corpus, *context)
            if not any(r.contains(t) for r in roots):
                return _fail(op, "O novo texto contém dispositivo fora da hierarquia do destino.")
            if t in new_targets:
                return _fail(op, "A nova redação repete um destino na mesma hierarquia.")
            new_targets.append(t)
        if [t for t in new_targets if t.level == op.target.level] != targets:
            return _fail(op, "Os dispositivos da nova redação não correspondem aos destinos reunidos.")
        descendants = [t for t in new_targets if t not in roots]
        if op.scope == "head" and descendants:
            # Só aceita filhos novos junto do caput quando o dispositivo ainda não os tinha.
            child_levels = {t.level for t in descendants}
            existing = [n for n in base.nodes if node and node.start < n.start < node.end and n.target.level in child_levels
                        and n.target.parent() == src]
            if op.action != "replace" or existing or any(base.by_target.get(t) for t in descendants):
                return _fail(op, "Há descendentes na nova redação. Escolha Dispositivo completo ou separe as operações.")
        if node and op.action == "replace" and not op.renumber_from and comparable_lines(op.before if len(targets) > 1 else base.current_lines(node, op.scope == "subtree")) == comparable_lines(op.new_lines):
            return _fail(op, "A mesma redação já consta no destino.", "já consta")
    if op.action == "revoke" and node:
        if all(re.search(r"\b(?:REVOGAD|SUPRIMID)[OA]\b|\((?:revogad|suprimid)[oa]\)", (base.find(t) or node).text, re.I) for t in targets):
            return _fail(op, "O destino já está marcado como revogado.", "já consta")
        op.scope = "subtree"
    if op.action == "rename":
        level = op.target.level
        number = op.rename_to.strip().upper() if level != 3 else op.rename_to.strip().lower()
        formats = {1: r"\d+(?:-[A-Z])?", 2: r"[IVXLCDM]+(?:-[A-Z])?", 3: r"[a-z](?:-[A-Z])?", 4: r"\d+(?:\.\d+)*"}
        if level < 1:
            return _fail(op, "Artigos e títulos não são renumerados (LC 95/98, art. 12, III, b). Use o número com letra (art. 5º-A).")
        if not re.fullmatch(formats[level], number):
            return _fail(op, "Informe o novo número no formato do dispositivo (§: 3; inciso: VIII; alínea: d).")
        op.rename_to = number
        destination = Target(op.target.corpus, *[number if k == level else v for k, v in enumerate(op.target.parts)])
        if destination == op.target:
            return _fail(op, "O novo número é igual ao atual.")
        if base.by_target.get(destination):
            return _fail(op, "O número de destino da renumeração já existe.")
    for t in targets:
        n = base.find(source_target(op) if t == op.target else t)
        if not n:
            continue
        stop = n.end if op.scope == "subtree" else n.head_end
        if any(b.role in ("table", "other") for b in base.blocks[n.start:stop]):
            return _fail(op, "O intervalo contém tabela ou estrutura não textual. Faça essa alteração no Word.")
    if node and op.scope == "subtree" and op.action == "replace" and not op.manual and not op.full_text:
        old_children = [n for n in base.nodes if node.start < n.start < node.end]
        mapped = set(new_targets)
        if op.renumber_from:
            mapped = {replace(t, paragraph=op.renumber_from) if t.paragraph == op.target.paragraph else t for t in mapped}
        missing = [n for n in old_children if n.target not in mapped]
        if missing:
            return _fail(op, "A citação completa omitiria descendentes existentes. Confira se é substituição integral ou parcial.")
    if op.effects and node:
        for b in base.blocks[node.lead:node.start]:
            if b.role == "new_note" and note_matches(b.text, src):
                m = re.search(r"efeitos\s+a\s+partir\s+de\s+(\d{1,2}[./]\d{1,2}[./]\d{2,4})", b.text, re.I)
                if m:
                    try:
                        prior = date_iso(m[1])
                    except ValueError:
                        continue
                    # Redação futura alterada ou revogada antes de produzir efeitos:
                    # a nota anterior passa a "sem efeitos" (padrão SEFA desde 2025).
                    never_in_force = bool(op.publication and op.publication < prior)
                    if op.effects <= prior and not never_in_force:
                        return _fail(op, "A data proposta não é posterior à data da redação atual. Confira as versões temporais antes de consolidar.")
                    if op.effects <= prior:
                        never_note = "A redação atual ainda não produzia efeitos quando o ato foi publicado: no histórico, a nota dela receberá “sem efeitos”."
                        if never_note not in op.note:
                            op.note = (op.note + " " if op.note else "") + never_note
    if op.later_versions:
        previous = op.effects
        for v in op.later_versions:
            lines = v.get("lines", [])
            first = label(lines[0]) if lines else None
            if not first or first[0] != op.target.level or first[1] != op.target.parts[op.target.level]:
                return _fail(op, "Uma das redações por período não começa pelo rótulo do destino.")
            if any(re.search(r"\.{3,}|…", s) for s in lines):
                return _fail(op, "Há reticências numa das redações por período.")
            if not v.get("effects") or not previous or v["effects"] <= previous:
                return _fail(op, "As redações por período precisam de datas de efeitos crescentes.")
            previous = v["effects"]
    if op.extraction_warning and not op.manual:
        return _fail(op, op.extraction_warning.strip())
    op.status = "pronto"
    op.reason = "Destino e redação conferidos por regras." if not op.manual else "Correção manual validada pela estrutura."
    if op.later_versions:
        op.reason += f" O ato traz {len(op.later_versions) + 1} redações deste dispositivo, por período; todas entram no texto vigente, a mais recente primeiro."
    if op.note and not op.manual:
        op.reason += " " + op.note


def heading_group(op, base):
    kind, _, number = op.heading_key.partition(":")
    parents = [tuple(x.split(":", 1)) for x in op.heading_parents.split("/") if x]
    corpus = op.target.corpus
    group = headmod.find(base, kind, number, corpus, parents)
    if group is None and op.target.article:
        group = headmod.find(base, kind, number, corpus, article=op.target.article)
    return group


def _validate_heading(op, base):
    """Nova denominação de Título, Capítulo, Seção ou Subseção."""
    if not op.heading_key and op.new_lines:
        info = headmod.parse_line(op.new_lines[0])
        if info:
            op.heading_key = f"{info[0]}:{info[1]}"
    if not op.heading_key:
        return _fail(op, "Informe o título com o seu rótulo, por exemplo: SEÇÃO V / Da Justiça Militar.")
    group = heading_group(op, base)
    if group is None:
        return _fail(op, "O título não foi localizado na base de forma única. Informe o artigo que vem logo depois dele.")
    op.before = list(group.lines)
    op.base_location = f"{group.corpus}, bloco {group.start + 1} do Word"
    if not op.new_lines or any(not s.strip() for s in op.new_lines):
        return _fail(op, "Informe a nova redação do título, uma linha por parágrafo (rótulo e denominação).")
    first = headmod.parse_line(op.new_lines[0])
    if op.heading_rename:
        if not first or first[0] != op.heading_key.partition(":")[0] or len(op.new_lines) != 1:
            return _fail(op, "Informe só o novo rótulo do título (por exemplo, Seção I).")
    elif not first or f"{first[0]}:{first[1]}" != op.heading_key:
        return _fail(op, "A nova redação deve começar pelo mesmo rótulo do título (por exemplo, SEÇÃO V).")
    if norm(" ".join(op.new_lines)) == norm(" ".join(group.lines)):
        return _fail(op, "A mesma redação já consta no título.", "já consta")
    if op.extraction_warning and not op.manual:
        return _fail(op, op.extraction_warning.strip())
    op.status = "pronto"
    op.reason = ("Título localizado na base; a nova denominação foi conferida." if not op.manual
                 else "Correção manual validada pela estrutura.")


def _apply_heading(doc, base, op, citation, profile, highlight_changes, style):
    group = heading_group(op, base)
    blocks = base.blocks[group.start:group.end]
    if op.heading_rename:
        blocks = blocks[:1]
    fresh = [paragraph(doc, text, profile["new_note"]) for text in notes_mod.new_notes(op, citation, style)]
    for k, line in enumerate(op.new_lines):
        el = deepcopy(blocks[min(k, len(blocks) - 1)].element)
        fresh.append(replace_text(el, line))
    highlight_new_elements(doc, fresh, highlight_changes)
    kind, _, number = op.heading_key.partition(":")
    wanted = norm(headmod.name(kind, number))
    prior = []
    k = group.start - 1
    while k >= 0 and base.blocks[k].role == "new_note" and headmod.HEADING_NOTE.match(base.blocks[k].text):
        if wanted in norm(base.blocks[k].text):
            prior.insert(0, base.blocks[k])
        k -= 1
    if prior:
        archived = [replace_text(historical(b.element, doc, profile["old_note"]), prior_note(b.text, op.effects, citation, op.publication, style)) for b in prior]
    else:
        archived = [paragraph(doc, notes_mod.fallback_note(True, op.effects, style), profile["old_note"])]
    archived += [historical(b.element, doc, profile["old"]) for b in blocks]
    anchor = blocks[0].element
    for el in fresh + archived:
        anchor.addprevious(el)
    for el in [b.element for b in blocks] + [b.element for b in prior]:
        if el.getparent() is not None:
            el.getparent().remove(el)
    op.applied = True


def _validate_note(op, base):
    """Nota sem mudança de texto: Vide, decisão judicial (ADI), regulamentação."""
    node = base.find(op.target)
    if not op.target.article or node is None:
        return _fail(op, "Informe o dispositivo que recebe a nota (ex.: art. 12 / inciso I / alínea b).")
    lines = [s.strip() for s in op.new_lines if s.strip()]
    if not lines:
        return _fail(op, "Escreva a nota na aba Nova redação (uma nota por linha).")
    if any(label(s) for s in lines):
        return _fail(op, "A nota não deve começar por rótulo de dispositivo. Para mudar o texto, use Substituir redação.")
    op.new_lines = lines
    op.before = base.current_lines(node)
    op.base_location = base.location(node)
    op.status = "pronto"
    op.reason = "A nota entra logo acima do dispositivo; o texto não muda."


_ACT_CACHE = {}


def act_document(path):
    if path not in _ACT_CACHE:
        _ACT_CACHE[path] = Document(path)
    return _ACT_CACHE[path]


def _text_of(el):
    from .index import paragraph_text
    if el.tag == qn("w:tbl"):
        rows = []
        for tr in el.iter(qn("w:tr")):
            cells = [" ".join(paragraph_text(p) for p in tc.iter(qn("w:p"))).strip() for tc in tr.iter(qn("w:tc"))]
            rows.append(" | ".join(cells))
        return "[TABELA] " + " / ".join(rows[:3]) + (" / ..." if len(rows) > 3 else "")
    return paragraph_text(el)


def _validate_annex(op, base):
    """Anexo inteiro: nova redação, acréscimo ou revogação."""
    if op.annex_mode not in ("replace", "add", "revoke") or not op.annex_key:
        return _fail(op, "Indique o anexo e se ele recebe nova redação, é acrescentado ou revogado.")
    regions = annexmod.regions(base)
    region = regions.get(op.annex_key)
    if op.annex_mode in ("replace", "revoke") and region is None:
        return _fail(op, f"O Anexo {op.annex_key} não foi localizado de forma única na base.")
    if op.annex_mode == "add" and region is not None:
        return _fail(op, f"O Anexo {op.annex_key} já existe na base. Confira se é nova redação.")
    if region is not None:
        op.before = [b.text for b in base.blocks[region.start:region.current_end] if b.text]
        op.base_location = f"ANEXO {op.annex_key}, bloco {region.start + 1} do Word"
        mixed = [b for b in base.blocks[region.current_end:region.end] if b.role in ("new_note", "heading") and b.text]
        if mixed:
            # Ex.: Anexo III da LC 123, com histórico por tabela no meio do texto vigente.
            return _fail(op, f"O Anexo {op.annex_key} tem histórico por tabela no meio do texto vigente. Faça esta mudança no Word.")
    if op.annex_mode in ("replace", "add"):
        if not op.annex_path:
            return _fail(op, "O texto do anexo só é copiado de ato em Word (.docx ou .doc). Faça esta mudança no Word.")
        content = annexmod.act_annexes(act_document(op.annex_path)).get(op.annex_source)
        if not content:
            return _fail(op, f"O Anexo {op.annex_source} não foi encontrado no ato alterador.")
        if annexmod.has_external_parts(content):
            return _fail(op, "O anexo do ato tem imagens, objetos ou links. Copie este anexo no Word.")
        op.new_lines = [f"ANEXO {op.annex_key}"] + [t for t in (_text_of(el) for el in content) if t.strip()]
    if op.extraction_warning and not op.manual:
        return _fail(op, op.extraction_warning.strip())
    op.status = "pronto"
    op.reason = {"replace": "Anexo localizado na base e no ato; o anexo atual vai para o histórico.",
                 "add": "Anexo novo localizado no ato.",
                 "revoke": "Anexo localizado na base; o texto vai para o histórico."}[op.annex_mode]


def _style_ids(doc):
    ids, by_name = set(), {}
    for st in doc.styles.element.findall(qn("w:style")):
        sid = st.get(qn("w:styleId"))
        name = st.find(qn("w:name"))
        ids.add(sid)
        if name is not None:
            by_name[name.get(qn("w:val"))] = sid
    return ids, by_name


def _ensure_template_style(doc, name):
    """Copia do modelo SEFA um estilo que a base ainda não tem (estilos 'ant' das tabelas)."""
    ids, by_name = _style_ids(doc)
    if name in by_name:
        return by_name[name]
    source = etree.parse(str(ROOT / "modelos/estilos_sefa.xml")).getroot()
    for st in source:
        n = st.find(qn("w:name"))
        if n is not None and n.get(qn("w:val")) == name:
            doc.styles.element.append(deepcopy(st))
            return st.get(qn("w:styleId"))
    return None


def _import_elements(doc, act, elements, heading_id, text_id, table_id):
    """Copia parágrafos e tabelas do anexo do ato, com estilos da base."""
    base_ids, base_by_name = _style_ids(doc)
    act_names = {st.get(qn("w:styleId")): (st.find(qn("w:name")).get(qn("w:val")) if st.find(qn("w:name")) is not None else "")
                 for st in act.styles.element.findall(qn("w:style"))}
    from .index import paragraph_text
    out = []
    for el in elements:
        c = deepcopy(el)
        in_table = c.tag == qn("w:tbl")
        for tbl_style in c.iter(qn("w:tblStyle")):
            if act_names.get(tbl_style.get(qn("w:val"))) not in base_by_name:
                tbl_style.getparent().remove(tbl_style)
            else:
                tbl_style.set(qn("w:val"), base_by_name[act_names[tbl_style.get(qn("w:val"))]])
        for par in c.iter(qn("w:p")):
            ppr = par.find(qn("w:pPr"))
            if ppr is None:
                ppr = etree.SubElement(par, qn("w:pPr"))
                par.remove(ppr)
                par.insert(0, ppr)
            for num in list(ppr.findall(qn("w:numPr"))):
                ppr.remove(num)
            ps = ppr.find(qn("w:pStyle"))
            name = act_names.get(ps.get(qn("w:val"))) if ps is not None else None
            # Nas tabelas, mantém o estilo de mesmo nome da base; fora delas, títulos
            # curtos recebem o estilo de título dos anexos da base (como nas bases SEFA).
            target = base_by_name.get(name) if in_table and name and name not in ("Normal", "Default") else None
            if target is None:
                if in_table:
                    target = table_id
                else:
                    text = paragraph_text(par).strip()
                    target = heading_id if text and len(text) <= 140 and not text.endswith((".", ";", ":")) else text_id
            if ps is None and target:
                ps = etree.SubElement(ppr, qn("w:pStyle"))
                ppr.remove(ps)
                ppr.insert(0, ps)
            if target:
                ps.set(qn("w:val"), target)
            elif ps is not None:
                ppr.remove(ps)
        for rs in list(c.iter(qn("w:rStyle"))):
            if act_names.get(rs.get(qn("w:val"))) not in base_by_name:
                rs.getparent().remove(rs)
            else:
                rs.set(qn("w:val"), base_by_name[act_names[rs.get(qn("w:val"))]])
        out.append(c)
    return out


def _archive_table(doc, el):
    """Tabela que vai para o histórico: estilos 'ant' (verdes) do modelo SEFA."""
    c = deepcopy(el)
    ids, by_name = _style_ids(doc)
    names = {v: k for k, v in by_name.items()}
    default = _ensure_template_style(doc, "a19_tabela_centroant")
    for par in c.iter(qn("w:p")):
        ppr = par.find(qn("w:pPr"))
        if ppr is None:
            ppr = OxmlElement("w:pPr")
            par.insert(0, ppr)
        ps = ppr.find(qn("w:pStyle"))
        current = names.get(ps.get(qn("w:val"))) if ps is not None else None
        wanted = annexmod.ANT.get(current) if current not in annexmod.ANT.values() else current
        target = _ensure_template_style(doc, wanted) if wanted else default
        if target:
            if ps is None:
                ps = OxmlElement("w:pStyle")
                ppr.insert(0, ps)
            ps.set(qn("w:val"), target)
    for rpr in c.iter(qn("w:rPr")):
        for color in list(rpr.findall(qn("w:color"))):
            rpr.remove(color)
    return c


def _apply_annex(doc, base, op, citation, profile, highlight_changes, style):
    regions = annexmod.regions(base)
    region = regions.get(op.annex_key)
    body = doc.element.body
    fresh = [paragraph(doc, text, profile["new_note"]) for text in notes_mod.new_notes(op, citation, style)]
    label_el = base.blocks[region.start].element if region else None
    hstyle = heading_style(base)
    hid = _style_ids(doc)[1].get(hstyle) or _style_ids(doc)[1].get(profile["current"])
    if label_el is not None:
        new_label = deepcopy(label_el)
        replace_text(new_label, f"ANEXO {op.annex_key}" + (" - REVOGADO" if op.annex_mode == "revoke" else ""))
    else:
        new_label = paragraph(doc, f"ANEXO {op.annex_key}", hstyle if hstyle in doc.styles else profile["current"])
    fresh.append(new_label)
    if op.annex_mode in ("replace", "add"):
        act = act_document(op.annex_path)
        content = annexmod.act_annexes(act)[op.annex_source]
        ids, by_name = _style_ids(doc)
        table_id = by_name.get("a15_tabela_centro")
        fresh.extend(_import_elements(doc, act, content, hid, by_name.get(profile["current"]), table_id))
    highlight_new_elements(doc, [e for e in fresh if e.tag == qn("w:p")], highlight_changes)
    if op.annex_mode == "add":
        later = sorted((r for r in regions.values() if annexmod.value(r.key) > annexmod.value(op.annex_key)), key=lambda r: r.lead)
        if later:
            anchor = base.blocks[later[0].lead].element
            for el in fresh:
                anchor.addprevious(el)
        else:
            last = body[-1]
            for el in fresh:
                if last.tag == qn("w:sectPr"):
                    last.addprevious(el)
                else:
                    body.append(el)
        op.applied = True
        return
    wanted = norm(f"anexo {op.annex_key}")
    prior = [b for b in base.blocks[region.lead:region.start] if b.role == "new_note" and re.search(r"\b" + re.escape(wanted) + r"\b", norm(b.text))]
    if prior:
        archived = [replace_text(historical(b.element, doc, profile["old_note"]), prior_note(b.text, op.effects, citation, op.publication, style)) for b in prior]
    else:
        archived = [paragraph(doc, notes_mod.fallback_note(True, op.effects, style), profile["old_note"])]
    current = base.blocks[region.start:region.current_end]
    for b in current:
        if b.element.tag == qn("w:tbl"):
            archived.append(_archive_table(doc, b.element))
        elif b.element.tag == qn("w:p"):
            archived.append(historical(b.element, doc, profile["old"]))
    for el in fresh + archived:
        label_el.addprevious(el)
    for el in [b.element for b in current] + [b.element for b in prior]:
        if el.getparent() is not None:
            el.getparent().remove(el)
    op.applied = True


def heading_style(base):
    from collections import Counter
    counts = Counter(b.style for b in base.blocks if b.role == "heading" and headmod.parse_line(b.text))
    return counts.most_common(1)[0][0] if counts else "a08_titulo"


DEPENDENT = ("Destino não encontrado", "superior não foi encontrado", "parágrafo único", "já existe", "Existe parágrafo")


def validate_plan(plan, base=None):
    base = base or Legislation(plan.converted_base)
    for op in plan.operations:
        was_selected = op.selected
        validate_operation(op, base, plan.effects)
        if op.status == "pronto":
            op.selected = was_selected
    # Uma mudança pode depender de outra do mesmo plano (renumeração, novo artigo ou
    # novo parágrafo), venha ela antes ou depois na ordem do ato: "acrescido o § 2º
    # ao art. 5º" (art. 2º do decreto) só cabe depois de "renumerado o parágrafo único
    # do art. 5º para § 1º" (art. 4º). Simula as mudanças numa cópia, na mesma ordem
    # da geração: as dependentes que ainda não cabem esperam a rodada seguinte.
    waiting = [i for i, op in enumerate(plan.operations)
               if op.status == "revisar" and any(s in op.reason for s in DEPENDENT)
               and any(x.status == "pronto" and x.selected for j, x in enumerate(plan.operations) if j != i)]
    if waiting:
        stream = BytesIO()
        base.doc.save(stream)
        stream.seek(0)
        simulation = Document(stream)
        settings = note_settings(plan, base)
        queue = [i for i, op in enumerate(plan.operations) if i in waiting or (op.status == "pronto" and op.selected)]
        while queue:
            later = []
            for i in queue:
                op = plan.operations[i]
                try:
                    if i in waiting:
                        validate_operation(op, Legislation(simulation), plan.effects)
                        if op.status != "pronto":
                            if any(s in op.reason for s in DEPENDENT):
                                later.append(i)
                            continue
                        op.selected = True
                        op.base_location = "Aplicada depois das outras mudanças deste plano de que depende."
                    apply_operation(simulation, deepcopy(op), plan.citation or "Ato alterador", base.profile, **settings)
                except ValueError:
                    continue
            if len(later) == len(queue) or not later:
                break
            queue = later
    for i, op in enumerate(plan.operations):
        if op.status != "pronto" or op.action in ("heading", "annex", "note"):
            continue
        for other in plan.operations[i + 1:]:
            if other.status != "pronto" or other.action in ("heading", "annex", "note"):
                continue
            if "rename" in (op.action, other.action) and {op.action, other.action} <= {"rename", "add"}:
                # Renumeração libera o número que o acréscimo ocupa; a ordem é conferida na simulação.
                continue
            conflict = any(a == b or (op.scope == "subtree" and a.contains(b)) or (other.scope == "subtree" and b.contains(a))
                           for a in op.targets for b in other.targets)
            if conflict:
                op.status = other.status = "revisar"
                op.selected = other.selected = False
                op.reason = other.reason = f"Operações sobrepostas: {op.id} e {other.id}. Una ou separe as mudanças antes de aplicar."
    return plan


def note_settings(plan, base=None):
    style = plan.resolved_date_style if plan.date_style in ("", "auto") else plan.date_style
    return {"style": style or "4", "corpus_suffix": plan.adct_suffix}


def analyze(base_path, amender_path, cache, citation="", publication="", effects="", date_style="auto"):
    converted = prepare_base(base_path, cache)
    base = Legislation(converted)
    if not base.nodes:
        raise ValueError("Não foram encontrados artigos no texto atual da legislação base.")
    lines, warnings = read_amender(amender_path, cache)
    resolved_style = notes_mod.detect_date_style(base.blocks)
    style = resolved_style if date_style in ("", "auto", None) else date_style
    citation = citation.strip() or notes_mod.format_citation(source_citation(lines), style)
    publication = date_iso(publication) or publication_date(lines)
    plan = Plan(str(Path(base_path).resolve()), str(Path(amender_path).resolve()), str(converted), citation, publication, date_iso(effects),
                digest(base_path), digest(amender_path), warnings=base.warnings + warnings, style_profile=base.profile, identity=base.identity)
    plan.date_style = date_style or "auto"
    plan.resolved_date_style = resolved_style
    plan.adct_suffix = notes_mod.detect_adct_suffix(base)
    plan.operations = parse_operations(lines, base)
    if str(amender_path).lower().endswith((".docx", ".doc")):
        try:
            plan.converted_amender = str(prepare_base(amender_path, cache))
        except ValueError:
            plan.converted_amender = ""
    for op in plan.operations:
        if op.action == "annex":
            op.annex_path = plan.converted_amender
    plan.warnings.extend(assign_effects(lines, plan.operations, base, plan.effects, publication))
    plan.related_provisions = collect_remissions(lines, plan.operations, base, citation)
    nested_effects(plan.operations, base)
    plan.operations = implicit_heading_renumbering(plan.operations, base)
    plan.operations = merge_versions(plan.operations)
    plan.operations = merge_replacements(plan.operations, base)
    for op in plan.operations:
        op.selected = True
        op.publication = publication
    if not citation:
        plan.warnings.append("Informe a identificação do ato alterador antes de gerar o documento.")
    if any(not op.effects for op in plan.operations):
        plan.warnings.append("Data de efeitos não identificada para todas as mudanças. As notas não inventam datas; informe a publicação ou a data de efeitos, se necessário.")
    if publication:
        plan.warnings.append("A data de publicação só é usada como data de efeitos quando o ato diz que entra em vigor ou produz efeitos na data de sua publicação.")
    style = note_settings(plan)["style"]
    plan.warnings.append(f"Datas e anos nas notas: {'dois dígitos (dd.mm.aa)' if style == '2' else 'quatro dígitos (dd.mm.aaaa)'}, conforme as notas mais recentes da base.")
    if any(re.search(r"consolidad[ao]\s+ate|atualizad[ao]\s+ate|alterad[ao]\s+pel", norm(b.text)) for b in base.blocks[:18]):
        plan.warnings.append("Confira ao final as referências iniciais de atualização da legislação. Essas notas são preservadas e não afirmam a completude das mudanças selecionadas.")
    if str(base_path).lower().endswith(".doc"):
        plan.warnings.append("A base .doc foi convertida para .docx. Confira no Word a fidelidade da conversão.")
    validate_plan(plan, base)
    if str(amender_path).lower().endswith((".pdf", ".html", ".htm")):
        for op in plan.operations:
            if op.status == "pronto":
                op.status = "revisar"
                op.reason = "Confira a extração de PDF/HTML com a origem e clique em Validar correção (ou em Aprovar extração do PDF/HTML para todas)."
                op.selected = False
                op.extraction_warning = op.reason
    for op in plan.operations:
        op.selected = op.status == "pronto"
    return plan, base


MODEL_STYLE = re.compile(r"a\d\d_")


def _style_source(name):
    """Estilos de um modelo: estilos_sefa.xml (modelo antigo, a10_nr_remiss ...) ou estilos_sefa_atual.xml
    (o mesmo modelo do Formatador SEFA, a30_remiss_nova_red ...)."""
    root = etree.parse(str(ROOT / "modelos" / name)).getroot()
    styles = [s for s in root if s.tag == qn("w:style") and s.find(qn("w:name")) is not None]
    return {s.find(qn("w:name")).get(qn("w:val")): s for s in styles}, {s.get(qn("w:styleId")): s for s in styles}


def ensure_styles(doc, profile):
    """Garante os estilos usados nas notas e no texto e, como no Word do Formatador, a lista completa de estilos
    do modelo SEFA da base (1.4.3): o atual, se a base usa a30_remiss_nova_red, ou o antigo, se usa a10_nr_remiss.
    Os estilos que a base já tem não são alterados."""
    current = profile.get("new_note") == "a30_remiss_nova_red"
    sources = [_style_source("estilos_sefa_atual.xml"), _style_source("estilos_sefa.xml")]
    if not current:
        sources.reverse()
    ids = {x.get(qn("w:styleId")) for x in doc.styles.element.findall(qn("w:style"))}

    def copy_style(s, by_id):
        sid = s.get(qn("w:styleId"))
        if sid in ids:
            return
        parent = s.find(qn("w:basedOn"))
        if parent is not None and parent.get(qn("w:val")) in by_id:
            copy_style(by_id[parent.get(qn("w:val"))], by_id)
        doc.styles.element.append(deepcopy(s))
        ids.add(sid)
    for name in profile.values():
        if name not in doc.styles:
            source = next((src for src in sources if name in src[0]), None)
            if source is None:
                raise ValueError("Estilo não disponível: " + name)
            copy_style(source[0][name], source[1])
    by_name, by_id = sources[0]
    for name, s in by_name.items():
        if MODEL_STYLE.match(name) and name not in doc.styles:
            copy_style(s, by_id)


def paragraph(doc, text, style, bold_label=False):
    el = OxmlElement("w:p")
    p = Paragraph(el, doc._body)
    p.style = style
    if p._p.pPr is not None:
        for name in ("ind", "spacing", "jc"):
            for prop in list(p._p.pPr.findall(qn("w:" + name))):
                p._p.pPr.remove(prop)
    info = label(text) if bold_label else None
    if info and info[0] == 0:
        end = info[2]
        p.add_run(text[:end]).bold = True
        p.add_run(text[end:])
    else:
        p.add_run(text)
    return el


def historical(el, doc, style):
    el = deepcopy(el)
    if el.tag != qn("w:p"):
        return el
    p = Paragraph(el, doc._body)
    p.style = style
    if p._p.pPr is not None:
        for name in ("ind", "spacing", "jc"):
            for prop in list(p._p.pPr.findall(qn("w:" + name))):
                p._p.pPr.remove(prop)
    # A cor do histórico vem do estilo, inclusive em hyperlinks arquivados.
    # Mantém negrito, itálico, campos e os relacionamentos da própria base.
    for rpr in el.iter(qn("w:rPr")):
        for name in ("color", "sz", "szCs", "rStyle"):
            for attr in list(rpr.findall(qn("w:" + name))):
                rpr.remove(attr)
    return el


def prior_note(text, effects, citation="", publication="", style="4"):
    return notes_mod.prior_note(text, effects, citation, publication, style)


NOTE_START = ("redacao dada", "nova redacao", "acrescido", "acrescida", "renumerado", "renumerada", "revogado", "revogada")


def note_targets(text, corpus="PRINCIPAL"):
    """Destinos de uma nota de redação, sem combinar números de dispositivos diferentes."""
    found = refparser.parse(text)
    targets = []
    for r in found:
        if refparser.is_self(r.law):
            continue
        ts = r.targets(corpus)
        if r.warning:
            return None
        targets.extend(ts)
    return targets


def note_matches(text, target):
    """Nota individual do próprio destino (uma nota coletiva fica com os demais)."""
    n = norm(text)
    if not n.startswith(NOTE_START):
        return False
    if re.search(r"\b(?:incisos|alineas|paragrafos|arts\.|artigos|itens)\b", n) or "§§" in text:
        return False
    ts = note_targets(text.split(" pel")[0] if " pel" in text else text, target.corpus)
    if not ts:
        return False
    if n.startswith(("renumerado", "renumerada")):
        return any(t.article == target.article and t.level <= target.level and t.contains(target) or t == target for t in ts)
    return len(ts) == 1 and ts[0] == target


def attached_note(text, target):
    """Nota colocada logo acima do dispositivo que remete a ele ou a um ancestral
    ("Redação dada ao art. 78 ..." acima da alínea “b”): pertence a este texto.
    Também vale quando só o nível intermediário mudou de nome (inciso I do
    parágrafo único que passou a ser inciso I do § 1º)."""
    n = norm(text)
    if not n.startswith(NOTE_START):
        return False
    ts = note_targets(text.split(" pel")[0] if " pel" in text else text, target.corpus)
    if not ts or len(ts) != 1:
        return False
    t = ts[0]
    if t.contains(target):
        return True
    level = target.level
    return (t.article == target.article and t.level == level and level >= 2
            and t.parts[level] == target.parts[level] and t.parts[1] in ("UNICO", "1") and target.parts[1] in ("UNICO", "1"))


def is_original(base, node):
    """Sem nota de origem: 'Redação original', salvo se a base já guarda outra versão deste rótulo."""
    article = base.find(Target(node.target.corpus, node.target.article))
    start, end = (article.start, article.end) if article else (node.start, node.end)
    level = node.target.level
    number = node.target.parts[level]
    for b in base.blocks[start:end]:
        if b.role == "old":
            info = label(b.text)
            if info and info[0] == level and info[1] == number:
                return False
    return True


def new_note(op, citation, style="4", corpus_suffix=True):
    return "\n".join(notes_mod.new_notes(op, citation, style, corpus_suffix))


def effects_text(effects, style="4", until=""):
    return notes_mod.effects_text(effects, style, until)


def highlight_new_elements(doc, elements, enabled):
    """Realça só os parágrafos novos; os estilos e a cor da fonte são mantidos."""
    if enabled:
        for el in elements:
            for run in Paragraph(el, doc._body).runs:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW


def _history(doc, base, node, op, citation, profile, style, scope):
    """Notas e textos que vão para o histórico de um dispositivo; e o que sai da base."""
    end = node.end if scope == "subtree" else node.head_end
    target = node.target
    attached = []
    for b in base.blocks[node.lead:node.start]:
        if b.role == "new_note" and (note_matches(b.text, target) or attached_note(b.text, target)):
            attached.append(b)
    origin, origin_targets = origin_note(base, node)
    collective = bool(origin and len(origin_targets) > 1 and target in origin_targets and origin not in attached)
    inserts = []
    archived = []
    # Revogação: "Vide" que acompanhava o texto revogado vai com ele para o histórico
    # (Constituição do Pará, EC 93/2024, inciso XIX do art. 31).
    vides = [b for b in base.blocks[node.lead:node.start]
             if b.role == "new_note" and norm(b.text).startswith("vide") and b not in attached] if op.action == "revoke" else []
    archived.extend(historical(b.element, doc, profile["old_note"]) for b in vides)
    if collective:
        remaining = [t for t in origin_targets if t != target]
        remaining_nodes = [base.find(t) for t in remaining]
        remaining_nodes = [n for n in remaining_nodes if n]
        if remaining_nodes:
            first = min(remaining_nodes, key=lambda n: n.start)
            inserts.append((base.blocks[first.start].element, paragraph(doc, note_for_targets(origin.text, remaining), profile["new_note"])))
            attached.append(origin)
    if attached:
        for b in attached:
            text = note_for_targets(b.text, [target]) if collective and b is origin else b.text
            archived.append(replace_text(historical(b.element, doc, profile["old_note"]), prior_note(text, op.effects, citation, op.publication, style)))
    elif origin:
        text = note_for_targets(origin.text, [target])
        archived.append(paragraph(doc, prior_note(text, op.effects, citation, op.publication, style), profile["old_note"]))
    else:
        archived.append(paragraph(doc, notes_mod.fallback_note(is_original(base, node), op.effects, style), profile["old_note"]))
    for b in base.blocks[node.start:end]:
        if b.role == "current":
            archived.append(historical(b.element, doc, profile["old"]))
        elif b.role == "new_note":
            archived.append(replace_text(historical(b.element, doc, profile["old_note"]), prior_note(b.text, op.effects, citation, op.publication, style)))
        else:
            archived.append(deepcopy(b.element))
    removed = [b.element for b in base.blocks[node.start:end]] + [b.element for b in attached] + [b.element for b in vides]
    return archived, removed, inserts


def shared_history(base, nodes, scope):
    """Irmãos consecutivos que podem dividir uma só nota de histórico, como em
    "Redação dada às alíneas “b” e “c” do art. 41 pela EC 15/1999, efeitos de ..."
    ou "Revogados os incisos IV e V ... / Redação original, efeitos até ...".
    Devolve (nota de origem, destinos da nota) ou None."""
    if len(nodes) < 2:
        return None
    origins = [origin_note(base, n) for n in nodes]
    if len({o[0].index if o[0] is not None else None for o in origins}) != 1:
        return None
    origin, origin_targets = origins[0]
    for k, n in enumerate(nodes):
        notes_above = [b for b in base.blocks[n.lead:n.start] if b.role == "new_note"]
        if any(b is not origin for b in notes_above) or (k and notes_above):
            return None
        if k and nodes[k - 1].end != n.lead:
            return None
        inner = base.blocks[n.head_end:n.end] if scope == "head" else base.blocks[n.start:n.end]
        if scope == "head" and any(b.text for b in inner):
            return None
        if any(b.role in ("old", "old_note") for b in inner):
            return None
    return origin, origin_targets


def _group_history(doc, base, nodes, op, citation, profile, style, scope, shared):
    """Uma nota de histórico para o grupo, seguida das redações anteriores de cada um."""
    origin, origin_targets = shared
    targets = [n.target for n in nodes]
    archived, removed, inserts = [], [], []
    if origin is None:
        original = all(is_original(base, n) for n in nodes)
        archived.append(paragraph(doc, notes_mod.fallback_note(original, op.effects, style), profile["old_note"]))
    elif len(origin_targets) > 1 and all(t in origin_targets for t in targets):
        # Nota coletiva de irmãos: o grupo leva a sua parte; o restante continua anotado.
        remaining = [t for t in origin_targets if t not in targets]
        remaining_nodes = [n for n in (base.find(t) for t in remaining) if n]
        if remaining_nodes:
            first = min(remaining_nodes, key=lambda n: n.start)
            inserts.append((base.blocks[first.start].element, paragraph(doc, note_for_targets(origin.text, remaining), profile["new_note"])))
        text = note_for_targets(origin.text, targets) if remaining else origin.text
        archived.append(replace_text(historical(origin.element, doc, profile["old_note"]), prior_note(text, op.effects, citation, op.publication, style)))
        removed.append(origin.element)
    else:
        # Nota de um dispositivo superior ("Redação dada ao art. 41"): continua no lugar.
        archived.append(paragraph(doc, prior_note(note_for_targets(origin.text, targets), op.effects, citation, op.publication, style), profile["old_note"]))
    for n in nodes:
        end = n.end if scope == "subtree" else n.head_end
        for b in base.blocks[n.start:end]:
            if b.role == "current":
                archived.append(historical(b.element, doc, profile["old"]))
            elif b.role == "new_note":
                archived.append(replace_text(historical(b.element, doc, profile["old_note"]), prior_note(b.text, op.effects, citation, op.publication, style)))
            else:
                archived.append(deepcopy(b.element))
        removed.extend(b.element for b in base.blocks[n.start:end])
    return archived, removed, inserts


def implicit_heading_renumbering(operations, base):
    """Seção II acrescida a capítulo que tinha "Seção única": "Renumerada a Seção única
    para Seção I" (LC 123, LC 147/2014). Vale para capítulo, subseção etc."""
    out = []
    for op in operations:
        info = headmod.parse_line(op.heading[0]) if op.action == "add" and op.heading else None
        if info and info[1] not in ("ÚNICA", "I"):
            kind = info[0]
            anchor, reason = insertion_anchor(base, op.target)
            parent = headmod.parent_at(base, anchor, kind) if not reason else None
            single = [g for g in headmod.groups(base) if g.kind == kind and g.number == "ÚNICA" and g.corpus == op.target.corpus
                      and (parent in g.parents if parent else not g.parents)]
            if len(single) == 1:
                g = single[0]
                label_line = re.sub(r"(?i)\b[uú]nic[ao]\b", "I", g.lines[0])
                rename = Operation("", "heading", Target(op.target.corpus, g.anchor.article if g.anchor else ""), new_lines=[label_line],
                                   source=op.source, source_location=op.source_location, source_article=op.source_article,
                                   note="Renumeração decorrente do acréscimo de outro título do mesmo nível.")
                rename.heading_key = f"{kind}:ÚNICA"
                rename.heading_parents = "/".join(f"{k}:{n}" for k, n in g.parents)
                rename.heading_rename = True
                rename.effects = op.effects
                out.append(rename)
        out.append(op)
    return out


def merge_versions(operations):
    """Duas ou mais propostas do mesmo ato para o mesmo dispositivo, com datas de
    efeitos diferentes, viram uma só mudança com redações por período."""
    out, by_key = [], {}
    for op in operations:
        key = (op.target, op.scope, tuple(op.extra_targets))
        first = by_key.get(key)
        if (first is not None and op.action in ("add", "replace") and first.action in ("add", "replace")
                and not op.foreign and not first.foreign and not op.manual and not first.manual
                and not op.extraction_warning and not first.extraction_warning
                and op.effects and first.effects and op.effects != first.effects and not op.later_versions
                and op.via == first.via
                and not op.renumber_from and not op.heading):
            versions = [dict(effects=first.effects, effects_until=first.effects_until, lines=first.new_lines, source=first.source)]
            versions += first.later_versions
            versions.append(dict(effects=op.effects, effects_until=op.effects_until, lines=op.new_lines, source=op.source))
            versions.sort(key=lambda v: v["effects"])
            if versions[0]["effects"] != first.effects:
                # A redação mais antiga é a que fecha o histórico atual; o acréscimo, se houver, vem primeiro.
                if first.action == "add":
                    out.append(op)
                    continue
            first.effects, first.effects_until, first.new_lines = versions[0]["effects"], versions[0]["effects_until"], versions[0]["lines"]
            first.later_versions = [dict(effects=v["effects"], effects_until=v["effects_until"], lines=v["lines"]) for v in versions[1:]]
            if op.source not in first.source:
                first.source = first.source + "\n" + op.source
            first.effects_source = "\n".join(x for x in (first.effects_source, op.effects_source) if x)
            first.remissions = list(dict.fromkeys(first.remissions + op.remissions))
            continue
        if op.action in ("add", "replace"):
            by_key.setdefault(key, op)
        out.append(op)
    for i, op in enumerate(out, 1):
        op.id = f"A{i:03d}"
    return out


def _via_dates(base, blocks, via):
    """Datas 'efeitos a partir de ...' das notas que citam o ato alterado (LC 214/25)."""
    number = re.search(r"\s([\d.]+)/", " " + via)
    if not number:
        return []
    pattern = re.compile(r"\b" + re.escape(via.split()[0]) + r"\s*" + re.escape(number[1]) + r"/\d{2,4}\b", re.I)
    out = []
    for b in blocks:
        if b.role in ("new_note",) and pattern.search(b.text):
            m = re.search(r"efeitos\s+a\s+partir\s+de\s+(\d{1,2}[./]\d{1,2}[./]\d{2,4})", b.text, re.I)
            if m:
                try:
                    out.append(date_iso(m[1]))
                except ValueError:
                    pass
    return out


def nested_effects(operations, base):
    """Mudança feita na redação de outro ato alterador: vale a data desse ato quando a
    redação dele ainda não vigorava (LC 123: "Redação dada ao § 1º do art. 23 pela LC 214/25,
    alterada pela LC 227/26, efeitos a partir de 01.01.27")."""
    for op in operations:
        if not op.via or not op.effects:
            continue
        node = base.find(source_target(op))
        own = []
        if node:
            origin, _ = origin_note(base, node)
            notes = list(base.blocks[node.lead:node.start]) + ([origin] if origin is not None else [])
            own = _via_dates(base, notes, op.via)
        if own:
            if max(own) > op.effects:
                op.effects = max(own)
                op.note = (op.note + " " if op.note else "") + (
                    f"Efeitos da {op.via} para este dispositivo: a redação dada por ela ainda não vigorava quando este ato foi publicado.")
            continue
        parent = base.find(Target(op.target.corpus, op.target.article))
        around = _via_dates(base, base.blocks[parent.lead:parent.end], op.via) if parent else []
        later = sorted(d for d in set(around) if d > op.effects)
        if later:
            shown = ", ".join("/".join(reversed(d.split("-"))) for d in later)
            op.extraction_warning = (op.extraction_warning + " " if op.extraction_warning else "") + (
                f"Confirme os efeitos: a {op.via} prevê {shown} para outros dispositivos deste artigo; "
                f"a proposta usa a data deste ato. Ajuste em Efeitos desta mudança, se for o caso, e valide.")


def merge_replacements(operations, base):
    """Novas redações de irmãos consecutivos, do mesmo artigo do ato e com os
    mesmos efeitos, numa só nota ("Redação dada aos arts. 168 e 169")."""
    from .parser import _next_number
    out = []
    for op in operations:
        prev = out[-1] if out else None
        if (prev and prev.action == op.action == "replace" and not prev.manual and not op.manual and prev.via == op.via
                and not prev.extraction_warning and not op.extraction_warning and not prev.foreign and not op.foreign
                and prev.source_article == op.source_article and prev.scope == op.scope == "head"
                and not any(x.renumber_from or x.fold or x.caput_forced or x.full_text or x.later_versions for x in (prev, op))
                and prev.effects == op.effects and prev.effects_until == op.effects_until
                and prev.targets[-1].level == op.target.level and prev.targets[-1].parent() == op.target.parent()
                and _next_number(prev.targets[-1], op.target)):
            nodes = [base.find(t) for t in prev.targets + [op.target]]
            if all(nodes) and shared_history(base, nodes, "head") is not None:
                prev.extra_targets = prev.extra_targets + [op.target]
                prev.new_lines = prev.new_lines + op.new_lines
                if op.source not in prev.source:
                    prev.source = prev.source + "\n" + op.source
                prev.remissions = list(dict.fromkeys(prev.remissions + op.remissions))
                prev.related_source = list(dict.fromkeys(prev.related_source + op.related_source))
                continue
        out.append(op)
    for i, op in enumerate(out, 1):
        op.id = f"A{i:03d}"
    return out


def renamed_label(level, number):
    """Rótulo do dispositivo renumerado: '§ 3º ', '§ 10. ', 'VIII - ', 'd) ', '2. '."""
    if level == 1:
        head = number.split("-")[0]
        return f"§ {number.replace(head, head + 'º', 1) if head.isdigit() and int(head) <= 9 else number + '.'} "
    if level == 2:
        return f"{number} - "
    if level == 3:
        return f"{number}) "
    return f"{number}. "


def revoked_line(node, suppressed=False):
    word = "SUPRIMID" if suppressed else "REVOGAD"
    return node.label_text.rstrip() + " " + word + ("A" if node.target.level == 3 else "O")


def apply_operation(doc, op, citation, profile, highlight_changes=False, style="4", corpus_suffix=True):
    base = Legislation(doc)
    original_before = list(op.before)
    original_location = op.base_location
    validate_operation(op, base)
    if op.status != "pronto":
        raise ValueError(f"{op.id}: {op.reason}")
    if original_location:
        op.before = original_before
        op.base_location = original_location
    if op.action == "heading":
        return _apply_heading(doc, base, op, citation, profile, highlight_changes, style)
    if op.action == "annex":
        return _apply_annex(doc, base, op, citation, profile, highlight_changes, style)
    if op.action == "note":
        node = base.find(op.target)
        fresh = [paragraph(doc, text, profile["new_note"]) for text in op.new_lines]
        highlight_new_elements(doc, fresh, highlight_changes)
        anchor = base.blocks[node.start].element
        for el in fresh:
            anchor.addprevious(el)
        op.applied = True
        return
    node = base.find(source_target(op))
    body = doc.element.body
    cite = notes_mod.format_citation(citation, style)
    staged = []
    if op.later_versions:
        fresh = []
        for k, (notes, lines) in enumerate(notes_mod.version_blocks(op, citation, style, corpus_suffix)):
            fresh.extend(paragraph(doc, text, profile["new_note"]) for text in notes)
            if k == 0:
                fresh.extend(paragraph(doc, text.replace(citation, cite), profile["new_note"]) for text in op.remissions)
            block = [paragraph(doc, s, profile["current"], True) for s in lines]
            staged.extend(block)
            fresh.extend(block)
    else:
        fresh = [paragraph(doc, text, profile["new_note"]) for text in notes_mod.new_notes(op, citation, style, corpus_suffix)]
        if op.action in ("replace", "add"):
            fresh.extend(paragraph(doc, text.replace(citation, cite), profile["new_note"]) for text in op.remissions)
    if op.action == "add":
        anchor, reason = insertion_anchor(base, op.target)
        if reason:
            raise ValueError(reason)
        if op.heading:
            hstyle = heading_style(base)
            if hstyle not in doc.styles:
                hstyle = profile["current"]
            fresh.extend(paragraph(doc, s, hstyle) for s in op.heading)
        if not op.later_versions:
            fresh.extend(paragraph(doc, s, profile["current"], True) for s in op.new_lines)
        highlight_new_elements(doc, fresh, highlight_changes)
        for offset, el in enumerate(fresh):
            body.insert(anchor + offset, el)
        op.applied = True
        return
    nodes = [node] if op.action not in ("revoke", "replace") or len(op.targets) == 1 else sorted((base.find(t) for t in op.targets), key=lambda n: n.start)
    if op.action == "revoke":
        fresh.extend(paragraph(doc, revoked_line(n, op.suppressed), profile["current"], True) for n in nodes)
    elif op.action == "rename":
        info = label(node.text)
        fresh.append(paragraph(doc, renamed_label(op.target.level, op.rename_to) + node.text[info[2]:].lstrip(), profile["current"], True))
    elif not op.later_versions:
        fresh.extend(paragraph(doc, s, profile["current"], True) for s in op.new_lines)
    highlight_new_elements(doc, fresh, highlight_changes)
    scope = "subtree" if op.action == "revoke" else op.scope
    archived, removed, inserts = [], [], []
    shared = shared_history(base, nodes, scope) if len(nodes) > 1 else None
    if shared is not None:
        archived, removed, inserts = _group_history(doc, base, nodes, op, citation, profile, style, scope, shared)
    else:
        for n in nodes:
            a, r, ins = _history(doc, base, n, op, citation, profile, style, scope)
            archived.extend(a)
            removed.extend(r)
            inserts.extend(ins)
    anchor = nodes[0].element if hasattr(nodes[0], "element") else base.blocks[nodes[0].start].element
    for el in fresh + archived:
        anchor.addprevious(el)
    for where, el in inserts:
        where.addprevious(el)
    for el in removed:
        if el.getparent() is not None:
            body.remove(el)
    op.applied = True


def protect_outputs(plan, paths):
    protected = {Path(p).resolve() for p in (plan.base_path, plan.amender_path, plan.converted_base)}
    if any(Path(p).resolve() in protected for p in paths):
        raise ValueError("Escolha outro nome para a saída. Os arquivos de entrada são preservados.")


def generate(plan, output_path, highlight_changes=False):
    if not plan.citation.strip():
        raise ValueError("Informe a identificação do ato alterador, como Lei 12.345/2026.")
    if digest(plan.base_path) != plan.base_hash or digest(plan.amender_path) != plan.amender_hash:
        raise ValueError("Um dos arquivos mudou depois da análise. Analise novamente antes de gerar.")
    out = Path(output_path).resolve()
    protect_outputs(plan, [out])
    if out.suffix.lower() != ".docx":
        raise ValueError("A saída deve ter a extensão .docx.")
    selected = [op for op in plan.operations if op.selected]
    if not selected:
        raise ValueError("Selecione pelo menos uma mudança validada.")
    if any(op.status != "pronto" for op in selected):
        raise ValueError("Há mudanças selecionadas que ainda precisam de revisão.")
    # Revalida conflitos sem alterar o plano que a interface apresenta.
    check = Plan.from_dict(plan.to_dict())
    validate_plan(check)
    if any(op.selected and op.status != "pronto" for op in check.operations) or any(check.operations[i].status != "pronto" for i, op in enumerate(plan.operations) if op.selected):
        raise ValueError("As mudanças selecionadas têm conflito estrutural. Revise o plano.")
    doc = Document(plan.converted_base)
    original_links = links_signature(doc)
    profile = Legislation(doc).profile
    ensure_styles(doc, profile)
    settings = note_settings(plan)
    originals = {op.id: op.to_dict() for op in selected}
    try:
        # Na ordem do ato; a mudança que depende de outra posterior (o § 2º novo depois
        # da renumeração do parágrafo único) é aplicada na rodada seguinte.
        pending = list(selected)
        while pending:
            later, errors = [], []
            for op in pending:
                saved = (list(op.before), op.base_location)
                try:
                    apply_operation(doc, op, plan.citation, profile, highlight_changes=highlight_changes, **settings)
                    op.selected = True
                except ValueError as error:
                    if not any(s in str(error) for s in DEPENDENT):
                        raise
                    op.before, op.base_location = saved
                    later.append(op)
                    errors.append(error)
            if len(later) == len(pending):
                raise errors[0]
            pending = later
        update_initial_reference(doc, notes_mod.format_citation(plan.citation, settings["style"]))
        if links_signature(doc) != original_links:
            raise ValueError("A geração alteraria os hyperlinks originais. O Word não foi salvo; revise os trechos com links.")
        # Continua sendo obrigatório abrir como DOCX válido antes da publicação.
        out.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(suffix=".docx", dir=out.parent, delete=False) as temp:
            temporary = Path(temp.name)
        try:
            save_preserving_package(doc, plan.converted_base, temporary)
            Document(temporary)
            if links_signature(Document(temporary)) != original_links:
                raise ValueError("A validação final detectou alteração nos hyperlinks. O Word não foi salvo.")
            os.replace(temporary, out)
        finally:
            temporary.unlink(missing_ok=True)
    except Exception:
        for op in selected:
            restored = Operation.from_dict(originals[op.id])
            op.__dict__.update(restored.__dict__)
        raise
    return out


def save_plan(plan, path):
    protect_outputs(plan, [path])
    if Path(path).suffix.lower() != ".json":
        raise ValueError("Salve o plano com a extensão .json.")
    value = {"format": "consolidador-sefa", "version": 1, "created": datetime.now().isoformat(timespec="seconds"), "plan": plan.to_dict()}
    Path(path).write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")


def load_plan(path, cache):
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    if data.get("format") != "consolidador-sefa" or data.get("version") != 1:
        raise ValueError("Esse arquivo não é um plano desta versão do Consolidador.")
    plan = Plan.from_dict(data["plan"])
    if digest(plan.base_path) != plan.base_hash or digest(plan.amender_path) != plan.amender_hash:
        raise ValueError("Os arquivos deste plano foram alterados. Faça uma nova análise.")
    plan.converted_base = str(prepare_base(plan.base_path, cache))
    base = Legislation(plan.converted_base)
    validate_plan(plan, base)
    return plan, base
