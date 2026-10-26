"""Datas e remissões extraídas de cláusulas expressas, sem inferir vigência tácita."""
from __future__ import annotations
from dataclasses import replace
from datetime import date, timedelta
import re
from .index import ARTICLE, label
from .model import Target, date_iso, norm, of_act
from .parser import VERBS, END_QUOTE, reference_lists, foreign_target, line_location
from . import refs as refparser

MONTHS={name:i for i,name in enumerate(("janeiro","fevereiro","marco","abril","maio","junho",
                                      "julho","agosto","setembro","outubro","novembro","dezembro"),1)}
DATES=re.compile(r"\b(\d{1,2})[º°o]?\s+de\s+([a-zÀ-ÿ]+)\s+de\s+(\d{4})\b|\b(\d{1,2}[./]\d{1,2}[./]\d{2,4})\b",re.I)
EFFECTS=re.compile(r"produz\w*\s+efeitos|produzindo\s+efeitos|efeitos\s+a\s+partir|efeitos\s+retroativos|retroag\w+(?:\s+seus)?\s+efeitos|surtir[aá]?\s+efeitos|surtindo\s+efeitos",re.I)
VIGOR=re.compile(r"entra(?:r[aá])?\s+em\s+vigor|passa(?:r[aá])?\s+a\s+vigorar\s+(?:na|a\s+partir|em)\b",re.I)
NUMBERS={"um":1,"uma":1,"dois":2,"duas":2,"tres":3,"quatro":4,"cinco":5,"seis":6,"sete":7,"oito":8,"nove":9,"dez":10,
         "quinze":15,"trinta":30,"quarenta e cinco":45,"sessenta":60,"noventa":90,"cento e vinte":120,"cento e oitenta":180}
PUBLICATION=re.compile(r"\b(?:d?a\s+)?data\s+(?:de\s+|da\s+)?(?:sua\s+)?publica[cç][aã]o\b|\bpublica[cç][aã]o\s+dest[ae]\s+\w+",re.I)
AFTER_DAYS=re.compile(r"(\d+|[a-zç ]+?)\s*(?:\([^)]*\)\s*)?dias\s+(?:ap[oó]s|depois\s+d[ae]|contados\s+d[ae]|a\s+partir\s+d[ae])\s+(?:a\s+|da\s+)?(?:data\s+d[ae]\s+)?(?:sua\s+)?publica[cç][aã]o",re.I)
FIRST_DAY=re.compile(r"primeiro\s+dia\s+do\s+(?:(segundo|terceiro|quarto)\s+)?m[eê]s\s+(?:subsequente|seguinte)\s+ao\s+d[ae]\s+(?:sua\s+)?publica[cç][aã]o",re.I)
NEXT_YEAR=re.compile(r"1[º°o]?\s+de\s+janeiro\s+do\s+(?:ano|exerc[ií]cio)(?:\s+financeiro)?\s+(?:subsequente|seguinte)\s+ao\s+d[ae]\s+(?:sua\s+)?publica[cç][aã]o",re.I)
SIGNATURE=re.compile(r"^(?:pal[aá]cio|bras[ií]lia|bel[eé]m|gabinete|sala\s+das\s+sess[oõ]es)\b|^[A-ZÁÉÍÓÚÂÊÔÃÕÇ .'-]{3,},\s*(?:em\s+)?\d{1,2}[º°]?\s+de\s+\w+\s+de\s+\d{4}",re.I)
SELF=re.compile(r"\b(?:dest[ae]|nest[ae])\s+(?:Lei|Decreto|Emenda|Instru[cç][aã]o|Portaria|Resolu[cç][aã]o|Ato)\b",re.I)
# END_QUOTE vem do parser: a citação também termina em "(AC)" e "(N.R.)", não só em "(NR)".
# Ligação entre dispositivos de uma mesma enumeração: "o inciso V do art. 2º e às alíneas ...".
COORDINATED=re.compile(r"^\s*(?:,\s*)?(?:(?:e|ou)\s+)?(?:(?:a|à|ao|aos|às|as|o|os|no|na|nos|nas)\s+)?$",re.I)


def dates_in(text,publication=""):
    result=[]
    for m in DATES.finditer(text):
        try:
            value=date_iso(m[4]) if m[4] else date(int(m[3]),MONTHS[norm(m[2])],int(m[1])).isoformat()
        except (ValueError,KeyError):
            continue
        result.append((m.start(),m.end(),value))
    if publication:
        # Prazos contados da publicação ("na data de sua publicação", "90 dias após").
        pub=date.fromisoformat(publication);used=[]
        for m in AFTER_DAYS.finditer(text):
            raw=norm(m[1]).strip()
            n=int(raw) if raw.isdigit() else NUMBERS.get(raw)
            if n:
                result.append((m.start(),m.end(),(pub+timedelta(days=n)).isoformat()));used.append((m.start(),m.end()))
        for m in FIRST_DAY.finditer(text):
            months=1+{"segundo":1,"terceiro":2,"quarto":3}.get(norm(m[1] or ""),0)
            y,mo=pub.year,pub.month+months
            y+=(mo-1)//12;mo=(mo-1)%12+1
            result.append((m.start(),m.end(),date(y,mo,1).isoformat()));used.append((m.start(),m.end()))
        for m in NEXT_YEAR.finditer(text):
            result.append((m.start(),m.end(),date(pub.year+1,1,1).isoformat()));used.append((m.start(),m.end()))
        for m in PUBLICATION.finditer(text):
            if any(a<=m.start()<b for a,b in used) or re.search(r"(?:at[eé]|antes\s+d[ae]|anteriormente\s+[aà])\s*$",text[:m.start()],re.I):
                continue
            result.append((m.start(),m.end(),publication))
    return sorted(result)


def publication_date(lines):
    for line in lines[:30]:
        if re.search(r"publicad[ao].*(?:dou|doe|diario|d\.o\.)",norm(line.text)):
            values=dates_in(line.text)
            if values:return values[0][2]
    return ""


def source_articles(lines):
    """Separa artigos do alterador sem confundi-los com os artigos citados entre aspas."""
    groups=[]; current=[]; quoted=False
    for line in lines:
        text=line.text.strip()
        starts_quote=text.startswith(('“','"')) and label(text)
        if ARTICLE.match(text) and not quoted and not starts_quote:
            if current:groups.append(current)
            current=[line]
        elif current:
            current.append(line)
        if starts_quote:quoted=True
        if quoted and END_QUOTE.search(text):quoted=False
    if current:groups.append(current)
    return groups


def referenced_targets(text):
    result=[]
    for r in refparser.parse(text):
        if refparser.is_self(r.law):
            continue
        ts=r.targets()
        if r.warning:
            return []
        result.extend(t for t in ts if t not in result)
    return result


def list_laws(text,parsed):
    """A lei citada no fim de uma enumeração vale para todos os dispositivos enumerados antes dela:
    "quanto ao inciso V do art. 2º e às alíneas “b” e “d” do inciso I do art. 5º da Lei nº 99.100"
    (o inciso V também é da Lei 99.100, e não do art. 2º do próprio ato)."""
    laws=[r.law for r in parsed]
    for k in range(len(parsed)-2,-1,-1):
        if not laws[k] and laws[k+1] and COORDINATED.match(text[parsed[k].end:parsed[k+1].start]):
            laws[k]=laws[k+1]
    return laws


def self_references(text):
    """Artigos (e incisos) do próprio ato citados como 'art. 2º desta Lei'."""
    out=[]
    for r in refparser.parse(text):
        if refparser.is_self(r.law) or (not r.law and SELF.search(text[r.end:r.end+40])):
            for a in r.parts.get("article",[]):
                for i in r.parts.get("inciso",[""]) or [""]:
                    out.append((a,i))
    return out


def assign_effects(lines,operations,base,fallback="",publication=""):
    rules={"effects":dict(defaults=[],specific=[],self=[]),"vigor":dict(defaults=[],specific=[],self=[])}
    warnings=[]; ambiguous=[]; inline_rules={}; unresolved=""
    for group in source_articles(lines):
        header=[]
        items=[]
        for i,line in enumerate(group):
            if SIGNATURE.match(line.text.strip()):
                break
            info=label(line.text)
            if i and info:
                items.append(line)
                continue
            if items:
                continue
            header.append(line.text)
        header_text=' '.join(header)
        if EFFECTS.search(header_text):
            kind,marker_re="effects",EFFECTS
        elif VIGOR.search(header_text) and not re.search(r"seguintes?\s+(?:altera|reda)",norm(header_text)):
            kind,marker_re="vigor",VIGOR
        else:
            continue
        bucket=rules[kind]
        inline=bool(VERBS.search(header_text)) and kind=="effects"
        inline_ops=[o for o in operations if group[0].text in o.source] if inline else []
        candidates=[replace(group[0],text=header_text)]+([] if inline else items)
        if not publication and PUBLICATION.search(" ".join(l.text for l in group)):
            unresolved=" ".join(l.text for l in group)
        for line in candidates:
            text=line.text; values=dates_in(text,publication)
            if not values:continue
            scope=re.search(r"\bquanto\b|\brelativamente\b|\bno que se refere\b|\bem rela[cç][aã]o\b",text,re.I)
            scope_text=None
            if scope:
                before=[v for v in values if v[0]<scope.start()]
                if before:
                    values=before
                else:
                    # "I - quanto ao art. 1º, a partir de 1º de janeiro de 2026"
                    after=[v for v in values if v[0]>scope.start()]
                    if after:
                        scope_text=text[scope.start():after[0][0]]
                        values=after[:1] if len({v[2] for v in after})==1 else after
            if not values:continue
            # "entra em vigor em X, produzindo efeitos a partir de Y": usa Y.
            marker=marker_re.search(text)
            if marker:
                values=[v for v in values if v[0]>=marker.end()]
            if not values:continue
            if len({v[2] for v in values})!=1:
                warnings.append("Cláusula com mais de uma data: confira os efeitos na revisão. "+text)
                tail=text[scope.start():] if scope else text[values[-1][1]:]
                ambiguous.append([o.target for o in inline_ops] if inline else referenced_targets(tail))
                continue
            value=values[0][2]
            if inline:
                for op in inline_ops:inline_rules.setdefault(op.id,[]).append((value,text))
                continue
            tail=scope_text if scope_text is not None else text[values[0][1]:]
            parsed=refparser.parse(tail)
            laws=list_laws(tail,parsed)
            targets=[]
            for r,law in zip(parsed,laws):
                if refparser.is_self(law):
                    continue
                if r.warning:
                    targets=[];break
                targets.extend(t for t in r.targets() if t not in targets)
            own=self_references(tail)+[(a,i) for r,law in zip(parsed,laws) if refparser.is_self(law) and not refparser.is_self(r.law)
                                       for a in r.parts.get("article",[]) for i in (r.parts.get("inciso") or [""])]
            # Sem indicação da lei, "o art. 1º" de uma cláusula de efeitos é do próprio ato.
            amender_articles={o.source_article for o in operations if o.source_article}
            plain=[r for r,law in zip(parsed,laws) if not law and r.parts.get("article")]
            if plain and all(a in amender_articles for r in plain for a in r.parts["article"]):
                own=own+[(a,i) for r in plain for a in r.parts["article"] for i in (r.parts.get("inciso") or [""])]
                targets=[t for t in targets if t.article not in amender_articles]
            foreign,_=foreign_target(tail,base.identity,base.identity_year)
            if targets and not foreign:
                bucket["specific"].append((targets,value,text))
            elif own:
                bucket["self"].append((own,value,text))
            elif re.search(r"demais\s+dispositivos|demais\s+(?:alteracoes|disposicoes|artigos)|restantes",norm(tail)):
                bucket["defaults"].append((value,text))
            elif (marker or kind=="vigor") and not targets and not re.search(r"quanto|relativ|exceto|ressalvad|em relacao",norm(tail)):
                bucket["defaults"].append((value,text))
    effects_found=any(rules["effects"].values())
    active=rules["effects"] if effects_found else rules["vigor"]
    defaults,specific,own_rules=active["defaults"],active["specific"],active["self"]
    if effects_found and rules["vigor"]["defaults"]:
        # "Entra em vigor" e "produz efeitos" em artigos distintos: valem os efeitos.
        pass
    default_values={v for v,_ in defaults}
    if len(default_values)>1:
        warnings.append("Há datas gerais de efeitos divergentes no ato; confirme cada proposta.")
    for op in operations:
        targets_of_op=op.targets
        if any(not targets or any(t.contains(x) for t in targets for x in targets_of_op) for targets in ambiguous):
            op.extraction_warning+=" A cláusula apresenta mais de uma data possível para este destino; confirme os efeitos."
        matched=inline_rules.get(op.id,[])+[(v,s) for ts,v,s in specific if any(t.contains(x) for t in ts for x in targets_of_op)]
        matched+=[(v,s) for pairs,v,s in own_rules if any(a==op.source_article and (not i or i==op.source_item) for a,i in pairs)]
        values={v for v,_ in matched}
        if len(values)>1:
            op.extraction_warning+=" Há regras expressas de efeitos divergentes para este destino."
            continue
        chosen=matched if matched else defaults if len(default_values)==1 else []
        if not matched and len(default_values)>1:
            op.extraction_warning+=" Há datas gerais de efeitos divergentes; confirme os efeitos desta mudança."
        op.effects=chosen[0][0] if chosen else fallback
        op.effects_source="\n".join(dict.fromkeys(s for _,s in chosen))
    if fallback and default_values and fallback not in default_values:
        warnings.append("A data geral informada difere da cláusula do ato. As propostas usam as datas expressas identificadas; confira a data de cada mudança.")
    if unresolved and any(not op.effects for op in operations):
        warnings.append("O ato vincula a vigência ou os efeitos à data de publicação. Informe a data de publicação para que as notas recebam a data de efeitos.")
    return warnings


def _source_number(group):
    return label(group[0].text)[1]


def collect_remissions(lines,operations,base,citation):
    """Preserva regras autônomas como remissões, nunca como novos artigos da base."""
    groups=source_articles(lines); related=[]; limit_conditions=[]
    for group in groups:
        text="\n".join(l.text for l in group)
        first=group[0].text
        if VERBS.search(first) or EFFECTS.search(first) or "entra em vigor" in norm(first):continue
        number=_source_number(group)
        if "ampliacao dos limites" in norm(text):
            limit_conditions.append((number,ARTICLE.sub("",first,count=1).strip(),text))
        context=[]
        for line in group:
            foreign,_=foreign_target(line.text,base.identity,base.identity_year)
            info=label(line.text)
            reference=line.text[info[2]:] if info else line.text
            ts=referenced_targets(reference)
            if foreign:continue
            for t in ts:
                changed=any(x==t or x.contains(t) for op in operations for x in op.targets if op.action in ("add","replace"))
                if base.find(t) or changed:context.append((t,line))
        if context:
            related.append({"article":number,"source":text,"location":line_location(group),
                            "targets":[t.__dict__ for t,_ in context]})
        for target,line in context:
            # Regras transitórias quantitativas: redação fiel, sem trocar os
            # valores do texto permanente que só produzirá efeitos depois.
            if not (re.search(r"ano-calend[aá]rio|exerc[ií]cio\s+de",line.text,re.I)
                    and re.search(r"\blimites?\b|R\$",line.text)):
                continue
            info=label(line.text)
            body=line.text[info[2]:].strip() if info else line.text.strip()
            sole=bool(re.match(r"Par[aá]grafo\s+[uú]nico",line.text,re.I))
            note=f"Vide {'parágrafo único do ' if sole else ''}art. {number}º {of_act(citation)} {citation}: {body}"
            for op in operations:
                if op.target==target:
                    op.remissions.append(note);op.related_source.append(text)
    for op in operations:
        if any(re.search(r"\blimites?\b",s,re.I) for s in op.remissions):
            for number,body,source in limit_conditions:
                op.remissions.append(f"Vide art. {number}º {of_act(citation)} {citation}: {body}")
                op.related_source.append(source)
    # Uma remissão de créditos costuma ser disciplinada por artigos que se
    # remetem entre si. Expande só essa rede expressa de referências.
    remission_groups=[g for g in groups if re.search(r"\bremiss[aã]o\b|cr[eé]ditos\s+remitidos",' '.join(l.text for l in g),re.I)]
    if remission_groups:
        numbers={_source_number(g) for g in remission_groups}
        for group in groups:
            if "disciplin" in norm(group[0].text):
                refs,_=reference_lists(ARTICLE.sub("",group[0].text,count=1))
                if set(refs.get("article",[])) & numbers:
                    remission_groups.append(group);numbers.add(_source_number(group))
        targets=[]
        for related_group in related:
            if related_group['article'] in numbers:
                targets.extend(Target.from_dict(t) for t in related_group['targets'])
        refs=refparser.enumerate_values(sorted(numbers,key=lambda n:int(re.match(r"\d+",n)[0])),0) if numbers else ""
        note=f"Vide arts. {refs} {of_act(citation)} {citation}, quanto à remissão de créditos e à sua regulamentação."
        for op in operations:
            if any(op.target.contains(t) for t in targets):
                op.remissions.append(note)
                op.related_source.extend('\n'.join(l.text for l in g) for g in remission_groups)
    # Demais disposições autônomas que tratam expressamente de um dispositivo
    # alterado por este ato: "Vide art. 8º da EC 136/25." (padrão das bases SEFA).
    for group in related:
        number=group['article'];ords=refparser.ordinal(number)
        targets=[Target.from_dict(t) for t in group['targets']]
        for op in operations:
            if op.action not in ("add","replace"):continue
            if not any(x==t or x.contains(t) for x in op.targets for t in targets):continue
            if any(re.search(r"\bVide\b.*\b"+re.escape(ords)+r"(?!\d)",r) for r in op.remissions):continue
            op.remissions.append(f"Vide art. {ords} {of_act(citation)} {citation}.")
            op.related_source.append(group['source'])
    for op in operations:
        op.remissions=list(dict.fromkeys(op.remissions));op.related_source=list(dict.fromkeys(op.related_source))
    return related
