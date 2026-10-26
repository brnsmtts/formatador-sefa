"""Conservação de notas, links e componentes do documento base."""
from collections import Counter
from copy import deepcopy
from io import BytesIO
from zipfile import ZipFile
import re
from lxml import etree
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from .model import norm
from dataclasses import replace
from .temporal import referenced_targets


def origin_note(base,node):
    candidates=[]
    # Só notas do mesmo corpo (texto principal, ADCT ou anexo) podem dar origem ao texto.
    corpus_start=min((n.lead for n in base.nodes if n.target.corpus==node.target.corpus),default=0)
    for b in base.blocks[corpus_start:node.start]:
        # A renumeração não define a origem da redação; as demais notas, sim.
        if b.role!='new_note' or not norm(b.text).startswith(('redacao dada','nova redacao','acrescido','acrescida')):continue
        ts=[replace(t,corpus=node.target.corpus) for t in referenced_targets(b.text.split(' pel')[0])]
        if not ts:continue
        matched=[t for t in ts if t.contains(node.target)]
        if not matched:continue
        if 'caput' in norm(b.text) and all(t!=node.target for t in matched):continue
        candidates.append((max(t.level for t in matched),b.index,b,ts))
    if not candidates:return None,[]
    _,_,block,targets=max(candidates,key=lambda x:x[:2])
    return block,targets


def _ordinal(value):
    if value.isdigit() and int(value)<=9:return value+'º'
    return value


def _enumerate(values):
    return ', '.join(values[:-1])+' e '+values[-1] if len(values)>1 else values[0]


def note_for_targets(text,targets):
    tail=re.search(r'\b(pel[ao])\s+(.+)$',text,re.I)
    if not tail or not targets:return text
    first=targets[0];level=first.level;plural=len(targets)>1
    if not plural:description=first.description()
    elif all(t.parent()==first.parent() and t.level==level for t in targets):
        labels={0:'arts.',1:'§§',2:'incisos',3:'alíneas',4:'itens'}
        values=[_ordinal(t.parts[level]) if level<2 else t.parts[level] for t in targets]
        description=labels[level]+' '+_enumerate(values)
        if level>0:description+=' do '+first.parent().description()
    else:return text
    feminine=level==3
    n=norm(text)
    if n.startswith(('acrescido','acrescida')):
        verb=('Acrescidas as ' if feminine else 'Acrescidos os ') if plural else ('Acrescida a ' if feminine else 'Acrescido o ')
        description=description.replace(' do ',' ao ',1)
    elif n.startswith('redacao dada'):
        verb=('Redação dada às ' if feminine else 'Redação dada aos ') if plural else ('Redação dada à ' if feminine else 'Redação dada ao ')
    else:return text
    return verb+description+' '+tail[1]+' '+tail[2]


def replace_text(element,new_text):
    """Altera só os caracteres necessários, preservando runs, campos e hyperlinks."""
    nodes=list(element.iter(qn('w:t')))
    old=''.join(t.text or '' for t in nodes)
    if old==new_text:return element
    prefix=0
    while prefix<min(len(old),len(new_text)) and old[prefix]==new_text[prefix]:prefix+=1
    suffix=0
    while suffix<min(len(old)-prefix,len(new_text)-prefix) and old[-1-suffix]==new_text[-1-suffix]:suffix+=1
    end=len(old)-suffix
    middle=new_text[prefix:len(new_text)-suffix if suffix else len(new_text)]
    cursor=0;placed=False
    for node in nodes:
        text=node.text or '';stop=cursor+len(text)
        if stop>=prefix and cursor<=end:
            left=text[:max(0,prefix-cursor)] if cursor<=prefix else ''
            right=text[max(0,end-cursor):] if stop>=end else ''
            node.text=left+(middle if not placed else '')+right
            node.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
            placed=True
        cursor=stop
    if not placed and nodes:nodes[-1].text=(nodes[-1].text or '')+middle
    return element


def links_signature(doc):
    """Conta destinos e campos; não transforma texto simples em link."""
    records=[]
    for part in doc.part.package.parts:
        element=getattr(part,'element',None)
        if element is None:continue
        for h in element.iter(qn('w:hyperlink')):
            rid=h.get(qn('r:id'))
            rel=part.rels.get(rid) if rid else None
            records.append(('link',str(part.partname),str(rel.target_ref) if rel else '',h.get(qn('w:anchor'),'')))
        for instruction in element.iter(qn('w:instrText')):
            if 'HYPERLINK' in (instruction.text or '').upper():records.append(('field',str(part.partname),instruction.text))
        for field in element.iter(qn('w:fldSimple')):
            if 'HYPERLINK' in field.get(qn('w:instr'),'').upper():records.append(('simple',str(part.partname),field.get(qn('w:instr'))))
    return Counter(records)


FAMILIES=[(r'\b(?:lei complementar|lc)\b','leis complementares',r'(?:lei complementar|lc)'),
          (r'\b(?:emenda constitucional|ec)\b','emendas constitucionais',r'(?:emenda constitucional|ec)'),
          (r'\blei\b','leis',r'lei'),(r'\bdecreto\b','decretos',r'decreto')]


def _number_key(value):
    number,_,year=value.partition('/')
    digits=int(re.sub(r'\D','',number) or 0)
    y=int(year) if year.isdigit() else 0
    if y and y<100:y+=2000 if y<70 else 1900
    return (y,digits)


def update_initial_reference(doc,citation):
    """Acrescenta o ato à lista "Alterada pelas ..." e avança "Atualizada até a EC ..."."""
    actual=norm(citation)
    family=next(((plural,word) for pattern,plural,word in FAMILIES if re.search(pattern,actual)),None)
    number=re.search(r'(\d[\d.]*)/(\d{2,4})\b',citation)
    if not family or not number:return
    plural,word=family
    value=number[0]
    for p in doc.paragraphs[:18]:
        n=norm(p.text)
        m=re.match(r'^(?:atualizad[ao]|consolidad[ao])\s+ate\s+(?:a|o)\s+'+word+r'\s*(?:n[o.]?\s*)?([\d.]+/\d{2,4})',n)
        if m:
            old=re.search(re.escape(m[1]).replace('\\.', '\\.'),p.text)
            if old and _number_key(value)>_number_key(m[1]):
                new_text=p.text[:old.start()]+value+p.text[old.end():]
                replace_text(p._p,new_text)
            continue
        if not n.startswith(('alterada','alterado')) or plural not in n:continue
        if re.search(r'(?<![\d.])'+re.escape(value)+r'(?!\d)',p.text):return
        if re.search(r'(?<![\d.])'+re.escape(number[1])+r'/(?:20)?'+number[2][-2:]+r'(?!\d)',p.text):return
        ts=list(p._p.iter(qn('w:t')))
        if ts and (ts[-1].text or '').endswith('.'):
            ts[-1].text=ts[-1].text[:-1]
        p.add_run(', '+value+'.')
        return


def save_preserving_package(doc,source,destination):
    """Só o corpo e, quando necessário, os estilos são regravados."""
    stream=BytesIO();doc.save(stream);stream.seek(0)
    with ZipFile(source) as original,ZipFile(stream) as edited,ZipFile(destination,'w') as out:
        for entry in original.infolist():
            data=original.read(entry.filename)
            if entry.filename in ('word/document.xml','word/styles.xml'):
                candidate=edited.read(entry.filename)
                if etree.tostring(etree.fromstring(data),method='c14n')!=etree.tostring(etree.fromstring(candidate),method='c14n'):
                    data=candidate
            out.writestr(entry,data)
