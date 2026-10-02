#!/usr/bin/env python3
"""Gera relatórios Fleet a partir do modelo, da planilha e das evidências por TAG."""
import argparse
import math
import os
import re
import shutil
import subprocess
import tempfile
from datetime import datetime, date
from io import BytesIO
from pathlib import Path
from PIL import Image, ImageDraw, ImageOps
from openpyxl import load_workbook
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.shared import Emu
from docx.text.paragraph import Paragraph

ROOT=Path(__file__).resolve().parent
SLOTS=('foto_ativo','foto_antes','foto_antes_adicionais','foto_painel','foto_gnss','foto_chicotes','print_classe','print_configuracao','foto_comissionado','checklist_assinado')
GROUPS={'PAINEL':7,'GNSS':8,'CHICOTES':9,'SISTEMA':10}
IMAGE_EXTENSIONS=('.jpg','.jpeg','.png','.webp','.tif','.tiff')

def find_office():
    found=shutil.which('libreoffice') or shutil.which('soffice')
    if found:return found
    for name in ('ProgramFiles','ProgramFiles(x86)'):
        root=os.environ.get(name)
        if root:
            path=Path(root)/'LibreOffice'/'program'/'soffice.exe'
            if path.is_file():return str(path)
    return None

def export_pdf(target):
    """Exporta PDF no Windows pelo Word; usa LibreOffice quando disponível."""
    output=target.with_suffix('.pdf')
    output.unlink(missing_ok=True)
    failures=[]
    with tempfile.TemporaryDirectory(prefix='fleet_pdf_') as work:
        temp=Path(work);input_file=temp/target.name
        temp_pdf=input_file.with_suffix('.pdf')
        shutil.copy2(target,input_file)
        if os.name=='nt':
            powershell=shutil.which('powershell.exe')
            if powershell:
                script=temp/'export_word_pdf.ps1'
                script.write_text('''
$ErrorActionPreference = 'Stop'
$word = $null
$document = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3
    $document = $word.Documents.Open($args[0], $false, $true)
    $document.ExportAsFixedFormat($args[1], 17)
} finally {
    if ($null -ne $document) { $document.Close(0) }
    if ($null -ne $word) { $word.Quit(0) }
}
''',encoding='utf-8')
                try:
                    proc=subprocess.run([powershell,'-NoProfile','-NonInteractive','-Sta',
                                         '-ExecutionPolicy','Bypass','-File',str(script),
                                         str(input_file),str(temp_pdf)],
                                        capture_output=True,text=True,timeout=120)
                    if proc.returncode==0 and temp_pdf.is_file() and temp_pdf.stat().st_size:
                        shutil.copy2(temp_pdf,output)
                        return 'Microsoft Word'
                    failures.append('Word: '+(proc.stderr.strip() or proc.stdout.strip() or 'PDF não criado'))
                except (OSError,subprocess.TimeoutExpired) as e:
                    failures.append(f'Word: {e}')
            else:
                failures.append('PowerShell não encontrado para exportar pelo Word')
        office=find_office()
        if office:
            try:
                temp_pdf.unlink(missing_ok=True)
                proc=subprocess.run([office,f'-env:UserInstallation={(temp / "profile").as_uri()}',
                                     '--headless','--convert-to','pdf','--outdir',str(temp),str(input_file)],
                                    capture_output=True,text=True,timeout=120)
                if proc.returncode==0 and temp_pdf.is_file() and temp_pdf.stat().st_size:
                    shutil.copy2(temp_pdf,output)
                    return 'LibreOffice'
                failures.append('LibreOffice: '+(proc.stderr.strip() or proc.stdout.strip() or 'PDF não criado'))
            except (OSError,subprocess.TimeoutExpired) as e:
                failures.append(f'LibreOffice: {e}')
        else:
            failures.append('LibreOffice não encontrado')
    raise RuntimeError('; '.join(failures))

def paragraphs(part):
    for p in part.paragraphs: yield p
    for t in part.tables:
        for row in t.rows:
            for cell in row.cells: yield from paragraphs(cell)

def replace(p,old,new):
    if old==new:return
    runs=p.runs
    if not runs or old not in ''.join(r.text for r in runs):return
    full=''.join(r.text for r in runs)
    while old in full:
        pos=full.index(old);end=pos+len(old)
        spans=[];offset=0
        for r in runs:spans.append((offset,offset+len(r.text)));offset+=len(r.text)
        affected=[i for i,(a,b) in enumerate(spans) if a<end and b>pos]
        first,last=affected[0],affected[-1]
        prefix=runs[first].text[:pos-spans[first][0]]
        suffix=runs[last].text[end-spans[last][0]:]
        runs[first].text=prefix+new+(suffix if first==last else '')
        for j in range(first+1,last):runs[j].text=''
        if last!=first:runs[last].text=suffix
        full=''.join(r.text for r in runs)

def data(value,field):
    if isinstance(value,(date,datetime)):return value.strftime('%d/%m/%Y')
    if isinstance(value,str):
        try:return datetime.strptime(value.strip(),'%d/%m/%Y').strftime('%d/%m/%Y')
        except ValueError:pass
    raise ValueError(f'{field}: preencha uma data válida do Excel ou dd/mm/aaaa')

def text(v):return str(v).strip() if v is not None else ''

def read_rows(book):
    wb=load_workbook(book,read_only=True,data_only=True)
    a=list(wb['Ativos'].values); c=list(wb['Verificacoes'].values)
    assert a[0][0:3]==('GERAR','TAG','MODELO'), 'Cabeçalhos da aba Ativos modificados'
    assert c[0][0:3]==('TAG','GRUPO','NUMERO'), 'Cabeçalhos da aba Verificacoes modificados'
    return [dict(zip(a[0],r)) for r in a[1:] if r[1]], [dict(zip(c[0],r)) for r in c[1:] if r[0]]

def evidence(folder,slot,part):
    found=[p for p in folder.iterdir() if p.is_file() and p.stem.casefold()==slot.casefold() and p.suffix.lower() in IMAGE_EXTENSIONS]
    if len(found)!=1:raise ValueError(f'{slot}: esperado 1 arquivo em {folder} (encontrados {len(found)})')
    try:
        im=ImageOps.exif_transpose(Image.open(found[0])).convert('RGB')
        current=Image.open(BytesIO(part.blob))
        # Retain exactly the picture's aspect ratio, box and Word layout.
        w,h=current.size
        factor=max(1,round(1300/max(w,h)))
        size=(w*factor,h*factor)
        canvas=Image.new('RGB',size,'#FFFFFF')
        im=ImageOps.contain(im,size,method=Image.Resampling.LANCZOS)
        canvas.paste(im,((size[0]-im.width)//2,(size[1]-im.height)//2))
        raw=BytesIO()
        fmt='JPEG' if part.partname.lower().endswith(('.jpeg','.jpg')) else 'PNG'
        canvas.save(raw,format=fmt,quality=90,optimize=True)
        part._blob=raw.getvalue()
    except Exception as e:raise ValueError(f'{slot}: não foi possível abrir/processar {found[0].name}: {e}') from e

def checklist_pdf(tag,folder,checklists):
    """Procura um único PDF 'TAG - Checklist ...' na pasta separada ou na pasta da TAG."""
    prefix=re.compile(r'^'+re.escape(tag)+r'(?:\s*[-_–]\s*|\s+)checklist\b',re.IGNORECASE)
    found=[]
    for directory in (checklists,folder):
        if directory.is_dir():
            found.extend(p for p in directory.iterdir() if p.is_file()
                         and p.suffix.casefold()=='.pdf' and prefix.match(p.stem))
    if len(found)>1:raise ValueError(f'checklist_assinado: mais de um PDF para {tag}: {", ".join(p.name for p in found)}')
    return found[0] if found else None

def render_checklist_pdf(path,placeholder):
    try:
        import pymupdf
    except ImportError as e:
        raise ValueError('checklist em PDF requer PyMuPDF; execute py -3 -m pip install -r requirements.txt') from e
    original=Image.open(BytesIO(placeholder.blob))
    width,height=original.size
    factor=max(1,math.ceil(1900/width));size=(width*factor,height*factor)
    pages=[]
    try:
        with pymupdf.open(path) as pdf:
            if pdf.needs_pass:raise ValueError('PDF protegido por senha')
            if not len(pdf):raise ValueError('PDF sem páginas')
            for page in pdf:
                pix=page.get_pixmap(dpi=200,alpha=False,annots=True)
                with Image.open(BytesIO(pix.tobytes('png'))) as source:
                    photo=ImageOps.contain(source.convert('RGB'),size,method=Image.Resampling.LANCZOS)
                canvas=Image.new('RGB',size,'white')
                canvas.paste(photo,((size[0]-photo.width)//2,(size[1]-photo.height)//2))
                raw=BytesIO();canvas.save(raw,format='PNG',optimize=True)
                pages.append(raw.getvalue())
    except Exception as e:
        raise ValueError(f'checklist_assinado: não foi possível ler {path.name}: {e}') from e
    return pages

def add_checklist_pages(first_paragraph,pages):
    """Adiciona as páginas 2+ após a figura existente, antes de sua legenda."""
    if len(pages)<2:return
    extent=first_paragraph._p.xpath('.//wp:inline/wp:extent')
    if len(extent)!=1:raise ValueError('checklist_assinado: dimensão da figura não encontrada no modelo')
    width,height=Emu(int(extent[0].get('cx'))),Emu(int(extent[0].get('cy')))
    first_paragraph.paragraph_format.keep_with_next=False
    anchor=first_paragraph
    for number,blob in enumerate(pages[1:],2):
        element=OxmlElement('w:p');anchor._p.addnext(element)
        paragraph=Paragraph(element,anchor._parent)
        paragraph.alignment=WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.page_break_before=True
        paragraph.paragraph_format.keep_with_next=(number==len(pages))
        picture=paragraph.add_run().add_picture(BytesIO(blob),width=width,height=height)
        picture._inline.docPr.set('descr',f'CHECKLIST_PDF:pagina_{number}')
        anchor=paragraph
    following=anchor._p.getnext()
    for _ in range(5):
        if following is None:break
        candidate=Paragraph(following,anchor._parent)
        if candidate.text.strip().upper()=='CONCLUSÃO':
            candidate.paragraph_format.page_break_before=True
            break
        following=following.getnext()

def chicote_files(folder):
    numbered={};legacy=[]
    for path in folder.iterdir():
        if not path.is_file() or not path.stem.casefold().startswith('foto_chicotes'):continue
        suffix=path.stem.casefold().removeprefix('foto_chicotes')
        if path.suffix.lower() not in IMAGE_EXTENSIONS:
            raise ValueError(f'foto_chicotes: extensão não aceita em {path.name}')
        if not suffix:legacy.append(path)
        elif suffix in ('1','2','3','4'):
            if suffix in numbered:raise ValueError(f'foto_chicotes{suffix}: duas imagens para o mesmo número')
            numbered[suffix]=path
        else:raise ValueError(f'foto_chicotes: nome inválido {path.name}; use foto_chicotes1 a foto_chicotes4')
    if legacy and numbered:raise ValueError('foto_chicotes: não misture o nome antigo com fotos numeradas')
    if len(legacy)>1:raise ValueError('foto_chicotes: mais de um arquivo com o nome antigo')
    paths=[numbered[str(i)] for i in range(1,5) if str(i) in numbered] if numbered else legacy
    if not paths:raise ValueError(f'foto_chicotes: coloque de 1 a 4 imagens numeradas em {folder}')
    return paths

def mosaic_rectangles(count,width,height,gap):
    x0=y0=gap; x1=width-gap; y1=height-gap
    xm=width//2; ym=height//2
    if count==1:return [(x0,y0,x1,y1)]
    if count==2:return [(x0,y0,xm-gap//2,y1),(xm+gap//2,y0,x1,y1)]
    if count==3:
        split=round(width*.63)
        return [(x0,y0,split-gap//2,y1),(split+gap//2,y0,x1,ym-gap//2),(split+gap//2,ym+gap//2,x1,y1)]
    return [(x0,y0,xm-gap//2,ym-gap//2),(xm+gap//2,y0,x1,ym-gap//2),
            (x0,ym+gap//2,xm-gap//2,y1),(xm+gap//2,ym+gap//2,x1,y1)]

def chicotes_mosaic(folder,part):
    paths=chicote_files(folder)
    try:
        original=Image.open(BytesIO(part.blob))
        width,height=original.size
        scale=max(1,math.ceil(2400/width))
        width*=scale;height*=scale
        gap=max(12,round(width*.008))
        canvas=Image.new('RGB',(width,height),'white')
        draw=ImageDraw.Draw(canvas)
        for path,rect in zip(paths,mosaic_rectangles(len(paths),width,height,gap)):
            left,top,right,bottom=rect
            draw.rectangle((left,top,right-1,bottom-1),fill='#F1F3F5')
            with Image.open(path) as source:
                photo=ImageOps.exif_transpose(source).convert('RGB')
            photo=ImageOps.contain(photo,(right-left,bottom-top),method=Image.Resampling.LANCZOS)
            canvas.paste(photo,(left+(right-left-photo.width)//2,top+(bottom-top-photo.height)//2))
        raw=BytesIO();canvas.save(raw,format='PNG',optimize=True)
        part._blob=raw.getvalue()
    except Exception as e:
        raise ValueError(f'foto_chicotes: não foi possível montar o mosaico: {e}') from e

def fotos_antes_adicionais(folder):
    """Aceita de zero a oito imagens: foto_antes_adicional1 a 8."""
    files={}
    prefix='foto_antes_adicional'
    for path in folder.iterdir():
        if not path.is_file() or not path.stem.casefold().startswith(prefix):continue
        index=path.stem.casefold()[len(prefix):]
        if index not in tuple(str(i) for i in range(1,9)) or path.suffix.casefold() not in IMAGE_EXTENSIONS:
            raise ValueError(f'foto_antes_adicionais: nome/extensão inválidos em {path.name}; use {prefix}1 a {prefix}8')
        if index in files:
            raise ValueError(f'{prefix}{index}: mais de um arquivo para o mesmo número')
        files[index]=path
    return [files[str(i)] for i in range(1,9) if str(i) in files]

def inserir_mosaico_antes(paths,part,drawing):
    """Monta linhas de uma ou duas imagens, inteiras e sem deformação."""
    count=len(paths)
    width=2400
    height={1:1350,2:1350,3:1700,4:1700,5:2200,6:2200,7:2700,8:2700}[count]
    gap=20
    canvas=Image.new('RGB',(width,height),'white')
    draw=ImageDraw.Draw(canvas)
    rows=math.ceil(count/2)
    row_h=(height-(rows+1)*gap)//rows
    try:
        for n,path in enumerate(paths):
            row,col=divmod(n,2)
            lone=col==0 and n==count-1 and count%2==1
            left=gap if lone or col==0 else width//2+gap//2
            right=width-gap if lone or col==1 else width//2-gap//2
            top=gap+row*(row_h+gap)
            bottom=top+row_h
            draw.rectangle((left,top,right-1,bottom-1),fill='#F1F3F5')
            with Image.open(path) as source:
                image=ImageOps.exif_transpose(source).convert('RGB')
            image=ImageOps.contain(image,(right-left,bottom-top),method=Image.Resampling.LANCZOS)
            canvas.paste(image,(left+(right-left-image.width)//2,
                                top+(bottom-top-image.height)//2))
        raw=BytesIO();canvas.save(raw,format='PNG',optimize=True)
        part._blob=raw.getvalue()
    except Exception as e:
        raise ValueError(f'foto_antes_adicionais: não foi possível montar o mosaico: {e}') from e
    original=drawing.xpath('.//wp:inline/wp:extent')
    if len(original)!=1:raise ValueError('foto_antes_adicionais: dimensões da figura 7.2 ausentes')
    new_cx=int(original[0].get('cx'))
    new_cy=round(new_cx*height/width)
    original[0].set('cy',str(new_cy))
    for ext in drawing.xpath('.//a:xfrm/a:ext'):
        ext.set('cy',str(new_cy))

def remover_figura_antes(p):
    """Sem fotos adicionais, remove a imagem e sua legenda do relatório gerado."""
    node=p._p
    caption=node.getnext()
    if caption is None or not ''.join(caption.itertext()).strip().startswith('Figura 7.2'):
        raise ValueError('foto_antes_adicionais: legenda da Figura 7.2 ausente no modelo')
    caption.getparent().remove(caption)
    node.getparent().remove(node)

def generate(row,checks,template,photos,outdir,draft=False,checklists=None):
    tag=text(row['TAG']).upper()
    if not re.fullmatch(r'[A-Z0-9_-]{2,32}',tag):raise ValueError('TAG inválida; use 2 a 32 letras/números, _ ou -')
    folder=photos/tag
    if not folder.is_dir():raise ValueError(f'pasta de evidências ausente: fotos/{tag}')
    entries={};generic={}
    for r in checks:
        source_tag=text(r['TAG']).upper()
        if source_tag not in (tag,'PADRAO'):continue
        group=text(r['GRUPO']).upper()
        try:num=int(r['NUMERO'])
        except (ValueError,TypeError):raise ValueError('NUMERO da verificação inválido')
        key=(group,num)
        destination=generic if source_tag=='PADRAO' else entries
        if key in destination:raise ValueError(f'verificação duplicada em {source_tag}: {key}')
        destination[key]=r
    for f in ('MODELO','VERSAO_MODELO'):
        if not text(row[f]):raise ValueError(f'{f} vazio')
    if text(row['AVALIACAO_INICIAL']).upper()!='APROVADO':raise ValueError('avaliação inicial não aprovada')
    start_date=data(row['DATA_INICIO'],'DATA_INICIO')
    values={'{{TAG}}':tag,'{{VERSAO_MODELO}}':text(row['VERSAO_MODELO']),
            '{{DATA_INICIO}}':start_date,'{{DATA_INÍCIO}}':start_date,
            '{{DATA_CONCLUSAO}}':data(row['DATA_CONCLUSAO'],'DATA_CONCLUSAO'),
            '{{DATA_COMISSIONAMENTO}}':data(row['DATA_COMISSIONAMENTO'],'DATA_COMISSIONAMENTO')}
    values['{{DATA_EMISSAO}}']=datetime.strptime(data(row['DATA_EMISSAO'],'DATA_EMISSAO'),'%d/%m/%Y').strftime('%d/%m/%y')
    doc=Document(template)
    # Every condition belongs to the inspected asset; no CH14 status is inherited.
    expected=set()
    for group,table_i in GROUPS.items():
        table=doc.tables[table_i]
        for number in range(1,len(table.rows)):
            key=(group,number);expected.add(key)
            if key not in entries and key not in generic:raise ValueError(f'verificação ausente: {group} {number}')
            specific=key in entries
            r=entries[key] if specific else generic[key]
            label=table.cell(number,0 if group!='SISTEMA' else 1).text.strip()
            # A lista PADRAO usa a redação atual do Word; itens específicos ainda são conferidos.
            if specific and text(r['ITEM'])!=label:raise ValueError(f'ITEM divergente em {group} {number}')
            status=text(r['CONDICAO']).upper()
            if status!='OK':raise ValueError(f'{group} {number}: condição {status or "VAZIA"}; este modelo afirma conclusão sem pendências')
            table.cell(number,1 if group!='SISTEMA' else 2).text=status
            table.cell(number,2 if group!='SISTEMA' else 3).text=text(r['OBSERVACOES'])
    unexpected=(set(entries)|set(generic))-expected
    if unexpected:raise ValueError(f'verificações extras: {sorted(unexpected)}')
    for p in paragraphs(doc):
        for old,new in values.items():replace(p,old,new)
    # Modelo da tabela é editável por ativo, sem trocar a classe textual do template.
    for p in doc.tables[4].cell(0,1).paragraphs:
        for old in [p.text]:
            if old:replace(p,old,text(row['MODELO']))
    for section in doc.sections:
        for header in (section.header,section.even_page_header):
            for p in paragraphs(header):
                for old,new in values.items():replace(p,old,new)
    seen=set();extra_checklist_pages=None;first_checklist_paragraph=None
    for p in doc.paragraphs:
        for draw in p._p.xpath('.//w:drawing'):
            props=draw.xpath('.//wp:docPr')
            if not props:continue
            descr=props[0].get('descr','')
            if not descr.startswith('EVIDENCIA:'):continue
            slot=descr.split(':',1)[1]
            if slot not in SLOTS or slot in seen:raise ValueError(f'posição de imagem inválida: {slot}')
            rid=draw.xpath('.//a:blip')[0].get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
            part=doc.part.related_parts[rid]
            if slot=='foto_chicotes':chicotes_mosaic(folder,part)
            elif slot=='foto_antes_adicionais':
                files=fotos_antes_adicionais(folder)
                if files:inserir_mosaico_antes(files,part,draw)
                else:remover_figura_antes(p)
            elif slot=='checklist_assinado':
                pdf=checklist_pdf(tag,folder,checklists or ROOT/'checklists')
                if pdf:
                    pages=render_checklist_pdf(pdf,part)
                    part._blob=pages[0]
                    extra_checklist_pages=pages
                    first_checklist_paragraph=p
                else:evidence(folder,slot,part)
            else:evidence(folder,slot,part)
            seen.add(slot)
    if seen!=set(SLOTS):raise ValueError(f'posições de foto ausentes: {set(SLOTS)-seen}')
    if extra_checklist_pages:add_checklist_pages(first_checklist_paragraph,extra_checklist_pages)
    filename=('RASCUNHO_' if draft else '')+'Checklist_Entrega_'+tag
    target=outdir/(filename+'.docx')
    # Save only after all validations pass; failed assets never produce a partial report.
    doc.save(target)
    try:
        method=export_pdf(target)
        print(f'  PDF gerado para {tag} por {method}: {target.with_suffix(".pdf").name}')
    except Exception as e:
        print(f'  PDF pendente para {tag}: {e}; DOCX gerado.')
    return target

def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--planilha',type=Path,default=ROOT/'Controle_Ativos.xlsx')
    parser.add_argument('--modelo',type=Path,default=ROOT/'modelo'/'Modelo_Caminhonete_Fleet.docx')
    parser.add_argument('--fotos',type=Path,default=ROOT/'fotos')
    parser.add_argument('--checklists',type=Path,default=ROOT/'checklists',help='pasta dos checklists PDF nomeados pela TAG')
    parser.add_argument('--saida',type=Path,default=ROOT/'saida')
    parser.add_argument('--tag',help='gera apenas esta TAG, se marcada para gerar')
    parser.add_argument('--teste',action='store_true',help='gera também linhas TESTE com prefixo RASCUNHO')
    args=parser.parse_args();args.saida.mkdir(parents=True,exist_ok=True)
    assets,checks=read_rows(args.planilha)
    selected=[r for r in assets if (not args.tag or text(r['TAG']).upper()==args.tag.upper()) and (text(r['GERAR']).upper()=='SIM' or (args.teste and text(r['GERAR']).upper()=='TESTE'))]
    if not selected:
        print('Nenhuma TAG elegível. Marque GERAR=SIM, ou use --teste para uma TAG TESTE.');return 0
    seen=set();errors=0
    for r in selected:
        tag=text(r['TAG']).upper()
        try:
            if tag in seen:raise ValueError('TAG repetida na aba Ativos')
            seen.add(tag)
            draft=text(r['GERAR']).upper()=='TESTE'
            result=generate(r,checks,args.modelo,args.fotos,args.saida,draft,args.checklists)
            print(f'OK {tag}: {result.name}')
        except Exception as e:print(f'ERRO {tag}: {e}');errors+=1
    print(f'Concluído: {len(selected)-errors}/{len(selected)} relatórios.')
    return 1 if errors else 0
if __name__=='__main__':raise SystemExit(main())
