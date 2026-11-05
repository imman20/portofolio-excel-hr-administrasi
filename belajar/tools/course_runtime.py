"""Saya menyiapkan workbook modul dan merekonsiliasi hasil dari CSV secara ketat."""
from __future__ import annotations
import datetime as dt
import importlib.util
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from zoneinfo import ZoneInfo
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import belajar_lib as L

ALLOWED=set('SUM AVERAGE MIN MAX COUNT COUNTA COUNTBLANK ROUND IF IFS AND OR SUMIFS COUNTIFS AVERAGEIFS MAXIFS MINIFS SUMPRODUCT VLOOKUP INDEX MATCH IFERROR IFNA TRIM CLEAN SUBSTITUTE LEFT RIGHT MID LEN VALUE TEXT DATE DATEVALUE YEAR MONTH DAY EOMONTH EDATE TODAY NETWORKDAYS NETWORKDAYS.INTL DATEDIF RANK.EQ CORREL STDEV.S VAR.S CONFIDENCE.T T.TEST T.INV.2T NORM.S.INV LINEST FORECAST.LINEAR TREND SWITCH TEXTJOIN ABS SQRT PERCENTILE.INC QUARTILE.INC MEDIAN LARGE SMALL'.split())
FUTURE=set('IFNA IFS MAXIFS MINIFS RANK.EQ CORREL STDEV.S VAR.S CONFIDENCE.T T.TEST T.INV.2T NORM.S.INV FORECAST.LINEAR SWITCH TEXTJOIN PERCENTILE.INC QUARTILE.INC NETWORKDAYS.INTL'.split())-{'CORREL','NETWORKDAYS.INTL'}


def finish(wb):
    for ws in wb:
        if not ws.freeze_panes:ws.freeze_panes='B2'
        ws.sheet_view.showGridLines=False
        ws.sheet_properties.pageSetUpPr.fitToPage=True
        ws.page_setup.orientation='landscape';ws.page_setup.paperSize=ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth=1;ws.page_setup.fitToHeight=1
        for row in ws:
            for c in row:
                if c.number_format in ['yyyy-mm-dd','yyyy/mm/dd']:c.number_format='dd/mm/yyyy'
                if c.number_format=='"Rp"#,##0':c.number_format='#,##0'
                if c.data_type=='f':
                    for f in FUTURE:c.value=re.sub(r'(?<![\w.])'+re.escape(f)+r'\(', '_xlfn.'+f+'(',c.value)
        ws.print_options.horizontalCentered=True
    return wb


def new(root,title,df,solution=True,data_name='Data',formats=None):
    wb=L.wb_baru()
    L.sheet_teks(wb,'Instruksi',[
        '# '+title,'Saya memakai data sintetis untuk latihan, bukan pengalaman kerja.',
        'Saya mengisi sel kuning pada Latihan dan membandingkan hasil dengan solusi.',
        'Data mentah/input berfont biru; rumus hitam; tautan antar-sheet hijau.',
        'Formulas > Calculation Options > Automatic, lalu Ctrl+Alt+F9.',
        'Saya membaca README.md dan latihan/kunci_jawaban.md untuk langkah lengkap.',
        'Fitur manual yang dijelaskan belum tertanam kecuali dinyatakan secara khusus.'
    ])
    ws=wb.create_sheet(data_name);L.tulis_df(ws,df,font=L.INPUT_FONT,formats=formats or {})
    ws.auto_filter.ref=ws.dimensions
    lat=wb.create_sheet('Latihan');L.judul(lat,title,sub='Saya mengisi sel kuning; rumus tetap hidup pada solusi.')
    lat.column_dimensions['A'].width=4;lat.column_dimensions['B'].width=53;lat.column_dimensions['C'].width=24
    for c in 'DEFGHIJ':lat.column_dimensions[c].width=18
    return wb,lat


def answer(ws,row,label,formula,solution,fmt='#,##0.00',column='C'):
    L.label(ws,f'B{row}',label,bold=False)
    L.sel_jawab(ws,f'{column}{row}',formula if solution else None,fmt)


def check(label,cell,expected,sheet='Latihan',file=None):
    return dict(label=label,cell=cell,expected=expected,sheet=sheet,**({'file':file} if file else {}))


def run(command,timeout=600):
    res=subprocess.run(command,capture_output=True,text=True,timeout=timeout,env={**os.environ,"TZ":"Asia/Jakarta"})
    if res.returncode:raise RuntimeError(res.stderr[-1500:] or res.stdout[-1500:])
    return res.stdout


def recalc(path):
    with tempfile.TemporaryDirectory(prefix='modul-recalc-') as temp:
        run(['soffice','--headless','--convert-to','xlsx','--outdir',temp,str(path)])
        out=Path(temp)/path.name
        if not out.exists():raise RuntimeError('Penghitungan ulang belum berhasil: '+path.name)
        shutil.copyfile(out,path)


def links(root,files):
    for p in files:
        source=p.read_text()
        for target in re.findall(r'!?\[[^\]]*\]\(([^)]+)\)',source):
            target=target.split(' "')[0].strip('<>')
            if target.startswith(('http:','https:','mailto:','#')):continue
            from urllib.parse import unquote
            destination=(p.parent/unquote(target.split('#')[0])).resolve()
            assert destination.exists(),(p,target)


def render(path,root,sheet='Latihan',area='A1:G24'):
    ns={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    with tempfile.TemporaryDirectory(prefix='modul-render-') as temp:
        source=Path(temp)/'preview.xlsx'
        with zipfile.ZipFile(path) as inp,zipfile.ZipFile(source,'w',zipfile.ZIP_DEFLATED) as out:
            for item in inp.infolist():
                data=inp.read(item.filename)
                if item.filename=='xl/workbook.xml':
                    tree=ET.fromstring(data)
                    target=None
                    for i,s in enumerate(tree.findall('m:sheets/m:sheet',ns)):
                        s.set('state','visible' if s.get('name')==sheet else 'hidden')
                        if s.get('name')==sheet:target=i
                    assert target is not None,sheet
                    for v in tree.findall('m:bookViews/m:workbookView',ns):v.set('activeTab',str(target))
                    defined=tree.find('m:definedNames',ns)
                    if defined is None:defined=ET.SubElement(tree,'{'+ns['m']+'}definedNames')
                    for name in list(defined):
                        if name.get('name')=='_xlnm.Print_Area':defined.remove(name)
                    n=ET.SubElement(defined,'{'+ns['m']+'}definedName',{'name':'_xlnm.Print_Area','localSheetId':str(target)})
                    n.text="'"+sheet+"'!"+area
                    data=ET.tostring(tree,encoding='utf-8',xml_declaration=True)
                out.writestr(item,data)
        run(['soffice','--headless','--convert-to','pdf','--outdir',temp,str(source)])
        pdf=Path(temp)/'preview.pdf';assert pdf.exists()
        assets=root/'assets';assets.mkdir(exist_ok=True)
        shutil.copyfile(pdf,assets/'preview.pdf')
        run(['pdftoppm','-f','1','-singlefile','-scale-to','1800','-png',str(pdf),str(assets/'preview')])


def verify(root):
    root=Path(root).resolve()
    spec=importlib.util.spec_from_file_location('unit_build',root/'tools/build.py')
    build=importlib.util.module_from_spec(spec);spec.loader.exec_module(build)
    files=sorted((root/'latihan').glob('*.xlsx'));assert len(files)>=2
    results=[];values={};comparisons=[]
    for path in files:
        recalc(path)
        wf=load_workbook(path,data_only=False);wv=load_workbook(path,data_only=True)
        count=0;errors=[]
        for ws in wf:
            assert ws.freeze_panes,(path.name,ws.title)
            for row in ws:
                for c in row:
                    if c.data_type=='f':
                        assert isinstance(c.value,str),'Rumus array tidak diperbolehkan'
                        expr=re.sub(r'"(?:[^"]|"")*"','""',c.value).replace('_xlfn.','')
                        funcs={f.upper() for f in re.findall(r'([A-Za-z_][A-Za-z_0-9.]*)\s*\(',expr)};assert funcs<=ALLOWED,(c.coordinate,funcs-ALLOWED)
                        assert not re.search(r'\[[^]]+\.xlsx',expr,re.I),'Tautan eksternal'
                        assert wv[ws.title][c.coordinate].value is not None,(ws.title,c.coordinate,'cache kosong')
                        count+=1
        for ws in wv:
            errors += [(ws.title,c.coordinate,c.value) for row in ws for c in row if c.data_type=='e']
        assert not errors,errors[:15]
        assert path.stat().st_size<5_000_000
        assert any(str(c.fill.fgColor.rgb).endswith('FFF2CC') and c.value is None for ws in wf for row in ws for c in row) if path.name.startswith('latihan_') else True
        values[path.name]=wv
        results.append(dict(file=path.name,sheets=wv.sheetnames,formulas=count,errors=len(errors),size_kb=round(path.stat().st_size/1024,1),recalc='success'))
    for c in build.checks():
        file=c.get('file','solusi_'+root.name[:2]+'.xlsx')
        actual=values[file][c['sheet']][c['cell']].value;expected=c['expected']
        if isinstance(expected,(int,float)):
            assert isinstance(actual,(int,float)),(c['label'],actual)
            rawdiff=float(actual)-float(expected)
            if 'rtol' in c:
                assert abs(rawdiff)<=c.get('atol',0)+c['rtol']*abs(float(expected)),(c['label'],expected,actual,rawdiff)
                diff=0
            else:
                diff=round(rawdiff,6)
                assert diff==0,(c['label'],expected,actual,diff)
        elif isinstance(expected,(dt.date,dt.datetime)):
            assert actual.strftime('%Y-%m-%d')==expected.strftime('%Y-%m-%d');diff=0
        else:assert actual==expected,(c['label'],expected,actual);diff=0
        comparisons.append(dict(label=c['label'],location=file+':'+c['sheet']+'!'+c['cell'],expected=str(expected),actual=str(actual),difference=diff,status='PASS'))
    assert comparisons,'Pembanding pandas wajib tersedia'
    render(root/'latihan'/('solusi_'+root.name[:2]+'.xlsx'),root,getattr(build,'PREVIEW_SHEET','Latihan'),getattr(build,'PREVIEW_AREA','A1:G24'))
    date=dt.datetime.now(ZoneInfo('Asia/Jakarta')).date().isoformat()
    report=dict(module=root.name,date=date,files=results,comparisons=comparisons)
    (root/'latihan/verifikasi.json').write_text(json.dumps(report,ensure_ascii=False,indent=2,default=str)+'\n')
    md=[f'# Verifikasi modul `{root.name}`','',f'Tanggal verifikasi: {date} (Asia/Jakarta). Mesin: LibreOffice headless (recalc) + openpyxl + pandas/scipy.','Skrip: `python tools/course_runtime.py '+root.name+'` dari folder `belajar/`.','','## Hasil recalc per file','','| File | Sheet | Jumlah rumus | Jumlah error | Status recalc | Ukuran (KB) |','|---|---|---:|---:|---|---:|']
    for f in results:md.append(f'| `{f["file"]}` | {", ".join(f["sheets"])} | {f["formulas"]} | {f["errors"]} | success | {f["size_kb"]} |')
    md+=['','## Rekonsiliasi pandas vs sel hasil recalc','','| Pemeriksaan | Lokasi sel | Nilai pandas | Nilai sel | Selisih | Status |','|---|---|---:|---:|---:|---|']
    for c in comparisons:md.append('| '+' | '.join(str(c[k]) for k in ['label','location','expected','actual','difference','status'])+' |')
    md+=['',f'Ringkasan: {len(files)} file, total error = 0, rekonsiliasi gagal = 0.','','## Catatan','','- Sel kuning pada latihan sengaja kosong. Semua perhitungan pada solusi memakai rumus hidup.', '- Pembanding dibaca dari CSV dan dihitung dengan pandas/scipy; statistik dibandingkan sampai enam desimal, selisih harus nol.', '- Preview berasal dari render nyata sheet '+getattr(build,'PREVIEW_SHEET','Latihan')+' ke PDF dengan LibreOffice, kemudian PNG melalui pdftoppm.', '- Fungsi diperiksa terhadap daftar yang diperbolehkan; tidak ada rumus array dinamis atau tautan workbook eksternal.', '- Interaksi manual di Excel dan Google Sheets belum diuji langsung; fitur yang memerlukan langkah manual dijelaskan jujur di README.']+['- '+n for n in getattr(build,'CATATAN_VERIFIKASI',[])]+['']
    (root/'VERIFIKASI.md').write_text('\n'.join(md))
    links(root,list(root.rglob('*.md')))
    print(json.dumps(dict(module=root.name,files=results,checks=len(comparisons)),ensure_ascii=False))
    return report


if __name__=='__main__':verify(Path(sys.argv[1]))
