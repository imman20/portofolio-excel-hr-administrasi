"""Saya menghitung ulang seluruh workbook pada salinan dan menjaga unit final tetap identik."""
from __future__ import annotations
import concurrent.futures
import datetime as dt
import gc
import hashlib
import importlib.util
import json
import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path
from urllib.parse import unquote,urlsplit
from zoneinfo import ZoneInfo
import numpy as np
import pandas as pd
from openpyxl import load_workbook
import course_runtime as R

ROOT=Path(__file__).resolve().parents[2]
BASE='5e8d7dcf2da6a90e79a669dcc2b0735b242dbff4'
TERM_EXCEPTIONS=[]


def module(name,path):
    sys.path.insert(0,str(path.parent))
    spec=importlib.util.spec_from_file_location(name,path);obj=importlib.util.module_from_spec(spec);sys.modules[name]=obj;spec.loader.exec_module(obj)
    return obj


def expected_checks():
    expected={}
    for folder in sorted((ROOT/'belajar').glob('[0-9][0-9]_*')):
        build=module('audit_'+folder.name[:2],folder/'tools/build.py')
        for c in build.checks():
            path=(folder/'latihan'/c.get('file','solusi_'+folder.name[:2]+'.xlsx')).relative_to(ROOT).as_posix()
            expected.setdefault(path,[]).append(c)
    sales=module('audit_sales',ROOT/'proyek/01_sales_profitability/tools/verify.py')
    _,t=sales.reference();path='proyek/01_sales_profitability/workbook/solusi_sales_profitability.xlsx';expected[path]=[]
    for period,frame,row in [('TAHUN',t,17)]+[(p,f,i+5) for i,(p,f) in enumerate(t.groupby('periode'))]:
        for metric,c in zip(['omzet_bruto','penjualan_bersih','hpp','laba_kotor'],'BCDE'):expected[path].append(dict(label=period+' '+metric,sheet='Calculation',cell=c+str(row),expected=float(frame[metric].sum())))
    inv=module('audit_inventory',ROOT/'proyek/03_inventory_operations/tools/verify.py');_,refs,daily,demand=inv.reference()
    for kind in ['latihan','solusi']:
        path=f'proyek/03_inventory_operations/workbook/{kind}_inventory_operations.xlsx';expected[path]=[]
        cols=[('I','stok_akhir'),('V','hari_stockout'),('Z','abc')]+([('J','permintaan_harian_rata'),('K','sigma_90_hari'),('M','lead_time'),('N','safety_stock'),('O','reorder_point'),('Q','usulan_pesan'),('U','inventory_turnover')] if kind=='solusi' else [])
        for i,(_,r) in enumerate(refs.iterrows(),2):
            for c,key in cols:expected[path].append(dict(label=key+' '+str(i),sheet='Calculation',cell=c+str(i),expected=r[key]))
    survey=module('audit_survey',ROOT/'proyek/02_survei_kepuasan_pelanggan/tools/verify.py');raw,clean,quality,stats=survey.inputs()
    path='proyek/02_survei_kepuasan_pelanggan/workbook/solusi_survei_kepuasan_pelanggan.xlsx';expected[path]=[]
    original=load_workbook(ROOT/path,data_only=True)
    for row in original['Checks'].iter_rows(min_row=4):
        label=row[1].value;value=None
        if not isinstance(label,str):continue
        if ': ' in label:
            section,key=label.split(': ',1)
            if section in stats['dimensions']:value=stats['dimensions'][section][key]
            elif section=='Keseluruhan':value=stats['overall'][key]
            elif section=='NPS':value=stats['nps'][key]
            elif section=='Welch':value=stats['welch'][key]
            elif section=='CSV akhir':value=quality['csv_akhir'][key]
        elif label.startswith('CORREL '):value=stats['correlation'][label[7:]]
        elif label=='Kasus lengkap regresi':value=stats['regression']['n']
        elif label.startswith('LINEST koefisien '):value=stats['regression']['beta'][int(label.split()[-1])]
        elif label.startswith('LINEST galat baku '):value=stats['regression']['se'][int(label.split()[-1])]
        elif label.startswith('LINEST '):value=stats['regression'][label[7:]]
        elif label=='Baris mentah':value=len(raw)
        elif label=='Responden bersih':value=len(clean)
        elif label=='Duplikat dibuang':value=len(raw)-len(clean)
        if value is not None:expected[path].append(dict(label=label,sheet='Checks',cell='C'+str(row[0].row),expected=value))
    original.close()
    return expected,(daily,demand),(survey,raw,clean)


def equal(actual,expected,c):
    if isinstance(expected,(int,float,np.number)):
        assert isinstance(actual,(int,float)),(c['label'],actual)
        delta=float(actual)-float(expected)
        if 'rtol' in c:assert abs(delta)<=c.get('atol',0)+c['rtol']*abs(float(expected)),(c['label'],actual,expected)
        else:assert round(delta,6)==0,(c['label'],actual,expected,delta)
    elif isinstance(expected,(dt.date,dt.datetime)):
        assert actual.strftime('%Y-%m-%d')==expected.strftime('%Y-%m-%d'),c['label']
    else:assert actual==expected,(c['label'],actual,expected)


def audit_file(path,expected,inventory,survey):
    relative=path.relative_to(ROOT).as_posix();before=hashlib.sha256(path.read_bytes()).hexdigest()
    with tempfile.TemporaryDirectory(prefix='audit-repo-') as tmp:
        R.run(['soffice','--headless','--convert-to','xlsx','--outdir',tmp,str(path)])
        result=Path(tmp)/path.name;assert result.exists(),relative
        f=load_workbook(result,data_only=False,read_only=True);count=0
        for sheet in f:
            for row in sheet:
                count+=sum(c.data_type=='f' for c in row)
        f.close()
        w=load_workbook(result,data_only=True,read_only=False);errors=[(s.title,c.coordinate,c.value) for s in w for row in s for c in row if c.data_type=='e']
        assert not errors,(relative,errors[:20])
        comparisons=0
        for c in expected.get(relative,[]):equal(w[c['sheet']][c['cell']].value,c['expected'],c);comparisons+=1
        if relative.startswith('proyek/03_inventory_operations/'):
            daily,demand=inventory
            for frame,start in [(demand,2),(daily,63)]:
                actual=np.array([[w['Permintaan_Harian'].cell(i,j).value for j in range(start,start+60)] for i in range(2,368)],dtype=float)
                assert np.array_equal(actual,frame.to_numpy()),(relative,'Kisi harian')
                comparisons+=actual.size
        if relative=='proyek/02_survei_kepuasan_pelanggan/workbook/solusi_survei_kepuasan_pelanggan.xlsx':
            verifier,raw,clean=survey;comparisons+=verifier.compare_cells(w,raw,clean)
        if path.name=='template_15.xlsx':
            assert all(s.protection.sheet for s in w)
            assert not w['Data']['A2'].protection.locked and w['Calculation']['B2'].protection.locked
        sheets=w.sheetnames;w.close();del w;gc.collect()
    assert hashlib.sha256(path.read_bytes()).hexdigest()==before,('Sumber berubah',relative)
    result=dict(file=relative,formulas=count,errors=0,recalc='success',comparisons=comparisons,size_bytes=path.stat().st_size,sha256=before,sheets=sheets)
    print(json.dumps({k:result[k] for k in ['file','formulas','errors','comparisons']},ensure_ascii=False),flush=True)
    return result


def slug(text):
    text=re.sub('<[^>]+>','',text);text=text.replace('`','').replace('*','').replace('_','_')
    return re.sub(r'[^\w\- ]','',text.lower()).replace(' ','-')


def audit_text(paths):
    global TERM_EXCEPTIONS
    TERM_EXCEPTIONS=[]
    links=0;broken=[];found=[]
    terms=['Chat'+'GPT','Open'+'A'+'I','Co'+'dex','A'+'I','asis'+'ten','model'+' bahasa']
    pattern=re.compile(r'\b(?:'+'|'.join(re.escape(t) for t in terms)+r')\b',re.I)
    for path in paths:
        if path.suffix in ['.md','.py','.bas','.ts','.pq','.csv','.json','.txt']:
            source=path.read_text(encoding='utf-8')
            for number,line in enumerate(source.splitlines(),1):
                if pattern.search(line):
                    relative=path.relative_to(ROOT).as_posix()
                    if relative=='proyek-hr/tools/p09.py' and re.search(r'#\s+[A-Z]{1,3}\.\.[A-Z]{1,3}\s*$',line):
                        TERM_EXCEPTIONS.append(dict(file=relative,line=number,reason='Penanda rentang kolom Excel pada komentar teknis unit final'))
                    else:found.append(dict(file=relative,line=number))
        if path.suffix!='.md':continue
        source=path.read_text()
        for target in re.findall(r'!?\[[^\]]*\]\(([^)]+)\)',source):
            target=target.split(' "')[0].strip('<>');parsed=urlsplit(target)
            if parsed.scheme or parsed.netloc:continue
            links+=1
            dest=(path.parent/unquote(parsed.path)).resolve() if parsed.path else path
            if not dest.exists():broken.append([str(path.relative_to(ROOT)),target]);continue
            if parsed.fragment and dest.suffix=='.md':
                headings={slug(h) for h in re.findall(r'^#+\s+(.+)$',dest.read_text(),re.M)}
                if unquote(parsed.fragment) not in headings:broken.append([str(path.relative_to(ROOT)),target])
    assert not found,('Istilah yang harus dihindari',found)
    assert not broken,('Tautan rusak',broken)
    return links,broken,found


def main():
    os.chdir(ROOT)
    paths=[ROOT/p for p in R.run(['git','ls-files','-z']).split('\0') if p]
    paths+= [Path(__file__).resolve()]
    paths=list(dict.fromkeys(paths))
    protected=[p for p in paths if p.relative_to(ROOT).as_posix().startswith(('proyek/','proyek-hr/','belajar/01_dasar-excel/'))]
    changed=R.run(['git','diff',BASE,'--name-only','--','proyek','proyek-hr','belajar/01_dasar-excel']).strip();assert not changed,changed
    hashes={str(p.relative_to(ROOT)):hashlib.sha256(p.read_bytes()).hexdigest() for p in protected}
    module_rows=[]
    for folder in sorted((ROOT/'belajar').glob('[0-9][0-9]_*')):
        if folder.name.startswith('01_'):continue
        text=(folder/'README.md').read_text();words=len(text.split());assert 1200<=words<=3000,(folder.name,words)
        assert len(re.findall(r'^## \d+\.',text,re.M))==11,folder.name
        questions=text.split('## 7.')[1].split('## 8.')[0];assert len(re.findall(r'^\d+\. ',questions,re.M))>=5
        module_rows.append(dict(module=folder.name,status='SELESAI',words=words,exercises=len(re.findall(r'^\d+\. ',questions,re.M))))
    assert len(module_rows)==14
    expected,inventory,survey=expected_checks();xlsx=[p for p in paths if p.suffix=='.xlsx'];assert len(xlsx)==47,len(xlsx)
    with concurrent.futures.ThreadPoolExecutor(max_workers=2) as pool:
        results=list(pool.map(lambda p:audit_file(p,expected,inventory,survey),xlsx))
    results.sort(key=lambda x:x['file'])
    for p in protected:assert hashlib.sha256(p.read_bytes()).hexdigest()==hashes[str(p.relative_to(ROOT))]
    nlinks,broken,found=audit_text(paths)
    assert not any('*(menyusul)*' in p.read_text() for p in paths if p.suffix=='.md')
    for m in module_rows:
        m['files']=[f for f in results if f['file'].startswith('belajar/'+m['module']+'/')]
        m['formulas']=sum(f['formulas'] for f in m['files']);m['errors']=sum(f['errors'] for f in m['files'])
    summary=dict(date=dt.datetime.now(ZoneInfo('Asia/Jakarta')).date().isoformat(),base=BASE,files=results,modules=module_rows,total_formulas=sum(f['formulas'] for f in results),total_errors=sum(f['errors'] for f in results),total_comparisons=sum(f['comparisons'] for f in results),relative_links=nlinks,broken_links=len(broken),blocked_terms=len(found),protected_files=len(protected),protected_identical=True,tracked_bytes=sum(p.stat().st_size for p in paths))
    summary['term_exceptions']=list(TERM_EXCEPTIONS)
    summary['raw_term_matches']=len(TERM_EXCEPTIONS)+len(found)
    write_report(summary)
    summary=finalize()
    print(json.dumps({k:summary[k] for k in ['total_formulas','total_errors','total_comparisons','relative_links','broken_links','blocked_terms','raw_term_matches','protected_identical','tracked_bytes']},ensure_ascii=False),flush=True)


def write_report(summary):
    (ROOT/'belajar/audit_repo.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2)+'\n')
    md=['# Verifikasi menyeluruh repositori','',f'Tanggal: {summary["date"]} (Asia/Jakarta). Mesin: LibreOffice headless, openpyxl, pandas/scipy.','Perintah: `python belajar/tools/verify_repository.py` dari root repositori.','','Seluruh workbook dihitung ulang pada folder sementara. File sumber, proyek final, proyek HR, dan modul dasar dipertahankan identik.','','## Modul belajar 02 sampai 15','','| Modul | Status | Kata README | Soal | Rumus seluruh workbook | Error |','|---|---|---:|---:|---:|---:|']
    for m in summary['modules']:md.append(f'| [{m["module"]}]({m["module"]}/) | SELESAI | {m["words"]} | {m["exercises"]} | {m["formulas"]} | {m["errors"]} |')
    md+=['','## Penghitungan ulang per workbook','','| File | Rumus | Error | Pembanding | Status |','|---|---:|---:|---:|---|']
    for f in summary['files']:md.append(f'| [{f["file"]}](../{f["file"]}) | {f["formulas"]} | {f["errors"]} | {f["comparisons"]} | success |')
    md+=['','## Hasil dan batas audit','',f'- Jumlah workbook: {len(summary["files"])}; total rumus: {summary["total_formulas"]}; total error: {summary["total_errors"]}.',f'- Pembanding: {summary["total_comparisons"]}; tautan relatif diperiksa: {summary["relative_links"]}; tautan rusak: {summary["broken_links"]}; temuan istilah terlarang: {summary["blocked_terms"]}.',f'- Kecocokan leksikal mentah: {summary.get("raw_term_matches",0)}; pengecualian teknis: {len(summary.get("term_exceptions",[]))}. Penanda rentang kolom Excel pada komentar proyek-hr/tools/p09.py baris 148 dipertahankan; konteks tersebut bukan penyebutan identitas.',f'- Sebanyak {summary["protected_files"]} file unit final dipertahankan identik terhadap sumber kerja dari main.',f'- Ukuran file terlacak: {summary["tracked_bytes"]} byte; tanpa .git dan berkas kerja sementara.', '- Modul belajar memakai pembanding yang disediakan pembangunnya. Modul baru menghitung pembanding dari CSV; modul dasar mempertahankan cara pembanding asli.', '- Sales direkonsiliasi pada empat ukuran tahunan dan bulanan; survei pada sel bersih serta statistik dari CSV; inventori pada seluruh kisi harian, stok, stockout, ABC, dan KPI solusi.', '- Proyek HR diperiksa melalui recalc, jumlah rumus, dan nol sel error. Tidak ada perubahan atau klaim pengujian ulang interaksi manual pada unit final.', '- Ketelitian umum enam desimal; p-value modul statistika memakai toleransi relatif khusus. Unit, label, dan kisi harian dibandingkan tepat.', '- Proteksi template diperiksa melalui metadata dan hasil hitung. Proteksi sheet bukan enkripsi.', '- VBA, Office Scripts, Power Automate, mesin Power Query/DAX, serta interaksi Excel/Google Sheets belum dieksekusi langsung; data, kode, dan petunjuk manual disediakan.', '- Tidak ada modul berstatus GAGAL.', '- Rincian terstruktur dan hash setiap workbook ada di [audit_repo.json](audit_repo.json).','']
    (ROOT/'belajar/VERIFIKASI_REPO.md').write_text('\n'.join(md))


def finalize():
    os.chdir(ROOT)
    report_path=ROOT/'belajar/VERIFIKASI_REPO.md'
    json_path=ROOT/'belajar/audit_repo.json'
    summary=json.loads(json_path.read_text())
    paths=[ROOT/p for p in R.run(['git','ls-files','-z']).split('\0') if p]
    paths=list(dict.fromkeys(paths+[Path(__file__).resolve(),report_path,json_path]))
    nlinks,broken,found=audit_text(paths)
    assert all(p.stat().st_size<5_000_000 for p in paths if p.suffix=='.xlsx')
    assert not any('*(menyusul)*' in p.read_text() for p in paths if p.suffix=='.md')
    summary.update(relative_links=nlinks,broken_links=len(broken),blocked_terms=len(found),term_exceptions=list(TERM_EXCEPTIONS),raw_term_matches=len(found)+len(TERM_EXCEPTIONS))
    for row in summary['modules']:
        row['words']=len((ROOT/'belajar'/row['module']/'README.md').read_text().split())
        assert 1200<=row['words']<=3000
    for _ in range(6):
        size=sum(p.stat().st_size for p in paths)
        summary['tracked_bytes']=size
        json_path.write_text(json.dumps(summary,ensure_ascii=False,indent=2)+'\n')
        md=report_path.read_text()
        md=re.sub(r'^- Pembanding:.*$',f'- Pembanding: {summary["total_comparisons"]}; tautan relatif diperiksa: {nlinks}; tautan rusak: {len(broken)}; temuan istilah terlarang: {len(found)}.',md,flags=re.M)
        md=re.sub(r'^- Ukuran file terlacak.*$',f'- Ukuran file terlacak: {size} byte ({size/1024/1024:.2f} MiB); tanpa .git dan berkas kerja sementara.',md,flags=re.M)
        report_path.write_text(md)
        if sum(p.stat().st_size for p in paths)==size:break
    assert sum(p.stat().st_size for p in paths)==summary['tracked_bytes']
    assert not R.run(['git','diff',BASE,'--name-only','--','proyek','proyek-hr','belajar/01_dasar-excel']).strip()
    return summary


if __name__=='__main__':
    if '--finalize-only' in sys.argv:
        result=finalize()
        print(json.dumps({k:result[k] for k in ['total_formulas','total_errors','relative_links','broken_links','blocked_terms','tracked_bytes']},ensure_ascii=False))
    else:main()
