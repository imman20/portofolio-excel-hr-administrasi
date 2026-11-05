"""Saya memeriksa input, menghitung ulang workbook, dan merekam bukti rekonsiliasi."""
from __future__ import annotations
import csv
import copy
import datetime as dt
import hashlib
import json
import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd
import scipy
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from analysis_support import BASE, inputs, audit

ALLOWED = set('SUM AVERAGE MIN MAX COUNT COUNTA COUNTBLANK ROUND IF IFS AND OR SUMIFS COUNTIFS AVERAGEIFS MAXIFS MINIFS SUMPRODUCT VLOOKUP INDEX MATCH IFERROR IFNA TRIM CLEAN SUBSTITUTE LEFT RIGHT MID LEN VALUE TEXT DATE DATEVALUE YEAR MONTH DAY EOMONTH EDATE TODAY NETWORKDAYS NETWORKDAYS.INTL DATEDIF RANK.EQ CORREL STDEV.S VAR.S CONFIDENCE.T T.TEST T.INV.2T NORM.S.INV LINEST FORECAST.LINEAR TREND SWITCH TEXTJOIN ABS SQRT PERCENTILE.INC QUARTILE.INC MEDIAN LARGE SMALL'.split())
REQUIRED = {'Readme','Codebook','Raw_Data','Master','Clean_Data','Calculation','Report','Checks','Reference','Check_Detail','Analisis_Set'}


def run(command):
    r = subprocess.run(command,capture_output=True,text=True,timeout=180)
    if r.returncode: raise RuntimeError(r.stderr[-2000:] or r.stdout[-2000:])
    return r.stdout


def recalc(path, temp):
    out = temp/'recalc'
    out.mkdir(exist_ok=True)
    run(['soffice','--headless','--norestore',f'-env:UserInstallation={(temp/"profile").as_uri()}', '--convert-to','xlsx','--outdir',str(out),str(path)])
    result=out/path.name
    assert result.exists() and result.stat().st_size>0, path.name
    shutil.copyfile(result,path)


def inspect(path):
    formula=load_workbook(path,data_only=False)
    value=load_workbook(path,data_only=True)
    assert REQUIRED<=set(formula.sheetnames)
    assert not formula._external_links and path.stat().st_size<5_000_000
    if path.name.startswith('solusi_'): assert len(formula['Report']._charts)==3
    count=0
    errors=[]
    functions=set()
    for ws in formula:
        assert ws.freeze_panes,ws.title
        for row in ws:
            for c in row:
                if c.data_type=='f':
                    count+=1
                    assert isinstance(c.value,str),(ws.title,c.coordinate,'rumus array')
                    expr=re.sub(r'"(?:[^"]|"")*"','""',c.value).replace('_xlfn.','').replace('_xlws.','')
                    f=set(re.findall(r'([A-Za-z_][A-Za-z_0-9.]*)\s*\(',expr))
                    assert f<=ALLOWED,(ws.title,c.coordinate,f-ALLOWED)
                    functions |= f
                if value[ws.title][c.coordinate].data_type=='e': errors.append((ws.title,c.coordinate,value[ws.title][c.coordinate].value))
    assert not errors, errors[:20]
    result=dict(file=path.name,sheets=formula.sheetnames,formula_count=count,error_count=0,status='success (LibreOffice headless)',size_kb=round(path.stat().st_size/1024,1),functions=sorted(functions))
    formula.close()
    return value,result


def compare_cells(wb, raw, ref):
    """Pemeriksaan Python membaca sel nyata, terpisah dari Checks di workbook."""
    clean=wb['Clean_Data']
    mapping={c.value:c.column for c in clean[1]}
    kept={clean.cell(r,1).value:r for r in range(2,422) if clean.cell(r,24).value==1}
    assert len(kept)==len(ref)==414
    assert sum(clean.cell(r,24).value==0 for r in range(2,422))==6
    compared=0
    for record in ref.to_dict('records'):
        r=kept[record['respondent_id']]
        for name,expected in record.items():
            actual=clean.cell(r,mapping[name]).value
            if name=='tanggal_survei': actual=actual.strftime('%Y-%m-%d')
            if pd.isna(expected) or expected=='': assert actual in (None,''),(r,name,actual,'kosong')
            elif isinstance(expected,(float,int,np.number)):
                assert isinstance(actual,(float,int)),(r,name,actual,expected)
                assert round(float(actual)-float(expected),12)==0,(r,name,actual,expected)
            else: assert actual==expected,(r,name,actual,expected)
            compared+=1
    # Raw_Data must preserve every raw input value, including whitespace and literals.
    for r,record in enumerate(raw.itertuples(index=False),2):
        for j,expected in enumerate(record,1):
            actual=wb['Raw_Data'].cell(r,j).value
            assert (actual if actual is not None else '')==expected,(r,j,actual,expected)
    for row in list(wb['Check_Detail'].values)[3:]: assert row[-1]=='PASS',row[:2]
    return compared


def preview(path,temp):
    ns={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    source=temp/'report.xlsx'
    with zipfile.ZipFile(path) as zin,zipfile.ZipFile(source,'w',zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data=zin.read(item.filename)
            if item.filename=='xl/workbook.xml':
                xml=ET.fromstring(data)
                report_index=next(i for i,s in enumerate(xml.findall('m:sheets/m:sheet',ns)) if s.get('name')=='Report')
                for i,sheet in enumerate(xml.findall('m:sheets/m:sheet',ns)):
                    sheet.set('state','visible' if sheet.get('name')=='Report' else 'hidden')
                    if sheet.get('name')=='Report':
                        for view in xml.findall('m:bookViews/m:workbookView',ns):view.set('activeTab',str(i))
                # LibreOffice can export explicit print areas on hidden worksheets.
                # Only the Report print definition belongs in the render copy.
                names=xml.find('m:definedNames',ns)
                if names is not None:
                    for name in list(names):
                        if name.get('name')=='_xlnm.Print_Area' and name.get('localSheetId')!=str(report_index):names.remove(name)
                data=ET.tostring(xml,encoding='utf-8',xml_declaration=True)
            zout.writestr(item,data)
    run(['soffice','--headless','--norestore',f'-env:UserInstallation={(temp/"preview-profile").as_uri()}','--convert-to','pdf','--outdir',str(temp),str(source)])
    assets=BASE/'assets';assets.mkdir(exist_ok=True)
    pdf=temp/'report.pdf';assert pdf.exists()
    shutil.copyfile(pdf,assets/'report.pdf')
    info=run(['pdfinfo',str(pdf)])
    pages=int(re.search(r'Pages:\s+(\d+)',info).group(1))
    assert pages==2,('Jumlah halaman Report',pages)
    for i in range(1,pages+1):
        name='report' if i==1 else f'report_{i}'
        run(['pdftoppm','-f',str(i),'-l',str(i),'-singlefile','-scale-to','2200','-png',str(pdf),str(assets/name)])
    return pages


def negative_tests(raw, ref, wb):
    """Saya memastikan pemeriksaan menolak kerusakan yang pernah relevan pada bahan."""
    notes=json.loads((BASE/'data/processed/catatan_masalah_disengaja.json').read_text())
    tested=[]
    def reject(name, operation):
        try: operation()
        except (AssertionError, ValueError): tested.append(dict(case=name,status='DITOLAK'))
        else: raise AssertionError('Kerusakan tidak terdeteksi: '+name)
    conflicting=raw.copy()
    i=conflicting.index[conflicting.respondent_id.duplicated()][0]
    conflicting.loc[i,'komentar']='Komentar bertentangan pada salinan ID'
    reject('Duplikat ID berkonflik',lambda:audit(conflicting,notes))
    incorrect=copy.deepcopy(notes);incorrect['sel_kosong_skor']+=1
    reject('Hitungan JSON salah',lambda:audit(raw,incorrect))
    wrong=ref.copy();wrong.loc[0,'tanggal_survei']='2024-07-05'
    reject('Tanggal ISO referensi tertukar',lambda:pd.testing.assert_frame_equal(ref,wrong))
    cell=wb['Clean_Data']['I2'];original=cell.value
    cell.value=1 if original!=1 else 5
    reject('Nilai sel hasil pembersihan berubah',lambda:compare_cells(wb,raw,ref))
    cell.value=original
    return tested


def main():
    raw,ref,quality,stats=inputs()
    files=sorted((BASE/'workbook').glob('*.xlsx'))
    assert len(files)==2
    rows=[];results=[];cell_count=0;pages=0;negatives=[]
    with tempfile.TemporaryDirectory(prefix='survei-verifikasi-') as folder:
        temp=Path(folder)
        for path in files:
            # Check allowed functions both before and after recalculation.
            before,_=inspect(path);before.close()
            recalc(path,temp)
            wb,result=inspect(path)
            solution=path.name.startswith('solusi_')
            if solution:
                cell_count=compare_cells(wb,raw,ref)
                for row in list(wb['Checks'].values)[3:-1]:
                    _,label,actual,expected,delta,status=row[:6]
                    assert isinstance(actual,(float,int)) and isinstance(expected,(float,int)),label
                    difference=round(actual-expected,10)
                    assert difference==0 and delta==0 and status=='PASS',(label,actual,expected,delta,status)
                    rows.append(dict(pemeriksaan=label,lokasi=f'Checks!C{len(rows)+4}',python=expected,workbook=actual,selisih=difference,status=status))
                assert list(wb['Checks'].values)[-1][5]=='SEMUA PASS'
                negatives=negative_tests(raw,ref,wb)
                result['reconciliation']='PASS'
            else:
                statuses=[r[5] for r in list(wb['Checks'].values)[3:]]
                assert all(s=='BELUM DIISI' for s in statuses)
                # Worked example must remain a real formula with a correct result.
                assert wb['Clean_Data']['A2'].value==raw.iloc[0].respondent_id
                result['reconciliation']='PASS contoh; BELUM DIISI pada sel latihan'
            results.append(result);wb.close()
        pages=preview(next(p for p in files if p.name.startswith('solusi_')),temp)
    proc=BASE/'data/processed'
    with (proc/'rekonsiliasi.csv').open('w',newline='',encoding='utf-8') as f:
        writer=csv.DictWriter(f,fieldnames=list(rows[0]),lineterminator="\n");writer.writeheader();writer.writerows(rows)
    summary=dict(date=dt.datetime.now(ZoneInfo('Asia/Jakarta')).isoformat(),engine=run(['soffice','--version']).strip(),pandas=pd.__version__,scipy=scipy.__version__,files=results,quality=quality,statistics=stats,clean_cells_compared=cell_count,comparisons=len(rows),preview_pages=pages,hashes={str(p.relative_to(BASE)):hashlib.sha256(p.read_bytes()).hexdigest() for p in [BASE/'data/raw/survei_raw.csv',proc/'survei_bersih_referensi.csv',proc/'catatan_masalah_disengaja.json',*files]})
    summary['negative_tests']=negatives
    (proc/'ringkasan.json').write_text(json.dumps(summary,indent=2,ensure_ascii=False)+'\n')
    lines=['# Verifikasi proyek `02_survei_kepuasan_pelanggan`','',f'Tanggal verifikasi: {summary["date"][:10]} (Asia/Jakarta). Mesin: {summary["engine"]}; pandas {pd.__version__}, scipy {scipy.__version__}.','Skrip: `python tools/verify.py` dari folder proyek.','','## Hasil recalc per file','','| File | Sheet | Jumlah rumus | Jumlah error | Status recalc | Ukuran (KB) |','|---|---|---:|---:|---|---:|']
    for r in results:lines.append(f'| `{r["file"]}` | {", ".join(r["sheets"])} | {r["formula_count"]} | 0 | {r["status"]} | {r["size_kb"]} |')
    lines += ['','## Rekonsiliasi pandas/scipy vs sel hasil recalc','','| Pemeriksaan | Lokasi sel | Nilai Python | Nilai sel | Selisih | Status |','|---|---|---:|---:|---:|---|']
    for r in rows:lines.append(f'| {r["pemeriksaan"]} | {r["lokasi"]} | {r["python"]:.12g} | {r["workbook"]:.12g} | {r["selisih"]} | PASS |')
    lines+=['','## Cakupan dan batas verifikasi','',f'- Saya mencocokkan {cell_count:,} sel bersih (414 responden × 27 variabel), termasuk tanggal, kategori, komentar, skor, dan nilai kosong. Selisih numerik dibulatkan 12 desimal; teks harus sama. Checks memuat 414 status per baris pada Check_Detail.',f'- Seluruh {len(raw)*len(raw.columns):,} input Raw_Data dicocokkan dengan CSV, termasuk spasi dan literal N/A. Raw_Data menyimpan semua input sebagai teks.',f'- {len(rows)} pembanding agregat/statistik PASS, selisih nol setelah pembulatan 10 desimal. Ini toleransi representasi desimal, bukan persentase. Selisih asli tersedia pada rekonsiliasi CSV melalui kedua nilai pembanding.','- Mean, SD, margin dan kedua batas CI setiap dimensi serta kepuasan keseluruhan, alpha, jumlah NPS, statistik/p-value Welch, CORREL, dan koefisien/galat baku LINEST diperiksa.','- Latihan sengaja menyisakan sel kuning kosong dan berstatus BELUM DIISI; contoh baris pertama dihitung ulang. Tidak ada PASS palsu untuk pekerjaan yang belum diisi.','- Fungsi diperiksa terhadap daftar yang diizinkan sebelum dan sesudah recalc. Tidak ada rumus dynamic array, sel error, atau tautan workbook eksternal. Semua sheet memiliki Freeze Panes; setiap file di bawah 5 MB.',f'- Preview berasal dari Report pada salinan sementara XLSX melalui LibreOffice PDF dan pdftoppm; PDF memuat {pages} halaman.','- Sampel non-acak, self-report, skala ordinal, dan tidak adanya hubungan sebab-akibat membatasi inferensi. Keberhasilan recalc tidak memvalidasi rancangan survei.','- PivotTable, Slicer, Power Query, dan Analysis ToolPak belum tertanam atau dijalankan. Interaksi Excel 2016/2019/365 dan impor Google Sheets belum diuji.','- [Rekonsiliasi CSV](data/processed/rekonsiliasi.csv), [ringkasan dan hash input](data/processed/ringkasan.json), [audit masalah data](docs/masalah_kualitas_data.md).','']
    lines+=['## Pengujian negatif','', '| Kerusakan pada salinan | Hasil |','|---|---|']
    lines += [f'| {r["case"]} | {r["status"]} |' for r in negatives]
    lines += ['','Saya menjalankan pengujian ini pada objek salinan; file sumber dan workbook hasil tidak diubah.','']
    (BASE/'VERIFIKASI.md').write_text('\n'.join(lines))
    print(json.dumps(dict(files=results,comparisons=len(rows),clean_cells=cell_count,statistics=stats,preview_pages=pages),ensure_ascii=False,indent=2))


if __name__=='__main__': main()
