"""Saya menghitung ulang workbook, memeriksa rumus, dan merekonsiliasi CSV secara independen.

Jalankan dari folder proyek: python tools/verify.py
"""
from __future__ import annotations

import csv
import datetime as dt
import hashlib
import json
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from zoneinfo import ZoneInfo

import pandas as pd
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data/raw"
ALLOWED = set('SUM AVERAGE MIN MAX COUNT COUNTA COUNTBLANK ROUND IF IFS AND OR SUMIFS COUNTIFS AVERAGEIFS MAXIFS MINIFS SUMPRODUCT VLOOKUP INDEX MATCH IFERROR IFNA TRIM CLEAN SUBSTITUTE LEFT RIGHT MID LEN VALUE TEXT DATE DATEVALUE YEAR MONTH DAY EOMONTH EDATE TODAY NETWORKDAYS NETWORKDAYS.INTL DATEDIF RANK.EQ CORREL STDEV.S VAR.S CONFIDENCE.T T.TEST T.INV.2T NORM.S.INV LINEST FORECAST.LINEAR TREND SWITCH TEXTJOIN ABS SQRT PERCENTILE.INC QUARTILE.INC MEDIAN LARGE SMALL'.split())
MONEY = ['omzet_bruto', 'penjualan_bersih', 'hpp', 'laba_kotor']
COLS = dict(zip(MONEY, 'BCDE'))


def reference():
    """Relasi divalidasi sebelum agregasi; satu order adalah satu baris produk."""
    d = {p.stem: pd.read_csv(p, keep_default_na=False) for p in RAW.glob('*.csv')}
    t, r = d['transaksi'].copy(), d['retur']
    for name, key in [('transaksi','order_id'), ('retur','retur_id'), ('master_produk','product_id'),
                      ('master_pelanggan','customer_id'), ('master_wilayah','region_id'), ('master_sales','salesperson_id')]:
        assert d[name][key].is_unique, (name, 'kunci duplikat')
        assert not d[name].eq('').any().any(), (name, 'kosong')
    assert not d['target'].duplicated(['bulan', 'region_id']).any()
    assert len(d['target']) == 12 * len(d['master_wilayah'])
    assert (d['target'].target_penjualan_bersih > 0).all()
    assert set(d['target'].bulan) == {f'2024-{i:02d}' for i in range(1,13)}
    assert set(d['target'].region_id) == set(d['master_wilayah'].region_id)
    for col, table in [('product_id','master_produk'), ('customer_id','master_pelanggan'),
                       ('region_id','master_wilayah'), ('salesperson_id','master_sales')]:
        assert t[col].isin(d[table][col]).all(), col
    assert r.order_id.isin(t.order_id).all()
    assert (r.qty_retur > 0).all()
    t = t.merge(d['master_produk'][['product_id','hpp_satuan','kategori']], validate='many_to_one')
    t = t.merge(d['master_pelanggan'][['customer_id','region_id','tanggal_bergabung']],
                on='customer_id', validate='many_to_one', suffixes=('', '_pelanggan'))
    assert (t.region_id == t.region_id_pelanggan).all()
    sales_region = d['master_sales'].set_index('salesperson_id').region_id
    assert (t.region_id == t.salesperson_id.map(sales_region)).all()
    dates = pd.to_datetime(t.tanggal, format='%Y-%m-%d')
    assert dates.dt.year.eq(2024).all()
    assert (dates >= pd.to_datetime(t.tanggal_bergabung)).all()
    retur_dates = pd.to_datetime(r.tanggal_retur, format='%Y-%m-%d')
    order_dates = pd.to_datetime(r.order_id.map(t.set_index('order_id').tanggal))
    assert (retur_dates >= order_dates).all() and retur_dates.dt.year.eq(2024).all()
    assert (t.qty > 0).all() and t.qty.mod(1).eq(0).all()
    assert (t.harga_satuan > 0).all() and t.diskon_pct.between(0,.25).all()
    t['qty_retur'] = t.order_id.map(r.groupby('order_id').qty_retur.sum()).fillna(0).astype(int)
    assert (t.qty_retur <= t.qty).all()
    t['omzet_bruto'] = t.qty * t.harga_satuan
    t['nilai_diskon'] = (t.omzet_bruto * t.diskon_pct).round(2)
    t['omzet_neto'] = t.omzet_bruto - t.nilai_diskon
    t['nilai_retur'] = (t.qty_retur * t.harga_satuan * (1-t.diskon_pct)).round(2)
    t['penjualan_bersih'] = t.omzet_neto - t.nilai_retur
    t['hpp'] = (t.qty-t.qty_retur) * t.hpp_satuan
    t['laba_kotor'] = t.penjualan_bersih-t.hpp
    t['periode'] = dates.dt.strftime('%Y-%m')
    return d, t


def run(command):
    result = subprocess.run(command, capture_output=True, text=True, timeout=600)
    if result.returncode:
        raise RuntimeError(result.stderr[-2000:] or result.stdout[-2000:])
    return result


def recalc(path):
    with tempfile.TemporaryDirectory(prefix='sales-recalc-') as temp:
        run(['soffice', '--headless', '--convert-to', 'xlsx', '--outdir', temp, str(path)])
        output = Path(temp)/path.name
        if not output.exists():
            raise RuntimeError('LibreOffice tidak menghasilkan workbook: '+path.name)
        shutil.copyfile(output, path)


def preview(path):
    """Salinan sementara menyembunyikan sheet lain; cache dan chart dipertahankan."""
    ns = {'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    with tempfile.TemporaryDirectory(prefix='sales-preview-') as temp:
        source = Path(temp)/'report.xlsx'
        with zipfile.ZipFile(path) as original, zipfile.ZipFile(source,'w',zipfile.ZIP_DEFLATED) as output:
            for item in original.infolist():
                data = original.read(item.filename)
                if item.filename == 'xl/workbook.xml':
                    root = ET.fromstring(data)
                    for index, sheet in enumerate(root.findall('m:sheets/m:sheet',ns)):
                        sheet.set('state', 'visible' if sheet.attrib['name']=='Report' else 'hidden')
                        if sheet.attrib['name']=='Report':
                            for view in root.findall('m:bookViews/m:workbookView',ns):
                                view.set('activeTab', str(index))
                    data = ET.tostring(root,encoding='utf-8',xml_declaration=True)
                output.writestr(item,data)
        run(['soffice','--headless','--convert-to','pdf','--outdir',temp,str(source)])
        pdf=Path(temp)/'report.pdf'
        assert pdf.exists(), 'PDF tidak dibuat'
        assets=ROOT/'assets'
        assets.mkdir(exist_ok=True)
        shutil.copyfile(pdf, assets/'report.pdf')
        run(['pdftoppm','-f','1','-singlefile','-scale-to','2200','-png',str(pdf),str(assets/'report')])


def main():
    d, t = reference()
    files = sorted((ROOT/'workbook').glob('*.xlsx'))
    assert len(files)==2
    results, rows = [], []
    values={}
    for path in files:
        recalc(path)
        is_solution=path.name.startswith('solusi_')
        formulas = load_workbook(path, data_only=False)
        wb = load_workbook(path, data_only=True)
        n=0
        for ws in formulas:
            assert ws.freeze_panes, ws.title
            for row in ws:
                for cell in row:
                    if cell.data_type=='f':
                        n+=1
                        assert isinstance(cell.value,str), (ws.title, cell.coordinate, 'rumus array')
                        expression=re.sub(r'"(?:[^"]|"")*"','""',cell.value)
                        functions=set(re.findall(r'([A-Za-z_][A-Za-z_0-9.]*)\s*\(',expression))
                        assert functions <= ALLOWED, (cell.coordinate, functions-ALLOWED)
                        assert '[' not in expression and not re.search(r'#[A-Za-z]',expression)
                        empty_exercise = (not is_solution and ws.title=='Checks' and cell.column==5
                                          and wb['Checks'].cell(cell.row,3).value is None)
                        assert wb[ws.title][cell.coordinate].value is not None or empty_exercise, (ws.title,cell.coordinate,'cache kosong')
        errors=[f'{ws.title}!{c.coordinate}: {c.value}' for ws in wb for row in ws for c in row if c.data_type=='e']
        assert not errors, errors[:20]
        assert path.stat().st_size < 5_000_000
        is_solution=path.name.startswith('solusi_')
        assert wb.sheetnames == (['Readme'] if is_solution else ['Instruksi'])+['Raw_Data','Retur','Master','Calculation','Report','Checks']
        if is_solution:
            assert len(formulas['Report']._charts)==4
            status=[c.value for row in wb['Checks'].iter_rows(min_col=6,max_col=6,min_row=5) for c in row if c.value]
            assert status and all(s=='PASS' for s in status), status
            values['solution']=wb
        else:
            assert wb['Checks']['F35'].value=='BELUM DIISI'
            assert formulas['Raw_Data']['J3'].value is None
            values['exercise']=wb
        results.append(dict(file=path.name,sheets=wb.sheetnames,formulas=n,errors=len(errors),size_kb=round(path.stat().st_size/1024,1),recalc='success (LibreOffice headless)'))

    def compare(label, sheet, cell, expected, exercise=False, monetary=True):
        actual = values['exercise' if exercise else 'solution'][sheet][cell].value
        assert isinstance(actual,(float,int)), (label,actual)
        difference=round(actual-expected,2) if monetary else round(actual-expected,12)
        assert difference==0, (label,expected,actual,difference)
        rows.append(dict(pemeriksaan=label,lokasi=('latihan' if exercise else 'solusi')+'_sales_profitability.xlsx:'+sheet+'!'+cell,pandas=round(float(expected),6),workbook=round(float(actual),6),selisih=difference,status='PASS'))

    for period, frame, row in [('TAHUN',t,17)]+[(period,frame,i+5) for i,(period,frame) in enumerate(t.groupby('periode'))]:
        for metric in MONEY:
            compare(period+' '+metric,'Calculation',COLS[metric]+str(row),float(frame[metric].sum()))
    compare('Margin tahunan','Calculation','F17',t.laba_kotor.sum()/t.penjualan_bersih.sum(),monetary=False)
    compare('Pencapaian tahunan','Calculation','H17',t.penjualan_bersih.sum()/d['target'].target_penjualan_bersih.sum(),monetary=False)
    region=[]
    for i, record in d['master_wilayah'].iterrows():
        f=t[t.region_id==record.region_id]
        target=d['target'].loc[d['target'].region_id==record.region_id,'target_penjualan_bersih'].sum()
        for metric,col in [('penjualan_bersih','C'),('hpp','D'),('laba_kotor','E')]:
            compare(record.region_id+' '+metric,'Calculation',col+str(i+21),float(f[metric].sum()))
        compare(record.region_id+' pencapaian','Calculation','H'+str(i+21),f.penjualan_bersih.sum()/target,monetary=False)
        region.append(dict(region=record.nama_region,penjualan_bersih=float(f.penjualan_bersih.sum()),target=int(target),pencapaian=float(f.penjualan_bersih.sum()/target)))
    categories=[]
    for i,k in enumerate(d['master_produk'].kategori.drop_duplicates()):
        f=t[t.kategori==k]
        margin=f.laba_kotor.sum()/f.penjualan_bersih.sum()
        compare(k+' penjualan','Calculation','B'+str(i+31),float(f.penjualan_bersih.sum()))
        compare(k+' margin','Calculation','E'+str(i+31),margin,monetary=False)
        categories.append(dict(kategori=k,penjualan_bersih=float(f.penjualan_bersih.sum()),margin=float(margin)))
    customers=(t.groupby('customer_id').penjualan_bersih.sum().reindex(d['master_pelanggan'].customer_id,fill_value=0)
               .reset_index().sort_values(['penjualan_bersih','customer_id'],ascending=[False,True],kind='stable'))
    n_top=round(.2*len(customers))
    top_share=customers.penjualan_bersih.iloc[:n_top].sum()/t.penjualan_bersih.sum()
    compare('Kontribusi 20% pelanggan','Calculation','AU7',top_share,monetary=False)
    for i,row in enumerate(customers.itertuples(index=False),5):
        assert values['solution']['Calculation']['AL'+str(i)].value==row.customer_id
        compare('Pelanggan '+row.customer_id,'Calculation','AO'+str(i),row.penjualan_bersih)
    # Cocokkan setiap baris KPI, bukan hanya agregat yang dapat saling menutup kesalahan.
    by_order=t.set_index('order_id')
    for row in values['solution']['Raw_Data'].iter_rows(min_row=2,values_only=True):
        ref=by_order.loc[row[0]]
        for metric,col in [('omzet_bruto',9),('nilai_diskon',10),('omzet_neto',11),('qty_retur',12),('nilai_retur',13),('penjualan_bersih',14),('hpp_satuan',15),('hpp',16),('laba_kotor',17)]:
            assert round(row[col]-float(ref[metric]),2)==0,(row[0],metric)
    first=t.set_index('order_id').loc[values['exercise']['Raw_Data']['A2'].value]
    for metric,col in [('omzet_bruto','J'),('penjualan_bersih','O'),('hpp','Q'),('laba_kotor','R')]:
        compare('Contoh baris 2 '+metric,'Raw_Data',col+'2',float(first[metric]),exercise=True)
    preview(ROOT/'workbook/solusi_sales_profitability.xlsx')
    out=ROOT/'data/processed'
    out.mkdir(exist_ok=True)
    pd.DataFrame(rows).to_csv(out/'rekonsiliasi.csv',index=False)
    summary=dict(tanggal=dt.datetime.now(ZoneInfo('Asia/Jakarta')).date().isoformat(),transaksi=len(t),retur=len(d['retur']),
                 total={m:float(t[m].sum()) for m in MONEY},margin=float(t.laba_kotor.sum()/t.penjualan_bersih.sum()),
                 pencapaian=float(t.penjualan_bersih.sum()/d['target'].target_penjualan_bersih.sum()),region=region,kategori=categories,
                 pelanggan_teratas=n_top,kontribusi_20pct=float(top_share),files=results,rekonsiliasi=len(rows),
                 hash_csv={p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in sorted(RAW.glob('*.csv'))})
    (out/'ringkasan.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2)+'\n')
    md=['# Verifikasi proyek `01_sales_profitability`','',f'Tanggal verifikasi: {summary["tanggal"]} (Asia/Jakarta). Mesin: LibreOffice headless (recalc) + openpyxl + pandas.',
        'Skrip: `python tools/verify.py` dari folder proyek.','', '## Hasil recalc per file','',
        '| File | Sheet | Jumlah rumus | Jumlah error | Status recalc | Ukuran (KB) |','|---|---|---:|---:|---|---:|']
    for f in results:
        md.append(f'| `{f["file"]}` | {", ".join(f["sheets"])} | {f["formulas"]} | {f["errors"]} | {f["recalc"]} | {f["size_kb"]} |')
    md+=['','## Rekonsiliasi pandas vs sel hasil recalc','',
         '| Pemeriksaan | Lokasi sel | Nilai pandas | Nilai sel | Selisih | Status |','|---|---|---:|---:|---:|---|']
    for r in rows:
        md.append('| '+' | '.join(str(r[k]) for k in ['pemeriksaan','lokasi','pandas','workbook','selisih','status'])+' |')
    md+=['',f'Ringkasan: {len(results)} file, total error = 0, rekonsiliasi gagal = 0. {len(rows)} pembanding agregat/contoh PASS.','',
         '## Catatan','',
         '- Empat KPI direkonsiliasi untuk setahun dan setiap bulan. Region, kategori, urutan Pareto lengkap, dan baris contoh latihan ikut diperiksa.',
         '- Seluruh baris solusi juga dicocokkan dengan pandas untuk omzet, diskon, neto, retur, HPP, serta laba; relasi dan tanggal diperiksa sebelum agregasi.',
         '- Selisih moneter harus nol setelah pembulatan dua desimal; rasio dibandingkan sampai dua belas desimal untuk mengatasi representasi biner. Angka data ini menghasilkan rupiah bulat.',
         '- Latihan sengaja menyisakan sel kuning kosong. Status Checks adalah BELUM DIISI; ini tidak dihitung sebagai error rumus.',
         '- Fungsi diperiksa terhadap daftar yang diizinkan; tidak ada rumus array dinamis atau tautan workbook eksternal.',
         '- Preview berasal dari Report pada salinan sementara workbook, diekspor LibreOffice ke PDF lalu pdftoppm ke PNG. Sheet lain disembunyikan hanya pada salinan render.',
         '- Pemeriksaan ini memakai LibreOfficeDev; tampilan dan interaksi Excel tetap perlu ditinjau saat dibuka di Excel. PivotTable dan Power Query belum tertanam.',
         '- Hasil terstruktur: [rekonsiliasi CSV](data/processed/rekonsiliasi.csv) dan [ringkasan JSON](data/processed/ringkasan.json).','']
    (ROOT/'VERIFIKASI.md').write_text('\n'.join(md),encoding='utf-8')
    print(json.dumps(summary,ensure_ascii=False,indent=2))


if __name__=='__main__':
    main()
