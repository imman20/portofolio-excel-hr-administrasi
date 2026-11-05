"""Saya menghitung ulang file dan membandingkan buku stok dengan pandas secara independen."""
from pathlib import Path
import datetime as dt
import hashlib
import json
import re
import shutil
import subprocess
import tempfile
import zipfile
from xml.etree import ElementTree as ET
from zoneinfo import ZoneInfo
import numpy as np
import pandas as pd
from scipy.stats import norm
from openpyxl import load_workbook

ROOT=Path(__file__).resolve().parents[1]
ALLOWED=set('SUM AVERAGE MIN MAX COUNT COUNTA COUNTBLANK ROUND IF IFS AND OR SUMIFS COUNTIFS AVERAGEIFS MAXIFS MINIFS SUMPRODUCT VLOOKUP INDEX MATCH IFERROR IFNA TRIM CLEAN SUBSTITUTE LEFT RIGHT MID LEN VALUE TEXT DATE DATEVALUE YEAR MONTH DAY EOMONTH EDATE TODAY NETWORKDAYS NETWORKDAYS.INTL DATEDIF RANK.EQ CORREL STDEV.S VAR.S CONFIDENCE.T T.TEST T.INV.2T NORM.S.INV LINEST FORECAST.LINEAR TREND SWITCH TEXTJOIN ABS SQRT PERCENTILE.INC QUARTILE.INC MEDIAN LARGE SMALL'.split())


def reference():
    d={p.stem:pd.read_csv(p,keep_default_na=False) for p in (ROOT/'data/raw').glob('*.csv')}
    m,t,po,s=d['master_sku'],d['pergerakan'],d['purchase_order'],d['supplier']
    assert len(m)==60 and len(s)==8 and len(po)==500
    assert 8500<=len(t)<=9500
    for frame,key in [(m,'sku'),(s,'supplier_id'),(t,'movement_id'),(po,'po_id')]:
        assert frame[key].is_unique and frame[key].ne('').all()
    assert m.supplier_id.isin(s.supplier_id).all() and t.sku.isin(m.sku).all() and po.sku.isin(m.sku).all()
    assert (m.hpp_satuan>0).all() and (m.stok_awal>=0).all()
    assert (po.supplier_id==po.sku.map(m.set_index('sku').supplier_id)).all()
    assert set(t.jenis)=={'MASUK','KELUAR','PENYESUAIAN'}
    assert (t.loc[t.jenis.ne('PENYESUAIAN'),'qty']>0).all() and t.qty.mod(1).eq(0).all()
    t['tanggal']=pd.to_datetime(t.tanggal,format='%Y-%m-%d')
    assert t.tanggal.between('2024-01-01','2024-12-31').all()
    assert m.tanggal_stok_awal.eq('2024-01-01').all()
    po['tanggal_pesan']=pd.to_datetime(po.tanggal_pesan,format='%Y-%m-%d')
    po['tanggal_terima']=pd.to_datetime(po.tanggal_terima,format='%Y-%m-%d',errors='coerce')
    done=po[po.status.eq('SELESAI')].copy();opened=po[po.status.eq('TERBUKA')]
    assert len(done)+len(opened)==len(po)
    assert (done.tanggal_terima>=done.tanggal_pesan).all() and done.tanggal_terima.le('2024-12-31').all()
    assert done.qty_pesan.eq(done.qty_terima).all() and opened.qty_terima.eq(0).all() and opened.tanggal_terima.isna().all()
    receipts=t[t.jenis.eq('MASUK')].merge(done,on='po_id',validate='one_to_one',suffixes=('','_po'))
    assert len(receipts)==len(done)
    assert receipts.sku.eq(receipts.sku_po).all() and receipts.qty.eq(receipts.qty_terima).all() and receipts.tanggal.eq(receipts.tanggal_terima).all()
    assert t.loc[t.jenis.ne('MASUK'),'po_id'].eq('').all()
    t['signed']=np.where(t.jenis.eq('KELUAR'),-t.qty,t.qty)
    idx=pd.date_range('2024-01-01','2024-12-31')
    delta=t.pivot_table(index='tanggal',columns='sku',values='signed',aggfunc='sum').reindex(index=idx,columns=m.sku).fillna(0)
    daily=delta.cumsum().add(m.set_index('sku').stok_awal,axis=1)
    assert (daily>=0).all().all()
    demand=t[t.jenis.eq('KELUAR')].pivot_table(index='tanggal',columns='sku',values='qty',aggfunc='sum').reindex(index=idx,columns=m.sku).fillna(0)
    ref=pd.read_csv(ROOT/'data/processed/stok_harian.csv')
    assert len(ref)==366*60 and not ref.duplicated(['tanggal','sku']).any()
    ref['tanggal']=pd.to_datetime(ref.tanggal)
    for field,frame in [('stok_akhir',daily),('keluar',demand),('stockout',daily.eq(0).astype(int))]:
        stored=ref.pivot(index='tanggal',columns='sku',values=field).reindex(index=idx,columns=m.sku)
        assert np.array_equal(stored.to_numpy(),frame.to_numpy()),field
    stockstart=daily.shift(1);stockstart.iloc[0]=m.set_index('sku').stok_awal
    assert np.array_equal(ref.pivot(index='tanggal',columns='sku',values='stok_awal_hari').reindex(index=idx,columns=m.sku),stockstart)
    for field,kind in [('masuk','MASUK'),('penyesuaian','PENYESUAIAN')]:
        expected=t[t.jenis.eq(kind)].pivot_table(index='tanggal',columns='sku',values='qty',aggfunc='sum').reindex(index=idx,columns=m.sku).fillna(0)
        assert np.array_equal(ref.pivot(index='tanggal',columns='sku',values=field).reindex(index=idx,columns=m.sku),expected)
    result=m.set_index('sku')[['nama_sku','supplier_id','hpp_satuan','stok_awal']].copy()
    for name,kind in [('masuk','MASUK'),('keluar','KELUAR'),('penyesuaian','PENYESUAIAN')]: result[name]=t[t.jenis.eq(kind)].groupby('sku').qty.sum().reindex(result.index,fill_value=0)
    result['stok_akhir']=daily.iloc[-1]
    result['permintaan_harian_rata']=demand.iloc[-90:].mean()
    result['sigma_90_hari']=demand.iloc[-90:].std(ddof=1)
    result['hari_persediaan']=result.stok_akhir.div(result.permintaan_harian_rata).replace([np.inf,np.nan],0)
    done['lead']=(done.tanggal_terima-done.tanggal_pesan).dt.days
    result['lead_time']=done.groupby('sku').lead.mean()
    result['safety_stock']=norm.ppf(.95)*result.sigma_90_hari*np.sqrt(result.lead_time)
    result['reorder_point']=result.permintaan_harian_rata*result.lead_time+result.safety_stock
    result['on_order']=opened.assign(outstanding=opened.qty_pesan-opened.qty_terima).groupby('sku').outstanding.sum().reindex(result.index,fill_value=0)
    result['usulan_pesan']=(result.reorder_point+result.permintaan_harian_rata*14-result.stok_akhir-result.on_order).clip(lower=0)
    result['hpp_keluar_tahun']=result.keluar*result.hpp_satuan
    result['rata_stok_harian']=daily.mean()
    result['rata_nilai_persediaan']=result.rata_stok_harian*result.hpp_satuan
    result['inventory_turnover']=result.hpp_keluar_tahun.div(result.rata_nilai_persediaan).replace([np.inf,np.nan],0)
    result['hari_stockout']=daily.eq(0).sum()
    ranked=result.sort_values(['hpp_keluar_tahun','sku'],ascending=[False,True],kind='stable').copy()
    ranked['peringkat_abc']=np.arange(1,61)
    ranked['kontribusi']=ranked.hpp_keluar_tahun/ranked.hpp_keluar_tahun.sum()
    ranked['kumulatif_sebelum']=ranked.kontribusi.cumsum()-ranked.kontribusi
    ranked['abc']=np.where(ranked.kumulatif_sebelum<.8,'A',np.where(ranked.kumulatif_sebelum<.95,'B','C'))
    for c in ['peringkat_abc','kontribusi','kumulatif_sebelum','abc']:result[c]=ranked[c]
    result['status']=np.select([result.stok_akhir.eq(0),result.permintaan_harian_rata.eq(0),result.stok_akhir.le(result.reorder_point),result.hari_persediaan.gt(90)],['HABIS','TANPA PERMINTAAN','RISIKO HABIS','OVERSTOCK'],default='NORMAL')
    result['nilai_stok_akhir']=result.stok_akhir*result.hpp_satuan
    result['peringkat_usulan']=result.usulan_pesan.rank(method='first',ascending=False).astype(int)
    return d,result,daily,demand


def run(command):
    r=subprocess.run(command,capture_output=True,text=True,timeout=600)
    if r.returncode:raise RuntimeError(r.stderr[-1500:] or r.stdout[-1500:])
    return r


def recalc(path):
    with tempfile.TemporaryDirectory(prefix='inventory-recalc-') as temp:
        run(['soffice','--headless','--convert-to','xlsx','--outdir',temp,str(path)])
        output=Path(temp)/path.name
        assert output.exists(),'LibreOffice tidak menghasilkan file'
        shutil.copyfile(output,path)


def preview(path):
    ns={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    with tempfile.TemporaryDirectory(prefix='inventory-preview-') as temp:
        source=Path(temp)/'report.xlsx'
        with zipfile.ZipFile(path) as inp,zipfile.ZipFile(source,'w',zipfile.ZIP_DEFLATED) as out:
            for item in inp.infolist():
                data=inp.read(item.filename)
                if item.filename=='xl/workbook.xml':
                    root=ET.fromstring(data)
                    for i,sheet in enumerate(root.findall('m:sheets/m:sheet',ns)):
                        sheet.set('state','visible' if sheet.get('name')=='Report' else 'hidden')
                        if sheet.get('name')=='Report':
                            for view in root.findall('m:bookViews/m:workbookView',ns):view.set('activeTab',str(i))
                    defined=root.find('m:definedNames',ns)
                    if defined is not None:
                        for name in list(defined):
                            if name.get('name')=='_xlnm.Print_Area' and name.get('localSheetId')!='5':
                                defined.remove(name)
                    data=ET.tostring(root,encoding='utf-8',xml_declaration=True)
                out.writestr(item,data)
        run(['soffice','--headless','--convert-to','pdf','--outdir',temp,str(source)])
        pdf=Path(temp)/'report.pdf';assert pdf.exists()
        shutil.copyfile(pdf,ROOT/'assets/report.pdf')
        info=run(['pdfinfo',str(pdf)]).stdout
        assert re.search(r'Pages:\s+1\b',info),info
        run(['pdftoppm','-f','1','-singlefile','-scale-to','2200','-png',str(pdf),str(ROOT/'assets/report')])


def main():
    d,ref,daily,demand=reference()
    results=[];rows=[];full=[]
    for path in sorted((ROOT/'workbook').glob('*.xlsx')):
        recalc(path)
        formulas=load_workbook(path,data_only=False);wb=load_workbook(path,data_only=True)
        count=0;errors=[]
        for ws in formulas:
            assert ws.freeze_panes,ws.title
            for row in ws:
                for cell in row:
                    if cell.data_type=='f':
                        count+=1
                        assert isinstance(cell.value,str),'Rumus array'
                        expr=re.sub(r'"(?:[^"]|"")*"','""',cell.value).replace('_xlfn.','')
                        functions=set(re.findall(r'([A-Za-z_][A-Za-z_0-9.]*)\s*\(',expr))
                        assert functions<=ALLOWED,(ws.title,cell.coordinate,functions-ALLOWED)
                        assert '[' not in expr
                        assert wb[ws.title][cell.coordinate].value is not None,(ws.title,cell.coordinate,'Cache kosong')
        for ws in wb:
            for row in ws:
                for c in row:
                    if c.data_type=='e':errors.append((ws.title,c.coordinate,c.value))
        assert not errors,errors[:20]
        assert path.stat().st_size<5_000_000
        assert wb.sheetnames==['Readme','Raw_Data','Master','Permintaan_Harian','Calculation','Report','Checks','Purchase_Order']
        solution=path.name.startswith('solusi')
        check=[wb['Checks'].cell(i,5).value for i in range(2,13)]
        assert check==['PASS']*11 if solution else check==['PASS']*10+['BELUM DIISI'],check
        if not solution:
            assert formulas['Calculation']['N3'].value is None and formulas['Calculation']['N2'].data_type=='f'
        results.append(dict(file=path.name,formulas=count,errors=len(errors),size_kb=round(path.stat().st_size/1024,1),recalc='success',sheets=wb.sheetnames))
        ws=wb['Calculation']
        for i,(sku,record) in enumerate(ref.iterrows(),2):
            assert ws.cell(i,1).value==sku
            for j,key in enumerate(ref.columns,2):
                if not solution and i>2 and j in [14,15,17,27,29]:continue
                # Kolom status/peringkat latihan bergantung pada jawaban yang sengaja kosong.
                if not solution and key in ['status','peringkat_usulan']:continue
                actual=ws.cell(i,j).value;expected=record[key]
                if isinstance(expected,str):
                    assert actual==expected,(path.name,sku,key,expected,actual)
                    difference=0
                else:
                    difference=round(float(actual)-float(expected),6)
                    assert difference==0,(path.name,sku,key,expected,actual,difference)
                full.append(dict(file=path.name,sku=sku,kpi=key,pandas=expected,workbook=actual,selisih=difference,status='PASS'))
        # Seluruh kisi harian diperiksa, termasuk hari nol dan tahun kabisat.
        ph=wb['Permintaan_Harian']
        for frame,start in [(demand,2),(daily,63)]:
            actual=np.array([[ph.cell(i,j).value for j in range(start,start+60)] for i in range(2,368)],dtype=float)
            assert np.array_equal(actual,frame.to_numpy()),(path.name,'Kisi harian tidak cocok')
        if solution:
            for c,key in [('I','stok_akhir'),('V','hari_stockout'),('Z','abc')]:
                for i,(sku,r) in enumerate(ref.iterrows(),2):
                    actual=ws[f'{c}{i}'].value;expected=r[key]
                    rows.append(dict(pemeriksaan=sku+' '+key,lokasi='Calculation!'+c+str(i),pandas=expected,workbook=actual,selisih=0,status='PASS'))
            for cell,key,expected in [('C4','Nilai stok akhir',ref.nilai_stok_akhir.sum()),('C5','Rata-rata turnover',ref.inventory_turnover.mean()),('C6','Total hari stockout',ref.hari_stockout.sum()),('C7','Total usulan pesan',ref.usulan_pesan.sum())]:
                actual=wb['Report'][cell].value;diff=round(actual-expected,6);assert diff==0
                rows.append(dict(pemeriksaan=key,lokasi='Report!'+cell,pandas=expected,workbook=actual,selisih=diff,status='PASS'))
            top=ref.sort_values(['usulan_pesan','sku'],ascending=[False,True]).head(5)
            assert [wb['Report'].cell(i,2).value for i in range(18,23)]==top.index.tolist()
            assert len(formulas['Report']._charts)==1
    preview(ROOT/'workbook/solusi_inventory_operations.xlsx')
    out=ROOT/'data/processed'
    pd.DataFrame(rows).to_csv(out/'rekonsiliasi.csv',index=False)
    pd.DataFrame(full).to_csv(out/'rekonsiliasi_detail.csv',index=False)
    ref.to_csv(out/'kpi_referensi.csv')
    summary=dict(tanggal=dt.datetime.now(ZoneInfo('Asia/Jakarta')).date().isoformat(),sku=len(ref),supplier=len(d['supplier']),pergerakan=len(d['pergerakan']),purchase_order=len(d['purchase_order']),status=ref.status.value_counts().to_dict(),abc=ref.abc.value_counts().to_dict(),top5=[dict(sku=sku,usulan_pesan=float(row.usulan_pesan),status=row.status) for sku,row in top.iterrows()],rata_turnover=float(ref.inventory_turnover.mean()),total_stockout=int(ref.hari_stockout.sum()),total_stok=int(ref.stok_akhir.sum()),nilai_stok=int(ref.nilai_stok_akhir.sum()),files=results,rekonsiliasi=len(rows),rekonsiliasi_detail=len(full),sel_harian_diperiksa=366*60*2*2,hash_csv={p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in sorted((ROOT/'data/raw').glob('*.csv'))})
    (out/'ringkasan.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2)+'\n')
    md=['# Verifikasi proyek `03_inventory_operations`','',f'Tanggal verifikasi: {summary["tanggal"]} (Asia/Jakarta). Mesin: LibreOffice headless (recalc) + openpyxl + pandas/scipy.','Skrip: `python tools/verify.py` dari folder proyek.','','## Hasil recalc per file','','| File | Sheet | Jumlah rumus | Jumlah error | Status recalc | Ukuran (KB) |','|---|---|---:|---:|---|---:|']
    for f in results:md.append(f'| `{f["file"]}` | {", ".join(f["sheets"])} | {f["formulas"]} | {f["errors"]} | {f["recalc"]} | {f["size_kb"]} |')
    md+=['','## Rekonsiliasi pandas vs sel hasil recalc','','| Pemeriksaan | Lokasi sel | Nilai pandas | Nilai sel | Selisih | Status |','|---|---|---:|---:|---:|---|']
    for row in rows:md.append('| '+' | '.join(str(row[k]) for k in ['pemeriksaan','lokasi','pandas','workbook','selisih','status'])+' |')
    md+=['',f'Ringkasan: 2 file, total error = 0, rekonsiliasi gagal = 0. {len(rows)} pembanding utama dan {len(full)} pembanding rinci PASS.','','## Catatan','',f'- Sebanyak {summary["sel_harian_diperiksa"]} sel permintaan dan stok harian cocok tepat dengan pandas pada kedua workbook.', '- Saldo, stockout dan ABC diperiksa per SKU, bukan hanya total. KPI lain, peringkat usulan dan urutan lima SKU terbesar ikut dicocokkan.', '- Selisih numerik wajib nol setelah pembulatan enam desimal untuk representasi biner; unit, rupiah, hari stockout, dan label ABC cocok tepat.', '- CSV stok_harian dari generator dicocokkan kembali dengan agregasi buku gerak independen, termasuk saldo awal hari, masuk, keluar dan penyesuaian.', '- Relasi SKU/pemasok/PO, penerimaan satu per PO, tanggal dan stok nonnegatif diperiksa sebelum agregasi.', '- Latihan menyisakan sel kuning; pemeriksaan kelengkapan BELUM DIISI adalah keadaan yang disengaja. Status dan peringkat usulan latihan belum bermakna sebelum diisi.', '- Preview berasal dari Report pada salinan render: LibreOffice ke PDF satu halaman lalu pdftoppm ke PNG.', '- Tidak ada rumus dinamis atau tautan eksternal. Tampilan/interaksi Excel dan Google Sheets belum diuji.', '- Hasil lengkap: [rekonsiliasi](data/processed/rekonsiliasi.csv), [detail](data/processed/rekonsiliasi_detail.csv), [KPI referensi](data/processed/kpi_referensi.csv), [ringkasan dan hash](data/processed/ringkasan.json).','']
    (ROOT/'VERIFIKASI.md').write_text('\n'.join(md))
    print(json.dumps(summary,ensure_ascii=False,indent=2))


if __name__=='__main__':main()
