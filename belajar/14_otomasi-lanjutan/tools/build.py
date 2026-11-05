"""Saya menyiapkan permintaan sintetis dan hasil pembanding otomasi manual."""
from pathlib import Path
import sys
import numpy as np
import pandas as pd
from openpyxl.worksheet.table import Table,TableStyleInfo
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT.parent/'tools'))
import belajar_lib as L
import course_runtime as R
PREVIEW_SHEET='Ringkasan';PREVIEW_AREA='A1:E14'


def main():
    g=np.random.default_rng(114);n=80
    d=pd.DataFrame(dict(request_id=[f'R{i:04}' for i in range(1,n+1)],tanggal=pd.date_range('2024-03-01',periods=n).date,bagian=g.choice(['Operasi','Keuangan','Layanan'],n),jenis=g.choice(['ATK','Kemasan','Peralatan'],n),qty=g.integers(1,16,n),harga_satuan=g.integers(5,51,n)*1000,status=g.choice(['BARU','DIPROSES','SELESAI'],n)))
    L.simpan_csv(d,ROOT/'latihan/data_14.csv')
    for sol in [False,True]:
        wb,ws=R.new(ROOT,'Modul 14 — Otomasi lanjutan',d,sol,formats={'tanggal':'dd/mm/yyyy','harga_satuan':'#,##0'})
        src=wb['Data'];L.tulis_header(src,1,8,['nilai','validasi','frekuensi_id'])
        for i in range(2,n+2):
            forms={'H':f'=E{i}*F{i}','I':f'=IF(AND(COUNTA(A{i})=1,E{i}>0,F{i}>=0,OR(G{i}="BARU",G{i}="DIPROSES",G{i}="SELESAI")),"VALID","PERIKSA")','J':f'=COUNTIFS($A$2:$A$81,A{i})'}
            for c,f in forms.items():L.sel_rumus(src,c+str(i),f,'#,##0' if c=='H' else None)
        table=Table(displayName='tblPermintaan',ref='A1:J81');table.tableStyleInfo=TableStyleInfo(name='TableStyleMedium2',showRowStripes=True);src.add_table(table)
        p=wb.create_sheet(PREVIEW_SHEET);L.judul(p,'Pembanding otomasi permintaan',sub='Saya menyediakan kode terpisah; makro, Office Scripts, dan flow belum tertanam atau dieksekusi.')
        L.tulis_header(p,3,1,['Status','Jumlah catatan','Nilai'])
        for i,status in enumerate(['BARU','DIPROSES','SELESAI'],4):
            p.cell(i,1,status);L.sel_rumus(p,f'B{i}',f'=COUNTIFS(Data!$G$2:$G$81,A{i})');L.sel_rumus(p,f'C{i}',f'=SUMIFS(Data!$H$2:$H$81,Data!$G$2:$G$81,A{i})','#,##0')
        p['A9']='Status tidak diubah oleh kode contoh';p['A10']='Eksekusi ulang tidak menambah baris sumber'
        for c in 'ABC':p.column_dimensions[c].width=30
        items=[('Jumlah permintaan','=COUNTA(Data!A2:A81)'),('Nilai seluruh permintaan','=SUM(Data!H2:H81)'),('Jumlah BARU','=COUNTIFS(Data!G2:G81,"BARU")'),('Nilai BARU','=SUMIFS(Data!H2:H81,Data!G2:G81,"BARU")'),('Baris ID duplikat','=COUNTIFS(Data!J2:J81,">1")'),('Selisih jumlah status','=SUM(Ringkasan!B4:B6)-C4'),('Catatan perlu pemeriksaan','=COUNTIFS(Data!I2:I81,"PERIKSA")')]
        for i,(label,f) in enumerate(items,4):R.answer(ws,i,label,f,sol,'#,##0')
        L.simpan(R.finish(wb),ROOT/'latihan'/f'{"solusi" if sol else "latihan"}_14.xlsx')
    text=['# Kunci jawaban modul 14','','Saya memeriksa formula pembanding dengan CSV. Kode host belum dieksekusi di Excel atau Microsoft 365.','','| Pemeriksaan | Sel | Nilai |','|---|---|---:|']
    for c in checks():text.append(f'| {c["label"]} | {c["cell"]} | {c["expected"]} |')
    text+=['','VBA dipakai untuk desktop; Office Scripts untuk lingkungan Microsoft 365; Power Automate untuk trigger, orkestrasi, serta log. Pengulangan kode contoh menulis formula sama tanpa menambah catatan. Sebelum pengembangan yang mengubah status, saya memerlukan kunci idempotensi dan jejak eksekusi.','']
    (ROOT/'latihan/kunci_jawaban.md').write_text('\n'.join(text))


def checks():
    d=pd.read_csv(ROOT/'latihan/data_14.csv');v=d.qty*d.harga_satuan;new=d.status.eq('BARU');assert d.request_id.is_unique
    vals=[len(d),v.sum(),new.sum(),v[new].sum(),0,0,0]
    out=[R.check('Ringkasan '+str(i),f'C{i}',int(x)) for i,x in enumerate(vals,4)]
    for i,status in enumerate(['BARU','DIPROSES','SELESAI'],4):
        mask=d.status.eq(status);out.extend([R.check('Jumlah '+status,f'B{i}',int(mask.sum()),'Ringkasan'),R.check('Nilai '+status,f'C{i}',int(v[mask].sum()),'Ringkasan')])
    return out


if __name__=='__main__':main()
