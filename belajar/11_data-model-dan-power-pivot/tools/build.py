"""Saya menyiapkan model bintang dan pembanding formula untuk measure DAX manual."""
from pathlib import Path
import sys
import numpy as np
import pandas as pd
from openpyxl.worksheet.table import Table,TableStyleInfo
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT.parent/'tools'))
import belajar_lib as L
import course_runtime as R
PREVIEW_SHEET='Hasil_Model_Diharapkan';PREVIEW_AREA='A1:F25'


def main():
    g=np.random.default_rng(111);n=240
    product=pd.DataFrame(dict(product_id=[f'M{i:03}' for i in range(1,13)],produk=[f'Barang Simulasi {i:02}' for i in range(1,13)],kategori=['ATK','Kemasan','Peralatan']*4))
    orderdates=pd.to_datetime(g.choice(pd.date_range('2024-01-01','2024-12-31'),120))
    d=pd.DataFrame(dict(line_id=[f'L{i:04}' for i in range(1,n+1)],order_id=[f'ORD{i//2+1:03}' for i in range(n)],tanggal=[orderdates[i//2].date() for i in range(n)],product_id=g.choice(product.product_id,n),qty=g.integers(1,9,n),harga_satuan=g.integers(10,81,n)*1000))
    dates=pd.date_range('2024-01-01','2024-12-31');calendar=pd.DataFrame(dict(tanggal=dates.date,tahun=dates.year,bulan=dates.month,periode=dates.strftime('%Y-%m')))
    for name,frame in [('data_11',d),('produk_11',product),('kalender_11',calendar)]:L.simpan_csv(frame,ROOT/f'latihan/{name}.csv')
    for sol in [False,True]:
        wb,ws=R.new(ROOT,'Modul 11 — Data Model manual',d,sol,formats={'tanggal':'dd/mm/yyyy','harga_satuan':'#,##0'})
        p=wb.create_sheet('Produk');L.tulis_df(p,product,font=L.INPUT_FONT)
        cal=wb.create_sheet('Kalender');L.tulis_df(cal,calendar[['tanggal']],formats={'tanggal':'dd/mm/yyyy'},font=L.INPUT_FONT);L.tulis_header(cal,1,2,['tahun','bulan','periode'])
        for i in range(2,368):
            for c,f in [('B',f'=YEAR(A{i})'),('C',f'=MONTH(A{i})'),('D',f'=TEXT(A{i},"yyyy-mm")')]:L.sel_rumus(cal,c+str(i),f)
        src=wb['Data'];L.tulis_header(src,1,7,['nilai_pembanding','kategori','pertama_order','pertama_order_kategori'])
        for i in range(2,242):
            forms={'G':f'=E{i}*F{i}','H':f'=INDEX(Produk!$C$2:$C$13,MATCH(D{i},Produk!$A$2:$A$13,0))','I':f'=IF(COUNTIFS($B$2:B{i},B{i})=1,1,0)','J':f'=IF(COUNTIFS($B$2:B{i},B{i},$H$2:H{i},H{i})=1,1,0)'}
            for c,f in forms.items():L.sel_rumus(src,c+str(i),f,'#,##0' if c=='G' else None)
        for sheet,name,area in [(src,'Fakt','A1:J241'),(p,'Produk','A1:C13'),(cal,'Kalender','A1:D367')]:
            table=Table(displayName=name,ref=area);table.tableStyleInfo=TableStyleInfo(name='TableStyleMedium2',showRowStripes=True);sheet.add_table(table)
        out=wb.create_sheet(PREVIEW_SHEET);L.judul(out,'Pembanding model bintang dan DAX',sub='Data Model, relasi, dan measure harus dibuat manual; angka berikut memakai rumus worksheet.')
        L.tulis_header(out,4,1,['Awal bulan','Penjualan','Pesanan unik','Baris fakta'])
        for i in range(5,17):
            L.sel_rumus(out,f'A{i}','=DATE(2024,1,1)' if i==5 else f'=EDATE(A{i-1},1)','dd/mm/yyyy')
            for c,source in [('B','G'),('C','I')]:L.sel_rumus(out,f'{c}{i}',f'=SUMIFS(Data!${source}$2:${source}$241,Data!$C$2:$C$241,">="&A{i},Data!$C$2:$C$241,"<"&EDATE(A{i},1))','#,##0')
            L.sel_rumus(out,f'D{i}',f'=COUNTIFS(Data!$C$2:$C$241,">="&A{i},Data!$C$2:$C$241,"<"&EDATE(A{i},1))')
        L.tulis_header(out,21,1,['Kategori','Penjualan','Pesanan unik kategori'])
        for i,cat in enumerate(['ATK','Kemasan','Peralatan'],22):
            out.cell(i,1,cat);L.sel_rumus(out,f'B{i}',f'=SUMIFS(Data!$G$2:$G$241,Data!$H$2:$H$241,A{i})','#,##0');L.sel_rumus(out,f'C{i}',f'=SUMIFS(Data!$J$2:$J$241,Data!$H$2:$H$241,A{i})')
        for c in 'ABCD':out.column_dimensions[c].width=28
        items=[('Penjualan seluruh fakta','=SUM(Data!G2:G241)'),('Jumlah pesanan unik','=SUM(Data!I2:I241)'),('Jumlah baris fakta','=COUNTA(Data!A2:A241)'),('Nilai per pesanan','=IF(C5=0,0,C4/C5)'),('Penjualan ATK','=Hasil_Model_Diharapkan!B22'),('Rekonsiliasi bulanan','=SUM(Hasil_Model_Diharapkan!B5:B16)-C4'),('Jumlah hari kalender','=COUNT(Kalender!A2:A367)'),('Jumlah distinct kategori dijumlahkan','=SUM(Hasil_Model_Diharapkan!C22:C24)'),('Overlap distinct antarkategori','=C11-C5')]
        for i,(lab,f) in enumerate(items,4):R.answer(ws,i,lab,f,sol,'#,##0.00')
        L.simpan(R.finish(wb),ROOT/'latihan'/f'{"solusi" if sol else "latihan"}_11.xlsx')
    text=['# Kunci jawaban modul 11','','Saya menghitung nilai dengan pandas groupby dan pesanan unik dengan nunique. Model DAX belum dijalankan pada lingkungan ini.','','| Pemeriksaan | Sel | Nilai |','|---|---|---:|']
    for c in checks()[:9]:text.append(f'| {c["label"]} | {c["cell"]} | {c["expected"]} |')
    text+=['','Fakta bergrain baris produk pesanan; satu pesanan dapat mempunyai beberapa baris dan kategori. Jumlah distinct kategori tidak harus menjumlah ke grand total. Dimensi mengalirkan filter ke fakta melalui relasi satu-ke-banyak; calculated column dihitung per baris, measure merespons filter.','']
    (ROOT/'latihan/kunci_jawaban.md').write_text('\n'.join(text))


def checks():
    d=pd.read_csv(ROOT/'latihan/data_11.csv',parse_dates=['tanggal']);p=pd.read_csv(ROOT/'latihan/produk_11.csv');calendar=pd.read_csv(ROOT/'latihan/kalender_11.csv');assert p.product_id.is_unique and calendar.tanggal.is_unique
    assert d.groupby('order_id').tanggal.nunique().eq(1).all()
    j=d.merge(p,on='product_id',validate='many_to_one');j['nilai']=j.qty*j.harga_satuan;unique=j.order_id.nunique();percat=j.groupby('kategori').order_id.nunique();total=j.nilai.sum()
    vals=[total,unique,len(j),total/unique,j.loc[j.kategori.eq('ATK'),'nilai'].sum(),0,len(calendar),percat.sum(),percat.sum()-unique]
    out=[R.check('Ringkasan '+str(i),f'C{i}',float(v)) for i,v in enumerate(vals,4)]
    for i,month in enumerate(range(1,13),5):
        f=j[j.tanggal.dt.month.eq(month)]
        out.extend([R.check('Penjualan bulan '+str(month),f'B{i}',int(f.nilai.sum()),PREVIEW_SHEET),R.check('Pesanan bulan '+str(month),f'C{i}',f.order_id.nunique(),PREVIEW_SHEET)])
    for i,cat in enumerate(['ATK','Kemasan','Peralatan'],22):out.append(R.check('Distinct '+cat,f'C{i}',int(percat[cat]),PREVIEW_SHEET))
    return out


if __name__=='__main__':main()
