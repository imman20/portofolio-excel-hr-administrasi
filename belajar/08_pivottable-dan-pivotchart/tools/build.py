"""Saya menyediakan sumber Table dan pembanding rumus untuk PivotTable manual."""
from pathlib import Path
import sys
import numpy as np
import pandas as pd
from openpyxl.worksheet.table import Table,TableStyleInfo
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT.parent/'tools'))
import belajar_lib as L
import course_runtime as R
PREVIEW_SHEET='Hasil_Pivot_Diharapkan'
PREVIEW_AREA='A1:F25'


def main():
    g=np.random.default_rng(108);n=240
    d=pd.DataFrame(dict(transaksi_id=[f'D{i:04}' for i in range(1,n+1)],tanggal=pd.to_datetime(g.choice(pd.date_range('2024-01-01','2024-06-30'),n)).date,wilayah=g.choice(['Jawa','Sumatra','Bali'],n),kategori=g.choice(['ATK','Kemasan','Peralatan'],n),qty=g.integers(1,12,n),harga_satuan=g.integers(10,76,n)*1000))
    L.simpan_csv(d,ROOT/'latihan/data_08.csv')
    for sol in [False,True]:
        wb,ws=R.new(ROOT,'Modul 08 — PivotTable manual',d,sol,formats={'tanggal':'dd/mm/yyyy','harga_satuan':'#,##0'})
        data=wb['Data'];L.tulis_header(data,1,7,['nilai'])
        for i in range(2,n+2):L.sel_rumus(data,f'G{i}',f'=E{i}*F{i}','#,##0')
        table=Table(displayName='tblDistribusi',ref='A1:G241');table.tableStyleInfo=TableStyleInfo(name='TableStyleMedium2',showRowStripes=True);data.add_table(table)
        p=wb.create_sheet(PREVIEW_SHEET);L.judul(p,'Pembanding PivotTable — hasil rumus biasa',sub='Saya membuat PivotTable, PivotChart, Slicer, dan Timeline secara manual; objek tersebut belum tertanam.')
        L.tulis_header(p,4,1,['Wilayah','ATK','Kemasan','Peralatan','Total'])
        for i,region in enumerate(['Jawa','Sumatra','Bali'],5):
            p.cell(i,1,region)
            for c in 'BCD':L.sel_rumus(p,f'{c}{i}',f'=SUMIFS(Data!$G$2:$G$241,Data!$C$2:$C$241,$A{i},Data!$D$2:$D$241,{c}$4)','#,##0')
            L.sel_rumus(p,f'E{i}',f'=SUM(B{i}:D{i})','#,##0')
        p['A8']='TOTAL'
        for c in 'BCDE':L.sel_rumus(p,f'{c}8',f'=SUM({c}5:{c}7)','#,##0')
        L.tulis_header(p,11,1,['Wilayah','Kontribusi grand total'])
        for i,j in enumerate(range(5,8),12):L.sel_rumus(p,f'A{i}',f'=A{j}');L.sel_rumus(p,f'B{i}',f'=IF($E$8=0,0,E{j}/$E$8)','0.0%')
        L.tulis_header(p,18,1,['Awal bulan','Nilai','Qty','Jumlah transaksi'])
        for i in range(19,25):
            L.sel_rumus(p,f'A{i}','=DATE(2024,1,1)' if i==19 else f'=EDATE(A{i-1},1)','dd/mm/yyyy')
            for c,source in [('B','G'),('C','E')]:L.sel_rumus(p,f'{c}{i}',f'=SUMIFS(Data!${source}$2:${source}$241,Data!$B$2:$B$241,">="&A{i},Data!$B$2:$B$241,"<"&EDATE(A{i},1))','#,##0')
            L.sel_rumus(p,f'D{i}',f'=COUNTIFS(Data!$B$2:$B$241,">="&A{i},Data!$B$2:$B$241,"<"&EDATE(A{i},1))')
        for c in 'ABCDE':p.column_dimensions[c].width=24
        items=[('Grand total','=Hasil_Pivot_Diharapkan!E8'),('Kontribusi Jawa','=Hasil_Pivot_Diharapkan!B12'),('Jawa kategori ATK','=Hasil_Pivot_Diharapkan!B5'),('Qty Januari','=Hasil_Pivot_Diharapkan!C19'),('Jumlah transaksi Jawa','=COUNTIFS(Data!C2:C241,"Jawa")'),('Nilai Januari','=Hasil_Pivot_Diharapkan!B19'),('Total kontribusi','=SUM(Hasil_Pivot_Diharapkan!B12:B14)'),('Selisih total wilayah','=SUM(Hasil_Pivot_Diharapkan!E5:E7)-SUM(Data!G2:G241)')]
        for i,(lab,f) in enumerate(items,4):R.answer(ws,i,lab,f,sol,'0.0%' if i in [5,10] else '#,##0')
        L.simpan(R.finish(wb),ROOT/'latihan'/f'{"solusi" if sol else "latihan"}_08.xlsx')
    text=['# Kunci jawaban modul 08','','Saya mencocokkan PivotTable manual dengan agregasi pandas dan sheet pembanding berumus, bukan mengklaim objek pivot sudah ada.','','| Pemeriksaan | Sheet / sel | Nilai |','|---|---|---:|']
    for x in checks():text.append(f'| {x["label"]} | {x["sheet"]}!{x["cell"]} | {x["expected"]} |')
    text+=['','Rows wilayah, Columns kategori, Values Sum of nilai memberi matriks pertama. Show Values As % of Grand Total pada pivot wilayah tanpa kategori memberi kontribusi. Group tanggal menurut Years dan Months memberi nilai bulanan. Clear Filter sebelum membandingkan keadaan awal; setelah penambahan sumber lakukan Refresh.','']
    (ROOT/'latihan/kunci_jawaban.md').write_text('\n'.join(text))


def checks():
    d=pd.read_csv(ROOT/'latihan/data_08.csv',parse_dates=['tanggal']);d['nilai']=d.qty*d.harga_satuan;total=d.nilai.sum();j=d.wilayah.eq('Jawa');jan=d.tanggal.dt.month.eq(1)
    vals=[total,d.loc[j,'nilai'].sum()/total,d.loc[j&d.kategori.eq('ATK'),'nilai'].sum(),d.loc[jan,'qty'].sum(),j.sum(),d.loc[jan,'nilai'].sum(),1,0]
    out=[R.check('Ringkasan '+str(i),f'C{i}',float(v)) for i,v in enumerate(vals,4)]
    for i,reg in enumerate(['Jawa','Sumatra','Bali'],5):
        for c,cat in zip('BCD',['ATK','Kemasan','Peralatan']):out.append(R.check(reg+' '+cat,f'{c}{i}',int(d.loc[d.wilayah.eq(reg)&d.kategori.eq(cat),'nilai'].sum()),PREVIEW_SHEET))
    for i,month in enumerate(range(1,7),19):out.append(R.check('Bulan '+str(month),f'B{i}',int(d.loc[d.tanggal.dt.month.eq(month),'nilai'].sum()),PREVIEW_SHEET))
    return out


if __name__=='__main__':main()
