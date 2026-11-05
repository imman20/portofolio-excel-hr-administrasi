"""Saya menyiapkan CSV ganda dan pembanding hasil transformasi Power Query manual."""
from pathlib import Path
import sys
import numpy as np
import pandas as pd
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT.parent/'tools'))
import belajar_lib as L
import course_runtime as R
PREVIEW_SHEET='Hasil_Query_Diharapkan';PREVIEW_AREA='A1:F21'


def main():
    g=np.random.default_rng(110);parts=[]
    master=pd.DataFrame(dict(kode=[f'Q{i:03}' for i in range(1,13)],kategori=['ATK','Kemasan','Peralatan']*4,harga=g.integers(15,76,12)*1000))
    for month,name in [(1,'januari'),(2,'februari')]:
        start=pd.Timestamp(2024,month,1);n=100
        d=pd.DataFrame(dict(order_id=[f'Q{month}{i:03}' for i in range(1,n+1)],tanggal=(start+pd.to_timedelta(g.integers(0,start.days_in_month,n),unit='D')).date,kode=g.choice(master.kode,n),qty=g.integers(1,16,n)))
        L.simpan_csv(d,ROOT/f'latihan/pesanan_{name}.csv');parts.append(d)
    data=pd.concat(parts,ignore_index=True);wide=pd.DataFrame({'wilayah':['Barat','Tengah','Timur'],'2024-01':[8000000,7000000,6000000],'2024-02':[8500000,7500000,6500000]})
    L.simpan_csv(data,ROOT/'latihan/data_10.csv');L.simpan_csv(master,ROOT/'latihan/master_produk.csv');L.simpan_csv(wide,ROOT/'latihan/target_wide.csv')
    joined=data.merge(master,on='kode',validate='many_to_one');joined['nilai']=joined.qty*joined.harga;L.simpan_csv(joined,ROOT/'latihan/hasil_referensi.csv')
    for sol in [False,True]:
        wb,ws=R.new(ROOT,'Modul 10 — Power Query manual',data,sol,formats={'tanggal':'dd/mm/yyyy'})
        m=wb.create_sheet('Master');L.tulis_df(m,master,font=L.INPUT_FONT,formats={'harga':'#,##0'})
        w=wb.create_sheet('Target_Wide');L.tulis_df(w,wide,font=L.INPUT_FONT)
        src=wb['Data'];L.tulis_header(src,1,5,['harga','nilai','kategori'])
        for i in range(2,202):
            L.sel_rumus(src,f'E{i}',f'=INDEX(Master!$C$2:$C$13,MATCH(C{i},Master!$A$2:$A$13,0))','#,##0');L.sel_rumus(src,f'F{i}',f'=D{i}*E{i}','#,##0');L.sel_rumus(src,f'G{i}',f'=INDEX(Master!$B$2:$B$13,MATCH(C{i},Master!$A$2:$A$13,0))')
        p=wb.create_sheet(PREVIEW_SHEET);L.judul(p,'Pembanding Append, Merge, dan Unpivot',sub='Skrip M disediakan terpisah; query belum tertanam dan belum dijalankan pada mesin Power Query.')
        L.tulis_header(p,4,1,['Bulan','Baris','Qty','Nilai'])
        for i,month in enumerate([1,2],5):
            p.cell(i,1,pd.Timestamp(2024,month,1).date()).number_format='dd/mm/yyyy'
            for c,func,source in [('B','COUNTIFS',''),('C','SUMIFS','D'),('D','SUMIFS','F')]:
                prefix=f'Data!${source}$2:${source}$201,' if source else ''
                L.sel_rumus(p,f'{c}{i}',f'={func}({prefix}Data!$B$2:$B$201,">="&A{i},Data!$B$2:$B$201,"<"&EDATE(A{i},1))','#,##0')
        L.tulis_header(p,11,1,['Wilayah','Periode','Target panjang'])
        row=12
        for i in range(2,5):
            for c in ['B','C']:
                for dest,f in [('A',f'=Target_Wide!A{i}'),('B',f'=Target_Wide!{c}1'),('C',f'=Target_Wide!{c}{i}')]:L.sel_rumus(p,f'{dest}{row}',f,'#,##0' if dest=='C' else None)
                row+=1
        p.column_dimensions['A'].width=28;p.column_dimensions['B'].width=24;p.column_dimensions['C'].width=24;p.column_dimensions['D'].width=24
        items=[('Baris hasil Append','=COUNTA(Data!A2:A201)'),('Nilai setelah Merge','=SUM(Data!F2:F201)'),('Nilai Januari','=Hasil_Query_Diharapkan!D5'),('Nilai Februari','=Hasil_Query_Diharapkan!D6'),('Baris hasil Unpivot','=COUNTA(Hasil_Query_Diharapkan!A12:A17)'),('Total target panjang','=SUM(Hasil_Query_Diharapkan!C12:C17)'),('Selisih target long vs wide','=C9-SUM(Target_Wide!B2:C4)')]
        for i,(lab,f) in enumerate(items,4):R.answer(ws,i,lab,f,sol,'#,##0')
        L.simpan(R.finish(wb),ROOT/'latihan'/f'{"solusi" if sol else "latihan"}_10.xlsx')
    keys=['# Kunci jawaban modul 10','','Saya memakai concat, merge many-to-one, dan melt pandas untuk pembanding. Query M belum dijalankan pada mesin Power Query.','','| Pemeriksaan | Sel | Nilai |','|---|---|---:|']
    for c in checks()[:7]:keys.append(f'| {c["label"]} | {c["cell"]} | {c["expected"]} |')
    keys+=['','Append menambah baris; Merge menambah atribut. Master unik menjaga jumlah baris. Unpivot menghasilkan pasangan wilayah dan periode dengan total tetap. Penambahan master duplikat harus ditolak sebelum Expand; skrip M mempunyai pemeriksaan eksplisit untuk kondisi tersebut.','']
    (ROOT/'latihan/kunci_jawaban.md').write_text('\n'.join(keys))


def checks():
    parts=[pd.read_csv(ROOT/f'latihan/pesanan_{name}.csv') for name in ['januari','februari']];d=pd.concat(parts,ignore_index=True);master=pd.read_csv(ROOT/'latihan/master_produk.csv');j=d.merge(master,on='kode',validate='many_to_one');j['nilai']=j.qty*j.harga;wide=pd.read_csv(ROOT/'latihan/target_wide.csv');long=wide.melt(id_vars='wilayah',var_name='periode',value_name='target')
    vals=[len(d),j.nilai.sum(),j.loc[pd.to_datetime(j.tanggal).dt.month.eq(1),'nilai'].sum(),j.loc[pd.to_datetime(j.tanggal).dt.month.eq(2),'nilai'].sum(),len(long),long.target.sum(),0]
    out=[R.check('Ringkasan '+str(i),f'C{i}',int(x)) for i,x in enumerate(vals,4)]
    for i,value in enumerate(j.nilai,2):out.append(R.check('Nilai gabungan '+str(i),f'F{i}',int(value),'Data'))
    return out


if __name__=='__main__':main()
