"""Saya membuat lookup produk dengan kunci unik dan kode yang sengaja tidak dikenal."""
from pathlib import Path
import sys
import numpy as np
import pandas as pd
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT.parent/'tools'))
import belajar_lib as L
import course_runtime as R


def main():
    g=np.random.default_rng(105)
    m=pd.DataFrame(dict(kode=[f'P{i:03}' for i in range(1,21)],produk=[f'Produk Simulasi {i:02}' for i in range(1,21)],harga=g.integers(8,96,20)*1000))
    d=pd.DataFrame(dict(order_id=[f'O{i:04}' for i in range(1,161)],kode=g.choice(m.kode,160),qty=g.integers(1,16,160),tanggal=pd.date_range('2024-02-01',periods=160).date))
    d.loc[:2,'kode']=['P999','P998','P997']
    L.simpan_csv(d,ROOT/'latihan/data_05.csv');L.simpan_csv(m,ROOT/'latihan/master_05.csv')
    for sol in [False,True]:
        wb,ws=R.new(ROOT,'Modul 05 — Lookup produk',d,sol,formats={'tanggal':'dd/mm/yyyy'})
        master=wb.create_sheet('Master');L.tulis_df(master,m,font=L.INPUT_FONT,formats={'harga':'#,##0'})
        src=wb['Data'];L.tulis_header(src,1,5,['harga_index_match','nilai_dikenal','status_kunci','harga_vlookup','jumlah_kunci'])
        for i in range(2,162):
            forms={'E':f'=IFNA(INDEX(Master!$C$2:$C$21,MATCH(B{i},Master!$A$2:$A$21,0)),"TIDAK ADA")','F':f'=IF(E{i}="TIDAK ADA",0,C{i}*E{i})','G':f'=IF(COUNTIFS(Master!$A$2:$A$21,B{i})=1,"VALID","PERIKSA KUNCI")','H':f'=IFNA(VLOOKUP(B{i},Master!$A$2:$C$21,3,0),"TIDAK ADA")','I':f'=COUNTIFS(Master!$A$2:$A$21,B{i})'}
            for c,f in forms.items():L.sel_rumus(src,c+str(i),f,'#,##0' if c in 'EFH' else None)
        audit=wb.create_sheet('Audit_Master');bad=pd.concat([m.iloc[:4],m.iloc[[0]]],ignore_index=True);L.tulis_df(audit,bad);L.tulis_header(audit,1,4,['frekuensi_kunci'])
        for i in range(2,7):L.sel_rumus(audit,f'D{i}',f'=COUNTIFS($A$2:$A$6,A{i})')
        items=[('Nilai pesanan dengan kode dikenal','=SUM(Data!F2:F161)'),('Pesanan valid','=COUNTIFS(Data!G2:G161,"VALID")'),('Pesanan perlu pemeriksaan','=COUNTIFS(Data!G2:G161,"PERIKSA KUNCI")'),('Qty kode tidak dikenal','=SUMIFS(Data!C2:C161,Data!G2:G161,"PERIKSA KUNCI")'),('Perbedaan metode lookup','=SUMPRODUCT(--(Data!E2:E161<>Data!H2:H161))'),('Frekuensi kode master pertama','=COUNTIFS(Master!A2:A21,Master!A2)'),('Baris audit berkunci duplikat','=COUNTIFS(Audit_Master!D2:D6,">1")'),('IFNA untuk kunci tidak ada','=IFNA(MATCH("TIDAK_ADA",Master!A2:A21,0),"KUNCI TIDAK ADA")'),('IFERROR menyembunyikan pembagi nol','=IFERROR(1/0,"GALAT DITUTUP")')]
        for i,(label,f) in enumerate(items,4):R.answer(ws,i,label,f,sol,'General' if i in [11,12] else '#,##0')
        L.label(ws,'B16','Nilai nol untuk kode tak dikenal disertai status; total bukan nilai seluruh pesanan.',False)
        L.simpan(R.finish(wb),ROOT/'latihan'/f'{"solusi" if sol else "latihan"}_05.xlsx')
    keys=['# Kunci jawaban modul 05','','Saya memakai merge many-to-one pandas dan mempertahankan kode tidak dikenal sebagai kasus pemeriksaan.','','| Pemeriksaan | Sel | Jawaban |','|---|---|---:|']
    for x in checks()[:9]:keys.append(f'| {x["label"]} | {x["cell"]} | {x["expected"]} |')
    keys+=['','INDEX/MATCH dan VLOOKUP exact harus cocok. Master produksi unik; Audit_Master sengaja duplikat dan tidak dipakai lookup. IFNA hanya menangani nilai tidak ditemukan; IFERROR juga menutup pembagi nol sehingga tidak tepat dijadikan pembungkus semua rumus.','']
    (ROOT/'latihan/kunci_jawaban.md').write_text('\n'.join(keys))


def checks():
    d=pd.read_csv(ROOT/'latihan/data_05.csv');m=pd.read_csv(ROOT/'latihan/master_05.csv');assert m.kode.is_unique
    j=d.merge(m,on='kode',how='left',validate='many_to_one');known=j.harga.notna();values=(j.qty*j.harga).fillna(0)
    v=[values.sum(),known.sum(),(~known).sum(),j.loc[~known,'qty'].sum(),0,1,2,'KUNCI TIDAK ADA','GALAT DITUTUP']
    out=[R.check('Ringkasan '+str(i),f'C{i}',x.item() if hasattr(x,'item') else x) for i,x in enumerate(v,4)]
    for i,(price,value) in enumerate(zip(j.harga,values),2):
        out.append(R.check('Harga '+str(i),f'E{i}','TIDAK ADA' if pd.isna(price) else int(price),'Data'))
        out.append(R.check('Nilai '+str(i),f'F{i}',int(value),'Data'))
    return out


if __name__=='__main__':main()
