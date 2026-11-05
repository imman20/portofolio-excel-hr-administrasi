"""Saya melatih aturan layanan dan agregasi bersyarat melalui data tiket sintetis."""
from pathlib import Path
import sys
import numpy as np
import pandas as pd
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT.parent/'tools'))
import belajar_lib as L
import course_runtime as R


def main():
    g=np.random.default_rng(104);n=220
    d=pd.DataFrame(dict(ticket_id=[f'T{i:04}' for i in range(1,n+1)],tanggal=pd.to_datetime(g.choice(pd.date_range('2024-01-01','2024-03-31'),n)).date,kanal=g.choice(['Email','Telepon','Portal'],n),kategori=g.choice(['Pengiriman','Pengembalian','Teknis'],n),jam=g.integers(1,73,n),biaya=g.integers(5,81,n)*10000,prioritas=g.choice(['Urgent','Normal'],n,p=[.2,.8]),selesai=g.choice(['YA','TIDAK'],n,p=[.85,.15])))
    L.simpan_csv(d,ROOT/'latihan/data_04.csv')
    for sol in [False,True]:
        wb,ws=R.new(ROOT,'Modul 04 — Logika layanan',d,sol,formats={'tanggal':'dd/mm/yyyy','biaya':'#,##0'})
        m=wb.create_sheet('Master');L.tulis_header(m,1,1,['Parameter','Nilai']);L.label(m,'A2','Batas jam layanan');L.sel_input(m,'B2',24);L.label(m,'A3','Biaya prioritas');L.sel_input(m,'B3',500000,'#,##0');m.column_dimensions['A'].width=28
        src=wb['Data'];L.tulis_header(src,1,9,['hasil_layanan','tinjauan_prioritas','kelas_ifs','kelas_if'])
        for i in range(2,n+2):
            f={'I':f'=IF(AND(H{i}="YA",E{i}<=Master!$B$2),"TEPAT","TINJAU")','J':f'=IF(OR(G{i}="Urgent",F{i}>Master!$B$3),"PRIORITAS","BIASA")','K':f'=IFS(E{i}<=8,"CEPAT",E{i}<=Master!$B$2,"SEDANG",E{i}>Master!$B$2,"LAMA")','L':f'=IF(E{i}<=8,"CEPAT",IF(E{i}<=Master!$B$2,"SEDANG","LAMA"))'}
            for c,v in f.items():L.sel_rumus(src,c+str(i),v)
        items=[('Jumlah tiket Email','=COUNTIFS(Data!C2:C221,"Email")'),('Biaya Email selesai','=SUMIFS(Data!F2:F221,Data!C2:C221,"Email",Data!H2:H221,"YA")'),('Rata-rata jam Email','=AVERAGEIFS(Data!E2:E221,Data!C2:C221,"Email")'),('Kategori berawalan Peng','=COUNTIFS(Data!D2:D221,"Peng*")'),('Tiket Januari','=COUNTIFS(Data!B2:B221,">="&DATE(2024,1,1),Data!B2:B221,"<"&DATE(2024,2,1))'),('Selesai dalam batas','=COUNTIFS(Data!H2:H221,"YA",Data!E2:E221,"<=24")'),('Prioritas tanpa hitung ganda','=COUNTIFS(Data!G2:G221,"Urgent")+COUNTIFS(Data!F2:F221,">500000")-COUNTIFS(Data!G2:G221,"Urgent",Data!F2:F221,">500000")'),('Perbedaan IFS dan IF','=SUMPRODUCT(--(Data!K2:K221<>Data!L2:L221))'),('Biaya seluruh tiket','=SUM(Data!F2:F221)'),('Proporsi tepat dari tiket selesai','=IF(COUNTIFS(Data!H2:H221,"YA")=0,0,C9/COUNTIFS(Data!H2:H221,"YA"))'),('Kelompok tanpa pengamatan','=IF(COUNTIFS(Data!C2:C221,"Surat")=0,"TIDAK ADA DATA",AVERAGEIFS(Data!E2:E221,Data!C2:C221,"Surat"))')]
        for i,(lab,f) in enumerate(items,4):R.answer(ws,i,lab,f,sol,'0.0%' if i==13 else '#,##0.00')
        R.answer(ws,18,'Contoh AND tiket pertama','=IF(AND(Data!H2="YA",Data!E2<=Master!$B$2),"TEPAT","TINJAU")',sol,'General')
        R.answer(ws,19,'Contoh OR tiket pertama','=IF(OR(Data!G2="Urgent",Data!F2>Master!$B$3),"PRIORITAS","BIASA")',sol,'General')
        L.simpan(R.finish(wb),ROOT/'latihan'/f'{"solusi" if sol else "latihan"}_04.xlsx')
    c=checks();text=['# Kunci jawaban modul 04','','Saya menghitung ulang syarat dari CSV dengan pandas. AND menyaring irisan, OR memakai gabungan tanpa menghitung overlap dua kali.','','| Pemeriksaan | Sel | Jawaban |','|---|---|---:|']
    for x in c[:11]:text.append(f'| {x["label"]} | {x["cell"]} | {x["expected"]} |')
    text+=['','IFS membaca syarat dari atas; syarat paling sempit ditempatkan lebih dulu. Alternatif IF bersarang menghasilkan kelas sama. Kelompok tanpa pengamatan diberi pesan, bukan rata-rata nol. Wildcard Peng* mencakup Pengiriman dan Pengembalian.','']
    (ROOT/'latihan/kunci_jawaban.md').write_text('\n'.join(text))


def checks():
    d=pd.read_csv(ROOT/'latihan/data_04.csv');email=d.kanal.eq('Email');closed=d.selesai.eq('YA');good=closed&d.jam.le(24)
    vals=[email.sum(),d.loc[email&closed,'biaya'].sum(),d.loc[email,'jam'].mean(),d.kategori.str.startswith('Peng').sum(),pd.to_datetime(d.tanggal).dt.month.eq(1).sum(),good.sum(),(d.prioritas.eq('Urgent')|d.biaya.gt(500000)).sum(),0,d.biaya.sum(),good.sum()/closed.sum(),'TIDAK ADA DATA']
    out=[R.check('Ringkasan '+str(i),f'C{i}',v.item() if hasattr(v,'item') else v) for i,v in enumerate(vals,4)]
    for i,v in enumerate(good,2):out.append(R.check('Aturan baris '+str(i),'I'+str(i),'TEPAT' if v else 'TINJAU','Data'))
    return out


if __name__=='__main__':main()
