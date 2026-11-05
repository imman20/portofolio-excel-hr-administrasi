"""Saya membuat latihan formula dari simulasi tagihan jasa operasional."""
from pathlib import Path
import sys
import pandas as pd
import numpy as np
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT.parent/'tools'))
import belajar_lib as L
import course_runtime as R


def generate():
    g=np.random.default_rng(103);n=180
    return pd.DataFrame(dict(tagihan_id=[f'J{i:04}' for i in range(1,n+1)],tanggal=pd.date_range('2024-01-01',periods=n).date,unit=g.choice(['Perawatan','Pengiriman','Penyimpanan'],n),qty=g.integers(1,16,n),tarif=g.choice([12500,27500,45000,85000],n),diskon_pct=g.choice([0,.05,.10],n),catatan=['' if i%5 else 'Periksa bukti' for i in range(n)]))


def build(df,solution):
    wb,ws=R.new(ROOT,'Modul 03 — Formula dasar',df,solution,formats={'tanggal':'dd/mm/yyyy','tarif':'#,##0','diskon_pct':'0.0%'})
    master=wb.create_sheet('Master');L.tulis_header(master,1,1,['Parameter','Nilai']);master['A2']='tarif_pajak_simulasi';L.sel_input(master,'B2',.11,'0.0%')
    data=wb['Data'];L.tulis_header(data,1,8,['bruto','neto','pajak','tagihan'])
    for i in range(2,len(df)+2):
        for c,f in [('H',f'=D{i}*E{i}'),('I',f'=ROUND(H{i}*(1-F{i}),0)'),('J',f'=ROUND(I{i}*Master!$B$2,0)'),('K',f'=I{i}+J{i}')]:L.sel_rumus(data,c+str(i),f,'#,##0')
    end=len(df)+1
    items=[('Total bruto',f'=SUM(Data!H2:H{end})'),('Total neto',f'=SUM(Data!I2:I{end})'),('Total pajak',f'=SUM(Data!J2:J{end})'),('Total tagihan',f'=SUM(Data!K2:K{end})'),('Rata-rata neto',f'=AVERAGE(Data!I2:I{end})'),('Neto terkecil',f'=MIN(Data!I2:I{end})'),('Neto terbesar',f'=MAX(Data!I2:I{end})'),('Banyak nilai neto',f'=COUNT(Data!I2:I{end})'),('ID terisi',f'=COUNTA(Data!A2:A{end})'),('Catatan kosong',f'=COUNTBLANK(Data!G2:G{end})'),('Rata-rata dibulatkan', '=ROUND(C8,0)'),('Urutan operasi tanpa kurung','=2+3*4'),('Urutan operasi dengan kurung','=(2+3)*4')]
    for i,(label,formula) in enumerate(items,4):R.answer(ws,i,label,formula,solution)
    L.tulis_header(ws,20,2,['Tarif input','Hasil referensi relatif','Hasil referensi absolut'])
    for i in range(21,24):
        L.sel_input(ws,'B'+str(i),i-19)
        L.sel_jawab(ws,'C'+str(i),f'=B{i}*Data!E{i-19}' if solution else None)
        L.sel_jawab(ws,'D'+str(i),f'=B{i}*Data!$E$2' if solution else None)
    L.tulis_header(ws,29,2,['qty / tarif',12500,27500,45000])
    for i,q in enumerate([2,5,9],30):
        L.sel_input(ws,f'B{i}',q)
        for c in 'CDE':L.sel_jawab(ws,f'{c}{i}',f'=$B{i}*{c}$29' if solution else None,'#,##0')
    L.label(ws,'B35','Saya membandingkan selisih tagihan dan komponen.',False)
    L.sel_jawab(ws,'C35','=C7-C5-C6' if solution else None)
    return R.finish(wb)


def checks():
    d=pd.read_csv(ROOT/'latihan/data_03.csv',keep_default_na=False)
    bruto=d.qty*d.tarif;neto=np.floor(bruto*(1-d.diskon_pct)+.5);pajak=np.floor(neto*.11+.5)
    expected=[bruto.sum(),neto.sum(),pajak.sum(),(neto+pajak).sum(),neto.mean(),neto.min(),neto.max(),len(d),len(d),d.catatan.eq('').sum(),np.floor(neto.mean()+.5),14,20]
    out=[R.check(label,f'C{i}',float(v)) for i,(label,v) in enumerate(zip(['Total bruto','Total neto','Pajak','Tagihan','Mean neto','Minimum','Maksimum','COUNT','COUNTA','COUNTBLANK','ROUND','Prioritas kali','Prioritas kurung'],expected),4)]
    for i in range(2,len(d)+2):out.append(R.check('Neto baris '+str(i),'I'+str(i),float(neto.iloc[i-2]),'Data'))
    out.extend([R.check('Referensi campuran','E32',9*45000),R.check('Komponen konsisten','C35',0)])
    return out


def main():
    d=generate();L.simpan_csv(d,ROOT/'latihan/data_03.csv')
    for sol in [False,True]:L.simpan(build(d,sol),ROOT/'latihan'/f'{"solusi" if sol else "latihan"}_03.xlsx')
    c=checks();lines=['# Kunci jawaban modul 03','','Saya memakai CSV dan pandas sebagai pembanding, dengan pembulatan setengah ke atas untuk nilai positif.','','| Pertanyaan / pemeriksaan | Sel solusi | Nilai |','|---|---|---:|']
    for x in c[:13]+c[-2:]:lines.append(f'| {x["label"]} | {x["sheet"]}!{x["cell"]} | {x["expected"]} |')
    lines+=['','Referensi relatif menggeser sumber saat disalin; absolut mengunci sumber tarif. Pada matriks, $B mengunci kolom qty dan $29 mengunci baris tarif. ROUND mengubah nilai, sedangkan format hanya mengubah tampilan.','']
    (ROOT/'latihan/kunci_jawaban.md').write_text('\n'.join(lines))


if __name__=='__main__':main()
