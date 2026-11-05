"""Saya menyediakan model optimasi linear dan pembanding fitur what-if manual."""
from pathlib import Path
import json,sys
import numpy as np
import pandas as pd
from scipy.optimize import linprog
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT.parent/'tools'))
import belajar_lib as L
import course_runtime as R


def optimum():
    p=pd.read_csv(ROOT/'latihan/data_13.csv');cap=pd.read_csv(ROOT/'latihan/kapasitas_13.csv')
    result=linprog(-p.margin.to_numpy(),A_ub=p[['jam','bahan']].to_numpy().T,b_ub=cap.kapasitas.to_numpy(),bounds=(0,None),method='highs')
    assert result.success,result.message
    return result,p,cap


def main():
    g=np.random.default_rng(113)
    products=pd.DataFrame(dict(produk=['A','B'],margin=[40000,30000],jam=[2,1],bahan=[1,2]));cap=pd.DataFrame(dict(sumber_daya=['Jam','Bahan'],kapasitas=[100,120]));scenarios=pd.DataFrame(dict(nama=['Dasar','Alternatif Harga','Alternatif Volume'],harga=[120000,int(g.choice([110000,130000])),120000],qty=[80,80,int(g.choice([60,100]))]))
    for name,frame in [('data_13',products),('kapasitas_13',cap),('skenario_13',scenarios)]:L.simpan_csv(frame,ROOT/f'latihan/{name}.csv')
    result,_,_=optimum()
    (ROOT/'latihan/solusi_optimal.json').write_text(json.dumps(dict(metode='scipy.optimize.linprog, HiGHS',produk_a=float(result.x[0]),produk_b=float(result.x[1]),laba_maksimum=float(-result.fun),slack=result.ineqlin.residual.tolist(),status=int(result.status),keterangan="Solusi optimal ditemukan"),ensure_ascii=False,indent=2)+'\n')
    for sol in [False,True]:
        wb,ws=R.new(ROOT,'Modul 13 — What-if dan optimasi',products,sol)
        m=wb.create_sheet('Master');L.tulis_df(m,cap,font=L.INPUT_FONT);m['A4']='Toleransi numerik';L.sel_input(m,'B4',.000001);m.column_dimensions['A'].width=28
        k=wb.create_sheet('Kandidat');L.tulis_header(k,1,1,['Titik','Produk A','Produk B','Jam dipakai','Bahan dipakai','Laba layak'])
        cand=[('Asal','=0','=0'),('Sumbu A','=MIN(Master!B2/Data!C2,Master!B3/Data!D2)','=0'),('Sumbu B','=0','=MIN(Master!B2/Data!C3,Master!B3/Data!D3)'),('Perpotongan','=(Master!B2*Data!D3-Master!B3*Data!C3)/(Data!C2*Data!D3-Data!D2*Data!C3)','=(Data!C2*Master!B3-Data!D2*Master!B2)/(Data!C2*Data!D3-Data!D2*Data!C3)')]
        for i,(lab,x,y) in enumerate(cand,2):
            k.cell(i,1,lab)
            forms={'B':x,'C':y,'D':f'=B{i}*Data!C2+C{i}*Data!C3','E':f'=B{i}*Data!D2+C{i}*Data!D3','F':f'=IF(AND(B{i}>=0,C{i}>=0,D{i}<=Master!B2+Master!B4,E{i}<=Master!B3+Master!B4),B{i}*Data!B2+C{i}*Data!B3,"TIDAK LAYAK")'}
            for c,f in forms.items():L.sel_rumus(k,c+str(i),f,'#,##0.000')
        model=wb.create_sheet('Model');L.tulis_header(model,1,1,['Ukuran','Nilai'])
        for i,label in enumerate(['Produk A','Produk B','Laba','Jam dipakai','Bahan dipakai','Sisa jam','Sisa bahan'],2):model.cell(i,1,label)
        for i,c in [(2,'B'),(3,'C')]:
            if sol:L.sel_rumus(model,f'B{i}',f'=INDEX(Kandidat!{c}2:{c}5,MATCH(MAX(Kandidat!F2:F5),Kandidat!F2:F5,0))','#,##0.000')
            else:L.sel_input(model,f'B{i}',0,'#,##0.000')
        for i,f in [(4,'=SUMPRODUCT(B2:B3,Data!B2:B3)'),(5,'=SUMPRODUCT(B2:B3,Data!C2:C3)'),(6,'=SUMPRODUCT(B2:B3,Data!D2:D3)'),(7,'=IF(ABS(Master!B2-B5)<Master!B4,0,Master!B2-B5)'),(8,'=IF(ABS(Master!B3-B6)<Master!B4,0,Master!B3-B6)')]:L.sel_rumus(model,f'B{i}',f,'#,##0.000')
        model.column_dimensions['A'].width=30;model.column_dimensions['B'].width=24
        what=wb.create_sheet('What_If');L.tulis_header(what,1,1,['Parameter','Nilai'])
        for i,(label,v) in enumerate([('Harga',120000),('Biaya variabel per unit',70000),('Biaya tetap',2000000),('Qty untuk Goal Seek',80)],2):what.cell(i,1,label);L.sel_input(what,f'B{i}',v,'#,##0')
        what['A6']='Laba saat ini';L.sel_rumus(what,'B6','=(B2-B3)*B5-B4','#,##0');what['A7']='Target laba';L.sel_input(what,'B7',1500000,'#,##0');what['A8']='Qty target pembanding';L.sel_rumus(what,'B8','=(B7+B4)/(B2-B3)');what['A9']='Laba qty pembanding';L.sel_rumus(what,'B9','=(B2-B3)*B8-B4','#,##0');what.column_dimensions['A'].width=32
        L.tulis_header(what,1,4,['Qty satu variabel','Laba pembanding'])
        for i,q in enumerate([20,40,60,80,100,120],2):L.sel_input(what,f'D{i}',q);L.sel_rumus(what,f'E{i}',f'=($B$2-$B$3)*D{i}-$B$4','#,##0')
        L.sel_rumus(what,'G2','=B6')
        for c,price in zip('HIJ',[100000,120000,140000]):L.sel_input(what,f'{c}2',price)
        for i,q in enumerate([40,80,120],3):
            L.sel_input(what,f'G{i}',q)
            for c in 'HIJ':L.sel_rumus(what,f'{c}{i}',f'=({c}$2-$B$3)*$G{i}-$B$4','#,##0')
        L.tulis_header(what,11,4,['Skenario','Harga','Qty','Laba'])
        for i,row in enumerate(scenarios.itertuples(index=False),12):
            for c,val in zip('DEF',row):L.sel_input(what,f'{c}{i}',val)
            L.sel_rumus(what,f'G{i}',f'=(E{i}-$B$3)*F{i}-$B$4','#,##0')
        items=[('Produk A optimum','=INDEX(Kandidat!B2:B5,MATCH(MAX(Kandidat!F2:F5),Kandidat!F2:F5,0))'),('Produk B optimum','=INDEX(Kandidat!C2:C5,MATCH(MAX(Kandidat!F2:F5),Kandidat!F2:F5,0))'),('Laba optimum','=MAX(Kandidat!F2:F5)'),('Qty Goal Seek pembanding','=What_If!B8'),('Laba pada qty target','=What_If!B9'),('Sensitivitas harga 120000 qty 80','=What_If!I4'),('Jam optimum','=C4*Data!C2+C5*Data!C3'),('Bahan optimum','=C4*Data!D2+C5*Data!D3')]
        for i,(label,f) in enumerate(items,4):R.answer(ws,i,label,f,sol,'#,##0.000')
        L.simpan(R.finish(wb),ROOT/'latihan'/f'{"solusi" if sol else "latihan"}_13.xlsx')
    text=['# Kunci jawaban modul 13','',f'Saya menjalankan scipy.optimize.linprog dengan metode HiGHS. Produk A optimum = {result.x[0]:.9f}, produk B = {result.x[1]:.9f}, dan laba maksimum = {(-result.fun):.9f} rupiah. Solusi ini kontinu; pembulatan unit belum merupakan solusi integer optimal.','','| Pemeriksaan | Sel | Nilai |','|---|---|---:|']
    for c in checks():text.append(f'| {c["label"]} | {c["cell"]} | {c["expected"]} |')
    text+=['','Kedua sumber daya aktif pada solusi kontinu. Goal Seek mengubah satu input untuk mencapai target; Data Table menghitung kombinasi skenario; Scenario Manager menyimpan kelompok input; Solver mengoptimalkan tujuan dengan kendala. File hanya menyediakan rumus pembanding dan model, bukan konfigurasi fitur manual yang sudah tertanam.','']
    (ROOT/'latihan/kunci_jawaban.md').write_text('\n'.join(text))


def checks():
    result,p,cap=optimum();x,y=result.x
    vals=[x,y,-result.fun,(1500000+2000000)/(120000-70000),1500000,(120000-70000)*80-2000000,float(np.dot(result.x,p.jam)),float(np.dot(result.x,p.bahan))]
    out=[R.check('Pembanding '+str(i),f'C{i}',float(v)) for i,v in enumerate(vals,4)]
    out.extend([R.check('Model produk A','B2',float(x),'Model'),R.check('Model produk B','B3',float(y),'Model'),R.check('Model laba','B4',float(-result.fun),'Model'),R.check('Sisa jam','B7',0,'Model'),R.check('Sisa bahan','B8',0,'Model')])
    scenarios=pd.read_csv(ROOT/'latihan/skenario_13.csv')
    for i,row in enumerate(scenarios.itertuples(index=False),12):out.append(R.check('Laba skenario '+row.nama,f'G{i}',int((row.harga-70000)*row.qty-2000000),'What_If'))
    for i,q in enumerate([20,40,60,80,100,120],2):out.append(R.check('Sensitivitas qty '+str(q),f'E{i}',int((120000-70000)*q-2000000),'What_If'))
    for i,q in enumerate([40,80,120],3):
        for c,price in zip('HIJ',[100000,120000,140000]):out.append(R.check('Sensitivitas '+str(price)+' '+str(q),f'{c}{i}',int((price-70000)*q-2000000),'What_If'))
    return out


if __name__=='__main__':main()
