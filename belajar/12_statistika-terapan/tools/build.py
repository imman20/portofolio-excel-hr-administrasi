"""Saya membandingkan statistika worksheet dengan pandas dan scipy pada data sintetis."""
from pathlib import Path
import sys
import numpy as np
import pandas as pd
from scipy import stats
from openpyxl.chart import BarChart,Reference
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT.parent/'tools'))
import belajar_lib as L
import course_runtime as R


def main():
    g=np.random.default_rng(112);n=60;hours=np.round(g.uniform(2,10,n),2);before=np.round(g.normal(55,6,n),2);groups=np.array(['A']*30+['B']*30);after=np.round(np.clip(40+2*hours+.3*before+(groups=='B')*4+g.normal(0,3,n),0,100),2)
    data=pd.DataFrame(dict(catatan_id=[f'S{i:03}' for i in range(1,n+1)],kelompok=groups,jam_latihan=hours,skor_awal=before,skor_akhir=after))
    x=np.arange(1,25);trend=pd.DataFrame(dict(indeks=x,bulan=pd.date_range('2023-01-01',periods=24,freq='MS').date,unit=np.round(200+8*x+8*np.sin(x*np.pi/6)+g.normal(0,4,24)).astype(int)))
    L.simpan_csv(data,ROOT/'latihan/data_12.csv');L.simpan_csv(trend,ROOT/'latihan/tren_12.csv')
    for sol in [False,True]:
        wb,ws=R.new(ROOT,'Modul 12 — Statistika terapan',data,sol)
        forms=[('n skor akhir','=COUNT(Data!E2:E61)'),('Mean skor akhir','=AVERAGE(Data!E2:E61)'),('Median skor akhir','=MEDIAN(Data!E2:E61)'),('Simpangan baku sampel','=STDEV.S(Data!E2:E61)'),('Varians sampel','=VAR.S(Data!E2:E61)'),('Minimum','=MIN(Data!E2:E61)'),('Maksimum','=MAX(Data!E2:E61)'),('Margin CI 95%','=CONFIDENCE.T(0.05,STDEV.S(Data!E2:E61),COUNT(Data!E2:E61))'),('CI batas bawah','=C5-C11'),('CI batas atas','=C5+C11'),('Korelasi jam dan skor','=CORREL(Data!C2:C61,Data!E2:E61)'),('Slope OLS','=INDEX(LINEST(Data!E2:E61,Data!C2:C61,1,1),1,1)'),('Intercept OLS','=INDEX(LINEST(Data!E2:E61,Data!C2:C61,1,1),1,2)'),('R kuadrat OLS','=INDEX(LINEST(Data!E2:E61,Data!C2:C61,1,1),3,1)'),('p Welch dua sisi A vs B','=T.TEST(Data!E2:E31,Data!E32:E61,2,3)'),('p berpasangan awal vs akhir','=T.TEST(Data!D2:D61,Data!E2:E61,2,1)'),('Cohen d B minus A','=(AVERAGE(Data!E32:E61)-AVERAGE(Data!E2:E31))/SQRT(((COUNT(Data!E2:E31)-1)*VAR.S(Data!E2:E31)+(COUNT(Data!E32:E61)-1)*VAR.S(Data!E32:E61))/(COUNT(Data!E2:E31)+COUNT(Data!E32:E61)-2))'),('p equal variance dua sisi','=T.TEST(Data!E2:E31,Data!E32:E61,2,2)'),('FORECAST skor untuk jam input','=FORECAST.LINEAR(8,Data!E2:E61,Data!C2:C61)'),('TREND skor untuk jam input','=INDEX(TREND(Data!E2:E61,Data!C2:C61,8),1,1)'),('Selisih dua prediksi linear','=C22-C23')]
        for i,(lab,f) in enumerate(forms,4):R.answer(ws,i,lab,f,sol,'0.000000' if i in [18,19,21] else '#,##0.000')
        tr=wb.create_sheet('Tren');L.tulis_df(tr,trend,formats={'bulan':'dd/mm/yyyy'});L.tulis_header(tr,1,4,['moving_average_3'])
        for i in range(2,26):L.sel_rumus(tr,f'D{i}','="BELUM TIGA PERIODE"' if i<4 else f'=AVERAGE(C{i-2}:C{i})','#,##0.00')
        R.answer(ws,26,'Ramalan unit periode berikut','=FORECAST.LINEAR(25,Tren!C2:C25,Tren!A2:A25)',sol)
        R.answer(ws,27,'Rata-rata tiga unit terakhir','=AVERAGE(Tren!C23:C25)',sol)
        h=wb.create_sheet('Histogram');L.tulis_header(h,1,1,['Batas bawah','Batas atas','Label bin','Frekuensi'])
        for i,lower in enumerate(range(0,100,10),2):
            h.cell(i,1,lower);h.cell(i,2,lower+10);h.cell(i,3,f'{lower}–{lower+10}')
            op='<=' if lower==90 else '<'
            L.sel_rumus(h,f'D{i}',f'=COUNTIFS(Data!$E$2:$E$61,">="&A{i},Data!$E$2:$E$61,"{op}"&B{i})')
        if sol:
            chart=BarChart();chart.title='Distribusi skor akhir';chart.y_axis.title='Frekuensi';chart.x_axis.title='Interval skor sama lebar';chart.y_axis.scaling.min=0;chart.add_data(Reference(h,min_col=4,min_row=1,max_row=11),titles_from_data=True);chart.set_categories(Reference(h,min_col=3,min_row=2,max_row=11));chart.gapWidth=0;chart.legend=None;h.add_chart(chart,'F2')
        L.simpan(R.finish(wb),ROOT/'latihan'/f'{"solusi" if sol else "latihan"}_12.xlsx')
    text=['# Kunci jawaban modul 12','','Saya memakai pandas/scipy; p-value adalah ilustrasi pada data sintetis dan tidak membuktikan efek nyata.','','| Ukuran | Sel | Nilai |','|---|---|---:|']
    for c in checks()[:23]:text.append(f'| {c["label"]} | {c["cell"]} | {c["expected"]} |')
    text+=['','Welch adalah uji utama perbandingan kelompok independen; paired memakai pengukuran awal/akhir pada ID sama. Equal-variance hanya pembanding metodologis. Cohen d memakai selisih mean B minus A dibagi SD pooled, bukan ukuran efek berpasangan. CI mengasumsikan pengamatan independen; korelasi dan regresi tidak membuktikan sebab-akibat. Histogram memakai interval sama lebar dan interval terakhir menyertakan batas atas. Moving average memakai tiga observasi berjalan, bukan validasi ramalan di luar sampel.','']
    (ROOT/'latihan/kunci_jawaban.md').write_text('\n'.join(text))


def checks():
    d=pd.read_csv(ROOT/'latihan/data_12.csv');t=pd.read_csv(ROOT/'latihan/tren_12.csv');x=d.jam_latihan.to_numpy();y=d.skor_akhir.to_numpy();a=y[:30];b=y[30:];n=len(y);sd=y.std(ddof=1);ci=stats.t.ppf(.975,n-1)*sd/np.sqrt(n);reg=stats.linregress(x,y);pooled=np.sqrt(((len(a)-1)*a.var(ddof=1)+(len(b)-1)*b.var(ddof=1))/(len(a)+len(b)-2));pred=reg.intercept+reg.slope*8
    vals=[n,y.mean(),np.median(y),sd,y.var(ddof=1),y.min(),y.max(),ci,y.mean()-ci,y.mean()+ci,np.corrcoef(x,y)[0,1],reg.slope,reg.intercept,reg.rvalue**2,stats.ttest_ind(a,b,equal_var=False).pvalue,stats.ttest_rel(d.skor_awal,y).pvalue,(b.mean()-a.mean())/pooled,stats.ttest_ind(a,b,equal_var=True).pvalue,pred,pred,0]
    labels=['n','Mean','Median','SD sampel','Varians sampel','Minimum','Maksimum','Margin CI','CI bawah','CI atas','Pearson r','Slope OLS','Intercept OLS','R kuadrat','p Welch','p paired','Cohen d','p equal variance','FORECAST','TREND','Selisih prediksi']
    out=[R.check(label,f'C{i}',float(v)) for i,(label,v) in enumerate(zip(labels,vals),4)]
    regt=stats.linregress(t.indeks,t.unit);out.extend([R.check('Ramalan periode 25','C26',float(regt.intercept+regt.slope*25)),R.check('Moving average terakhir','C27',float(t.unit.tail(3).mean()))])
    for i,lower in enumerate(range(0,100,10),2):out.append(R.check('Bin '+str(lower),f'D{i}',int(((y>=lower)&(y<lower+10 if lower<90 else y<=100)).sum()),'Histogram'))
    for i in range(4,26):out.append(R.check('Moving average baris '+str(i),f'D{i}',float(t.unit.iloc[i-4:i-1].mean()),'Tren'))
    for c in out:
        if c['label'].startswith('p '):c.update(rtol=1e-9,atol=1e-50)
    return out


CATATAN_VERIFIKASI=['P-value diuji juga dengan toleransi relatif 1e-9 dan absolut 1e-50, agar hasil yang sangat kecil tidak dianggap cocok hanya karena pembulatan enam desimal.']

if __name__=='__main__':main()
