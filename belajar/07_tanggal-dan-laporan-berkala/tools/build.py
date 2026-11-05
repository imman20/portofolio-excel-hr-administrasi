"""Saya membuat laporan kalender dua tahun dan membedakan pembanding nol dari pertumbuhan nol."""
from pathlib import Path
import datetime as dt
from zoneinfo import ZoneInfo
import sys
import numpy as np
import pandas as pd
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT.parent/'tools'))
import belajar_lib as L
import course_runtime as R


def main():
    g=np.random.default_rng(107);rows=[]
    for month in pd.date_range('2023-01-01','2024-12-01',freq='MS'):
        for j in range(20):rows.append(dict(transaksi_id=f'B{len(rows)+1:04}',tanggal=(month+pd.Timedelta(days=int(g.integers(0,month.days_in_month)))).date(),kanal=g.choice(['Toko','Online']),nilai=0 if month==pd.Timestamp('2023-02-01') else int(g.integers(10,101))*10000))
    d=pd.DataFrame(rows);holidays=pd.DataFrame({'tanggal_libur':pd.to_datetime(['2023-01-02','2023-05-01','2023-12-25','2024-01-01','2024-05-01','2024-12-25']).date})
    L.simpan_csv(d,ROOT/'latihan/data_07.csv');L.simpan_csv(holidays,ROOT/'latihan/libur_07.csv')
    for sol in [False,True]:
        wb,ws=R.new(ROOT,'Modul 07 — Tanggal dan laporan berkala',d,sol,formats={'tanggal':'dd/mm/yyyy','nilai':'#,##0'})
        h=wb.create_sheet('Libur');L.tulis_df(h,holidays,formats={'tanggal_libur':'dd/mm/yyyy'})
        m=wb.create_sheet('Master');L.tulis_header(m,1,1,['Parameter','Nilai']);m['A2']='Tanggal acuan historis';L.sel_input(m,'B2',dt.date(2024,12,31),'dd/mm/yyyy');m.column_dimensions['A'].width=30
        mon=wb.create_sheet('Laporan_Bulanan');L.tulis_header(mon,1,1,['awal_bulan','periode','nilai','mom','yoy','hari_kerja','tahun','bulan','hari','hari_kalender'])
        for i in range(2,26):
            f={'A':'=DATE(2023,1,1)' if i==2 else f'=EDATE(A{i-1},1)','B':f'=TEXT(A{i},"yyyy-mm")','C':f'=SUMIFS(Data!$D$2:$D$481,Data!$B$2:$B$481,">="&A{i},Data!$B$2:$B$481,"<"&EDATE(A{i},1))','D':'="BELUM ADA PEMBANDING"' if i==2 else f'=IF(C{i-1}=0,"PEMBANDING NOL",(C{i}-C{i-1})/C{i-1})','E':'="BELUM ADA PEMBANDING"' if i<14 else f'=IF(C{i-12}=0,"PEMBANDING NOL",(C{i}-C{i-12})/C{i-12})','F':f'=NETWORKDAYS.INTL(A{i},EOMONTH(A{i},0),"0000011",Libur!$A$2:$A$7)','G':f'=YEAR(A{i})','H':f'=MONTH(A{i})','I':f'=DAY(A{i})','J':f'=DATEDIF(A{i},EDATE(A{i},1),"d")'}
            for c,v in f.items():L.sel_rumus(mon,c+str(i),v,'dd/mm/yyyy' if c=='A' else '0.0%' if c in 'DE' else '#,##0' if c!='B' else None)
        items=[('Total 2023','=SUM(Laporan_Bulanan!C2:C13)'),('Total 2024','=SUM(Laporan_Bulanan!C14:C25)'),('Pertumbuhan tahunan','=IF(C4=0,"PEMBANDING NOL",(C5-C4)/C4)'),('Hari kerja Januari 2024','=Laporan_Bulanan!F14'),('Akhir Februari kabisat','=EOMONTH(DATE(2024,2,1),0)'),('Satu bulan setelah akhir Januari','=EDATE(DATE(2024,1,31),1)'),('Tahun acuan','=YEAR(Master!B2)'),('Bulan acuan','=MONTH(Master!B2)'),('Hari acuan','=DAY(Master!B2)'),('Tanggal berjalan — dinamis','=TODAY()'),('Hari kerja dengan akhir pekan Jumat–Sabtu','=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,31),"0000110",Libur!A2:A7)'),('Bulan penuh dalam periode acuan','=DATEDIF(DATE(2024,1,1),Master!B2,"m")'),('YoY Februari dengan pembanding nol','=Laporan_Bulanan!E15')]
        for i,(lab,f) in enumerate(items,4):R.answer(ws,i,lab,f,sol,'dd/mm/yyyy' if i in [8,9,13] else '0.0%' if i==6 else '#,##0')
        L.simpan(R.finish(wb),ROOT/'latihan'/f'{"solusi" if sol else "latihan"}_07.xlsx')
    text=['# Kunci jawaban modul 07','','Saya menghitung agregasi bulanan dari tanggal CSV dan memakai kalender libur simulasi, bukan daftar libur resmi.','','| Pemeriksaan | Sel | Nilai |','|---|---|---:|']
    for x in checks()[:13]:text.append(f'| {x["label"]} | {x["cell"]} | {x["expected"]} |')
    text+=['','MoM dan YoY memakai penyebut periode pembanding. Pembanding nol menghasilkan label, bukan angka pertumbuhan. TODAY berubah saat dihitung ulang; laporan historis memakai acuan tetap di Master. Rentang hari kerja inklusif terhadap awal dan akhir, dikurangi akhir pekan serta libur yang jatuh pada hari kerja.','']
    (ROOT/'latihan/kunci_jawaban.md').write_text('\n'.join(text))


def checks():
    d=pd.read_csv(ROOT/'latihan/data_07.csv',parse_dates=['tanggal']);h=pd.read_csv(ROOT/'latihan/libur_07.csv').tanggal_libur.to_numpy(dtype='datetime64[D]')
    sums=d.groupby(d.tanggal.dt.to_period('M')).nilai.sum().reindex(pd.period_range('2023-01','2024-12',freq='M'),fill_value=0)
    a=sums.iloc[:12].sum();b=sums.iloc[12:].sum()
    business=lambda mask:int(np.busday_count('2024-01-01','2024-02-01',weekmask=mask,holidays=h))
    vals=[int(a),int(b),float((b-a)/a),business('1111100'),pd.Timestamp('2024-02-01').to_period('M').end_time.date(),(pd.Timestamp('2024-01-31')+pd.DateOffset(months=1)).date(),2024,12,31,dt.datetime.now(ZoneInfo('Asia/Jakarta')).date(),business('1111001'),11,'PEMBANDING NOL']
    out=[R.check('Ringkasan '+str(i),f'C{i}',x) for i,x in enumerate(vals,4)]
    for i,(p,value) in enumerate(sums.items(),2):
        out.append(R.check('Nilai '+str(p),f'C{i}',int(value),'Laporan_Bulanan'))
        if i>=3:
            prev=sums.iloc[i-3];expected='PEMBANDING NOL' if prev==0 else float((value-prev)/prev)
            out.append(R.check('MoM '+str(p),f'D{i}',expected,'Laporan_Bulanan'))
        if i>=14:
            prev=sums.iloc[i-14];expected='PEMBANDING NOL' if prev==0 else float((value-prev)/prev)
            out.append(R.check('YoY '+str(p),f'E{i}',expected,'Laporan_Bulanan'))
    return out


CATATAN_VERIFIKASI=['NETWORKDAYS.INTL disimpan dengan nama fungsi langsung tanpa awalan _xlfn, karena pengujian mesin ini menunjukkan awalan tersebut menghasilkan nama tidak dikenal. Interaksi Excel belum diuji langsung.']

if __name__=='__main__':main()
