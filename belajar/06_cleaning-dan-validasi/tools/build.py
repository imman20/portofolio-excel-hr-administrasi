"""Saya mempertahankan data kotor dan membuat kolom pembersih yang dapat ditelusuri."""
from pathlib import Path
import datetime as dt
import sys
import numpy as np
import pandas as pd
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT.parent/'tools'))
import belajar_lib as L
import course_runtime as R


def main():
    g=np.random.default_rng(106);n=100;amount=g.integers(100,201,n)*1000;amount[5]=1500000
    ids=[f'{i:04}' for i in range(1,101)];ids[-2:]=ids[:2]
    dates=pd.date_range('2024-01-01',periods=n)
    values=[str(x) for x in amount]
    for i in range(n):
        if i%17==0:values[i]='N/A'
        elif i%19==0:values[i]='-'
        elif i%23==0:values[i]=''
        elif i%7==0:values[i]=' '+values[i]+'\u00a0'
    d=pd.DataFrame(dict(catatan_id=ids,nama_raw=[f'  Pegawai\u00a0{i:03}  ' for i in range(1,n+1)],email_raw=[f'  pegawai{i:03}@contoh.test\t' for i in range(1,n+1)],nilai_raw=values,tanggal_raw=[x.strftime('%Y-%m-%d' if i%2==0 else '%d/%m/%Y') for i,x in enumerate(dates)],kode_raw=[f'JKT-{i:03}' for i in range(1,n+1)]))
    L.simpan_csv(d,ROOT/'latihan/data_06.csv')
    for sol in [False,True]:
        wb,ws=R.new(ROOT,'Modul 06 — Cleaning dan validasi',d,sol)
        src=wb['Data'];L.tulis_header(src,1,7,['nama_bersih','nilai_bersih','tanggal_bersih','kota_kode','nomor_kode','segmen_mid','panjang_kode','frekuensi_id','flag_iqr','email_bersih'])
        for i in range(2,102):
            f={'G':f'=TRIM(CLEAN(SUBSTITUTE(B{i}," "," ")))','H':f'=IF(OR(TRIM(D{i})="N/A",TRIM(D{i})="-",TRIM(D{i})=""),"KOSONG",VALUE(TRIM(SUBSTITUTE(D{i}," "," "))))','I':f'=IF(MID(E{i},5,1)="-",DATE(VALUE(LEFT(E{i},4)),VALUE(MID(E{i},6,2)),VALUE(RIGHT(E{i},2))),DATE(VALUE(RIGHT(E{i},4)),VALUE(MID(E{i},4,2)),VALUE(LEFT(E{i},2))))','J':f'=LEFT(F{i},3)','K':f'=RIGHT(F{i},3)','L':f'=MID(F{i},5,3)','M':f'=LEN(F{i})','N':f'=COUNTIFS($A$2:$A$101,A{i})','O':f'=IF(H{i}="KOSONG","KOSONG",IF(OR(H{i}<Latihan!$C$7,H{i}>Latihan!$C$8),"TINJAU","WAJAR"))','P':f'=TRIM(CLEAN(C{i}))'}
            for c,v in f.items():L.sel_rumus(src,c+str(i),v,'dd/mm/yyyy' if c=='I' else '#,##0' if c=='H' else None)
        # Ambang IQR tetap dihitung pada latihan agar kolom flag tidak bergantung pada jawaban kosong.
        stats=[('Kuartil pertama','=QUARTILE.INC(Data!H2:H101,1)'),('Kuartil ketiga','=QUARTILE.INC(Data!H2:H101,3)'),('IQR','=C5-C4'),('Batas bawah','=C4-1.5*C6'),('Batas atas','=C5+1.5*C6')]
        for i,(lab,f) in enumerate(stats,4):R.answer(ws,i,lab,f,True,'#,##0.00')
        items=[('Nilai kosong','=COUNTIFS(Data!H2:H101,"KOSONG")'),('Total nilai numerik','=SUM(Data!H2:H101)'),('Jumlah nilai numerik','=COUNT(Data!H2:H101)'),('Baris ID berulang','=COUNTIFS(Data!N2:N101,">1")'),('Outlier kandidat','=COUNTIFS(Data!O2:O101,"TINJAU")'),('Jumlah catatan','=COUNTA(Data!A2:A101)'),('Rata-rata nilai valid','=AVERAGE(Data!H2:H101)'),('Jumlah ID unik','=SUMPRODUCT(1/Data!N2:N101)')]
        for i,(lab,f) in enumerate(items,9):R.answer(ws,i,lab,f,sol)
        R.answer(ws,20,'DATEVALUE untuk ISO yang dikenali','=DATEVALUE("2024-01-15")',sol,'dd/mm/yyyy')
        R.answer(ws,21,'LEN kode pertama','=LEN(Data!F2)',sol)
        inputs=wb.create_sheet('Input_Validasi');L.tulis_header(inputs,1,1,['ID','Qty integer 1–100','Kategori']);inputs.append(['Uji',10,'ATK']);inputs.column_dimensions['B'].width=25
        for dv,area in [(DataValidation(type='whole',operator='between',formula1=1,formula2=100),'B2:B101'),(DataValidation(type='list',formula1='"ATK,Kemasan,Peralatan"'),'C2:C101')]:
            dv.showErrorMessage=True;dv.errorStyle='stop';dv.error='Saya mengisi nilai sesuai aturan.';inputs.add_data_validation(dv);dv.add(area)
        src.conditional_formatting.add('N2:N101',CellIsRule(operator='greaterThan',formula=['1'],fill=PatternFill('solid',fgColor='FFC7CE')))
        L.simpan(R.finish(wb),ROOT/'latihan'/f'{"solusi" if sol else "latihan"}_06.xlsx')
    text=['# Kunci jawaban modul 06','','Saya tidak menghapus missing values, duplikat, atau outlier tanpa aturan. Pembanding berikut berasal dari CSV dengan penanda kosong dipertahankan.','','| Pemeriksaan | Sel | Nilai |','|---|---|---:|']
    for c in checks()[:13]:text.append(f'| {c["label"]} | {c["cell"]} | {c["expected"]} |')
    text+=['','Nama dibersihkan dengan mengganti spasi nonbreaking sebelum TRIM/CLEAN. Format tanggal ditentukan dari posisinya, bukan ditebak. ID berulang ditandai; outlier IQR adalah kandidat tinjauan, bukan bukti kesalahan. Data Validation menolak ketikan tidak sesuai tetapi penempelan dapat melewati aturan.','']
    (ROOT/'latihan/kunci_jawaban.md').write_text('\n'.join(text))


def checks():
    d=pd.read_csv(ROOT/'latihan/data_06.csv',dtype=str,keep_default_na=False)
    v=pd.to_numeric(d.nilai_raw.str.replace('\u00a0','',regex=False).str.strip(),errors='coerce');q1=v.quantile(.25);q3=v.quantile(.75);iqr=q3-q1
    counts=d.catatan_id.value_counts();flags=v.lt(q1-1.5*iqr)|v.gt(q3+1.5*iqr)
    vals=[q1,q3,iqr,q1-1.5*iqr,q3+1.5*iqr,v.isna().sum(),v.sum(),v.count(),d.catatan_id.map(counts).gt(1).sum(),flags.sum(),len(d),v.mean(),d.catatan_id.nunique()]
    out=[R.check('Pemeriksaan '+str(i),f'C{i}',float(x)) for i,x in enumerate(vals,4)]
    out.append(R.check('DATEVALUE ISO','C20',dt.date(2024,1,15)))
    for i,row in enumerate(d.itertuples(index=False),2):
        name=' '.join(row.nama_raw.replace('\u00a0',' ').split())
        date=pd.to_datetime(row.tanggal_raw,format='%Y-%m-%d' if row.tanggal_raw[4]=='-' else '%d/%m/%Y').date()
        out.extend([R.check('Nama '+str(i),f'G{i}',name,'Data'),R.check('Tanggal '+str(i),f'I{i}',date,'Data'),R.check('Nilai '+str(i),f'H{i}','KOSONG' if pd.isna(v.iloc[i-2]) else float(v.iloc[i-2]),'Data')])
    return out


if __name__=='__main__':main()
