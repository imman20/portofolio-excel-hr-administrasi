"""Saya memisahkan input, proses, dan output serta membuat template terproteksi."""
from pathlib import Path
import datetime as dt
import sys
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Protection
from openpyxl.worksheet.datavalidation import DataValidation
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT.parent/'tools'))
import belajar_lib as L
import course_runtime as R
PREVIEW_SHEET='Report';PREVIEW_AREA='A1:E17'


def main():
    g=np.random.default_rng(115);n=60;dates=pd.date_range('2024-02-01',periods=n)
    d=pd.DataFrame(dict(invoice_id=[f'INV{i:03}' for i in range(1,n+1)],tanggal=dates.date,vendor=g.choice([f'Vendor Simulasi {i:02}' for i in range(1,9)],n),qty=g.integers(1,11,n),harga_satuan=g.integers(5,101,n)*1000,jatuh_tempo=(dates+pd.Timedelta(days=30)).date))
    bills=d.qty*d.harga_satuan;records=[]
    for i in range(55):records.append(dict(payment_id=f'PAY{len(records)+1:03}',invoice_id=d.invoice_id.iloc[i],nominal=int(bills.iloc[i]+50000 if i==0 else bills.iloc[i]*.5 if i==1 else bills.iloc[i]),tanggal=(dates[i]+pd.Timedelta(days=10)).date()))
    records.append(dict(payment_id='PAY056',invoice_id=d.invoice_id.iloc[1],nominal=int(bills.iloc[1]*.25),tanggal=(dates[1]+pd.Timedelta(days=20)).date()))
    records.append(dict(payment_id='PAY057',invoice_id='INV999',nominal=90000,tanggal=dt.date(2024,5,15)))
    pay=pd.DataFrame(records);L.simpan_csv(d,ROOT/'latihan/data_15.csv');L.simpan_csv(pay,ROOT/'latihan/pembayaran_15.csv')
    for kind in ['latihan','solusi','template']:
        sol=kind!='latihan';wb,ws=R.new(ROOT,'Modul 15 — Praktik kerja profesional',d,sol,formats={'tanggal':'dd/mm/yyyy','jatuh_tempo':'dd/mm/yyyy','harga_satuan':'#,##0'})
        payment=wb.create_sheet('Pembayaran');L.tulis_df(payment,pay,formats={'tanggal':'dd/mm/yyyy','nominal':'#,##0'},font=L.INPUT_FONT)
        m=wb.create_sheet('Master');L.tulis_header(m,1,1,['Parameter','Nilai']);m['A2']='Tanggal acuan';L.sel_input(m,'B2',dt.date(2024,6,30),'dd/mm/yyyy')
        cal=wb.create_sheet('Calculation');L.tulis_header(cal,1,1,['invoice_id','nilai_tagihan','pembayaran_dikenal','saldo','status','jatuh_tempo','lewat_tempo','frekuensi_invoice'])
        for i in range(2,62):
            forms={'A':f'=Data!A{i}','B':f'=Data!D{i}*Data!E{i}','C':f'=SUMIFS(Pembayaran!$C$2:$C$58,Pembayaran!$B$2:$B$58,A{i})','D':f'=B{i}-C{i}','E':f'=IF(D{i}<0,"LEBIH BAYAR",IF(D{i}=0,"LUNAS",IF(C{i}>0,"SEBAGIAN","BELUM BAYAR")))','F':f'=Data!F{i}','G':f'=IF(AND(F{i}<Master!$B$2,D{i}>0),"LEWAT TEMPO","LAIN")','H':f'=COUNTIFS(Data!$A$2:$A$61,A{i})'}
            for c,f in forms.items():L.sel_rumus(cal,c+str(i),f,'dd/mm/yyyy' if c=='F' else '#,##0' if c in 'BCD' else None)
        audit=wb.create_sheet('Check_Pembayaran');L.tulis_header(audit,1,1,['payment_id','invoice_id','nominal','status_kunci','nilai_dikenal','nilai_tanpa_kunci','frekuensi_payment'])
        for i in range(2,59):
            forms={'A':f'=Pembayaran!A{i}','B':f'=Pembayaran!B{i}','C':f'=Pembayaran!C{i}','D':f'=IF(COUNTIFS(Data!$A$2:$A$61,B{i})=1,"DIKENAL","PERIKSA KUNCI")','E':f'=IF(D{i}="DIKENAL",C{i},0)','F':f'=IF(D{i}="DIKENAL",0,C{i})','G':f'=COUNTIFS(Pembayaran!$A$2:$A$58,A{i})'}
            for c,f in forms.items():L.sel_rumus(audit,c+str(i),f,'#,##0' if c in 'CEF' else None)
        ck=wb.create_sheet('Checks');L.tulis_header(ck,1,1,['Pemeriksaan','Aktual','Pembanding','Selisih','Status'])
        checks=[('Alokasi pembayaran','=SUM(Calculation!C2:C61)','=SUM(Check_Pembayaran!E2:E58)'),('Klasifikasi pembayaran','=SUM(Check_Pembayaran!E2:F58)','=SUM(Pembayaran!C2:C58)'),('Identitas saldo','=SUM(Calculation!D2:D61)','=SUM(Calculation!B2:B61)-SUM(Calculation!C2:C61)'),('Kunci pembayaran unik','=COUNTIFS(Check_Pembayaran!G2:G58,">1")','=0'),('Kunci invoice unik','=COUNTIFS(Calculation!H2:H61,">1")','=0')]
        for i,(label,a,b) in enumerate(checks,2):ck.cell(i,1,label);L.sel_rumus(ck,f'B{i}',a);L.sel_rumus(ck,f'C{i}',b);L.sel_rumus(ck,f'D{i}',f'=B{i}-C{i}');L.sel_rumus(ck,f'E{i}',f'=IF(D{i}=0,"PASS","FAIL")')
        report=wb.create_sheet('Report');L.judul(report,'Rekonsiliasi tagihan dan pembayaran',sub='Data sintetis | Acuan 30/06/2024 | PASS matematis bukan persetujuan seluruh pengecualian.')
        metrics=[('Nilai tagihan','=SUM(Calculation!B2:B61)'),('Pembayaran dikenal','=SUM(Calculation!C2:C61)'),('Sisa positif','=SUMIFS(Calculation!D2:D61,Calculation!D2:D61,">0")'),('Lebih bayar terpisah','=-SUMIFS(Calculation!D2:D61,Calculation!D2:D61,"<0")'),('Pembayaran tanpa kunci','=SUM(Check_Pembayaran!F2:F58)'),('Saldo neto','=SUM(Calculation!D2:D61)'),('Invoice lewat tempo','=COUNTIFS(Calculation!G2:G61,"LEWAT TEMPO")'),('Pemeriksaan FAIL','=COUNTIFS(Checks!E2:E6,"FAIL")')]
        for i,(label,f) in enumerate(metrics,4):L.label(report,f'A{i}',label);L.sel_rumus(report,f'C{i}',f,'#,##0')
        report['A14']='Saya meninjau pembayaran tanpa kunci dan lebih bayar sebelum tindakan bisnis.';report.column_dimensions['A'].width=38;report.column_dimensions['B'].width=8;report.column_dimensions['C'].width=24
        items=[('Nilai tagihan','=Report!C4'),('Pembayaran dikenal','=Report!C5'),('Sisa positif','=Report!C6'),('Jumlah invoice lebih bayar','=COUNTIFS(Calculation!E2:E61,"LEBIH BAYAR")'),('Pembayaran tanpa kunci','=Report!C8'),('Selisih identitas saldo','=SUM(Calculation!D2:D61)-(SUM(Calculation!B2:B61)-SUM(Calculation!C2:C61))'),('Invoice lewat tempo','=Report!C10')]
        for i,(label,f) in enumerate(items,4):R.answer(ws,i,label,f,sol,'#,##0')
        R.finish(wb)
        if kind=='template':
            for sheet in wb:
                for row in sheet:
                    for c in row:c.protection=Protection(locked=True,hidden=False)
                sheet.protection.sheet=True;sheet.protection.set_password('latihan')
            for sheet,area in [(wb['Data'],'A2:F61'),(payment,'A2:D58'),(m,'B2')]:
                for row in sheet[area] if ':' in area else [(sheet[area],)]:
                    for c in row:c.protection=Protection(locked=False)
            dv=DataValidation(type='whole',operator='between',formula1=1,formula2=10000);dv.showErrorMessage=True;wb['Data'].add_data_validation(dv);dv.add('D2:D61')
        L.simpan(wb,ROOT/'latihan'/f'{kind}_15.xlsx')
    text=['# Kunci jawaban modul 15','','Saya mempertahankan pengecualian bisnis walaupun rekonsiliasi matematis PASS.','','| Pemeriksaan | File / sel | Nilai |','|---|---|---:|']
    for c in checks_reference()[:7]:text.append(f'| {c["label"]} | solusi_15.xlsx:{c["cell"]} | {c["expected"]} |')
    text+=['','Pembayaran invoice yang terpecah adalah hubungan satu-ke-banyak yang sah; payment_id tetap harus unik. Lebih bayar tidak dikurangi diam-diam dari sisa invoice lain, dan pembayaran berkunci asing tidak dialokasikan ke invoice rekaan. Template membuka hanya input yang dinyatakan; password contoh latihan bukan perlindungan kerahasiaan.','']
    (ROOT/'latihan/kunci_jawaban.md').write_text('\n'.join(text))


def checks_reference():
    d=pd.read_csv(ROOT/'latihan/data_15.csv',parse_dates=['tanggal','jatuh_tempo']);p=pd.read_csv(ROOT/'latihan/pembayaran_15.csv',parse_dates=['tanggal']);assert d.invoice_id.is_unique and p.payment_id.is_unique
    assert d.tanggal.le('2024-06-30').all() and p.tanggal.le('2024-06-30').all()
    bills=d.qty*d.harga_satuan;paid=d.invoice_id.map(p.groupby('invoice_id').nominal.sum()).fillna(0);balance=bills-paid;unknown=~p.invoice_id.isin(d.invoice_id);late=d.jatuh_tempo.lt('2024-06-30')&balance.gt(0)
    vals=[bills.sum(),paid.sum(),balance[balance.gt(0)].sum(),balance.lt(0).sum(),p.loc[unknown,'nominal'].sum(),0,late.sum()]
    out=[R.check('Ringkasan '+str(i),f'C{i}',float(v)) for i,v in enumerate(vals,4)]
    for i,v in enumerate(balance,2):out.append(R.check('Saldo invoice '+str(i),f'D{i}',float(v),'Calculation'))
    return out


def checks():
    out=checks_reference()
    for c in checks_reference()[:7]:out.append({**c,'file':'template_15.xlsx'})
    template=load_workbook(ROOT/'latihan/template_15.xlsx')
    assert all(ws.protection.sheet for ws in template)
    assert not template['Data']['A2'].protection.locked and not template['Master']['B2'].protection.locked
    assert template['Calculation']['B2'].protection.locked and template['Calculation']['B2'].data_type=='f'
    return out


if __name__=='__main__':main()
