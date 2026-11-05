"""Saya membangun dashboard dengan chart native dan ukuran yang dapat ditelusuri."""
from pathlib import Path
import datetime as dt
import sys
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import LineChart,BarChart,Reference
from openpyxl.styles import Font,PatternFill,Alignment
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT.parent/'tools'))
import belajar_lib as L
import course_runtime as R
PREVIEW_SHEET='Report';PREVIEW_AREA='A1:H38'


def main():
    g=np.random.default_rng(109);n=360;price=g.integers(20,151,n)*1000
    d=pd.DataFrame(dict(transaksi_id=[f'V{i:04}' for i in range(1,n+1)],tanggal=pd.to_datetime(g.choice(pd.date_range('2024-01-01','2024-06-30'),n)).date,kategori=g.choice(['ATK','Kemasan','Peralatan'],n),wilayah=g.choice(['Barat','Timur'],n),qty=g.integers(1,11,n),harga_satuan=price,hpp_satuan=np.floor(price*g.uniform(.6,.9,n)).astype(int),diskon_pct=g.choice([0,.05,.1],n)))
    target=pd.DataFrame(dict(bulan=pd.date_range('2024-01-01',periods=6,freq='MS').date,target=[22000000,24000000,26000000,25000000,28000000,30000000]))
    L.simpan_csv(d,ROOT/'latihan/data_09.csv');L.simpan_csv(target,ROOT/'latihan/target_09.csv')
    for sol in [False,True]:
        wb,ws=R.new(ROOT,'Modul 09 — Visualisasi dan dashboard',d,sol,formats={'tanggal':'dd/mm/yyyy','harga_satuan':'#,##0','hpp_satuan':'#,##0','diskon_pct':'0.0%'})
        m=wb.create_sheet('Master');L.tulis_df(m,target,formats={'bulan':'dd/mm/yyyy','target':'#,##0'},font=L.INPUT_FONT);m['D1']='Acuan data';L.sel_input(m,'E2',dt.date(2024,6,30),'dd/mm/yyyy')
        src=wb['Data'];L.tulis_header(src,1,9,['penjualan_bersih','hpp','laba_kotor'])
        for i in range(2,n+2):
            for c,f in [('I',f'=ROUND(E{i}*F{i}*(1-H{i}),0)'),('J',f'=E{i}*G{i}'),('K',f'=I{i}-J{i}')]:L.sel_rumus(src,c+str(i),f,'#,##0')
        cal=wb.create_sheet('Calculation');L.tulis_header(cal,1,1,['Bulan','Penjualan','Target','Laba kotor','Pencapaian','Margin'])
        for i in range(2,8):
            for c,f in [('A',f'=Master!A{i}'),('B',f'=SUMIFS(Data!$I$2:$I$361,Data!$B$2:$B$361,">="&A{i},Data!$B$2:$B$361,"<"&EDATE(A{i},1))'),('C',f'=Master!B{i}'),('D',f'=SUMIFS(Data!$K$2:$K$361,Data!$B$2:$B$361,">="&A{i},Data!$B$2:$B$361,"<"&EDATE(A{i},1))'),('E',f'=IF(C{i}=0,0,B{i}/C{i})'),('F',f'=IF(B{i}=0,0,D{i}/B{i})')]:L.sel_rumus(cal,c+str(i),f,'dd/mm/yyyy' if c=='A' else '0.0%' if c in 'EF' else '#,##0')
        L.tulis_header(cal,11,1,['Kategori','Penjualan','Laba kotor','Margin'])
        for i,cat in enumerate(['ATK','Kemasan','Peralatan'],12):
            cal.cell(i,1,cat)
            L.sel_rumus(cal,f'B{i}',f'=SUMIFS(Data!$I$2:$I$361,Data!$C$2:$C$361,A{i})','#,##0');L.sel_rumus(cal,f'C{i}',f'=SUMIFS(Data!$K$2:$K$361,Data!$C$2:$C$361,A{i})','#,##0');L.sel_rumus(cal,f'D{i}',f'=IF(B{i}=0,0,C{i}/B{i})','0.0%')
        report=wb.create_sheet('Report')
        for c in 'ABCDEFGH':report.column_dimensions[c].width=18
        for row,text in [(1,'DASHBOARD PENJUALAN | SIMULASI'),(2,'="Sumber: data_09.csv dan target_09.csv | Acuan data: "&TEXT(Master!E2,"dd/mm/yyyy")'),(3,'="Tanggal penghitungan ulang: "&TEXT(TODAY(),"dd/mm/yyyy")')]:
            report.merge_cells(start_row=row,start_column=1,end_row=row,end_column=8);report.cell(row,1,text);report.cell(row,1).alignment=Alignment(wrap_text=True);report.row_dimensions[row].height=25
        report['A1'].fill=PatternFill('solid',fgColor=L.COLOR_HEADER);report['A1'].font=Font(name='Arial',bold=True,color='FFFFFF',size=14)
        for cell,label in [('A4','Penjualan (Rp)'),('A5','Laba kotor (Rp)'),('E4','Pencapaian'),('E5','Margin kotor')]:L.label(report,cell,label)
        for cell,f,fmt in [('C4','=SUM(Calculation!B2:B7)','#,##0'),('C5','=SUM(Calculation!D2:D7)','#,##0'),('G4','=IF(SUM(Calculation!C2:C7)=0,0,C4/SUM(Calculation!C2:C7))','0.0%'),('G5','=IF(C4=0,0,C5/C4)','0.0%')]:L.sel_rumus(report,cell,f,fmt,bold=True)
        if sol:
            line=LineChart();line.title='Tren bulanan vs target';line.y_axis.title='Rupiah';line.x_axis.title='Bulan';line.add_data(Reference(cal,min_col=2,max_col=3,min_row=1,max_row=7),titles_from_data=True);line.set_categories(Reference(cal,min_col=1,min_row=2,max_row=7));line.width=12.5;line.height=8;line.legend.position='b';line.smooth=False;line.y_axis.scaling.min=0
            for series in line.series:series.smooth=False;series.marker.symbol='circle';series.marker.size=4
            report.add_chart(line,'A8')
            bar=BarChart();bar.title='Penjualan per kategori';bar.y_axis.title='Rupiah';bar.y_axis.scaling.min=0;bar.add_data(Reference(cal,min_col=2,min_row=11,max_row=14),titles_from_data=True);bar.set_categories(Reference(cal,min_col=1,min_row=12,max_row=14));bar.width=12.5;bar.height=8;bar.legend=None;report.add_chart(bar,'E8')
        L.tulis_header(report,26,1,['Bulan','Penjualan','Target','Pencapaian'])
        for i,j in enumerate(range(2,8),27):
            for c,source in [('A','A'),('B','B'),('C','C'),('D','E')]:L.sel_rumus(report,c+str(i),f'=Calculation!{source}{j}','dd/mm/yyyy' if c=='A' else '0.0%' if c=='D' else '#,##0')
        L.tulis_header(report,26,5,['Kategori','Penjualan','Laba','Margin'])
        for i,j in enumerate(range(12,15),27):
            for c,source in [('E','A'),('F','B'),('G','C'),('H','D')]:L.sel_rumus(report,c+str(i),f'=Calculation!{source}{j}','0.0%' if c=='H' else '#,##0' if c in 'FG' else None)
        notes=[(35,'=IF(Calculation!E7>=1,"Bulan terakhir mencapai target; saya tetap meninjau margin.","Bulan terakhir belum mencapai target; saya meninjau volume dan bauran.")'),(36,'Saya membedakan penjualan dengan laba; kategori terbesar belum tentu bermargin tertinggi.'),(37,'Batasan: data sintetis, tanpa biaya operasi, pajak, retur, atau bukti sebab-akibat.')]
        for i,text in notes:report.merge_cells(start_row=i,start_column=1,end_row=i,end_column=8);report.cell(i,1,text).alignment=Alignment(wrap_text=True);report.row_dimensions[i].height=24
        items=[('Penjualan bersih','=Report!C4'),('Laba kotor','=Report!C5'),('Margin kotor','=Report!G5'),('Pencapaian target','=Report!G4'),('Pencapaian Juni','=Calculation!E7'),('Kategori margin terendah','=INDEX(Calculation!A12:A14,MATCH(MIN(Calculation!D12:D14),Calculation!D12:D14,0))'),('Kategori penjualan terbesar','=INDEX(Calculation!A12:A14,MATCH(MAX(Calculation!B12:B14),Calculation!B12:B14,0))'),('Rekonsiliasi tren vs sumber','=SUM(Calculation!B2:B7)-SUM(Data!I2:I361)')]
        for i,(lab,f) in enumerate(items,4):R.answer(ws,i,lab,f,sol,'0.0%' if i in [6,7,8] else 'General' if i in [9,10] else '#,##0')
        L.simpan(R.finish(wb),ROOT/'latihan'/f'{"solusi" if sol else "latihan"}_09.xlsx')
    lines=['# Kunci jawaban modul 09','','Saya menghitung ukuran dari CSV; margin adalah rasio total laba terhadap total penjualan, bukan rata-rata margin transaksi.','','| Pemeriksaan | Sel | Nilai |','|---|---|---:|']
    for x in checks()[:8]:lines.append(f'| {x["label"]} | {x["cell"]} | {x["expected"]} |')
    lines+=['','Chart garis menyajikan waktu dan target pada satuan sama; batang membandingkan kategori. Saya tidak menyamakan penjualan terbesar dengan margin tertinggi, tidak menambahkan efek tiga dimensi, serta menyatakan data dan tanggal acuannya.','']
    (ROOT/'latihan/kunci_jawaban.md').write_text('\n'.join(lines))


def checks():
    d=pd.read_csv(ROOT/'latihan/data_09.csv',parse_dates=['tanggal']);t=pd.read_csv(ROOT/'latihan/target_09.csv');d['neto']=np.floor(d.qty*d.harga_satuan*(1-d.diskon_pct)+.5);d['laba']=d.neto-d.qty*d.hpp_satuan
    cats=d.groupby('kategori')[['neto','laba']].sum().reindex(['ATK','Kemasan','Peralatan']);margins=cats.laba/cats.neto;month=d.groupby(d.tanggal.dt.month)[['neto','laba']].sum()
    vals=[d.neto.sum(),d.laba.sum(),d.laba.sum()/d.neto.sum(),d.neto.sum()/t.target.sum(),month.loc[6,'neto']/t.target.iloc[5],margins.idxmin(),cats.neto.idxmax(),0]
    out=[R.check('Ringkasan '+str(i),f'C{i}',v.item() if hasattr(v,'item') else v) for i,v in enumerate(vals,4)]
    for i in range(2,8):
        for col,val in [('B',month.loc[i-1,'neto']),('D',month.loc[i-1,'laba']),('E',month.loc[i-1,'neto']/t.target.iloc[i-2]),('F',month.loc[i-1,'laba']/month.loc[i-1,'neto'])]:out.append(R.check('Bulan '+str(i-1)+' '+col,f'{col}{i}',float(val),'Calculation'))
    assert len(load_workbook(ROOT/'latihan/solusi_09.xlsx')['Report']._charts)==2
    return out


if __name__=='__main__':main()
