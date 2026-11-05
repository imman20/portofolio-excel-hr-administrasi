"""Saya membangun workbook latihan dan solusi inventori dengan rumus hidup."""
from pathlib import Path
import datetime as dt
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as col
from openpyxl.chart import BarChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from openpyxl.workbook.properties import CalcProperties

ROOT=Path(__file__).resolve().parents[1]
BLUE='DDEBF7'
NAVY='17365D'


def put(ws,r,c,v,fmt=None):
    cell=ws.cell(r,c,v)
    cell.font=Font(name='Arial',size=10,color='008000' if isinstance(v,str) and v.startswith('=') and '!' in v else '000000')
    if fmt: cell.number_format=fmt
    return cell


def head(ws,r,labels,start=1):
    for c,label in enumerate(labels,start):
        x=put(ws,r,c,label)
        x.font=Font(name='Arial',bold=True,color='FFFFFF',size=10)
        x.fill=PatternFill('solid',fgColor=NAVY)
        x.alignment=Alignment(wrap_text=True,vertical='center')
    ws.row_dimensions[r].height=34


def raw_table(ws,frame,date_cols=()):
    head(ws,1,list(frame.columns))
    for r,record in enumerate(frame.itertuples(index=False,name=None),2):
        for c,v in enumerate(record,1):
            if frame.columns[c-1] in date_cols and v: v=dt.date.fromisoformat(v)
            put(ws,r,c,v,'dd/mm/yyyy' if frame.columns[c-1] in date_cols else None)
    ws.auto_filter.ref=ws.dimensions


def build(exercise=False):
    data={p.stem:pd.read_csv(p,keep_default_na=False) for p in (ROOT/'data/raw').glob('*.csv')}
    m,t,po,sup=data['master_sku'],data['pergerakan'],data['purchase_order'],data['supplier']
    wb=Workbook();wb.remove(wb.active)
    names=['Readme','Raw_Data','Master','Permintaan_Harian','Calculation','Report','Checks','Purchase_Order']
    for name in names: wb.create_sheet(name)
    for ws in wb:
        ws.freeze_panes='B2';ws.sheet_view.showGridLines=False
        ws.sheet_properties.pageSetUpPr.fitToPage=True
        ws.page_setup.orientation='landscape';ws.page_setup.paperSize=ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth=1;ws.page_setup.fitToHeight=1
        for c in range(1,32): ws.column_dimensions[col(c)].width=18
        ws.print_options.horizontalCentered=True
        ws.sheet_properties.outlinePr.summaryRight=False
    r=wb['Readme'];head(r,1,['Inventory & Operations — '+('latihan' if exercise else 'solusi')])
    notes=[
        'Saya menganalisis simulasi persediaan tahun 2024. Data sintetis; bukan pengalaman kerja.',
        'Alur: Raw_Data + Master + Purchase_Order > Permintaan_Harian > Calculation > Report > Checks.',
        'Input parameter berwarna biru pada Master!B2:B10; data mentah dipertahankan.',
        'Permintaan_Harian: B:BI = keluar harian; BK:DR = stok akhir harian, termasuk hari tanpa transaksi.',
        'Calculation: stok, permintaan 90 hari, lead time, safety stock, ROP, usulan, turnover, stockout, ABC.',
        'Saya memakai HPP tetap per SKU. Turnover memakai rata-rata 366 saldo akhir harian.',
        'Service level adalah probabilitas tidak stockout per siklus dalam pendekatan normal; bukan fill rate.',
        'Keluar tercatat dapat tersensor saat stok habis. Usulan harus ditinjau sebelum pemesanan.',
        'Stockout = stok akhir hari nol; bukan hitungan lost sales atau jam tanpa stok.',
        'ABC berdasarkan HPP keluar tahunan. Urutan tie: SKU menaik; item pelintas batas masuk kelas sebelumnya.',
        'Formulas > Calculation Options > Automatic; Ctrl+Alt+F9. Filter Calculation untuk meninjau SKU.',
        'PivotTable, Slicer, Power Query, Solver dan makro tidak tertanam. Report memakai chart native.',
        'Latihan: sel kuning Calculation!N2:O61 dan Q2:Q61 sengaja kosong; contoh baris 2 tersedia.' if exercise else 'Solusi: semua Checks harus PASS; angka pembanding ada pada data/processed/rekonsiliasi.csv.',
        'Periode acuan 31/12/2024, jendela permintaan 03/10/2024–31/12/2024 (90 hari kalender).',
        'Baca docs/instruksi_pengerjaan.md, docs/definisi_kpi.md, dan VERIFIKASI.md.'
    ]
    for i,n in enumerate(notes,3): put(r,i,1,n)
    r.column_dimensions['A'].width=140;r.print_area='A1:A18'
    ws=wb['Master'];head(ws,1,['Parameter','Nilai','Keterangan'])
    params=[('tanggal_acuan',dt.date(2024,12,31),'Akhir periode simulasi'),('service_level',.95,'Probabilitas per siklus'),('periode_review',14,'Hari kalender'),('periode_permintaan',90,'Hari kalender; minimum 2'),('ambang_overstock_hari',90,'Hari persediaan di atas ambang'),('faktor_risiko',1,'Stok <= ROP × faktor'),('ambang_abc_a',.80,'Kontribusi HPP kumulatif sebelum item'),('ambang_abc_b',.95,'A+B; sisa C'),('awal_tahun',dt.date(2024,1,1),'Awal simulasi; kalender tetap 2024')]
    for i,(name,v,note) in enumerate(params,2):
        put(ws,i,1,name);cell=put(ws,i,2,v,'dd/mm/yyyy' if isinstance(v,dt.date) else ('0.0%' if i in [3,8,9] else '0'))
        cell.fill=PatternFill('solid',fgColor=BLUE);cell.font=Font(name='Arial',color='0000FF',bold=True)
        put(ws,i,3,note)
    for area,kind,lo,hi in [('B3','decimal',.001,.999),('B4','whole',0,366),('B5','whole',2,366),('B6','whole',1,10000),('B7','decimal',0,10),('B8:B9','decimal',0,1)]:
        dv=DataValidation(type=kind,operator='between',formula1=lo,formula2=hi);dv.errorTitle='Input tidak valid';dv.error='Saya memakai nilai dalam rentang parameter.';dv.showErrorMessage=True;dv.errorStyle='stop';ws.add_data_validation(dv);dv.add(area)
    dv=DataValidation(type='date',operator='equal',formula1='DATE(2024,12,31)');dv.showErrorMessage=True;ws.add_data_validation(dv);dv.add('B2')
    dv=DataValidation(type='date',operator='equal',formula1='DATE(2024,1,1)');dv.showErrorMessage=True;ws.add_data_validation(dv);dv.add('B10')
    head(ws,13,list(m.columns))
    for i,rec in enumerate(m.itertuples(index=False,name=None),14):
        for j,v in enumerate(rec,1): put(ws,i,j,dt.date.fromisoformat(v) if j==7 else v,'dd/mm/yyyy' if j==7 else '#,##0' if j==5 else None)
    head(ws,77,list(sup.columns))
    for i,rec in enumerate(sup.itertuples(index=False,name=None),78):
        for j,v in enumerate(rec,1): put(ws,i,j,v)
    ws.column_dimensions['A'].width=28;ws.column_dimensions['C'].width=48
    ws=wb['Raw_Data'];raw_table(ws,t,['tanggal']);head(ws,1,['qty_sig','hpp_satuan','hpp_keluar'],7)
    end=len(t)+1
    for i in range(2,end+1):
        put(ws,i,7,f'=IF(D{i}="KELUAR",-E{i},E{i})')
        put(ws,i,8,f'=INDEX(Master!$E$14:$E$73,MATCH(C{i},Master!$A$14:$A$73,0))','#,##0')
        put(ws,i,9,f'=IF(D{i}="KELUAR",E{i}*H{i},0)','#,##0')
    ws.auto_filter.ref=f'A1:I{end}'
    ws=wb['Purchase_Order'];raw_table(ws,po,['tanggal_pesan','tanggal_terima']);head(ws,1,['lead_time_aktual','on_order'],9)
    for i in range(2,len(po)+2):
        put(ws,i,9,f'=IF(AND(H{i}="SELESAI",E{i}<=Master!$B$2),E{i}-D{i},0)')
        put(ws,i,10,f'=IF(AND(D{i}<=Master!$B$2,H{i}="TERBUKA"),MAX(0,F{i}-G{i}),0)')
    ws=wb['Permintaan_Harian'];head(ws,1,['tanggal']+m.sku.tolist()+['pemisah']+m.sku.tolist())
    for di in range(366):
        i=di+2;put(ws,i,1,'=Master!$B$10' if i==2 else f'=A{i-1}+1','dd/mm/yyyy')
        for si in range(60):
            c=col(si+2);s=col(si+63);prev=f'Master!F{si+14}' if i==2 else f'{s}{i-1}'
            put(ws,i,si+2,f'=SUMIFS(Raw_Data!$E$2:$E${end},Raw_Data!$C$2:$C${end},{c}$1,Raw_Data!$B$2:$B${end},$A{i},Raw_Data!$D$2:$D${end},"KELUAR")')
            put(ws,i,si+63,f'={prev}+SUMIFS(Raw_Data!$G$2:$G${end},Raw_Data!$C$2:$C${end},{s}$1,Raw_Data!$B$2:$B${end},$A{i})')
    ws.auto_filter.ref='A1:BI367'
    ws=wb['Calculation']
    headers=['sku','nama_sku','supplier','hpp_satuan','stok_awal','masuk','keluar','penyesuaian','stok_akhir','permintaan_harian_rata','sigma_90_hari','hari_persediaan','lead_time','safety_stock','reorder_point','on_order','usulan_pesan','hpp_keluar_tahun','rata_stok_harian','rata_nilai_persediaan','inventory_turnover','hari_stockout','peringkat_abc','kontribusi','kumulatif_sebelum','abc','status','nilai_stok_akhir']
    head(ws,1,headers)
    for i in range(2,62):
        si=i-2;c=col(si+2);s=col(si+63);mr=si+14
        formulas={1:f'=Master!A{mr}',2:f'=Master!B{mr}',3:f'=Master!C{mr}',4:f'=Master!E{mr}',5:f'=Master!F{mr}'}
        for cno,kind in [(6,'MASUK'),(7,'KELUAR'),(8,'PENYESUAIAN')]:
            formulas[cno]=f'=SUMIFS(Raw_Data!$E$2:$E${end},Raw_Data!$C$2:$C${end},A{i},Raw_Data!$D$2:$D${end},"{kind}",Raw_Data!$B$2:$B${end},"<="&Master!$B$2)'
        formulas.update({9:f'=E{i}+F{i}-G{i}+H{i}',10:f'=SUM(INDEX(Permintaan_Harian!{c}$2:{c}$367,367-Master!$B$5):Permintaan_Harian!{c}$367)/Master!$B$5',11:f'=_xlfn.STDEV.S(INDEX(Permintaan_Harian!{c}$2:{c}$367,367-Master!$B$5):Permintaan_Harian!{c}$367)',12:f'=IF(J{i}=0,0,I{i}/J{i})',13:f'=IF(COUNTIFS(Purchase_Order!$B$2:$B$501,A{i},Purchase_Order!$H$2:$H$501,"SELESAI")=0,INDEX(Master!$C$78:$C$85,MATCH(C{i},Master!$A$78:$A$85,0)),AVERAGEIFS(Purchase_Order!$I$2:$I$501,Purchase_Order!$B$2:$B$501,A{i},Purchase_Order!$H$2:$H$501,"SELESAI"))',14:f'=_xlfn.NORM.S.INV(Master!$B$3)*K{i}*SQRT(M{i})',15:f'=J{i}*M{i}+N{i}',16:f'=SUMIFS(Purchase_Order!$J$2:$J$501,Purchase_Order!$B$2:$B$501,A{i})',17:f'=MAX(0,O{i}+J{i}*Master!$B$4-I{i}-P{i})',18:f'=G{i}*D{i}',19:f'=AVERAGE(Permintaan_Harian!{s}$2:{s}$367)',20:f'=S{i}*D{i}',21:f'=IF(T{i}=0,0,R{i}/T{i})',22:f'=COUNTIFS(Permintaan_Harian!{s}$2:{s}$367,0)',23:f'=COUNTIFS($R$2:$R$61,">"&R{i})+COUNTIFS($R$2:R{i},R{i})',24:f'=IF(SUM($R$2:$R$61)=0,0,R{i}/SUM($R$2:$R$61))',25:f'=SUMIFS($X$2:$X$61,$W$2:$W$61,"<"&W{i})',26:f'=IF(Y{i}<Master!$B$8,"A",IF(Y{i}<Master!$B$9,"B","C"))',27:f'=IF(I{i}=0,"HABIS",IF(J{i}=0,"TANPA PERMINTAAN",IF(I{i}<=O{i}*Master!$B$7,"RISIKO HABIS",IF(L{i}>Master!$B$6,"OVERSTOCK","NORMAL"))))',28:f'=I{i}*D{i}'})
        for j,f in formulas.items():
            if exercise and i>2 and j in [14,15,17]:
                put(ws,i,j,None).fill=PatternFill('solid',fgColor='FFF2CC')
            else: put(ws,i,j,f,'0.0%' if j in [24,25] else '#,##0.00' if j in [10,11,12,13,14,15,17,21] else '#,##0' if j not in [1,2,3,26,27] else None)
    ws.auto_filter.ref='A1:AB61';ws.freeze_panes='D2'
    ws.conditional_formatting.add('I2:I61',CellIsRule(operator='equal',formula=['0'],fill=PatternFill('solid',fgColor='FFC7CE')))
    ws=wb['Report'];head(ws,1,['INVENTORY & OPERATIONS | 2024']+['']*7)
    put(ws,2,1,'Saya menilai stok pada 31/12/2024 | Simulasi | HPP tetap per SKU')
    for i,label,f in [(4,'Nilai stok akhir','=SUM(Calculation!AB2:AB61)'),(5,'Rata-rata turnover SKU','=AVERAGE(Calculation!U2:U61)'),(6,'Total hari stockout SKU','=SUM(Calculation!V2:V61)'),(7,'Usulan pesan (unit teoritis)','=SUM(Calculation!Q2:Q61)')]:
        put(ws,i,1,label);put(ws,i,3,f,'#,##0.00' if i in [5,7] else '#,##0')
    head(ws,9,['Status','Jumlah SKU'])
    statuses=['HABIS','RISIKO HABIS','NORMAL','OVERSTOCK','TANPA PERMINTAAN']
    for i,status in enumerate(statuses,10): put(ws,i,1,status);put(ws,i,2,f'=COUNTIFS(Calculation!$AA$2:$AA$61,A{i})')
    head(ws,17,['Peringkat','SKU','Usulan unit','Stok akhir','On order','ABC','Status'])
    # Urutan deterministik memakai peringkat bantu; tidak memerlukan dynamic array.
    head(wb['Calculation'],1,['peringkat_usulan'],29)
    for i in range(2,62): put(wb['Calculation'],i,29,f'=COUNTIFS($Q$2:$Q$61,">"&Q{i})+COUNTIFS($Q$2:Q{i},Q{i})')
    for i in range(18,23):
        put(ws,i,1,i-17)
        for j,source in [(2,'A'),(3,'Q'),(4,'I'),(5,'P'),(6,'Z'),(7,'AA')]: put(ws,i,j,f'=INDEX(Calculation!${source}$2:${source}$61,MATCH(A{i},Calculation!$AC$2:$AC$61,0))','#,##0.00' if j==3 else None)
    if exercise:
        put(ws,2,1,'LATIHAN: hasil usulan dan status belum bermakna sebelum semua sel kuning terisi.')
        for row in ws.iter_rows(min_row=18,max_row=22,min_col=2,max_col=7):
            for cell in row:
                cell.value='=IF(COUNT(Calculation!$Q$2:$Q$61)=COUNTA(Calculation!$A$2:$A$61),'+cell.value[1:]+',"BELUM DIISI")'
    head(ws,25,['Kelas ABC','Jumlah SKU','Kontribusi HPP'])
    for i,abc in enumerate('ABC',26):
        put(ws,i,1,abc);put(ws,i,2,f'=COUNTIFS(Calculation!$Z$2:$Z$61,A{i})');put(ws,i,3,f'=SUMIFS(Calculation!$X$2:$X$61,Calculation!$Z$2:$Z$61,A{i})','0.0%')
    for i,note in [(31,'Saya memprioritaskan HABIS/RISIKO HABIS; on_order sudah dikurangkan dari usulan.'),(32,'OVERSTOCK: hari persediaan di atas ambang Master. Saya meninjau transfer sebelum pembelian.'),(33,'Usulan berupa unit teoritis pecahan. Saya membulatkan ke kemasan/MOQ setelah tinjauan pemasok.'),(34,'Keluar tercatat dapat lebih kecil daripada permintaan nyata saat stok habis.'),(35,'Normal, independen, lead time tetap: asumsi safety stock; service level bukan jaminan.'),(36,'Status tanpa permintaan memakai indikator tersendiri; hari persediaan nol adalah penanda.')]: put(ws,i,1,note)
    chart=BarChart();chart.title='Jumlah SKU per status';chart.y_axis.title='SKU';chart.add_data(Reference(ws,min_col=2,min_row=9,max_row=14),titles_from_data=True);chart.set_categories(Reference(ws,min_col=1,min_row=10,max_row=14));chart.width=15;chart.height=7;chart.legend=None;ws.add_chart(chart,'D4')
    ws.column_dimensions['A'].width=31
    for c in ['B','C','D','E','F','G','H']: ws.column_dimensions[c].width=17
    for rr in [1,2,31,32,33,34,35,36]:
        ws.merge_cells(start_row=rr,start_column=1,end_row=rr,end_column=8)
        ws.cell(rr,1).alignment=Alignment(wrap_text=True,vertical='center')
        ws.row_dimensions[rr].height=27
    ws.print_area='A1:H37'
    ws=wb['Checks'];head(ws,1,['Pemeriksaan','Aktual','Pembanding','Selisih','Status'])
    checks=[('Stok akhir vs buku gerak','=SUM(Calculation!I2:I61)','=SUM(Master!F14:F73)+SUM(Raw_Data!G2:G'+str(end)+')'),('Saldo harian terakhir','=SUM(Permintaan_Harian!BK367:DR367)','=SUM(Calculation!I2:I61)'),('Jumlah SKU','=COUNTA(Calculation!A2:A61)','=COUNTA(Master!A14:A73)'),('Klasifikasi ABC','=SUM(Report!B26:B28)','=COUNTA(Calculation!A2:A61)'),('Jumlah status','=SUM(Report!B10:B14)','=COUNTA(Calculation!A2:A61)'),('HPP keluar','=SUM(Calculation!R2:R61)','=SUM(Raw_Data!I2:I'+str(end)+')'),('Stok negatif harian','=COUNTIFS(Permintaan_Harian!BK2:DR367,"<0")','=0'),('Baris permintaan harian','=COUNT(Permintaan_Harian!A2:A367)','=Master!B2-Master!B10+1'),('Penerimaan PO','=SUMIFS(Raw_Data!E2:E'+str(end)+',Raw_Data!D2:D'+str(end)+',"MASUK")','=SUM(Purchase_Order!G2:G501)'),('Parameter ABC berurutan','=IF(AND(Master!B8>0,Master!B8<Master!B9,Master!B9<1),1,0)','=1'),('Latihan safety stock terisi','=COUNT(Calculation!N2:O61)+COUNT(Calculation!Q2:Q61)','=COUNTA(Calculation!A2:A61)*3')]
    for i,(label,a,b) in enumerate(checks,2):
        put(ws,i,1,label);put(ws,i,2,a);put(ws,i,3,b);put(ws,i,4,f'=B{i}-C{i}');put(ws,i,5,f'=IF(D{i}=0,"PASS",'+('"BELUM DIISI"' if exercise and i==12 else '"FAIL"')+')')
    ws.column_dimensions['A'].width=35
    wb.calculation=CalcProperties(calcId=191029,fullCalcOnLoad=True)
    wb.active=5
    out=ROOT/'workbook'/f'{"latihan" if exercise else "solusi"}_inventory_operations.xlsx';wb.save(out)
    print(out.name)


if __name__=='__main__':
    build();build(True)
