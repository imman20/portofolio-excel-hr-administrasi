"""Proyek 05 - Pengelolaan Cuti (tingkat menengah). Tahun 2024."""

import datetime as dt
import random

from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill

from data_gen import buat_karyawan
from util import (DEPARTEMEN, FMT_TGL, LIBUR_2024, WARNA_HEADER, WARNA_INPUT, WARNA_SUB, cf_kurang_dari, cf_rumus,
                  cf_sama_dengan, chart_bar, dropdown, font, header, judul, kolom, label_nilai, lebar, rapikan, ref,
                  tabel, teks_panduan, tulis_baris)

N = 60
JENIS_CUTI = ["Tahunan", "Sakit", "Melahirkan", "Khusus (menikah/duka)"]
STATUS = ["Diajukan", "Disetujui", "Ditolak"]
ALASAN = ["Keperluan keluarga", "Liburan", "Urusan pribadi", "Pernikahan saudara", "Pemulihan kesehatan",
          "Mengurus dokumen", "Acara keluarga di luar kota", "Istirahat"]


def build():
    rng = random.Random(505)
    wb = Workbook()
    data = buat_karyawan(N, seed=101)[:N]
    nlib = len(LIBUR_2024)

    # ---------------- Hari Libur
    wl = wb.active
    wl.title = "Hari Libur 2024"
    judul(wl, "Daftar Hari Libur Nasional dan Cuti Bersama 2024", "Rentang A5:A31 dipakai sebagai argumen holidays pada NETWORKDAYS.INTL.")
    header(wl, 4, 1, ["Tanggal", "Hari", "Keterangan", "Jatuh di Hari Kerja?"])
    for i, (d, ket) in enumerate(LIBUR_2024, start=5):
        tulis_baris(wl, i, 1, [d, f'=TEXT(A{i},"dddd")', ket, f'=IF(WEEKDAY(A{i},2)<=5,"Ya","Tidak (akhir pekan)")'],
                    [FMT_TGL, None, None, None])
    akhir_lib = 4 + nlib
    label_nilai(wl, akhir_lib + 2, 1, "Jumlah hari libur", f"=COUNTA(A5:A{akhir_lib})", "0")
    label_nilai(wl, akhir_lib + 3, 1, "Libur yang jatuh di hari kerja", f'=COUNTIF(D5:D{akhir_lib},"Ya")', "0")
    label_nilai(wl, akhir_lib + 4, 1, "Hari kerja efektif 2024 (NETWORKDAYS.INTL)",
                f'=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,12,31),1,A5:A{akhir_lib})', "0")
    lebar(wl, [14, 12, 36, 22])

    # ---------------- Karyawan & saldo
    wk = wb.create_sheet("Saldo Cuti")
    judul(wk, "Saldo Cuti Tahunan 2024", "Hak cuti 12 hari setelah masa kerja 12 bulan. Cuti terpakai dihitung SUMIFS dari pengajuan yang disetujui.")
    header(wk, 4, 1, ["NIK", "Nama", "Departemen", "Tanggal Masuk", "Masa Kerja (bln) per 1 Jan 2024", "Hak Cuti",
                      "Cuti Terpakai (Disetujui)", "Sedang Diajukan", "Sisa Saldo", "Status Saldo"])
    label_nilai(wk, 2, 7, "Hak cuti penuh (hari)", 12, "0", fill=WARNA_INPUT)
    label_nilai(wk, 3, 7, "Tanggal awal tahun cuti", dt.date(2024, 1, 1), FMT_TGL, fill=WARNA_INPUT)
    for i, k in enumerate(data):
        r = 5 + i
        tulis_baris(wk, r, 1, [
            k["NIK"], k["Nama"], k["Departemen"], k["TglMasuk"],
            f'=IF(D{r}>$H$3,0,DATEDIF(D{r},$H$3,"M"))',
            f"=IF(E{r}>=12,$H$2,0)",
            f"=SUMIFS('Pengajuan Cuti'!$I$5:$I$500,'Pengajuan Cuti'!$C$5:$C$500,A{r},'Pengajuan Cuti'!$E$5:$E$500,\"Tahunan\",'Pengajuan Cuti'!$J$5:$J$500,\"Disetujui\")",
            f"=SUMIFS('Pengajuan Cuti'!$I$5:$I$500,'Pengajuan Cuti'!$C$5:$C$500,A{r},'Pengajuan Cuti'!$E$5:$E$500,\"Tahunan\",'Pengajuan Cuti'!$J$5:$J$500,\"Diajukan\")",
            f"=F{r}-G{r}",
            f'=IF(F{r}=0,"Belum berhak",IF(I{r}<=2,"Hampir habis","Tersedia"))',
        ], ["@", None, None, FMT_TGL, "0", "0", "0", "0", "0", None])
    akhir = 4 + N
    tabel(wk, f"A4:J{akhir}", "tblSaldo")
    wk.freeze_panes = "C5"
    lebar(wk, [10, 24, 13, 13, 14, 9, 13, 11, 10, 14])
    cf_kurang_dari(wk, f"I5:I{akhir}", 3)
    cf_sama_dengan(wk, f"J5:J{akhir}", "Belum berhak", "D9D9D9", "595959")

    # ---------------- Pengajuan Cuti
    wp = wb.create_sheet("Pengajuan Cuti")
    judul(wp, "Formulir dan Log Pengajuan Cuti 2024", "Hari kerja cuti dihitung NETWORKDAYS.INTL; validasi menolak cuti tahunan yang melebihi saldo.")
    header(wp, 4, 1, ["No", "Tanggal Pengajuan", "NIK", "Nama", "Jenis Cuti", "Tanggal Mulai", "Tanggal Selesai",
                      "Hari Kalender", "Hari Kerja (NETWORKDAYS.INTL)", "Status", "Saldo Tersedia Saat Ini", "Validasi",
                      "Alasan", "Disetujui Oleh"])
    pengajuan = []
    tgl = dt.date(2024, 1, 8)
    for i in range(90):
        k = rng.choice(data)
        tgl = tgl + dt.timedelta(days=rng.randint(0, 5))
        if tgl > dt.date(2024, 10, 31):
            break
        jenis = rng.choices(JENIS_CUTI, weights=[70, 18, 2, 10])[0]
        mulai = tgl + dt.timedelta(days=rng.randint(3, 30))
        if mulai.weekday() >= 5:
            mulai += dt.timedelta(days=7 - mulai.weekday())
        lama = {"Tahunan": rng.choice([1, 1, 2, 2, 3, 4, 5]), "Sakit": rng.choice([1, 2, 3]),
                "Melahirkan": 90, "Khusus (menikah/duka)": rng.choice([2, 3])}[jenis]
        selesai = mulai + dt.timedelta(days=lama - 1)
        status = rng.choices(STATUS, weights=[12, 80, 8])[0] if mulai < dt.date(2024, 10, 1) else "Diajukan"
        pengajuan.append((tgl, k, jenis, mulai, selesai, status))
    pengajuan.sort(key=lambda p: p[0])
    for i, (t, k, jenis, mulai, selesai, status) in enumerate(pengajuan):
        r = 5 + i
        tulis_baris(wp, r, 1, [
            "=ROW()-4", t, k["NIK"],
            f'=IFERROR(INDEX(\'Saldo Cuti\'!$B$5:$B${akhir},MATCH(C{r},\'Saldo Cuti\'!$A$5:$A${akhir},0)),"")',
            jenis, mulai, selesai,
            f"=G{r}-F{r}+1",
            f"=NETWORKDAYS.INTL(F{r},G{r},1,'Hari Libur 2024'!$A$5:$A${akhir_lib})",
            status,
            # saldo saat baris ini diproses: hak cuti - cuti tahunan disetujui pada baris-baris di atasnya
            f"=IFERROR(INDEX('Saldo Cuti'!$F$5:$F${akhir},MATCH(C{r},'Saldo Cuti'!$A$5:$A${akhir},0)),0)"
            f"-SUMIFS($I$5:I{r - 1 if r > 5 else 5},$C$5:C{r - 1 if r > 5 else 5},C{r},$E$5:E{r - 1 if r > 5 else 5},\"Tahunan\",$J$5:J{r - 1 if r > 5 else 5},\"Disetujui\")" if r > 5 else
            f"=IFERROR(INDEX('Saldo Cuti'!$F$5:$F${akhir},MATCH(C{r},'Saldo Cuti'!$A$5:$A${akhir},0)),0)",
            f'=IF(E{r}<>"Tahunan","Tidak memakai saldo",IF(I{r}>K{r},"MELEBIHI SALDO",IF(F{r}<B{r},"Tanggal mulai sebelum pengajuan","OK")))',
            rng.choice(ALASAN), "Manager HR" if status != "Diajukan" else "",
        ], ["0", FMT_TGL, "@", None, None, FMT_TGL, FMT_TGL, "0", "0", None, "0", None, None, None])
    akhir_p = 4 + len(pengajuan)
    tabel(wp, f"A4:N{akhir_p}", "tblPengajuan")
    wp.freeze_panes = "E5"
    lebar(wp, [5, 13, 10, 22, 16, 13, 13, 9, 12, 11, 11, 20, 24, 13])
    dropdown(wp, f"C5:C{akhir_p + 200}", f"='Saldo Cuti'!$A$5:$A${akhir}")
    dropdown(wp, f"E5:E{akhir_p + 200}", "=Daftar!$A$2:$A$5")
    dropdown(wp, f"J5:J{akhir_p + 200}", "=Daftar!$B$2:$B$4")
    cf_sama_dengan(wp, f"L5:L{akhir_p}", "MELEBIHI SALDO")
    cf_sama_dengan(wp, f"J5:J{akhir_p}", "Disetujui", "C6EFCE", "006100")
    cf_sama_dengan(wp, f"J5:J{akhir_p}", "Ditolak", "FFC7CE", "9C0006")
    cf_sama_dengan(wp, f"J5:J{akhir_p}", "Diajukan", "FFEB9C", "9C5700")
    # Data validation tanggal: selesai >= mulai
    from openpyxl.worksheet.datavalidation import DataValidation
    dv = DataValidation(type="date", operator="greaterThanOrEqual", formula1="F5", allow_blank=True,
                        showErrorMessage=True, errorTitle="Tanggal tidak valid",
                        error="Tanggal selesai harus sama atau sesudah tanggal mulai.")
    wp.add_data_validation(dv)
    dv.add(f"G5:G{akhir_p + 200}")

    # ---------------- Rekap
    wr = wb.create_sheet("Rekap Cuti")
    judul(wr, "Rekap Cuti per Departemen dan per Bulan", "COUNTIFS dan SUMIFS atas log pengajuan.")
    header(wr, 4, 1, ["Departemen", "Jumlah Karyawan", "Total Hak Cuti", "Terpakai", "Sisa", "% Terpakai",
                      "Pengajuan Ditolak"])
    for i, d in enumerate(DEPARTEMEN, start=5):
        tulis_baris(wr, i, 1, [
            d, f"=COUNTIF('Saldo Cuti'!$C$5:$C${akhir},A{i})",
            f"=SUMIF('Saldo Cuti'!$C$5:$C${akhir},A{i},'Saldo Cuti'!$F$5:$F${akhir})",
            f"=SUMIF('Saldo Cuti'!$C$5:$C${akhir},A{i},'Saldo Cuti'!$G$5:$G${akhir})",
            f"=C{i}-D{i}", f"=IFERROR(D{i}/C{i},0)",
            f"=COUNTIFS('Pengajuan Cuti'!$J$5:$J$500,\"Ditolak\",'Pengajuan Cuti'!$O$5:$O$500,A{i})",
        ], [None, "0", "0", "0", "0", "0.0%", "0"])
    tulis_baris(wr, 11, 1, ["Total"] + [f"=SUM({c}5:{c}10)" for c in "BCDE"] + ["=IFERROR(D11/C11,0)", "=SUM(G5:G10)"],
                [None, "0", "0", "0", "0", "0.0%", "0"])
    for c in range(1, 8):
        wr.cell(row=11, column=c).font = font(bold=True)
        wr.cell(row=11, column=c).fill = PatternFill("solid", fgColor=WARNA_SUB)
    header(wr, 13, 1, ["Bulan", "Pengajuan Masuk", "Disetujui", "Ditolak", "Hari Cuti Disetujui"])
    for m in range(1, 13):
        r = 13 + m
        wr.cell(row=r, column=8, value=dt.date(2024, m, 1)).number_format = FMT_TGL
        wr.cell(row=r, column=9, value=dt.date(2024 + (m == 12), (m % 12) + 1, 1)).number_format = FMT_TGL
        pc = "'Pengajuan Cuti'"
        tulis_baris(wr, r, 1, [
            f'=TEXT(H{r},"mmmm")',
            f'=COUNTIFS({pc}!$F$5:$F$500,">="&H{r},{pc}!$F$5:$F$500,"<"&I{r})',
            f'=COUNTIFS({pc}!$F$5:$F$500,">="&H{r},{pc}!$F$5:$F$500,"<"&I{r},{pc}!$J$5:$J$500,"Disetujui")',
            f'=COUNTIFS({pc}!$F$5:$F$500,">="&H{r},{pc}!$F$5:$F$500,"<"&I{r},{pc}!$J$5:$J$500,"Ditolak")',
            f'=SUMIFS({pc}!$I$5:$I$500,{pc}!$F$5:$F$500,">="&H{r},{pc}!$F$5:$F$500,"<"&I{r},{pc}!$J$5:$J$500,"Disetujui")',
        ], [None, "0", "0", "0", "0"])
    wr["H13"], wr["I13"] = "Awal bulan (bantu)", "Awal bulan berikut (bantu)"
    wr["H13"].font = wr["I13"].font = font(italic=True, size=9, color="595959")
    lebar(wr, [16, 14, 13, 11, 10, 11, 14, 2, 16, 20])
    chart_bar(wr, "K4", "Hari Cuti Disetujui per Bulan 2024", ref(wr, 1, 14, 1, 25), ref(wr, 5, 13, 5, 25),
              judul_y="Hari", lebar_cm=16, tinggi_cm=8)
    chart_bar(wr, "K21", "Pemakaian Cuti per Departemen", ref(wr, 1, 5, 1, 10), ref(wr, 4, 4, 5, 10),
              judul_y="Hari", lebar_cm=16, tinggi_cm=8, tumpuk=True)

    # ---------------- Kalender ringkas
    wc = wb.create_sheet("Kalender")
    judul(wc, "Kalender Cuti Ringkas", "Pilih bulan di sel C4; angka di tiap tanggal = jumlah karyawan yang sedang cuti (disetujui) hari itu.")
    wc["B4"] = "Bulan (1-12):"
    wc["B4"].font = font(bold=True)
    wc["C4"] = 6
    wc["C4"].fill = PatternFill("solid", fgColor=WARNA_INPUT)
    wc["C4"].font = font(bold=True)
    dropdown(wc, "C4", [str(i) for i in range(1, 13)])
    wc["D4"] = "Tahun:"
    wc["D4"].font = font(bold=True)
    wc["E4"] = 2024
    wc["E4"].fill = PatternFill("solid", fgColor=WARNA_INPUT)
    wc["B5"] = "Tanggal awal bulan:"
    wc["C5"] = "=DATE($E$4,$C$4,1)"
    wc["C5"].number_format = FMT_TGL
    wc["B6"] = "Awal grid (Senin minggu pertama):"
    wc["C6"] = "=C5-WEEKDAY(C5,2)+1"
    wc["C6"].number_format = FMT_TGL
    wc.merge_cells("B8:H8")
    wc["B8"] = '=UPPER(TEXT(C5,"mmmm yyyy"))'
    wc["B8"].font = font(bold=True, size=13, color="FFFFFF")
    wc["B8"].fill = PatternFill("solid", fgColor=WARNA_HEADER)
    wc["B8"].alignment = Alignment(horizontal="center")
    header(wc, 9, 2, ["Sen", "Sel", "Rab", "Kam", "Jum", "Sab", "Min"], fill=WARNA_SUB, warna_teks="000000", wrap=False)
    pc = "'Pengajuan Cuti'"
    for minggu in range(6):
        r_tgl = 10 + minggu * 2
        r_jml = r_tgl + 1
        for hari in range(7):
            col = 2 + hari
            offset = minggu * 7 + hari
            kol = kolom(col)
            c1 = wc.cell(row=r_tgl, column=col,
                         value=f'=IF(MONTH($C$6+{offset})=$C$4,$C$6+{offset},"")')
            c1.number_format = "d"
            c1.font = font(bold=True, size=9, color="595959")
            c1.alignment = Alignment(horizontal="right", vertical="top")
            c2 = wc.cell(row=r_jml, column=col,
                         value=f'=IF({kol}{r_tgl}="","",COUNTIFS({pc}!$F$5:$F$500,"<="&{kol}{r_tgl},{pc}!$G$5:$G$500,">="&{kol}{r_tgl},{pc}!$J$5:$J$500,"Disetujui"))')
            c2.alignment = Alignment(horizontal="center", vertical="center")
            c2.font = font(bold=True, size=12, color=WARNA_HEADER)
            from util import BORDER
            c1.border = BORDER
            c2.border = BORDER
        wc.row_dimensions[r_jml].height = 26
    cf_rumus(wc, "B11:H21", "AND(ISNUMBER(B11),B11>=3)")
    cf_rumus(wc, "B10:H20", 'COUNTIF(\'Hari Libur 2024\'!$A$5:$A$31,B10)>0', warna="FFEB9C", teks="9C5700")
    wc["B23"] = "Keterangan: sel merah = 3 orang atau lebih cuti bersamaan; tanggal kuning = hari libur nasional/cuti bersama."
    wc["B23"].font = font(italic=True, size=9, color="595959")
    wc["B25"] = "Daftar karyawan cuti pada bulan terpilih (status Disetujui):"
    wc["B25"].font = font(bold=True)
    header(wc, 26, 2, ["No", "Nama", "Mulai", "Selesai", "Hari Kerja", "Jenis", "Peringkat (bantu)"], fill=WARNA_SUB,
           warna_teks="000000", wrap=False)
    # kolom bantu di Pengajuan Cuti: nomor urut pengajuan yang beririsan dengan bulan terpilih
    wp["P4"] = "Bantu Kalender: urutan dalam bulan terpilih"
    wp["P4"].font = font(italic=True, size=9, color="595959")
    wp["O4"] = "Departemen (bantu)"
    wp["O4"].font = font(italic=True, size=9, color="595959")
    for i in range(len(pengajuan)):
        r = 5 + i
        wp.cell(row=r, column=15, value=f"=IFERROR(INDEX('Saldo Cuti'!$C$5:$C${akhir},MATCH(C{r},'Saldo Cuti'!$A$5:$A${akhir},0)),\"\")")
        wp.cell(row=r, column=16, value=(
            f'=IF(AND(J{r}="Disetujui",F{r}<=EOMONTH(Kalender!$C$5,0),G{r}>=Kalender!$C$5),'
            f'COUNTIFS($J$5:J{r},"Disetujui",$F$5:F{r},"<="&EOMONTH(Kalender!$C$5,0),$G$5:G{r},">="&Kalender!$C$5),"")'))
    for i in range(15):
        r = 27 + i
        tulis_baris(wc, r, 2, [
            i + 1,
            f'=IFERROR(INDEX({pc}!$D$5:$D$500,MATCH(B{r},{pc}!$P$5:$P$500,0)),"")',
            f'=IFERROR(INDEX({pc}!$F$5:$F$500,MATCH(B{r},{pc}!$P$5:$P$500,0)),"")',
            f'=IFERROR(INDEX({pc}!$G$5:$G$500,MATCH(B{r},{pc}!$P$5:$P$500,0)),"")',
            f'=IFERROR(INDEX({pc}!$I$5:$I$500,MATCH(B{r},{pc}!$P$5:$P$500,0)),"")',
            f'=IFERROR(INDEX({pc}!$E$5:$E$500,MATCH(B{r},{pc}!$P$5:$P$500,0)),"")',
            "",
        ], ["0", None, FMT_TGL, FMT_TGL, "0", None, None])
    lebar(wc, [3, 12, 12, 12, 12, 12, 12, 12])
    wc.column_dimensions["C"].width = 22
    wc.column_dimensions["B"].width = 30

    wd = wb.create_sheet("Daftar")
    header(wd, 1, 1, ["Jenis Cuti", "Status"])
    for i, j in enumerate(JENIS_CUTI, start=2):
        wd.cell(row=i, column=1, value=j)
    for i, s in enumerate(STATUS, start=2):
        wd.cell(row=i, column=2, value=s)
    lebar(wd, [24, 12])

    wpt = wb.create_sheet("Petunjuk")
    teks_panduan(wpt, 1, [
        "# Petunjuk Penggunaan",
        "1. Hari Libur 2024: daftar tanggal merah dan cuti bersama; dipakai NETWORKDAYS.INTL agar hari libur tidak memotong saldo cuti.",
        "2. Saldo Cuti: hak cuti 12 hari untuk karyawan bermasa kerja >= 12 bulan per 1 Januari (sel kuning H2:H3 dapat diubah).",
        "3. Pengajuan Cuti: isi tanggal pengajuan, NIK (dropdown), jenis, tanggal mulai dan selesai. Kolom hari kerja, saldo, dan validasi otomatis.",
        "4. Validasi menampilkan MELEBIHI SALDO (merah) bila cuti tahunan lebih besar dari saldo yang tersedia saat baris itu diproses.",
        "5. Kalender: ganti bulan di C4; grid tanggal dibangun dengan DATE, WEEKDAY, dan MONTH; jumlah orang cuti dihitung COUNTIFS.",
        "6. Rekap Cuti: ringkasan per departemen dan per bulan beserta chart.",
    ])
    wb.active = wb.sheetnames.index("Pengajuan Cuti")
    return rapikan(wb)


SCREENSHOT = [("Pengajuan Cuti", "05_pengajuan_cuti.png", "A1:N34"), ("Kalender", "05_kalender.png", "A1:I42"),
              ("Saldo Cuti", "05_saldo_cuti.png", "A1:J30")]
