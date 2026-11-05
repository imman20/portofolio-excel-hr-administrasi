"""Proyek 08 - HR Dashboard Analitik (tingkat lanjut). 250 karyawan, data turnover 2023-2024."""

import csv
import datetime as dt
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill

from data_gen import buat_karyawan
from util import (BORDER, DEPARTEMEN, FMT_PCT, FMT_RP, FMT_TGL, WARNA_HEADER, WARNA_INPUT, WARNA_SUB, chart_bar,
                  chart_line, chart_pie, dropdown, font, header, judul, kartu_kpi, label_nilai, lebar, rapikan, ref,
                  tabel, teks_panduan, tulis_baris)

N = 250
ACUAN = dt.date(2024, 12, 31)
KELOMPOK_USIA = [("< 25", 0, 25), ("25-34", 25, 35), ("35-44", 35, 45), ("45-54", 45, 55), (">= 55", 55, 200)]
FOLDER_DATA = Path(__file__).resolve().parent.parent / "08_hr-dashboard-analitik" / "data"


def tulis_csv(data):
    """Dua file CSV untuk latihan Power Query (gabung karyawan + turnover)."""
    FOLDER_DATA.mkdir(parents=True, exist_ok=True)
    with open(FOLDER_DATA / "karyawan.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["NIK", "Nama", "JenisKelamin", "TanggalLahir", "Departemen", "Jabatan", "Level", "StatusKepegawaian",
                    "TanggalMasuk", "GajiPokok", "Kota"])
        for k in data:
            w.writerow([k["NIK"], k["Nama"], k["JK"], k["TglLahir"].isoformat(), k["Departemen"], k["Jabatan"],
                        k["Level"], k["Status"], k["TglMasuk"].isoformat(), k["Gaji"], k["Kota"]])
    with open(FOLDER_DATA / "turnover_2023_2024.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["NIK", "TanggalKeluar", "AlasanKeluar"])
        for k in data:
            if k["TglKeluar"]:
                w.writerow([k["NIK"], k["TglKeluar"].isoformat(), k["AlasanKeluar"]])


def build():
    wb = Workbook()
    data = buat_karyawan(N, seed=808, dengan_keluar=True)
    tulis_csv(data)
    akhir = 4 + N
    D = "'Data Karyawan'"
    P = "Perhitungan"

    # ================= Dashboard (sheet pertama)
    wd = wb.active
    wd.title = "Dashboard"
    wd.sheet_view.showGridLines = False
    wd.merge_cells("B2:P2")
    wd["B2"] = "HR DASHBOARD - PT CONTOH PERUSAHAAN INDONESIA"
    wd["B2"].font = font(bold=True, size=16, color="FFFFFF")
    wd["B2"].fill = PatternFill("solid", fgColor=WARNA_HEADER)
    wd["B2"].alignment = Alignment(horizontal="center", vertical="center")
    wd.row_dimensions[2].height = 30
    wd["B3"] = "Tanggal acuan:"
    wd["C3"] = ACUAN
    wd["C3"].number_format = FMT_TGL
    wd["C3"].fill = PatternFill("solid", fgColor=WARNA_INPUT)
    wd["C3"].font = font(bold=True)
    wd["E3"] = "Filter departemen:"
    wd["F3"] = "Semua"
    wd["F3"].fill = PatternFill("solid", fgColor=WARNA_INPUT)
    wd["F3"].font = font(bold=True)
    dropdown(wd, "F3", ["Semua"] + DEPARTEMEN)
    wd["H3"] = "Kriteria (bantu):"
    wd["I3"] = '=IF(F3="Semua","*",F3)'
    wd["I3"].font = font(italic=True, color="595959")
    for c in ("B3", "E3", "H3"):
        wd[c].font = font(bold=True)

    kpi = [
        ("HEADCOUNT AKTIF", f'=COUNTIFS({D}!$L$5:$L${akhir},"Aktif",{D}!$F$5:$F${akhir},$I$3)', "#,##0"),
        ("TURNOVER RATE 2024", f"=IFERROR({P}!C64/{P}!C65,0)", FMT_PCT),
        ("RATA-RATA MASA KERJA (THN)", f'=IFERROR(AVERAGEIFS({D}!$K$5:$K${akhir},{D}!$L$5:$L${akhir},"Aktif",{D}!$F$5:$F${akhir},$I$3),0)', "0.0"),
        ("RATA-RATA USIA (THN)", f'=IFERROR(AVERAGEIFS({D}!$E$5:$E${akhir},{D}!$L$5:$L${akhir},"Aktif",{D}!$F$5:$F${akhir},$I$3),0)', "0.0"),
        ("RASIO GENDER (L : P)", f'=TEXT(IFERROR(COUNTIFS({D}!$C$5:$C${akhir},"L",{D}!$L$5:$L${akhir},"Aktif",{D}!$F$5:$F${akhir},$I$3)/C7,0),"0%")&" : "&TEXT(IFERROR(COUNTIFS({D}!$C$5:$C${akhir},"P",{D}!$L$5:$L${akhir},"Aktif",{D}!$F$5:$F${akhir},$I$3)/C7,0),"0%")', None),
        ("MASUK 2024", f'=COUNTIFS({D}!$J$5:$J${akhir},">="&DATE(2024,1,1),{D}!$J$5:$J${akhir},"<="&DATE(2024,12,31),{D}!$F$5:$F${akhir},$I$3)', "0"),
        ("KELUAR 2024", f'=COUNTIFS({D}!$M$5:$M${akhir},">="&DATE(2024,1,1),{D}!$M$5:$M${akhir},"<="&DATE(2024,12,31),{D}!$F$5:$F${akhir},$I$3)', "0"),
        ("RATA-RATA GAJI AKTIF", f'=IFERROR(AVERAGEIFS({D}!$O$5:$O${akhir},{D}!$L$5:$L${akhir},"Aktif",{D}!$F$5:$F${akhir},$I$3),0)', FMT_RP),
    ]
    col = 2
    for label, rumus, fmt in kpi:
        kartu_kpi(wd, 6, col, label, rumus, fmt, lebar_kolom=2)
        col += 2
    wd["B9"] = "Catatan: turnover rate 2024 = jumlah keluar 2024 / rata-rata headcount (awal dan akhir tahun), dihitung untuk seluruh perusahaan; KPI lain mengikuti filter departemen."
    wd["B9"].font = font(italic=True, size=9, color="595959")
    for c in range(2, 18):
        wd.column_dimensions[chr(64 + c) if c <= 26 else "A"].width = 11
    wd.column_dimensions["A"].width = 2

    # ================= Data Karyawan
    ws = wb.create_sheet("Data Karyawan")
    judul(ws, "Dataset Karyawan (250 baris)", "Kolom Usia, Masa Kerja, Kelompok Usia, dan Tahun adalah rumus; tanggal acuan ada di Dashboard!C3.")
    head = ["NIK", "Nama", "JK", "Tanggal Lahir", "Usia", "Departemen", "Jabatan", "Level", "Status Kepegawaian",
            "Tanggal Masuk", "Masa Kerja (thn)", "Status Aktif", "Tanggal Keluar", "Alasan Keluar", "Gaji Pokok",
            "Kota", "Kelompok Usia", "Tahun Masuk", "Tahun Keluar", "Pendidikan"]
    header(ws, 4, 1, head)
    for i, k in enumerate(data):
        r = 5 + i
        tulis_baris(ws, r, 1, [
            k["NIK"], k["Nama"], k["JK"], k["TglLahir"],
            f'=DATEDIF(D{r},Dashboard!$C$3,"Y")',
            k["Departemen"], k["Jabatan"], k["Level"], k["Status"], k["TglMasuk"],
            f'=ROUND((IF(L{r}="Aktif",Dashboard!$C$3,M{r})-J{r})/365.25,1)',
            k["StatusAktif"], k["TglKeluar"], k["AlasanKeluar"], k["Gaji"], k["Kota"],
            f'=IF(E{r}<25,"< 25",IF(E{r}<35,"25-34",IF(E{r}<45,"35-44",IF(E{r}<55,"45-54",">= 55"))))',
            f"=YEAR(J{r})", f'=IF(M{r}="","",YEAR(M{r}))', k["Pendidikan"],
        ], ["@", None, None, FMT_TGL, "0", None, None, None, None, FMT_TGL, "0.0", None, FMT_TGL, None, FMT_RP, None,
            None, "0", "0", None])
    tabel(ws, f"A4:T{akhir}", "tblKaryawan")
    ws.freeze_panes = "C5"
    lebar(ws, [10, 22, 5, 12, 6, 12, 24, 12, 12, 12, 10, 10, 12, 14, 14, 11, 11, 9, 9, 10])

    # ================= Perhitungan (tabel bantu untuk chart)
    wp = wb.create_sheet("Perhitungan")
    judul(wp, "Tabel Perhitungan Dashboard", "Semua tabel di sini adalah sumber chart dan KPI. Rumus mengikuti filter departemen di Dashboard!F3 kecuali turnover tahunan.")
    header(wp, 4, 1, ["Departemen", "Headcount Aktif", "Laki-laki", "Perempuan", "Tetap", "Kontrak", "Magang",
                      "Rata-rata Usia", "Rata-rata Masa Kerja", "Rata-rata Gaji", "Keluar 2024", "Turnover 2024"])
    for i, d in enumerate(DEPARTEMEN, start=5):
        tulis_baris(wp, i, 1, [
            d, f'=COUNTIFS({D}!$F$5:$F${akhir},A{i},{D}!$L$5:$L${akhir},"Aktif")',
            f'=COUNTIFS({D}!$F$5:$F${akhir},A{i},{D}!$L$5:$L${akhir},"Aktif",{D}!$C$5:$C${akhir},"L")',
            f'=COUNTIFS({D}!$F$5:$F${akhir},A{i},{D}!$L$5:$L${akhir},"Aktif",{D}!$C$5:$C${akhir},"P")',
            f'=COUNTIFS({D}!$F$5:$F${akhir},A{i},{D}!$L$5:$L${akhir},"Aktif",{D}!$I$5:$I${akhir},"Tetap")',
            f'=COUNTIFS({D}!$F$5:$F${akhir},A{i},{D}!$L$5:$L${akhir},"Aktif",{D}!$I$5:$I${akhir},"Kontrak")',
            f'=COUNTIFS({D}!$F$5:$F${akhir},A{i},{D}!$L$5:$L${akhir},"Aktif",{D}!$I$5:$I${akhir},"Magang")',
            f'=IFERROR(AVERAGEIFS({D}!$E$5:$E${akhir},{D}!$F$5:$F${akhir},A{i},{D}!$L$5:$L${akhir},"Aktif"),0)',
            f'=IFERROR(AVERAGEIFS({D}!$K$5:$K${akhir},{D}!$F$5:$F${akhir},A{i},{D}!$L$5:$L${akhir},"Aktif"),0)',
            f'=IFERROR(AVERAGEIFS({D}!$O$5:$O${akhir},{D}!$F$5:$F${akhir},A{i},{D}!$L$5:$L${akhir},"Aktif"),0)',
            f'=COUNTIFS({D}!$F$5:$F${akhir},A{i},{D}!$S$5:$S${akhir},2024)',
            f"=IFERROR(K{i}/((B{i}+K{i}+B{i})/2),0)",
        ], [None, "0", "0", "0", "0", "0", "0", "0.0", "0.0", FMT_RP, "0", FMT_PCT])
    tulis_baris(wp, 11, 1, ["Total"] + [f"=SUM({c}5:{c}10)" for c in "BCDEFG"] +
                [f'=IFERROR(AVERAGEIFS({D}!$E$5:$E${akhir},{D}!$L$5:$L${akhir},"Aktif"),0)',
                 f'=IFERROR(AVERAGEIFS({D}!$K$5:$K${akhir},{D}!$L$5:$L${akhir},"Aktif"),0)',
                 f'=IFERROR(AVERAGEIFS({D}!$O$5:$O${akhir},{D}!$L$5:$L${akhir},"Aktif"),0)',
                 "=SUM(K5:K10)", "=IFERROR(K11/((B11+K11+B11)/2),0)"],
                [None, "0", "0", "0", "0", "0", "0", "0.0", "0.0", FMT_RP, "0", FMT_PCT])
    for c in range(1, 13):
        wp.cell(row=11, column=c).font = font(bold=True)
        wp.cell(row=11, column=c).fill = PatternFill("solid", fgColor=WARNA_SUB)

    header(wp, 14, 1, ["Kelompok Usia", "Jumlah Aktif", "% Aktif"])
    for i, (lab, lo, hi) in enumerate(KELOMPOK_USIA, start=15):
        tulis_baris(wp, i, 1, [
            lab, f'=COUNTIFS({D}!$Q$5:$Q${akhir},A{i},{D}!$L$5:$L${akhir},"Aktif",{D}!$F$5:$F${akhir},Dashboard!$I$3)',
            "=IFERROR(B%d/SUM($B$15:$B$19),0)" % i], [None, "0", FMT_PCT])
    header(wp, 14, 5, ["Level", "Jumlah Aktif"])
    for i, lv in enumerate(["Staff", "Senior Staff", "Supervisor", "Manager"], start=15):
        tulis_baris(wp, i, 5, [lv, f'=COUNTIFS({D}!$H$5:$H${akhir},E{i},{D}!$L$5:$L${akhir},"Aktif",{D}!$F$5:$F${akhir},Dashboard!$I$3)'],
                    [None, "0"])
    header(wp, 14, 8, ["Alasan Keluar", "Jumlah 2023-2024"])
    for i, al in enumerate(["Resign", "Habis Kontrak", "Pensiun", "Pindah Kota", "PHK"], start=15):
        tulis_baris(wp, i, 8, [al, f'=COUNTIFS({D}!$N$5:$N${akhir},H{i},{D}!$F$5:$F${akhir},Dashboard!$I$3)'], [None, "0"])
    header(wp, 14, 11, ["Status Kepegawaian", "Jumlah Aktif"])
    for i, st in enumerate(["Tetap", "Kontrak", "Magang"], start=15):
        tulis_baris(wp, i, 11, [st, f'=COUNTIFS({D}!$I$5:$I${akhir},K{i},{D}!$L$5:$L${akhir},"Aktif",{D}!$F$5:$F${akhir},Dashboard!$I$3)'],
                    [None, "0"])

    # turnover bulanan 2023-2024 (rentang baris 23-46 = 24 bulan)
    header(wp, 22, 1, ["Bulan", "Awal Bulan", "Awal Bulan Berikut", "Headcount Awal", "Masuk", "Keluar",
                       "Headcount Akhir", "Turnover Bulanan", "Tahun", "Nama Bulan"])
    r = 23
    for th in (2023, 2024):
        for m in range(1, 13):
            awal = dt.date(th, m, 1)
            nxt = dt.date(th + (m == 12), (m % 12) + 1, 1)
            tulis_baris(wp, r, 1, [
                f'=TEXT(B{r},"mmm yyyy")', awal, nxt,
                f'=COUNTIFS({D}!$J$5:$J${akhir},"<"&B{r})-COUNTIFS({D}!$M$5:$M${akhir},"<"&B{r})',
                f'=COUNTIFS({D}!$J$5:$J${akhir},">="&B{r},{D}!$J$5:$J${akhir},"<"&C{r})',
                f'=COUNTIFS({D}!$M$5:$M${akhir},">="&B{r},{D}!$M$5:$M${akhir},"<"&C{r})',
                f"=D{r}+E{r}-F{r}", f"=IFERROR(F{r}/AVERAGE(D{r},G{r}),0)", th, f'=TEXT(B{r},"mmm")',
            ], [None, FMT_TGL, FMT_TGL, "0", "0", "0", "0", FMT_PCT, "0", None])
            r += 1
    # ringkasan tahunan (baris 38-41 di kolom A-C? gunakan baris 48+ agar tidak bentrok) -> taruh di kolom M
    header(wp, 22, 12, ["Ringkasan Tahunan", "2023", "2024"])
    ringkas = [("Headcount awal tahun", "=D23", "=D35"), ("Headcount akhir tahun", "=G34", "=G46"),
               ("Masuk", "=SUM(E23:E34)", "=SUM(E35:E46)"), ("Keluar", "=SUM(F23:F34)", "=SUM(F35:F46)"),
               ("Rata-rata headcount", "=AVERAGE(M23,M24)", "=AVERAGE(N23,N24)"),
               ("Turnover rate tahunan", "=IFERROR(M26/M27,0)", "=IFERROR(N26/N27,0)")]
    for i, (lab, a, b) in enumerate(ringkas, start=23):
        tulis_baris(wp, i, 12, [lab, a, b], [None, "0" if i < 28 else FMT_PCT, "0" if i < 28 else FMT_PCT])
    # sel yang dipakai KPI dashboard
    wp["B64"], wp["C64"] = "Keluar 2024 (untuk KPI)", "=N26"
    wp["B65"], wp["C65"] = "Rata-rata headcount 2024 (untuk KPI)", "=N27"
    for c in ("B64", "B65"):
        wp[c].font = font(italic=True, color="595959")
    # tabel tren per bulan 2 seri (untuk line chart)
    header(wp, 48, 1, ["Bulan", "Keluar 2023", "Keluar 2024", "Turnover 2023", "Turnover 2024"])
    for m in range(1, 13):
        rr = 48 + m
        tulis_baris(wp, rr, 1, [f"=J{22 + m}", f"=F{22 + m}", f"=F{34 + m}", f"=H{22 + m}", f"=H{34 + m}"],
                    [None, "0", "0", FMT_PCT, FMT_PCT])
    lebar(wp, [22, 13, 13, 12, 10, 10, 12, 12, 10, 12, 18, 24, 10, 10])
    wp.freeze_panes = "A5"

    # ================= chart di Dashboard
    chart_bar(wd, "B11", "Headcount Aktif per Departemen", ref(wp, 1, 5, 1, 10), ref(wp, 2, 4, 2, 10),
              judul_y="Orang", lebar_cm=13, tinggi_cm=7.5)
    chart_bar(wd, "H11", "Komposisi Gender per Departemen", ref(wp, 1, 5, 1, 10), ref(wp, 3, 4, 4, 10),
              judul_y="Orang", lebar_cm=13, tinggi_cm=7.5, tumpuk=True)
    chart_pie(wd, "N11", "Status Kepegawaian", ref(wp, 11, 15, 11, 17), ref(wp, 12, 14, 12, 17), lebar_cm=9, tinggi_cm=7.5)
    chart_line(wd, "B27", "Tren Karyawan Keluar per Bulan (2023 vs 2024)", ref(wp, 1, 49, 1, 60), ref(wp, 2, 48, 3, 60),
               judul_y="Orang", lebar_cm=13, tinggi_cm=7.5)
    chart_bar(wd, "H27", "Distribusi Usia Karyawan Aktif", ref(wp, 1, 15, 1, 19), ref(wp, 2, 14, 2, 19),
              judul_y="Orang", lebar_cm=13, tinggi_cm=7.5)
    chart_bar(wd, "N27", "Alasan Keluar 2023-2024", ref(wp, 8, 15, 8, 19), ref(wp, 9, 14, 9, 19),
              judul_y="Orang", lebar_cm=9, tinggi_cm=7.5, horizontal=True)
    # tabel ringkas di dashboard
    header(wd, 43, 2, ["Departemen", "Headcount", "Turnover 2024", "Rata-rata Usia", "Rata-rata Masa Kerja", "Rata-rata Gaji"])
    for i in range(6):
        r = 44 + i
        pr = 5 + i
        tulis_baris(wd, r, 2, [f"={P}!A{pr}", f"={P}!B{pr}", f"={P}!L{pr}", f"={P}!H{pr}", f"={P}!I{pr}", f"={P}!J{pr}"],
                    [None, "0", FMT_PCT, "0.0", "0.0", FMT_RP])
    wd.column_dimensions["B"].width = 14
    wd.column_dimensions["G"].width = 14

    # ================= Panduan Power Query
    wq = wb.create_sheet("Panduan Power Query")
    teks_panduan(wq, 1, [
        "# Panduan Power Query: Menggabungkan karyawan.csv dan turnover_2023_2024.csv",
        "Kedua file ada di folder data/ proyek ini (pemisah titik koma). Tujuannya membentuk tabel seperti sheet Data Karyawan secara otomatis.",
        "## 1. Impor file karyawan",
        "1. Buka workbook baru. Pilih Data > Get Data > From File > From Text/CSV, pilih data/karyawan.csv.",
        "2. Pada jendela pratinjau pastikan Delimiter = Semicolon dan File Origin = 65001: Unicode (UTF-8). Klik Transform Data.",
        "3. Di Power Query Editor, periksa tipe kolom: TanggalLahir dan TanggalMasuk harus Date (klik ikon tipe di header kolom > Date); GajiPokok = Whole Number; NIK = Text.",
        "4. Ganti nama query menjadi 'Karyawan' (panel Query Settings di kanan).",
        "## 2. Impor file turnover",
        "1. Di Power Query Editor: Home > New Source > File > Text/CSV > pilih data/turnover_2023_2024.csv > OK.",
        "2. Pastikan TanggalKeluar bertipe Date dan NIK bertipe Text. Beri nama query 'Turnover'.",
        "## 3. Gabungkan (Merge) dengan Left Outer Join",
        "1. Pilih query Karyawan. Home > Merge Queries > Merge Queries as New.",
        "2. Tabel atas: Karyawan, klik kolom NIK. Tabel bawah: Turnover, klik kolom NIK. Join Kind: Left Outer (all from first, matching from second). OK.",
        "3. Kolom baru 'Turnover' muncul; klik ikon panah ganda di headernya, centang TanggalKeluar dan AlasanKeluar, hilangkan centang 'Use original column name as prefix'. OK.",
        "4. Add Column > Custom Column: nama 'StatusAktif', rumus: if [TanggalKeluar] = null then \"Aktif\" else \"Keluar\".",
        "5. Add Column > Custom Column: nama 'Usia', rumus: Number.RoundDown(Duration.Days(#date(2024,12,31) - [TanggalLahir]) / 365.25).",
        "6. Ganti nama query gabungan menjadi 'DataKaryawan'. Home > Close & Load To > Table, pilih New worksheet.",
        "## 4. Memperbarui",
        "Saat CSV berubah, cukup klik Data > Refresh All; seluruh langkah transformasi dijalankan ulang.",
        "## 5. Pivot dan dashboard dari hasil Power Query",
        "Buat PivotTable dari tabel hasil (Insert > PivotTable), lalu tambahkan slicer Departemen dan timeline TanggalMasuk (lihat sheet Panduan Slicer).",
    ])
    wsl = wb.create_sheet("Panduan Slicer")
    teks_panduan(wsl, 1, [
        "# Panduan Slicer, Timeline, dan Versi Excel 365",
        "## A. Slicer untuk tabel (tanpa PivotTable)",
        "1. Klik sel di dalam tabel tblKaryawan (sheet Data Karyawan).",
        "2. Table Design > Insert Slicer > centang Departemen, Status Aktif, Level > OK.",
        "3. Klik tombol slicer untuk memfilter tabel; klik ikon corong bersilang untuk menghapus filter.",
        "4. Klik slicer > Slicer tab > Columns = 3 agar tombol tersusun mendatar.",
        "## B. Slicer dan Timeline untuk PivotTable",
        "1. Insert > PivotTable dari tblKaryawan. Rows: Departemen; Values: Count of NIK; Columns: Status Aktif.",
        "2. PivotTable Analyze > Insert Slicer > Level, Kelompok Usia.",
        "3. PivotTable Analyze > Insert Timeline > Tanggal Masuk. Ubah tingkat (Years/Quarters/Months) di pojok kanan timeline.",
        "4. Untuk beberapa pivot: klik slicer > Slicer > Report Connections > centang semua PivotTable.",
        "## C. Rumus dinamis Excel 365 (pengganti tabel bantu)",
        "Workbook ini memakai COUNTIFS/INDEX/MATCH agar berjalan di Excel 2016/2019. Di Excel 365 tabel bantu dapat diganti rumus array dinamis:",
        "  =UNIQUE(tblKaryawan[Departemen])  -> daftar departemen unik (tumpah ke bawah otomatis).",
        "  =SORT(UNIQUE(tblKaryawan[Departemen]))  -> daftar departemen terurut.",
        "  =FILTER(tblKaryawan;(tblKaryawan[Departemen]=\"IT\")*(tblKaryawan[Status Aktif]=\"Aktif\"))  -> semua baris karyawan IT aktif.",
        "  =SORTBY(tblKaryawan[Nama];tblKaryawan[Gaji Pokok];-1)  -> nama diurutkan dari gaji tertinggi.",
        "  =XLOOKUP(\"20190012\";tblKaryawan[NIK];tblKaryawan[Nama];\"tidak ada\")  -> pengganti INDEX/MATCH.",
        "  =LET(d;tblKaryawan[Departemen];s;tblKaryawan[Status Aktif];COUNTIFS(d;\"HR\";s;\"Aktif\"))  -> variabel dalam rumus.",
        "Alternatif Excel 2016: UNIQUE -> Remove Duplicates atau PivotTable; FILTER -> AutoFilter/Advanced Filter; XLOOKUP -> INDEX/MATCH.",
    ])
    wb.active = 0
    return rapikan(wb)


SCREENSHOT = [("Dashboard", "08_dashboard.png", "A1:Q50"), ("Perhitungan", "08_perhitungan.png", "A1:N47"),
              ("Data Karyawan", "08_data_karyawan.png", "A1:T32")]
