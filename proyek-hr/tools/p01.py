"""Proyek 01 - Database Karyawan (tingkat pemula)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment

from data_gen import buat_karyawan
from util import (DEPARTEMEN, FMT_RP, FMT_TGL, STATUS_KARYAWAN, WARNA_SUB, chart_bar, cf_rumus,
                  dropdown, font, header, judul, label_nilai, lebar, rapikan, ref, tabel, tulis_baris)

N = 80
KOLOM = ["No", "NIK", "Nama Lengkap", "Jenis Kelamin", "Tanggal Lahir", "Usia (Thn)", "Departemen",
         "Jabatan", "Status", "Tanggal Masuk", "Masa Kerja (Thn)", "Masa Kerja (Teks)", "Gaji Pokok",
         "Pendidikan", "Kota", "Email", "No. HP"]


def build():
    wb = Workbook()
    ws = wb.active
    ws.title = "Data Karyawan"
    data = buat_karyawan(N, seed=101)

    header(ws, 1, 1, KOLOM)
    for i, k in enumerate(data, start=1):
        r = i + 1
        baris = [
            i, k["NIK"], k["Nama"], k["JK"], k["TglLahir"],
            f'=DATEDIF(E{r},TODAY(),"Y")',
            k["Departemen"], k["Jabatan"], k["Status"], k["TglMasuk"],
            f'=DATEDIF(J{r},TODAY(),"Y")',
            f'=DATEDIF(J{r},TODAY(),"Y")&" thn "&DATEDIF(J{r},TODAY(),"YM")&" bln"',
            k["Gaji"], k["Pendidikan"], k["Kota"], k["Email"], k["HP"],
        ]
        fmts = [None, "@", None, None, FMT_TGL, "0", None, None, None, FMT_TGL, "0", None, FMT_RP, None, None, None, "@"]
        tulis_baris(ws, r, 1, baris, fmts)
    akhir = N + 1
    tabel(ws, f"A1:Q{akhir}", "tblKaryawan")
    ws.freeze_panes = "D2"
    lebar(ws, [5, 11, 24, 8, 13, 9, 13, 24, 9, 13, 10, 16, 15, 11, 12, 36, 15])
    dropdown(ws, f"G2:G{akhir}", "=Daftar!$A$2:$A$7")
    dropdown(ws, f"I2:I{akhir}", "=Daftar!$B$2:$B$4")
    dropdown(ws, f"D2:D{akhir}", ["L", "P"])
    # tandai gaji di bawah UMK Surabaya 2024 (sel acuan di Ringkasan)
    cf_rumus(ws, f"M2:M{akhir}", "M2<Ringkasan!$B$5")

    # ---- sheet daftar (sumber dropdown)
    wd = wb.create_sheet("Daftar")
    header(wd, 1, 1, ["Departemen", "Status", "Jenis Kelamin"])
    for i, d in enumerate(DEPARTEMEN, start=2):
        wd.cell(row=i, column=1, value=d)
    for i, s in enumerate(STATUS_KARYAWAN, start=2):
        wd.cell(row=i, column=2, value=s)
    wd["C2"], wd["C3"] = "L", "P"
    lebar(wd, [16, 12, 14])

    # ---- sheet ringkasan
    wr = wb.create_sheet("Ringkasan")
    judul(wr, "Ringkasan Database Karyawan", "Semua angka dihitung dengan rumus dari sheet Data Karyawan.")
    label_nilai(wr, 4, 1, "Tanggal acuan (hari ini)", "=TODAY()", FMT_TGL)
    label_nilai(wr, 5, 1, "UMK Surabaya 2024 (acuan)", 4725479, FMT_RP, fill="FFF2CC")
    label_nilai(wr, 6, 1, "Jumlah karyawan (COUNTA)", "=COUNTA('Data Karyawan'!B2:B1000)", "0")
    label_nilai(wr, 7, 1, "Karyawan laki-laki (COUNTIF)", "=COUNTIF('Data Karyawan'!D2:D1000,\"L\")", "0")
    label_nilai(wr, 8, 1, "Karyawan perempuan (COUNTIF)", "=COUNTIF('Data Karyawan'!D2:D1000,\"P\")", "0")
    label_nilai(wr, 9, 1, "Rata-rata gaji pokok (AVERAGE)", "=AVERAGE('Data Karyawan'!M2:M1000)", FMT_RP)
    label_nilai(wr, 10, 1, "Gaji tertinggi (MAX)", "=MAX('Data Karyawan'!M2:M1000)", FMT_RP)
    label_nilai(wr, 11, 1, "Gaji terendah (MIN)", "=MIN('Data Karyawan'!M2:M1000)", FMT_RP)
    label_nilai(wr, 12, 1, "Total gaji pokok per bulan (SUM)", "=SUM('Data Karyawan'!M2:M1000)", FMT_RP)
    label_nilai(wr, 13, 1, "Rata-rata usia (AVERAGE)", "=AVERAGE('Data Karyawan'!F2:F1000)", "0.0")
    label_nilai(wr, 14, 1, "Rata-rata masa kerja (thn)", "=AVERAGE('Data Karyawan'!K2:K1000)", "0.0")
    label_nilai(wr, 15, 1, "Gaji di bawah UMK (COUNTIF)", "=COUNTIF('Data Karyawan'!M2:M1000,\"<\"&B5)", "0")

    header(wr, 17, 1, ["Departemen", "Jumlah", "Tetap", "Kontrak", "Magang", "Rata-rata Gaji", "% dari Total"])
    for i, d in enumerate(DEPARTEMEN, start=18):
        tulis_baris(wr, i, 1, [
            d,
            f"=COUNTIF('Data Karyawan'!$G$2:$G$1000,A{i})",
            f"=COUNTIFS('Data Karyawan'!$G$2:$G$1000,$A{i},'Data Karyawan'!$I$2:$I$1000,C$17)",
            f"=COUNTIFS('Data Karyawan'!$G$2:$G$1000,$A{i},'Data Karyawan'!$I$2:$I$1000,D$17)",
            f"=COUNTIFS('Data Karyawan'!$G$2:$G$1000,$A{i},'Data Karyawan'!$I$2:$I$1000,E$17)",
            f"=IFERROR(AVERAGEIF('Data Karyawan'!$G$2:$G$1000,A{i},'Data Karyawan'!$M$2:$M$1000),0)",
            f"=IFERROR(B{i}/$B$24,0)",
        ], [None, "0", "0", "0", "0", FMT_RP, "0.0%"])
    tulis_baris(wr, 24, 1, ["Total", "=SUM(B18:B23)", "=SUM(C18:C23)", "=SUM(D18:D23)", "=SUM(E18:E23)",
                            "=AVERAGE('Data Karyawan'!M2:M1000)", "=SUM(G18:G23)"],
                [None, "0", "0", "0", "0", FMT_RP, "0.0%"])
    for c in range(1, 8):
        wr.cell(row=24, column=c).font = font(bold=True)
        wr.cell(row=24, column=c).fill = __import__("openpyxl").styles.PatternFill("solid", fgColor=WARNA_SUB)

    header(wr, 27, 1, ["Pendidikan", "Jumlah"])
    for i, p in enumerate(["SMA/SMK", "D3", "S1", "S2"], start=28):
        tulis_baris(wr, i, 1, [p, f"=COUNTIF('Data Karyawan'!$N$2:$N$1000,A{i})"], [None, "0"])
    lebar(wr, [34, 16, 10, 10, 10, 16, 12])
    wr["A4"].alignment = Alignment(wrap_text=False)
    chart_bar(wr, "I4", "Jumlah Karyawan per Departemen", ref(wr, 1, 18, 1, 23), ref(wr, 2, 17, 2, 23),
              judul_y="Orang", lebar_cm=14, tinggi_cm=8)
    chart_bar(wr, "I21", "Komposisi Status per Departemen", ref(wr, 1, 18, 1, 23), ref(wr, 3, 17, 5, 23),
              judul_y="Orang", lebar_cm=14, tinggi_cm=8, tumpuk=True)

    # ---- sheet petunjuk
    wp = wb.create_sheet("Petunjuk")
    from util import teks_panduan
    teks_panduan(wp, 1, [
        "# Petunjuk Penggunaan",
        "1. Sheet Data Karyawan adalah tabel master (Format as Table bernama tblKaryawan). Tambahkan karyawan baru di baris paling bawah; tabel otomatis memanjang.",
        "2. Kolom Usia, Masa Kerja, dan Masa Kerja (Teks) berisi rumus DATEDIF; jangan diketik manual.",
        "3. Kolom Departemen, Status, dan Jenis Kelamin memakai dropdown (Data Validation) yang sumbernya ada di sheet Daftar.",
        "4. Sheet Ringkasan menghitung COUNTA, COUNTIF, COUNTIFS, AVERAGE, AVERAGEIF, MAX, MIN, dan SUM secara otomatis.",
        "5. Sel kuning di Ringkasan!B5 (UMK) boleh diubah; gaji di bawah UMK akan disorot merah oleh Conditional Formatting.",
        "6. Rumus di file ini tampil dengan pemisah titik koma (;) pada Excel regional Indonesia, dan koma (,) pada regional Inggris.",
    ])
    wb.active = 0
    return rapikan(wb)


SCREENSHOT = [("Data Karyawan", "01_data_karyawan.png", "A1:O32"), ("Ringkasan", "01_ringkasan.png")]
