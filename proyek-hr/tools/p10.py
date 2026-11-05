"""Proyek 10 - Perencanaan Tenaga Kerja dan Otomasi (tingkat lanjut)."""

import random

from openpyxl import Workbook
from openpyxl.styles import PatternFill

from data_gen import buat_karyawan
from util import (DEPARTEMEN, FMT_PCT, FMT_RP, FMT_TGL, WARNA_INPUT, WARNA_SUB, chart_line, dropdown, font, header,
                  judul, label_nilai, lebar, rapikan, ref, tabel, teks_panduan, tulis_baris)

TAHUN = [2019, 2020, 2021, 2022, 2023, 2024]
PROYEKSI = [2025, 2026, 2027]
AWAL = {"HR": 9, "Finance": 14, "Operasional": 38, "Marketing": 20, "IT": 15, "Produksi": 52}
TUMBUH = {"HR": .04, "Finance": .03, "Operasional": .07, "Marketing": .06, "IT": .12, "Produksi": .05}


def build():
    rng = random.Random(1010)
    wb = Workbook()

    # ---------------- Historis
    wh = wb.active
    wh.title = "Historis"
    judul(wh, "Headcount Historis per Departemen 2019-2024", "Data akhir tahun. Baris Total dan kolom CAGR adalah rumus.")
    header(wh, 4, 1, ["Departemen"] + [str(t) for t in TAHUN] + ["CAGR 2019-2024", "Rata-rata Tambahan/Tahun"])
    for i, d in enumerate(DEPARTEMEN, start=5):
        hc = [AWAL[d]]
        for _ in TAHUN[1:]:
            hc.append(max(1, round(hc[-1] * (1 + TUMBUH[d] + rng.gauss(0, .03)))))
        tulis_baris(wh, i, 1, [d] + hc + [f"=(G{i}/B{i})^(1/5)-1", f"=(G{i}-B{i})/5"], [None] + ["0"] * 6 + [FMT_PCT, "0.0"])
    tulis_baris(wh, 11, 1, ["Total"] + [f"=SUM({c}5:{c}10)" for c in "BCDEFG"] + ["=(G11/B11)^(1/5)-1", "=(G11-B11)/5"],
                [None] + ["0"] * 6 + [FMT_PCT, "0.0"])
    for c in range(1, 10):
        wh.cell(row=11, column=c).font = font(bold=True)
        wh.cell(row=11, column=c).fill = PatternFill("solid", fgColor=WARNA_SUB)
    wh["A13"] = "Tahun (angka, untuk FORECAST/TREND)"
    wh["A13"].font = font(italic=True, color="595959")
    for j, t in enumerate(TAHUN):
        wh.cell(row=13, column=2 + j, value=t)
    lebar(wh, [16, 8, 8, 8, 8, 8, 8, 16, 20])
    chart_line(wh, "A15", "Tren Headcount Total 2019-2024", ref(wh, 2, 4, 7, 4), ref(wh, 1, 11, 7, 11), judul_y="Orang",
               lebar_cm=16, tinggi_cm=8)

    # ---------------- Proyeksi
    wp = wb.create_sheet("Proyeksi")
    judul(wp, "Proyeksi Headcount 2025-2027", "FORECAST.LINEAR dan TREND memakai tahun sebagai x dan headcount historis sebagai y.")
    header(wp, 4, 1, ["Departemen", "2024 Aktual", "2025 (FORECAST.LINEAR)", "2026 (FORECAST.LINEAR)", "2027 (FORECAST.LINEAR)",
                      "2025 (TREND)", "2026 (TREND)", "2027 (TREND)", "2027 (Metode CAGR)", "Tambahan 2025-2027"])
    for i, d in enumerate(DEPARTEMEN, start=5):
        hr = 5 + i - 5
        ys = f"Historis!$B${hr}:$G${hr}"
        xs = "Historis!$B$13:$G$13"
        tulis_baris(wp, i, 1, [
            d, f"=Historis!G{hr}",
            f"=ROUND(_xlfn.FORECAST.LINEAR(2025,{ys},{xs}),0)",
            f"=ROUND(_xlfn.FORECAST.LINEAR(2026,{ys},{xs}),0)",
            f"=ROUND(_xlfn.FORECAST.LINEAR(2027,{ys},{xs}),0)",
            f"=ROUND(TREND({ys},{xs},2025),0)", f"=ROUND(TREND({ys},{xs},2026),0)", f"=ROUND(TREND({ys},{xs},2027),0)",
            f"=ROUND(B{i}*(1+Historis!H{hr})^3,0)", f"=E{i}-B{i}",
        ], [None] + ["0"] * 9)
    tulis_baris(wp, 11, 1, ["Total"] + [f"=SUM({c}5:{c}10)" for c in "BCDEFGHIJ"], [None] + ["0"] * 9)
    for c in range(1, 11):
        wp.cell(row=11, column=c).font = font(bold=True)
        wp.cell(row=11, column=c).fill = PatternFill("solid", fgColor=WARNA_SUB)
    label_nilai(wp, 13, 1, "Asumsi turnover tahunan", 0.12, FMT_PCT, fill=WARNA_INPUT)
    header(wp, 15, 1, ["Departemen", "Headcount 2025", "Pengganti Turnover", "Penambahan Baru", "Kebutuhan Rekrutmen 2025"])
    for i, d in enumerate(DEPARTEMEN, start=16):
        tulis_baris(wp, i, 1, [d, f"=C{i - 11}", f"=ROUND(B{i - 11}*$B$13,0)", f"=MAX(C{i - 11}-B{i - 11},0)", f"=C{i}+D{i}"],
                    [None, "0", "0", "0", "0"])
    tulis_baris(wp, 22, 1, ["Total"] + [f"=SUM({c}16:{c}21)" for c in "BCDE"], [None] + ["0"] * 4)
    lebar(wp, [16, 12, 14, 14, 14, 12, 12, 12, 14, 14])
    chart_line(wp, "L4", "Proyeksi Headcount Total", ref(wp, 2, 4, 5, 4), ref(wp, 1, 11, 5, 11), judul_y="Orang",
               lebar_cm=15, tinggi_cm=8)

    # ---------------- Anggaran & What-If
    wa = wb.create_sheet("Anggaran Gaji")
    judul(wa, "Simulasi Anggaran Gaji 2025-2027 dan What-If Analysis", "Sel kuning = input. Skenario aktif dipilih di C6.")
    label_nilai(wa, 4, 2, "Kenaikan gaji tahunan", "=INDEX(C14:E14,MATCH(C6,C13:E13,0))", FMT_PCT)
    label_nilai(wa, 5, 2, "Pertumbuhan headcount tahunan", "=INDEX(C15:E15,MATCH(C6,C13:E13,0))", FMT_PCT)
    label_nilai(wa, 6, 2, "Skenario aktif", "Moderat", fill=WARNA_INPUT)
    dropdown(wa, "C6", ["Pesimis", "Moderat", "Optimis"])
    label_nilai(wa, 7, 2, "Bulan gaji per tahun (12 + THR)", 13, "0", fill=WARNA_INPUT)
    label_nilai(wa, 8, 2, "Batas anggaran 2025 (untuk Goal Seek)", 55_000_000_000, FMT_RP, fill=WARNA_INPUT)
    header(wa, 12, 2, ["Parameter Skenario", "Pesimis", "Moderat", "Optimis"])
    tulis_baris(wa, 13, 2, ["Nama", "Pesimis", "Moderat", "Optimis"])
    tulis_baris(wa, 14, 2, ["Kenaikan gaji", 0.03, 0.06, 0.09], [None, FMT_PCT, FMT_PCT, FMT_PCT])
    tulis_baris(wa, 15, 2, ["Pertumbuhan headcount", 0.0, 0.04, 0.08], [None, FMT_PCT, FMT_PCT, FMT_PCT])
    for r in (14, 15):
        for c in (3, 4, 5):
            wa.cell(row=r, column=c).fill = PatternFill("solid", fgColor=WARNA_INPUT)
    header(wa, 18, 2, ["Departemen", "Headcount 2024", "Rata-rata Gaji 2024", "Anggaran 2024", "Anggaran 2025",
                       "Anggaran 2026", "Anggaran 2027"])
    gaji_rata = {"HR": 8_100_000, "Finance": 8_600_000, "Operasional": 7_900_000, "Marketing": 8_900_000,
                 "IT": 9_800_000, "Produksi": 7_200_000}
    for i, d in enumerate(DEPARTEMEN, start=19):
        tulis_baris(wa, i, 2, [
            d, f"=Historis!G{i - 14}", gaji_rata[d], f"=C{i}*D{i}*$C$7",
            f"=E{i}*(1+$C$4)*(1+$C$5)", f"=F{i}*(1+$C$4)*(1+$C$5)", f"=G{i}*(1+$C$4)*(1+$C$5)",
        ], [None, "0", FMT_RP, FMT_RP, FMT_RP, FMT_RP, FMT_RP])
        wa.cell(row=i, column=4).fill = PatternFill("solid", fgColor=WARNA_INPUT)
    tulis_baris(wa, 25, 2, ["Total", "=SUM(C19:C24)", "=E25/C25/$C$7", "=SUM(E19:E24)", "=SUM(F19:F24)", "=SUM(G19:G24)",
                            "=SUM(H19:H24)"], [None, "0", FMT_RP, FMT_RP, FMT_RP, FMT_RP, FMT_RP])
    for c in range(2, 9):
        wa.cell(row=25, column=c).font = font(bold=True)
        wa.cell(row=25, column=c).fill = PatternFill("solid", fgColor=WARNA_SUB)
    label_nilai(wa, 26, 2, "Kenaikan anggaran 2025 vs 2024", "=F25/E25-1", FMT_PCT)
    label_nilai(wa, 27, 2, "Selisih terhadap batas anggaran", "=C8-F25", FMT_RP)
    label_nilai(wa, 28, 2, "Kenaikan gaji maksimal agar 2025 <= batas (jawaban Goal Seek)", "=C8/E25/(1+C5)-1", FMT_PCT)

    # Data Table 1 variabel (setara)
    header(wa, 30, 2, ["Kenaikan Gaji", "Anggaran 2025", "Anggaran 2027"])
    wa["B31"] = "Simulasi satu variabel (setara What-If > Data Table, column input cell = C4)"
    wa["B31"].font = font(italic=True, size=9, color="595959")
    for i, k in enumerate([0.03, 0.04, 0.05, 0.06, 0.07, 0.08, 0.09, 0.10], start=32):
        tulis_baris(wa, i, 2, [k, f"=$E$25*(1+B{i})*(1+$C$5)", f"=$E$25*((1+B{i})*(1+$C$5))^3"], [FMT_PCT, FMT_RP, FMT_RP])
    # Data Table 2 variabel (setara)
    wa["F30"] = "Simulasi dua variabel: anggaran 2025 (baris = kenaikan gaji, kolom = pertumbuhan headcount)"
    wa["F30"].font = font(italic=True, size=9, color="595959")
    header(wa, 31, 6, ["Kenaikan \\ HC"] + ["0%", "2%", "4%", "6%", "8%"])
    for j, g in enumerate([0, .02, .04, .06, .08]):
        wa.cell(row=31, column=7 + j, value=g).number_format = FMT_PCT
    for i, k in enumerate([0.03, 0.05, 0.07, 0.09, 0.11], start=32):
        wa.cell(row=i, column=6, value=k).number_format = FMT_PCT
        for j in range(5):
            col = chr(ord("G") + j)
            c = wa.cell(row=i, column=7 + j, value=f"=$E$25*(1+$F{i})*(1+{col}$31)")
            c.number_format = FMT_RP
    lebar(wa, [2, 44, 14, 18, 18, 18, 18, 18, 18, 18, 18])

    # ---------------- Data Karyawan (untuk makro)
    wd = wb.create_sheet("Data Karyawan")
    data = buat_karyawan(120, seed=1212)
    judul(wd, "Data Karyawan (bahan uji makro VBA)", "Sengaja berisi 3 masalah: NIK ganda, email kosong, dan gaji di bawah UMK (lihat Petunjuk).")
    header(wd, 4, 1, ["NIK", "Nama", "Departemen", "Jabatan", "Status", "Tanggal Masuk", "Gaji Pokok", "Email"])
    for i, k in enumerate(data):
        r = 5 + i
        tulis_baris(wd, r, 1, [k["NIK"], k["Nama"], k["Departemen"], k["Jabatan"], k["Status"], k["TglMasuk"], k["Gaji"],
                               k["Email"]], ["@", None, None, None, None, FMT_TGL, FMT_RP, None])
    wd["A40"] = wd["A39"].value  # NIK ganda
    wd["H61"] = None  # email kosong
    wd["G85"] = 4_200_000  # di bawah UMK
    tabel(wd, "A4:H124", "tblKaryawan")
    wd.freeze_panes = "C5"
    lebar(wd, [10, 24, 13, 24, 9, 13, 15, 40])
    label_nilai(wd, 126, 1, "UMK acuan (dipakai makro ValidasiData)", 4_725_479, FMT_RP, fill=WARNA_INPUT)
    label_nilai(wd, 127, 1, "Cek rumus: jumlah NIK ganda", '=SUMPRODUCT((COUNTIF(A5:A124,A5:A124)>1)*1)', "0")
    label_nilai(wd, 128, 1, "Cek rumus: email kosong", '=COUNTBLANK(H5:H124)', "0")
    label_nilai(wd, 129, 1, "Cek rumus: gaji di bawah UMK", '=COUNTIF(G5:G124,"<"&B126)', "0")

    # ---------------- Panduan
    wg = wb.create_sheet("Panduan What-If & VBA")
    teks_panduan(wg, 1, [
        "# Panduan What-If Analysis",
        "## Data Table satu variabel (Anggaran Gaji!B32:C39)",
        "1. Ketik nilai kenaikan gaji 3%..10% di B32:B39, dan di C31 ketik rumus =F25 (sel hasil).",
        "2. Blok B31:C39 > Data > What-If Analysis > Data Table. Column input cell: C4. OK. Excel mengisi {=TABLE(;C4)}.",
        "## Data Table dua variabel (F31:K36)",
        "1. Sel sudut F31 harus berisi rumus hasil (=F25); baris atas = pertumbuhan headcount, kolom kiri = kenaikan gaji.",
        "2. Blok F31:K36 > What-If Analysis > Data Table. Row input cell: C5, Column input cell: C4. OK.",
        "## Goal Seek",
        "1. Data > What-If Analysis > Goal Seek. Set cell: F25; To value: nilai di C8; By changing cell: C4. OK.",
        "2. Bandingkan dengan jawaban analitik di C28. Klik Cancel bila tidak ingin menimpa C4.",
        "## Scenario Manager",
        "1. Data > What-If Analysis > Scenario Manager > Add. Nama: Pesimis; Changing cells: C4:C5; isi 3% dan 0%.",
        "2. Ulangi untuk Moderat (6%, 4%) dan Optimis (9%, 8%). Klik Summary > Result cells: F25:H25 untuk laporan ringkasan.",
        "3. Workbook ini juga menyediakan versi rumus: pilih skenario di C6 dan nilai C4:C5 berubah lewat INDEX/MATCH.",
        "# Panduan Modul VBA (folder vba/)",
        "1. Buka file .xlsx, tekan Alt+F11 untuk membuka VBA Editor. Jika tab Developer belum ada: File > Options > Customize Ribbon > centang Developer.",
        "2. Di VBA Editor: File > Import File, pilih vba/ValidasiData.bas, lalu ulangi untuk RekapPerDepartemen.bas dan SlipGajiPDF.bas.",
        "3. Simpan sebagai Excel Macro-Enabled Workbook (.xlsm): File > Save As > tipe .xlsm. File .xlsx di repo tetap bebas makro.",
        "4. Jalankan: Developer > Macros > pilih nama makro > Run. Hasil ValidasiData: sel bermasalah di sheet Data Karyawan disorot dan dirangkum di sheet 'Hasil Validasi'.",
        "5. RekapPerDepartemen menyalin baris tiap departemen ke sheet baru bernama departemennya. SlipGajiPDF dipakai pada workbook proyek 04 (sheet Slip Gaji, sel C6).",
        "6. Jika makro diblokir: klik kanan file > Properties > Unblock, atau File > Options > Trust Center > Macro Settings.",
    ])
    wb.active = wb.sheetnames.index("Anggaran Gaji")
    return rapikan(wb)


SCREENSHOT = [("Anggaran Gaji", "10_anggaran_gaji.png", "A1:K40"), ("Proyeksi", "10_proyeksi.png", "A1:T24")]
