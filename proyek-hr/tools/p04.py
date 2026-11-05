"""Proyek 04 - Penggajian dan Slip Gaji (tingkat menengah). Periode: Maret 2024."""

import random

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, PatternFill, Side

from data_gen import buat_karyawan
from util import (DEPARTEMEN, FMT_PCT, FMT_RP, WARNA_HEADER, WARNA_INPUT, WARNA_SUB, chart_bar, dropdown, font,
                  header, judul, label_nilai, lebar, rapikan, ref, tabel, teks_panduan, tulis_baris)

N = 70
TUNJ_JABATAN = [("Staff", 300_000), ("Senior Staff", 600_000), ("Supervisor", 1_250_000), ("Manager", 3_000_000)]
PTKP = [("TK/0", "A", 54_000_000), ("TK/1", "A", 58_500_000), ("K/0", "A", 58_500_000),
        ("TK/2", "B", 63_000_000), ("TK/3", "B", 67_500_000), ("K/1", "B", 63_000_000), ("K/2", "B", 67_500_000),
        ("K/3", "C", 72_000_000)]
# Tabel TER bulanan yang disederhanakan (batas bawah penghasilan bruto bulanan -> tarif). Sumber: cuplikan
# lampiran PP 58/2023, dibulatkan dan dipotong sampai Rp43 juta untuk keperluan belajar.
TER = {
    "A": [(0, 0), (5_400_001, .0025), (5_650_001, .005), (5_950_001, .0075), (6_300_001, .01), (6_750_001, .0125),
          (7_500_001, .015), (8_550_001, .0175), (9_650_001, .02), (10_050_001, .0225), (10_350_001, .025),
          (10_700_001, .03), (11_050_001, .035), (11_600_001, .04), (12_500_001, .05), (13_750_001, .06),
          (15_100_001, .07), (16_950_001, .08), (19_750_001, .09), (24_150_001, .10), (26_450_001, .11),
          (28_000_001, .12), (30_050_001, .13), (32_400_001, .14), (35_400_001, .15)],
    "B": [(0, 0), (6_200_001, .0025), (6_500_001, .005), (6_850_001, .0075), (7_300_001, .01), (9_200_001, .015),
          (10_750_001, .02), (11_250_001, .025), (11_600_001, .03), (12_600_001, .04), (13_600_001, .05),
          (14_950_001, .06), (16_400_001, .07), (18_450_001, .08), (21_850_001, .09), (26_000_001, .10),
          (27_700_001, .11), (29_350_001, .12), (31_450_001, .13), (33_950_001, .14), (37_100_001, .15)],
    "C": [(0, 0), (6_600_001, .0025), (6_950_001, .005), (7_350_001, .0075), (7_800_001, .01), (8_850_001, .0125),
          (9_800_001, .015), (10_950_001, .0175), (11_200_001, .02), (12_050_001, .03), (12_950_001, .04),
          (14_150_001, .05), (15_550_001, .06), (17_050_001, .07), (19_500_001, .08), (22_700_001, .09),
          (26_600_001, .10), (28_100_001, .11), (30_100_001, .12), (32_600_001, .13), (35_400_001, .14),
          (38_900_001, .15)],
}
BANK = ["Mandiri", "BCA", "BNI", "BRI"]


def build():
    rng = random.Random(404)
    wb = Workbook()
    data = buat_karyawan(N, seed=101)[:N]

    # ------------------------------------------------ Parameter
    wp = wb.active
    wp.title = "Parameter"
    judul(wp, "Parameter Penggajian", "Sel kuning adalah input yang boleh diubah. Semua sheet lain merujuk ke sini.")
    param = [
        ("Periode gaji", "Maret 2024", None),
        ("Hari kerja dalam periode", 18, "0"),
        ("UMK Surabaya 2024", 4_725_479, FMT_RP),
        ("Tunjangan makan per hari hadir", 25_000, FMT_RP),
        ("Tunjangan transport per hari hadir", 20_000, FMT_RP),
        ("Pembagi upah lembur per jam (jam kerja sebulan)", 173, "0"),
        ("Pengali upah lembur (disederhanakan, rata)", 1.5, "0.0"),
        ("Iuran BPJS Kesehatan ditanggung karyawan", 0.01, FMT_PCT),
        ("Batas atas upah BPJS Kesehatan", 12_000_000, FMT_RP),
        ("Iuran JHT ditanggung karyawan", 0.02, FMT_PCT),
        ("Iuran Jaminan Pensiun ditanggung karyawan", 0.01, FMT_PCT),
        ("Batas atas upah Jaminan Pensiun 2024", 10_042_300, FMT_RP),
        ("Tambahan PPh 21 bila tanpa NPWP", 0.20, FMT_PCT),
    ]
    for i, (lab, val, fmt) in enumerate(param, start=4):
        label_nilai(wp, i, 1, lab, val, fmt, fill=WARNA_INPUT)
    # B4..B16 ; nama sel untuk dokumentasi
    header(wp, 19, 1, ["Level Jabatan", "Tunjangan Jabatan"])
    for i, (lv, t) in enumerate(TUNJ_JABATAN, start=20):
        tulis_baris(wp, i, 1, [lv, t], [None, FMT_RP])
        wp.cell(row=i, column=2).fill = PatternFill("solid", fgColor=WARNA_INPUT)
    header(wp, 26, 1, ["Status PTKP", "Kategori TER", "PTKP Setahun"])
    for i, (st, kat, ptkp) in enumerate(PTKP, start=27):
        tulis_baris(wp, i, 1, [st, kat, ptkp], [None, None, FMT_RP])
    # Tabel TER
    header(wp, 19, 4, ["TER A: Batas Bawah Bruto", "Tarif A"])
    header(wp, 19, 7, ["TER B: Batas Bawah Bruto", "Tarif B"])
    header(wp, 19, 10, ["TER C: Batas Bawah Bruto", "Tarif C"])
    for kat, col in (("A", 4), ("B", 7), ("C", 10)):
        for i, (bb, tr) in enumerate(TER[kat], start=20):
            tulis_baris(wp, i, col, [bb, tr], [FMT_RP, "0.00%"])
    wp["D18"] = ("Tabel TER (Tarif Efektif Rata-rata) bulanan di bawah ini adalah cuplikan yang disederhanakan dan "
                 "dibulatkan dari PP 58/2023 untuk keperluan belajar; payroll nyata wajib memakai tabel resmi lengkap.")
    wp["D18"].font = font(italic=True, size=9, color="C00000")
    lebar(wp, [46, 16, 14, 24, 10, 2, 24, 10, 2, 24, 10])
    wp.freeze_panes = "A4"

    # ------------------------------------------------ Master Karyawan
    wm = wb.create_sheet("Master Karyawan")
    judul(wm, "Master Data Gaji Karyawan", "Tunjangan jabatan ditarik dengan VLOOKUP dari sheet Parameter.")
    header(wm, 4, 1, ["NIK", "Nama", "Departemen", "Jabatan", "Level", "Status PTKP", "NPWP", "Gaji Pokok",
                      "Tunjangan Jabatan", "Bank", "No. Rekening"])
    for i, k in enumerate(data):
        r = 5 + i
        ptkp = rng.choices([p[0] for p in PTKP], weights=[30, 8, 15, 5, 3, 18, 14, 7])[0]
        tulis_baris(wm, r, 1, [
            k["NIK"], k["Nama"], k["Departemen"], k["Jabatan"], k["Level"], ptkp,
            rng.choices(["Ya", "Tidak"], weights=[85, 15])[0], k["Gaji"],
            f"=VLOOKUP(E{r},Parameter!$A$20:$B$23,2,FALSE)",
            rng.choice(BANK), f"{rng.randint(100, 999)}{rng.randint(1000000, 9999999)}",
        ], ["@", None, None, None, None, None, None, FMT_RP, FMT_RP, None, "@"])
    akhir = 4 + N
    tabel(wm, f"A4:K{akhir}", "tblMasterGaji")
    wm.freeze_panes = "C5"
    lebar(wm, [10, 24, 12, 24, 12, 11, 7, 15, 15, 9, 13])
    dropdown(wm, f"F5:F{akhir}", "=Parameter!$A$27:$A$34")
    dropdown(wm, f"G5:G{akhir}", ["Ya", "Tidak"])
    dropdown(wm, f"E5:E{akhir}", "=Parameter!$A$20:$A$23")

    # ------------------------------------------------ Kehadiran & Lembur
    wl = wb.create_sheet("Kehadiran & Lembur")
    judul(wl, "Kehadiran dan Lembur - Maret 2024", "Hari hadir dan jam lembur diisi dari rekap absensi; tarif lembur = gaji pokok / 173.")
    header(wl, 4, 1, ["NIK", "Nama", "Hari Kerja", "Hari Hadir", "Jam Lembur", "Gaji Pokok", "Tarif Lembur / Jam",
                      "Upah Lembur"])
    for i, k in enumerate(data):
        r = 5 + i
        hadir = rng.choices([18, 17, 16, 15], weights=[60, 25, 10, 5])[0]
        lembur = 0 if k["Level"] == "Manager" else rng.choices([0, 4, 8, 12, 16, 20], weights=[35, 20, 20, 12, 8, 5])[0]
        tulis_baris(wl, r, 1, [
            f"='Master Karyawan'!A{r}",
            f"=INDEX('Master Karyawan'!$B$5:$B${akhir},MATCH(A{r},'Master Karyawan'!$A$5:$A${akhir},0))",
            "=Parameter!$B$5", hadir, lembur,
            f"=INDEX('Master Karyawan'!$H$5:$H${akhir},MATCH(A{r},'Master Karyawan'!$A$5:$A${akhir},0))",
            f"=ROUND(F{r}/Parameter!$B$9,0)",
            f"=E{r}*G{r}*Parameter!$B$10",
        ], ["@", None, "0", "0", "0", FMT_RP, FMT_RP, FMT_RP])
        for c in (4, 5):
            wl.cell(row=r, column=c).fill = PatternFill("solid", fgColor=WARNA_INPUT)
    tabel(wl, f"A4:H{akhir}", "tblLembur")
    wl.freeze_panes = "C5"
    lebar(wl, [10, 24, 10, 10, 10, 15, 15, 15])

    # ------------------------------------------------ Perhitungan Gaji
    wg = wb.create_sheet("Perhitungan Gaji")
    judul(wg, "Perhitungan Gaji - Maret 2024", "Bruto, potongan BPJS, PPh 21 metode TER (disederhanakan), dan gaji bersih.")
    head = ["NIK", "Nama", "Departemen", "Level", "Gaji Pokok", "Tunj. Jabatan", "Tunj. Makan", "Tunj. Transport",
            "Upah Lembur", "Gaji Bruto", "BPJS Kesehatan 1%", "JHT 2%", "JP 1%", "Kategori TER", "Tarif TER",
            "PPh 21", "Total Potongan", "Gaji Bersih (THP)"]
    header(wg, 4, 1, head)
    for i, k in enumerate(data):
        r = 5 + i
        tulis_baris(wg, r, 1, [
            f"='Master Karyawan'!A{r}",
            f"=INDEX('Master Karyawan'!$B$5:$B${akhir},MATCH($A{r},'Master Karyawan'!$A$5:$A${akhir},0))",
            f"=INDEX('Master Karyawan'!$C$5:$C${akhir},MATCH($A{r},'Master Karyawan'!$A$5:$A${akhir},0))",
            f"=INDEX('Master Karyawan'!$E$5:$E${akhir},MATCH($A{r},'Master Karyawan'!$A$5:$A${akhir},0))",
            f"=INDEX('Master Karyawan'!$H$5:$H${akhir},MATCH($A{r},'Master Karyawan'!$A$5:$A${akhir},0))",
            f"=INDEX('Master Karyawan'!$I$5:$I${akhir},MATCH($A{r},'Master Karyawan'!$A$5:$A${akhir},0))",
            f"=INDEX('Kehadiran & Lembur'!$D$5:$D${akhir},MATCH($A{r},'Kehadiran & Lembur'!$A$5:$A${akhir},0))*Parameter!$B$7",
            f"=INDEX('Kehadiran & Lembur'!$D$5:$D${akhir},MATCH($A{r},'Kehadiran & Lembur'!$A$5:$A${akhir},0))*Parameter!$B$8",
            f"=SUMIF('Kehadiran & Lembur'!$A$5:$A${akhir},$A{r},'Kehadiran & Lembur'!$H$5:$H${akhir})",
            f"=SUM(E{r}:I{r})",
            f"=ROUND(MIN(E{r}+F{r},Parameter!$B$12)*Parameter!$B$11,0)",
            f"=ROUND((E{r}+F{r})*Parameter!$B$13,0)",
            f"=ROUND(MIN(E{r}+F{r},Parameter!$B$15)*Parameter!$B$14,0)",
            f"=INDEX(Parameter!$B$27:$B$34,MATCH(INDEX('Master Karyawan'!$F$5:$F${akhir},MATCH($A{r},'Master Karyawan'!$A$5:$A${akhir},0)),Parameter!$A$27:$A$34,0))",
            f'=IF(N{r}="A",VLOOKUP(J{r},Parameter!$D$20:$E$44,2,TRUE),IF(N{r}="B",VLOOKUP(J{r},Parameter!$G$20:$H$40,2,TRUE),VLOOKUP(J{r},Parameter!$J$20:$K$41,2,TRUE)))',
            f"=ROUND(J{r}*O{r}*IF(INDEX('Master Karyawan'!$G$5:$G${akhir},MATCH($A{r},'Master Karyawan'!$A$5:$A${akhir},0))=\"Tidak\",1+Parameter!$B$16,1),0)",
            f"=SUM(K{r}:M{r})+P{r}",
            f"=J{r}-Q{r}",
        ], ["@", None, None, None] + [FMT_RP] * 9 + [None, "0.00%", FMT_RP, FMT_RP, FMT_RP])
    tulis_baris(wg, akhir + 1, 1, ["Total", "", "", ""] + [f"=SUM({c}5:{c}{akhir})" for c in "EFGHIJKLM"] + ["", ""]
                + [f"=SUM({c}5:{c}{akhir})" for c in "PQR"], [None] * 4 + [FMT_RP] * 9 + [None, None] + [FMT_RP] * 3)
    for c in range(1, 19):
        wg.cell(row=akhir + 1, column=c).font = font(bold=True)
        wg.cell(row=akhir + 1, column=c).fill = PatternFill("solid", fgColor=WARNA_SUB)
    wg.freeze_panes = "C5"
    wg.auto_filter.ref = f"A4:R{akhir}"
    lebar(wg, [10, 22, 12, 11, 14, 13, 12, 12, 13, 14, 13, 12, 12, 9, 9, 13, 14, 15])

    # ------------------------------------------------ Rekap Departemen
    wr = wb.create_sheet("Rekap Departemen")
    judul(wr, "Rekap Biaya Gaji per Departemen - Maret 2024", "SUMIFS dan COUNTIF terhadap sheet Perhitungan Gaji.")
    header(wr, 4, 1, ["Departemen", "Jumlah Karyawan", "Gaji Pokok", "Total Tunjangan", "Upah Lembur", "Gaji Bruto",
                      "Potongan BPJS", "PPh 21", "Gaji Bersih", "Rata-rata THP", "% dari Total Bruto"])
    for i, d in enumerate(DEPARTEMEN, start=5):
        pg = "'Perhitungan Gaji'"
        tulis_baris(wr, i, 1, [
            d, f"=COUNTIF({pg}!$C$5:$C${akhir},A{i})",
            f"=SUMIFS({pg}!$E$5:$E${akhir},{pg}!$C$5:$C${akhir},$A{i})",
            f"=SUMIFS({pg}!$F$5:$F${akhir},{pg}!$C$5:$C${akhir},$A{i})+SUMIFS({pg}!$G$5:$G${akhir},{pg}!$C$5:$C${akhir},$A{i})+SUMIFS({pg}!$H$5:$H${akhir},{pg}!$C$5:$C${akhir},$A{i})",
            f"=SUMIFS({pg}!$I$5:$I${akhir},{pg}!$C$5:$C${akhir},$A{i})",
            f"=SUMIFS({pg}!$J$5:$J${akhir},{pg}!$C$5:$C${akhir},$A{i})",
            f"=SUMIFS({pg}!$K$5:$K${akhir},{pg}!$C$5:$C${akhir},$A{i})+SUMIFS({pg}!$L$5:$L${akhir},{pg}!$C$5:$C${akhir},$A{i})+SUMIFS({pg}!$M$5:$M${akhir},{pg}!$C$5:$C${akhir},$A{i})",
            f"=SUMIFS({pg}!$P$5:$P${akhir},{pg}!$C$5:$C${akhir},$A{i})",
            f"=SUMIFS({pg}!$R$5:$R${akhir},{pg}!$C$5:$C${akhir},$A{i})",
            f"=IFERROR(I{i}/B{i},0)", f"=IFERROR(F{i}/$F$11,0)",
        ], [None, "0"] + [FMT_RP] * 8 + [FMT_PCT])
    tulis_baris(wr, 11, 1, ["Total", "=SUM(B5:B10)"] + [f"=SUM({c}5:{c}10)" for c in "CDEFGHI"] +
                ["=IFERROR(I11/B11,0)", "=SUM(K5:K10)"], [None, "0"] + [FMT_RP] * 8 + [FMT_PCT])
    for c in range(1, 12):
        wr.cell(row=11, column=c).font = font(bold=True)
        wr.cell(row=11, column=c).fill = PatternFill("solid", fgColor=WARNA_SUB)
    lebar(wr, [16, 10, 16, 16, 15, 16, 15, 14, 16, 15, 12])
    chart_bar(wr, "A14", "Gaji Bruto vs Gaji Bersih per Departemen", ref(wr, 1, 5, 1, 10), ref(wr, 6, 4, 6, 10),
              judul_y="Rupiah", lebar_cm=15, tinggi_cm=8)
    wr._charts[-1].add_data(ref(wr, 9, 4, 9, 10), titles_from_data=True)
    chart_bar(wr, "G14", "Komposisi Potongan per Departemen", ref(wr, 1, 5, 1, 10), ref(wr, 7, 4, 8, 10),
              judul_y="Rupiah", lebar_cm=15, tinggi_cm=8, tumpuk=True)

    # ------------------------------------------------ Slip Gaji
    wsl = wb.create_sheet("Slip Gaji")
    tebal = Side(style="medium", color=WARNA_HEADER)
    wsl.merge_cells("B2:F2")
    wsl["B2"] = "PT CONTOH PERUSAHAAN INDONESIA"
    wsl["B2"].font = font(bold=True, size=14, color=WARNA_HEADER)
    wsl["B2"].alignment = Alignment(horizontal="center")
    wsl.merge_cells("B3:F3")
    wsl["B3"] = "SLIP GAJI KARYAWAN"
    wsl["B3"].font = font(bold=True, size=12)
    wsl["B3"].alignment = Alignment(horizontal="center")
    wsl.merge_cells("B4:F4")
    wsl["B4"] = '="Periode: "&Parameter!B4'
    wsl["B4"].alignment = Alignment(horizontal="center")
    wsl["B6"] = "Pilih nama karyawan:"
    wsl["B6"].font = font(bold=True)
    wsl["C6"] = data[7]["Nama"]
    wsl["C6"].fill = PatternFill("solid", fgColor=WARNA_INPUT)
    wsl["C6"].font = font(bold=True)
    wsl.merge_cells("C6:D6")
    dropdown(wsl, "C6", f"='Master Karyawan'!$B$5:$B${akhir}", pesan="Pilih nama dari daftar karyawan.")
    mk = "'Master Karyawan'"
    pg = "'Perhitungan Gaji'"
    cari = f"MATCH($C$6,{mk}!$B$5:$B${akhir},0)"
    info = [("NIK", f"=INDEX({mk}!$A$5:$A${akhir},{cari})", "@"),
            ("Departemen", f"=INDEX({mk}!$C$5:$C${akhir},{cari})", None),
            ("Jabatan", f"=INDEX({mk}!$D$5:$D${akhir},{cari})", None),
            ("Status PTKP", f"=INDEX({mk}!$F$5:$F${akhir},{cari})", None),
            ("Bank / Rekening", f'=INDEX({mk}!$J$5:$J${akhir},{cari})&" - "&INDEX({mk}!$K$5:$K${akhir},{cari})', None)]
    for i, (lab, rum, fmt) in enumerate(info, start=8):
        wsl.cell(row=i, column=2, value=lab).font = font()
        c = wsl.cell(row=i, column=3, value=rum)
        c.font = font(bold=True)
        if fmt:
            c.number_format = fmt
    cari_g = f"MATCH($C$6,{pg}!$B$5:$B${akhir},0)"
    header(wsl, 14, 2, ["PENERIMAAN", "Jumlah"])
    header(wsl, 14, 5, ["POTONGAN", "Jumlah"])
    terima = [("Gaji Pokok", "E"), ("Tunjangan Jabatan", "F"), ("Tunjangan Makan", "G"), ("Tunjangan Transport", "H"),
              ("Upah Lembur", "I")]
    potong = [("BPJS Kesehatan (1%)", "K"), ("JHT (2%)", "L"), ("Jaminan Pensiun (1%)", "M"), ("PPh 21", "P")]
    for i, (lab, col) in enumerate(terima, start=15):
        tulis_baris(wsl, i, 2, [lab, f"=INDEX({pg}!${col}$5:${col}${akhir},{cari_g})"], [None, FMT_RP])
    for i, (lab, col) in enumerate(potong, start=15):
        tulis_baris(wsl, i, 5, [lab, f"=INDEX({pg}!${col}$5:${col}${akhir},{cari_g})"], [None, FMT_RP])
    tulis_baris(wsl, 20, 2, ["Total Penerimaan", "=SUM(C15:C19)"], [None, FMT_RP])
    tulis_baris(wsl, 20, 5, ["Total Potongan", "=SUM(F15:F18)"], [None, FMT_RP])
    for c in ("B20", "C20", "E20", "F20"):
        wsl[c].font = font(bold=True)
        wsl[c].fill = PatternFill("solid", fgColor=WARNA_SUB)
    wsl["E19"].border = wsl["F19"].border = Border(left=Side(style="thin", color="BFBFBF"), right=Side(style="thin", color="BFBFBF"))
    wsl.merge_cells("B22:E22")
    wsl["B22"] = "GAJI BERSIH DITERIMA (TAKE HOME PAY)"
    wsl["B22"].font = font(bold=True, size=12, color="FFFFFF")
    wsl["B22"].fill = PatternFill("solid", fgColor=WARNA_HEADER)
    wsl["B22"].alignment = Alignment(horizontal="right", vertical="center")
    wsl["F22"] = "=C20-F20"
    wsl["F22"].number_format = FMT_RP
    wsl["F22"].font = font(bold=True, size=12, color="FFFFFF")
    wsl["F22"].fill = PatternFill("solid", fgColor=WARNA_HEADER)
    wsl.row_dimensions[22].height = 24
    wsl["B23"] = '="Cek: sama dengan kolom THP di Perhitungan Gaji? "&IF(F22=INDEX(' + pg + f'!$R$5:$R${akhir},{cari_g}),"Ya","Tidak")'
    wsl["B23"].font = font(italic=True, size=9, color="595959")
    wsl["B25"] = "Hari hadir:"
    wsl["C25"] = f"=INDEX('Kehadiran & Lembur'!$D$5:$D${akhir},MATCH($C$6,'Kehadiran & Lembur'!$B$5:$B${akhir},0))&\" dari \"&Parameter!B5&\" hari kerja\""
    wsl["B26"] = "Jam lembur:"
    wsl["C26"] = f"=INDEX('Kehadiran & Lembur'!$E$5:$E${akhir},MATCH($C$6,'Kehadiran & Lembur'!$B$5:$B${akhir},0))&\" jam\""
    wsl["B28"] = "Surabaya, 25 Maret 2024"
    wsl["E28"] = "Mengetahui,"
    wsl["E32"] = "Manager HR"
    for rr in range(2, 33):
        for cc in range(2, 7):
            cell = wsl.cell(row=rr, column=cc)
            b = Border(left=tebal if cc == 2 else cell.border.left, right=tebal if cc == 6 else cell.border.right,
                       top=tebal if rr == 2 else cell.border.top, bottom=tebal if rr == 32 else cell.border.bottom)
            cell.border = b
    lebar(wsl, [3, 24, 18, 4, 24, 18])
    wsl.sheet_view.showGridLines = False

    wpt = wb.create_sheet("Petunjuk")
    teks_panduan(wpt, 1, [
        "# Petunjuk Penggunaan",
        "1. Ubah parameter (sel kuning) di sheet Parameter: hari kerja, tarif tunjangan, iuran BPJS, dan tabel TER.",
        "2. Master Karyawan: data gaji pokok, status PTKP, dan NPWP. Tunjangan jabatan dihitung VLOOKUP dari level.",
        "3. Kehadiran & Lembur: isi hari hadir dan jam lembur (sel kuning); upah lembur = jam x (gaji pokok/173) x 1,5.",
        "4. Perhitungan Gaji: seluruh kolom berupa rumus (INDEX/MATCH, SUMIF, MIN, ROUND, VLOOKUP approximate untuk TER).",
        "5. PPh 21 memakai metode TER bulanan yang DISEDERHANAKAN (lihat catatan merah di Parameter). Pada bulan Desember Excel nyata perlu perhitungan tahunan tarif progresif Pasal 17.",
        "6. Slip Gaji: pilih nama di sel C6 (dropdown); seluruh isi slip berubah lewat INDEX/MATCH. Cetak dengan Ctrl+P.",
        "7. Rekap Departemen memakai SUMIFS untuk melaporkan biaya gaji per departemen beserta chart.",
    ])
    wb.active = wb.sheetnames.index("Slip Gaji")
    return rapikan(wb)


SCREENSHOT = [("Slip Gaji", "04_slip_gaji.png", "A1:G33"), ("Perhitungan Gaji", "04_perhitungan_gaji.png", "A1:R30"),
              ("Rekap Departemen", "04_rekap_departemen.png", "A1:K32")]
