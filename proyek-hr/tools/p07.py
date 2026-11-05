"""Proyek 07 - Penilaian Kinerja KPI (tingkat menengah). Periode penilaian: Semester 1 2024."""

import random

from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill

from data_gen import buat_karyawan
from util import (BORDER, DEPARTEMEN, FMT_PCT, WARNA_HEADER, WARNA_INPUT, WARNA_SUB, cf_sama_dengan, cf_skala_warna,
                  chart_bar, dropdown, font, header, judul, label_nilai, lebar, rapikan, ref, tabel, teks_panduan,
                  tulis_baris, validasi_angka)

N = 80
KPI = {
    "Staff": [("Kualitas hasil kerja", .30), ("Ketepatan waktu", .25), ("Kedisiplinan & kehadiran", .20),
              ("Kerja sama tim", .15), ("Inisiatif", .10)],
    "Senior Staff": [("Kualitas hasil kerja", .30), ("Ketepatan waktu", .20), ("Pemecahan masalah", .20),
                     ("Kerja sama tim", .15), ("Mentoring junior", .15)],
    "Supervisor": [("Pencapaian target tim", .35), ("Kualitas supervisi", .20), ("Pengembangan anggota", .15),
                   ("Komunikasi & koordinasi", .15), ("Efisiensi proses", .15)],
    "Manager": [("Pencapaian target departemen", .40), ("Pengelolaan anggaran", .20), ("Kepemimpinan", .15),
                ("Inovasi & perbaikan", .15), ("Kolaborasi lintas fungsi", .10)],
}
RATING = [(4.5, "Istimewa"), (3.5, "Baik"), (2.5, "Cukup"), (0, "Perlu Perbaikan")]
BOX = {("Tinggi", "Tinggi"): "Bintang (Star)", ("Tinggi", "Sedang"): "Pemain Kunci", ("Tinggi", "Rendah"): "Ahli Andal",
       ("Sedang", "Tinggi"): "Calon Bintang", ("Sedang", "Sedang"): "Inti Organisasi", ("Sedang", "Rendah"): "Kontributor Stabil",
       ("Rendah", "Tinggi"): "Potensi Belum Terwujud", ("Rendah", "Sedang"): "Perlu Pendampingan",
       ("Rendah", "Rendah"): "Perlu Tindakan"}


def build():
    rng = random.Random(707)
    wb = Workbook()
    data = buat_karyawan(N, seed=101)[:N]

    # ---------------- Master KPI
    wm = wb.active
    wm.title = "Master KPI"
    judul(wm, "Master Indikator Kinerja (KPI) dan Bobot per Level Jabatan", "Bobot tiap level harus berjumlah 100%. Sel kuning dapat diubah.")
    header(wm, 4, 1, ["Level", "KPI 1", "Bobot 1", "KPI 2", "Bobot 2", "KPI 3", "Bobot 3", "KPI 4", "Bobot 4", "KPI 5",
                      "Bobot 5", "Total Bobot", "Cek"])
    for i, (lv, items) in enumerate(KPI.items(), start=5):
        baris = [lv]
        for nama, bobot in items:
            baris += [nama, bobot]
        baris += [f"=C{i}+E{i}+G{i}+I{i}+K{i}", f'=IF(ROUND(L{i},4)=1,"OK","Bobot harus 100%")']
        tulis_baris(wm, i, 1, baris, [None] + [None, FMT_PCT] * 5 + [FMT_PCT, None])
        for c in (3, 5, 7, 9, 11):
            wm.cell(row=i, column=c).fill = PatternFill("solid", fgColor=WARNA_INPUT)
    header(wm, 11, 1, ["Skor", "Arti", "", "Batas Bawah Nilai Akhir", "Rating"])
    skala = [(1, "Jauh di bawah harapan"), (2, "Di bawah harapan"), (3, "Sesuai harapan"), (4, "Melampaui harapan"),
             (5, "Jauh melampaui harapan")]
    for i, (s, a) in enumerate(skala, start=12):
        tulis_baris(wm, i, 1, [s, a])
    for i, (b, r_) in enumerate(reversed(RATING), start=12):
        tulis_baris(wm, i, 4, [b, r_], ["0.0", None])
    wm["D17"] = "Batas performa 9-box: Rendah < 3,0 ; Sedang 3,0 - 3,99 ; Tinggi >= 4,0 (sama untuk potensi)."
    wm["D17"].font = font(italic=True, size=9, color="595959")
    lebar(wm, [13, 26, 9, 26, 9, 26, 9, 26, 9, 26, 9, 11, 16])

    # ---------------- Penilaian
    wp = wb.create_sheet("Penilaian")
    judul(wp, "Penilaian Kinerja Semester 1 2024", "Skor 1-5 per KPI (sel kuning). Nilai akhir = SUMPRODUCT(skor; bobot). Rating dengan IFS, peringkat dengan RANK.EQ.")
    head = ["NIK", "Nama", "Departemen", "Jabatan", "Level", "Skor 1", "Skor 2", "Skor 3", "Skor 4", "Skor 5",
            "Bobot 1", "Bobot 2", "Bobot 3", "Bobot 4", "Bobot 5", "Nilai Akhir", "Rating", "Peringkat",
            "Potensi (1-5)", "Kategori Performa", "Kategori Potensi", "Kuadran 9-Box", "Atasan Penilai"]
    header(wp, 4, 1, head)
    for i, k in enumerate(data):
        r = 5 + i
        dasar = rng.gauss(3.6, 0.55)
        skor = [max(1, min(5, round(dasar + rng.gauss(0, 0.6)))) for _ in range(5)]
        potensi = max(1, min(5, round(dasar + rng.gauss(0, 0.9))))
        tulis_baris(wp, r, 1, [
            k["NIK"], k["Nama"], k["Departemen"], k["Jabatan"], k["Level"]] + skor + [
            f"=INDEX('Master KPI'!$C$5:$C$8,MATCH($E{r},'Master KPI'!$A$5:$A$8,0))",
            f"=INDEX('Master KPI'!$E$5:$E$8,MATCH($E{r},'Master KPI'!$A$5:$A$8,0))",
            f"=INDEX('Master KPI'!$G$5:$G$8,MATCH($E{r},'Master KPI'!$A$5:$A$8,0))",
            f"=INDEX('Master KPI'!$I$5:$I$8,MATCH($E{r},'Master KPI'!$A$5:$A$8,0))",
            f"=INDEX('Master KPI'!$K$5:$K$8,MATCH($E{r},'Master KPI'!$A$5:$A$8,0))",
            f"=ROUND(SUMPRODUCT(F{r}:J{r},K{r}:O{r}),2)",
            f'=_xlfn.IFS(P{r}>=4.5,"Istimewa",P{r}>=3.5,"Baik",P{r}>=2.5,"Cukup",TRUE,"Perlu Perbaikan")',
            f"=_xlfn.RANK.EQ(P{r},$P$5:$P${4 + N},0)",
            potensi,
            f'=IF(P{r}>=4,"Tinggi",IF(P{r}>=3,"Sedang","Rendah"))',
            f'=IF(S{r}>=4,"Tinggi",IF(S{r}>=3,"Sedang","Rendah"))',
            f"=INDEX('9-Box'!$C$6:$E$8,MATCH(T{r},'9-Box'!$B$6:$B$8,0),MATCH(U{r},'9-Box'!$C$5:$E$5,0))",
            f"Manager {k['Departemen']}" if k["Level"] != "Manager" else "Direktur",
        ], ["@", None, None, None, None] + ["0"] * 5 + [FMT_PCT] * 5 + ["0.00", None, "0", "0", None, None, None, None])
        for c in range(6, 11):
            wp.cell(row=r, column=c).fill = PatternFill("solid", fgColor=WARNA_INPUT)
        wp.cell(row=r, column=19).fill = PatternFill("solid", fgColor=WARNA_INPUT)
    akhir = 4 + N
    tabel(wp, f"A4:W{akhir}", "tblPenilaian")
    wp.freeze_panes = "C5"
    lebar(wp, [10, 22, 12, 22, 12, 7, 7, 7, 7, 7, 8, 8, 8, 8, 8, 10, 15, 9, 9, 12, 12, 22, 18])
    validasi_angka(wp, f"F5:J{akhir}", 1, 5, "Skor harus bilangan bulat 1 sampai 5.")
    validasi_angka(wp, f"S5:S{akhir}", 1, 5, "Potensi harus bilangan bulat 1 sampai 5.")
    cf_skala_warna(wp, f"P5:P{akhir}")
    cf_skala_warna(wp, f"F5:J{akhir}")
    cf_sama_dengan(wp, f"Q5:Q{akhir}", "Istimewa", "C6EFCE", "006100")
    cf_sama_dengan(wp, f"Q5:Q{akhir}", "Perlu Perbaikan", "FFC7CE", "9C0006")

    # ---------------- 9-Box
    w9 = wb.create_sheet("9-Box")
    judul(w9, "Matriks 9-Box: Performa vs Potensi", "Baris 5-8 = label kuadran (dipakai INDEX/MATCH); baris 12-15 = jumlah karyawan tiap kuadran (COUNTIFS).")
    w9["B5"] = "Performa \\ Potensi"
    w9["B5"].font = font(bold=True)
    for j, pot in enumerate(["Rendah", "Sedang", "Tinggi"], start=3):
        w9.cell(row=5, column=j, value=pot).font = font(bold=True)
    for i, perf in enumerate(["Rendah", "Sedang", "Tinggi"], start=6):
        w9.cell(row=i, column=2, value=perf).font = font(bold=True)
        for j, pot in enumerate(["Rendah", "Sedang", "Tinggi"], start=3):
            c = w9.cell(row=i, column=j, value=BOX[(perf, pot)])
            c.border = BORDER
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    w9["B11"] = "Jumlah karyawan per kuadran"
    w9["B11"].font = font(bold=True, color=WARNA_HEADER)
    w9["B12"] = "Performa \\ Potensi"
    w9["B12"].font = font(bold=True)
    for j, pot in enumerate(["Rendah", "Sedang", "Tinggi"], start=3):
        w9.cell(row=12, column=j, value=pot).font = font(bold=True)
    for i, perf in enumerate(["Rendah", "Sedang", "Tinggi"], start=13):
        w9.cell(row=i, column=2, value=perf).font = font(bold=True)
        for j in range(3, 6):
            kol = "CDE"[j - 3]
            c = w9.cell(row=i, column=j,
                        value=f"=COUNTIFS(Penilaian!$T$5:$T${akhir},$B{i},Penilaian!$U$5:$U${akhir},{kol}$12)")
            c.border = BORDER
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.font = font(bold=True, size=14)
            c.number_format = "0"
    w9["F13"], w9["F14"], w9["F15"] = "=SUM(C13:E13)", "=SUM(C14:E14)", "=SUM(C15:E15)"
    w9["C16"], w9["D16"], w9["E16"], w9["F16"] = "=SUM(C13:C15)", "=SUM(D13:D15)", "=SUM(E13:E15)", "=SUM(F13:F15)"
    for c in ("F13", "F14", "F15", "C16", "D16", "E16", "F16"):
        w9[c].font = font(bold=True)
    cf_skala_warna(w9, "C13:E15")
    w9["B6"].fill = w9["B13"].fill = PatternFill("solid", fgColor="FFC7CE")
    w9["B8"].fill = w9["B15"].fill = PatternFill("solid", fgColor="C6EFCE")
    for r in (6, 7, 8, 13, 14, 15):
        w9.row_dimensions[r].height = 34
    lebar(w9, [3, 20, 22, 22, 22, 8])

    # ---------------- Rekap
    wr = wb.create_sheet("Rekap")
    judul(wr, "Rekap Distribusi Rating per Departemen", "COUNTIFS, AVERAGEIF, dan peringkat departemen.")
    header(wr, 4, 1, ["Departemen", "Jumlah Dinilai", "Istimewa", "Baik", "Cukup", "Perlu Perbaikan", "Rata-rata Nilai",
                      "% Baik ke Atas", "Peringkat Departemen", "Karyawan Terbaik"])
    for i, d in enumerate(DEPARTEMEN, start=5):
        tulis_baris(wr, i, 1, [
            d, f"=COUNTIF(Penilaian!$C$5:$C${akhir},A{i})",
            f"=COUNTIFS(Penilaian!$C$5:$C${akhir},$A{i},Penilaian!$Q$5:$Q${akhir},C$4)",
            f"=COUNTIFS(Penilaian!$C$5:$C${akhir},$A{i},Penilaian!$Q$5:$Q${akhir},D$4)",
            f"=COUNTIFS(Penilaian!$C$5:$C${akhir},$A{i},Penilaian!$Q$5:$Q${akhir},E$4)",
            f"=COUNTIFS(Penilaian!$C$5:$C${akhir},$A{i},Penilaian!$Q$5:$Q${akhir},F$4)",
            f"=IFERROR(AVERAGEIF(Penilaian!$C$5:$C${akhir},A{i},Penilaian!$P$5:$P${akhir}),0)",
            f"=IFERROR((C{i}+D{i})/B{i},0)",
            f"=_xlfn.RANK.EQ(G{i},$G$5:$G$10,0)",
            f"=INDEX(Penilaian!$B$5:$B${akhir},MATCH(_xlfn.MAXIFS(Penilaian!$P$5:$P${akhir},Penilaian!$C$5:$C${akhir},A{i}),Penilaian!$P$5:$P${akhir},0))",
        ], [None, "0", "0", "0", "0", "0", "0.00", FMT_PCT, "0", None])
    tulis_baris(wr, 11, 1, ["Total"] + [f"=SUM({c}5:{c}10)" for c in "BCDEF"] +
                [f"=AVERAGE(Penilaian!$P$5:$P${akhir})", "=IFERROR((C11+D11)/B11,0)", "", ""],
                [None, "0", "0", "0", "0", "0", "0.00", FMT_PCT, None, None])
    for c in range(1, 11):
        wr.cell(row=11, column=c).font = font(bold=True)
        wr.cell(row=11, column=c).fill = PatternFill("solid", fgColor=WARNA_SUB)
    header(wr, 14, 1, ["Rating", "Jumlah", "% Karyawan", "Target Distribusi (acuan)", "Selisih"])
    target = [(RATING[0][1], .10), (RATING[1][1], .40), (RATING[2][1], .40), (RATING[3][1], .10)]
    for i, (rt, tg) in enumerate(target, start=15):
        tulis_baris(wr, i, 1, [rt, f"=COUNTIF(Penilaian!$Q$5:$Q${akhir},A{i})", f"=IFERROR(B{i}/$B$11,0)", tg,
                               f"=C{i}-D{i}"], [None, "0", FMT_PCT, FMT_PCT, "+0.0%;-0.0%"])
        wr.cell(row=i, column=4).fill = PatternFill("solid", fgColor=WARNA_INPUT)
    header(wr, 21, 1, ["Peringkat", "Nama", "Departemen", "Nilai Akhir", "Rating"])
    for i in range(10):
        r = 22 + i
        tulis_baris(wr, r, 1, [
            i + 1,
            f"=INDEX(Penilaian!$B$5:$B${akhir},MATCH(A{r},Penilaian!$X$5:$X${akhir},0))",
            f"=INDEX(Penilaian!$C$5:$C${akhir},MATCH(A{r},Penilaian!$X$5:$X${akhir},0))",
            f"=INDEX(Penilaian!$P$5:$P${akhir},MATCH(A{r},Penilaian!$X$5:$X${akhir},0))",
            f"=INDEX(Penilaian!$Q$5:$Q${akhir},MATCH(A{r},Penilaian!$X$5:$X${akhir},0))",
        ], ["0", None, None, "0.00", None])
    # kolom bantu peringkat unik di Penilaian (X)
    wp["X4"] = "Peringkat Unik (bantu)"
    wp["X4"].font = font(italic=True, size=9, color="595959")
    for i in range(N):
        r = 5 + i
        wp.cell(row=r, column=24, value=f"=R{r}+COUNTIF($P$5:P{r},P{r})-1").number_format = "0"
    lebar(wr, [18, 12, 10, 10, 10, 14, 13, 13, 12, 22])
    chart_bar(wr, "L4", "Distribusi Rating per Departemen", ref(wr, 1, 5, 1, 10), ref(wr, 3, 4, 6, 10),
              judul_y="Orang", lebar_cm=16, tinggi_cm=8, tumpuk=True)
    chart_bar(wr, "L21", "Rata-rata Nilai Akhir per Departemen", ref(wr, 1, 5, 1, 10), ref(wr, 7, 4, 7, 10),
              judul_y="Nilai", lebar_cm=16, tinggi_cm=8)

    wpt = wb.create_sheet("Petunjuk")
    teks_panduan(wpt, 1, [
        "# Petunjuk Penggunaan",
        "1. Master KPI: lima KPI dan bobotnya per level jabatan; kolom Cek memastikan bobot = 100%.",
        "2. Penilaian: atasan mengisi skor 1-5 (sel kuning) dan potensi. Bobot ditarik INDEX/MATCH sesuai level.",
        "3. Nilai akhir = SUMPRODUCT(skor; bobot); rating = IFS; peringkat = RANK.EQ; kuadran 9-box = INDEX/MATCH dua arah.",
        "4. Sheet 9-Box berisi label kuadran dan jumlah karyawan tiap kuadran (COUNTIFS).",
        "5. Rekap: distribusi rating per departemen, perbandingan dengan target distribusi, dan 10 besar karyawan.",
        "6. Conditional Formatting skala warna memudahkan melihat skor tinggi (hijau) dan rendah (merah).",
    ])
    wb.active = wb.sheetnames.index("Penilaian")
    return rapikan(wb)


SCREENSHOT = [("Penilaian", "07_penilaian.png", "A1:W32"), ("9-Box", "07_9box.png", "A1:F17"),
              ("Rekap", "07_rekap.png", "A1:T32")]
