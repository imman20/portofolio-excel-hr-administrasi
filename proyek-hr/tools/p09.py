"""Proyek 09 - Analisis Survei Engagement dengan Statistika (tingkat lanjut). 200 responden."""

import random

from openpyxl import Workbook
from openpyxl.styles import PatternFill

from util import (DEPARTEMEN, WARNA_INPUT, WARNA_SUB, cf_skala_warna, chart_bar, font, header, judul, label_nilai,
                  lebar, rapikan, ref, tabel, teks_panduan, tulis_baris)

N = 200
DIMENSI = ["Kepemimpinan", "Kompensasi", "Pengembangan Karier", "Lingkungan Kerja", "Keseimbangan Kerja-Hidup"]
BOBOT_ASLI = [0.30, 0.20, 0.25, 0.15, 0.10]


def build():
    rng = random.Random(909)
    wb = Workbook()
    ws = wb.active
    ws.title = "Data Survei"
    judul(ws, "Data Survei Engagement Karyawan 2024 (skala Likert 1-5)",
          "Skor dimensi = AVERAGE tiga item; engagement = AVERAGE tiga item E. Data diurutkan menurut status kepegawaian untuk uji t.")
    head = ["ID", "Departemen", "JK", "Status", "Masa Kerja"] + [f"{d[:3].upper()}{i}" for d in DIMENSI for i in (1, 2, 3)] + \
           ["ENG1", "ENG2", "ENG3"] + DIMENSI + ["Engagement"] + [f"Eng {d}" for d in DEPARTEMEN] + ["Prediksi Regresi", "Residual"]
    header(ws, 4, 1, head)
    efek_dept = {"HR": .2, "Finance": 0, "Operasional": -.25, "Marketing": .1, "IT": .15, "Produksi": -.35}
    baris = []
    for i in range(N):
        dept = rng.choices(DEPARTEMEN, weights=[10, 12, 22, 14, 14, 28])[0]
        status = rng.choices(["Tetap", "Kontrak"], weights=[70, 30])[0]
        lat = [max(1, min(5, rng.gauss(3.6 + efek_dept[dept] + (.15 if status == "Tetap" else -.2), .7))) for _ in DIMENSI]
        item = []
        for l in lat:
            item += [max(1, min(5, round(l + rng.gauss(0, .55)))) for _ in range(3)]
        eng = sum(b * l for b, l in zip(BOBOT_ASLI, lat)) + rng.gauss(0, .35)
        item += [max(1, min(5, round(eng + rng.gauss(0, .5)))) for _ in range(3)]
        baris.append((f"R{i + 1:03d}", dept, rng.choice("LP"), status, rng.choice(["< 2 thn", "2-5 thn", "> 5 thn"]), item))
    baris.sort(key=lambda b: (b[3] != "Tetap", b[0]))
    n_tetap = sum(1 for b in baris if b[3] == "Tetap")
    for i, (idr, dept, jk, status, mk, item) in enumerate(baris):
        r = 5 + i
        rumus = [f"=AVERAGE({chr(70 + 3 * j)}{r}:{chr(72 + 3 * j)}{r})" for j in range(5)] + [f"=AVERAGE(U{r}:W{r})"]
        helper = [f'=IF($B{r}="{d}",$AC{r},"")' for d in DEPARTEMEN]
        pred = (f"=Regresi!$G$5+Regresi!$F$5*X{r}+Regresi!$E$5*Y{r}+Regresi!$D$5*Z{r}+Regresi!$C$5*AA{r}+Regresi!$B$5*AB{r}")
        tulis_baris(ws, r, 1, [idr, dept, jk, status, mk] + item + rumus + helper + [pred, f"=AC{r}-AJ{r}"],
                    [None] * 5 + ["0"] * 18 + ["0.00"] * 6 + ["0.00"] * 6 + ["0.00", "0.00"])
    akhir = 4 + N
    tabel(ws, f"A4:AK{akhir}", "tblSurvei")
    ws.freeze_panes = "F5"
    lebar(ws, [7, 12, 5, 9, 10] + [5] * 18 + [13, 12, 14, 13, 14, 12] + [9] * 6 + [10, 9])
    r_tetap = (5, 4 + n_tetap)
    r_kontrak = (5 + n_tetap, akhir)

    # ---------------- Statistik Deskriptif
    wsd = wb.create_sheet("Statistik Deskriptif")
    judul(wsd, "Statistik Deskriptif per Dimensi", "Semua ukuran dihitung dengan fungsi statistik Excel atas kolom skor di Data Survei.")
    ukuran = [("N", "COUNT"), ("Rata-rata", "AVERAGE"), ("Median", "MEDIAN"), ("Modus", "_xlfn.MODE.SNGL"),
              ("Simpangan baku (s)", "_xlfn.STDEV.S"), ("Varians", "_xlfn.VAR.S"), ("Minimum", "MIN"), ("Maksimum", "MAX"),
              ("Kuartil 1", None), ("Kuartil 3", None), ("Skewness", "SKEW"), ("Kurtosis", "KURT"),
              ("Standard error", None), ("Batas bawah CI 95%", None), ("Batas atas CI 95%", None)]
    kol = ["X", "Y", "Z", "AA", "AB", "AC"]
    header(wsd, 4, 1, ["Ukuran"] + DIMENSI + ["Engagement"])
    for i, (lab, fn) in enumerate(ukuran, start=5):
        row = [lab]
        for j, k in enumerate(kol):
            rng_ = f"'Data Survei'!${k}$5:${k}${akhir}"
            c = chr(66 + j)
            if fn:
                row.append(f"={fn}({rng_})")
            elif lab == "Kuartil 1":
                row.append(f"=QUARTILE({rng_},1)")
            elif lab == "Kuartil 3":
                row.append(f"=QUARTILE({rng_},3)")
            elif lab == "Standard error":
                row.append(f"={c}9/SQRT({c}5)")
            elif lab.startswith("Batas bawah"):
                row.append(f"={c}6-_xlfn.T.INV.2T(0.05,{c}5-1)*{c}17")
            else:
                row.append(f"={c}6+_xlfn.T.INV.2T(0.05,{c}5-1)*{c}17")
        tulis_baris(wsd, i, 1, row, [None] + ["0.00"] * 6)
    wsd["B5"].number_format = "0"
    for c in "CDEFG":
        wsd[f"{c}5"].number_format = "0"
    label_nilai(wsd, 22, 1, "Dimensi dengan rata-rata terendah", "=INDEX($B$4:$F$4,MATCH(MIN(B6:F6),B6:F6,0))")
    label_nilai(wsd, 23, 1, "Dimensi dengan rata-rata tertinggi", "=INDEX($B$4:$F$4,MATCH(MAX(B6:F6),B6:F6,0))")
    label_nilai(wsd, 24, 1, "Dimensi paling bervariasi (s terbesar)", "=INDEX($B$4:$F$4,MATCH(MAX(B9:F9),B9:F9,0))")
    lebar(wsd, [30, 14, 14, 18, 16, 20, 13])
    chart_bar(wsd, "I4", "Rata-rata Skor per Dimensi", ref(wsd, 2, 4, 7, 4), ref(wsd, 1, 6, 7, 6), judul_y="Skor",
              lebar_cm=15, tinggi_cm=8)
    # catatan: series dari baris -> gunakan tabel transposisi sederhana
    header(wsd, 27, 1, ["Dimensi", "Rata-rata", "Simpangan Baku"])
    for j, d in enumerate(DIMENSI + ["Engagement"], start=28):
        c = chr(66 + j - 28)
        tulis_baris(wsd, j, 1, [d, f"={c}6", f"={c}9"], [None, "0.00", "0.00"])
    wsd._charts.clear()
    chart_bar(wsd, "I4", "Rata-rata Skor per Dimensi", ref(wsd, 1, 28, 1, 33), ref(wsd, 2, 27, 2, 33), judul_y="Skor",
              lebar_cm=15, tinggi_cm=8)

    # ---------------- Korelasi
    wk = wb.create_sheet("Korelasi")
    judul(wk, "Matriks Korelasi Pearson (CORREL)", "Warna hijau = korelasi kuat positif. Nilai > 0,5 dianggap kuat untuk konteks survei.")
    nama = DIMENSI + ["Engagement"]
    header(wk, 4, 1, [""] + nama)
    for i, k1 in enumerate(kol):
        r = 5 + i
        row = [nama[i]]
        for k2 in kol:
            row.append(f"=CORREL('Data Survei'!${k1}$5:${k1}${akhir},'Data Survei'!${k2}$5:${k2}${akhir})")
        tulis_baris(wk, r, 1, row, [None] + ["0.00"] * 6)
    cf_skala_warna(wk, "B5:G10")
    label_nilai(wk, 13, 1, "Dimensi berkorelasi terkuat dengan engagement", "=INDEX($A$5:$A$9,MATCH(MAX(G5:G9),G5:G9,0))")
    label_nilai(wk, 14, 1, "Nilai korelasinya", "=MAX(G5:G9)", "0.00")
    label_nilai(wk, 15, 1, "Uji signifikansi r (t = r*SQRT(n-2)/SQRT(1-r^2))", f"=B14*SQRT({N}-2)/SQRT(1-B14^2)", "0.00")
    label_nilai(wk, 16, 1, "p-value (dua sisi)", f"=_xlfn.T.DIST.2T(ABS(B15),{N}-2)", "0.0000")
    lebar(wk, [40, 14, 14, 18, 16, 22, 13])

    # ---------------- Uji t
    wt = wb.create_sheet("Uji t")
    judul(wt, "Uji t Dua Sampel: Engagement Karyawan Tetap vs Kontrak",
          "H0: rata-rata engagement kedua kelompok sama. Data diurutkan per status sehingga tiap kelompok adalah satu rentang.")
    rt = f"'Data Survei'!$AC${r_tetap[0]}:$AC${r_tetap[1]}"
    rk = f"'Data Survei'!$AC${r_kontrak[0]}:$AC${r_kontrak[1]}"
    header(wt, 4, 1, ["Statistik", "Tetap", "Kontrak"])
    stat = [("n", f"=COUNT({rt})", f"=COUNT({rk})", "0"), ("Rata-rata", f"=AVERAGE({rt})", f"=AVERAGE({rk})", "0.000"),
            ("Varians", f"=_xlfn.VAR.S({rt})", f"=_xlfn.VAR.S({rk})", "0.000"),
            ("Simpangan baku", f"=_xlfn.STDEV.S({rt})", f"=_xlfn.STDEV.S({rk})", "0.000")]
    for i, (lab, a, b, fmt) in enumerate(stat, start=5):
        tulis_baris(wt, i, 1, [lab, a, b], [None, fmt, fmt])
    label_nilai(wt, 10, 1, "Alpha", 0.05, "0.00", fill=WARNA_INPUT)
    label_nilai(wt, 11, 1, "Selisih rata-rata", "=B6-C6", "0.000")
    label_nilai(wt, 12, 1, "Varians gabungan (pooled)", "=((B5-1)*B7+(C5-1)*C7)/(B5+C5-2)", "0.000")
    label_nilai(wt, 13, 1, "t hitung (manual)", "=B11/SQRT(B12*(1/B5+1/C5))", "0.000")
    label_nilai(wt, 14, 1, "Derajat bebas", "=B5+C5-2", "0")
    label_nilai(wt, 15, 1, "t tabel (dua sisi)", "=_xlfn.T.INV.2T(B10,B14)", "0.000")
    label_nilai(wt, 16, 1, "p-value manual (T.DIST.2T)", "=_xlfn.T.DIST.2T(ABS(B13),B14)", "0.0000")
    label_nilai(wt, 17, 1, "p-value T.TEST (varians sama, tipe 2)", f"=_xlfn.T.TEST({rt},{rk},2,2)", "0.0000")
    label_nilai(wt, 18, 1, "p-value T.TEST Welch (varians beda, tipe 3)", f"=_xlfn.T.TEST({rt},{rk},2,3)", "0.0000")
    label_nilai(wt, 19, 1, "Keputusan", '=IF(B17<B10,"Tolak H0: engagement berbeda signifikan","Gagal tolak H0: tidak ada perbedaan signifikan")')
    label_nilai(wt, 20, 1, "Cohen's d (ukuran efek)", "=ABS(B11)/SQRT(B12)", "0.00")
    lebar(wt, [44, 16, 16])

    # ---------------- ANOVA
    wa = wb.create_sheet("ANOVA")
    judul(wa, "ANOVA Satu Arah: Engagement antar Departemen",
          "Dihitung manual dengan rumus (COUNT, AVERAGE, VAR.S, SUMPRODUCT, F.DIST.RT) lalu dibandingkan dengan Analysis ToolPak.")
    header(wa, 4, 1, ["Departemen", "n", "Rata-rata", "Varians", "Jumlah"])
    for i, d in enumerate(DEPARTEMEN, start=5):
        k = chr(ord("AD"[1]) + i - 5)  # AD..AI
        kk = "A" + k
        rg = f"'Data Survei'!${kk}$5:${kk}${akhir}"
        tulis_baris(wa, i, 1, [d, f"=COUNT({rg})", f"=AVERAGE({rg})", f"=_xlfn.VAR.S({rg})", f"=SUM({rg})"],
                    [None, "0", "0.000", "0.000", "0.00"])
    label_nilai(wa, 12, 1, "Alpha", 0.05, "0.00", fill=WARNA_INPUT)
    label_nilai(wa, 13, 1, "N total", "=SUM(B5:B10)", "0")
    label_nilai(wa, 14, 1, "k (jumlah kelompok)", "=COUNT(B5:B10)", "0")
    label_nilai(wa, 15, 1, "Rata-rata keseluruhan", "=SUM(E5:E10)/B13", "0.000")
    header(wa, 17, 1, ["Sumber Variasi", "SS", "df", "MS", "F", "p-value", "F tabel"])
    tulis_baris(wa, 18, 1, ["Antar kelompok", "=SUMPRODUCT(B5:B10,(C5:C10-B15)^2)", "=B14-1", "=B18/C18", "=D18/D19",
                            "=_xlfn.F.DIST.RT(E18,C18,C19)", "=_xlfn.F.INV.RT(B12,C18,C19)"],
                [None, "0.000", "0", "0.000", "0.000", "0.0000", "0.000"])
    tulis_baris(wa, 19, 1, ["Dalam kelompok", "=SUMPRODUCT(B5:B10-1,D5:D10)", "=B13-B14", "=B19/C19", "", "", ""],
                [None, "0.000", "0", "0.000", None, None, None])
    tulis_baris(wa, 20, 1, ["Total", "=B18+B19", "=C18+C19", "", "", "", ""], [None, "0.000", "0", None, None, None, None])
    for c in range(1, 8):
        wa.cell(row=20, column=c).font = font(bold=True)
        wa.cell(row=20, column=c).fill = PatternFill("solid", fgColor=WARNA_SUB)
    label_nilai(wa, 22, 1, "Keputusan", '=IF(F18<B12,"Tolak H0: minimal satu departemen berbeda","Gagal tolak H0")')
    label_nilai(wa, 23, 1, "Eta kuadrat (ukuran efek)", "=B18/B20", "0.000")
    label_nilai(wa, 24, 1, "Departemen engagement terendah", "=INDEX(A5:A10,MATCH(MIN(C5:C10),C5:C10,0))")
    lebar(wa, [30, 12, 12, 12, 12, 12, 12])
    chart_bar(wa, "I4", "Rata-rata Engagement per Departemen", ref(wa, 1, 5, 1, 10), ref(wa, 3, 4, 3, 10), judul_y="Skor",
              lebar_cm=14, tinggi_cm=8)
    teks_panduan(wa, 27, [
        "## Langkah Analysis ToolPak (untuk pembanding)",
        "1. File > Options > Add-ins > Manage: Excel Add-ins > Go > centang Analysis ToolPak > OK.",
        "2. Data > Data Analysis > Anova: Single Factor > OK.",
        "3. Input Range: blok kolom bantu 'Eng HR' sampai 'Eng Produksi' di Data Survei (AD4:AI204), Grouped By Columns, centang Labels in first row.",
        "4. Alpha 0,05; Output Range: sel kosong (misal I20). OK. Nilai SS, df, MS, F, p-value, F crit harus sama dengan tabel di atas.",
    ], lebar_kolom=30)

    # ---------------- Regresi
    wr = wb.create_sheet("Regresi")
    judul(wr, "Regresi Linear Berganda: Engagement ~ 5 Dimensi (LINEST)",
          "Matriks LINEST 5x6 diambil per sel dengan INDEX; koefisien keluar dalam urutan terbalik (X5 ... X1, konstanta).")
    y = f"'Data Survei'!$AC$5:$AC${akhir}"
    x = f"'Data Survei'!$X$5:$AB${akhir}"
    header(wr, 4, 1, ["Matriks LINEST", "X5 Keseimbangan", "X4 Lingkungan", "X3 Karier", "X2 Kompensasi", "X1 Kepemimpinan", "Konstanta"])
    lab = ["Koefisien", "Standard error", "R2 | SE y", "F | df residual", "SS regresi | SS residual"]
    wr["A5"] = lab[0]
    for rr in range(5):
        wr.cell(row=5 + rr, column=1, value=lab[rr]).font = font(bold=True)
        for cc in range(6):
            c = wr.cell(row=5 + rr, column=2 + cc, value=f'=IFERROR(INDEX(LINEST({y},{x},TRUE,TRUE),{rr + 1},{cc + 1}),"")')
            c.number_format = "0.0000"
    wr["A5"].font = font(bold=True)
    # tabel koefisien rapi (baris 7-12 kolom A-C dipakai oleh prediksi di Data Survei) -> letakkan di kolom I
    header(wr, 6, 9, ["Variabel", "Koefisien", "SE", "t hitung", "p-value", "Signifikan (5%)"])
    urut = [("Konstanta", 7), ("Kepemimpinan", 6), ("Kompensasi", 5), ("Pengembangan Karier", 4), ("Lingkungan Kerja", 3),
            ("Keseimbangan Kerja-Hidup", 2)]
    for i, (nm, kc) in enumerate(urut, start=7):
        kl = chr(64 + kc)
        tulis_baris(wr, i, 9, [nm, f"={kl}5", f"={kl}6", f"=IFERROR(J{i}/K{i},0)", f"=_xlfn.T.DIST.2T(ABS(L{i}),$C$8)",
                               f'=IF(M{i}<0.05,"Ya","Tidak")'], [None, "0.0000", "0.0000", "0.000", "0.0000", None])
    # salinan koefisien di kolom C baris 7-12 untuk prediksi (dirujuk Data Survei)
    wr["A6"].value = "Standard error"
    header(wr, 14, 1, ["Ringkasan Model", "Nilai"])
    ringkas = [("R kuadrat", "=B7", "0.0000"), ("R kuadrat disesuaikan", f"=1-(1-B7)*({N}-1)/({N}-5-1)", "0.0000"),
               ("Standard error estimasi", "=C7", "0.0000"), ("F hitung", "=B8", "0.000"),
               ("df regresi", 5, "0"), ("df residual", "=C8", "0"),
               ("p-value model (F.DIST.RT)", "=_xlfn.F.DIST.RT(B18,B19,B20)", "0.0000"),
               ("Jumlah observasi", f"=COUNT({y})", "0")]
    for i, (lab_, v, fmt) in enumerate(ringkas, start=15):
        tulis_baris(wr, i, 1, [lab_, v], [None, fmt])
    # koefisien untuk prediksi di C7:C12? bentrok dengan matriks (baris 5-9). Gunakan baris 7..12 kolom C bentrok -> pindahkan prediksi ke kolom Q
    lebar(wr, [26, 14, 14, 14, 14, 16, 12, 2, 26, 12, 12, 12, 12, 14])
    teks_panduan(wr, 25, [
        "## Cara membaca",
        "Koefisien positif berarti kenaikan satu poin pada dimensi tersebut (dimensi lain tetap) menaikkan engagement sebesar koefisien itu.",
        "p-value < 0,05 berarti dimensi berpengaruh signifikan. R kuadrat menunjukkan proporsi variasi engagement yang dijelaskan kelima dimensi.",
        "Cara ToolPak: Data > Data Analysis > Regression; Input Y Range = AC4:AC204, Input X Range = X4:AB204, centang Labels, OK.",
        "Prediksi dan residual tiap responden ada di kolom AJ-AK Data Survei (konstanta + jumlah koefisien x skor).",
    ], lebar_kolom=26)

    # ---------------- Histogram
    wh = wb.create_sheet("Histogram")
    judul(wh, "Histogram Skor Engagement", "Frekuensi tiap kelas dihitung COUNTIFS (alternatif FREQUENCY). Chart kolom dengan gap kecil.")
    header(wh, 4, 1, ["Kelas", "Batas Bawah", "Batas Atas", "Frekuensi", "Frekuensi Relatif", "Kumulatif"])
    batas = [1, 1.5, 2, 2.5, 3, 3.5, 4, 4.5, 5.01]
    for i in range(8):
        r = 5 + i
        tulis_baris(wh, r, 1, [f"{batas[i]:.1f} - {min(batas[i + 1], 5):.1f}", batas[i], batas[i + 1],
                               f'=COUNTIFS({y},">="&B{r},{y},"<"&C{r})', f"=D{r}/SUM($D$5:$D$12)", f"=SUM($D$5:D{r})"],
                    [None, "0.0", "0.0", "0", "0.0%", "0"])
    label_nilai(wh, 14, 1, "Kelas modus", "=INDEX(A5:A12,MATCH(MAX(D5:D12),D5:D12,0))")
    label_nilai(wh, 15, 1, "Skewness engagement", f"=SKEW({y})", "0.000")
    label_nilai(wh, 16, 1, "Interpretasi bentuk", '=IF(ABS(B15)<0.5,"Hampir simetris",IF(B15<0,"Menceng kiri (banyak skor tinggi)","Menceng kanan (banyak skor rendah)"))')
    lebar(wh, [24, 12, 12, 12, 16, 12])
    ch = chart_bar(wh, "H4", "Histogram Engagement", ref(wh, 1, 5, 1, 12), ref(wh, 4, 4, 4, 12), judul_y="Responden",
                   judul_x="Kelas skor", lebar_cm=15, tinggi_cm=8)
    ch.gapWidth = 10

    # ---------------- Interpretasi
    wi = wb.create_sheet("Interpretasi")
    teks_panduan(wi, 1, [
        "# Interpretasi untuk Tim HR",
        "1. Statistik deskriptif memperlihatkan dimensi dengan rata-rata terendah (lihat sel dinamis di sheet Statistik Deskriptif); dimensi ini menjadi prioritas program perbaikan.",
        "2. Korelasi: dimensi dengan r tertinggi terhadap engagement adalah pengungkit utama; intervensi di sana memberi dampak terbesar.",
        "3. Uji t: bila p-value < 0,05, engagement karyawan tetap dan kontrak berbeda nyata; HR perlu meninjau kebijakan untuk kelompok yang lebih rendah.",
        "4. ANOVA: bila signifikan, lakukan uji lanjut (misalnya Tukey atau uji t berpasangan dengan koreksi Bonferroni) untuk menemukan departemen mana yang berbeda.",
        "5. Regresi: koefisien signifikan menunjukkan dimensi yang benar-benar mendorong engagement setelah dimensi lain dikendalikan; gunakan R kuadrat untuk menilai kecukupan model.",
        "6. Histogram: distribusi menceng kiri berarti mayoritas karyawan cukup engaged, tetapi perhatikan ekor kiri (skor < 2,5) sebagai kelompok berisiko.",
        "7. Keterbatasan: data sintetis, sampel 200, asumsi normalitas dan homogenitas varians sebaiknya diuji (Levene, Shapiro-Wilk) pada data nyata.",
    ])
    wb.active = 0
    return rapikan(wb)


SCREENSHOT = [("Regresi", "09_regresi.png", "A1:N32"), ("ANOVA", "09_anova.png", "A1:P32"),
              ("Statistik Deskriptif", "09_statistik_deskriptif.png", "A1:P25")]
