"""Proyek 06 - Rekrutmen Tracker (tingkat menengah). Periode: Januari - September 2024."""

import datetime as dt
import random

from openpyxl import Workbook
from openpyxl.styles import PatternFill

from data_gen import NAMA_BELAKANG, NAMA_DEPAN_L, NAMA_DEPAN_P
from util import (DEPARTEMEN, FMT_PCT, FMT_TGL, WARNA_INPUT, WARNA_SUB, cf_sama_dengan, chart_bar, dropdown, font,
                  header, judul, label_nilai, lebar, rapikan, ref, tabel, teks_panduan, tulis_baris)

TAHAP = ["Applied", "Screening", "Interview HR", "Interview User", "Offer", "Hired"]
SUMBER = ["JobStreet", "LinkedIn", "Referral Karyawan", "Website Perusahaan", "Job Fair Kampus"]
POSISI = [("Staff HR", "HR"), ("Staff Akuntansi", "Finance"), ("Staff Operasional", "Operasional"),
          ("Staff Marketing", "Marketing"), ("Programmer", "IT"), ("Operator Produksi", "Produksi"),
          ("Supervisor Produksi", "Produksi"), ("Digital Marketing Specialist", "Marketing")]
REKRUTER = ["Dewi Saputra", "Wisnu Pratama", "Farah Sudrajat"]
N = 150


def build():
    rng = random.Random(606)
    wb = Workbook()
    ws = wb.active
    ws.title = "Kandidat"
    judul(ws, "Pipeline Rekrutmen 2024", "Satu baris per kandidat. Tanggal tiap tahap diisi saat kandidat lolos; kolom status dan time-to-hire otomatis.")
    head = ["ID Kandidat", "Nama Kandidat", "Posisi", "Departemen", "Sumber", "Tanggal Lamar", "Tgl Screening",
            "Tgl Interview HR", "Tgl Interview User", "Tgl Offer", "Tgl Hired", "Tgl Ditolak", "Tahap Terakhir",
            "Status", "Time to Hire (hari)", "Lama di Pipeline (hari)", "Rekruter"]
    header(ws, 4, 1, head)
    tgl = dt.date(2024, 1, 8)
    for i in range(N):
        r = 5 + i
        tgl += dt.timedelta(days=rng.choice([0, 1, 1, 2, 3]))
        jk = rng.choice("LP")
        nama = f"{rng.choice(NAMA_DEPAN_L if jk == 'L' else NAMA_DEPAN_P)} {rng.choice(NAMA_BELAKANG)}"
        pos, dept = rng.choice(POSISI)
        sumber = rng.choices(SUMBER, weights=[35, 25, 15, 15, 10])[0]
        # simulasi perjalanan di funnel
        tanggal = [tgl]
        lolos = True
        prob = {"JobStreet": [.55, .6, .65, .85, .9], "LinkedIn": [.65, .65, .7, .85, .9],
                "Referral Karyawan": [.8, .75, .75, .9, .95], "Website Perusahaan": [.5, .6, .65, .85, .9],
                "Job Fair Kampus": [.6, .55, .6, .8, .9]}[sumber]
        for s in range(5):
            if rng.random() < prob[s]:
                tanggal.append(tanggal[-1] + dt.timedelta(days=rng.randint(3, 12)))
            else:
                lolos = False
                break
        ditolak = None
        status = "Hired" if lolos else "Rejected"
        if not lolos:
            ditolak = tanggal[-1] + dt.timedelta(days=rng.randint(2, 7))
            if tgl > dt.date(2024, 8, 15) and rng.random() < .6:
                status, ditolak = "Aktif", None
        if status == "Hired" and tanggal[-1] > dt.date(2024, 9, 30):
            status = "Aktif"
            tanggal = tanggal[:-1]
        kolom_tgl = tanggal + [None] * (6 - len(tanggal))
        tulis_baris(ws, r, 1, [
            f"KD-2024-{i + 1:03d}", nama, pos, dept, sumber] + kolom_tgl + [ditolak,
            f'=IF(K{r}<>"","Hired",IF(J{r}<>"","Offer",IF(I{r}<>"","Interview User",IF(H{r}<>"","Interview HR",IF(G{r}<>"","Screening","Applied")))))',
            status,
            f'=IF(N{r}="Hired",DATEDIF(F{r},K{r},"D"),"")',
            f'=IF(N{r}="Hired",K{r}-F{r},IF(N{r}="Rejected",L{r}-F{r},TODAY()-F{r}))',
            rng.choice(REKRUTER),
        ], [None, None, None, None, None] + [FMT_TGL] * 7 + [None, None, "0", "0", None])
    akhir = 4 + N
    tabel(ws, f"A4:Q{akhir}", "tblKandidat")
    ws.freeze_panes = "C5"
    lebar(ws, [12, 22, 24, 12, 18, 12, 12, 12, 12, 12, 12, 12, 14, 10, 11, 11, 16])
    dropdown(ws, f"C5:C{akhir + 200}", "=Daftar!$A$2:$A$9")
    dropdown(ws, f"E5:E{akhir + 200}", "=Daftar!$C$2:$C$6")
    dropdown(ws, f"N5:N{akhir + 200}", ["Aktif", "Hired", "Rejected"])
    cf_sama_dengan(ws, f"N5:N{akhir}", "Hired", "C6EFCE", "006100")
    cf_sama_dengan(ws, f"N5:N{akhir}", "Rejected", "FFC7CE", "9C0006")
    cf_sama_dengan(ws, f"N5:N{akhir}", "Aktif", "FFEB9C", "9C5700")

    # ---------------- Funnel
    wf = wb.create_sheet("Funnel")
    judul(wf, "Funnel Rekrutmen", "Jumlah kandidat yang mencapai tiap tahap (COUNTIFS pada kolom tanggal) dan conversion rate antar tahap.")
    header(wf, 4, 1, ["Tahap", "Jumlah Kandidat", "% dari Applied", "Conversion dari Tahap Sebelumnya", "Drop-off"])
    kol_tgl = ["F", "G", "H", "I", "J", "K"]
    for i, (t, kc) in enumerate(zip(TAHAP, kol_tgl), start=5):
        tulis_baris(wf, i, 1, [
            t, f'=COUNTIFS(Kandidat!${kc}$5:${kc}${akhir},"<>")',
            f"=IFERROR(B{i}/$B$5,0)",
            "=1" if i == 5 else f"=IFERROR(B{i}/B{i - 1},0)",
            "=0" if i == 5 else f"=B{i - 1}-B{i}",
        ], [None, "0", FMT_PCT, FMT_PCT, "0"])
    label_nilai(wf, 12, 1, "Kandidat masih aktif di pipeline", f'=COUNTIF(Kandidat!$N$5:$N${akhir},"Aktif")', "0")
    label_nilai(wf, 13, 1, "Kandidat ditolak", f'=COUNTIF(Kandidat!$N$5:$N${akhir},"Rejected")', "0")
    label_nilai(wf, 14, 1, "Total kandidat diterima (Hired)", f'=COUNTIF(Kandidat!$N$5:$N${akhir},"Hired")', "0")
    label_nilai(wf, 15, 1, "Rata-rata time-to-hire (hari)", f"=IFERROR(AVERAGE(Kandidat!$O$5:$O${akhir}),0)", "0.0")
    label_nilai(wf, 16, 1, "Time-to-hire tercepat (hari)", f"=IFERROR(MIN(Kandidat!$O$5:$O${akhir}),0)", "0")
    label_nilai(wf, 17, 1, "Time-to-hire terlama (hari)", f"=IFERROR(MAX(Kandidat!$O$5:$O${akhir}),0)", "0")
    label_nilai(wf, 18, 1, "Offer acceptance rate", "=IFERROR(B10/B9,0)", FMT_PCT)
    lebar(wf, [32, 16, 14, 18, 10])
    chart_bar(wf, "G4", "Funnel Rekrutmen 2024", ref(wf, 1, 5, 1, 10), ref(wf, 2, 4, 2, 10), judul_y="Kandidat",
              lebar_cm=16, tinggi_cm=9, horizontal=True)
    wf._charts[-1].x_axis.scaling.orientation = "maxMin"

    # ---------------- Ringkasan (pengganti PivotTable, dengan rumus)
    wr = wb.create_sheet("Ringkasan Rumus")
    judul(wr, "Ringkasan ala PivotTable (dibuat dengan rumus)", "Bandingkan dengan PivotTable yang Anda buat sendiri mengikuti sheet Panduan Pivot.")
    header(wr, 4, 1, ["Sumber Kandidat", "Jumlah Pelamar", "Hired", "Hire Rate", "Rejected", "Aktif", "Rata-rata Time to Hire"])
    for i, s in enumerate(SUMBER, start=5):
        tulis_baris(wr, i, 1, [
            s, f"=COUNTIF(Kandidat!$E$5:$E${akhir},A{i})",
            f'=COUNTIFS(Kandidat!$E$5:$E${akhir},A{i},Kandidat!$N$5:$N${akhir},"Hired")',
            f"=IFERROR(C{i}/B{i},0)",
            f'=COUNTIFS(Kandidat!$E$5:$E${akhir},A{i},Kandidat!$N$5:$N${akhir},"Rejected")',
            f'=COUNTIFS(Kandidat!$E$5:$E${akhir},A{i},Kandidat!$N$5:$N${akhir},"Aktif")',
            f'=IFERROR(AVERAGEIFS(Kandidat!$O$5:$O${akhir},Kandidat!$E$5:$E${akhir},A{i},Kandidat!$N$5:$N${akhir},"Hired"),0)',
        ], [None, "0", "0", FMT_PCT, "0", "0", "0.0"])
    tulis_baris(wr, 10, 1, ["Total", "=SUM(B5:B9)", "=SUM(C5:C9)", "=IFERROR(C10/B10,0)", "=SUM(E5:E9)", "=SUM(F5:F9)",
                            f"=IFERROR(AVERAGE(Kandidat!$O$5:$O${akhir}),0)"], [None, "0", "0", FMT_PCT, "0", "0", "0.0"])
    for c in range(1, 8):
        wr.cell(row=10, column=c).font = font(bold=True)
        wr.cell(row=10, column=c).fill = PatternFill("solid", fgColor=WARNA_SUB)

    header(wr, 13, 1, ["Posisi"] + TAHAP + ["Hired %"])
    for i, (pos, _) in enumerate(POSISI, start=14):
        baris = [pos]
        for kc in kol_tgl:
            baris.append(f'=COUNTIFS(Kandidat!$C$5:$C${akhir},$A{i},Kandidat!${kc}$5:${kc}${akhir},"<>")')
        baris.append(f"=IFERROR(G{i}/B{i},0)")
        tulis_baris(wr, i, 1, baris, [None] + ["0"] * 6 + [FMT_PCT])
    header(wr, 24, 1, ["Departemen", "Pelamar", "Hired", "Rata-rata Time to Hire", "Rekruter Terbanyak Hired"])
    for i, d in enumerate(DEPARTEMEN, start=25):
        tulis_baris(wr, i, 1, [
            d, f"=COUNTIF(Kandidat!$D$5:$D${akhir},A{i})",
            f'=COUNTIFS(Kandidat!$D$5:$D${akhir},A{i},Kandidat!$N$5:$N${akhir},"Hired")',
            f'=IFERROR(AVERAGEIFS(Kandidat!$O$5:$O${akhir},Kandidat!$D$5:$D${akhir},A{i},Kandidat!$N$5:$N${akhir},"Hired"),0)',
            "",
        ], [None, "0", "0", "0.0", None])
    header(wr, 33, 1, ["Rekruter", "Kandidat Ditangani", "Hired", "Rata-rata Time to Hire"])
    for i, rk in enumerate(REKRUTER, start=34):
        tulis_baris(wr, i, 1, [
            rk, f"=COUNTIF(Kandidat!$Q$5:$Q${akhir},A{i})",
            f'=COUNTIFS(Kandidat!$Q$5:$Q${akhir},A{i},Kandidat!$N$5:$N${akhir},"Hired")',
            f'=IFERROR(AVERAGEIFS(Kandidat!$O$5:$O${akhir},Kandidat!$Q$5:$Q${akhir},A{i},Kandidat!$N$5:$N${akhir},"Hired"),0)',
        ], [None, "0", "0", "0.0"])
    header(wr, 39, 1, ["Bulan Lamar", "Pelamar", "Hired"])
    for m in range(1, 10):
        r = 39 + m
        wr.cell(row=r, column=6, value=dt.date(2024, m, 1)).number_format = FMT_TGL
        wr.cell(row=r, column=7, value=dt.date(2024, m + 1, 1)).number_format = FMT_TGL
        tulis_baris(wr, r, 1, [
            f'=TEXT(F{r},"mmm yyyy")',
            f'=COUNTIFS(Kandidat!$F$5:$F${akhir},">="&F{r},Kandidat!$F$5:$F${akhir},"<"&G{r})',
            f'=COUNTIFS(Kandidat!$F$5:$F${akhir},">="&F{r},Kandidat!$F$5:$F${akhir},"<"&G{r},Kandidat!$N$5:$N${akhir},"Hired")',
        ], [None, "0", "0"])
    wr["F39"], wr["G39"] = "Awal bulan (bantu)", "Bulan berikut (bantu)"
    wr["F39"].font = wr["G39"].font = font(italic=True, size=9, color="595959")
    lebar(wr, [28, 13, 12, 13, 13, 12, 14, 10])
    chart_bar(wr, "J4", "Hire Rate per Sumber Kandidat", ref(wr, 1, 5, 1, 9), ref(wr, 4, 4, 4, 9), judul_y="%",
              lebar_cm=14, tinggi_cm=8)
    chart_bar(wr, "J21", "Pelamar vs Hired per Bulan", ref(wr, 1, 40, 1, 48), ref(wr, 2, 39, 3, 48), judul_y="Orang",
              lebar_cm=14, tinggi_cm=8)

    # ---------------- Panduan Pivot
    wp = wb.create_sheet("Panduan Pivot")
    teks_panduan(wp, 1, [
        "# Panduan Membuat PivotTable dan Slicer dari tblKandidat",
        "PivotTable tidak dapat dibuat otomatis oleh skrip pembangun, jadi ikuti langkah berikut lalu bandingkan hasilnya dengan sheet Ringkasan Rumus.",
        "## A. PivotTable: Hired per Sumber",
        "1. Klik sel mana pun di dalam tabel Kandidat (misalnya B5).",
        "2. Pilih menu Insert > PivotTable. Pastikan Table/Range terisi tblKandidat, pilih New Worksheet, klik OK.",
        "3. Di panel PivotTable Fields, seret 'Sumber' ke area Rows.",
        "4. Seret 'ID Kandidat' ke area Values (otomatis menjadi Count of ID Kandidat).",
        "5. Seret 'Status' ke area Columns. Kini terlihat jumlah Aktif / Hired / Rejected per sumber.",
        "6. Untuk hire rate: klik kanan angka > Show Values As > % of Row Total.",
        "## B. PivotTable: Rata-rata Time to Hire per Departemen",
        "1. Insert > PivotTable lagi (sheet baru). Seret 'Departemen' ke Rows.",
        "2. Seret 'Time to Hire (hari)' ke Values, klik panah kecil > Value Field Settings > Average > OK.",
        "3. Klik kanan angka > Number Format > Number, 1 desimal.",
        "## C. Slicer dan Timeline",
        "1. Klik PivotTable, lalu PivotTable Analyze > Insert Slicer. Centang 'Posisi' dan 'Rekruter' > OK.",
        "2. Klik tombol di slicer untuk memfilter; tahan Ctrl untuk memilih beberapa.",
        "3. PivotTable Analyze > Insert Timeline > centang 'Tanggal Lamar' > OK. Geser rentang bulan untuk memfilter periode.",
        "4. Agar satu slicer mengendalikan beberapa PivotTable: klik slicer > Slicer > Report Connections > centang semua pivot.",
        "## D. Pivot Chart Funnel",
        "1. Buat PivotTable dengan Rows = 'Tahap Terakhir' dan Values = Count of ID Kandidat.",
        "2. Klik PivotTable > PivotTable Analyze > PivotChart > Bar > Clustered Bar. Urutkan tahap manual dengan drag label.",
        "## E. Memperbarui",
        "Setelah menambah kandidat baru pada tabel, klik kanan PivotTable > Refresh (atau Ctrl+Alt+F5 untuk semua).",
    ])

    wd = wb.create_sheet("Daftar")
    header(wd, 1, 1, ["Posisi", "Departemen", "Sumber"])
    for i, (p, d) in enumerate(POSISI, start=2):
        wd.cell(row=i, column=1, value=p)
        wd.cell(row=i, column=2, value=d)
    for i, s in enumerate(SUMBER, start=2):
        wd.cell(row=i, column=3, value=s)
    lebar(wd, [28, 14, 20])
    wb.active = 0
    return rapikan(wb)


SCREENSHOT = [("Kandidat", "06_kandidat.png", "A1:Q32"), ("Funnel", "06_funnel.png", "A1:P22"),
              ("Ringkasan Rumus", "06_ringkasan_rumus.png", "A1:Q37")]
