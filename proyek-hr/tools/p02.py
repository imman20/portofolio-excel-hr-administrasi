"""Proyek 02 - Rekap Absensi Bulanan (tingkat pemula). Bulan: Maret 2024."""

import datetime as dt
import random

from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill

from data_gen import buat_karyawan
from util import (DEPARTEMEN, FMT_PCT, FMT_TGL, LIBUR_2024, WARNA_SUB, cf_kurang_dari, cf_sama_dengan, chart_bar,
                  dropdown, font, header, judul, kolom, label_nilai, lebar, rapikan, ref, tabel, teks_panduan,
                  tulis_baris)

N = 60
BULAN = dt.date(2024, 3, 1)
KODE = [("H", "Hadir"), ("S", "Sakit (dengan surat dokter)"), ("I", "Izin"), ("A", "Alpa / tanpa keterangan"),
        ("C", "Cuti")]


def hari_kerja(bulan):
    libur = {d for d, _ in LIBUR_2024}
    hari = []
    d = bulan
    while d.month == bulan.month:
        if d.weekday() < 5 and d not in libur:
            hari.append(d)
        d += dt.timedelta(days=1)
    return hari


def build():
    rng = random.Random(202)
    wb = Workbook()
    ws = wb.active
    ws.title = "Absensi"
    data = buat_karyawan(N, seed=101)[:N]
    hari = hari_kerja(BULAN)
    nh = len(hari)

    judul(ws, "Absensi Harian - Maret 2024", "Isi kode H / S / I / A / C per hari kerja. Kolom rekap di kanan dihitung otomatis.")
    k0 = 5  # kolom pertama tanggal (E)
    head = ["NIK", "Nama", "Departemen", "Jabatan"] + [h for h in hari] + \
           ["Hadir (H)", "Sakit (S)", "Izin (I)", "Alpa (A)", "Cuti (C)", "Hari Kerja", "% Kehadiran", "Keterangan",
            "Peringkat Terendah"]
    header(ws, 4, 1, head)
    for i, h in enumerate(hari):
        c = ws.cell(row=4, column=k0 + i)
        c.number_format = "dd"
        c.alignment = Alignment(horizontal="center", vertical="center", text_rotation=0)
    # baris 3: nama hari
    for i, h in enumerate(hari):
        c = ws.cell(row=3, column=k0 + i, value=f"=TEXT({kolom(k0 + i)}4,\"ddd\")")
        c.font = font(size=8, color="595959")
        c.alignment = Alignment(horizontal="center")
    ka = k0 + nh  # kolom Hadir
    kol_h, kol_s, kol_i, kol_a, kol_c, kol_hk, kol_pct, kol_ket, kol_rank = [kolom(ka + j) for j in range(9)]
    tgl_awal, tgl_akhir = kolom(k0), kolom(k0 + nh - 1)
    for r, k in enumerate(data, start=5):
        kode = []
        for _ in hari:
            kode.append(rng.choices(["H", "S", "I", "A", "C"], weights=[94, 2, 1.5, 1, 1.5])[0])
        if rng.random() < 0.08:  # beberapa karyawan sering absen
            for j in rng.sample(range(nh), 3):
                kode[j] = rng.choice(["A", "I", "S"])
        rentang = f"${tgl_awal}{r}:${tgl_akhir}{r}"
        nilai = [k["NIK"], k["Nama"], k["Departemen"], k["Jabatan"]] + kode + [
            f'=COUNTIF({rentang},"H")', f'=COUNTIF({rentang},"S")', f'=COUNTIF({rentang},"I")',
            f'=COUNTIF({rentang},"A")', f'=COUNTIF({rentang},"C")',
            f"=COUNT(${tgl_awal}$4:${tgl_akhir}$4)",
            f"=IFERROR({kol_h}{r}/{kol_hk}{r},0)",
            f'=IF({kol_a}{r}>=2,"Peringatan alpa",IF({kol_pct}{r}<0.9,"Kehadiran rendah","Baik"))',
            f"=_xlfn.RANK.EQ({kol_pct}{r},${kol_pct}$5:${kol_pct}${4 + N},1)+COUNTIF(${kol_pct}$5:{kol_pct}{r},{kol_pct}{r})-1",
        ]
        fmts = [None] * (4 + nh) + ["0"] * 6 + [FMT_PCT, None, "0"]
        tulis_baris(ws, r, 1, nilai, fmts)
        for j in range(nh):
            ws.cell(row=r, column=k0 + j).alignment = Alignment(horizontal="center")
    akhir = 4 + N
    ws.freeze_panes = "E5"
    lebar(ws, {"A": 10, "B": 22, "C": 12, "D": 20})
    for j in range(nh):
        ws.column_dimensions[kolom(k0 + j)].width = 3.6
    for c in [kol_h, kol_s, kol_i, kol_a, kol_c, kol_hk]:
        ws.column_dimensions[c].width = 7
    ws.column_dimensions[kol_pct].width = 10
    ws.column_dimensions[kol_ket].width = 17
    dropdown(ws, f"{tgl_awal}5:{tgl_akhir}{akhir}", [k for k, _ in KODE], pesan="Kode hanya H, S, I, A, atau C.")
    cf_kurang_dari(ws, f"{kol_pct}5:{kol_pct}{akhir}", 0.9)
    rentang_kode = f"{tgl_awal}5:{tgl_akhir}{akhir}"
    cf_sama_dengan(ws, rentang_kode, "A", "FFC7CE", "9C0006")
    cf_sama_dengan(ws, rentang_kode, "S", "FFEB9C", "9C5700")
    cf_sama_dengan(ws, rentang_kode, "I", "DDEBF7", "1F4E78")
    cf_sama_dengan(ws, rentang_kode, "C", "E2EFDA", "375623")
    ws.auto_filter.ref = f"A4:{kol_ket}{akhir}"

    # ---- Rekap departemen
    wr = wb.create_sheet("Rekap Departemen")
    judul(wr, "Rekap Kehadiran per Departemen - Maret 2024", "Rumus COUNTIF, SUMIF, dan pembagian persentase.")
    label_nilai(wr, 4, 1, "Jumlah hari kerja bulan ini", f"=Absensi!{kol_hk}5", "0")
    label_nilai(wr, 5, 1, "Jumlah karyawan", "=COUNTA(Absensi!A5:A500)", "0")
    label_nilai(wr, 6, 1, "Rata-rata kehadiran perusahaan", f"=AVERAGE(Absensi!{kol_pct}5:{kol_pct}{akhir})", FMT_PCT)
    label_nilai(wr, 7, 1, "Karyawan dengan kehadiran < 90%", f'=COUNTIF(Absensi!{kol_pct}5:{kol_pct}{akhir},"<0.9")', "0")
    header(wr, 9, 1, ["Departemen", "Jumlah Karyawan", "Total Hadir", "Total Sakit", "Total Izin", "Total Alpa",
                      "Total Cuti", "Total Hari Kerja", "% Kehadiran", "% Alpa"])
    for i, d in enumerate(DEPARTEMEN, start=10):
        tulis_baris(wr, i, 1, [
            d,
            f"=COUNTIF(Absensi!$C$5:$C${akhir},A{i})",
            f"=SUMIF(Absensi!$C$5:$C${akhir},A{i},Absensi!${kol_h}$5:${kol_h}${akhir})",
            f"=SUMIF(Absensi!$C$5:$C${akhir},A{i},Absensi!${kol_s}$5:${kol_s}${akhir})",
            f"=SUMIF(Absensi!$C$5:$C${akhir},A{i},Absensi!${kol_i}$5:${kol_i}${akhir})",
            f"=SUMIF(Absensi!$C$5:$C${akhir},A{i},Absensi!${kol_a}$5:${kol_a}${akhir})",
            f"=SUMIF(Absensi!$C$5:$C${akhir},A{i},Absensi!${kol_c}$5:${kol_c}${akhir})",
            f"=B{i}*$B$4",
            f"=IFERROR(C{i}/H{i},0)",
            f"=IFERROR(F{i}/H{i},0)",
        ], [None, "0", "0", "0", "0", "0", "0", "0", FMT_PCT, FMT_PCT])
    tulis_baris(wr, 16, 1, ["Total"] + [f"=SUM({kolom(c)}10:{kolom(c)}15)" for c in range(2, 9)] +
                ["=IFERROR(C16/H16,0)", "=IFERROR(F16/H16,0)"],
                [None] + ["0"] * 7 + [FMT_PCT, FMT_PCT])
    for c in range(1, 11):
        wr.cell(row=16, column=c).font = font(bold=True)
        wr.cell(row=16, column=c).fill = PatternFill("solid", fgColor=WARNA_SUB)
    cf_kurang_dari(wr, "I10:I15", 0.9)
    lebar(wr, [30, 12, 11, 11, 11, 11, 11, 13, 12, 10])
    chart_bar(wr, "A19", "Persentase Kehadiran per Departemen", ref(wr, 1, 10, 1, 15), ref(wr, 9, 9, 9, 15),
              judul_y="% Kehadiran", lebar_cm=15, tinggi_cm=8)
    chart_bar(wr, "G19", "Komposisi Ketidakhadiran per Departemen", ref(wr, 1, 10, 1, 15), ref(wr, 4, 9, 7, 15),
              judul_y="Hari", lebar_cm=15, tinggi_cm=8, tumpuk=True)

    # ---- Daftar karyawan bermasalah (rumus)
    header(wr, 37, 1, ["Peringkat", "Nama (kehadiran terendah)", "% Kehadiran", "Alpa"])
    for i in range(10):
        r = 38 + i
        tulis_baris(wr, r, 1, [
            i + 1,
            f"=INDEX(Absensi!$B$5:$B${akhir},MATCH(A{r},Absensi!${kol_rank}$5:${kol_rank}${akhir},0))",
            f"=INDEX(Absensi!${kol_pct}$5:${kol_pct}${akhir},MATCH(B{r},Absensi!$B$5:$B${akhir},0))",
            f"=INDEX(Absensi!${kol_a}$5:${kol_a}${akhir},MATCH(B{r},Absensi!$B$5:$B${akhir},0))",
        ], ["0", None, FMT_PCT, "0"])

    # ---- Kode & libur
    wk = wb.create_sheet("Kode & Libur")
    judul(wk, "Kode Absensi dan Hari Libur")
    header(wk, 3, 1, ["Kode", "Arti"])
    for i, (k, a) in enumerate(KODE, start=4):
        tulis_baris(wk, i, 1, [k, a])
    header(wk, 3, 4, ["Tanggal Libur 2024", "Keterangan"])
    for i, (d, ket) in enumerate(LIBUR_2024, start=4):
        tulis_baris(wk, i, 4, [d, ket], [FMT_TGL, None])
    lebar(wk, [8, 34, 3, 18, 36])

    # ---- Petunjuk
    wp = wb.create_sheet("Petunjuk")
    teks_panduan(wp, 1, [
        "# Petunjuk Penggunaan",
        "1. Sheet Absensi: satu baris per karyawan, satu kolom per hari kerja (Sabtu, Minggu, dan libur nasional sudah dikecualikan).",
        "2. Isi sel tanggal dengan kode H/S/I/A/C (dropdown). Warna sel berubah otomatis lewat Conditional Formatting.",
        "3. Kolom rekap (Hadir, Sakit, Izin, Alpa, Cuti) memakai COUNTIF terhadap rentang tanggal di baris yang sama.",
        "4. % Kehadiran = Hadir / Hari Kerja; sel merah bila di bawah 90%.",
        "5. Sheet Rekap Departemen memakai COUNTIF dan SUMIF untuk menjumlahkan per departemen, lengkap dengan chart.",
        "6. Untuk bulan lain: salin sheet Absensi, ganti tanggal di baris 4, kosongkan kode.",
    ])
    wb.active = 0
    return rapikan(wb)


SCREENSHOT = [("Absensi", "02_absensi.png", "A1:AF30"), ("Rekap Departemen", "02_rekap_departemen.png", "A1:J36")]
