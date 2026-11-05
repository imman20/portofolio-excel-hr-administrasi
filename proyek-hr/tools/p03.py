"""Proyek 03 - Administrasi Surat dan ATK (tingkat pemula)."""

import datetime as dt
import random

from openpyxl import Workbook
from openpyxl.styles import PatternFill

from util import (DEPARTEMEN, FMT_ANGKA, FMT_RP, FMT_TGL, WARNA_SUB, cf_rumus, cf_sama_dengan, chart_bar, dropdown,
                  font, header, judul, label_nilai, lebar, rapikan, ref, tabel, teks_panduan, tulis_baris)

PENGIRIM = ["PT Maju Jaya Abadi", "Dinas Tenaga Kerja Kota Surabaya", "BPJS Ketenagakerjaan Cabang Surabaya",
            "CV Sumber Rejeki", "Bank Mandiri KCP Darmo", "PT Telkom Indonesia", "Universitas Airlangga",
            "Kantor Pajak Pratama Surabaya", "PT Asuransi Sinar Mas", "Kementerian Ketenagakerjaan"]
PERIHAL_MASUK = ["Undangan rapat koordinasi", "Penawaran kerja sama pelatihan", "Permohonan data karyawan",
                 "Pemberitahuan iuran BPJS", "Permintaan magang mahasiswa", "Tagihan layanan internet",
                 "Konfirmasi rekening payroll", "Surat keputusan UMK", "Pengajuan audit internal",
                 "Undangan sosialisasi peraturan"]
PERIHAL_KELUAR = ["Surat keterangan kerja", "Surat panggilan interview", "Surat peringatan 1", "Surat tugas",
                  "Balasan permohonan magang", "Surat pengangkatan karyawan tetap", "Surat referensi",
                  "Surat permohonan pelatihan", "Surat undangan rapat", "Surat perjanjian kerja"]
JENIS = ["Undangan", "Pemberitahuan", "Permohonan", "Penawaran", "Keputusan", "Tagihan"]
STATUS_DISPOSISI = ["Belum", "Proses", "Selesai"]

ATK = [
    ("ATK-001", "Kertas HVS A4 80gsm", "Rim", 40, 10, 48000),
    ("ATK-002", "Kertas HVS F4 70gsm", "Rim", 20, 8, 52000),
    ("ATK-003", "Pulpen Hitam", "Box", 15, 5, 24000),
    ("ATK-004", "Pensil 2B", "Box", 8, 4, 18000),
    ("ATK-005", "Spidol Whiteboard", "Pcs", 30, 10, 8500),
    ("ATK-006", "Map Plastik Folio", "Pcs", 100, 30, 2500),
    ("ATK-007", "Amplop Coklat Folio", "Pack", 12, 5, 21000),
    ("ATK-008", "Stapler Besar", "Pcs", 6, 2, 35000),
    ("ATK-009", "Isi Stapler No. 10", "Box", 25, 10, 3500),
    ("ATK-010", "Sticky Notes", "Pack", 20, 8, 12000),
    ("ATK-011", "Binder Clip 25mm", "Box", 18, 6, 9500),
    ("ATK-012", "Tinta Printer Hitam", "Botol", 10, 4, 95000),
    ("ATK-013", "Tinta Printer Warna", "Botol", 8, 3, 110000),
    ("ATK-014", "Ordner Besar", "Pcs", 24, 8, 27000),
    ("ATK-015", "Buku Agenda", "Pcs", 10, 3, 32000),
]
ROMAWI = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII"]


def build():
    rng = random.Random(303)
    wb = Workbook()

    # ---------------- Surat Masuk
    ws = wb.active
    ws.title = "Surat Masuk"
    judul(ws, "Agenda Surat Masuk 2024", "Nomor agenda dibuat otomatis dengan TEXT dan ROW. Status disposisi memakai dropdown.")
    head = ["No. Agenda", "Tanggal Terima", "Nomor Surat Asal", "Pengirim", "Perihal", "Jenis", "Disposisi Ke",
            "Status Disposisi", "Tanggal Selesai", "Lama Proses (hari)", "Keterangan"]
    header(ws, 4, 1, head)
    tgl = dt.date(2024, 1, 3)
    n_masuk = 60
    for i in range(n_masuk):
        r = 5 + i
        tgl += dt.timedelta(days=rng.randint(1, 4))
        if tgl.weekday() >= 5:
            tgl += dt.timedelta(days=7 - tgl.weekday())
        status = rng.choices(STATUS_DISPOSISI, weights=[15, 25, 60])[0]
        if i > n_masuk - 8:
            status = rng.choice(["Belum", "Proses"])
        selesai = tgl + dt.timedelta(days=rng.randint(1, 10)) if status == "Selesai" else None
        tulis_baris(ws, r, 1, [
            f'="SM-"&TEXT(B{r},"yyyy")&"-"&TEXT(ROW()-4,"000")',
            tgl,
            f"{rng.randint(10, 999):03d}/{rng.choice(['DIR', 'HRD', 'KEU', 'UMUM'])}/{ROMAWI[tgl.month - 1]}/2024",
            rng.choice(PENGIRIM), rng.choice(PERIHAL_MASUK), rng.choice(JENIS), rng.choice(DEPARTEMEN),
            status, selesai,
            f'=IF(H{r}="Selesai",I{r}-B{r},IF(B{r}="","",TODAY()-B{r}))',
            "",
        ], [None, FMT_TGL, None, None, None, None, None, None, FMT_TGL, "0", None])
    akhir_m = 4 + n_masuk
    tabel(ws, f"A4:K{akhir_m}", "tblSuratMasuk")
    ws.freeze_panes = "C5"
    lebar(ws, [15, 13, 20, 34, 30, 15, 13, 14, 13, 11, 20])
    dropdown(ws, f"F5:F{akhir_m + 200}", "=Daftar!$A$2:$A$7")
    dropdown(ws, f"G5:G{akhir_m + 200}", "=Daftar!$B$2:$B$7")
    dropdown(ws, f"H5:H{akhir_m + 200}", "=Daftar!$C$2:$C$4")
    cf_sama_dengan(ws, f"H5:H{akhir_m}", "Belum", "FFC7CE", "9C0006")
    cf_sama_dengan(ws, f"H5:H{akhir_m}", "Proses", "FFEB9C", "9C5700")
    cf_sama_dengan(ws, f"H5:H{akhir_m}", "Selesai", "C6EFCE", "006100")
    cf_rumus(ws, f"J5:J{akhir_m}", 'AND($H5<>"Selesai",$J5>7)')

    # ---------------- Surat Keluar
    wk = wb.create_sheet("Surat Keluar")
    judul(wk, "Agenda Surat Keluar 2024", "Nomor surat: urut/HRD/bulan romawi/tahun, dibuat dengan TEXT, CHOOSE, dan MONTH.")
    header(wk, 4, 1, ["No. Urut", "Tanggal Surat", "Nomor Surat", "Tujuan", "Perihal", "Jenis", "Penanggung Jawab",
                      "Status Kirim", "Tanggal Kirim", "Arsip (Ya/Tidak)"])
    tgl = dt.date(2024, 1, 4)
    n_keluar = 45
    for i in range(n_keluar):
        r = 5 + i
        tgl += dt.timedelta(days=rng.randint(1, 6))
        if tgl.weekday() >= 5:
            tgl += dt.timedelta(days=7 - tgl.weekday())
        status = rng.choices(["Draft", "Terkirim"], weights=[15, 85])[0]
        if i > n_keluar - 5:
            status = "Draft"
        tulis_baris(wk, r, 1, [
            "=ROW()-4", tgl,
            f'=TEXT(A{r},"000")&"/HRD/"&CHOOSE(MONTH(B{r}),"I","II","III","IV","V","VI","VII","VIII","IX","X","XI","XII")&"/"&YEAR(B{r})',
            rng.choice(PENGIRIM + ["Karyawan ybs.", "Pelamar ybs."]), rng.choice(PERIHAL_KELUAR), rng.choice(JENIS),
            rng.choice(["Staff HR", "Supervisor HR", "Manager HR"]), status,
            tgl + dt.timedelta(days=rng.randint(0, 2)) if status == "Terkirim" else None,
            "Ya" if status == "Terkirim" else "Tidak",
        ], ["0", FMT_TGL, None, None, None, None, None, None, FMT_TGL, None])
    akhir_k = 4 + n_keluar
    tabel(wk, f"A4:J{akhir_k}", "tblSuratKeluar")
    wk.freeze_panes = "D5"
    lebar(wk, [8, 13, 22, 34, 30, 15, 16, 12, 13, 14])
    dropdown(wk, f"H5:H{akhir_k + 200}", ["Draft", "Terkirim"])
    dropdown(wk, f"J5:J{akhir_k + 200}", ["Ya", "Tidak"])
    cf_sama_dengan(wk, f"H5:H{akhir_k}", "Draft", "FFEB9C", "9C5700")

    # ---------------- Ringkasan surat
    wr = wb.create_sheet("Ringkasan Surat")
    judul(wr, "Ringkasan Agenda Surat", "COUNTIF / COUNTIFS terhadap tabel surat masuk dan keluar.")
    label_nilai(wr, 4, 1, "Total surat masuk", f"=COUNTA('Surat Masuk'!B5:B500)", "0")
    label_nilai(wr, 5, 1, "Surat masuk belum didisposisi", f'=COUNTIF(\'Surat Masuk\'!H5:H500,"Belum")', "0")
    label_nilai(wr, 6, 1, "Surat masuk dalam proses", f'=COUNTIF(\'Surat Masuk\'!H5:H500,"Proses")', "0")
    label_nilai(wr, 7, 1, "Surat masuk selesai", f'=COUNTIF(\'Surat Masuk\'!H5:H500,"Selesai")', "0")
    label_nilai(wr, 8, 1, "Rata-rata lama proses selesai (hari)",
                "=IFERROR(AVERAGEIF('Surat Masuk'!H5:H500,\"Selesai\",'Surat Masuk'!J5:J500),0)", "0.0")
    label_nilai(wr, 9, 1, "Total surat keluar", "=COUNTA('Surat Keluar'!B5:B500)", "0")
    label_nilai(wr, 10, 1, "Surat keluar masih draft", "=COUNTIF('Surat Keluar'!H5:H500,\"Draft\")", "0")
    header(wr, 12, 1, ["Bulan", "Surat Masuk", "Surat Keluar", "Masuk Selesai", "% Selesai"])
    for m in range(1, 13):
        r = 12 + m
        awal = dt.date(2024, m, 1)
        akhir = dt.date(2024 if m < 12 else 2025, m + 1 if m < 12 else 1, 1)
        wr.cell(row=r, column=7, value=awal).number_format = FMT_TGL
        wr.cell(row=r, column=8, value=akhir).number_format = FMT_TGL
        tulis_baris(wr, r, 1, [
            f'=TEXT(G{r},"mmmm yyyy")',
            f"=COUNTIFS('Surat Masuk'!$B$5:$B$500,\">=\"&G{r},'Surat Masuk'!$B$5:$B$500,\"<\"&H{r})",
            f"=COUNTIFS('Surat Keluar'!$B$5:$B$500,\">=\"&G{r},'Surat Keluar'!$B$5:$B$500,\"<\"&H{r})",
            f"=COUNTIFS('Surat Masuk'!$B$5:$B$500,\">=\"&G{r},'Surat Masuk'!$B$5:$B$500,\"<\"&H{r},'Surat Masuk'!$H$5:$H$500,\"Selesai\")",
            f"=IFERROR(D{r}/B{r},0)",
        ], [None, "0", "0", "0", "0.0%"])
    wr["G12"], wr["H12"] = "Awal bulan (bantu)", "Akhir bulan (bantu)"
    wr["G12"].font = wr["H12"].font = font(italic=True, color="595959")
    header(wr, 27, 1, ["Disposisi Ke", "Jumlah Surat", "Belum", "Proses", "Selesai"])
    for i, d in enumerate(DEPARTEMEN, start=28):
        tulis_baris(wr, i, 1, [
            d, f"=COUNTIF('Surat Masuk'!$G$5:$G$500,A{i})",
            f"=COUNTIFS('Surat Masuk'!$G$5:$G$500,$A{i},'Surat Masuk'!$H$5:$H$500,C$27)",
            f"=COUNTIFS('Surat Masuk'!$G$5:$G$500,$A{i},'Surat Masuk'!$H$5:$H$500,D$27)",
            f"=COUNTIFS('Surat Masuk'!$G$5:$G$500,$A{i},'Surat Masuk'!$H$5:$H$500,E$27)",
        ], [None, "0", "0", "0", "0"])
    lebar(wr, [36, 13, 13, 13, 11, 2, 18, 18])
    chart_bar(wr, "J4", "Volume Surat per Bulan 2024", ref(wr, 1, 13, 1, 24), ref(wr, 2, 12, 3, 24), judul_y="Surat",
              lebar_cm=16, tinggi_cm=8)

    # ---------------- Master ATK
    wa = wb.create_sheet("Master ATK")
    judul(wa, "Master Barang ATK", "Daftar barang, satuan, stok awal tahun, stok minimum, dan harga satuan.")
    header(wa, 4, 1, ["Kode", "Nama Barang", "Satuan", "Stok Awal", "Stok Minimum", "Harga Satuan"])
    for i, a in enumerate(ATK, start=5):
        tulis_baris(wa, i, 1, list(a), [None, None, None, "0", "0", FMT_RP])
    tabel(wa, f"A4:F{4 + len(ATK)}", "tblATK")
    lebar(wa, [10, 26, 9, 10, 13, 14])

    # ---------------- Transaksi ATK
    wt = wb.create_sheet("Transaksi ATK")
    judul(wt, "Transaksi Barang Masuk / Keluar", "Nama barang dan harga diambil dengan INDEX/MATCH dari Master ATK.")
    header(wt, 4, 1, ["Tanggal", "Kode", "Nama Barang", "Jenis", "Jumlah", "Satuan", "Harga Satuan", "Nilai",
                      "Departemen Pemakai", "Keterangan"])
    tgl = dt.date(2024, 1, 2)
    r = 5
    n_trx = 120
    stok = {a[0]: a[3] for a in ATK}
    stok_min = {a[0]: a[4] for a in ATK}
    for i in range(n_trx):
        tgl += dt.timedelta(days=rng.randint(0, 3))
        if tgl > dt.date(2024, 6, 30):
            break
        kode = rng.choice(ATK)[0]
        jenis = rng.choices(["Masuk", "Keluar"], weights=[25, 75])[0]
        jumlah = rng.randint(5, 30) if jenis == "Masuk" else rng.randint(1, 6)
        if jenis == "Keluar" and stok[kode] - jumlah < 0:
            jenis, jumlah = "Masuk", rng.randint(10, 30)
        # pembelian dilakukan saat stok mendekati minimum, kecuali beberapa barang yang sengaja dibiarkan
        if jenis == "Keluar" and stok[kode] - jumlah <= stok_min[kode] and kode not in ("ATK-012", "ATK-014"):
            jenis, jumlah = "Masuk", rng.randint(10, 30)
        stok[kode] += jumlah if jenis == "Masuk" else -jumlah
        tulis_baris(wt, r, 1, [
            tgl, kode,
            f'=IFERROR(INDEX(\'Master ATK\'!$B$5:$B$19,MATCH(B{r},\'Master ATK\'!$A$5:$A$19,0)),"")',
            jenis, jumlah,
            f'=IFERROR(INDEX(\'Master ATK\'!$C$5:$C$19,MATCH(B{r},\'Master ATK\'!$A$5:$A$19,0)),"")',
            f'=IFERROR(INDEX(\'Master ATK\'!$F$5:$F$19,MATCH(B{r},\'Master ATK\'!$A$5:$A$19,0)),0)',
            f"=E{r}*G{r}",
            rng.choice(DEPARTEMEN) if jenis == "Keluar" else "Gudang",
            "Pembelian" if jenis == "Masuk" else "Permintaan rutin",
        ], [FMT_TGL, None, None, None, "0", None, FMT_RP, FMT_RP, None, None])
        r += 1
    akhir_t = r - 1
    tabel(wt, f"A4:J{akhir_t}", "tblTransaksi")
    wt.freeze_panes = "C5"
    lebar(wt, [12, 10, 26, 9, 9, 9, 14, 14, 18, 18])
    dropdown(wt, f"B5:B{akhir_t + 300}", "='Master ATK'!$A$5:$A$19")
    dropdown(wt, f"D5:D{akhir_t + 300}", ["Masuk", "Keluar"])
    dropdown(wt, f"I5:I{akhir_t + 300}", "=Daftar!$B$2:$B$8")

    # ---------------- Kartu Stok
    wsk = wb.create_sheet("Kartu Stok")
    judul(wsk, "Kartu Stok ATK", "Stok akhir = stok awal + masuk - keluar (SUMIFS). Status PESAN jika di bawah stok minimum.")
    header(wsk, 4, 1, ["Kode", "Nama Barang", "Satuan", "Stok Awal", "Masuk", "Keluar", "Stok Akhir", "Stok Minimum",
                       "Status", "Harga Satuan", "Nilai Persediaan"])
    for i in range(len(ATK)):
        r = 5 + i
        mr = 5 + i
        tulis_baris(wsk, r, 1, [
            f"='Master ATK'!A{mr}", f"='Master ATK'!B{mr}", f"='Master ATK'!C{mr}", f"='Master ATK'!D{mr}",
            f"=SUMIFS('Transaksi ATK'!$E$5:$E$500,'Transaksi ATK'!$B$5:$B$500,A{r},'Transaksi ATK'!$D$5:$D$500,\"Masuk\")",
            f"=SUMIFS('Transaksi ATK'!$E$5:$E$500,'Transaksi ATK'!$B$5:$B$500,A{r},'Transaksi ATK'!$D$5:$D$500,\"Keluar\")",
            f"=D{r}+E{r}-F{r}", f"='Master ATK'!E{mr}",
            f'=IF(G{r}<=H{r},"PESAN","Aman")', f"='Master ATK'!F{mr}", f"=G{r}*J{r}",
        ], [None, None, None, "0", "0", "0", "0", "0", None, FMT_RP, FMT_RP])
    akhir_s = 4 + len(ATK)
    tulis_baris(wsk, akhir_s + 1, 1, ["Total", "", "", f"=SUM(D5:D{akhir_s})", f"=SUM(E5:E{akhir_s})",
                                      f"=SUM(F5:F{akhir_s})", f"=SUM(G5:G{akhir_s})", "", f'=COUNTIF(I5:I{akhir_s},"PESAN")&" item perlu dipesan"',
                                      "", f"=SUM(K5:K{akhir_s})"],
                [None, None, None, "0", "0", "0", "0", None, None, None, FMT_RP])
    for c in range(1, 12):
        wsk.cell(row=akhir_s + 1, column=c).font = font(bold=True)
        wsk.cell(row=akhir_s + 1, column=c).fill = PatternFill("solid", fgColor=WARNA_SUB)
    cf_sama_dengan(wsk, f"I5:I{akhir_s}", "PESAN")
    cf_rumus(wsk, f"G5:G{akhir_s}", "$G5<=$H5")
    lebar(wsk, [10, 26, 9, 10, 9, 9, 10, 12, 10, 14, 16])
    chart_bar(wsk, "A23", "Stok Akhir vs Stok Minimum", ref(wsk, 2, 5, 2, akhir_s), ref(wsk, 7, 4, 8, akhir_s),
              judul_y="Unit", lebar_cm=24, tinggi_cm=9)

    # ---------------- Laporan Bulanan ATK
    wl = wb.create_sheet("Laporan Bulanan ATK")
    judul(wl, "Laporan Pemakaian ATK per Bulan (Semester 1 2024)", "SUMIFS dengan kriteria kode barang, jenis Keluar, dan rentang tanggal.")
    bulan = [dt.date(2024, m, 1) for m in range(1, 8)]
    header(wl, 4, 1, ["Kode", "Nama Barang"] + [b.strftime("%b %Y") for b in bulan[:6]] + ["Total Keluar", "Nilai Pemakaian"])
    for j, b in enumerate(bulan):
        wl.cell(row=3, column=3 + j, value=b).number_format = FMT_TGL
        wl.cell(row=3, column=3 + j).font = font(italic=True, size=8, color="595959")
    wl["A3"] = "Baris 3 = awal bulan (bantu untuk rumus)"
    wl["A3"].font = font(italic=True, size=8, color="595959")
    for i in range(len(ATK)):
        r = 5 + i
        mr = 5 + i
        baris = [f"='Master ATK'!A{mr}", f"='Master ATK'!B{mr}"]
        for j in range(6):
            kc = chr(ord("C") + j)
            kn = chr(ord("C") + j + 1)
            baris.append(
                f"=SUMIFS('Transaksi ATK'!$E$5:$E$500,'Transaksi ATK'!$B$5:$B$500,$A{r},'Transaksi ATK'!$D$5:$D$500,\"Keluar\","
                f"'Transaksi ATK'!$A$5:$A$500,\">=\"&{kc}$3,'Transaksi ATK'!$A$5:$A$500,\"<\"&{kn}$3)")
        baris.append(f"=SUM(C{r}:H{r})")
        baris.append(f"=I{r}*'Master ATK'!F{mr}")
        tulis_baris(wl, r, 1, baris, [None, None] + ["0"] * 7 + [FMT_RP])
    akhir_l = 4 + len(ATK)
    tulis_baris(wl, akhir_l + 1, 1, ["Total", ""] + [f"=SUM({chr(ord('C') + j)}5:{chr(ord('C') + j)}{akhir_l})" for j in range(7)]
                + [f"=SUM(J5:J{akhir_l})"], [None, None] + ["0"] * 7 + [FMT_RP])
    for c in range(1, 11):
        wl.cell(row=akhir_l + 1, column=c).font = font(bold=True)
        wl.cell(row=akhir_l + 1, column=c).fill = PatternFill("solid", fgColor=WARNA_SUB)
    lebar(wl, [10, 26, 10, 10, 10, 10, 10, 10, 12, 16])
    chart_bar(wl, "L4", "Nilai Pemakaian ATK per Barang", ref(wl, 2, 5, 2, akhir_l), ref(wl, 10, 4, 10, akhir_l),
              judul_y="Rupiah", lebar_cm=16, tinggi_cm=10, horizontal=True)

    # ---------------- Daftar
    wd = wb.create_sheet("Daftar")
    header(wd, 1, 1, ["Jenis Surat", "Departemen", "Status Disposisi"])
    for i, j in enumerate(JENIS, start=2):
        wd.cell(row=i, column=1, value=j)
    for i, d in enumerate(DEPARTEMEN + ["Gudang"], start=2):
        wd.cell(row=i, column=2, value=d)
    for i, s in enumerate(STATUS_DISPOSISI, start=2):
        wd.cell(row=i, column=3, value=s)
    lebar(wd, [16, 14, 16])

    wp = wb.create_sheet("Petunjuk")
    teks_panduan(wp, 1, [
        "# Petunjuk Penggunaan",
        "1. Surat Masuk: tambahkan baris baru di bawah tabel; nomor agenda SM-2024-001 dst. terbentuk otomatis dari TEXT(ROW()-4;\"000\").",
        "2. Surat Keluar: nomor surat (001/HRD/I/2024) dibuat otomatis dari nomor urut, bulan romawi (CHOOSE+MONTH), dan tahun.",
        "3. Transaksi ATK: cukup pilih kode barang dan jenis Masuk/Keluar; nama, satuan, dan harga ditarik dari Master ATK dengan INDEX/MATCH.",
        "4. Kartu Stok: SUMIFS menjumlahkan masuk dan keluar per kode. Status PESAN muncul bila stok akhir <= stok minimum.",
        "5. Laporan Bulanan: SUMIFS dengan rentang tanggal (>= awal bulan dan < awal bulan berikutnya).",
        "6. Gunakan filter pada header tabel untuk menyaring status disposisi atau departemen tertentu.",
    ])
    wb.active = 0
    return rapikan(wb)


SCREENSHOT = [("Surat Masuk", "03_surat_masuk.png", "A1:K34"), ("Kartu Stok", "03_kartu_stok.png", "A1:K40"),
              ("Laporan Bulanan ATK", "03_laporan_bulanan_atk.png", "A1:U22")]
