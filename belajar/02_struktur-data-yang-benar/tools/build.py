"""
Pembangun modul 02 - Struktur data yang benar.
Jalankan dari folder modul: python tools/build.py
Menghasilkan latihan/data_02.csv, latihan/latihan_02.xlsx, latihan/solusi_02.xlsx.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font

HERE = Path(__file__).resolve().parent
MOD = HERE.parent
sys.path.insert(0, str(MOD.parent / "tools"))
import belajar_lib as L  # noqa: E402
import course_runtime as R

SEED = 102
N = 240
OUT = MOD / "latihan"
CABANG = {"001": "Surabaya", "002": "Malang", "003": "Sidoarjo"}
BULAN = ["2024-01", "2024-02", "2024-03", "2024-04"]


def buat_data() -> pd.DataFrame:
    r = L.rng(SEED)
    tgl = L.tanggal_acak(r, N, "2024-01-01", "2024-04-30")
    prod = [L.PRODUK_TOKO[i] for i in r.integers(0, len(L.PRODUK_TOKO), N)]
    cab = r.choice(list(CABANG.keys()), N, p=[0.5, 0.3, 0.2])
    df = pd.DataFrame({
        "no": range(1, N + 1),
        "tanggal": tgl,
        "cabang_id": cab,
        "nama_cabang": [CABANG[c] for c in cab],
        "kode_produk": [p[0] for p in prod],
        "kategori": [p[2] for p in prod],
        "qty": r.integers(1, 15, N),
        "harga_satuan": [p[3] for p in prod],
    })
    df["total"] = df["qty"] * df["harga_satuan"]
    df["periode"] = pd.to_datetime(df["tanggal"]).dt.strftime("%Y-%m")
    return df


def ringkasan_wide(df: pd.DataFrame) -> pd.DataFrame:
    pv = df.pivot_table(index="nama_cabang", columns="periode", values="total", aggfunc="sum", fill_value=0)
    return pv.reindex(list(CABANG.values()))[BULAN]


INSTRUKSI = [
    "# Modul 02 - Struktur data yang benar: workbook latihan",
    "",
    "Sheet dalam workbook ini:",
    "- 'Laporan_Lama': contoh ANTI-POLA (judul merged, blok kolom per bulan, subtotal di tengah data, angka tersimpan sebagai teks). Jangan ditiru; dipakai untuk latihan membaca masalah.",
    "- 'Data': 240 transaksi dalam bentuk tidy (satu baris = satu transaksi). Di file latihan masih berupa range biasa; ubah menjadi Excel Table bernama tblPenjualan.",
    "- 'ID_Rusak': daftar cabang yang ID-nya kehilangan nol di depan.",
    "- 'Latihan': sel kuning untuk diisi.",
    "",
    "## Tugas",
    "1. Di sheet Data: klik A1, tekan Ctrl+T, centang 'My table has headers', lalu di Table Design ubah Table Name menjadi tblPenjualan.",
    "2. Latihan bagian A: periksa tipe data dengan COUNT vs COUNTA.",
    "3. Bagian B: perbaiki ID dengan TEXT(...;\"000\") dan dengan custom number format 000; amati bedanya.",
    "4. Bagian C: tulis rumus dengan structured references (tblPenjualan[total]).",
    "5. Bagian D: buktikan bahwa angka Laporan_Lama bisa dihasilkan ulang dari tabel tidy dengan SUMIFS.",
    "6. Bagian E: ubah tabel lebar (wide) kecil menjadi tabel panjang (long) 12 baris.",
    "",
    "Bandingkan dengan solusi_02.xlsx dan kunci_jawaban.md.",
] + [""] + L.kalimat_latihan()


def bangun(df: pd.DataFrame, solusi: bool):
    wb = L.wb_baru()
    L.sheet_teks(wb, "Instruksi", INSTRUKSI)

    # ---------- Laporan_Lama (anti-pola) ----------
    wl = wb.create_sheet("Laporan_Lama")
    wl.merge_cells("A1:F1")
    wl["A1"] = "LAPORAN PENJUALAN CABANG JAWA TIMUR  JAN - APR 2024"
    wl["A1"].font = Font(name=L.FONT_NAME, bold=True, size=12)
    wl["A1"].alignment = Alignment(horizontal="center")
    wl["A2"] = "(angka dalam rupiah; disusun manual)"
    wide = ringkasan_wide(df)
    L.tulis_header(wl, 4, 1, ["Cabang", "Jan", "Feb", "Mar", "Apr", "Total"])
    r = 5
    for i, (cab, row) in enumerate(wide.iterrows()):
        wl.cell(row=r, column=1, value=cab)
        for j, b in enumerate(BULAN):
            v = int(row[b])
            c = wl.cell(row=r, column=2 + j)
            # anti-pola: beberapa angka tersimpan sebagai teks
            if (i + j) % 3 == 0:
                c.value = str(v)
                c.alignment = Alignment(horizontal="left")
            else:
                c.value = v
                c.number_format = L.FMT_INT
        wl.cell(row=r, column=6, value=f"=VALUE(B{r})+VALUE(C{r})+VALUE(D{r})+VALUE(E{r})").number_format = L.FMT_INT
        r += 1
        if cab == "Malang":
            wl.cell(row=r, column=1, value="Subtotal 2 cabang").font = L.BOLD_FONT
            sub = wide.loc[["Surabaya", "Malang"]].sum()
            for j, b in enumerate(BULAN):
                wl.cell(row=r, column=2 + j, value=f"=VALUE({chr(66+j)}5)+VALUE({chr(66+j)}6)").number_format = L.FMT_INT
            wl.cell(row=r, column=6, value="=F5+F6").number_format = L.FMT_INT
            r += 1
    wl.cell(row=r, column=1, value="TOTAL").font = L.BOLD_FONT
    tot = wide.sum()
    for j, b in enumerate(BULAN):
        wl.cell(row=r, column=2 + j, value=f"=VALUE({chr(66+j)}5)+VALUE({chr(66+j)}6)+VALUE({chr(66+j)}8)").number_format = L.FMT_INT
    wl.cell(row=r, column=6, value="=F5+F6+F8").number_format = L.FMT_INT
    wl.cell(row=r + 2, column=1, value="Catatan: sheet ini sengaja dibuat sebagai contoh struktur yang SALAH untuk analisis.")
    L.autofit(wl)
    wl.column_dimensions["A"].width = 22

    # ---------- Data (tidy) ----------
    ws = wb.create_sheet("Data")
    kolom = ["no", "tanggal", "cabang_id", "nama_cabang", "kode_produk", "kategori", "qty", "harga_satuan", "total"]
    fmts = {"tanggal": L.FMT_DATE, "harga_satuan": L.FMT_INT, "total": L.FMT_INT, "qty": "0"}
    L.tulis_df(ws, df[kolom], formats=fmts, font=L.INPUT_FONT, as_table="tblPenjualan" if solusi else None)
    for i in range(2, N + 2):
        ws.cell(row=i, column=3).number_format = "@"
    for i in range(2,N+2):
        ws.cell(i,9,f"=G{i}*H{i}")
    last = N + 1

    # ---------- ID_Rusak ----------
    wi = wb.create_sheet("ID_Rusak")
    L.judul(wi, "Daftar cabang: ID kehilangan nol di depan", "A1",
            "Kolom A diketik sebagai angka (1, 2, 3) padahal ID resmi adalah 001, 002, 003.")
    L.tulis_header(wi, 3, 1, ["cabang_id (angka)", "nama_cabang", "ID benar via TEXT", "ID benar via format 000", "Sama dengan tabel Data? (COUNTIFS)"])
    for i, (k, v) in enumerate(CABANG.items(), start=4):
        L.sel_input(wi, f"A{i}", int(k))
        wi[f"B{i}"] = v
        wi[f"B{i}"].border = L.BORDER
        L.sel_jawab(wi, f"C{i}", f'=TEXT(A{i},"000")' if solusi else None)
        L.sel_jawab(wi, f"D{i}", f"=A{i}" if solusi else None, "000")
        L.sel_jawab(wi, f"E{i}", f"=COUNTIFS(Data!$C$2:$C${last},C{i})" if solusi else None, L.FMT_INT)
    wi["A8"] = "Perhatikan: kolom D tampil 001 tetapi nilainya tetap angka 1; kolom C benar-benar teks \"001\"."
    wi["A9"] = "COUNTIFS pada kolom E memakai teks dari kolom C sehingga cocok dengan cabang_id di sheet Data (teks)."
    L.autofit(wi)

    # ---------- Latihan ----------
    wx = wb.create_sheet("Latihan")
    L.judul(wx, "Modul 02 - Latihan struktur data", "A1", "Isi sel kuning. Data: Data!A2:I241 (tabel tblPenjualan).")
    wx.column_dimensions["A"].width = 3
    wx.column_dimensions["B"].width = 60
    for col in "CDEFG":
        wx.column_dimensions[col].width = 16

    L.subjudul(wx, "B3", "Bagian A - Pemeriksaan tipe data (COUNT hanya menghitung angka, COUNTA semua yang terisi)")
    L.tulis_header(wx, 4, 2, ["Rentang", "COUNT", "COUNTA", "Sel bukan angka (COUNTA-COUNT)"])
    barisA = [
        ("Data!B2:B241 (tanggal)", "Data!$B$2:$B$241"),
        ("Data!C2:C241 (cabang_id, sengaja teks)", "Data!$C$2:$C$241"),
        ("Data!G2:G241 (qty)", "Data!$G$2:$G$241"),
        ("Laporan_Lama!B5:E8 (angka bulanan, sebagian teks)", "Laporan_Lama!$B$5:$E$8"),
    ]
    for i, (lab, rng) in enumerate(barisA, start=5):
        wx[f"B{i}"] = lab
        wx[f"B{i}"].border = L.BORDER
        L.sel_jawab(wx, f"C{i}", f"=COUNT({rng})" if solusi else None, L.FMT_INT)
        L.sel_jawab(wx, f"D{i}", f"=COUNTA({rng})" if solusi else None, L.FMT_INT)
        L.sel_jawab(wx, f"E{i}", f"=D{i}-C{i}" if solusi else None, L.FMT_INT)

    L.subjudul(wx, "B10", "Bagian B - ID dengan nol di depan (lihat sheet ID_Rusak)")
    wx["B11"] = "Ketik 7 di C11, lalu beri format custom 0000 (tampil 0007, nilai tetap 7)"
    L.sel_jawab(wx, "C11", 7 if solusi else None, "0000", font=L.INPUT_FONT)
    wx["B12"] = "=TEXT(C11;\"0000\")  -> teks \"0007\""
    L.sel_jawab(wx, "C12", '=TEXT(C11,"0000")' if solusi else None)
    wx["B13"] = "=C11*2 (angka berformat masih bisa dihitung)"
    L.sel_jawab(wx, "C13", "=C11*2" if solusi else None)
    wx["B14"] = "=COUNT(C12) -> 0 karena C12 teks"
    L.sel_jawab(wx, "C14", "=COUNT(C12)" if solusi else None)
    wx["B15"] = "=VALUE(C12) mengembalikan teks menjadi angka 7"
    L.sel_jawab(wx, "C15", "=VALUE(C12)" if solusi else None)

    L.subjudul(wx, "B17", "Bagian C - Structured references pada tblPenjualan")
    barisC = [
        ("Total penjualan: =SUM(tblPenjualan[total])", "=SUM(tblPenjualan[total])", L.FMT_INT),
        ("Jumlah transaksi: =COUNTA(tblPenjualan[no])", "=COUNTA(tblPenjualan[no])", L.FMT_INT),
        ("Penjualan Surabaya: =SUMIFS(tblPenjualan[total];tblPenjualan[nama_cabang];\"Surabaya\")",
         '=SUMIFS(tblPenjualan[total],tblPenjualan[nama_cabang],"Surabaya")', L.FMT_INT),
        ("Qty terbesar satu transaksi: =MAX(tblPenjualan[qty])", "=MAX(tblPenjualan[qty])", L.FMT_INT),
        ("Rata-rata harga satuan: =AVERAGE(tblPenjualan[harga_satuan])", "=AVERAGE(tblPenjualan[harga_satuan])", L.FMT_DEC2),
    ]
    for i, (lab, rumus, fmt) in enumerate(barisC, start=18):
        wx[f"B{i}"] = lab
        wx[f"B{i}"].border = L.BORDER
        L.sel_jawab(wx, f"C{i}", rumus if solusi else None, fmt)
    wx["D17"] = "Di file latihan, buat tabel dulu (Ctrl+T, nama tblPenjualan); tanpa tabel, rumus ini #NAME?."
    wx["D17"].font = Font(name=L.FONT_NAME, size=9, italic=True)

    L.subjudul(wx, "B24", "Bagian D - Laporan_Lama dihasilkan ulang dari tabel tidy (SUMIFS per cabang per bulan)")
    L.tulis_header(wx, 25, 2, ["Cabang", "Jan", "Feb", "Mar", "Apr", "Total", "Laporan_Lama!F", "Selisih"])
    # kunci periode: tanggal >= awal bulan dan < awal bulan berikutnya
    awal = {"Jan": ("2024-01-01", "2024-02-01"), "Feb": ("2024-02-01", "2024-03-01"),
            "Mar": ("2024-03-01", "2024-04-01"), "Apr": ("2024-04-01", "2024-05-01")}
    baris_lama = {"Surabaya": 5, "Malang": 6, "Sidoarjo": 8}
    for i, cab in enumerate(CABANG.values(), start=26):
        wx[f"B{i}"] = cab
        wx[f"B{i}"].border = L.BORDER
        for j, b in enumerate(["Jan", "Feb", "Mar", "Apr"]):
            col = "CDEF"[j]
            a0, a1 = awal[b]
            rumus = (f'=SUMIFS(Data!$I$2:$I${last},Data!$D$2:$D${last},$B{i},'
                     f'Data!$B$2:$B${last},">="&DATE({a0[:4]},{int(a0[5:7])},1),'
                     f'Data!$B$2:$B${last},"<"&DATE({a1[:4]},{int(a1[5:7])},1))')
            L.sel_jawab(wx, f"{col}{i}", rumus if solusi else None, L.FMT_INT)
        L.sel_jawab(wx, f"G{i}", f"=SUM(C{i}:F{i})" if solusi else None, L.FMT_INT)
        L.sel_rumus(wx, f"H{i}", f"=Laporan_Lama!F{baris_lama[cab]}", L.FMT_INT, link=True)
        L.sel_jawab(wx, f"I{i}", f"=G{i}-H{i}" if solusi else None, L.FMT_INT)
    wx["B29"] = "Total"
    wx["B29"].font = L.BOLD_FONT
    L.sel_jawab(wx, "G29", "=SUM(G26:G28)" if solusi else None, L.FMT_INT)
    L.sel_rumus(wx, "H29", "=Laporan_Lama!F9", L.FMT_INT, link=True)
    L.sel_jawab(wx, "I29", "=G29-H29" if solusi else None, L.FMT_INT)
    wx.column_dimensions["H"].width = 16
    wx.column_dimensions["I"].width = 12

    L.subjudul(wx, "B31", "Bagian E - Wide ke long: 3 cabang x 4 bulan menjadi 12 baris")
    wx["B32"] = "Tabel lebar (sumber):"
    L.tulis_header(wx, 33, 2, ["cabang", "Jan", "Feb", "Mar", "Apr"])
    for i, cab in enumerate(CABANG.values(), start=34):
        wx[f"B{i}"] = cab
        wx[f"B{i}"].border = L.BORDER
        for j, col in enumerate("CDEF"):
            L.sel_rumus(wx, f"{col}{i}", f"={col}{26 + i - 34}", L.FMT_INT)
    wx["B38"] = "Tabel panjang (hasil): satu baris per cabang per bulan"
    L.tulis_header(wx, 39, 2, ["cabang", "bulan", "penjualan"])
    k = 40
    for i, cab in enumerate(CABANG.values()):
        for j, b in enumerate(["Jan", "Feb", "Mar", "Apr"]):
            L.sel_jawab(wx, f"B{k}", cab if solusi else None, font=L.INPUT_FONT)
            L.sel_jawab(wx, f"C{k}", b if solusi else None, font=L.INPUT_FONT)
            L.sel_jawab(wx, f"D{k}", f"={'CDEF'[j]}{34 + i}" if solusi else None, L.FMT_INT)
            k += 1
    wx[f"B{k}"] = "Pemeriksaan: SUM tabel panjang = SUM tabel lebar"
    L.sel_jawab(wx, f"D{k}", f"=SUM(D40:D{k - 1})-SUM(C34:F36)" if solusi else None, L.FMT_INT)
    L.legenda(wx, f"B{k + 3}")
    wx.sheet_view.showGridLines = False
    for addr in ["B12","B13","B14","B15"]:
        wx[addr]="Contoh: "+wx[addr].value
    wb.active = 0
    return R.finish(wb)


def main() -> None:
    df = buat_data()
    L.simpan_csv(df.drop(columns=["periode","total"]), OUT / "data_02.csv")
    L.simpan(bangun(df, solusi=False), OUT / "latihan_02.xlsx")
    L.simpan(bangun(df, solusi=True), OUT / "solusi_02.xlsx")
    print("selesai:", OUT)


def checks() -> list[dict]:
    df = pd.read_csv(OUT/"data_02.csv",dtype={"cabang_id":str})
    df["tanggal"]=pd.to_datetime(df["tanggal"])
    df["total"]=df.qty*df.harga_satuan
    df["periode"]=df.tanggal.dt.strftime("%Y-%m")
    f = "solusi_02.xlsx"
    wide = ringkasan_wide(df)
    return [
        {"label": "COUNT tanggal", "file": f, "sheet": "Latihan", "cell": "C5", "expected": N},
        {"label": "COUNT cabang_id (teks) = 0", "file": f, "sheet": "Latihan", "cell": "C6", "expected": 0},
        {"label": "Total penjualan (structured ref)", "file": f, "sheet": "Latihan", "cell": "C18", "expected": int(df["total"].sum())},
        {"label": "Penjualan Surabaya", "file": f, "sheet": "Latihan", "cell": "C20",
         "expected": int(df.loc[df["nama_cabang"] == "Surabaya", "total"].sum())},
        {"label": "Surabaya Jan (SUMIFS tanggal)", "file": f, "sheet": "Latihan", "cell": "C26",
         "expected": int(wide.loc["Surabaya", "2024-01"])},
        {"label": "Selisih total vs Laporan_Lama = 0", "file": f, "sheet": "Latihan", "cell": "I29", "expected": 0},
        {"label": "Selisih long vs wide = 0", "file": f, "sheet": "Latihan", "cell": "D52", "expected": 0},
        {"label": "ID 001 ditemukan di Data", "file": f, "sheet": "ID_Rusak", "cell": "E4",
         "expected": int((df["cabang_id"] == "001").sum())},
    ]


CATATAN_VERIFIKASI = [
    "Di latihan_02.xlsx sheet Data masih range biasa (peserta membuat tabel tblPenjualan sendiri); rumus Bagian C di file solusi memakai structured references dan terverifikasi di LibreOffice.",
    "Sheet Laporan_Lama sengaja berisi anti-pola (merged cell, subtotal di tengah, angka sebagai teks) sebagai bahan diskusi, bukan untuk ditiru.",
]

if __name__ == "__main__":
    main()
