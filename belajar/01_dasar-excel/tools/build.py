"""
Pembangun modul 01 - Dasar Excel.
Jalankan dari folder modul: python tools/build.py
Menghasilkan latihan/data_01.csv, latihan/latihan_01.xlsx, latihan/solusi_01.xlsx.
"""
from __future__ import annotations

import sys
from pathlib import Path

import numpy as np
import pandas as pd

HERE = Path(__file__).resolve().parent
MOD = HERE.parent
sys.path.insert(0, str(MOD.parent / "tools"))
import belajar_lib as L  # noqa: E402

SEED = 101
N = 150
OUT = MOD / "latihan"


def buat_data() -> pd.DataFrame:
    r = L.rng(SEED)
    tgl = L.tanggal_acak(r, N, "2024-01-01", "2024-03-31")
    prod = [L.PRODUK_TOKO[i] for i in r.integers(0, len(L.PRODUK_TOKO), N)]
    qty = r.integers(1, 12, N)
    diskon = r.choice([0, 0, 0, 0.05, 0.1, 0.15], N)
    df = pd.DataFrame({
        "no_transaksi": [f"TRX-{i:04d}" for i in range(1, N + 1)],
        "tanggal": tgl,
        "kasir": r.choice(L.KASIR, N),
        "kode_produk": [p[0] for p in prod],
        "nama_produk": [p[1] for p in prod],
        "kategori": [p[2] for p in prod],
        "qty": qty,
        "harga_satuan": [p[3] for p in prod],
        "diskon_pct": diskon,
    })
    df["total"] = (df["qty"] * df["harga_satuan"] * (1 - df["diskon_pct"])).round(0).astype(int)
    return df


INSTRUKSI = [
    "# Modul 01 - Dasar Excel: workbook latihan",
    "",
    "Workbook ini berisi dua sheet data dan satu sheet Latihan.",
    "- Sheet 'Data': 150 transaksi Toko Sumber Rejeki (Jan-Mar 2024). Kolom tanggal, harga, dan diskon SENGAJA belum diformat.",
    "- Sheet 'Latihan': sel kuning adalah sel yang harus Anda isi.",
    "",
    "## Tugas",
    "1. Di sheet Data: format kolom tanggal menjadi yyyy-mm-dd, harga_satuan dan total menjadi #,##0, diskon_pct menjadi 0%.",
    "2. Di sheet Data: aktifkan Freeze Panes di bawah baris header (View > Freeze Panes > Freeze Top Row).",
    "3. Di sheet Latihan bagian A: isi rumus COUNTA, SUM, MAX, MIN, AVERAGE pada kolom total (Data!J2:J151).",
    "4. Bagian B: pelajari perbedaan nilai dan tampilan (tanggal = angka serial; 0,1 tampil 10%).",
    "5. Bagian C: gunakan Filter di sheet Data, baca jumlah baris pada status bar, dan ketik angkanya.",
    "6. Bagian D: urutkan Data berdasarkan total dari besar ke kecil, salin 3 no_transaksi teratas sebagai nilai (Paste Special > Values).",
    "7. Bagian E: salin tabel ringkasan kategori, lalu tempel dengan Paste Special > Transpose ke area yang disediakan.",
    "",
    "Bandingkan hasil Anda dengan solusi_01.xlsx dan kunci_jawaban.md.",
] + [""] + L.kalimat_latihan()

RINGKASAN_KATEGORI = ["Sembako", "Minuman", "Makanan", "Perawatan", "Rumah Tangga"]


def bangun(df: pd.DataFrame, solusi: bool):
    wb = L.wb_baru()
    L.sheet_teks(wb, "Instruksi", INSTRUKSI)

    # ---- Sheet Data ----
    ws = wb.create_sheet("Data")
    if solusi:
        fmts = {"tanggal": L.FMT_DATE, "harga_satuan": L.FMT_INT, "total": L.FMT_INT, "diskon_pct": "0%"}
        L.tulis_df(ws, df, formats=fmts, font=L.INPUT_FONT, freeze=True)
    else:
        fmts = {"tanggal": "General", "harga_satuan": "General", "total": "General", "diskon_pct": "General"}
        L.tulis_df(ws, df, formats=fmts, font=L.INPUT_FONT, freeze=False)
        ws.column_dimensions["B"].width = 12
    last = len(df) + 1  # baris terakhir data

    # ---- Sheet Latihan ----
    wl = wb.create_sheet("Latihan")
    L.judul(wl, "Modul 01 - Latihan Dasar Excel", "A1",
            "Isi sel kuning. Rentang data: Data!A2:J151 (150 transaksi).")
    wl.column_dimensions["A"].width = 3
    wl.column_dimensions["B"].width = 62
    wl.column_dimensions["C"].width = 18
    wl.column_dimensions["D"].width = 18
    wl.column_dimensions["E"].width = 18
    wl.column_dimensions["F"].width = 18
    wl.column_dimensions["G"].width = 18

    L.subjudul(wl, "B3", "Bagian A - Statistik dasar kolom total (Data!J2:J151)")
    bagianA = [
        ("B4", "Jumlah transaksi (COUNTA pada no_transaksi)", f"=COUNTA(Data!A2:A{last})", L.FMT_INT),
        ("B5", "Total penjualan (SUM kolom total)", f"=SUM(Data!J2:J{last})", L.FMT_INT),
        ("B6", "Transaksi terbesar (MAX)", f"=MAX(Data!J2:J{last})", L.FMT_INT),
        ("B7", "Transaksi terkecil (MIN)", f"=MIN(Data!J2:J{last})", L.FMT_INT),
        ("B8", "Rata-rata nilai transaksi (AVERAGE)", f"=AVERAGE(Data!J2:J{last})", L.FMT_DEC2),
        ("B9", "Total qty terjual (SUM kolom qty)", f"=SUM(Data!G2:G{last})", L.FMT_INT),
    ]
    for cell, teks, rumus, fmt in bagianA:
        L.label(wl, cell, teks, bold=False)
        target = "C" + cell[1:]
        L.sel_jawab(wl, target, rumus if solusi else None, fmt)

    L.subjudul(wl, "B11", "Bagian B - Nilai vs tampilan sel")
    L.label(wl, "B12", "Tanggal transaksi pertama (=Data!B2), format yyyy-mm-dd", bold=False)
    L.sel_jawab(wl, "C12", "=Data!B2" if solusi else None, L.FMT_DATE, font=L.LINK_FONT)
    L.label(wl, "B13", "Nilai yang sama (=C12*1) dengan format General -> tampil angka serial", bold=False)
    L.sel_jawab(wl, "C13", "=C12*1" if solusi else None, "General")
    L.label(wl, "B14", "Ketik 0,1 lalu format 0% -> tampil 10%", bold=False)
    L.sel_jawab(wl, "C14", 0.1 if solusi else None, "0%", font=L.INPUT_FONT)
    L.label(wl, "B15", "Nilai di C14 dikali 100 (bukti nilai sebenarnya 0,1)", bold=False)
    L.sel_jawab(wl, "C15", "=C14*100" if solusi else None, L.FMT_DEC2)
    L.label(wl, "B16", "Tanggal transaksi terakhir (=Data!B151)", bold=False)
    L.sel_jawab(wl, "C16", f"=Data!B{last}" if solusi else None, L.FMT_DATE, font=L.LINK_FONT)
    L.label(wl, "B17", "Selisih hari antara tanggal terakhir dan pertama (=C16-C12)", bold=False)
    L.sel_jawab(wl, "C17", "=C16-C12" if solusi else None, L.FMT_INT)
    L.label(wl, "B18", "Bulan dari tanggal pertama (=MONTH(C12))", bold=False)
    L.sel_jawab(wl, "C18", "=MONTH(C12)" if solusi else None, "General")

    L.subjudul(wl, "B20", "Bagian C - Hasil Filter (ketik angka dari status bar; pembanding memakai COUNTIFS/SUMIFS)")
    rk = f"Data!$F$2:$F${last}"
    rt = f"Data!$J$2:$J${last}"
    rkasir = f"Data!$C$2:$C${last}"
    rq = f"Data!$G$2:$G${last}"
    bagianC = [
        ("B21", "Jumlah transaksi kategori Sembako", f'=COUNTIFS({rk},"Sembako")', L.FMT_INT),
        ("B22", "Jumlah transaksi kasir Rina dengan qty >= 5", f'=COUNTIFS({rkasir},"Rina",{rq},">=5")', L.FMT_INT),
        ("B23", "Total penjualan kategori Minuman (SUM pada hasil filter)", f'=SUMIFS({rt},{rk},"Minuman")', L.FMT_INT),
        ("B24", "Jumlah transaksi dengan diskon_pct > 0", f'=COUNTIFS(Data!$I$2:$I${last},">0")', L.FMT_INT),
    ]
    for cell, teks, rumus, fmt in bagianC:
        L.label(wl, cell, teks, bold=False)
        L.sel_jawab(wl, "C" + cell[1:], rumus if solusi else None, fmt)
    wl["D20"] = "Catatan: rumus COUNTIFS/SUMIFS dipelajari di modul 04; di sini hanya sebagai pembanding hasil filter."
    wl["D20"].font = L.Font(name=L.FONT_NAME, size=9, italic=True)

    L.subjudul(wl, "B26", "Bagian D - Sort: 3 transaksi dengan total terbesar")
    L.tulis_header(wl, 27, 2, ["Peringkat", "no_transaksi (tempel sebagai nilai)", "total (rumus LARGE)"])
    top3 = df.sort_values(["total", "no_transaksi"], ascending=[False, True]).head(3)
    for i, (_, rec) in enumerate(top3.iterrows(), start=1):
        wl.cell(row=27 + i, column=2, value=i).border = L.BORDER
        L.sel_jawab(wl, f"C{27 + i}", rec["no_transaksi"] if solusi else None, font=L.INPUT_FONT)
        L.sel_jawab(wl, f"D{27 + i}", f"=LARGE({rt},{i})" if solusi else None, L.FMT_INT)

    L.subjudul(wl, "B32", "Bagian E - Paste Special > Transpose")
    L.label(wl, "B33", "Tabel ringkasan (vertikal) - salin B34:C38 lalu tempel Transpose mulai sel E34", bold=False)
    L.tulis_header(wl, 34, 2, ["kategori", "jumlah_transaksi"])
    for i, k in enumerate(RINGKASAN_KATEGORI, start=1):
        c = wl.cell(row=34 + i, column=2, value=k)
        c.border = L.BORDER
        L.sel_rumus(wl, f"C{34 + i}", f'=COUNTIFS({rk},B{34 + i})', L.FMT_INT)
    L.label(wl, "E33", "Hasil transpose (nilai statis, bukan rumus)", bold=False)
    if solusi:
        wl.cell(row=34, column=5, value="kategori").font = L.BOLD_FONT
        wl.cell(row=35, column=5, value="jumlah_transaksi").font = L.BOLD_FONT
        counts = df["kategori"].value_counts()
        for j, k in enumerate(RINGKASAN_KATEGORI):
            c1 = wl.cell(row=34, column=6 + j, value=k)
            c2 = wl.cell(row=35, column=6 + j, value=int(counts.get(k, 0)))
            c1.fill = L.ANSWER_FILL
            c2.fill = L.ANSWER_FILL
            c1.border = L.BORDER
            c2.border = L.BORDER
            c2.number_format = L.FMT_INT
    else:
        for rr in (34, 35):
            for cc in range(5, 11):
                wl.cell(row=rr, column=cc).fill = L.ANSWER_FILL
                wl.cell(row=rr, column=cc).border = L.BORDER

    L.legenda(wl, "B41")
    wl.sheet_view.showGridLines = False
    wb.active = wb.sheetnames.index("Instruksi")
    return wb


def main() -> None:
    df = buat_data()
    L.simpan_csv(df, OUT / "data_01.csv")
    L.simpan(bangun(df, solusi=False), OUT / "latihan_01.xlsx")
    L.simpan(bangun(df, solusi=True), OUT / "solusi_01.xlsx")
    print("selesai:", OUT)


def checks() -> list[dict]:
    df = buat_data()
    f = "solusi_01.xlsx"
    return [
        {"label": "Jumlah transaksi", "file": f, "sheet": "Latihan", "cell": "C4", "expected": int(len(df))},
        {"label": "Total penjualan", "file": f, "sheet": "Latihan", "cell": "C5", "expected": int(df["total"].sum())},
        {"label": "Transaksi terbesar", "file": f, "sheet": "Latihan", "cell": "C6", "expected": int(df["total"].max())},
        {"label": "Rata-rata transaksi", "file": f, "sheet": "Latihan", "cell": "C8", "expected": float(df["total"].mean())},
        {"label": "Transaksi Sembako", "file": f, "sheet": "Latihan", "cell": "C21",
         "expected": int((df["kategori"] == "Sembako").sum())},
        {"label": "Kasir Rina qty>=5", "file": f, "sheet": "Latihan", "cell": "C22",
         "expected": int(((df["kasir"] == "Rina") & (df["qty"] >= 5)).sum())},
        {"label": "Total Minuman", "file": f, "sheet": "Latihan", "cell": "C23",
         "expected": int(df.loc[df["kategori"] == "Minuman", "total"].sum())},
        {"label": "Selisih hari", "file": f, "sheet": "Latihan", "cell": "C17",
         "expected": int((pd.Timestamp(df["tanggal"].iloc[-1]) - pd.Timestamp(df["tanggal"].iloc[0])).days)},
    ]


CATATAN_VERIFIKASI = [
    "latihan_01.xlsx sengaja berisi sedikit rumus (hanya pembanding Bagian E); sel kuning kosong diisi peserta.",
    "Sheet Data pada latihan_01.xlsx sengaja belum diformat (tanggal tampil sebagai angka serial) sebagai bahan latihan format.",
]

if __name__ == "__main__":
    main()
