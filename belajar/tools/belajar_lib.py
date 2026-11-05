"""
Pustaka bersama untuk skrip pembangun modul belajar (belajar/NN_slug/tools/build.py).

Semua workbook dibangun dengan openpyxl secara deterministik (seed tetap).
Konvensi warna sel:
  - input / data mentah yang boleh diubah : font biru (0000FF)
  - formula                                : font hitam
  - link antar-sheet                        : font hijau (008000)
  - sel yang harus diisi peserta (latihan)  : fill kuning (FFF2CC)
  - header                                  : fill biru tua (1F4E78), font putih tebal
"""
from __future__ import annotations

import datetime as dt
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

FONT_NAME = "Arial"
COLOR_HEADER = "1F4E78"
COLOR_INPUT = "0000FF"
COLOR_LINK = "008000"
COLOR_FILL_ANSWER = "FFF2CC"
COLOR_FILL_SUBHEAD = "D9E1F2"
COLOR_FILL_NOTE = "F2F2F2"

FMT_INT = "#,##0"
FMT_DEC1 = "#,##0.0"
FMT_DEC2 = "#,##0.00"
FMT_PCT = "0.0%"
FMT_PCT2 = "0.00%"
FMT_DATE = "yyyy-mm-dd"
FMT_RP = '"Rp"#,##0'

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

HEADER_FONT = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=10)
HEADER_FILL = PatternFill("solid", fgColor=COLOR_HEADER)
BODY_FONT = Font(name=FONT_NAME, size=10)
INPUT_FONT = Font(name=FONT_NAME, size=10, color=COLOR_INPUT)
LINK_FONT = Font(name=FONT_NAME, size=10, color=COLOR_LINK)
BOLD_FONT = Font(name=FONT_NAME, size=10, bold=True)
TITLE_FONT = Font(name=FONT_NAME, size=14, bold=True, color=COLOR_HEADER)
SUB_FONT = Font(name=FONT_NAME, size=11, bold=True, color=COLOR_HEADER)
ANSWER_FILL = PatternFill("solid", fgColor=COLOR_FILL_ANSWER)
SUBHEAD_FILL = PatternFill("solid", fgColor=COLOR_FILL_SUBHEAD)
NOTE_FILL = PatternFill("solid", fgColor=COLOR_FILL_NOTE)

TANGGAL_PEMBUATAN = "2025-11-05"

# ----------------------------------------------------------------------------
# Data sintetis: nama orang, kota, produk (tema bisnis Indonesia)
# ----------------------------------------------------------------------------
NAMA_DEPAN = [
    "Andi", "Budi", "Citra", "Dewi", "Eko", "Fajar", "Gita", "Hadi", "Indah", "Joko",
    "Kartika", "Lestari", "Maya", "Nugroho", "Putri", "Rizky", "Sari", "Tono", "Utami",
    "Wahyu", "Yuni", "Zaki", "Agus", "Bayu", "Dian", "Fitri", "Galih", "Hana", "Irfan",
    "Kurnia", "Lina", "Mulyadi", "Nanda", "Rahma", "Surya", "Tri", "Vina", "Widodo",
    "Yoga", "Ayu", "Dimas", "Intan", "Reza", "Sinta", "Teguh",
]
NAMA_BELAKANG = [
    "Pratama", "Saputra", "Wijaya", "Santoso", "Hidayat", "Kusuma", "Lestari", "Nugraha",
    "Permata", "Rahayu", "Setiawan", "Susanto", "Utomo", "Wibowo", "Yulianti", "Hartono",
    "Firmansyah", "Maulana", "Ramadhan", "Anggraini", "Puspita", "Handayani", "Siregar",
    "Nasution", "Simanjuntak", "Gunawan", "Halim", "Salim", "Prasetyo", "Mahendra",
]
KOTA = ["Surabaya", "Jakarta", "Bandung", "Semarang", "Yogyakarta", "Malang", "Medan",
        "Makassar", "Denpasar", "Palembang", "Balikpapan", "Sidoarjo"]
PROVINSI = {
    "Surabaya": "Jawa Timur", "Malang": "Jawa Timur", "Sidoarjo": "Jawa Timur",
    "Jakarta": "DKI Jakarta", "Bandung": "Jawa Barat", "Semarang": "Jawa Tengah",
    "Yogyakarta": "DI Yogyakarta", "Medan": "Sumatera Utara", "Makassar": "Sulawesi Selatan",
    "Denpasar": "Bali", "Palembang": "Sumatera Selatan", "Balikpapan": "Kalimantan Timur",
}
# (kode, nama produk, kategori, harga satuan, hpp satuan)
PRODUK_TOKO = [
    ("P001", "Beras Premium 5 kg", "Sembako", 72000, 64000),
    ("P002", "Minyak Goreng 2 L", "Sembako", 36000, 32500),
    ("P003", "Gula Pasir 1 kg", "Sembako", 17500, 15800),
    ("P004", "Telur Ayam 1 kg", "Sembako", 29000, 26500),
    ("P005", "Tepung Terigu 1 kg", "Sembako", 13000, 11200),
    ("P006", "Kopi Bubuk 200 g", "Minuman", 24000, 19000),
    ("P007", "Teh Celup 25 pcs", "Minuman", 9500, 7200),
    ("P008", "Susu UHT 1 L", "Minuman", 18500, 15900),
    ("P009", "Air Mineral 600 ml (dus)", "Minuman", 42000, 36000),
    ("P010", "Sirup Rasa Jeruk 500 ml", "Minuman", 21000, 16500),
    ("P011", "Mi Instan Goreng (dus)", "Makanan", 112000, 101000),
    ("P012", "Biskuit Kaleng 400 g", "Makanan", 38000, 29500),
    ("P013", "Keripik Singkong 250 g", "Makanan", 15000, 10500),
    ("P014", "Roti Tawar", "Makanan", 14000, 10800),
    ("P015", "Kecap Manis 600 ml", "Makanan", 23000, 19200),
    ("P016", "Sabun Mandi 90 g", "Perawatan", 4500, 3600),
    ("P017", "Sampo 170 ml", "Perawatan", 27000, 21500),
    ("P018", "Pasta Gigi 190 g", "Perawatan", 16500, 13000),
    ("P019", "Deterjen Bubuk 800 g", "Rumah Tangga", 24500, 20500),
    ("P020", "Pembersih Lantai 800 ml", "Rumah Tangga", 19000, 15200),
    ("P021", "Tisu Gulung 10 roll", "Rumah Tangga", 32000, 26500),
    ("P022", "Kantong Sampah 30 pcs", "Rumah Tangga", 12000, 8800),
    ("P023", "Baterai AA 4 pcs", "Rumah Tangga", 21000, 15000),
    ("P024", "Lampu LED 9 W", "Rumah Tangga", 28000, 20000),
]
DEPARTEMEN = ["Keuangan", "Pemasaran", "Operasional", "SDM", "TI", "Penjualan", "Gudang"]
JABATAN = ["Staf", "Staf Senior", "Supervisor", "Manajer"]
KASIR = ["Rina", "Dedi", "Sinta", "Agus"]


def rng(seed: int) -> np.random.Generator:
    return np.random.default_rng(seed)


def nama_acak(r: np.random.Generator) -> str:
    return f"{r.choice(NAMA_DEPAN)} {r.choice(NAMA_BELAKANG)}"


def nama_unik(r: np.random.Generator, n: int) -> list[str]:
    out, seen = [], set()
    while len(out) < n:
        nm = nama_acak(r)
        if nm not in seen:
            seen.add(nm)
            out.append(nm)
    return out


def tanggal_acak(r: np.random.Generator, n: int, mulai="2024-01-01", akhir="2024-12-31",
                 bobot_bulan: list[float] | None = None) -> list[dt.date]:
    d0 = pd.Timestamp(mulai)
    d1 = pd.Timestamp(akhir)
    hari = pd.date_range(d0, d1, freq="D")
    if bobot_bulan is None:
        idx = r.integers(0, len(hari), n)
    else:
        w = np.array([bobot_bulan[h.month - 1] for h in hari], dtype=float)
        w = w / w.sum()
        idx = r.choice(len(hari), size=n, p=w)
    return sorted(hari[i].date() for i in idx)


# ----------------------------------------------------------------------------
# Helper workbook
# ----------------------------------------------------------------------------

def wb_baru() -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)
    return wb


def simpan(wb: Workbook, path: str | Path) -> Path:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def autofit(ws, min_width: float = 8, max_width: float = 60, pad: float = 2) -> None:
    widths: dict[int, float] = {}
    for row in ws.iter_rows():
        for c in row:
            if c.value is None:
                continue
            v = c.value
            if isinstance(v, (dt.date, dt.datetime)):
                ln = 10
            elif isinstance(v, float):
                ln = len(f"{v:,.2f}")
            elif isinstance(v, str) and v.startswith("="):
                ln = 12
            else:
                ln = len(str(v))
            widths[c.column] = max(widths.get(c.column, 0), ln)
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = max(min_width, min(max_width, w + pad))


def tulis_header(ws, row: int, col: int, headers: list[str]) -> None:
    for j, h in enumerate(headers):
        c = ws.cell(row=row, column=col + j, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER


def tulis_df(ws, df: pd.DataFrame, row: int = 1, col: int = 1, formats: dict | None = None,
             font: Font = BODY_FONT, border: bool = True, freeze: bool = True,
             as_table: str | None = None, table_style: str = "TableStyleMedium2") -> int:
    """Tulis DataFrame mulai (row, col). Mengembalikan nomor baris terakhir data."""
    formats = formats or {}
    headers = [str(h) for h in df.columns]
    tulis_header(ws, row, col, headers)
    for i, rec in enumerate(df.itertuples(index=False), start=1):
        for j, v in enumerate(rec):
            if isinstance(v, (np.integer,)):
                v = int(v)
            elif isinstance(v, (np.floating,)):
                v = float(v)
            elif isinstance(v, pd.Timestamp):
                v = v.date()
            elif v is pd.NaT or (isinstance(v, float) and np.isnan(v)):
                v = None
            c = ws.cell(row=row + i, column=col + j, value=v)
            c.font = font
            if border:
                c.border = BORDER
            fmt = formats.get(headers[j])
            if fmt:
                c.number_format = fmt
            elif isinstance(v, (dt.date, dt.datetime)):
                c.number_format = FMT_DATE
    last = row + len(df)
    if freeze:
        ws.freeze_panes = ws.cell(row=row + 1, column=col)
    if as_table:
        ref = f"{get_column_letter(col)}{row}:{get_column_letter(col + len(headers) - 1)}{max(last, row + 1)}"
        t = Table(displayName=as_table, ref=ref)
        t.tableStyleInfo = TableStyleInfo(name=table_style, showRowStripes=True)
        ws.add_table(t)
    autofit(ws)
    return last


def judul(ws, text: str, cell: str = "A1", sub: str | None = None) -> None:
    ws[cell] = text
    ws[cell].font = TITLE_FONT
    if sub:
        c = ws.cell(row=ws[cell].row + 1, column=ws[cell].column, value=sub)
        c.font = Font(name=FONT_NAME, size=10, italic=True, color="595959")


def label(ws, cell: str, text: str, bold: bool = True) -> None:
    ws[cell] = text
    ws[cell].font = BOLD_FONT if bold else BODY_FONT


def sel_jawab(ws, cell: str, value=None, fmt: str | None = None, font: Font = BODY_FONT) -> None:
    """Sel yang harus diisi peserta (kuning). Pada solusi, isi value (rumus)."""
    c = ws[cell]
    c.value = value
    c.fill = ANSWER_FILL
    c.border = BORDER
    c.font = font
    if fmt:
        c.number_format = fmt


def sel_rumus(ws, cell: str, formula: str, fmt: str | None = None, link: bool = False,
              bold: bool = False) -> None:
    c = ws[cell]
    c.value = formula
    c.font = Font(name=FONT_NAME, size=10, color=COLOR_LINK if link else "000000", bold=bold)
    c.border = BORDER
    if fmt:
        c.number_format = fmt


def sel_input(ws, cell: str, value, fmt: str | None = None) -> None:
    c = ws[cell]
    c.value = value
    c.font = INPUT_FONT
    c.border = BORDER
    if fmt:
        c.number_format = fmt


def subjudul(ws, cell: str, text: str, span: int = 6) -> None:
    ws[cell] = text
    ws[cell].font = SUB_FONT
    r = ws[cell].row
    c0 = ws[cell].column
    for j in range(span):
        ws.cell(row=r, column=c0 + j).fill = SUBHEAD_FILL


def sheet_teks(wb: Workbook, nama: str, baris: list[str], lebar: int = 110, judul_sheet: str | None = None):
    """Sheet berisi teks (Instruksi / Readme). Baris yang diawali '# ' menjadi judul,
    '## ' menjadi subjudul, '' baris kosong."""
    ws = wb.create_sheet(nama)
    ws.column_dimensions["A"].width = lebar
    r = 1
    if judul_sheet:
        ws.cell(row=r, column=1, value=judul_sheet).font = TITLE_FONT
        r += 2
    for b in baris:
        c = ws.cell(row=r, column=1)
        if b.startswith("# "):
            c.value = b[2:]
            c.font = TITLE_FONT
        elif b.startswith("## "):
            c.value = b[3:]
            c.font = SUB_FONT
            c.fill = SUBHEAD_FILL
        else:
            c.value = b
            c.font = BODY_FONT
            c.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1
    ws.sheet_view.showGridLines = False
    return ws


def legenda(ws, cell: str) -> None:
    """Legenda warna sel standar."""
    r = ws[cell].row
    c = ws[cell].column
    items = [
        ("Legenda warna:", BOLD_FONT, None),
        ("Font biru = input / data mentah yang boleh diubah", INPUT_FONT, None),
        ("Font hitam = rumus (jangan ditimpa angka)", BODY_FONT, None),
        ("Font hijau = tautan ke sheet lain", LINK_FONT, None),
        ("Fill kuning = sel yang harus Anda isi pada workbook latihan", BODY_FONT, ANSWER_FILL),
    ]
    for i, (t, f, fill) in enumerate(items):
        cc = ws.cell(row=r + i, column=c, value=t)
        cc.font = f
        if fill:
            cc.fill = fill


def col_ref(df: pd.DataFrame, nama_kolom: str, col0: int = 1) -> str:
    """Huruf kolom Excel untuk nama kolom DataFrame (kolom pertama DF di col0)."""
    return get_column_letter(col0 + list(df.columns).index(nama_kolom))


def rentang(df: pd.DataFrame, nama_kolom: str, row0: int = 1, col0: int = 1, sheet: str | None = None,
            absolut: bool = True) -> str:
    """Rentang data (tanpa header) untuk kolom DF, misal Data!$C$2:$C$121."""
    L = col_ref(df, nama_kolom, col0)
    a, b = row0 + 1, row0 + len(df)
    s = f"${L}${a}:${L}${b}" if absolut else f"{L}{a}:{L}{b}"
    return f"{sheet}!{s}" if sheet else s


def simpan_csv(df: pd.DataFrame, path: str | Path) -> Path:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(path, index=False, lineterminator="\n")
    return path


def kalimat_latihan() -> list[str]:
    return [
        "Materi ini adalah LATIHAN dengan data sintetis (dibangkitkan skrip, bukan data perusahaan nyata).",
        "Pemisah argumen rumus di dalam file ini mengikuti format internal Excel (koma). Pada Excel regional",
        "Indonesia, rumus yang Anda ketik di layar memakai titik koma (;), misalnya =SUM(A1;A5).",
    ]
