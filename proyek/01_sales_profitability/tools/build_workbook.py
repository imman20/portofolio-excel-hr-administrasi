"""
Pembangun workbook proyek "Sales & Profitability Analysis".

Menghasilkan dua file di folder workbook/:
  - solusi_sales_profitability.xlsx  : rumus lengkap (Readme, Raw_Data, Retur, Master,
                                       Calculation, Report, Checks)
  - latihan_sales_profitability.xlsx : data mentah + master + sheet Instruksi; kolom
                                       perhitungan kosong berlabel (baris 2 diisi contoh)

Semua perhitungan memakai rumus hidup (SUMIFS, INDEX/MATCH, SMALL, SUMPRODUCT, TEXT,
IFERROR). Tidak ada XLOOKUP/FILTER/array dinamis supaya file dapat diverifikasi di
LibreOffice dan dibuka di Excel 2016+.

Cara menjalankan (dari folder proyek):
    python tools/generate_data.py      # bila CSV belum ada
    python tools/build_workbook.py
Setelah itu lakukan recalc (LibreOffice headless) dan verifikasi: python tools/verify.py
"""

from __future__ import annotations

import datetime as dt
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.workbook.properties import CalcProperties
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

PROJECT_DIR = Path(__file__).resolve().parent.parent
RAW_DIR = PROJECT_DIR / "data" / "raw"
WB_DIR = PROJECT_DIR / "workbook"
TANGGAL_BUAT = dt.date(2026, 11, 5)

# ---------------------------------------------------------------------------
# Gaya
# ---------------------------------------------------------------------------
FONT_NAME = "Arial"
F_BASE = Font(name=FONT_NAME, size=10)
F_BOLD = Font(name=FONT_NAME, size=10, bold=True)
F_TITLE = Font(name=FONT_NAME, size=14, bold=True, color="1F4E78")
F_SUB = Font(name=FONT_NAME, size=10, italic=True, color="595959")
F_HEADER = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
F_INPUT = Font(name=FONT_NAME, size=10, color="0000FF")          # input / data mentah
F_FORMULA = Font(name=FONT_NAME, size=10, color="000000")        # rumus dalam sheet
F_LINK = Font(name=FONT_NAME, size=10, color="008000")           # rumus merujuk sheet lain
F_LINK_BOLD = Font(name=FONT_NAME, size=10, color="008000", bold=True)
F_KPI = Font(name=FONT_NAME, size=16, bold=True, color="008000")
F_KPI_LABEL = Font(name=FONT_NAME, size=9, bold=True, color="595959")

FILL_HEADER = PatternFill("solid", fgColor="1F4E78")
FILL_BLOCK = PatternFill("solid", fgColor="DDEBF7")
FILL_TOTAL = PatternFill("solid", fgColor="F2F2F2")
FILL_CARD = PatternFill("solid", fgColor="F7FAFD")
FILL_INPUT = PatternFill("solid", fgColor="FFF2CC")
FILL_TOP = PatternFill("solid", fgColor="E2EFDA")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

FMT_RP = "#,##0"
FMT_PCT = "0.0%"
FMT_DATE = "dd/mm/yyyy"
FMT_INT = "#,##0"

N_TRANS = 6000  # diperbarui dari CSV saat build
N_RETUR = 120
N_PRODUK = 40
N_CUST = 120
N_TARGET = 72


def style_header(ws, row, col_start, col_end, fill=FILL_HEADER, font=F_HEADER):
    for c in range(col_start, col_end + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


def set_widths(ws, widths: dict):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def put(ws, ref, value, font=F_BASE, fmt=None, fill=None, align=None, border=False):
    cell = ws[ref]
    cell.value = value
    cell.font = font
    if fmt:
        cell.number_format = fmt
    if fill:
        cell.fill = fill
    if align:
        cell.alignment = align
    if border:
        cell.border = BORDER
    return cell


def write_table(ws, df: pd.DataFrame, row0: int, col0: int, date_cols=(), fmt_map=None,
                font=F_INPUT):
    """Tulis DataFrame sebagai tabel: header di row0, data mulai row0+1."""
    fmt_map = fmt_map or {}
    for j, name in enumerate(df.columns):
        ws.cell(row=row0, column=col0 + j, value=name)
    style_header(ws, row0, col0, col0 + len(df.columns) - 1)
    for i, rec in enumerate(df.itertuples(index=False), start=1):
        for j, name in enumerate(df.columns):
            v = rec[j]
            if name in date_cols:
                v = dt.date.fromisoformat(v)
            elif hasattr(v, "item"):
                v = v.item()
            cell = ws.cell(row=row0 + i, column=col0 + j, value=v)
            cell.font = font
            if name in date_cols:
                cell.number_format = FMT_DATE
            elif name in fmt_map:
                cell.number_format = fmt_map[name]


# ---------------------------------------------------------------------------
# Data
# ---------------------------------------------------------------------------
def load_data():
    d = {
        "trans": pd.read_csv(RAW_DIR / "transaksi.csv"),
        "retur": pd.read_csv(RAW_DIR / "retur.csv"),
        "produk": pd.read_csv(RAW_DIR / "master_produk.csv"),
        "pelanggan": pd.read_csv(RAW_DIR / "master_pelanggan.csv"),
        "wilayah": pd.read_csv(RAW_DIR / "master_wilayah.csv"),
        "sales": pd.read_csv(RAW_DIR / "master_sales.csv"),
        "target": pd.read_csv(RAW_DIR / "target.csv"),
    }
    return d


# ---------------------------------------------------------------------------
# Sheet Raw_Data
# ---------------------------------------------------------------------------
RAW_CALC_COLS = [
    # (header, rumus baris 2 (dengan {r}), format, font)
    ("omzet_bruto", "=G{r}*H{r}", FMT_RP, F_FORMULA),
    ("nilai_diskon", "=J{r}*I{r}", FMT_RP, F_FORMULA),
    ("omzet_neto", "=J{r}-K{r}", FMT_RP, F_FORMULA),
    ("qty_retur", "=SUMIFS(Retur!$D$2:$D${nr},Retur!$B$2:$B${nr},A{r})", FMT_INT, F_LINK),
    ("nilai_retur", "=M{r}*H{r}*(1-I{r})", FMT_RP, F_FORMULA),
    ("penjualan_bersih", "=L{r}-N{r}", FMT_RP, F_FORMULA),
    ("hpp_satuan", "=INDEX(Master!$F$3:$F${np},MATCH(D{r},Master!$A$3:$A${np},0))", FMT_RP, F_LINK),
    ("hpp", "=(G{r}-M{r})*P{r}", FMT_RP, F_FORMULA),
    ("laba_kotor", "=O{r}-Q{r}", FMT_RP, F_FORMULA),
    ("margin_kotor", "=IF(O{r}=0,0,R{r}/O{r})", FMT_PCT, F_FORMULA),
    ("periode", '=TEXT(B{r},"yyyy-mm")', None, F_FORMULA),
    ("kategori", "=INDEX(Master!$C$3:$C${np},MATCH(D{r},Master!$A$3:$A${np},0))", None, F_LINK),
    ("segmen", "=INDEX(Master!$J$3:$J${nc},MATCH(C{r},Master!$H$3:$H${nc},0))", None, F_LINK),
    ("kunci_unik", '=A{r}&"|"&D{r}', None, F_FORMULA),
    ("duplikat", "=IF(COUNTIFS($W$2:$W${nt},W{r})>1,1,0)", FMT_INT, F_FORMULA),
]


def build_raw_data(wb, d, solusi: bool):
    ws = wb.create_sheet("Raw_Data")
    trans = d["trans"]
    write_table(ws, trans, 1, 1, date_cols=("tanggal",),
                fmt_map={"harga_satuan": FMT_RP, "qty": FMT_INT, "diskon_pct": "0%"})
    nt = len(trans) + 1
    np_, nc, nr = N_PRODUK + 2, N_CUST + 2, N_RETUR + 1
    col0 = len(trans.columns) + 1  # kolom J
    for j, (name, _, _, _) in enumerate(RAW_CALC_COLS):
        ws.cell(row=1, column=col0 + j, value=name)
    style_header(ws, 1, col0, col0 + len(RAW_CALC_COLS) - 1, fill=PatternFill("solid", fgColor="375623"))
    last_row = nt if solusi else 2  # latihan: hanya baris 2 sebagai contoh
    for r in range(2, last_row + 1):
        for j, (name, formula, fmt, font) in enumerate(RAW_CALC_COLS):
            cell = ws.cell(row=r, column=col0 + j,
                           value=formula.format(r=r, nt=nt, np=np_, nc=nc, nr=nr))
            cell.font = font
            if fmt:
                cell.number_format = fmt
    if not solusi:
        for r in range(3, nt + 1):
            for j in range(len(RAW_CALC_COLS)):
                ws.cell(row=r, column=col0 + j).fill = FILL_INPUT
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:{get_column_letter(col0 + len(RAW_CALC_COLS) - 1)}{nt}"
    set_widths(ws, {"A": 16, "B": 12, "C": 12, "D": 11, "E": 10, "F": 14, "G": 7, "H": 13,
                    "I": 10, "J": 15, "K": 13, "L": 15, "M": 9, "N": 13, "O": 16, "P": 12,
                    "Q": 15, "R": 15, "S": 11, "T": 10, "U": 22, "V": 10, "W": 22, "X": 9})
    return ws


def build_retur(wb, d, solusi: bool):
    ws = wb.create_sheet("Retur")
    write_table(ws, d["retur"], 1, 1, date_cols=("tanggal_retur",), fmt_map={"qty_retur": FMT_INT})
    nr = len(d["retur"]) + 1
    nt = len(d["trans"]) + 1
    ws.cell(row=1, column=6, value="order_ditemukan")
    style_header(ws, 1, 6, 6, fill=PatternFill("solid", fgColor="375623"))
    for r in range(2, nr + 1):
        if solusi or r == 2:
            c = ws.cell(row=r, column=6,
                        value=f'=IF(COUNTIFS(Raw_Data!$A$2:$A${nt},B{r})>0,"Ya","Tidak")')
            c.font = F_LINK
        else:
            ws.cell(row=r, column=6).fill = FILL_INPUT
    ws.freeze_panes = "A2"
    set_widths(ws, {"A": 14, "B": 16, "C": 13, "D": 10, "E": 24, "F": 16})
    return ws


def build_master(wb, d):
    ws = wb.create_sheet("Master")
    blocks = [
        ("Master Produk", d["produk"], 1, (), {"harga_list": FMT_RP, "hpp_satuan": FMT_RP}),
        ("Master Pelanggan", d["pelanggan"], 8, ("tanggal_bergabung",), {}),
        ("Master Wilayah", d["wilayah"], 15, (), {}),
        ("Master Sales", d["sales"], 19, (), {}),
        ("Target Penjualan Bersih", d["target"], 23, (), {"target_penjualan_bersih": FMT_RP}),
    ]
    for title, df, col0, date_cols, fmt_map in blocks:
        c = ws.cell(row=1, column=col0, value=title)
        c.font = F_BOLD
        c.fill = FILL_BLOCK
        for j in range(1, len(df.columns)):
            ws.cell(row=1, column=col0 + j).fill = FILL_BLOCK
        write_table(ws, df, 2, col0, date_cols=date_cols, fmt_map=fmt_map)
    ws.freeze_panes = "A3"
    set_widths(ws, {"A": 11, "B": 32, "C": 24, "D": 14, "E": 12, "F": 12, "G": 3,
                    "H": 12, "I": 36, "J": 10, "K": 16, "L": 10, "M": 17, "N": 3,
                    "O": 10, "P": 20, "Q": 12, "R": 3, "S": 14, "T": 18, "U": 10, "V": 3,
                    "W": 10, "X": 10, "Y": 22})
    return ws


# ---------------------------------------------------------------------------
# Sheet Calculation
# ---------------------------------------------------------------------------
def block_title(ws, ref, text):
    c = ws[ref]
    c.value = text
    c.font = F_BOLD
    c.fill = FILL_BLOCK


def build_calculation(wb, d, solusi: bool):
    ws = wb.create_sheet("Calculation")
    nt = len(d["trans"]) + 1
    R = lambda col: f"Raw_Data!${col}$2:${col}${nt}"  # noqa: E731
    T_BULAN = f"Master!$W$3:$W${N_TARGET + 2}"
    T_REG = f"Master!$X$3:$X${N_TARGET + 2}"
    T_VAL = f"Master!$Y$3:$Y${N_TARGET + 2}"

    put(ws, "A1", "Calculation: rekap KPI (semua sel berumus, sumber Raw_Data dan Master)", F_TITLE)
    put(ws, "A2", "Satuan rupiah kecuali disebut lain. Margin = laba kotor / penjualan bersih. "
                  "Pencapaian = penjualan bersih / target. MoM = (bulan ini - bulan lalu) / bulan lalu.", F_SUB)

    def f(ws_ref, value, font=F_LINK, fmt=None):
        put(ws, ws_ref, value if solusi else None, font, fmt, border=True)
        if not solusi:
            ws[ws_ref].fill = FILL_INPUT

    # ---- Blok A: bulanan (A4:I17)
    block_title(ws, "A3", "A. Rekap bulanan")
    hdr = ["periode", "omzet_bruto", "penjualan_bersih", "hpp", "laba_kotor", "margin_kotor",
           "target", "pencapaian_target", "pertumbuhan_mom"]
    for j, h in enumerate(hdr, start=1):
        ws.cell(row=4, column=j, value=h)
    style_header(ws, 4, 1, 9)
    for i, m in enumerate(range(1, 13)):
        r = 5 + i
        put(ws, f"A{r}", f"2024-{m:02d}", F_INPUT, border=True)
        f(f"B{r}", f"=SUMIFS({R('J')},{R('T')},$A{r})", fmt=FMT_RP)
        f(f"C{r}", f"=SUMIFS({R('O')},{R('T')},$A{r})", fmt=FMT_RP)
        f(f"D{r}", f"=SUMIFS({R('Q')},{R('T')},$A{r})", fmt=FMT_RP)
        f(f"E{r}", f"=C{r}-D{r}", F_FORMULA, FMT_RP)
        f(f"F{r}", f"=IF(C{r}=0,0,E{r}/C{r})", F_FORMULA, FMT_PCT)
        f(f"G{r}", f"=SUMIFS({T_VAL},{T_BULAN},$A{r})", fmt=FMT_RP)
        f(f"H{r}", f'=IF(G{r}=0,"-",C{r}/G{r})', F_FORMULA, FMT_PCT)
        if i == 0:
            put(ws, f"I{r}", "-", F_FORMULA, border=True, align=Alignment(horizontal="right"))
        else:
            f(f"I{r}", f'=IF(C{r - 1}=0,"-",(C{r}-C{r - 1})/C{r - 1})', F_FORMULA, FMT_PCT)
    put(ws, "A17", "TOTAL 2024", F_BOLD, fill=FILL_TOTAL, border=True)
    for col in "BCDEG":
        f(f"{col}17", f"=SUM({col}5:{col}16)", F_BOLD, FMT_RP)
        ws[f"{col}17"].fill = FILL_TOTAL
    f("F17", "=IF(C17=0,0,E17/C17)", F_BOLD, FMT_PCT)
    f("H17", '=IF(G17=0,"-",C17/G17)', F_BOLD, FMT_PCT)
    for col in "FHI":
        ws[f"{col}17"].fill = FILL_TOTAL
    ws["I17"].border = BORDER

    # ---- Blok B: region (A20:I28)
    block_title(ws, "A19", "B. Rekap per region")
    hdr = ["region_id", "nama_region", "penjualan_bersih", "hpp", "laba_kotor", "margin_kotor",
           "target", "pencapaian_target", "kontribusi"]
    for j, h in enumerate(hdr, start=1):
        ws.cell(row=20, column=j, value=h)
    style_header(ws, 20, 1, 9)
    regs = list(d["wilayah"]["region_id"])
    for i, reg in enumerate(regs):
        r = 21 + i
        put(ws, f"A{r}", reg, F_INPUT, border=True)
        f(f"B{r}", f"=INDEX(Master!$P$3:$P$8,MATCH(A{r},Master!$O$3:$O$8,0))")
        f(f"C{r}", f"=SUMIFS({R('O')},{R('E')},$A{r})", fmt=FMT_RP)
        f(f"D{r}", f"=SUMIFS({R('Q')},{R('E')},$A{r})", fmt=FMT_RP)
        f(f"E{r}", f"=C{r}-D{r}", F_FORMULA, FMT_RP)
        f(f"F{r}", f"=IF(C{r}=0,0,E{r}/C{r})", F_FORMULA, FMT_PCT)
        f(f"G{r}", f"=SUMIFS({T_VAL},{T_REG},$A{r})", fmt=FMT_RP)
        f(f"H{r}", f'=IF(G{r}=0,"-",C{r}/G{r})', F_FORMULA, FMT_PCT)
        f(f"I{r}", f"=IFERROR(C{r}/$C$27,0)", F_FORMULA, FMT_PCT)
    put(ws, "A27", "TOTAL", F_BOLD, fill=FILL_TOTAL, border=True)
    ws["B27"].fill = FILL_TOTAL
    ws["B27"].border = BORDER
    for col in "CDEG":
        f(f"{col}27", f"=SUM({col}21:{col}26)", F_BOLD, FMT_RP)
        ws[f"{col}27"].fill = FILL_TOTAL
    f("F27", "=IFERROR(E27/C27,0)", F_BOLD, FMT_PCT)
    f("H27", "=IFERROR(C27/G27,0)", F_BOLD, FMT_PCT)
    f("I27", "=SUM(I21:I26)", F_BOLD, FMT_PCT)
    for col in "FHI":
        ws[f"{col}27"].fill = FILL_TOTAL

    # ---- Blok C: kategori (A30:F37)
    block_title(ws, "A29", "C. Rekap per kategori produk")
    hdr = ["kategori", "penjualan_bersih", "hpp", "laba_kotor", "margin_kotor", "kontribusi"]
    for j, h in enumerate(hdr, start=1):
        ws.cell(row=30, column=j, value=h)
    style_header(ws, 30, 1, 6)
    kats = list(dict.fromkeys(d["produk"]["kategori"]))
    for i, k in enumerate(kats):
        r = 31 + i
        put(ws, f"A{r}", k, F_INPUT, border=True)
        f(f"B{r}", f"=SUMIFS({R('O')},{R('U')},$A{r})", fmt=FMT_RP)
        f(f"C{r}", f"=SUMIFS({R('Q')},{R('U')},$A{r})", fmt=FMT_RP)
        f(f"D{r}", f"=B{r}-C{r}", F_FORMULA, FMT_RP)
        f(f"E{r}", f"=IF(B{r}=0,0,D{r}/B{r})", F_FORMULA, FMT_PCT)
        f(f"F{r}", f"=IFERROR(B{r}/$B$36,0)", F_FORMULA, FMT_PCT)
    put(ws, "A36", "TOTAL", F_BOLD, fill=FILL_TOTAL, border=True)
    for col in "BCD":
        f(f"{col}36", f"=SUM({col}31:{col}35)", F_BOLD, FMT_RP)
        ws[f"{col}36"].fill = FILL_TOTAL
    f("E36", "=IFERROR(D36/B36,0)", F_BOLD, FMT_PCT)
    f("F36", "=SUM(F31:F35)", F_BOLD, FMT_PCT)
    for col in "EF":
        ws[f"{col}36"].fill = FILL_TOTAL

    # ---- Blok D: segmen (A39:G44)
    block_title(ws, "A38", "D. Rekap per segmen pelanggan")
    hdr = ["segmen", "jumlah_order", "penjualan_bersih", "hpp", "laba_kotor", "margin_kotor", "kontribusi"]
    for j, h in enumerate(hdr, start=1):
        ws.cell(row=39, column=j, value=h)
    style_header(ws, 39, 1, 7)
    for i, s in enumerate(["Retail", "Grosir", "Korporat"]):
        r = 40 + i
        put(ws, f"A{r}", s, F_INPUT, border=True)
        f(f"B{r}", f"=COUNTIFS({R('V')},$A{r})", fmt=FMT_INT)
        f(f"C{r}", f"=SUMIFS({R('O')},{R('V')},$A{r})", fmt=FMT_RP)
        f(f"D{r}", f"=SUMIFS({R('Q')},{R('V')},$A{r})", fmt=FMT_RP)
        f(f"E{r}", f"=C{r}-D{r}", F_FORMULA, FMT_RP)
        f(f"F{r}", f"=IF(C{r}=0,0,E{r}/C{r})", F_FORMULA, FMT_PCT)
        f(f"G{r}", f"=IFERROR(C{r}/$C$43,0)", F_FORMULA, FMT_PCT)
    put(ws, "A43", "TOTAL", F_BOLD, fill=FILL_TOTAL, border=True)
    for col in "BCDE":
        f(f"{col}43", f"=SUM({col}40:{col}42)", F_BOLD, FMT_RP)
        ws[f"{col}43"].fill = FILL_TOTAL
    f("F43", "=IFERROR(E43/C43,0)", F_BOLD, FMT_PCT)
    f("G43", "=SUM(G40:G42)", F_BOLD, FMT_PCT)
    for col in "FG":
        ws[f"{col}43"].fill = FILL_TOTAL

    # ---- Blok E: per produk (K4:R44) dan Blok F: top 10 produk (T4:Y14)
    block_title(ws, "K3", "E. Rekap per produk (bantu untuk Top 10)")
    hdr = ["product_id", "nama_produk", "kategori", "qty_terjual_neto", "penjualan_bersih",
           "laba_kotor", "margin_kotor", "kunci_peringkat"]
    for j, h in enumerate(hdr, start=11):
        ws.cell(row=4, column=j, value=h)
    style_header(ws, 4, 11, 18)
    for i, pid in enumerate(d["produk"]["product_id"]):
        r = 5 + i
        put(ws, f"K{r}", pid, F_INPUT, border=True)
        f(f"L{r}", f"=INDEX(Master!$B$3:$B$42,MATCH(K{r},Master!$A$3:$A$42,0))")
        f(f"M{r}", f"=INDEX(Master!$C$3:$C$42,MATCH(K{r},Master!$A$3:$A$42,0))")
        f(f"N{r}", f"=SUMIFS({R('G')},{R('D')},$K{r})-SUMIFS({R('M')},{R('D')},$K{r})", fmt=FMT_INT)
        f(f"O{r}", f"=SUMIFS({R('O')},{R('D')},$K{r})", fmt=FMT_RP)
        f(f"P{r}", f"=SUMIFS({R('R')},{R('D')},$K{r})", fmt=FMT_RP)
        f(f"Q{r}", f"=IF(O{r}=0,0,P{r}/O{r})", F_FORMULA, FMT_PCT)
        f(f"R{r}", f'=COUNTIFS($O$5:$O$44,">"&O{r})+COUNTIFS($O$5:O{r},O{r})', F_FORMULA, FMT_INT)
    block_title(ws, "T3", "F. Top 10 produk (SMALL + INDEX/MATCH)")
    hdr = ["peringkat", "kunci", "product_id", "nama_produk", "penjualan_bersih", "margin_kotor"]
    for j, h in enumerate(hdr, start=20):
        ws.cell(row=4, column=j, value=h)
    style_header(ws, 4, 20, 25)
    for k in range(1, 11):
        r = 4 + k
        put(ws, f"T{r}", k, F_INPUT, FMT_INT, border=True)
        f(f"U{r}", f"=SMALL($R$5:$R$44,T{r})", F_FORMULA, FMT_INT)
        f(f"V{r}", f"=INDEX($K$5:$K$44,MATCH(U{r},$R$5:$R$44,0))", F_FORMULA)
        f(f"W{r}", f"=INDEX($L$5:$L$44,MATCH(U{r},$R$5:$R$44,0))", F_FORMULA)
        f(f"X{r}", f"=INDEX($O$5:$O$44,MATCH(U{r},$R$5:$R$44,0))", F_FORMULA, FMT_RP)
        f(f"Y{r}", f"=INDEX($Q$5:$Q$44,MATCH(U{r},$R$5:$R$44,0))", F_FORMULA, FMT_PCT)

    # ---- Blok G: per pelanggan (AA4:AH124) dan Blok H: Pareto (AJ4:AR124)
    block_title(ws, "AA3", "G. Rekap per pelanggan (bantu untuk Top 10 dan Pareto)")
    hdr = ["customer_id", "nama_pelanggan", "segmen", "region_id", "jumlah_order",
           "penjualan_bersih", "laba_kotor", "kunci_peringkat"]
    for j, h in enumerate(hdr, start=27):
        ws.cell(row=4, column=j, value=h)
    style_header(ws, 4, 27, 34)
    ncl = N_CUST + 4  # baris terakhir pelanggan (124)
    for i, cid in enumerate(d["pelanggan"]["customer_id"]):
        r = 5 + i
        put(ws, f"AA{r}", cid, F_INPUT, border=True)
        f(f"AB{r}", f"=INDEX(Master!$I$3:$I$122,MATCH(AA{r},Master!$H$3:$H$122,0))")
        f(f"AC{r}", f"=INDEX(Master!$J$3:$J$122,MATCH(AA{r},Master!$H$3:$H$122,0))")
        f(f"AD{r}", f"=INDEX(Master!$L$3:$L$122,MATCH(AA{r},Master!$H$3:$H$122,0))")
        f(f"AE{r}", f"=COUNTIFS({R('C')},$AA{r})", fmt=FMT_INT)
        f(f"AF{r}", f"=SUMIFS({R('O')},{R('C')},$AA{r})", fmt=FMT_RP)
        f(f"AG{r}", f"=SUMIFS({R('R')},{R('C')},$AA{r})", fmt=FMT_RP)
        f(f"AH{r}", f'=COUNTIFS($AF$5:$AF${ncl},">"&AF{r})+COUNTIFS($AF$5:AF{r},AF{r})', F_FORMULA, FMT_INT)
    block_title(ws, "AJ3", "H. Peringkat pelanggan dan Pareto (baris 5-14 = Top 10)")
    hdr = ["peringkat", "kunci", "customer_id", "nama_pelanggan", "segmen", "penjualan_bersih",
           "kumulatif", "pct_kumulatif_penjualan", "pct_kumulatif_pelanggan"]
    for j, h in enumerate(hdr, start=36):
        ws.cell(row=4, column=j, value=h)
    style_header(ws, 4, 36, 44)
    for k in range(1, N_CUST + 1):
        r = 4 + k
        put(ws, f"AJ{r}", k, F_INPUT, FMT_INT, border=True)
        f(f"AK{r}", f"=SMALL($AH$5:$AH${ncl},AJ{r})", F_FORMULA, FMT_INT)
        f(f"AL{r}", f"=INDEX($AA$5:$AA${ncl},MATCH(AK{r},$AH$5:$AH${ncl},0))", F_FORMULA)
        f(f"AM{r}", f"=INDEX($AB$5:$AB${ncl},MATCH(AK{r},$AH$5:$AH${ncl},0))", F_FORMULA)
        f(f"AN{r}", f"=INDEX($AC$5:$AC${ncl},MATCH(AK{r},$AH$5:$AH${ncl},0))", F_FORMULA)
        f(f"AO{r}", f"=INDEX($AF$5:$AF${ncl},MATCH(AK{r},$AH$5:$AH${ncl},0))", F_FORMULA, FMT_RP)
        f(f"AP{r}", f"=AO{r}" if k == 1 else f"=AP{r - 1}+AO{r}", F_FORMULA, FMT_RP)
        f(f"AQ{r}", f"=IFERROR(AP{r}/$AP${ncl},0)", F_FORMULA, FMT_PCT)
        f(f"AR{r}", f"=AJ{r}/{N_CUST}", F_FORMULA, FMT_PCT)
        if k <= 10 and solusi:
            for col in ["AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR"]:
                ws[f"{col}{r}"].fill = FILL_TOP
    # Ringkasan Pareto
    block_title(ws, "AT3", "I. Ringkasan Pareto")
    put(ws, "AT4", "ukuran", F_HEADER, fill=FILL_HEADER, border=True)
    put(ws, "AU4", "nilai", F_HEADER, fill=FILL_HEADER, border=True)
    put(ws, "AT5", "jumlah_pelanggan", F_BASE, border=True)
    f("AU5", f"=COUNTA($AA$5:$AA${ncl})", F_FORMULA, FMT_INT)
    put(ws, "AT6", "pelanggan_20pct_teratas (n)", F_BASE, border=True)
    f("AU6", "=ROUND(0.2*AU5,0)", F_FORMULA, FMT_INT)
    put(ws, "AT7", "kontribusi_20pct_teratas", F_BASE, border=True)
    f("AU7", f"=INDEX($AQ$5:$AQ${ncl},AU6)", F_FORMULA, FMT_PCT)
    put(ws, "AT8", "pelanggan_untuk_80pct_penjualan", F_BASE, border=True)
    f("AU8", f'=COUNTIFS($AQ$5:$AQ${ncl},"<0.8")+1', F_FORMULA, FMT_INT)
    put(ws, "AT9", "pelanggan_terbesar", F_BASE, border=True)
    f("AU9", "=AM5", F_FORMULA)
    put(ws, "AT10", "share_pelanggan_terbesar", F_BASE, border=True)
    f("AU10", "=AQ5", F_FORMULA, FMT_PCT)
    put(ws, "AT11", "kontribusi_top10_pelanggan", F_BASE, border=True)
    f("AU11", "=AQ14", F_FORMULA, FMT_PCT)

    ws.freeze_panes = "A5"
    set_widths(ws, {"A": 24, "B": 16, "C": 17, "D": 15, "E": 15, "F": 12, "G": 15, "H": 12,
                    "I": 12, "J": 3, "K": 11, "L": 30, "M": 22, "N": 12, "O": 16, "P": 15,
                    "Q": 11, "R": 16, "S": 3, "T": 9, "U": 16, "V": 11, "W": 30, "X": 16,
                    "Y": 11, "Z": 3, "AA": 11, "AB": 34, "AC": 10, "AD": 10, "AE": 11,
                    "AF": 16, "AG": 15, "AH": 16, get_column_letter(35): 3, "AJ": 9, "AK": 16, "AL": 11,
                    "AM": 34, "AN": 10, "AO": 16, "AP": 17, "AQ": 12, "AR": 12, "AS": 3,
                    "AT": 30, "AU": 16})
    return ws


# ---------------------------------------------------------------------------
# Sheet Report (dashboard)
# ---------------------------------------------------------------------------
def build_report(wb, ws_calc, solusi: bool):
    ws = wb.create_sheet("Report")
    ws.sheet_view.showGridLines = False
    put(ws, "B1", "Dashboard Sales & Profitability 2024 (data simulasi)", F_TITLE)
    put(ws, "B2", "Satuan: rupiah (Rp) dan persen. Sumber data: sheet Raw_Data, Master, Calculation. "
                  "Periode: 2024-01-01 s.d. 2024-12-31.", F_SUB)
    put(ws, "J2", "Diperbarui:", F_SUB, align=Alignment(horizontal="right"))
    put(ws, "K2", TANGGAL_BUAT, F_INPUT, FMT_DATE)

    if not solusi:
        put(ws, "B4", "Sheet ini sengaja kosong. Bangun KPI card, chart, dan kotak insight sesuai "
                      "sheet Instruksi (langkah 6).", F_BOLD)
        set_widths(ws, {"A": 2, "B": 18})
        return ws

    # KPI card: label di baris 4, nilai di baris 5, keterangan baris 6
    cards = [
        ("Penjualan bersih 2024 (Rp)", "=Calculation!C17", FMT_RP, "omzet neto - nilai retur"),
        ("Laba kotor 2024 (Rp)", "=Calculation!E17", FMT_RP, "penjualan bersih - HPP"),
        ("Margin kotor", "=Calculation!F17", FMT_PCT, "laba kotor / penjualan bersih"),
        ("Pencapaian target", "=Calculation!H17", FMT_PCT, "penjualan bersih / target"),
        ("Jumlah order", "=COUNTA(Raw_Data!A2:A6001)", FMT_INT, "baris transaksi"),
        ("Nilai retur (Rp)", "=SUM(Raw_Data!N2:N6001)", FMT_RP, "=TEXT(SUM(Raw_Data!N2:N6001)/SUM(Raw_Data!L2:L6001),\"0.00%\")&\" dari omzet neto\""),
    ]
    col = 2
    for label, formula, fmt, ket in cards:
        c1, c2 = get_column_letter(col), get_column_letter(col + 1)
        ws.merge_cells(f"{c1}4:{c2}4")
        ws.merge_cells(f"{c1}5:{c2}5")
        ws.merge_cells(f"{c1}6:{c2}6")
        put(ws, f"{c1}4", label, F_KPI_LABEL, fill=FILL_CARD, align=Alignment(horizontal="center"))
        put(ws, f"{c1}5", formula, F_KPI, fmt, fill=FILL_CARD, align=Alignment(horizontal="center"))
        put(ws, f"{c1}6", ket, F_SUB if not ket.startswith("=") else Font(name=FONT_NAME, size=9, italic=True, color="008000"),
            fill=FILL_CARD, align=Alignment(horizontal="center"))
        for r in (4, 5, 6):
            for cc in (c1, c2):
                ws[f"{cc}{r}"].border = BORDER
        col += 2

    # ---- sel bantu untuk kalimat insight (kolom P:Q)
    put(ws, "P3", "Sel bantu insight (berumus; dipakai kotak insight)", F_BOLD, fill=FILL_BLOCK)
    bantu = [
        ("bulan_tertinggi", "=INDEX(Calculation!A5:A16,MATCH(MAX(Calculation!C5:C16),Calculation!C5:C16,0))", None),
        ("nilai_bulan_tertinggi", "=MAX(Calculation!C5:C16)", FMT_RP),
        ("bulan_terendah", "=INDEX(Calculation!A5:A16,MATCH(MIN(Calculation!C5:C16),Calculation!C5:C16,0))", None),
        ("nilai_bulan_terendah", "=MIN(Calculation!C5:C16)", FMT_RP),
        ("region_terbesar", "=INDEX(Calculation!B21:B26,MATCH(MAX(Calculation!C21:C26),Calculation!C21:C26,0))", None),
        ("kontribusi_region_terbesar", "=MAX(Calculation!I21:I26)", FMT_PCT),
        ("region_pencapaian_terendah", "=INDEX(Calculation!B21:B26,MATCH(MIN(Calculation!H21:H26),Calculation!H21:H26,0))", None),
        ("pencapaian_terendah", "=MIN(Calculation!H21:H26)", FMT_PCT),
        ("region_pencapaian_tertinggi", "=INDEX(Calculation!B21:B26,MATCH(MAX(Calculation!H21:H26),Calculation!H21:H26,0))", None),
        ("pencapaian_tertinggi", "=MAX(Calculation!H21:H26)", FMT_PCT),
        ("jumlah_region_di_bawah_target", '=COUNTIFS(Calculation!H21:H26,"<1")', FMT_INT),
        ("kategori_terbesar", "=INDEX(Calculation!A31:A35,MATCH(MAX(Calculation!B31:B35),Calculation!B31:B35,0))", None),
        ("kontribusi_kategori_terbesar", "=MAX(Calculation!F31:F35)", FMT_PCT),
        ("kategori_margin_tertipis", "=INDEX(Calculation!A31:A35,MATCH(MIN(Calculation!E31:E35),Calculation!E31:E35,0))", None),
        ("margin_tertipis", "=MIN(Calculation!E31:E35)", FMT_PCT),
        ("kontribusi_kategori_tertipis", "=INDEX(Calculation!F31:F35,MATCH(MIN(Calculation!E31:E35),Calculation!E31:E35,0))", FMT_PCT),
        ("kategori_margin_tertebal", "=INDEX(Calculation!A31:A35,MATCH(MAX(Calculation!E31:E35),Calculation!E31:E35,0))", None),
        ("margin_tertebal", "=MAX(Calculation!E31:E35)", FMT_PCT),
        ("pelanggan_terbesar", "=Calculation!AU9", None),
        ("share_pelanggan_terbesar", "=Calculation!AU10", FMT_PCT),
        ("kontribusi_20pct_teratas", "=Calculation!AU7", FMT_PCT),
        ("pelanggan_untuk_80pct", "=Calculation!AU8", FMT_INT),
        ("jumlah_pelanggan", "=Calculation!AU5", FMT_INT),
        ("mom_desember", "=Calculation!I16", FMT_PCT),
        ("segmen_terbesar", "=INDEX(Calculation!A40:A42,MATCH(MAX(Calculation!C40:C42),Calculation!C40:C42,0))", None),
        ("kontribusi_segmen_terbesar", "=MAX(Calculation!G40:G42)", FMT_PCT),
    ]
    B = {}
    for i, (name, formula, fmt) in enumerate(bantu):
        r = 4 + i
        put(ws, f"P{r}", name, F_BASE, border=True)
        put(ws, f"Q{r}", formula, F_LINK, fmt, border=True)
        B[name] = f"$Q${r}"

    # ---- Chart 1: tren bulanan vs target (line)
    ch = LineChart()
    ch.title = "Tren bulanan: penjualan bersih vs target (Rp)"
    ch.y_axis.title = "Rp"
    ch.x_axis.title = "Bulan"
    ch.y_axis.number_format = "#,##0"
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.add_data(Reference(ws_calc, min_col=3, min_row=4, max_row=16), titles_from_data=True)
    ch.add_data(Reference(ws_calc, min_col=7, min_row=4, max_row=16), titles_from_data=True)
    ch.set_categories(Reference(ws_calc, min_col=1, min_row=5, max_row=16))
    ch.series[0].graphicalProperties.line.solidFill = "1F4E78"
    ch.series[0].graphicalProperties.line.width = 28000
    ch.series[1].graphicalProperties.line.solidFill = "C00000"
    ch.series[1].graphicalProperties.line.dashStyle = "dash"
    ch.height, ch.width = 7.5, 15.5
    ch.legend.position = "b"
    ws.add_chart(ch, "B8")

    # ---- Chart 2: kontribusi region (bar)
    ch = BarChart()
    ch.type = "bar"
    ch.title = "Penjualan bersih per region (Rp)"
    ch.y_axis.number_format = "#,##0"
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.add_data(Reference(ws_calc, min_col=3, min_row=20, max_row=26), titles_from_data=True)
    ch.set_categories(Reference(ws_calc, min_col=2, min_row=21, max_row=26))
    ch.series[0].graphicalProperties.solidFill = "1F4E78"
    ch.legend = None
    ch.height, ch.width = 7.5, 12.5
    ws.add_chart(ch, "I8")

    # ---- Chart 3: kategori (bar: penjualan bersih dan laba kotor)
    ch = BarChart()
    ch.type = "col"
    ch.title = "Penjualan bersih dan laba kotor per kategori (Rp)"
    ch.y_axis.number_format = "#,##0"
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.add_data(Reference(ws_calc, min_col=2, min_row=30, max_row=35), titles_from_data=True)
    ch.add_data(Reference(ws_calc, min_col=4, min_row=30, max_row=35), titles_from_data=True)
    ch.set_categories(Reference(ws_calc, min_col=1, min_row=31, max_row=35))
    ch.series[0].graphicalProperties.solidFill = "1F4E78"
    ch.series[1].graphicalProperties.solidFill = "70AD47"
    ch.legend.position = "b"
    ch.height, ch.width = 7.5, 15.5
    ws.add_chart(ch, "B24")

    # ---- Chart 4: Pareto pelanggan (bar + line kumulatif, sumbu sekunder)
    bar = BarChart()
    bar.type = "col"
    bar.title = "Pareto pelanggan: penjualan bersih per peringkat dan % kumulatif"
    bar.y_axis.number_format = "#,##0"
    bar.x_axis.title = "Peringkat pelanggan"
    bar.x_axis.delete = False
    bar.y_axis.delete = False
    bar.add_data(Reference(ws_calc, min_col=41, min_row=4, max_row=124), titles_from_data=True)
    bar.set_categories(Reference(ws_calc, min_col=36, min_row=5, max_row=124))
    bar.series[0].graphicalProperties.solidFill = "1F4E78"
    bar.gapWidth = 30
    line = LineChart()
    line.add_data(Reference(ws_calc, min_col=43, min_row=4, max_row=124), titles_from_data=True)
    line.y_axis.axId = 200
    line.y_axis.number_format = "0%"
    line.y_axis.title = "% kumulatif"
    line.y_axis.crosses = "max"
    line.y_axis.scaling.min = 0
    line.y_axis.scaling.max = 1
    line.y_axis.delete = False
    line.series[0].graphicalProperties.line.solidFill = "C00000"
    line.series[0].graphicalProperties.line.width = 22000
    bar += line
    bar.legend.position = "b"
    bar.height, bar.width = 7.5, 12.5
    ws.add_chart(bar, "I24")

    # ---- Kotak insight dan rekomendasi (berumus, mengutip angka dari sel)
    put(ws, "B40", "Insight utama (angka dihitung langsung dari sel Calculation)", F_BOLD, fill=FILL_BLOCK)
    for cc in "CDEFGHIJKLM":
        ws[f"{cc}40"].fill = FILL_BLOCK
    insights = [
        '="1. Penjualan bersih 2024 mencapai Rp "&TEXT(Calculation!C17,"#,##0")&" dengan laba kotor Rp "'
        '&TEXT(Calculation!E17,"#,##0")&" (margin kotor "&TEXT(Calculation!F17,"0.0%")&"). '
        'Pencapaian target total "&TEXT(Calculation!H17,"0.0%")&"."',
        f'="2. Tren musiman: bulan tertinggi "&{B["bulan_tertinggi"]}&" (Rp "&TEXT({B["nilai_bulan_tertinggi"]},"#,##0")'
        f'&"), terendah "&{B["bulan_terendah"]}&" (Rp "&TEXT({B["nilai_bulan_terendah"]},"#,##0")&"). '
        f'Pertumbuhan MoM Desember "&TEXT({B["mom_desember"]},"0.0%")&"."',
        f'="3. Region "&{B["region_terbesar"]}&" menyumbang "&TEXT({B["kontribusi_region_terbesar"]},"0.0%")'
        f'&" penjualan bersih. "&{B["jumlah_region_di_bawah_target"]}&" dari 6 region di bawah target; terendah "'
        f'&{B["region_pencapaian_terendah"]}&" ("&TEXT({B["pencapaian_terendah"]},"0.0%")&"), tertinggi "'
        f'&{B["region_pencapaian_tertinggi"]}&" ("&TEXT({B["pencapaian_tertinggi"]},"0.0%")&")."',
        f'="4. Kategori terbesar "&{B["kategori_terbesar"]}&" ("&TEXT({B["kontribusi_kategori_terbesar"]},"0.0%")'
        f'&" penjualan). Margin tertipis pada "&{B["kategori_margin_tertipis"]}&" ("&TEXT({B["margin_tertipis"]},"0.0%")'
        f'&") padahal menyumbang "&TEXT({B["kontribusi_kategori_tertipis"]},"0.0%")&" penjualan; margin tertebal "'
        f'&{B["kategori_margin_tertebal"]}&" ("&TEXT({B["margin_tertebal"]},"0.0%")&")."',
        f'="5. Pareto: 20% pelanggan teratas (24 dari "&{B["jumlah_pelanggan"]}&") menyumbang "'
        f'&TEXT({B["kontribusi_20pct_teratas"]},"0.0%")&" penjualan; 80% penjualan berasal dari "&{B["pelanggan_untuk_80pct"]}'
        f'&" pelanggan. Pelanggan terbesar "&{B["pelanggan_terbesar"]}&" sendiri "&TEXT({B["share_pelanggan_terbesar"]},"0.0%")'
        f'&". Segmen "&{B["segmen_terbesar"]}&" = "&TEXT({B["kontribusi_segmen_terbesar"]},"0.0%")&"."',
    ]
    for i, fml in enumerate(insights):
        r = 41 + i
        ws.merge_cells(f"B{r}:M{r}")
        put(ws, f"B{r}", fml, F_LINK, align=Alignment(wrap_text=True, vertical="top"))
        ws.row_dimensions[r].height = 38
    put(ws, "B47", "Rekomendasi", F_BOLD, fill=FILL_BLOCK)
    for cc in "CDEFGHIJKLM":
        ws[f"{cc}47"].fill = FILL_BLOCK
    rekom = [
        f'="1. Tinjau kebijakan diskon dan harga pada kategori "&{B["kategori_margin_tertipis"]}&": margin "'
        f'&TEXT({B["margin_tertipis"]},"0.0%")&" jauh di bawah margin total "&TEXT(Calculation!F17,"0.0%")'
        f'&"; skenario margin +2 poin, penjualan tetap: tambahan laba Rp "'
        f'&TEXT(0.02*INDEX(Calculation!B31:B35,MATCH({B["kategori_margin_tertipis"]},Calculation!A31:A35,0)),"#,##0")&"."',
        f'="2. Fokus pemulihan target pada region "&{B["region_pencapaian_terendah"]}&" (pencapaian "'
        f'&TEXT({B["pencapaian_terendah"]},"0.0%")&"): tinjau target, cakupan tim sales, dan bauran produk; '
        f'pelajari praktik region "&{B["region_pencapaian_tertinggi"]}&"."',
        f'="3. Kelola risiko konsentrasi: "&{B["pelanggan_terbesar"]}&" menyumbang "&TEXT({B["share_pelanggan_terbesar"]},"0.0%")'
        f'&" penjualan. Siapkan kontrak/retensi untuk 24 pelanggan teratas ("&TEXT({B["kontribusi_20pct_teratas"]},"0.0%")'
        f'&" penjualan) dan program pengembangan pelanggan menengah."',
        f'="4. Siapkan stok dan tenaga penjualan untuk puncak "&{B["bulan_tertinggi"]}&" dan promosi pada bulan lesu "'
        f'&{B["bulan_terendah"]}&" agar beban kerja dan arus kas lebih merata."',
    ]
    for i, fml in enumerate(rekom):
        r = 48 + i
        ws.merge_cells(f"B{r}:M{r}")
        put(ws, f"B{r}", fml, F_LINK, align=Alignment(wrap_text=True, vertical="top"))
        ws.row_dimensions[r].height = 38
    put(ws, "B53", "Catatan: data pada proyek ini adalah data sintetis yang dibangkitkan dengan skrip; "
                   "tidak ada data perusahaan atau pribadi nyata.", F_SUB)

    set_widths(ws, {"A": 2, "B": 13, "C": 13, "D": 13, "E": 13, "F": 13, "G": 13, "H": 13,
                    "I": 13, "J": 13, "K": 13, "L": 13, "M": 13, "N": 3, "O": 3, "P": 30, "Q": 24})
    ws.row_dimensions[5].height = 26
    # Pengaturan cetak: satu halaman landscape (dipakai untuk preview PDF)
    ws.print_area = "A1:N39"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.4
    return ws


# ---------------------------------------------------------------------------
# Sheet Checks
# ---------------------------------------------------------------------------
def build_checks(wb, d, solusi: bool):
    ws = wb.create_sheet("Checks")
    nt, nr = len(d["trans"]) + 1, len(d["retur"]) + 1
    R = lambda col: f"Raw_Data!${col}$2:${col}${nt}"
    put(ws, "A1", "Checks: rekonsiliasi dan kualitas data", F_TITLE)
    put(ws, "A2", "Pembanding di sheet ini berupa rumus. Rekonsiliasi independen pandas ada di VERIFIKASI.md. Selisih dibulatkan 2 desimal harus nol.", F_SUB)
    headers = ["no", "pemeriksaan", "nilai_workbook", "nilai_pembanding", "selisih", "status", "keterangan"]
    for j, value in enumerate(headers, 1):
        ws.cell(4, j, value)
    style_header(ws, 4, 1, 7)
    checks = [
        ("Omzet bruto langsung", f"=SUM({R('J')})", f"=SUMPRODUCT({R('G')},{R('H')})"),
        ("Penjualan bersih langsung", f"=SUM({R('O')})", f"=SUMPRODUCT({R('G')}-{R('M')},{R('H')},1-{R('I')})"),
        ("HPP langsung", f"=SUM({R('Q')})", f"=SUMPRODUCT({R('G')}-{R('M')},{R('P')})"),
        ("Laba kotor langsung", f"=SUM({R('R')})", f"=SUM({R('O')})-SUM({R('Q')})"),
    ]
    for label, raw, monthly in [("Omzet bruto", "J", "B"), ("Penjualan bersih", "O", "C"), ("HPP", "Q", "D"), ("Laba kotor", "R", "E")]:
        checks.append((label + " vs bulanan", f"=SUM({R(raw)})", f"=Calculation!{monthly}17"))
    checks += [
        ("Penjualan vs region", f"=SUM({R('O')})", "=Calculation!C27"),
        ("Penjualan vs kategori", f"=SUM({R('O')})", "=Calculation!B36"),
        ("Penjualan vs pelanggan", f"=SUM({R('O')})", "=SUM(Calculation!AF5:AF124)"),
        ("Duplikat order_id", f'=SUMPRODUCT((COUNTIFS({R("A")},{R("A")})>1)*1)', "=0"),
        ("Duplikat order dan produk", f"=SUM({R('X')})", "=0"),
        ("Produk tanpa master", f'=SUMPRODUCT((COUNTIFS(Master!$A$3:$A$42,{R("D")})<>1)*1)', "=0"),
        ("Pelanggan tanpa master", f'=SUMPRODUCT((COUNTIFS(Master!$H$3:$H$122,{R("C")})<>1)*1)', "=0"),
        ("Region tanpa master", f'=SUMPRODUCT((COUNTIFS(Master!$O$3:$O$8,{R("E")})<>1)*1)', "=0"),
        ("Sales tanpa master", f'=SUMPRODUCT((COUNTIFS(Master!$S$3:$S$14,{R("F")})<>1)*1)', "=0"),
        ("Retur tanpa order", f'=COUNTIFS(Retur!$F$2:$F${nr},"Tidak")', "=0"),
        ("Retur melebihi pembelian", f"=SUMPRODUCT(({R('M')}>{R('G')})*1)", "=0"),
        ("Sel mentah kosong", f"=COUNTBLANK(Raw_Data!$A$2:$I${nt})", "=0"),
        ("Qty tidak positif", f'=COUNTIFS({R("G")},"<=0")', "=0"),
        ("Harga tidak positif", f'=COUNTIFS({R("H")},"<=0")', "=0"),
        ("Diskon di luar rentang", f'=COUNTIFS({R("I")},"<0")+COUNTIFS({R("I")},">0.25")', "=0"),
        ("Tanggal di luar periode", f'=COUNTIFS({R("B")},"<"&DATE(2024,1,1))+COUNTIFS({R("B")},">"&DATE(2024,12,31))', "=0"),
        ("Target tidak positif", '=COUNTIFS(Master!Y3:Y74,"<=0")', "=0"),
        ("Target total", "=Calculation!G17", "=SUM(Master!Y3:Y74)"),
        ("Kontribusi region", "=Calculation!I27", "=1"),
        ("Kontribusi kategori", "=Calculation!F36", "=1"),
        ("Pareto berakhir 100%", "=Calculation!AQ124", "=1"),
    ]
    for i, (label, actual, expected) in enumerate(checks, 5):
        put(ws, f"A{i}", f"={i}-4", F_FORMULA, border=True)
        put(ws, f"B{i}", label, F_BASE, border=True)
        put(ws, f"C{i}", actual if solusi else None, F_LINK, FMT_RP, border=True)
        if not solusi:
            ws[f"C{i}"].fill = FILL_INPUT
        put(ws, f"D{i}", expected, F_LINK, FMT_RP, border=True)
        put(ws, f"E{i}", f'=IF(C{i}="","",ROUND(C{i}-D{i},2))', F_FORMULA, "#,##0.00", border=True)
        put(ws, f"F{i}", f'=IF(C{i}="","BELUM DIISI",IF(E{i}=0,"PASS","FAIL"))', F_BOLD, border=True)
        put(ws, f"G{i}", "Selisih aktual dan pembanding; lihat pemeriksaan independen di dokumen.", F_SUB, border=True)
    last = 4 + len(checks)
    put(ws, f"B{last+2}", "STATUS KESELURUHAN", F_BOLD, fill=FILL_TOTAL)
    put(ws, f"F{last+2}", f'=IF(COUNTIFS(F5:F{last},"BELUM DIISI")>0,"BELUM DIISI",IF(COUNTIFS(F5:F{last},"FAIL")=0,"PASS","FAIL"))', F_BOLD)
    for status, color in [("PASS", "C6EFCE"), ("FAIL", "FFC7CE")]:
        ws.conditional_formatting.add(f"F5:F{last+2}", CellIsRule(operator="equal", formula=[f'"{status}"'], fill=PatternFill("solid", fgColor=color)))
    ws.freeze_panes = "A5"
    set_widths(ws, {"A": 5, "B": 35, "C": 20, "D": 20, "E": 15, "F": 18, "G": 60})
    return ws


# ---------------------------------------------------------------------------
# Sheet Readme / Instruksi
# ---------------------------------------------------------------------------
def build_readme(wb, solusi: bool):
    ws = wb.create_sheet("Readme" if solusi else "Instruksi", 0)
    ws.sheet_view.showGridLines = False
    judul = ("Sales & Profitability Analysis 2024: workbook solusi" if solusi
             else "Sales & Profitability Analysis 2024: workbook latihan")
    put(ws, "B2", judul, F_TITLE)
    put(ws, "B3", "Proyek portofolio dengan data SIMULASI. Data sintetis dibangkitkan dengan skrip "
                  "(tools/generate_data.py); tidak ada data perusahaan atau pribadi nyata.", F_SUB)
    put(ws, "B4", "Tanggal pembuatan:", F_BOLD)
    put(ws, "C4", TANGGAL_BUAT, F_INPUT, FMT_DATE)
    put(ws, "B5", "Periode data:", F_BOLD)
    put(ws, "C5", "2024-01-01 s.d. 2024-12-31 (12 bulan)", F_BASE)
    put(ws, "B6", "Satuan:", F_BOLD)
    put(ws, "C6", "Rupiah ditampilkan tanpa desimal (format #,##0) dan persen (0.0%)", F_BASE)

    r = 8
    put(ws, f"B{r}", "Tujuan", F_BOLD, fill=FILL_BLOCK)
    ws[f"C{r}"].fill = FILL_BLOCK
    put(ws, f"B{r + 1}", "Mengukur penjualan bersih, HPP, laba kotor, dan margin dari data transaksi; "
                         "membandingkan dengan target per region; menemukan kontribusi kategori, "
                         "segmen, produk, dan pelanggan (Pareto) sebagai dasar rekomendasi.", F_BASE,
        align=Alignment(wrap_text=True, vertical="top"))
    ws.merge_cells(f"B{r + 1}:G{r + 1}")
    ws.row_dimensions[r + 1].height = 45

    r = 11
    put(ws, f"B{r}", "Daftar sheet dan alur (Input -> Proses -> Output)", F_BOLD, fill=FILL_BLOCK)
    ws[f"C{r}"].fill = FILL_BLOCK
    sheets = [
        ("Raw_Data", "INPUT + PROSES. 6.000 baris transaksi (kolom A:I, data mentah) dan kolom perhitungan "
                     "per baris (J:X): omzet bruto, diskon, omzet neto, qty retur (SUMIFS ke Retur), nilai retur, "
                     "penjualan bersih, HPP satuan (INDEX/MATCH ke Master), HPP, laba kotor, margin, periode, "
                     "kategori, segmen, kunci unik, penanda duplikat."),
        ("Retur", "INPUT. 120 baris retur; kolom F memeriksa apakah order_id ada di Raw_Data."),
        ("Master", "INPUT. Lima tabel referensi berdampingan: Produk (A:F), Pelanggan (H:M), Wilayah (O:Q), "
                   "Sales (S:U), Target per bulan x region (W:Y)."),
        ("Calculation", "PROSES. Rekap bulanan (blok A), per region (B), per kategori (C), per segmen (D), "
                        "per produk dan Top 10 produk (E, F), per pelanggan, peringkat dan Pareto (G, H), "
                        "ringkasan Pareto (I). Semua SUMIFS/COUNTIFS/SMALL/INDEX-MATCH."),
        ("Report", "OUTPUT. Dashboard: 6 KPI card, chart tren bulanan vs target, bar region, bar kategori, "
                   "Pareto pelanggan, kotak insight dan rekomendasi yang mengutip angka dari sel."),
        ("Checks", "KONTROL. Pemeriksaan rekonsiliasi (jumlah baris, total nilai, keunikan kunci, relasi "
                   "ke master, missing, rentang nilai) dengan status PASS/FAIL berumus."),
    ]
    if not solusi:
        sheets = [
            ("Raw_Data", "Data mentah A:I. Kolom J:X berlabel dan KOSONG (kuning); baris 2 diisi contoh rumus."),
            ("Retur", "Data retur. Kolom F berlabel; baris 2 contoh."),
            ("Master", "Tabel referensi: Produk, Pelanggan, Wilayah, Sales, Target."),
            ("Calculation", "Kerangka rekap dengan header dan kunci (periode, region, kategori, segmen, produk, "
                            "pelanggan) sudah ada; sel kuning harus diisi rumus."),
            ("Report", "Kosong: bangun KPI card, chart, dan insight."),
            ("Checks", "Kerangka pemeriksaan; isi kolom nilai_workbook (kuning) dengan rumus."),
        ]
    for i, (nm, desc) in enumerate(sheets):
        rr = r + 1 + i
        put(ws, f"B{rr}", nm, F_BOLD, border=True)
        put(ws, f"C{rr}", desc, F_BASE, border=True, align=Alignment(wrap_text=True, vertical="top"))
        ws.merge_cells(f"C{rr}:G{rr}")
        ws.row_dimensions[rr].height = 48 if solusi else 30

    r = 19
    put(ws, f"B{r}", "Legenda warna sel", F_BOLD, fill=FILL_BLOCK)
    ws[f"C{r}"].fill = FILL_BLOCK
    put(ws, f"B{r + 1}", "Teks biru", F_INPUT, border=True)
    put(ws, f"C{r + 1}", "Input: data mentah, kunci, dan parameter. Boleh diganti.", F_BASE, border=True)
    put(ws, f"B{r + 2}", "Teks hitam", F_FORMULA, border=True)
    put(ws, f"C{r + 2}", "Rumus yang merujuk sel pada sheet yang sama. Jangan ditimpa.", F_BASE, border=True)
    put(ws, f"B{r + 3}", "Teks hijau", F_LINK, border=True)
    put(ws, f"C{r + 3}", "Rumus yang merujuk sheet lain (link antar-sheet).", F_BASE, border=True)
    put(ws, f"B{r + 4}", "Latar kuning", F_BASE, fill=FILL_INPUT, border=True)
    put(ws, f"C{r + 4}", "Sel yang harus diisi (hanya pada workbook latihan).", F_BASE, border=True)
    put(ws, f"B{r + 5}", "Latar hijau muda", F_BASE, fill=FILL_TOP, border=True)
    put(ws, f"C{r + 5}", "Baris Top 10 pada tabel peringkat pelanggan (Calculation blok H).", F_BASE, border=True)
    for i in range(1, 6):
        ws.merge_cells(f"C{r + i}:G{r + i}")

    r = 26
    put(ws, f"B{r}", "Definisi KPI (persis seperti docs/definisi_kpi.md)", F_BOLD, fill=FILL_BLOCK)
    ws[f"C{r}"].fill = FILL_BLOCK
    kpi = [
        "omzet_bruto = qty x harga_satuan",
        "nilai_diskon = omzet_bruto x diskon_pct; omzet_neto = omzet_bruto - nilai_diskon",
        "nilai_retur = qty_retur x harga_satuan x (1 - diskon_pct) order terkait; penjualan_bersih = omzet_neto - nilai_retur",
        "hpp = (qty - qty_retur) x hpp_satuan; laba_kotor = penjualan_bersih - hpp",
        "margin_kotor = laba_kotor / penjualan_bersih (IF pembagi nol -> 0 hanya saat penjualan bersih = 0, yaitu order yang diretur penuh)",
        "pencapaian_target = penjualan_bersih / target; pertumbuhan_mom = (bulan ini - bulan lalu) / bulan lalu (tanda '-' bila bulan lalu 0)",
    ]
    for i, t in enumerate(kpi):
        put(ws, f"B{r + 1 + i}", t, F_BASE)

    r = 34
    put(ws, f"B{r}", "Catatan kompatibilitas", F_BOLD, fill=FILL_BLOCK)
    ws[f"C{r}"].fill = FILL_BLOCK
    notes = [
        "Rumus memakai fungsi Excel 2016+ (SUMIFS, INDEX/MATCH, SMALL, SUMPRODUCT, TEXT, IFERROR) dan dapat dibuka di LibreOffice Calc.",
        "Pemisah argumen di file ini adalah koma (format internal). Di Excel regional Indonesia rumus tampil dengan titik koma.",
        "Jika TEXT(tanggal;\"yyyy-mm\") menampilkan hasil aneh pada Excel berbahasa lain, pakai =YEAR(B2)&\"-\"&TEXT(MONTH(B2);\"00\").",
        "PivotTable, Power Query, dan Slicer tidak tertanam di file; lihat docs/pivot_table.md dan docs/power_query.md untuk langkah manual.",
    ]
    for i, t in enumerate(notes):
        put(ws, f"B{r + 1 + i}", t, F_BASE)

    if not solusi:
        r = 40
        put(ws, f"B{r}", "Tugas bertahap (kerjakan berurutan; cocokkan dengan workbook solusi bila perlu)", F_BOLD, fill=FILL_BLOCK)
        ws[f"C{r}"].fill = FILL_BLOCK
        tugas = [
            "Langkah 1 (Raw_Data J:L): isi omzet_bruto, nilai_diskon, omzet_neto untuk semua baris. Salin rumus baris 2 ke bawah (klik ganda fill handle).",
            "Langkah 2 (Raw_Data M:O): qty_retur dengan SUMIFS ke sheet Retur berdasarkan order_id; nilai_retur; penjualan_bersih.",
            "Langkah 3 (Raw_Data P:S): hpp_satuan dengan INDEX/MATCH ke Master Produk; hpp; laba_kotor; margin_kotor dengan IF pembagi nol.",
            "Langkah 4 (Raw_Data T:X): periode dengan TEXT; kategori dan segmen dengan INDEX/MATCH; kunci_unik; duplikat dengan COUNTIFS.",
            "Langkah 5 (Calculation): isi semua sel kuning blok A sampai I dengan SUMIFS/COUNTIFS/SMALL/INDEX-MATCH. Periksa total blok A = total blok B = total blok C.",
            "Langkah 6 (Report): buat 6 KPI card merujuk Calculation; chart line bulanan vs target; bar region; bar kategori; combo Pareto; kotak insight berumus (TEXT untuk format angka).",
            "Langkah 7 (Checks): isi kolom nilai_workbook; pastikan semua status PASS. Bila FAIL, cari penyebabnya di Raw_Data atau Calculation.",
            "Langkah 8 (opsional): buat PivotTable dari Raw_Data (docs/pivot_table.md) dan bandingkan dengan Calculation; coba Power Query (docs/power_query.md).",
        ]
        for i, t in enumerate(tugas):
            put(ws, f"B{r + 1 + i}", t, F_BASE, align=Alignment(wrap_text=True, vertical="top"))
            ws.merge_cells(f"B{r + 1 + i}:G{r + 1 + i}")
            ws.row_dimensions[r + 1 + i].height = 30
    set_widths(ws, {"A": 2, "B": 20, "C": 24, "D": 20, "E": 20, "F": 20, "G": 20})
    return ws


def build(solusi: bool, d):
    wb = Workbook()
    wb.properties.creator = "Immantaka Juniarsa"
    wb.calculation = CalcProperties(calcId=0, fullCalcOnLoad=True)
    wb.remove(wb.active)
    build_readme(wb, solusi)
    build_raw_data(wb, d, solusi)
    build_retur(wb, d, solusi)
    build_master(wb, d)
    ws_calc = build_calculation(wb, d, solusi)
    build_report(wb, ws_calc, solusi)
    build_checks(wb, d, solusi)
    # Tabel membantu audit, filter, dan langkah PivotTable manual.
    tables = [("Raw_Data", "Transaksi", f"A1:X{N_TRANS+1}"),
              ("Retur", "DataRetur", f"A1:F{N_RETUR+1}"),
              ("Master", "Produk", "A2:F42"), ("Master", "Pelanggan", "H2:M122"),
              ("Master", "Wilayah", "O2:Q8"), ("Master", "Sales", "S2:U14"),
              ("Master", "Target", "W2:Y74")]
    for sheet, table_name, ref in tables:
        tab = Table(displayName=table_name, ref=ref)
        tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        wb[sheet].add_table(tab)
    for ws in wb:
        if not ws.freeze_panes:
            ws.freeze_panes = "B3"
    intro = wb["Readme" if solusi else "Instruksi"]
    for row in list(range(27, 33)) + list(range(35, 39)):
        intro.merge_cells(start_row=row, start_column=2, end_row=row, end_column=7)
        intro.cell(row, 2).alignment = Alignment(wrap_text=True, vertical="top")
        intro.row_dimensions[row].height = 32
    intro.print_options.horizontalCentered = True
    wb["Report"].freeze_panes = "B8"
    name = "solusi_sales_profitability.xlsx" if solusi else "latihan_sales_profitability.xlsx"
    WB_DIR.mkdir(parents=True, exist_ok=True)
    wb.save(WB_DIR / name)
    print("tersimpan:", WB_DIR / name)


def main():
    global N_TRANS, N_RETUR
    d = load_data()
    N_TRANS, N_RETUR = len(d["trans"]), len(d["retur"])
    assert len(d["produk"]) == N_PRODUK and len(d["pelanggan"]) == N_CUST and len(d["target"]) == N_TARGET
    build(True, d)
    build(False, d)


if __name__ == "__main__":
    main()
