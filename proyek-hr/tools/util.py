"""Fungsi bantu bersama untuk membangun semua workbook portofolio.

Semua rumus ditulis dengan pemisah koma (format internal file Excel).
Di Excel regional Indonesia rumus yang sama tampil dengan titik koma.
"""

from __future__ import annotations

import datetime as dt

from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

# ---------------------------------------------------------------- gaya dasar
FONT_NAME = "Arial"
WARNA_HEADER = "1F4E78"
WARNA_SUB = "D9E1F2"
WARNA_INPUT = "FFF2CC"
WARNA_KPI = "E2EFDA"

FMT_RP = '"Rp"#,##0'
FMT_RP_NEG = '"Rp"#,##0;[Red]-"Rp"#,##0'
FMT_PCT = "0.0%"
FMT_TGL = "dd/mm/yyyy"
FMT_ANGKA = "#,##0"
FMT_DES2 = "0.00"

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

DEPARTEMEN = ["HR", "Finance", "Operasional", "Marketing", "IT", "Produksi"]
STATUS_KARYAWAN = ["Tetap", "Kontrak", "Magang"]

LIBUR_2024 = [
    (dt.date(2024, 1, 1), "Tahun Baru Masehi"),
    (dt.date(2024, 2, 8), "Isra Mikraj"),
    (dt.date(2024, 2, 9), "Cuti Bersama Imlek"),
    (dt.date(2024, 2, 10), "Tahun Baru Imlek"),
    (dt.date(2024, 3, 11), "Hari Raya Nyepi"),
    (dt.date(2024, 3, 12), "Cuti Bersama Nyepi"),
    (dt.date(2024, 3, 29), "Wafat Isa Almasih"),
    (dt.date(2024, 3, 31), "Hari Paskah"),
    (dt.date(2024, 4, 8), "Cuti Bersama Idulfitri"),
    (dt.date(2024, 4, 9), "Cuti Bersama Idulfitri"),
    (dt.date(2024, 4, 10), "Hari Raya Idulfitri"),
    (dt.date(2024, 4, 11), "Hari Raya Idulfitri"),
    (dt.date(2024, 4, 12), "Cuti Bersama Idulfitri"),
    (dt.date(2024, 4, 15), "Cuti Bersama Idulfitri"),
    (dt.date(2024, 5, 1), "Hari Buruh Internasional"),
    (dt.date(2024, 5, 9), "Kenaikan Isa Almasih"),
    (dt.date(2024, 5, 10), "Cuti Bersama Kenaikan Isa Almasih"),
    (dt.date(2024, 5, 23), "Hari Raya Waisak"),
    (dt.date(2024, 5, 24), "Cuti Bersama Waisak"),
    (dt.date(2024, 6, 1), "Hari Lahir Pancasila"),
    (dt.date(2024, 6, 17), "Iduladha"),
    (dt.date(2024, 6, 18), "Cuti Bersama Iduladha"),
    (dt.date(2024, 7, 7), "Tahun Baru Islam"),
    (dt.date(2024, 8, 17), "Hari Kemerdekaan RI"),
    (dt.date(2024, 9, 16), "Maulid Nabi Muhammad SAW"),
    (dt.date(2024, 12, 25), "Hari Raya Natal"),
    (dt.date(2024, 12, 26), "Cuti Bersama Natal"),
]


def font(bold=False, color="000000", size=10, italic=False):
    return Font(name=FONT_NAME, bold=bold, color=color, size=size, italic=italic)


def judul(ws, teks, sub=None, baris=1):
    """Tulis judul besar (dan subjudul) di baris tertentu."""
    c = ws.cell(row=baris, column=1, value=teks)
    c.font = font(bold=True, size=14, color=WARNA_HEADER)
    if sub:
        s = ws.cell(row=baris + 1, column=1, value=sub)
        s.font = font(italic=True, size=9, color="595959")


def header(ws, baris, kolom_awal, labels, fill=WARNA_HEADER, warna_teks="FFFFFF", wrap=True):
    """Tulis baris header bergaya."""
    for i, lab in enumerate(labels):
        c = ws.cell(row=baris, column=kolom_awal + i, value=lab)
        c.font = font(bold=True, color=warna_teks)
        c.fill = PatternFill("solid", fgColor=fill)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=wrap)
        c.border = BORDER
    ws.row_dimensions[baris].height = 30 if wrap else 18


def tulis_baris(ws, baris, kolom_awal, nilai, formats=None, border=True):
    """Tulis satu baris data; nilai berupa list (string '=...' dianggap rumus)."""
    for i, v in enumerate(nilai):
        c = ws.cell(row=baris, column=kolom_awal + i, value=v)
        if border:
            c.border = BORDER
        if formats and formats[i]:
            c.number_format = formats[i]
    return baris


def tabel(ws, ref, nama, gaya="TableStyleMedium2"):
    """Jadikan rentang sebagai Format as Table."""
    t = Table(displayName=nama, ref=ref)
    t.tableStyleInfo = TableStyleInfo(
        name=gaya, showFirstColumn=False, showLastColumn=False,
        showRowStripes=True, showColumnStripes=False,
    )
    ws.add_table(t)
    return t


def lebar(ws, widths):
    """Atur lebar kolom: dict {'A': 12} atau list berurutan dari A."""
    if isinstance(widths, dict):
        for k, v in widths.items():
            ws.column_dimensions[k].width = v
    else:
        for i, v in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = v


def dropdown(ws, rentang, sumber, judul_err="Nilai tidak valid", pesan="Pilih dari daftar."):
    """Data validation tipe list. `sumber` bisa string rumus ('=Daftar!$A$2:$A$7') atau list."""
    if isinstance(sumber, (list, tuple)):
        f = '"' + ",".join(sumber) + '"'
    else:
        f = sumber
    dv = DataValidation(type="list", formula1=f, allow_blank=True, showErrorMessage=True,
                        errorTitle=judul_err, error=pesan)
    ws.add_data_validation(dv)
    dv.add(rentang)
    return dv


def validasi_angka(ws, rentang, minimum, maksimum, pesan="Masukkan angka pada rentang yang diizinkan."):
    dv = DataValidation(type="whole", operator="between", formula1=str(minimum), formula2=str(maksimum),
                        allow_blank=True, showErrorMessage=True, errorTitle="Angka tidak valid", error=pesan)
    ws.add_data_validation(dv)
    dv.add(rentang)
    return dv


def cf_kurang_dari(ws, rentang, batas, warna="FFC7CE", teks="9C0006"):
    ws.conditional_formatting.add(
        rentang,
        CellIsRule(operator="lessThan", formula=[str(batas)],
                   fill=PatternFill("solid", fgColor=warna), font=Font(color=teks)),
    )


def cf_sama_dengan(ws, rentang, teks_nilai, warna="FFC7CE", teks="9C0006"):
    ws.conditional_formatting.add(
        rentang,
        CellIsRule(operator="equal", formula=[f'"{teks_nilai}"'],
                   fill=PatternFill("solid", fgColor=warna), font=Font(color=teks)),
    )


def cf_rumus(ws, rentang, rumus, warna="FFC7CE", teks="9C0006"):
    ws.conditional_formatting.add(
        rentang,
        FormulaRule(formula=[rumus], fill=PatternFill("solid", fgColor=warna), font=Font(color=teks)),
    )


def cf_skala_warna(ws, rentang):
    ws.conditional_formatting.add(
        rentang,
        ColorScaleRule(start_type="min", start_color="F8696B", mid_type="percentile", mid_value=50,
                       mid_color="FFEB84", end_type="max", end_color="63BE7B"),
    )


def label_nilai(ws, baris, kolom, label, nilai, fmt=None, fill=None, bold_nilai=True):
    """Pasangan sel label (kiri) dan nilai/rumus (kanan)."""
    a = ws.cell(row=baris, column=kolom, value=label)
    a.font = font(bold=False)
    a.border = BORDER
    b = ws.cell(row=baris, column=kolom + 1, value=nilai)
    b.font = font(bold=bold_nilai)
    b.border = BORDER
    if fmt:
        b.number_format = fmt
    if fill:
        b.fill = PatternFill("solid", fgColor=fill)
    return b


def kartu_kpi(ws, baris, kolom, label, rumus, fmt=None, lebar_kolom=2):
    """Kartu KPI sederhana: label kecil di atas, angka besar di bawah (2 baris x n kolom)."""
    kol_akhir = kolom + lebar_kolom - 1
    ws.merge_cells(start_row=baris, start_column=kolom, end_row=baris, end_column=kol_akhir)
    ws.merge_cells(start_row=baris + 1, start_column=kolom, end_row=baris + 1, end_column=kol_akhir)
    a = ws.cell(row=baris, column=kolom, value=label)
    a.font = font(bold=True, color="FFFFFF", size=9)
    a.fill = PatternFill("solid", fgColor=WARNA_HEADER)
    a.alignment = Alignment(horizontal="center", vertical="center")
    b = ws.cell(row=baris + 1, column=kolom, value=rumus)
    b.font = font(bold=True, size=18, color=WARNA_HEADER)
    b.fill = PatternFill("solid", fgColor=WARNA_KPI)
    b.alignment = Alignment(horizontal="center", vertical="center")
    if fmt:
        b.number_format = fmt
    ws.row_dimensions[baris + 1].height = 32
    for r in (baris, baris + 1):
        for c in range(kolom, kol_akhir + 1):
            ws.cell(row=r, column=c).border = BORDER


def chart_bar(ws, anchor, judul_chart, kategori_ref, data_ref, judul_y=None, judul_x=None,
              lebar_cm=16, tinggi_cm=8, horizontal=False, tumpuk=False, warna=None):
    ch = BarChart()
    ch.type = "bar" if horizontal else "col"
    ch.title = judul_chart
    ch.style = 10
    if tumpuk:
        ch.grouping = "stacked"
        ch.overlap = 100
    ch.add_data(data_ref, titles_from_data=True)
    ch.set_categories(kategori_ref)
    ch.y_axis.title = judul_y
    ch.x_axis.title = judul_x
    ch.width = lebar_cm
    ch.height = tinggi_cm
    ch.legend.position = "b"
    if warna:
        for s, w in zip(ch.series, warna):
            s.graphicalProperties.solidFill = w
    ws.add_chart(ch, anchor)
    return ch


def chart_line(ws, anchor, judul_chart, kategori_ref, data_ref, judul_y=None, lebar_cm=16, tinggi_cm=8):
    ch = LineChart()
    ch.title = judul_chart
    ch.style = 12
    ch.add_data(data_ref, titles_from_data=True)
    ch.set_categories(kategori_ref)
    ch.y_axis.title = judul_y
    ch.width = lebar_cm
    ch.height = tinggi_cm
    ch.legend.position = "b"
    ws.add_chart(ch, anchor)
    return ch


def chart_pie(ws, anchor, judul_chart, kategori_ref, data_ref, lebar_cm=10, tinggi_cm=8):
    ch = PieChart()
    ch.title = judul_chart
    ch.add_data(data_ref, titles_from_data=True)
    ch.set_categories(kategori_ref)
    ch.width = lebar_cm
    ch.height = tinggi_cm
    ws.add_chart(ch, anchor)
    return ch


def ref(ws, min_col, min_row, max_col=None, max_row=None):
    return Reference(ws, min_col=min_col, min_row=min_row, max_col=max_col or min_col, max_row=max_row or min_row)


def teks_panduan(ws, baris_awal, baris_teks, lebar_kolom=110):
    """Tulis daftar paragraf panduan (satu per baris) di kolom A dengan wrap."""
    ws.column_dimensions["A"].width = lebar_kolom
    r = baris_awal
    for t in baris_teks:
        c = ws.cell(row=r, column=1, value=t)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if t.startswith("## "):
            c.value = t[3:]
            c.font = font(bold=True, size=12, color=WARNA_HEADER)
        elif t.startswith("# "):
            c.value = t[2:]
            c.font = font(bold=True, size=14, color=WARNA_HEADER)
        else:
            c.font = font()
        r += 1
    return r


def rapikan(wb):
    """Terapkan font Arial pada semua sel berisi dan atur halaman cetak (landscape, muat lebar)."""
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.value is None:
                    continue
                f = c.font
                if f is None or f.name != FONT_NAME:
                    c.font = Font(name=FONT_NAME, bold=f.bold, italic=f.italic, color=f.color,
                                  size=f.size or 10, underline=f.underline)
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.print_options.gridLines = False
        ws.sheet_view.showGridLines = True
    wb.calculation.fullCalcOnLoad = True
    return wb


def kolom(n):
    return get_column_letter(n)


def rp(x):
    return int(round(x / 1000.0)) * 1000
