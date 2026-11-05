"""Fungsi bantu openpyxl yang dipakai skrip build_workbook.py.

Konvensi warna (ditulis juga di sheet Readme):
- biru  : sel input/parameter yang boleh diubah
- hitam : rumus
- hijau : tautan antar-sheet
"""
from __future__ import annotations

import re

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

FONT_NAME = "Arial"
FONT = Font(name=FONT_NAME, size=10)
FONT_BOLD = Font(name=FONT_NAME, size=10, bold=True)
FONT_TITLE = Font(name=FONT_NAME, size=14, bold=True, color="1F3864")
FONT_SUB = Font(name=FONT_NAME, size=11, bold=True, color="1F3864")
FONT_HEADER = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
FONT_INPUT = Font(name=FONT_NAME, size=10, color="0000FF")
FONT_INPUT_BOLD = Font(name=FONT_NAME, size=10, color="0000FF", bold=True)
FONT_LINK = Font(name=FONT_NAME, size=10, color="008000")
FONT_NOTE = Font(name=FONT_NAME, size=9, italic=True, color="595959")
FONT_KPI = Font(name=FONT_NAME, size=16, bold=True, color="1F3864")

FILL_HEADER = PatternFill("solid", start_color="1F3864")
FILL_INPUT = PatternFill("solid", start_color="FFF2CC")
FILL_SECTION = PatternFill("solid", start_color="D9E1F2")
FILL_KPI = PatternFill("solid", start_color="EDF2F9")
FILL_PASS = PatternFill("solid", start_color="E2EFDA")
FILL_FAIL = PatternFill("solid", start_color="FCE4D6")

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

# Fungsi Excel pasca-2007 yang di dalam file XML harus diberi awalan _xlfn.
_XLFN = [
    "STDEV.S", "VAR.S", "CONFIDENCE.T", "T.TEST", "T.INV.2T", "NORM.S.INV",
    "RANK.EQ", "IFS", "MAXIFS", "MINIFS", "SWITCH", "TEXTJOIN", "FORECAST.LINEAR",
    "PERCENTILE.INC", "QUARTILE.INC", "MODE.SNGL", "NETWORKDAYS.INTL", "CONCAT",
    "DAYS",
]
_XLFN_RE = re.compile(r"(?<![A-Za-z0-9_.])(" + "|".join(re.escape(f) for f in sorted(_XLFN, key=len, reverse=True)) + r")\(")


def fx(formula: str) -> str:
    """Tambahkan awalan _xlfn. pada fungsi pasca-2007 agar Excel dan LibreOffice mengenalinya."""
    return _XLFN_RE.sub(lambda m: "_xlfn." + m.group(1) + "(", formula)


def col(n: int) -> str:
    return get_column_letter(n)


def set_widths(ws, widths: dict):
    for k, w in widths.items():
        ws.column_dimensions[k if isinstance(k, str) else col(k)].width = w


def autofit(ws, min_w=8, max_w=60, skip=()):
    widths = {}
    for row in ws.iter_rows():
        for c in row:
            if c.value is None or c.column_letter in skip:
                continue
            v = str(c.value)
            if v.startswith("="):
                length = 12
            else:
                length = max(len(x) for x in v.split("\n")) + 2
            widths[c.column_letter] = max(widths.get(c.column_letter, 0), length)
    for k, w in widths.items():
        if ws.column_dimensions[k].width in (None, 13.0) or ws.column_dimensions[k].width < w:
            ws.column_dimensions[k].width = max(min_w, min(max_w, w))


def title(ws, text, row=1, sub=None):
    ws.cell(row=row, column=1, value=text).font = FONT_TITLE
    if sub:
        ws.cell(row=row + 1, column=1, value=sub).font = FONT_NOTE


def section(ws, row, text, ncols=8):
    c = ws.cell(row=row, column=1, value=text)
    c.font = FONT_SUB
    for i in range(1, ncols + 1):
        ws.cell(row=row, column=i).fill = FILL_SECTION
    return row + 1


def header_row(ws, row, headers, start_col=1, fill=FILL_HEADER, font=FONT_HEADER):
    for i, h in enumerate(headers):
        c = ws.cell(row=row, column=start_col + i, value=h)
        c.font = font
        c.fill = fill
        c.alignment = CENTER
        c.border = BORDER
    ws.row_dimensions[row].height = 30


def put(ws, ref, value, font=FONT, fmt=None, fill=None, align=None, border=True):
    c = ws[ref]
    c.value = value
    c.font = font
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if align:
        c.alignment = align
    if border:
        c.border = BORDER
    return c


def put_formula(ws, ref, formula, fmt=None, link=False, fill=None, bold=False):
    f = FONT_LINK if link else (FONT_BOLD if bold else FONT)
    return put(ws, ref, fx(formula), font=f, fmt=fmt, fill=fill)


def put_input(ws, ref, value, fmt=None, bold=False):
    return put(ws, ref, value, font=FONT_INPUT_BOLD if bold else FONT_INPUT, fmt=fmt, fill=FILL_INPUT)


def put_text(ws, ref, text, bold=False, wrap=False, note=False):
    f = FONT_NOTE if note else (FONT_BOLD if bold else FONT)
    return put(ws, ref, text, font=f, align=WRAP if wrap else None, border=False)


def write_df(ws, df, start_row=1, start_col=1, header=True, num_formats=None, text_cols=(), date_cols=()):
    """Tulis DataFrame sebagai tabel statis (data mentah)."""
    r = start_row
    if header:
        header_row(ws, r, list(df.columns), start_col=start_col)
        r += 1
    num_formats = num_formats or {}
    for rec in df.itertuples(index=False):
        for j, v in enumerate(rec):
            name = df.columns[j]
            c = ws.cell(row=r, column=start_col + j)
            if v is None or (isinstance(v, float) and v != v):
                c.value = None
            else:
                c.value = v
            c.font = FONT
            c.border = BORDER
            if name in num_formats:
                c.number_format = num_formats[name]
        r += 1
    return r


def freeze(ws, ref="A2"):
    ws.freeze_panes = ref


def print_setup(ws, landscape=True, fit_width=True, area=None):
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1 if fit_width else 0
    ws.page_setup.fitToHeight = 0
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    if area:
        ws.print_area = area
