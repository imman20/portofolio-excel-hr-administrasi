"""Membangun workbook solusi dan latihan proyek Survei Kepuasan Pelanggan.

Jalankan: python tools/build_workbook.py
Setelah itu jalankan recalc LibreOffice (lihat VERIFIKASI.md) agar nilai rumus terisi.
"""
from __future__ import annotations

import datetime as dt
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, str(Path(__file__).resolve().parent))
from xlsx_helpers import (  # noqa: E402
    BORDER, CENTER, FILL_FAIL, FILL_INPUT, FILL_KPI, FILL_PASS, FONT, FONT_BOLD, FONT_INPUT,
    FONT_KPI, FONT_LINK, FONT_NOTE, WRAP, autofit, freeze, fx, header_row, print_setup, put,
    put_formula, put_input, put_text, section, set_widths, title, write_df,
)

BASE = Path(__file__).resolve().parents[1]
RAW_CSV = BASE / "data" / "raw" / "survei_raw.csv"
REF_CSV = BASE / "data" / "processed" / "survei_bersih_referensi.csv"
OUT_SOLUSI = BASE / "workbook" / "solusi_survei_kepuasan_pelanggan.xlsx"
OUT_LATIHAN = BASE / "workbook" / "latihan_survei_kepuasan_pelanggan.xlsx"
TANGGAL_BUAT = "2026-11-05"

ITEMS = [f"q{i:02d}" for i in range(1, 13)]
DIMENSI = [
    ("Produk", ["q01", "q02", "q03"]),
    ("Layanan", ["q04", "q05", "q06"]),
    ("Harga", ["q07", "q08", "q09"]),
    ("Pengiriman", ["q10", "q11", "q12"]),
]
PERTANYAAN = {
    "q01": "Kualitas produk sesuai harapan saya",
    "q02": "Produk awet dan dapat diandalkan",
    "q03": "Pilihan/variasi produk memadai",
    "q04": "Staf melayani dengan ramah dan sopan",
    "q05": "Staf tanggap menyelesaikan masalah saya",
    "q06": "Informasi yang diberikan jelas dan akurat",
    "q07": "Harga sebanding dengan kualitas",
    "q08": "Harga bersaing dibanding tempat lain",
    "q09": "Promo dan diskon menarik",
    "q10": "Pesanan tiba tepat waktu",
    "q11": "Barang tiba dalam kondisi baik",
    "q12": "Biaya pengiriman wajar",
    "kepuasan_keseluruhan": "Secara keseluruhan, seberapa puas Anda dengan kami?",
    "nps": "Seberapa besar kemungkinan Anda merekomendasikan kami kepada teman/rekan? (0-10)",
}
KANAL = ["Toko", "Online", "Call Center"]
KOTA = ["Surabaya", "Jakarta", "Bandung", "Semarang", "Medan", "Makassar"]
SEGMEN = ["Retail", "Grosir", "Korporat"]

RAW_FIRST, RAW_LAST = 2, 421  # 420 baris data
RNG = f"$2:${RAW_LAST}"


def rr(colletter, sheet="Clean_Data", absolute=True):
    """Referensi kolom penuh baris data, mis. Clean_Data!$Y$2:$Y$421."""
    if absolute:
        return f"{sheet}!${colletter}${RAW_FIRST}:${colletter}${RAW_LAST}"
    return f"{sheet}!{colletter}{RAW_FIRST}:{colletter}{RAW_LAST}"


# ---------------------------------------------------------------------------
# Data mentah
# ---------------------------------------------------------------------------
def muat_raw() -> pd.DataFrame:
    df = pd.read_csv(RAW_CSV, dtype=str, keep_default_na=False)
    return df


def nilai_sel_raw(v: str):
    """Angka polos ditulis sebagai angka; selain itu teks apa adanya (spasi, 'N/A', '-')."""
    if v == "":
        return None
    if v.strip() == v and v.lstrip("-").isdigit():
        return int(v)
    return v


def sheet_raw(wb, df, nama="Raw_Data"):
    ws = wb.create_sheet(nama)
    header_row(ws, 1, list(df.columns))
    for i, rec in enumerate(df.itertuples(index=False), start=2):
        for j, v in enumerate(rec, start=1):
            c = ws.cell(row=i, column=j)
            c.value = v if v != "" else None
            c.font = FONT
            c.border = BORDER
    freeze(ws, "B2")
    set_widths(ws, {"A": 13, "B": 13, "C": 12, "D": 11, "E": 10, "F": 7, "G": 7, "H": 9, "W": 40})
    for k in range(9, 23):
        ws.column_dimensions[ws.cell(row=1, column=k).column_letter].width = 6
    ws.auto_filter.ref = f"A1:W{RAW_LAST}"
    print_setup(ws, area="A1:W40")
    return ws


# ---------------------------------------------------------------------------
# Master dan Codebook
# ---------------------------------------------------------------------------
MASTER = {
    "kanal": "Master!$B$4:$B$6",
    "usia_min": "Master!$B$11", "usia_max": "Master!$B$12",
    "lik_min": "Master!$B$13", "lik_max": "Master!$B$14",
    "nps_min": "Master!$B$15", "nps_max": "Master!$B$16",
    "taraf": "Master!$B$17", "f_krit": "Master!$B$18",
    "promotor": "Master!$B$19", "detraktor": "Master!$B$20",
}


def sheet_master(wb):
    ws = wb.create_sheet("Master")
    title(ws, "Master: daftar nilai standar dan parameter", sub="Sel biru = parameter yang boleh diubah; rumus di sheet lain mengacu ke sini.")
    header_row(ws, 3, ["Kanal standar"], start_col=1)
    ws.merge_cells("A3:B3")
    for i, k in enumerate(KANAL, start=4):
        put(ws, f"A{i}", i - 3)
        put_input(ws, f"B{i}", k)
    put_text(ws, "A8", "Catatan: MATCH tidak membedakan huruf besar-kecil, sehingga 'toko', 'TOKO' dan 'Toko' semua cocok ke baris 'Toko'.", note=True)

    header_row(ws, 10, ["Parameter", "Nilai", "Keterangan"], start_col=1)
    params = [
        ("Usia minimum valid", 17, "Usia di bawah ini dianggap salah input (mis. 3) dan dikosongkan"),
        ("Usia maksimum valid", 80, "Usia di atas ini dianggap salah input (mis. 150) dan dikosongkan"),
        ("Skor Likert minimum", 1, "Nilai di luar 1-5 (mis. 0, 6, 7) dikosongkan"),
        ("Skor Likert maksimum", 5, ""),
        ("NPS minimum", 0, "Skala rekomendasi 0-10"),
        ("NPS maksimum", 10, ""),
        ("Taraf signifikansi (alpha)", 0.05, "Dipakai CONFIDENCE.T, T.INV.2T dan keputusan uji"),
        ("Catatan ANOVA", "Deskriptif", "Keputusan inferensial memakai p-value tepat melalui Analysis ToolPak"),
        ("Ambang promotor (NPS ≥)", 9, "Promotor = skor 9-10"),
        ("Ambang detraktor (NPS ≤)", 6, "Detraktor = skor 0-6; pasif = 7-8"),
    ]
    for i, (nama, nilai, ket) in enumerate(params, start=11):
        put(ws, f"A{i}", nama)
        put_input(ws, f"B{i}", nilai, fmt="0.00" if isinstance(nilai, float) else "0")
        put(ws, f"C{i}", ket, font=FONT_NOTE)

    header_row(ws, 3, ["Kota"], start_col=4)
    for i, k in enumerate(KOTA, start=4):
        put_input(ws, f"D{i}", k)
    header_row(ws, 3, ["Segmen"], start_col=6)
    for i, k in enumerate(SEGMEN, start=4):
        put_input(ws, f"F{i}", k)

    header_row(ws, 3, ["Item", "Dimensi", "Pernyataan"], start_col=8)
    r = 4
    for dim, items in DIMENSI:
        for it in items:
            put_input(ws, f"H{r}", it)
            put_input(ws, f"I{r}", dim)
            put(ws, f"J{r}", PERTANYAAN[it])
            r += 1
    set_widths(ws, {"A": 38, "B": 14, "C": 62, "D": 12, "E": 3, "F": 12, "G": 3, "H": 8, "I": 12, "J": 44})
    print_setup(ws)
    return ws


def sheet_codebook(wb):
    ws = wb.create_sheet("Codebook")
    title(ws, "Codebook: definisi variabel survei", sub="Sumber: codebook.md. Skala Likert 1 = Sangat tidak setuju ... 5 = Sangat setuju.")
    header_row(ws, 3, ["Variabel", "Label / pertanyaan lengkap", "Tipe", "Skala / kode", "Nilai valid", "Penanganan nilai tidak valid"])
    rows = [
        ("respondent_id", "Kode unik responden (R-0001 ...)", "Teks", "ID", "Unik per responden", "Duplikat: baris kedua dan seterusnya dibuang (COUNTIFS kumulatif)"),
        ("tanggal_survei", "Tanggal pengisian survei", "Tanggal", "YYYY-MM-DD; sebagian DD/MM/YYYY", "2024-01-01 s.d. 2024-12-31", "Diparse dengan LEFT/MID/RIGHT + DATE sesuai pola"),
        ("kanal", "Kanal transaksi yang dinilai", "Kategori", "Toko / Online / Call Center", "3 kategori di Master", "TRIM + MATCH (tidak peka huruf besar-kecil) ke daftar standar"),
        ("kota", "Kota domisili responden", "Kategori", "6 kota", "Lihat Master", "TRIM"),
        ("segmen", "Segmen pelanggan", "Kategori", "Retail / Grosir / Korporat", "3 kategori", "TRIM"),
        ("usia", "Usia responden (tahun)", "Angka", "Tahun", "17-80", "Di luar rentang dikosongkan (bukan dihapus barisnya)"),
        ("jenis_kelamin", "Jenis kelamin", "Kategori", "L = laki-laki, P = perempuan", "L/P", "TRIM"),
        ("lama_pelanggan_bulan", "Lama menjadi pelanggan (bulan)", "Angka", "Bulan", "1-120", "VALUE(TRIM())"),
    ]
    for it in ITEMS:
        dim = next(d for d, its in DIMENSI if it in its)
        rows.append((it, f"[{dim}] {PERTANYAAN[it]}", "Likert", "1=Sangat tidak setuju, 2=Tidak setuju, 3=Netral, 4=Setuju, 5=Sangat setuju", "1-5", "Kosong, 'N/A', '-', 0, 6, 7 → kosong; angka berspasi → TRIM lalu VALUE"))
    rows += [
        ("kepuasan_keseluruhan", PERTANYAAN["kepuasan_keseluruhan"], "Likert", "1=Sangat tidak puas ... 5=Sangat puas", "1-5", "Sama seperti item Likert"),
        ("nps", PERTANYAAN["nps"], "Angka", "0-10; detraktor 0-6, pasif 7-8, promotor 9-10", "0-10", "Kosong/teks → kosong; tidak ikut hitung NPS"),
        ("komentar", "Komentar bebas (opsional)", "Teks", "Teks bebas", "Boleh kosong", "Tidak dianalisis kuantitatif; dibaca kualitatif"),
        ("dim_produk ... dim_pengiriman", "Skor dimensi = rata-rata 3 item dimensi (dihitung di Clean_Data)", "Angka", "1-5", "Hanya jika ketiga item terisi", "Jika ada item kosong, skor dimensi kosong (listwise per dimensi)"),
    ]
    for i, row in enumerate(rows, start=4):
        for j, v in enumerate(row, start=1):
            put(ws, f"{ws.cell(row=i, column=j).column_letter}{i}", v, align=WRAP)
    set_widths(ws, {"A": 22, "B": 52, "C": 10, "D": 40, "E": 22, "F": 52})
    freeze(ws, "A4")
    print_setup(ws)
    return ws


# ---------------------------------------------------------------------------
# Clean_Data
# ---------------------------------------------------------------------------
def tgl_formula(r):
    x = f"Raw_Data!B{r}"
    return (
        f'=IFERROR(IF(MID({x},5,1)="-",DATE(VALUE(LEFT({x},4)),VALUE(MID({x},6,2)),VALUE(RIGHT({x},2))),'
        f'DATE(VALUE(RIGHT({x},4)),VALUE(MID({x},4,2)),VALUE(LEFT({x},2)))),"")'
    )


def angka_rentang(x, lo, hi):
    v = f"VALUE(TRIM({x}))"
    return f'=IF(TRIM({x})="","",IFERROR(IF(AND({v}>={lo},{v}<={hi},{v}=ROUND({v},0)),{v},""),""))'


CLEAN_COLS = []  # (letter, header, formula_fn, fmt)


def _def_clean_cols():
    cols = []
    cols.append(("A", "respondent_id", lambda r: f"=Raw_Data!A{r}", None, True))
    cols.append(("B", "tanggal_survei", tgl_formula, "dd/mm/yyyy", False))
    cols.append(("C", "kanal", lambda r: f'=IFERROR(INDEX({MASTER["kanal"]},MATCH(TRIM(Raw_Data!C{r}),{MASTER["kanal"]},0)),"")', None, False))
    cols.append(("D", "kota", lambda r: f"=TRIM(Raw_Data!D{r})", None, False))
    cols.append(("E", "segmen", lambda r: f"=TRIM(Raw_Data!E{r})", None, False))
    cols.append(("F", "usia", lambda r: angka_rentang(f"Raw_Data!F{r}", MASTER["usia_min"], MASTER["usia_max"]), "0", False))
    cols.append(("G", "jenis_kelamin", lambda r: f"=TRIM(Raw_Data!G{r})", None, False))
    cols.append(("H", "lama_pelanggan_bulan", lambda r: f'=IFERROR(VALUE(TRIM(Raw_Data!H{r})),"")', "0", False))
    raw_letters = "IJKLMNOPQRST"
    for it, L in zip(ITEMS, raw_letters):
        cols.append((L, it, (lambda L: lambda r: angka_rentang(f"Raw_Data!{L}{r}", MASTER["lik_min"], MASTER["lik_max"]))(L), "0", False))
    cols.append(("U", "kepuasan_keseluruhan", lambda r: angka_rentang(f"Raw_Data!U{r}", MASTER["lik_min"], MASTER["lik_max"]), "0", False))
    cols.append(("V", "nps", lambda r: angka_rentang(f"Raw_Data!V{r}", MASTER["nps_min"], MASTER["nps_max"]), "0", False))
    cols.append(("W", "status_duplikat", lambda r: f'=IF(COUNTIFS(Raw_Data!$A${RAW_FIRST}:$A${RAW_LAST},Raw_Data!A{r})=1,"Unik",IF(COUNTIFS(Raw_Data!$A${RAW_FIRST}:$A{r},Raw_Data!A{r})=1,"Duplikat - dipakai","Duplikat - dibuang"))', None, False))
    cols.append(("X", "dipakai", lambda r: f"=IF(COUNTIFS(Raw_Data!$A${RAW_FIRST}:$A{r},Raw_Data!A{r})=1,1,0)", "0", False))
    dim_src = {"Produk": ("I", "K"), "Layanan": ("L", "N"), "Harga": ("O", "Q"), "Pengiriman": ("R", "T")}
    dim_letters = ["Y", "Z", "AA", "AB"]
    for (dim, _), L in zip(DIMENSI, dim_letters):
        a, b = dim_src[dim]
        cols.append((L, f"dim_{dim.lower()}", (lambda a, b: lambda r: f'=IF(AND($X{r}=1,COUNT({a}{r}:{b}{r})=3),AVERAGE({a}{r}:{b}{r}),"")')(a, b), "0.00", False))
    cols.append(("AC", "kk_analisis", lambda r: f'=IF(AND($X{r}=1,COUNT(U{r})=1),U{r},"")', "0", False))
    cols.append(("AD", "nps_analisis", lambda r: f'=IF(AND($X{r}=1,COUNT(V{r})=1),V{r},"")', "0", False))
    cols.append(("AE", "lengkap", lambda r: f"=IF(COUNT(Y{r}:AC{r})=5,1,0)", "0", False))
    cols.append(("AF", "no_lengkap", lambda r: f'=IF(AE{r}=1,COUNTIFS($AE${RAW_FIRST}:$AE{r},1),"")', "0", False))
    alpha_letters = [get_column_letter(i) for i in range(33,45)]
    for it, L, src in zip(ITEMS, alpha_letters, raw_letters):
        dim = next(d for d, its in DIMENSI if it in its)
        dl = dim_letters[[d for d, _ in DIMENSI].index(dim)]
        cols.append((L, f"{it}_alpha", (lambda src, dl: lambda r: f'=IF(${dl}{r}<>"",{src}{r},"")')(src, dl), "0", False))
    for k, L in zip(KANAL, ["AS", "AT", "AU"]):
        cols.append((L, f"kk_{k.lower().replace(' ', '_')}", (lambda k: lambda r: f'=IF($C{r}="{k}",$AC{r},"")')(k), "0", False))
    cols.append(("AV", "komentar", lambda r: f'=IF(Raw_Data!W{r}="","",Raw_Data!W{r})', None, True))
    return cols


CLEAN_COLS = _def_clean_cols()
CLEAN = {h: L for L, h, _, _, _ in CLEAN_COLS}


def sheet_clean(wb, isi_rumus=True):
    ws = wb.create_sheet("Clean_Data")
    header_row(ws, 1, [h for _, h, _, _, _ in CLEAN_COLS])
    if isi_rumus:
        for r in range(RAW_FIRST, RAW_LAST + 1):
            for L, h, fn, fmt, link in CLEAN_COLS:
                put_formula(ws, f"{L}{r}", fn(r), fmt=fmt, link=link)
    freeze(ws, "B2")
    for L, h, _, _, _ in CLEAN_COLS:
        ws.column_dimensions[L].width = 11 if len(h) <= 8 else min(20, len(h) + 2)
    ws.column_dimensions["W"].width = 18
    # Conditional formatting: baris duplikat yang dibuang diberi warna
    ws.conditional_formatting.add(f"A2:X{RAW_LAST}", FormulaRule(formula=[f"$X2=0"], fill=PatternFill("solid", start_color="FCE4D6")))
    ws.auto_filter.ref = f"A1:AV{RAW_LAST}"
    print_setup(ws, area="A1:X40")
    audit_headers = ["skor_kosong", "teks_na", "teks_strip", "angka_berspasi", "likert_0", "likert_6", "likert_7", "usia_invalid", "tanggal_dmy", "kanal_tidak_baku"]
    header_row(ws, 1, audit_headers, start_col=49)
    if isi_rumus:
        for r in range(2, RAW_LAST+1):
            terms = [f'IF(TRIM(Raw_Data!{L}{r})="",0,IFERROR(IF(VALUE(TRIM(Raw_Data!{L}{r}))={{v}},1,0),0))' for L in "IJKLMNOPQRSTU"]
            fs = [f'=COUNTBLANK(Raw_Data!I{r}:V{r})', f'=COUNTIFS(Raw_Data!I{r}:V{r},"N/A")', f'=COUNTIFS(Raw_Data!I{r}:V{r},"-")', f'=SUMPRODUCT(--(LEN(Raw_Data!I{r}:V{r})<>LEN(TRIM(Raw_Data!I{r}:V{r}))))']
            fs += ['='+'+'.join(t.format(v=v) for t in terms) for v in [0,6,7]]
            fs += [f'=IF(AND(Raw_Data!F{r}<>"",F{r}=""),1,0)', f'=IF(MID(Raw_Data!B{r},3,1)="/",1,0)', f'=IF(OR(SUBSTITUTE(Raw_Data!C{r},"Toko","")="",SUBSTITUTE(Raw_Data!C{r},"Online","")="",SUBSTITUTE(Raw_Data!C{r},"Call Center","")=""),0,1)']
            for j,f in enumerate(fs,49): put_formula(ws, f"{ws.cell(1,j).column_letter}{r}", f)
    return ws


def sheet_analisis_set(wb, isi_rumus=True):
    ws = wb.create_sheet("Analisis_Set")
    title(ws, "Analisis_Set: kasus lengkap (4 skor dimensi + kepuasan keseluruhan terisi) ditumpuk rapat untuk CORREL dan LINEST")
    put_text(ws, "A2", "Baris diambil dari Clean_Data dengan INDEX/MATCH pada nomor urut 'no_lengkap'. Baris setelah n_lengkap otomatis kosong. Rentang analisis dibatasi dengan INDEX (Calculation).", note=True)
    header_row(ws, 3, ["no", "respondent_id", "Produk", "Layanan", "Harga", "Pengiriman", "kepuasan_keseluruhan", "nps", "kanal"])
    if isi_rumus:
        for i, r in enumerate(range(4, 4 + (RAW_LAST - RAW_FIRST + 1)), start=1):
            put(ws, f"A{r}", i, fmt="0")
            m = f"MATCH($A{r},{rr('AF')},0)"
            put_formula(ws, f"B{r}", f'=IFERROR(INDEX({rr("A")},{m}),"")', link=True)
            for L, src in zip("CDEFGH", ["Y", "Z", "AA", "AB", "AC", "AD"]):
                put_formula(ws, f"{L}{r}", f'=IFERROR(INDEX({rr(src)},{m}),"")', fmt="0.00", link=True)
            put_formula(ws, f"I{r}", f'=IFERROR(INDEX({rr("C")},{m}),"")', link=True)
    freeze(ws, "A4")
    set_widths(ws, {"A": 6, "B": 14, "C": 10, "D": 10, "E": 10, "F": 11, "G": 20, "H": 8, "I": 12})
    print_setup(ws, area="A1:I40")
    return ws


# ---------------------------------------------------------------------------
# Calculation
# ---------------------------------------------------------------------------
SET_FIRST, SET_LAST = 4, 4 + (RAW_LAST - RAW_FIRST)  # Analisis_Set rows 4..423


def set_rng(L, n_ref):
    return f"Analisis_Set!${L}${SET_FIRST}:INDEX(Analisis_Set!${L}${SET_FIRST}:${L}${SET_LAST},{n_ref})"


def sheet_calculation(wb, isi_rumus=True):
    ws = wb.create_sheet("Calculation")
    C = {}  # alamat sel penting
    title(ws, "Calculation: seluruh perhitungan analisis (rumus hidup, data dari Clean_Data dan Analisis_Set)",
          sub="Pemisah argumen di file ini mengikuti format internal (koma). Fungsi statistik: COUNTIFS, AVERAGEIFS, STDEV.S, CONFIDENCE.T, T.TEST, CORREL, LINEST.")
    r = 3
    r = section(ws, r, "A. Ukuran sampel dan parameter", 10)
    rows = [
        ("n baris data mentah (Raw_Data)", f"=COUNTA(Raw_Data!A{RAW_FIRST}:A{RAW_LAST})", "n_raw"),
        ("n baris duplikat dibuang", f"=COUNTIFS({rr('W')},\"Duplikat - dibuang\")", "n_dup"),
        ("n responden bersih (dipakai)", f"=SUM({rr('X')})", "n_bersih"),
        ("n kasus lengkap (4 dimensi + kepuasan)", f"=SUM({rr('AE')})", "n_lengkap"),
        ("Taraf signifikansi (alpha)", f"={MASTER['taraf']}", "taraf"),
        ("Tingkat kepercayaan", "=1-{taraf}", "conf"),
    ]
    for label, f, key in rows:
        put(ws, f"A{r}", label)
        C[key] = f"$B${r}"
        if isi_rumus:
            put_formula(ws, f"B{r}", f.format(taraf=C.get("taraf", "")), fmt="0.00" if key in ("taraf", "conf") else "0", link=key == "taraf")
        r += 1
    r += 1

    # ---- B. Distribusi item
    r = section(ws, r, "B. Distribusi skor per item (hanya baris dipakai = 1; nilai kosong tidak dihitung)", 17)
    hdr = ["Item", "Dimensi", "Pernyataan", "n", "Skor 1", "Skor 2", "Skor 3", "Skor 4", "Skor 5", "% 1", "% 2", "% 3", "% 4", "% 5", "Mean", "SD", "Top-2-box %"]
    header_row(ws, r, hdr)
    hdr_row = r
    for k, v in zip("EFGHI", [1, 2, 3, 4, 5]):
        ws[f"{k}{r}"].value = v  # header numerik supaya bisa dipakai di rumus
        ws[f"{k}{r}"].number_format = '"Skor "0'
    r += 1
    C["dist_first"] = r
    item_rows = {}
    for it in ITEMS + ["kepuasan_keseluruhan"]:
        L = CLEAN[it]
        dim = next((d for d, its in DIMENSI if it in its), "Keseluruhan")
        put(ws, f"A{r}", it)
        put(ws, f"B{r}", dim)
        put(ws, f"C{r}", PERTANYAAN[it], align=WRAP)
        item_rows[it] = r
        if isi_rumus:
            put_formula(ws, f"D{r}", f"=COUNTIFS({rr(L)},\">=1\",{rr('X')},1)", fmt="0")
            for k in "EFGHI":
                put_formula(ws, f"{k}{r}", f"=COUNTIFS({rr(L)},{k}${hdr_row},{rr('X')},1)", fmt="0")
            for k, src in zip("JKLMN", "EFGHI"):
                put_formula(ws, f"{k}{r}", f"=IFERROR({src}{r}/$D{r},0)", fmt="0.0%")
            put_formula(ws, f"O{r}", f"=IFERROR(AVERAGEIFS({rr(L)},{rr('X')},1),\"\")", fmt="0.00")
            put_formula(ws, f"P{r}", f"=IFERROR(SQRT((SUMPRODUCT(E{r}:I{r},$E${hdr_row}:$I${hdr_row}^2)-D{r}*O{r}^2)/(D{r}-1)),\"\")", fmt="0.00")
            put_formula(ws, f"Q{r}", f"=M{r}+N{r}", fmt="0.0%")
        r += 1
    C["dist_last"] = r - 1
    C["item_rows"] = item_rows
    put_text(ws, f"A{r}", "SD dihitung dari tabel frekuensi: SQRT((Σ n_k·k² − n·mean²)/(n−1)); hasilnya sama dengan STDEV.S pada data bersih.", note=True)
    r += 2

    # ---- C. Dimensi
    r = section(ws, r, "C. Skor dimensi: mean, interval kepercayaan 95 % (CONFIDENCE.T) dan Cronbach's alpha", 12)
    header_row(ws, r, ["Dimensi", "n", "Mean", "SD", "Margin CI", "Batas bawah", "Batas atas", "k item", "Σ var item", "Var total (skor jumlah)", "Cronbach's alpha", "Interpretasi alpha"])
    r += 1
    C["dim_first"] = r
    dim_rows = {}
    alpha_letters = [get_column_letter(i) for i in range(33,45)]
    for idx, (dim, items) in enumerate(DIMENSI):
        L = CLEAN[f"dim_{dim.lower()}"]
        put(ws, f"A{r}", dim)
        dim_rows[dim] = r
        if isi_rumus:
            put_formula(ws, f"B{r}", f"=COUNT({rr(L)})", fmt="0")
            put_formula(ws, f"C{r}", f"=AVERAGE({rr(L)})", fmt="0.00")
            put_formula(ws, f"D{r}", f"=STDEV.S({rr(L)})", fmt="0.00")
            put_formula(ws, f"E{r}", f"=CONFIDENCE.T({C['taraf']},D{r},B{r})", fmt="0.000")
            put_formula(ws, f"F{r}", f"=C{r}-E{r}", fmt="0.00")
            put_formula(ws, f"G{r}", f"=C{r}+E{r}", fmt="0.00")
            put(ws, f"H{r}", 3, fmt="0")
            al = alpha_letters[idx * 3: idx * 3 + 3]
            put_formula(ws, f"I{r}", "=" + "+".join(f"VAR.S({rr(a)})" for a in al), fmt="0.000")
            put_formula(ws, f"J{r}", f"=VAR.S({rr(L)})*H{r}^2", fmt="0.000")
            put_formula(ws, f"K{r}", f"=H{r}/(H{r}-1)*(1-I{r}/J{r})", fmt="0.000")
            put_formula(ws, f"L{r}", f'=IF(K{r}>=0.8,"Baik (≥0,8)",IF(K{r}>=0.7,"Dapat diterima (0,7-0,8)","Rendah (<0,7)"))')
        r += 1
    put(ws, f"A{r}", "Kepuasan keseluruhan")
    dim_rows["Keseluruhan"] = r
    if isi_rumus:
        put_formula(ws, f"B{r}", f"=COUNT({rr('AC')})", fmt="0")
        put_formula(ws, f"C{r}", f"=AVERAGE({rr('AC')})", fmt="0.00")
        put_formula(ws, f"D{r}", f"=STDEV.S({rr('AC')})", fmt="0.00")
        put_formula(ws, f"E{r}", f"=CONFIDENCE.T({C['taraf']},D{r},B{r})", fmt="0.000")
        put_formula(ws, f"F{r}", f"=C{r}-E{r}", fmt="0.00")
        put_formula(ws, f"G{r}", f"=C{r}+E{r}", fmt="0.00")
        put(ws, f"H{r}", "-"); put(ws, f"I{r}", "-"); put(ws, f"J{r}", "-"); put(ws, f"K{r}", "-"); put(ws, f"L{r}", "Satu item, alpha tidak berlaku")
    r += 1
    C["dim_rows"] = dim_rows
    put_text(ws, f"A{r}", "Skor dimensi = rata-rata 3 item (hanya responden yang mengisi ketiganya). Var total skor jumlah = VAR.S(skor dimensi) × k² karena jumlah item = k × rata-rata. Alpha = k/(k−1) × (1 − Σ var item / var total).", note=True)
    r += 2

    # ---- D. Per kanal / kota / segmen
    r = section(ws, r, "D. Ringkasan per kanal, kota, dan segmen (AVERAGEIFS / COUNTIFS pada data bersih)", 14)
    header_row(ws, r, ["Kanal", "n responden", "n kepuasan", "Mean kepuasan", "SD", "Margin CI", "Batas bawah", "Batas atas", "Produk", "Layanan", "Harga", "Pengiriman", "Promotor", "Detraktor", "n NPS", "NPS"])
    r += 1
    C["kanal_first"] = r
    kanal_rows = {}
    for k, hl in zip(KANAL, ["AS", "AT", "AU"]):
        put(ws, f"A{r}", k)
        kanal_rows[k] = r
        if isi_rumus:
            put_formula(ws, f"B{r}", f"=COUNTIFS({rr('C')},A{r},{rr('X')},1)", fmt="0")
            put_formula(ws, f"C{r}", f"=COUNT({rr(hl)})", fmt="0")
            put_formula(ws, f"D{r}", f"=AVERAGE({rr(hl)})", fmt="0.00")
            put_formula(ws, f"E{r}", f"=STDEV.S({rr(hl)})", fmt="0.00")
            put_formula(ws, f"F{r}", f"=CONFIDENCE.T({C['taraf']},E{r},C{r})", fmt="0.000")
            put_formula(ws, f"G{r}", f"=D{r}-F{r}", fmt="0.00")
            put_formula(ws, f"H{r}", f"=D{r}+F{r}", fmt="0.00")
            for L, dim in zip("IJKL", ["Y", "Z", "AA", "AB"]):
                put_formula(ws, f"{L}{r}", f"=IFERROR(AVERAGEIFS({rr(dim)},{rr('C')},$A{r}),\"\")", fmt="0.00")
            put_formula(ws, f"M{r}", f"=COUNTIFS({rr('AD')},\">=\"&{MASTER['promotor']},{rr('C')},$A{r})", fmt="0")
            put_formula(ws, f"N{r}", f"=COUNTIFS({rr('AD')},\"<>\",{rr('AD')},\"<=\"&{MASTER['detraktor']},{rr('C')},$A{r})", fmt="0")
            put_formula(ws, f"O{r}", f"=COUNTIFS({rr('AD')},\"<>\",{rr('AD')},\">=0\",{rr('C')},$A{r})", fmt="0")
            put_formula(ws, f"P{r}", f"=IFERROR((M{r}-N{r})/O{r}*100,\"\")", fmt="0.0")
        r += 1
    C["kanal_rows"] = kanal_rows
    r += 1
    header_row(ws, r, ["Kota", "n responden", "n kepuasan", "Mean kepuasan", "Produk", "Layanan", "Harga", "Pengiriman", "NPS"])
    r += 1
    C["kota_first"] = r
    for k in KOTA:
        put(ws, f"A{r}", k)
        if isi_rumus:
            put_formula(ws, f"B{r}", f"=COUNTIFS({rr('D')},A{r},{rr('X')},1)", fmt="0")
            put_formula(ws, f"C{r}", f"=COUNTIFS({rr('AC')},\">=1\",{rr('D')},A{r})", fmt="0")
            put_formula(ws, f"D{r}", f"=IFERROR(AVERAGEIFS({rr('AC')},{rr('D')},A{r}),\"\")", fmt="0.00")
            for L, dim in zip("EFGH", ["Y", "Z", "AA", "AB"]):
                put_formula(ws, f"{L}{r}", f"=IFERROR(AVERAGEIFS({rr(dim)},{rr('D')},$A{r}),\"\")", fmt="0.00")
            put_formula(ws, f"I{r}", f"=IFERROR((COUNTIFS({rr('AD')},\">=\"&{MASTER['promotor']},{rr('D')},A{r})-COUNTIFS({rr('AD')},\"<>\",{rr('AD')},\"<=\"&{MASTER['detraktor']},{rr('D')},A{r}))/COUNTIFS({rr('AD')},\"<>\",{rr('AD')},\">=0\",{rr('D')},A{r})*100,\"\")", fmt="0.0")
        r += 1
    C["kota_last"] = r - 1
    r += 1
    header_row(ws, r, ["Segmen", "n responden", "n kepuasan", "Mean kepuasan", "Produk", "Layanan", "Harga", "Pengiriman", "NPS"])
    r += 1
    C["segmen_first"] = r
    for k in SEGMEN:
        put(ws, f"A{r}", k)
        if isi_rumus:
            put_formula(ws, f"B{r}", f"=COUNTIFS({rr('E')},A{r},{rr('X')},1)", fmt="0")
            put_formula(ws, f"C{r}", f"=COUNTIFS({rr('AC')},\">=1\",{rr('E')},A{r})", fmt="0")
            put_formula(ws, f"D{r}", f"=IFERROR(AVERAGEIFS({rr('AC')},{rr('E')},A{r}),\"\")", fmt="0.00")
            for L, dim in zip("EFGH", ["Y", "Z", "AA", "AB"]):
                put_formula(ws, f"{L}{r}", f"=IFERROR(AVERAGEIFS({rr(dim)},{rr('E')},$A{r}),\"\")", fmt="0.00")
            put_formula(ws, f"I{r}", f"=IFERROR((COUNTIFS({rr('AD')},\">=\"&{MASTER['promotor']},{rr('E')},A{r})-COUNTIFS({rr('AD')},\"<>\",{rr('AD')},\"<=\"&{MASTER['detraktor']},{rr('E')},A{r}))/COUNTIFS({rr('AD')},\"<>\",{rr('AD')},\">=0\",{rr('E')},A{r})*100,\"\")", fmt="0.0")
        r += 1
    C["segmen_last"] = r - 1
    r += 1

    # ---- E. Uji t Welch dan ANOVA
    r = section(ws, r, "E. Uji beda rata-rata kepuasan keseluruhan antar kanal", 10)
    put_text(ws, f"A{r}", "E1. Uji t Welch (dua kelompok, varians tidak diasumsikan sama): Toko vs Online. T.TEST(array1, array2, tails = 2, type = 3).", bold=True)
    r += 1
    kt, ko = kanal_rows["Toko"], kanal_rows["Online"]
    t_rows = [
        ("n Toko", f"=C{kt}", "0", "n1"),
        ("n Online", f"=C{ko}", "0", "n2"),
        ("Mean Toko", f"=D{kt}", "0.000", "m1"),
        ("Mean Online", f"=D{ko}", "0.000", "m2"),
        ("Varians Toko (VAR.S)", f"=VAR.S({rr('AS')})", "0.000", "v1"),
        ("Varians Online (VAR.S)", f"=VAR.S({rr('AT')})", "0.000", "v2"),
        ("Selisih mean (Toko − Online)", "={m1}-{m2}", "0.000", "diff"),
        ("Galat baku selisih = SQRT(v1/n1 + v2/n2)", "=SQRT({v1}/{n1}+{v2}/{n2})", "0.000", "se"),
        ("Statistik t = selisih / galat baku", "={diff}/{se}", "0.000", "t"),
        ("Derajat bebas Welch–Satterthwaite", "=({v1}/{n1}+{v2}/{n2})^2/(({v1}/{n1})^2/({n1}-1)+({v2}/{n2})^2/({n2}-1))", "0.0", "df"),
        ("t kritis dua sisi = T.INV.2T(alpha, df)", "=T.INV.2T(" + C["taraf"] + ",{df})", "0.000", "tkrit"),
        ("p-value dua sisi = T.TEST(..., 2, 3)", f"=T.TEST({rr('AS')},{rr('AT')},2,3)", "0.0000", "p"),
        ("Cohen's d (ukuran efek, SD gabungan)", "={diff}/SQRT((({n1}-1)*{v1}+({n2}-1)*{v2})/({n1}+{n2}-2))", "0.00", "d"),
        ("Keputusan pada alpha", '=IF({p}<' + C["taraf"] + ',"Berbeda signifikan: H0 (mean sama) ditolak","Tidak berbeda signifikan: H0 tidak ditolak")', None, "keputusan"),
        ("Interpretasi", '="Rata-rata kepuasan Toko "&TEXT({m1},"0.00")&" vs Online "&TEXT({m2},"0.00")&"; selisih "&TEXT({diff},"0.00")&" poin, t = "&TEXT({t},"0.00")&", p = "&TEXT({p},"0.0000")&". "&IF({p}<' + C["taraf"] + ',"H0 mean sama ditolak dengan asumsi uji; p bukan probabilitas bahwa H0 benar.","Bukti belum cukup untuk menolak H0; ini tidak membuktikan kedua mean sama.")&" Ini bukan bukti sebab-akibat."', None, "interp"),
    ]
    refs = {}
    for label, f, fmt, key in t_rows:
        put(ws, f"A{r}", label)
        refs[key] = f"$B${r}"
        C[f"t_{key}"] = f"$B${r}"
        if isi_rumus:
            put_formula(ws, f"B{r}", f.format(**refs), fmt=fmt, link=key in ("n1", "n2", "m1", "m2"))
        r += 1
    r += 1
    put_text(ws, f"A{r}", "E2. ANOVA satu arah tiga kanal (dihitung dengan rumus; versi lengkap tersedia di Data > Data Analysis > Anova: Single Factor).", bold=True)
    r += 1
    kc = kanal_rows["Call Center"]
    a_rows = [
        ("N total (kepuasan terisi)", f"=C{kt}+C{ko}+C{kc}", "0", "N"),
        ("Grand mean", f"=AVERAGE({rr('AC')})", "0.000", "gm"),
        ("SS antar kelompok = Σ n_g (mean_g − grand mean)²", f"=C{kt}*(D{kt}-{{gm}})^2+C{ko}*(D{ko}-{{gm}})^2+C{kc}*(D{kc}-{{gm}})^2", "0.000", "ssb"),
        ("SS dalam kelompok = Σ (n_g − 1) var_g", f"=(C{kt}-1)*E{kt}^2+(C{ko}-1)*E{ko}^2+(C{kc}-1)*E{kc}^2", "0.000", "ssw"),
        ("df antar = k − 1", "=3-1", "0", "df1"),
        ("df dalam = N − k", "={N}-3", "0", "df2"),
        ("MS antar", "={ssb}/{df1}", "0.000", "msb"),
        ("MS dalam", "={ssw}/{df2}", "0.000", "msw"),
        ("F = MS antar / MS dalam", "={msb}/{msw}", "0.000", "F"),
        ("Keputusan", '="F deskriptif; keputusan ANOVA dilakukan manual dengan p-value tepat"', None, "akep"),
    ]
    for label, f, fmt, key in a_rows:
        put(ws, f"A{r}", label)
        refs[key] = f"$B${r}"
        C[f"a_{key}"] = f"$B${r}"
        if isi_rumus:
            put_formula(ws, f"B{r}", f.format(**refs), fmt=fmt, link=key == "fkrit")
        r += 1
    put_text(ws, f"A{r}", "p-value ANOVA tidak dihitung di sini (butuh F.DIST.RT); gunakan Analysis ToolPak untuk tabel ANOVA lengkap. Post-hoc: uji t Welch per pasangan dengan koreksi Bonferroni (alpha/3).", note=True)
    r += 2

    # ---- F. Korelasi dan regresi
    r = section(ws, r, "F. Korelasi dimensi dengan kepuasan keseluruhan (CORREL) dan regresi berganda deskriptif (LINEST)", 10)
    header_row(ws, r, ["Dimensi", "r Pearson dengan kepuasan", "r²", "n kasus lengkap", "Kekuatan hubungan"])
    r += 1
    C["corr_first"] = r
    nL = C["n_lengkap"]
    for dim, L in zip([d for d, _ in DIMENSI], "CDEF"):
        put(ws, f"A{r}", dim)
        if isi_rumus:
            put_formula(ws, f"B{r}", f"=CORREL({set_rng(L, nL)},{set_rng('G', nL)})", fmt="0.000")
            put_formula(ws, f"C{r}", f"=B{r}^2", fmt="0.000")
            put_formula(ws, f"D{r}", f"={nL}", fmt="0")
            put_formula(ws, f"E{r}", f'=IF(ABS(B{r})>=0.7,"Kuat",IF(ABS(B{r})>=0.4,"Sedang","Lemah"))')
        r += 1
    C["corr_last"] = r - 1
    r += 1
    put_text(ws, f"A{r}", "Regresi berganda deskriptif: kepuasan_keseluruhan = b0 + b1·Produk + b2·Layanan + b3·Harga + b4·Pengiriman. LINEST mengembalikan koefisien dalam urutan terbalik (Pengiriman dahulu).", bold=True)
    r += 1
    lin = f"LINEST({set_rng('G', nL)},Analisis_Set!$C${SET_FIRST}:INDEX(Analisis_Set!$C${SET_FIRST}:$F${SET_LAST},{nL},4),1,1)"
    header_row(ws, r, ["Prediktor", "Koefisien b", "Galat baku", "t = b / SE", "t kritis", "Signifikan pada alpha?", "Keterangan"])
    r += 1
    C["reg_first"] = r
    reg = [("Intersep (b0)", 5), ("Produk", 4), ("Layanan", 3), ("Harga", 2), ("Pengiriman", 1)]
    for name, idx in reg:
        put(ws, f"A{r}", name)
        if isi_rumus:
            put_formula(ws, f"B{r}", f"=INDEX({lin},1,{idx})", fmt="0.000")
            put_formula(ws, f"C{r}", f"=INDEX({lin},2,{idx})", fmt="0.000")
            put_formula(ws, f"D{r}", f"=B{r}/C{r}", fmt="0.00")
            put_formula(ws, f"E{r}", f"=T.INV.2T({C['taraf']},{nL}-5)", fmt="0.000")
            put_formula(ws, f"F{r}", f'=IF(ABS(D{r})>E{r},"Ya","Tidak")')
            put(ws, f"G{r}", "Perubahan rata-rata kepuasan per 1 poin skor dimensi, dimensi lain dianggap tetap" if idx != 5 else "Nilai prediksi saat semua dimensi = 0 (di luar rentang data; tidak ditafsirkan)", font=FONT_NOTE)
        r += 1
    C["reg_last"] = r - 1
    stat_rows = [("R² (proporsi variasi yang dijelaskan)", 3, 1, "0.000", "r2"), ("Galat baku estimasi", 3, 2, "0.000", "sey"), ("F regresi", 4, 1, "0.00", "freg"), ("df sisa", 4, 2, "0", "dfres")]
    for label, i, j, fmt, key in stat_rows:
        put(ws, f"A{r}", label)
        C[f"reg_{key}"] = f"$B${r}"
        if isi_rumus:
            put_formula(ws, f"B{r}", f"=INDEX({lin},{i},{j})", fmt=fmt)
        r += 1
    put_text(ws, f"A{r}", "PERINGATAN: prediktor saling berkorelasi (multikolinearitas), variabel respons ordinal 1-5, dan data observasional non-acak. Koefisien hanya menggambarkan asosiasi dalam sampel ini, BUKAN pengaruh sebab-akibat.", note=True)
    r += 2

    # ---- G. NPS
    r = section(ws, r, "G. Net Promoter Score (NPS) keseluruhan = % promotor − % detraktor", 8)
    n_rows = [
        ("Promotor (skor 9-10)", f"=COUNTIFS({rr('AD')},\">=\"&{MASTER['promotor']})", "0", "prom"),
        ("Pasif (skor 7-8)", "={nnps}-{prom}-{det}", "0", "pas"),
        ("Detraktor (skor 0-6)", f"=COUNTIFS({rr('AD')},\"<>\",{rr('AD')},\"<=\"&{MASTER['detraktor']})", "0", "det"),
        ("n responden menjawab NPS", f"=COUNT({rr('AD')})", "0", "nnps"),
        ("% promotor", "={prom}/{nnps}", "0.0%", "pprom"),
        ("% pasif", "={pas}/{nnps}", "0.0%", "ppas"),
        ("% detraktor", "={det}/{nnps}", "0.0%", "pdet"),
        ("NPS (skala −100 s.d. +100)", "=({pprom}-{pdet})*100", "0.0", "nps"),
        ("Mean skor rekomendasi (0-10)", f"=AVERAGE({rr('AD')})", "0.00", "mnps"),
    ]
    # tulis label dulu supaya referensi maju bisa diisi
    start = r
    for i, (label, _, _, key) in enumerate(n_rows):
        put(ws, f"A{start + i}", label)
        refs[key] = f"$B${start + i}"
        C[f"nps_{key}"] = f"$B${start + i}"
    for i, (label, f, fmt, key) in enumerate(n_rows):
        if isi_rumus:
            put_formula(ws, f"B{start + i}", f.format(**refs), fmt=fmt)
    r = start + len(n_rows) + 1

    set_widths(ws, {"A": 46, "B": 18, "C": 40, "D": 12, "E": 10, "F": 10, "G": 10, "H": 10, "I": 10, "J": 10, "K": 10, "L": 12, "M": 10, "N": 10, "O": 10, "P": 10, "Q": 12})
    freeze(ws, "A3")
    print_setup(ws)
    return ws, C


# ---------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------
def sheet_report(wb, C):
    from openpyxl.worksheet.pagebreak import Break
    ws=wb.create_sheet("Report")
    ws.sheet_view.showGridLines=False
    for L in "ABCDEFGHIJKL": ws.column_dimensions[L].width=12
    for r in range(1,84):ws.row_dimensions[r].height=20
    def text(row,content,bold=False,height=28):
        ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=12)
        put_text(ws,f"A{row}",content,bold=bold,wrap=True)
        ws.row_dimensions[row].height=height
    def live(row,formula,height=30):
        ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=12)
        put_formula(ws,f"A{row}",formula,link=True)
        ws[f"A{row}"].alignment=WRAP;ws.row_dimensions[row].height=height
    text(1,"Survei Kepuasan Pelanggan | 2024",True,30)
    text(2,"Proyek portofolio dengan data simulasi • Sumber: survei_raw.csv • Diperbarui "+TANGGAL_BUAT,False,22)
    text(3,"Data pada proyek ini adalah data sintetis yang saya bangkitkan dengan skrip; tidak ada data perusahaan atau pribadi nyata.",False,28)
    overall=C['dim_rows']['Keseluruhan']
    cards=[('Responden bersih',f"=Calculation!{C['n_bersih']}","0"),('Kepuasan / 5',f"=Calculation!C{overall}","0.00"),('NPS / −100–100',f"=Calculation!{C['nps_nps']}","0.00"),('Jawaban NPS',f"=Calculation!{C['nps_nnps']}","0")]
    for i,(label,f,fmt) in enumerate(cards):
        j=i*3+1
        for r in [5,6]:ws.merge_cells(start_row=r,start_column=j,end_row=r,end_column=j+2)
        L=ws.cell(5,j).column_letter
        put(ws,f"{L}5",label,font=FONT_BOLD,fill=FILL_KPI,align=CENTER)
        put(ws,f"{L}6",f,font=FONT_KPI,fmt=fmt,fill=FILL_KPI,align=CENTER)
    ws.row_dimensions[6].height=30
    text(8,"Skor dimensi dan interval kepercayaan 95%",True,24)
    header_row(ws,9,["Dimensi","n valid","Mean / 5","CI bawah","CI atas","Alpha"])
    for r,(name,_) in enumerate(DIMENSI,10):
        cr=C['dim_rows'][name]
        put(ws,f"A{r}",name)
        for L,src in zip("BCDEF","BCFGK"):
            put_formula(ws,f"{L}{r}",f"=Calculation!{src}{cr}",fmt="0" if L=='B' else "0.000",link=True)
    ch=BarChart();ch.type="bar";ch.style=10;ch.title="Mean dimensi (1–5)"
    ch.add_data(Reference(ws,min_col=3,min_row=9,max_row=13),titles_from_data=True)
    ch.set_categories(Reference(ws,min_col=1,min_row=10,max_row=13))
    ch.legend=None;ch.y_axis.scaling.min=0;ch.y_axis.scaling.max=5
    ch.width=13;ch.height=5.4;ws.add_chart(ch,"G8")
    live(15,f'="Kepuasan keseluruhan: "&TEXT(Calculation!C{overall},"0.000")&"; CI 95% "&TEXT(Calculation!F{overall},"0.000")&"–"&TEXT(Calculation!G{overall},"0.000")&"; n = "&Calculation!B{overall}&". Alpha tidak membuktikan validitas kuesioner."',32)
    text(17,"Ringkasan per kanal | pembagi n adalah jawaban kepuasan yang valid",True,22)
    header_row(ws,18,["Kanal","n valid","Mean / 5","CI bawah","CI atas","Produk","Layanan","Harga","Pengiriman","NPS"])
    for r,k in enumerate(KANAL,19):
        cr=C['kanal_rows'][k];put(ws,f"A{r}",k)
        for L,src in zip("BCDEFGHIJ","CDGHIJKLP"):
            put_formula(ws,f"{L}{r}",f"=Calculation!{src}{cr}",fmt="0" if L=='B' else "0.000",link=True)
    text(23,"Uji t Welch dua sisi | kepuasan Toko dibanding Online",True,22)
    for r,(label,key) in enumerate([('Selisih mean','t_diff'),('t Welch','t_t'),('df Welch','t_df'),('p dua sisi','t_p'),("Cohen's d",'t_d')],24):
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=3)
        put(ws,f"A{r}",label);put_formula(ws,f"D{r}",f"=Calculation!{C[key]}",fmt="0.000000",link=True)
    ch2=BarChart();ch2.type="col";ch2.style=10;ch2.title="NPS per kanal"
    ch2.add_data(Reference(ws,min_col=10,min_row=18,max_row=21),titles_from_data=True)
    ch2.set_categories(Reference(ws,min_col=1,min_row=19,max_row=21));ch2.legend=None
    ch2.width=13;ch2.height=4.3;ws.add_chart(ch2,"G23")
    live(30,f'="Keputusan: "&Calculation!{C["t_keputusan"]}&". Perbedaan kanal tidak membuktikan sebab-akibat."',30)
    live(32,f'="NPS = ("&Calculation!{C["nps_prom"]}&" promotor − "&Calculation!{C["nps_det"]}&" detraktor) / "&Calculation!{C["nps_nnps"]}&" jawaban valid × 100; pasif = "&Calculation!{C["nps_pas"]}&"."',30)
    text(34,"Cara membaca: CI adalah rentang berbasis asumsi statistik. Sampel non-acak dan self-report tidak mendukung generalisasi populasi atau klaim sebab-akibat.",False,34)
    text(36,"Pemeriksaan: buka Checks (SEMUA PASS pada solusi), Check_Detail (rekonsiliasi setiap sel), dan VERIFIKASI.md. Data kosong tidak diubah menjadi nol.",False,32)
    text(38,"Halaman 1 / 2 • Angka berasal dari Calculation • Grafik native workbook • Keterbatasan lengkap di halaman berikutnya",False,24)
    text(43,"Asosiasi, rekomendasi, dan batas penggunaan",True,30)
    text(44,"Saya membaca hasil berikut sebagai simulasi proses analisis; pola pembangkit data dirancang dan bukan temuan dari perusahaan nyata.",False,30)
    header_row(ws,46,["Dimensi","r Pearson","Koefisien b","Galat baku","t = b/SE","Signifikan?"])
    for i,(name,_) in enumerate(DIMENSI):
        r=47+i;cr=C['reg_first']+1+i;put(ws,f"A{r}",name)
        put_formula(ws,f"B{r}",f"=Calculation!B{C['corr_first']+i}",fmt="0.000",link=True)
        for L,src in zip("CDEF","BCDF"):
            put_formula(ws,f"{L}{r}",f"=Calculation!{src}{cr}",fmt="0.000" if L!='F' else None,link=True)
    ch3=BarChart();ch3.type="bar";ch3.style=10;ch3.title="Korelasi Pearson"
    ch3.add_data(Reference(ws,min_col=2,min_row=46,max_row=50),titles_from_data=True)
    ch3.set_categories(Reference(ws,min_col=1,min_row=47,max_row=50));ch3.legend=None
    ch3.y_axis.scaling.min=0;ch3.y_axis.scaling.max=1;ch3.width=13;ch3.height=5.2;ws.add_chart(ch3,"G45")
    live(53,f'="Regresi deskriptif: n kasus lengkap = "&Calculation!{C["n_lengkap"]}&"; R² = "&TEXT(Calculation!{C["reg_r2"]},"0.000")&". Koefisien adalah asosiasi bersyarat, bukan pengaruh kausal."',32)
    text(55,"Temuan dan tindak lanjut yang saya usulkan",True,24)
    dimA=f"Calculation!A{C['dim_first']}:A{C['dim_first']+3}";dimC=f"Calculation!C{C['dim_first']}:C{C['dim_first']+3}"
    live(56,f'="Dimensi terendah: "&INDEX({dimA},MATCH(MIN({dimC}),{dimC},0))&" ("&TEXT(MIN({dimC}),"0.000")&"). Saya memeriksa persepsi nilai dan komentar sebelum mengusulkan perubahan harga."',34)
    kt=C['kanal_rows']['Toko'];ko=C['kanal_rows']['Online']
    live(58,f'="Pengiriman: mean Online "&TEXT(Calculation!L{ko},"0.000")&", Toko "&TEXT(Calculation!L{kt},"0.000")&". Saya memeriksa waktu kirim aktual dan bauran pelanggan sebelum menyimpulkan penyebab."',34)
    text(60,"Saya membaca komentar detraktor untuk merumuskan hipotesis keluhan, lalu menilai perubahan melalui desain pembanding dan ukuran operasional yang independen. Seluruh komentar adalah sintetis.",False,34)
    text(62,"Keterbatasan metode dan alat",True,24)
    limits=["Sampel non-acak: CI dan p-value hanya ilustrasi berbasis asumsi responden independen; bias seleksi tidak tercakup dalam CI.","Self-report dan skala ordinal: rata-rata memperlakukan jarak skor sebagai setara. Tidak ada hubungan sebab-akibat dari perbandingan kanal, CORREL, atau LINEST.","Missing values tidak diimputasi. Skor dan alpha memakai tiga item lengkap per dimensi; regresi memakai kasus lengkap seluruh prediktor dan respons. Pembagi berbeda antar analisis.","Alpha mengukur konsistensi internal pada sampel ini; nilai tinggi bukan bukti validitas atau unidimensionalitas. Prediktor regresi saling berkorelasi.","PivotTable, Slicer, Power Query, dan Analysis ToolPak belum tertanam. Interaksi di Excel dan impor Google Sheets masih perlu diperiksa manual."]
    for r,t in zip([63,65,67,69,71],limits):text(r,t,False,34)
    text(74,"Halaman 2 / 2 • Data simulasi • Definisi, instruksi, dan asumsi tersedia pada README serta docs/analisis_statistik.md",False,24)
    freeze(ws,"A8");print_setup(ws,landscape=True,area="A1:L74")
    ws.page_setup.fitToHeight=2
    ws.row_breaks.append(Break(id=42))
    return ws


# ---------------------------------------------------------------------------
# Checks
# ---------------------------------------------------------------------------
def sheet_checks(wb, C, ref, raw, catatan, solusi=True):
    from analysis_support import inputs
    _, clean, quality, st = inputs()
    reference = wb.create_sheet("Reference")
    title(reference, "Input pembanding: salinan CSV referensi, bukan hasil rumus")
    write_df(reference, clean, start_row=3)
    freeze(reference, "B4")
    detail = wb.create_sheet("Check_Detail")
    title(detail, "Rekonsiliasi setiap sel referensi dengan Clean_Data (0 = cocok)")
    headers = ["respondent_id", "baris Clean_Data"] + list(clean.columns) + ["jumlah beda", "status"]
    header_row(detail, 3, headers)
    mapping = {h:L for L,h,_,_,_ in CLEAN_COLS}
    end_col = detail.cell(3,len(headers)-2).column_letter
    diff_col = detail.cell(3,len(headers)-1).column_letter
    status_col = detail.cell(3,len(headers)).column_letter
    for r in range(4,len(clean)+4):
        put_formula(detail, f"A{r}", f"=Reference!A{r}", link=True)
        put_formula(detail, f"B{r}", f'=IFERROR(MATCH(A{r},Clean_Data!A$2:A${RAW_LAST},0)+1,"")')
        for j,c in enumerate(clean.columns,3):
            L = detail.cell(3,j).column_letter
            E = f"Reference!{reference.cell(3,j-2).column_letter}{r}"
            V = f"INDEX(Clean_Data!{mapping[c]}$1:{mapping[c]}${RAW_LAST},$B{r})"
            if c=="tanggal_survei": V=f'TEXT({V},"yyyy-mm-dd")'
            f=f'=IF({E}="",IF({V}="",0,1),IF(COUNT({E})=1,IF(COUNT({V})=1,IF(ROUND({V}-{E},12)=0,0,1),1),IF(AND(LEN({V})=LEN({E}),SUBSTITUTE({V},{E},"")=""),0,1)))'
            f = f'=IF(COUNTA(Clean_Data!$X$2:$X${RAW_LAST})<{len(raw)},"",{f[1:]})'
            put_formula(detail,f"{L}{r}",f,fmt="0",link=True)
        put_formula(detail,f"{diff_col}{r}",f"=SUM(C{r}:{end_col}{r})")
        put_formula(detail,f"{status_col}{r}",f'=IF(COUNTA(Clean_Data!$X$2:$X${RAW_LAST})<{len(raw)},"BELUM DIISI",IF({diff_col}{r}=0,"PASS","FAIL"))')
    freeze(detail,"C4")
    set_widths(detail,{"A":15,"B":15})
    ws = wb.create_sheet("Checks")
    title(ws,"Checks: CSV, hasil pembersihan per sel, dan statistik independen")
    put_text(ws,"A2","Pembanding berasal dari CSV dan scipy. Selisih dihitung pada 10 desimal; latihan berstatus BELUM DIISI sampai rumus lengkap.",note=True)
    header_row(ws,3,["No","Pemeriksaan","Nilai workbook","Nilai pembanding","Selisih","Status"])
    checks=[("Baris mentah",f"=Calculation!{C['n_raw']}",len(raw)),("Responden bersih",f"=Calculation!{C['n_bersih']}",len(clean)),("Duplikat dibuang",f"=Calculation!{C['n_dup']}",6), ("Sel referensi berbeda",f"=SUM(Check_Detail!{diff_col}4:{diff_col}{len(clean)+3})",0), ("Sel pembersih terisi",f"=COUNTA(Clean_Data!A2:AV{RAW_LAST})",len(raw)*48)]
    keys=list(quality["csv_akhir"])
    for j,k in enumerate(keys,49):
        L=wb["Clean_Data"].cell(1,j).column_letter
        checks.append((f"CSV akhir: {k}",f"=SUM(Clean_Data!{L}2:{L}{RAW_LAST})",quality["csv_akhir"][k]))
    for name,v in st["dimensions"].items():
        r=C["dim_rows"][name]
        for col,key in [("B","n"),("C","mean"),("D","sd"),("E","margin"),("F","ci_low"),("G","ci_high"),("K","alpha")]:
            checks.append((f"{name}: {key}",f"=Calculation!{col}{r}",v[key]))
    for col,key in [("B","n"),("C","mean"),("D","sd"),("E","margin"),("F","ci_low"),("G","ci_high")]:
        checks.append((f"Keseluruhan: {key}",f"=Calculation!{col}{C['dim_rows']['Keseluruhan']}",st["overall"][key]))
    for key,cell in [("n","nnps"),("promotor","prom"),("pasif","pas"),("detraktor","det"),("value","nps")]:
        checks.append((f"NPS: {key}",f"=Calculation!{C['nps_'+cell]}",st["nps"][key]))
    for key,cell in [("n_toko","n1"),("n_online","n2"),("mean_toko","m1"),("mean_online","m2"),("diff","diff"),("t","t"),("df","df"),("p","p"),("cohen_d","d")]:
        checks.append((f"Welch: {key}",f"=Calculation!{C['t_'+cell]}",st["welch"][key]))
    for i,name in enumerate(DIMENSI):
        checks.append((f"CORREL {name[0]}",f"=Calculation!B{C['corr_first']+i}",st["correlation"][name[0]]))
    checks.append(("Kasus lengkap regresi",f"=Calculation!{C['n_lengkap']}",st["regression"]["n"]))
    for i,value in enumerate(st["regression"]["beta"]):
        checks.append((f"LINEST koefisien {i}",f"=Calculation!B{C['reg_first']+i}",value))
    for i,value in enumerate(st["regression"]["se"]):
        checks.append((f"LINEST galat baku {i}",f"=Calculation!C{C['reg_first']+i}",value))
    for key in ["r2","sey","F","dfres"]:
        checks.append((f"LINEST {key}",f"=Calculation!{C['reg_'+('freg' if key=='F' else key)]}",st["regression"][key]))
    for i,(label,f,value) in enumerate(checks,4):
        put(ws,f"A{i}",i-3);put(ws,f"B{i}",label)
        if solusi: put_formula(ws,f"C{i}",f,fmt="0.0000000000",link=True)
        else: put(ws,f"C{i}",None,fill=FILL_INPUT)
        put_input(ws,f"D{i}",value,fmt="0.0000000000")
        put_formula(ws,f"E{i}",f'=IF(C{i}="","",ROUND(C{i}-D{i},10))',fmt="0.0000000000")
        put_formula(ws,f"F{i}",f'=IF(C{i}="","BELUM DIISI",IF(E{i}=0,"PASS","FAIL"))')
    last=len(checks)+3
    put(ws,f"B{last+1}","Hasil seluruh pemeriksaan",font=FONT_BOLD)
    put_formula(ws,f"F{last+1}",f'=IF(COUNTIFS(F4:F{last},"BELUM DIISI")>0,"BELUM DIISI",IF(COUNTIFS(F4:F{last},"FAIL")=0,"SEMUA PASS","ADA FAIL"))')
    for label,fill in [("PASS",FILL_PASS),("FAIL",FILL_FAIL)]:
        ws.conditional_formatting.add(f"F4:F{last}",CellIsRule(operator="equal",formula=['"'+label+'"'],fill=fill))
    freeze(ws,"C4");set_widths(ws,{"A":6,"B":40,"C":23,"D":23,"E":20,"F":20})
    print_setup(ws)
    return ws


# ---------------------------------------------------------------------------
# Readme
# ---------------------------------------------------------------------------
def sheet_readme(wb, solusi=True):
    ws = wb.create_sheet("Readme", 0)
    ws.sheet_view.showGridLines = False
    title(ws, "Survei Kepuasan Pelanggan 2024: " + ("workbook solusi" if solusi else "workbook latihan"))
    lines = [
        ("PROYEK PORTOFOLIO DENGAN DATA SIMULASI", True),
        ("Data pada proyek ini adalah data sintetis yang saya bangkitkan dengan skrip; tidak ada data perusahaan atau pribadi nyata.", False),
        ("", False),
        ("Tujuan", True),
        ("Membersihkan data survei mentah yang bermasalah, lalu menghitung distribusi skor, skor dimensi dengan interval kepercayaan, perbandingan antar kanal (uji t Welch), korelasi dan regresi deskriptif, NPS, dan reliabilitas (Cronbach's alpha), seluruhnya dengan rumus Excel yang bisa ditelusuri.", False),
        ("", False),
        ("Sumber data", True),
        ("data/raw/survei_raw.csv (420 baris, 414 responden unik) dibangkitkan oleh tools/generate_data.py dengan seed tetap; masalah kualitas data disisipkan dengan sengaja dan didokumentasikan di docs/masalah_kualitas_data.md.", False),
        (f"Tanggal pembuatan workbook: {TANGGAL_BUAT}. Dibangun dengan tools/build_workbook.py (openpyxl), diverifikasi dengan LibreOffice (lihat VERIFIKASI.md).", False),
        ("", False),
        ("Daftar sheet (urutan membaca)", True),
        ("1. Readme: halaman ini.", False),
        ("2. Codebook: definisi tiap variabel, pertanyaan lengkap, skala, dan penanganan nilai tidak valid.", False),
        ("3. Raw_Data: data mentah apa adanya (termasuk duplikat, teks 'N/A', angka berspasi, tanggal campuran).", False),
        ("4. Master: daftar nilai standar (kanal, kota, segmen, pemetaan item ke dimensi) dan parameter pembersihan/analisis (sel biru).", False),
        ("5. Clean_Data: kolom berumus yang membersihkan tiap sel (TRIM, VALUE, IF rentang, INDEX/MATCH, COUNTIFS untuk duplikat) serta skor dimensi.", False),
        ("6. Analisis_Set: kasus lengkap ditumpuk rapat (INDEX/MATCH) agar CORREL dan LINEST bisa dipakai tanpa fungsi FILTER.", False),
        ("7. Calculation: seluruh perhitungan (A-G) dengan rumus hidup.", False),
        ("8. Report: dashboard ringkasan, grafik, temuan, rekomendasi, dan keterbatasan.", False),
        ("9. Checks: rekonsiliasi otomatis workbook vs CSV/skrip (PASS/FAIL).", False),
        ("", False),
        ("Konvensi warna", True),
        ("Biru dengan latar kuning muda = sel input/parameter yang boleh diubah; hitam = rumus; hijau = tautan ke sheet lain.", False),
        ("", False),
        ("Alur input → proses → output", True),
        ("Input: Raw_Data + parameter di Master → Proses: Clean_Data → Analisis_Set → Calculation → Output: Report, diperiksa oleh Checks.", False),
        ("Mengubah parameter di Master (mis. rentang usia valid atau alpha) akan memperbarui seluruh hasil secara otomatis.", False),
        ("", False),
        ("Kompatibilitas", True),
        ("Rumus memakai fungsi Excel 2016+; recalc diperiksa pada versi LibreOffice yang dicatat di VERIFIKASI.md (tanpa XLOOKUP/FILTER/dynamic array). Fungsi STDEV.S, VAR.S, CONFIDENCE.T, T.TEST dan T.INV.2T tersedia pada Excel 2016+.", False),
        ("Keterbatasan metode: sampel non-acak, jawaban self-report, dan tidak ada hubungan sebab-akibat (lihat bagian 8 di Report).", False),
    ]
    if not solusi:
        lines.insert(3, ("Isi kolom dan sel yang masih kosong sesuai sheet Instruksi; bandingkan dengan workbook solusi setelah selesai.", False))
    for i, (t, b) in enumerate(lines, start=3):
        c = ws.cell(row=i, column=1, value=t)
        c.font = FONT_BOLD if b else FONT
        c.alignment = WRAP
    ws.column_dimensions["A"].width = 130
    print_setup(ws, landscape=False)
    return ws


def sheet_instruksi(wb):
    ws = wb.create_sheet("Instruksi", 1)
    ws.sheet_view.showGridLines = False
    title(ws, "Instruksi latihan: dari data mentah ke laporan survei")
    tugas = [
        ("Tahap 1: Kenali data", [
            "Buka Raw_Data. Dengan Conditional Formatting (Home > Conditional Formatting > Highlight Cells Rules > Text that Contains) tandai sel 'N/A' dan '-'.",
            "Hitung di sel kosong: jumlah baris (COUNTA), jumlah sel kosong di blok skor (COUNTBLANK), jumlah nilai <1 dan >5 (COUNTIFS).",
            "Temukan duplikat respondent_id dengan COUNTIFS; catat 6 ID yang terduplikasi.",
        ]),
        ("Tahap 2: Bersihkan di Clean_Data (kolom sudah diberi judul, isi rumusnya)", [
            "Kolom tanggal: ubah teks 'YYYY-MM-DD' dan 'DD/MM/YYYY' menjadi tanggal sungguhan dengan IF(MID(x;5;1)=\"-\"; DATE(...); DATE(...)).",
            "Kolom kanal: TRIM lalu INDEX/MATCH ke daftar standar di Master (MATCH tidak peka huruf besar-kecil).",
            "Kolom usia dan skor: IFERROR(IF(AND(VALUE(TRIM(x))>=min; VALUE(TRIM(x))<=max); VALUE(TRIM(x)); \"\"); \"\") dengan min/max dari Master.",
            "Kolom dipakai: 1 jika COUNTIFS kumulatif respondent_id = 1 (baris pertama), 0 untuk duplikat berikutnya.",
            "Kolom dim_*: AVERAGE 3 item hanya jika COUNT = 3 dan dipakai = 1.",
        ]),
        ("Tahap 3: Analisis di Calculation", [
            "Distribusi per item: COUNTIFS untuk skor 1-5 dengan kriteria tambahan dipakai = 1; persen = count/n; mean dengan AVERAGEIFS.",
            "Skor dimensi: AVERAGE, STDEV.S, margin = CONFIDENCE.T(alpha; SD; n), batas bawah/atas.",
            "Cronbach's alpha: k/(k−1) × (1 − Σ VAR.S(item) / VAR.S(jumlah item)).",
            "Per kanal: AVERAGEIFS/COUNTIFS; uji t Welch Toko vs Online: T.TEST(array_toko; array_online; 2; 3) ditambah hitungan manual t dan df.",
            "Korelasi: CORREL dimensi vs kepuasan pada Analisis_Set; regresi: INDEX(LINEST(...);baris;kolom).",
            "NPS: (COUNTIFS(nps;\">=9\") − COUNTIFS(nps;\"<=6\")) / COUNT(nps) × 100.",
        ]),
        ("Tahap 4: Report dan Checks", [
            "Buat tabel ringkas yang menautkan ke Calculation, lalu Insert > Chart (bar untuk skor dimensi, 100 % stacked bar untuk distribusi, column untuk kanal).",
            "Tulis 5 temuan dan 3 rekomendasi; wajib ada bagian keterbatasan: sampel non-acak, self-report, tidak ada sebab-akibat.",
            "Di Checks, isi kolom 'Nilai workbook' dengan rumus dan bandingkan dengan nilai pembanding; status PASS/FAIL dengan IF(ABS(selisih)<0,0001; ...).",
        ]),
    ]
    r = 3
    put_text(ws, f"A{r}", "Catatan pemisah argumen: contoh di sini memakai titik koma (;) sesuai Excel regional Indonesia; Excel regional Inggris memakai koma (,).", note=True)
    r += 2
    for judul, langkah in tugas:
        put_text(ws, f"A{r}", judul, bold=True); r += 1
        for i, l in enumerate(langkah, start=1):
            put(ws, f"A{r}", f"{i}. {l}", align=WRAP, border=False); ws.row_dimensions[r].height = 30; r += 1
        r += 1
    ws.column_dimensions["A"].width = 130
    print_setup(ws, landscape=False)
    return ws


def build(solusi: bool):
    raw = muat_raw()
    ref = pd.read_csv(REF_CSV)
    import json
    catatan = json.loads((BASE / "data" / "processed" / "catatan_masalah_disengaja.json").read_text(encoding="utf-8"))
    wb = Workbook()
    wb.remove(wb.active)
    sheet_readme(wb, solusi)
    if not solusi:
        sheet_instruksi(wb)
    sheet_codebook(wb)
    sheet_raw(wb, raw)
    sheet_master(wb)
    sheet_clean(wb, isi_rumus=solusi)
    sheet_analisis_set(wb, isi_rumus=solusi)
    ws_calc, C = sheet_calculation(wb, isi_rumus=solusi)
    if solusi:
        sheet_report(wb, C)
    else:
        ws = wb.create_sheet("Report")
        title(ws, "Report: susun dashboard sesuai Instruksi")
    sheet_checks(wb, C, ref, raw, catatan, solusi)
    if not solusi:
        # A complete worked first row makes fill-down and absolute references reviewable.
        for L,h,fn,fmt,link in CLEAN_COLS:
            put_formula(wb["Clean_Data"], f"{L}2", fn(2), fmt=fmt, link=link)
        for ws in [wb["Clean_Data"], wb["Calculation"], wb["Analisis_Set"]]:
            for row in ws:
                for cell in row:
                    if cell.value is None: cell.fill = FILL_INPUT
    for ws in wb:
        if ws.freeze_panes is None: freeze(ws, "A3")
    out = OUT_SOLUSI if solusi else OUT_LATIHAN
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    print("tersimpan:", out)


if __name__ == "__main__":
    build(True)
    build(False)
