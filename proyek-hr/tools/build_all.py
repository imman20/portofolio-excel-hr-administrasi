#!/usr/bin/env python3
"""Membangun ulang seluruh file .xlsx portofolio, menghitung ulang rumus dengan LibreOffice,
memverifikasi tidak ada error, dan membuat screenshot PNG tiap sheet utama.

Pemakaian:
    python tools/build_all.py            # bangun semua proyek
    python tools/build_all.py 01 04      # bangun proyek tertentu saja
    python tools/build_all.py --tanpa-screenshot

Kebutuhan: Python 3.10+, openpyxl 3.1, LibreOffice (soffice), pdftoppm (poppler-utils).
"""

from __future__ import annotations

import importlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))

PROYEK = {
    "01": ("01_database-karyawan", "p01", "01_database_karyawan.xlsx"),
    "02": ("02_rekap-absensi-bulanan", "p02", "02_rekap_absensi_bulanan.xlsx"),
    "03": ("03_administrasi-surat-dan-atk", "p03", "03_administrasi_surat_dan_atk.xlsx"),
    "04": ("04_penggajian-dan-slip-gaji", "p04", "04_penggajian_dan_slip_gaji.xlsx"),
    "05": ("05_pengelolaan-cuti", "p05", "05_pengelolaan_cuti.xlsx"),
    "06": ("06_rekrutmen-tracker", "p06", "06_rekrutmen_tracker.xlsx"),
    "07": ("07_penilaian-kinerja-kpi", "p07", "07_penilaian_kinerja_kpi.xlsx"),
    "08": ("08_hr-dashboard-analitik", "p08", "08_hr_dashboard_analitik.xlsx"),
    "09": ("09_analisis-survei-engagement-statistik", "p09", "09_analisis_survei_engagement.xlsx"),
    "10": ("10_perencanaan-tenaga-kerja-dan-otomasi", "p10", "10_perencanaan_tenaga_kerja.xlsx"),
}

ERRORS = ["#VALUE!", "#DIV/0!", "#REF!", "#NAME?", "#NULL!", "#NUM!", "#N/A"]

MACRO = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">
    Sub HitungUlangSimpan()
      ThisComponent.calculateAll()
      ThisComponent.store()
      ThisComponent.close(True)
    End Sub
</script:module>"""


def _env():
    env = os.environ.copy()
    env["SAL_USE_VCLPLUGIN"] = "svp"
    return env


def _profil(tmp: Path) -> str:
    url = tmp.as_uri()
    subprocess.run(["soffice", "--headless", "--terminate_after_init", f"-env:UserInstallation={url}"],
                   capture_output=True, env=_env(), timeout=120)
    d = tmp / "user" / "basic" / "Standard"
    d.mkdir(parents=True, exist_ok=True)
    (d / "Module1.xba").write_text(MACRO)
    return url


def hitung_ulang(path: Path, timeout=180):
    """Buka file di LibreOffice headless, hitung semua rumus, simpan kembali di tempat."""
    with tempfile.TemporaryDirectory(prefix="lo-prof-") as tmp:
        url = _profil(Path(tmp))
        cmd = ["timeout", str(timeout), "soffice", "--headless", "--norestore",
               f"-env:UserInstallation={url}",
               "vnd.sun.star.script:Standard.Module1.HitungUlangSimpan?language=Basic&location=application",
               str(path.resolve())]
        r = subprocess.run(cmd, capture_output=True, text=True, env=_env(), timeout=timeout + 30)
        if r.returncode != 0:
            raise RuntimeError(f"LibreOffice gagal menghitung ulang {path.name}: {r.stderr.strip()}")


def pasang_gaya_tabel(path: Path, gaya="TableStyleMedium2"):
    """LibreOffice menyimpan tabel tanpa tableStyleInfo; tambahkan kembali gaya tabel ke XML."""
    import re
    import zipfile

    tag = (f'<tableStyleInfo name="{gaya}" showFirstColumn="0" showLastColumn="0" '
           f'showRowStripes="1" showColumnStripes="0"/>')
    tmp = path.with_suffix(".tmp.xlsx")
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("xl/tables/") and b"tableStyleInfo" not in data:
                data = re.sub(rb"</table>\s*$", tag.encode() + b"</table>", data)
            zout.writestr(item, data)
    tmp.replace(path)


def pangkas_png(path: Path, margin=12):
    """Pangkas area putih di sekeliling screenshot agar ringkas."""
    from PIL import Image, ImageChops

    im = Image.open(path).convert("RGB")
    bg = Image.new("RGB", im.size, (255, 255, 255))
    bbox = ImageChops.difference(im, bg).getbbox()
    if bbox:
        l, t, r, b = bbox
        im = im.crop((max(0, l - margin), max(0, t - margin), min(im.width, r + margin), min(im.height, b + margin)))
    im.save(path, optimize=True)


def verifikasi(path: Path):
    """Hitung jumlah rumus dan jumlah sel error setelah recalc."""
    from openpyxl import load_workbook

    wf = load_workbook(path)
    wv = load_workbook(path, data_only=True)
    jumlah_rumus = 0
    error = {}
    rumus_per_sheet = {}
    for ws in wf.worksheets:
        n = 0
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    n += 1
        rumus_per_sheet[ws.title] = n
        jumlah_rumus += n
        for row in wv[ws.title].iter_rows():
            for c in row:
                if isinstance(c.value, str) and any(e in c.value for e in ERRORS):
                    error.setdefault(c.value.split("!")[0] if False else c.value, []).append(f"{ws.title}!{c.coordinate}")
    return {"rumus": jumlah_rumus, "per_sheet": rumus_per_sheet, "error": error,
            "sheet": wf.sheetnames}


def screenshot(modul, nama_sheet: str, keluaran: Path, area=None, lebar_px=1600):
    """Bangun ulang workbook dengan hanya satu sheet terlihat, konversi ke PDF lalu PNG (halaman 1)."""
    wb = modul.build()
    for ws in wb.worksheets:
        if ws.title != nama_sheet:
            ws.sheet_state = "hidden"
    wb.active = wb.sheetnames.index(nama_sheet)
    target = wb[nama_sheet]
    target.sheet_view.tabSelected = True
    if area:
        target.print_area = area
    target.page_setup.fitToWidth = 1
    target.page_setup.fitToHeight = 1
    target.print_options.gridLines = bool(target.sheet_view.showGridLines)
    target.print_options.horizontalCentered = False
    target.page_margins.left = target.page_margins.right = 0.3
    target.page_margins.top = target.page_margins.bottom = 0.3
    with tempfile.TemporaryDirectory(prefix="shot-") as tmp:
        tmp = Path(tmp)
        f = tmp / "s.xlsx"
        wb.save(f)
        hitung_ulang(f)
        subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(tmp), str(f)],
                       capture_output=True, env=_env(), timeout=180)
        pdf = tmp / "s.pdf"
        if not pdf.exists():
            raise RuntimeError(f"PDF tidak terbentuk untuk sheet {nama_sheet}")
        subprocess.run(["pdftoppm", "-png", "-r", "130", "-f", "1", "-l", "1", "-scale-to-x", str(lebar_px),
                        "-scale-to-y", "-1", str(pdf), str(tmp / "p")], capture_output=True, timeout=120)
        png = sorted(tmp.glob("p*.png"))[0]
        keluaran.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy(png, keluaran)
    pangkas_png(keluaran)


def bangun(kode: str, dengan_screenshot=True):
    folder, modul_nama, nama_file = PROYEK[kode]
    modul = importlib.import_module(modul_nama)
    t0 = time.time()
    wb = modul.build()
    path = ROOT / folder / nama_file
    wb.save(path)
    hitung_ulang(path)
    pasang_gaya_tabel(path)
    hasil = verifikasi(path)
    hasil["file"] = f"{folder}/{nama_file}"
    hasil["ukuran_kb"] = round(path.stat().st_size / 1024, 1)
    if dengan_screenshot:
        for spec in getattr(modul, "SCREENSHOT", []):
            sheet, png = spec[0], spec[1]
            area = spec[2] if len(spec) > 2 else None
            screenshot(modul, sheet, ROOT / folder / "assets" / png, area=area)
    hasil["detik"] = round(time.time() - t0, 1)
    status = "OK" if not hasil["error"] else "ERROR"
    print(f"[{kode}] {nama_file}: {hasil['rumus']} rumus, error={sum(len(v) for v in hasil['error'].values())} "
          f"({status}), {hasil['detik']} dtk")
    return hasil


def tulis_log(hasil_semua):
    baris = ["# Log Verifikasi Workbook", "",
             "Dibuat otomatis oleh `tools/build_all.py`. Setiap file dibangun dengan openpyxl, "
             "dihitung ulang (recalculate) dengan LibreOffice headless, lalu seluruh sel diperiksa "
             "terhadap nilai error (#REF!, #NAME?, #DIV/0!, #VALUE!, #N/A, #NUM!, #NULL!).", "",
             "| Proyek | File | Sheet | Jumlah rumus | Sel error | Ukuran (KB) | Hasil |",
             "|---|---|---|---|---|---|---|"]
    total_rumus = 0
    total_err = 0
    for kode, h in hasil_semua.items():
        n_err = sum(len(v) for v in h["error"].values())
        total_rumus += h["rumus"]
        total_err += n_err
        baris.append(f"| {kode} | `{h['file']}` | {len(h['sheet'])} | {h['rumus']} | {n_err} | {h['ukuran_kb']} | "
                     f"{'LOLOS' if n_err == 0 else 'GAGAL'} |")
    baris += ["", f"**Total rumus hidup: {total_rumus}. Total sel error: {total_err}.**", "",
              "## Rincian rumus per sheet", ""]
    for kode, h in hasil_semua.items():
        baris.append(f"### Proyek {kode}: `{h['file']}`")
        baris.append("")
        baris.append("| Sheet | Jumlah rumus |")
        baris.append("|---|---|")
        for s, n in h["per_sheet"].items():
            baris.append(f"| {s} | {n} |")
        if h["error"]:
            baris.append("")
            baris.append("Sel error: " + json.dumps(h["error"], ensure_ascii=False))
        baris.append("")
    baris += ["## Catatan metode", "",
              "- Fungsi pasca-2007 (IFS, RANK.EQ, STDEV.S, T.TEST, F.DIST.RT, FORECAST.LINEAR, dsb.) ditulis "
              "dengan awalan internal `_xlfn.` sesuai format file Excel; di Excel awalan ini tidak terlihat.",
              "- Fungsi array dinamis Excel 365 (XLOOKUP, FILTER, UNIQUE, SORT) tidak dipakai di sel hidup agar "
              "file tetap terbuka sempurna di Excel 2016/2019 dan LibreOffice; alternatif INDEX/MATCH dipakai "
              "dan versi 365 dijelaskan di tutorial.",
              "- Screenshot di folder `assets/` dihasilkan dari file yang sama lewat konversi PDF -> PNG LibreOffice.",
              f"- Waktu verifikasi: {time.strftime('%Y-%m-%d %H:%M')}."]
    (ROOT / "tools" / "verifikasi.md").write_text("\n".join(baris) + "\n", encoding="utf-8")


def main(argv):
    dengan_screenshot = "--tanpa-screenshot" not in argv
    kode_list = [a for a in argv if a in PROYEK] or list(PROYEK)
    hasil = {}
    for kode in kode_list:
        hasil[kode] = bangun(kode, dengan_screenshot)
    if len(kode_list) == len(PROYEK):
        tulis_log(hasil)
    else:
        (ROOT / "tools" / "verifikasi_sebagian.json").write_text(json.dumps(hasil, indent=1, ensure_ascii=False))
    gagal = [k for k, h in hasil.items() if h["error"]]
    if gagal:
        print("GAGAL verifikasi:", gagal)
        return 1
    print("Semua file terverifikasi tanpa error.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
