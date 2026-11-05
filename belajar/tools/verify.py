"""
Verifikasi satu modul belajar.

Pemakaian (dari folder belajar/):
    python tools/verify.py 01_dasar-excel [--no-recalc]

Langkah:
 1. Menjalankan recalc LibreOffice headless (skrip recalc.py dari skill xlsx, atau
    konversi soffice bila skrip tidak ditemukan) pada setiap .xlsx di latihan/.
 2. Menghitung jumlah rumus dan jumlah sel error (#REF!, #NAME?, #DIV/0!, #VALUE!, #N/A, #NUM!).
 3. Menjalankan fungsi checks() dari tools/build.py modul (rekonsiliasi pandas vs sel hasil recalc).
 4. Menulis VERIFIKASI.md di folder modul.
"""
from __future__ import annotations

import datetime as dt
import importlib.util
import json
import os
import subprocess
import sys
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

ERRORS = ["#VALUE!", "#DIV/0!", "#REF!", "#NAME?", "#NULL!", "#NUM!", "#N/A"]

def cari_recalc() -> Path | None:
    """Skrip recalc eksternal (opsional) lewat variabel lingkungan RECALC_PY."""
    c = os.environ.get("RECALC_PY", "")
    return Path(c) if c and Path(c).exists() else None


def recalc_soffice(path: Path) -> dict:
    """Fallback: LibreOffice headless membuka file, menghitung semua rumus (sel tanpa nilai
    cache dihitung saat dimuat), lalu menyimpan ulang sebagai .xlsx di folder sementara;
    hasilnya menimpa file asal."""
    import shutil
    import tempfile
    with tempfile.TemporaryDirectory() as td:
        prof = Path(td) / "profil"
        cmd = ["soffice", "--headless", "--norestore", f"-env:UserInstallation={prof.as_uri()}",
               "--convert-to", "xlsx", "--outdir", td, str(path)]
        try:
            res = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
        except subprocess.TimeoutExpired:
            return {"error": "soffice timeout"}
        out = Path(td) / path.name
        if res.returncode != 0 or not out.exists():
            return {"error": f"soffice gagal: {res.stderr[:300]}"}
        shutil.copyfile(out, path)
    return {"status": "success (soffice convert)"}


def recalc_file(path: Path) -> dict:
    script = cari_recalc()
    if script is None:
        return recalc_soffice(path)
    res = subprocess.run([sys.executable, str(script), str(path), "120"], capture_output=True, text=True)
    out = res.stdout.strip()
    try:
        return json.loads(out[out.index("{"):])
    except Exception:
        return {"error": f"output recalc tidak terbaca: {out[:300]} {res.stderr[:300]}"}


def hitung_rumus_dan_error(path: Path) -> tuple[int, dict[str, list[str]], list[str]]:
    wb_f = load_workbook(path, data_only=False)
    n_formula = 0
    for ws in wb_f.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, ArrayFormula):
                    v = v.text
                if isinstance(v, str) and v.startswith("="):
                    n_formula += 1
    wb_f.close()
    wb_v = load_workbook(path, data_only=True)
    errs: dict[str, list[str]] = {e: [] for e in ERRORS}
    sheets = wb_v.sheetnames
    for ws in wb_v.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("#"):
                    for e in ERRORS:
                        if e in c.value:
                            errs[e].append(f"{ws.title}!{c.coordinate}")
                            break
    wb_v.close()
    return n_formula, {k: v for k, v in errs.items() if v}, sheets


def muat_build(mod_dir: Path):
    spec = importlib.util.spec_from_file_location("build_mod", mod_dir / "tools" / "build.py")
    m = importlib.util.module_from_spec(spec)
    sys.modules["build_mod"] = m
    spec.loader.exec_module(m)
    return m


def nilai_sel(path: Path, sheet: str, cell: str):
    wb = load_workbook(path, data_only=True)
    v = wb[sheet][cell].value
    wb.close()
    return v


def fmt(v) -> str:
    if isinstance(v, float):
        return f"{v:,.4f}".rstrip("0").rstrip(".")
    return str(v)


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    mod_dir = Path(sys.argv[1]).resolve()
    do_recalc = "--no-recalc" not in sys.argv
    lat = mod_dir / "latihan"
    files = sorted(lat.glob("*.xlsx"))
    hasil_file = []
    total_err = 0
    for f in files:
        rec = {"file": f.name}
        if do_recalc:
            r = recalc_file(f)
            if "error" in r:
                rec["recalc"] = f"GAGAL: {r['error']}"
            else:
                rec["recalc"] = r.get("status", "?")
        else:
            rec["recalc"] = "dilewati"
        n, errs, sheets = hitung_rumus_dan_error(f)
        rec["rumus"] = n
        rec["errors"] = errs
        rec["n_err"] = sum(len(v) for v in errs.values())
        rec["sheets"] = sheets
        rec["ukuran_kb"] = round(f.stat().st_size / 1024, 1)
        total_err += rec["n_err"]
        hasil_file.append(rec)
        print(f"{f.name}: recalc={rec['recalc']} rumus={n} error={rec['n_err']} {errs if errs else ''}")

    # rekonsiliasi
    build = muat_build(mod_dir)
    rows_check = []
    n_fail = 0
    if hasattr(build, "checks"):
        for chk in build.checks():
            nama, fpath, sheet, cell, expected = chk["label"], chk["file"], chk["sheet"], chk["cell"], chk["expected"]
            try:
                actual = nilai_sel(mod_dir / "latihan" / fpath, sheet, cell)
            except Exception as e:  # noqa: BLE001
                actual = f"ERR {e}"
            status = "PASS"
            selisih = ""
            if isinstance(expected, (int, float)) and isinstance(actual, (int, float)):
                d = abs(float(actual) - float(expected))
                ref = abs(float(expected)) if expected else 1.0
                selisih = f"{d:.6g}"
                if d > 1e-6 and d / ref > 0.0001:
                    status = "FAIL"
            elif isinstance(expected, (dt.date, dt.datetime)) and isinstance(actual, (dt.date, dt.datetime)):
                if actual.strftime("%Y-%m-%d") != expected.strftime("%Y-%m-%d"):
                    status = "FAIL"
            else:
                if str(actual).strip() != str(expected).strip():
                    status = "FAIL"
            if status == "FAIL":
                n_fail += 1
            rows_check.append((nama, f"{fpath}:{sheet}!{cell}", fmt(expected), fmt(actual), selisih, status))
            print(f"  [{status}] {nama}: pandas={fmt(expected)} sel={fmt(actual)}")

    catatan = getattr(build, "CATATAN_VERIFIKASI", [])
    tanggal = dt.date.today().isoformat()
    md = [f"# Verifikasi modul `{mod_dir.name}`", "",
          f"Tanggal verifikasi: {tanggal}. Mesin: LibreOffice headless (recalc) + openpyxl + pandas.",
          "Skrip: `python tools/verify.py " + mod_dir.name + "` dari folder `belajar/`.", "",
          "## Hasil recalc per file", "",
          "| File | Sheet | Jumlah rumus | Jumlah error | Status recalc | Ukuran (KB) |",
          "|---|---|---:|---:|---|---:|"]
    for rec in hasil_file:
        md.append(f"| `{rec['file']}` | {', '.join(rec['sheets'])} | {rec['rumus']} | {rec['n_err']} | {rec['recalc']} | {rec['ukuran_kb']} |")
    md.append("")
    if total_err:
        md.append("Detail error:")
        for rec in hasil_file:
            for e, cells in rec["errors"].items():
                md.append(f"- `{rec['file']}` {e}: {', '.join(cells[:30])}")
        md.append("")
    md += ["## Rekonsiliasi pandas vs sel hasil recalc", ""]
    if rows_check:
        md += ["| Pemeriksaan | Lokasi sel | Nilai pandas | Nilai sel | Selisih | Status |", "|---|---|---:|---:|---:|---|"]
        for r in rows_check:
            md.append("| " + " | ".join(r) + " |")
    else:
        md.append("Tidak ada pemeriksaan numerik untuk modul ini (lihat catatan).")
    md += ["", f"Ringkasan: {len(files)} file, total error = {total_err}, rekonsiliasi gagal = {n_fail}.", ""]
    if catatan:
        md += ["## Catatan", ""] + [f"- {c}" for c in catatan] + [""]
    (mod_dir / "VERIFIKASI.md").write_text("\n".join(md), encoding="utf-8")
    print(f"VERIFIKASI.md ditulis. error={total_err} fail={n_fail}")
    return 0 if (total_err == 0 and n_fail == 0) else 2


if __name__ == "__main__":
    sys.exit(main())
