# Verifikasi modul `14_otomasi-lanjutan`

Tanggal verifikasi: 2026-11-05 (Asia/Jakarta). Mesin: LibreOffice headless (recalc) + openpyxl + pandas/scipy.
Skrip: `python tools/course_runtime.py 14_otomasi-lanjutan` dari folder `belajar/`.

## Hasil recalc per file

| File | Sheet | Jumlah rumus | Jumlah error | Status recalc | Ukuran (KB) |
|---|---|---:|---:|---|---:|
| `latihan_14.xlsx` | Instruksi, Data, Latihan, Ringkasan | 246 | 0 | success | 19.2 |
| `solusi_14.xlsx` | Instruksi, Data, Latihan, Ringkasan | 253 | 0 | success | 19.4 |

## Rekonsiliasi pandas vs sel hasil recalc

| Pemeriksaan | Lokasi sel | Nilai pandas | Nilai sel | Selisih | Status |
|---|---|---:|---:|---:|---|
| Ringkasan 4 | solusi_14.xlsx:Latihan!C4 | 80 | 80 | 0.0 | PASS |
| Ringkasan 5 | solusi_14.xlsx:Latihan!C5 | 18106000 | 18106000 | 0.0 | PASS |
| Ringkasan 6 | solusi_14.xlsx:Latihan!C6 | 28 | 28 | 0.0 | PASS |
| Ringkasan 7 | solusi_14.xlsx:Latihan!C7 | 7511000 | 7511000 | 0.0 | PASS |
| Ringkasan 8 | solusi_14.xlsx:Latihan!C8 | 0 | 0 | 0.0 | PASS |
| Ringkasan 9 | solusi_14.xlsx:Latihan!C9 | 0 | 0 | 0.0 | PASS |
| Ringkasan 10 | solusi_14.xlsx:Latihan!C10 | 0 | 0 | 0.0 | PASS |
| Jumlah BARU | solusi_14.xlsx:Ringkasan!B4 | 28 | 28 | 0.0 | PASS |
| Nilai BARU | solusi_14.xlsx:Ringkasan!C4 | 7511000 | 7511000 | 0.0 | PASS |
| Jumlah DIPROSES | solusi_14.xlsx:Ringkasan!B5 | 20 | 20 | 0.0 | PASS |
| Nilai DIPROSES | solusi_14.xlsx:Ringkasan!C5 | 4307000 | 4307000 | 0.0 | PASS |
| Jumlah SELESAI | solusi_14.xlsx:Ringkasan!B6 | 32 | 32 | 0.0 | PASS |
| Nilai SELESAI | solusi_14.xlsx:Ringkasan!C6 | 6288000 | 6288000 | 0.0 | PASS |

Ringkasan: 2 file, total error = 0, rekonsiliasi gagal = 0.

## Catatan

- Sel kuning pada latihan sengaja kosong. Semua perhitungan pada solusi memakai rumus hidup.
- Pembanding dibaca dari CSV dan dihitung dengan pandas/scipy; statistik dibandingkan sampai enam desimal, selisih harus nol.
- Preview berasal dari render nyata sheet Ringkasan ke PDF dengan LibreOffice, kemudian PNG melalui pdftoppm.
- Fungsi diperiksa terhadap daftar yang diperbolehkan; tidak ada rumus array dinamis atau tautan workbook eksternal.
- Interaksi manual di Excel dan Google Sheets belum diuji langsung; fitur yang memerlukan langkah manual dijelaskan jujur di README.
