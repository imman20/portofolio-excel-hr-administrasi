# Verifikasi modul `13_what-if-dan-optimasi`

Tanggal verifikasi: 2026-11-05 (Asia/Jakarta). Mesin: LibreOffice headless (recalc) + openpyxl + pandas/scipy.
Skrip: `python tools/course_runtime.py 13_what-if-dan-optimasi` dari folder `belajar/`.

## Hasil recalc per file

| File | Sheet | Jumlah rumus | Jumlah error | Status recalc | Ukuran (KB) |
|---|---|---:|---:|---|---:|
| `latihan_13.xlsx` | Instruksi, Data, Latihan, Master, Kandidat, Model, What_If | 47 | 0 | success | 16.2 |
| `solusi_13.xlsx` | Instruksi, Data, Latihan, Master, Kandidat, Model, What_If | 57 | 0 | success | 16.4 |

## Rekonsiliasi pandas vs sel hasil recalc

| Pemeriksaan | Lokasi sel | Nilai pandas | Nilai sel | Selisih | Status |
|---|---|---:|---:|---:|---|
| Pembanding 4 | solusi_13.xlsx:Latihan!C4 | 26.66666666666667 | 26.6666666666667 | 0.0 | PASS |
| Pembanding 5 | solusi_13.xlsx:Latihan!C5 | 46.666666666666664 | 46.6666666666667 | 0.0 | PASS |
| Pembanding 6 | solusi_13.xlsx:Latihan!C6 | 2466666.666666667 | 2466666.66666667 | 0.0 | PASS |
| Pembanding 7 | solusi_13.xlsx:Latihan!C7 | 70.0 | 70 | 0.0 | PASS |
| Pembanding 8 | solusi_13.xlsx:Latihan!C8 | 1500000.0 | 1500000 | 0.0 | PASS |
| Pembanding 9 | solusi_13.xlsx:Latihan!C9 | 2000000.0 | 2000000 | 0.0 | PASS |
| Pembanding 10 | solusi_13.xlsx:Latihan!C10 | 100.0 | 100 | 0.0 | PASS |
| Pembanding 11 | solusi_13.xlsx:Latihan!C11 | 120.0 | 120 | 0.0 | PASS |
| Model produk A | solusi_13.xlsx:Model!B2 | 26.66666666666667 | 26.6666666666667 | 0.0 | PASS |
| Model produk B | solusi_13.xlsx:Model!B3 | 46.666666666666664 | 46.6666666666667 | 0.0 | PASS |
| Model laba | solusi_13.xlsx:Model!B4 | 2466666.666666667 | 2466666.66666667 | 0.0 | PASS |
| Sisa jam | solusi_13.xlsx:Model!B7 | 0 | 0 | 0.0 | PASS |
| Sisa bahan | solusi_13.xlsx:Model!B8 | 0 | 0 | 0.0 | PASS |
| Laba skenario Dasar | solusi_13.xlsx:What_If!G12 | 2000000 | 2000000 | 0.0 | PASS |
| Laba skenario Alternatif Harga | solusi_13.xlsx:What_If!G13 | 1200000 | 1200000 | 0.0 | PASS |
| Laba skenario Alternatif Volume | solusi_13.xlsx:What_If!G14 | 1000000 | 1000000 | 0.0 | PASS |
| Sensitivitas qty 20 | solusi_13.xlsx:What_If!E2 | -1000000 | -1000000 | 0.0 | PASS |
| Sensitivitas qty 40 | solusi_13.xlsx:What_If!E3 | 0 | 0 | 0.0 | PASS |
| Sensitivitas qty 60 | solusi_13.xlsx:What_If!E4 | 1000000 | 1000000 | 0.0 | PASS |
| Sensitivitas qty 80 | solusi_13.xlsx:What_If!E5 | 2000000 | 2000000 | 0.0 | PASS |
| Sensitivitas qty 100 | solusi_13.xlsx:What_If!E6 | 3000000 | 3000000 | 0.0 | PASS |
| Sensitivitas qty 120 | solusi_13.xlsx:What_If!E7 | 4000000 | 4000000 | 0.0 | PASS |
| Sensitivitas 100000 40 | solusi_13.xlsx:What_If!H3 | -800000 | -800000 | 0.0 | PASS |
| Sensitivitas 120000 40 | solusi_13.xlsx:What_If!I3 | 0 | 0 | 0.0 | PASS |
| Sensitivitas 140000 40 | solusi_13.xlsx:What_If!J3 | 800000 | 800000 | 0.0 | PASS |
| Sensitivitas 100000 80 | solusi_13.xlsx:What_If!H4 | 400000 | 400000 | 0.0 | PASS |
| Sensitivitas 120000 80 | solusi_13.xlsx:What_If!I4 | 2000000 | 2000000 | 0.0 | PASS |
| Sensitivitas 140000 80 | solusi_13.xlsx:What_If!J4 | 3600000 | 3600000 | 0.0 | PASS |
| Sensitivitas 100000 120 | solusi_13.xlsx:What_If!H5 | 1600000 | 1600000 | 0.0 | PASS |
| Sensitivitas 120000 120 | solusi_13.xlsx:What_If!I5 | 4000000 | 4000000 | 0.0 | PASS |
| Sensitivitas 140000 120 | solusi_13.xlsx:What_If!J5 | 6400000 | 6400000 | 0.0 | PASS |

Ringkasan: 2 file, total error = 0, rekonsiliasi gagal = 0.

## Catatan

- Sel kuning pada latihan sengaja kosong. Semua perhitungan pada solusi memakai rumus hidup.
- Pembanding dibaca dari CSV dan dihitung dengan pandas/scipy; statistik dibandingkan sampai enam desimal, selisih harus nol.
- Preview berasal dari render nyata sheet Latihan ke PDF dengan LibreOffice, kemudian PNG melalui pdftoppm.
- Fungsi diperiksa terhadap daftar yang diperbolehkan; tidak ada rumus array dinamis atau tautan workbook eksternal.
- Interaksi manual di Excel dan Google Sheets belum diuji langsung; fitur yang memerlukan langkah manual dijelaskan jujur di README.
