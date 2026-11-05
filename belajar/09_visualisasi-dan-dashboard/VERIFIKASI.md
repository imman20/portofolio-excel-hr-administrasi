# Verifikasi modul `09_visualisasi-dan-dashboard`

Tanggal verifikasi: 2026-11-05 (Asia/Jakarta). Mesin: LibreOffice headless (recalc) + openpyxl + pandas/scipy.
Skrip: `python tools/course_runtime.py 09_visualisasi-dan-dashboard` dari folder `belajar/`.

## Hasil recalc per file

| File | Sheet | Jumlah rumus | Jumlah error | Status recalc | Ukuran (KB) |
|---|---|---:|---:|---|---:|
| `latihan_09.xlsx` | Instruksi, Data, Latihan, Master, Calculation, Report | 1168 | 0 | success | 47.7 |
| `solusi_09.xlsx` | Instruksi, Data, Latihan, Master, Calculation, Report | 1176 | 0 | success | 54.8 |

## Rekonsiliasi pandas vs sel hasil recalc

| Pemeriksaan | Lokasi sel | Nilai pandas | Nilai sel | Selisih | Status |
|---|---|---:|---:|---:|---|
| Ringkasan 4 | solusi_09.xlsx:Latihan!C4 | 159469650.0 | 159469650 | 0.0 | PASS |
| Ringkasan 5 | solusi_09.xlsx:Latihan!C5 | 33213576.0 | 33213576 | 0.0 | PASS |
| Ringkasan 6 | solusi_09.xlsx:Latihan!C6 | 0.2082752172592089 | 0.208275217259209 | 0.0 | PASS |
| Ringkasan 7 | solusi_09.xlsx:Latihan!C7 | 1.0288364516129032 | 1.0288364516129 | -0.0 | PASS |
| Ringkasan 8 | solusi_09.xlsx:Latihan!C8 | 0.72357 | 0.72357 | 0.0 | PASS |
| Ringkasan 9 | solusi_09.xlsx:Latihan!C9 | Kemasan | Kemasan | 0 | PASS |
| Ringkasan 10 | solusi_09.xlsx:Latihan!C10 | ATK | ATK | 0 | PASS |
| Ringkasan 11 | solusi_09.xlsx:Latihan!C11 | 0 | 0 | 0.0 | PASS |
| Bulan 1 B | solusi_09.xlsx:Calculation!B2 | 24964750.0 | 24964750 | 0.0 | PASS |
| Bulan 1 D | solusi_09.xlsx:Calculation!D2 | 5500142.0 | 5500142 | 0.0 | PASS |
| Bulan 1 E | solusi_09.xlsx:Calculation!E2 | 1.1347613636363636 | 1.13476136363636 | -0.0 | PASS |
| Bulan 1 F | solusi_09.xlsx:Calculation!F2 | 0.22031632601968776 | 0.220316326019688 | 0.0 | PASS |
| Bulan 2 B | solusi_09.xlsx:Calculation!B3 | 30773050.0 | 30773050 | 0.0 | PASS |
| Bulan 2 D | solusi_09.xlsx:Calculation!D3 | 6521354.0 | 6521354 | 0.0 | PASS |
| Bulan 2 E | solusi_09.xlsx:Calculation!E3 | 1.2822104166666666 | 1.28221041666667 | 0.0 | PASS |
| Bulan 2 F | solusi_09.xlsx:Calculation!F3 | 0.2119177007153987 | 0.211917700715399 | 0.0 | PASS |
| Bulan 3 B | solusi_09.xlsx:Calculation!B4 | 30619300.0 | 30619300 | 0.0 | PASS |
| Bulan 3 D | solusi_09.xlsx:Calculation!D4 | 6321190.0 | 6321190 | 0.0 | PASS |
| Bulan 3 E | solusi_09.xlsx:Calculation!E4 | 1.1776653846153846 | 1.17766538461538 | -0.0 | PASS |
| Bulan 3 F | solusi_09.xlsx:Calculation!F4 | 0.2064446280613861 | 0.206444628061386 | -0.0 | PASS |
| Bulan 4 B | solusi_09.xlsx:Calculation!B5 | 22008600.0 | 22008600 | 0.0 | PASS |
| Bulan 4 D | solusi_09.xlsx:Calculation!D5 | 4742651.0 | 4742651 | 0.0 | PASS |
| Bulan 4 E | solusi_09.xlsx:Calculation!E5 | 0.880344 | 0.880344 | 0.0 | PASS |
| Bulan 4 F | solusi_09.xlsx:Calculation!F5 | 0.21549080813863672 | 0.215490808138637 | 0.0 | PASS |
| Bulan 5 B | solusi_09.xlsx:Calculation!B6 | 29396850.0 | 29396850 | 0.0 | PASS |
| Bulan 5 D | solusi_09.xlsx:Calculation!D6 | 6031935.0 | 6031935 | 0.0 | PASS |
| Bulan 5 E | solusi_09.xlsx:Calculation!E6 | 1.0498875 | 1.0498875 | 0.0 | PASS |
| Bulan 5 F | solusi_09.xlsx:Calculation!F6 | 0.20518984176876093 | 0.205189841768761 | 0.0 | PASS |
| Bulan 6 B | solusi_09.xlsx:Calculation!B7 | 21707100.0 | 21707100 | 0.0 | PASS |
| Bulan 6 D | solusi_09.xlsx:Calculation!D7 | 4096304.0 | 4096304 | 0.0 | PASS |
| Bulan 6 E | solusi_09.xlsx:Calculation!E7 | 0.72357 | 0.72357 | 0.0 | PASS |
| Bulan 6 F | solusi_09.xlsx:Calculation!F7 | 0.18870802640610676 | 0.188708026406107 | 0.0 | PASS |

Ringkasan: 2 file, total error = 0, rekonsiliasi gagal = 0.

## Catatan

- Sel kuning pada latihan sengaja kosong. Semua perhitungan pada solusi memakai rumus hidup.
- Pembanding dibaca dari CSV dan dihitung dengan pandas/scipy; statistik dibandingkan sampai enam desimal, selisih harus nol.
- Preview berasal dari render nyata sheet Report ke PDF dengan LibreOffice, kemudian PNG melalui pdftoppm.
- Fungsi diperiksa terhadap daftar yang diperbolehkan; tidak ada rumus array dinamis atau tautan workbook eksternal.
- Interaksi manual di Excel dan Google Sheets belum diuji langsung; fitur yang memerlukan langkah manual dijelaskan jujur di README.
