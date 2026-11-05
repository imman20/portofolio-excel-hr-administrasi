# Verifikasi modul `08_pivottable-dan-pivotchart`

Tanggal verifikasi: 2026-11-05 (Asia/Jakarta). Mesin: LibreOffice headless (recalc) + openpyxl + pandas/scipy.
Skrip: `python tools/course_runtime.py 08_pivottable-dan-pivotchart` dari folder `belajar/`.

## Hasil recalc per file

| File | Sheet | Jumlah rumus | Jumlah error | Status recalc | Ukuran (KB) |
|---|---|---:|---:|---|---:|
| `latihan_08.xlsx` | Instruksi, Data, Latihan, Hasil_Pivot_Diharapkan | 286 | 0 | success | 24.6 |
| `solusi_08.xlsx` | Instruksi, Data, Latihan, Hasil_Pivot_Diharapkan | 294 | 0 | success | 24.7 |

## Rekonsiliasi pandas vs sel hasil recalc

| Pemeriksaan | Lokasi sel | Nilai pandas | Nilai sel | Selisih | Status |
|---|---|---:|---:|---:|---|
| Ringkasan 4 | solusi_08.xlsx:Latihan!C4 | 63802000.0 | 63802000 | 0.0 | PASS |
| Ringkasan 5 | solusi_08.xlsx:Latihan!C5 | 0.39077144917087236 | 0.390771449170872 | -0.0 | PASS |
| Ringkasan 6 | solusi_08.xlsx:Latihan!C6 | 6672000.0 | 6672000 | 0.0 | PASS |
| Ringkasan 7 | solusi_08.xlsx:Latihan!C7 | 238.0 | 238 | 0.0 | PASS |
| Ringkasan 8 | solusi_08.xlsx:Latihan!C8 | 86.0 | 86 | 0.0 | PASS |
| Ringkasan 9 | solusi_08.xlsx:Latihan!C9 | 10095000.0 | 10095000 | 0.0 | PASS |
| Ringkasan 10 | solusi_08.xlsx:Latihan!C10 | 1.0 | 1 | 0.0 | PASS |
| Ringkasan 11 | solusi_08.xlsx:Latihan!C11 | 0.0 | 0 | 0.0 | PASS |
| Jawa ATK | solusi_08.xlsx:Hasil_Pivot_Diharapkan!B5 | 6672000 | 6672000 | 0.0 | PASS |
| Jawa Kemasan | solusi_08.xlsx:Hasil_Pivot_Diharapkan!C5 | 10399000 | 10399000 | 0.0 | PASS |
| Jawa Peralatan | solusi_08.xlsx:Hasil_Pivot_Diharapkan!D5 | 7861000 | 7861000 | 0.0 | PASS |
| Sumatra ATK | solusi_08.xlsx:Hasil_Pivot_Diharapkan!B6 | 4917000 | 4917000 | 0.0 | PASS |
| Sumatra Kemasan | solusi_08.xlsx:Hasil_Pivot_Diharapkan!C6 | 7056000 | 7056000 | 0.0 | PASS |
| Sumatra Peralatan | solusi_08.xlsx:Hasil_Pivot_Diharapkan!D6 | 8457000 | 8457000 | 0.0 | PASS |
| Bali ATK | solusi_08.xlsx:Hasil_Pivot_Diharapkan!B7 | 5288000 | 5288000 | 0.0 | PASS |
| Bali Kemasan | solusi_08.xlsx:Hasil_Pivot_Diharapkan!C7 | 4640000 | 4640000 | 0.0 | PASS |
| Bali Peralatan | solusi_08.xlsx:Hasil_Pivot_Diharapkan!D7 | 8512000 | 8512000 | 0.0 | PASS |
| Bulan 1 | solusi_08.xlsx:Hasil_Pivot_Diharapkan!B19 | 10095000 | 10095000 | 0.0 | PASS |
| Bulan 2 | solusi_08.xlsx:Hasil_Pivot_Diharapkan!B20 | 11400000 | 11400000 | 0.0 | PASS |
| Bulan 3 | solusi_08.xlsx:Hasil_Pivot_Diharapkan!B21 | 10197000 | 10197000 | 0.0 | PASS |
| Bulan 4 | solusi_08.xlsx:Hasil_Pivot_Diharapkan!B22 | 12328000 | 12328000 | 0.0 | PASS |
| Bulan 5 | solusi_08.xlsx:Hasil_Pivot_Diharapkan!B23 | 8823000 | 8823000 | 0.0 | PASS |
| Bulan 6 | solusi_08.xlsx:Hasil_Pivot_Diharapkan!B24 | 10959000 | 10959000 | 0.0 | PASS |

Ringkasan: 2 file, total error = 0, rekonsiliasi gagal = 0.

## Catatan

- Sel kuning pada latihan sengaja kosong. Semua perhitungan pada solusi memakai rumus hidup.
- Pembanding dibaca dari CSV dan dihitung dengan pandas/scipy; statistik dibandingkan sampai enam desimal, selisih harus nol.
- Preview berasal dari render nyata sheet Hasil_Pivot_Diharapkan ke PDF dengan LibreOffice, kemudian PNG melalui pdftoppm.
- Fungsi diperiksa terhadap daftar yang diperbolehkan; tidak ada rumus array dinamis atau tautan workbook eksternal.
- Interaksi manual di Excel dan Google Sheets belum diuji langsung; fitur yang memerlukan langkah manual dijelaskan jujur di README.
