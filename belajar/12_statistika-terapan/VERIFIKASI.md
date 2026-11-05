# Verifikasi modul `12_statistika-terapan`

Tanggal verifikasi: 2026-11-05 (Asia/Jakarta). Mesin: LibreOffice headless (recalc) + openpyxl + pandas/scipy.
Skrip: `python tools/course_runtime.py 12_statistika-terapan` dari folder `belajar/`.

## Hasil recalc per file

| File | Sheet | Jumlah rumus | Jumlah error | Status recalc | Ukuran (KB) |
|---|---|---:|---:|---|---:|
| `latihan_12.xlsx` | Instruksi, Data, Latihan, Tren, Histogram | 34 | 0 | success | 16.5 |
| `solusi_12.xlsx` | Instruksi, Data, Latihan, Tren, Histogram | 57 | 0 | success | 21.1 |

## Rekonsiliasi pandas vs sel hasil recalc

| Pemeriksaan | Lokasi sel | Nilai pandas | Nilai sel | Selisih | Status |
|---|---|---:|---:|---:|---|
| n | solusi_12.xlsx:Latihan!C4 | 60.0 | 60 | 0.0 | PASS |
| Mean | solusi_12.xlsx:Latihan!C5 | 70.82633333333334 | 70.8263333333333 | -0.0 | PASS |
| Median | solusi_12.xlsx:Latihan!C6 | 71.72 | 71.72 | 0.0 | PASS |
| SD sampel | solusi_12.xlsx:Latihan!C7 | 6.797683046616685 | 6.79768304661669 | 0.0 | PASS |
| Varians sampel | solusi_12.xlsx:Latihan!C8 | 46.2084948022599 | 46.2084948022599 | 0.0 | PASS |
| Minimum | solusi_12.xlsx:Latihan!C9 | 54.55 | 54.55 | 0.0 | PASS |
| Maksimum | solusi_12.xlsx:Latihan!C10 | 85.98 | 85.98 | 0.0 | PASS |
| Margin CI | solusi_12.xlsx:Latihan!C11 | 1.7560277365133417 | 1.75602773651334 | -0.0 | PASS |
| CI bawah | solusi_12.xlsx:Latihan!C12 | 69.07030559681999 | 69.07030559682 | 0.0 | PASS |
| CI atas | solusi_12.xlsx:Latihan!C13 | 72.58236106984668 | 72.5823610698467 | 0.0 | PASS |
| Pearson r | solusi_12.xlsx:Latihan!C14 | 0.8013895905942389 | 0.801389590594239 | 0.0 | PASS |
| Slope OLS | solusi_12.xlsx:Latihan!C15 | 2.3311880505283344 | 2.33118805052834 | 0.0 | PASS |
| Intercept OLS | solusi_12.xlsx:Latihan!C16 | 56.517501079190424 | 56.5175010791904 | -0.0 | PASS |
| R kuadrat | solusi_12.xlsx:Latihan!C17 | 0.642225275912802 | 0.642225275912803 | 0.0 | PASS |
| p Welch | solusi_12.xlsx:Latihan!C18 | 0.0011528540254506655 | 0.00115285402545054 | 0 | PASS |
| p paired | solusi_12.xlsx:Latihan!C19 | 1.3743300264479845e-24 | 1.37433002644808e-24 | 0 | PASS |
| Cohen d | solusi_12.xlsx:Latihan!C20 | 0.8851863034599317 | 0.885186303459934 | 0.0 | PASS |
| p equal variance | solusi_12.xlsx:Latihan!C21 | 0.001123660685938495 | 0.00112366068593837 | 0 | PASS |
| FORECAST | solusi_12.xlsx:Latihan!C22 | 75.1670054834171 | 75.1670054834171 | 0.0 | PASS |
| TREND | solusi_12.xlsx:Latihan!C23 | 75.1670054834171 | 75.1670054834171 | 0.0 | PASS |
| Selisih prediksi | solusi_12.xlsx:Latihan!C24 | 0.0 | 0 | 0.0 | PASS |
| Ramalan periode 25 | solusi_12.xlsx:Latihan!C26 | 396.78985507246375 | 396.789855072464 | 0.0 | PASS |
| Moving average terakhir | solusi_12.xlsx:Latihan!C27 | 384.0 | 384 | 0.0 | PASS |
| Bin 0 | solusi_12.xlsx:Histogram!D2 | 0 | 0 | 0.0 | PASS |
| Bin 10 | solusi_12.xlsx:Histogram!D3 | 0 | 0 | 0.0 | PASS |
| Bin 20 | solusi_12.xlsx:Histogram!D4 | 0 | 0 | 0.0 | PASS |
| Bin 30 | solusi_12.xlsx:Histogram!D5 | 0 | 0 | 0.0 | PASS |
| Bin 40 | solusi_12.xlsx:Histogram!D6 | 0 | 0 | 0.0 | PASS |
| Bin 50 | solusi_12.xlsx:Histogram!D7 | 5 | 5 | 0.0 | PASS |
| Bin 60 | solusi_12.xlsx:Histogram!D8 | 19 | 19 | 0.0 | PASS |
| Bin 70 | solusi_12.xlsx:Histogram!D9 | 33 | 33 | 0.0 | PASS |
| Bin 80 | solusi_12.xlsx:Histogram!D10 | 3 | 3 | 0.0 | PASS |
| Bin 90 | solusi_12.xlsx:Histogram!D11 | 0 | 0 | 0.0 | PASS |
| Moving average baris 4 | solusi_12.xlsx:Tren!D4 | 220.66666666666666 | 220.666666666667 | 0.0 | PASS |
| Moving average baris 5 | solusi_12.xlsx:Tren!D5 | 232.0 | 232 | 0.0 | PASS |
| Moving average baris 6 | solusi_12.xlsx:Tren!D6 | 237.33333333333334 | 237.333333333333 | -0.0 | PASS |
| Moving average baris 7 | solusi_12.xlsx:Tren!D7 | 246.0 | 246 | 0.0 | PASS |
| Moving average baris 8 | solusi_12.xlsx:Tren!D8 | 250.33333333333334 | 250.333333333333 | -0.0 | PASS |
| Moving average baris 9 | solusi_12.xlsx:Tren!D9 | 254.33333333333334 | 254.333333333333 | -0.0 | PASS |
| Moving average baris 10 | solusi_12.xlsx:Tren!D10 | 256.0 | 256 | 0.0 | PASS |
| Moving average baris 11 | solusi_12.xlsx:Tren!D11 | 263.3333333333333 | 263.333333333333 | -0.0 | PASS |
| Moving average baris 12 | solusi_12.xlsx:Tren!D12 | 273.0 | 273 | 0.0 | PASS |
| Moving average baris 13 | solusi_12.xlsx:Tren!D13 | 287.3333333333333 | 287.333333333333 | -0.0 | PASS |
| Moving average baris 14 | solusi_12.xlsx:Tren!D14 | 299.6666666666667 | 299.666666666667 | 0.0 | PASS |
| Moving average baris 15 | solusi_12.xlsx:Tren!D15 | 310.0 | 310 | 0.0 | PASS |
| Moving average baris 16 | solusi_12.xlsx:Tren!D16 | 318.0 | 318 | 0.0 | PASS |
| Moving average baris 17 | solusi_12.xlsx:Tren!D17 | 327.0 | 327 | 0.0 | PASS |
| Moving average baris 18 | solusi_12.xlsx:Tren!D18 | 333.6666666666667 | 333.666666666667 | 0.0 | PASS |
| Moving average baris 19 | solusi_12.xlsx:Tren!D19 | 338.3333333333333 | 338.333333333333 | -0.0 | PASS |
| Moving average baris 20 | solusi_12.xlsx:Tren!D20 | 341.0 | 341 | 0.0 | PASS |
| Moving average baris 21 | solusi_12.xlsx:Tren!D21 | 345.0 | 345 | 0.0 | PASS |
| Moving average baris 22 | solusi_12.xlsx:Tren!D22 | 350.3333333333333 | 350.333333333333 | -0.0 | PASS |
| Moving average baris 23 | solusi_12.xlsx:Tren!D23 | 362.0 | 362 | 0.0 | PASS |
| Moving average baris 24 | solusi_12.xlsx:Tren!D24 | 370.6666666666667 | 370.666666666667 | 0.0 | PASS |
| Moving average baris 25 | solusi_12.xlsx:Tren!D25 | 384.0 | 384 | 0.0 | PASS |

Ringkasan: 2 file, total error = 0, rekonsiliasi gagal = 0.

## Catatan

- Sel kuning pada latihan sengaja kosong. Semua perhitungan pada solusi memakai rumus hidup.
- Pembanding dibaca dari CSV dan dihitung dengan pandas/scipy; statistik dibandingkan sampai enam desimal, selisih harus nol.
- Preview berasal dari render nyata sheet Latihan ke PDF dengan LibreOffice, kemudian PNG melalui pdftoppm.
- Fungsi diperiksa terhadap daftar yang diperbolehkan; tidak ada rumus array dinamis atau tautan workbook eksternal.
- Interaksi manual di Excel dan Google Sheets belum diuji langsung; fitur yang memerlukan langkah manual dijelaskan jujur di README.
- P-value diuji juga dengan toleransi relatif 1e-9 dan absolut 1e-50, agar hasil yang sangat kecil tidak dianggap cocok hanya karena pembulatan enam desimal.
