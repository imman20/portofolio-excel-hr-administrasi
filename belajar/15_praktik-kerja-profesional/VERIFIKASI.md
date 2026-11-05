# Verifikasi modul `15_praktik-kerja-profesional`

Tanggal verifikasi: 2026-11-05 (Asia/Jakarta). Mesin: LibreOffice headless (recalc) + openpyxl + pandas/scipy.
Skrip: `python tools/course_runtime.py 15_praktik-kerja-profesional` dari folder `belajar/`.

## Hasil recalc per file

| File | Sheet | Jumlah rumus | Jumlah error | Status recalc | Ukuran (KB) |
|---|---|---:|---:|---|---:|
| `latihan_15.xlsx` | Instruksi, Data, Latihan, Pembayaran, Master, Calculation, Check_Pembayaran, Checks, Report | 907 | 0 | success | 31.1 |
| `solusi_15.xlsx` | Instruksi, Data, Latihan, Pembayaran, Master, Calculation, Check_Pembayaran, Checks, Report | 914 | 0 | success | 31.2 |
| `template_15.xlsx` | Instruksi, Data, Latihan, Pembayaran, Master, Calculation, Check_Pembayaran, Checks, Report | 914 | 0 | success | 31.5 |

## Rekonsiliasi pandas vs sel hasil recalc

| Pemeriksaan | Lokasi sel | Nilai pandas | Nilai sel | Selisih | Status |
|---|---|---:|---:|---:|---|
| Ringkasan 4 | solusi_15.xlsx:Latihan!C4 | 16303000.0 | 16303000 | 0.0 | PASS |
| Ringkasan 5 | solusi_15.xlsx:Latihan!C5 | 14861250.0 | 14861250 | 0.0 | PASS |
| Ringkasan 6 | solusi_15.xlsx:Latihan!C6 | 1491750.0 | 1491750 | 0.0 | PASS |
| Ringkasan 7 | solusi_15.xlsx:Latihan!C7 | 1.0 | 1 | 0.0 | PASS |
| Ringkasan 8 | solusi_15.xlsx:Latihan!C8 | 90000.0 | 90000 | 0.0 | PASS |
| Ringkasan 9 | solusi_15.xlsx:Latihan!C9 | 0.0 | 0 | 0.0 | PASS |
| Ringkasan 10 | solusi_15.xlsx:Latihan!C10 | 6.0 | 6 | 0.0 | PASS |
| Saldo invoice 2 | solusi_15.xlsx:Calculation!D2 | -50000.0 | -50000 | 0.0 | PASS |
| Saldo invoice 3 | solusi_15.xlsx:Calculation!D3 | 54750.0 | 54750 | 0.0 | PASS |
| Saldo invoice 4 | solusi_15.xlsx:Calculation!D4 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 5 | solusi_15.xlsx:Calculation!D5 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 6 | solusi_15.xlsx:Calculation!D6 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 7 | solusi_15.xlsx:Calculation!D7 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 8 | solusi_15.xlsx:Calculation!D8 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 9 | solusi_15.xlsx:Calculation!D9 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 10 | solusi_15.xlsx:Calculation!D10 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 11 | solusi_15.xlsx:Calculation!D11 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 12 | solusi_15.xlsx:Calculation!D12 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 13 | solusi_15.xlsx:Calculation!D13 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 14 | solusi_15.xlsx:Calculation!D14 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 15 | solusi_15.xlsx:Calculation!D15 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 16 | solusi_15.xlsx:Calculation!D16 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 17 | solusi_15.xlsx:Calculation!D17 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 18 | solusi_15.xlsx:Calculation!D18 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 19 | solusi_15.xlsx:Calculation!D19 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 20 | solusi_15.xlsx:Calculation!D20 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 21 | solusi_15.xlsx:Calculation!D21 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 22 | solusi_15.xlsx:Calculation!D22 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 23 | solusi_15.xlsx:Calculation!D23 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 24 | solusi_15.xlsx:Calculation!D24 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 25 | solusi_15.xlsx:Calculation!D25 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 26 | solusi_15.xlsx:Calculation!D26 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 27 | solusi_15.xlsx:Calculation!D27 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 28 | solusi_15.xlsx:Calculation!D28 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 29 | solusi_15.xlsx:Calculation!D29 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 30 | solusi_15.xlsx:Calculation!D30 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 31 | solusi_15.xlsx:Calculation!D31 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 32 | solusi_15.xlsx:Calculation!D32 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 33 | solusi_15.xlsx:Calculation!D33 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 34 | solusi_15.xlsx:Calculation!D34 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 35 | solusi_15.xlsx:Calculation!D35 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 36 | solusi_15.xlsx:Calculation!D36 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 37 | solusi_15.xlsx:Calculation!D37 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 38 | solusi_15.xlsx:Calculation!D38 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 39 | solusi_15.xlsx:Calculation!D39 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 40 | solusi_15.xlsx:Calculation!D40 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 41 | solusi_15.xlsx:Calculation!D41 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 42 | solusi_15.xlsx:Calculation!D42 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 43 | solusi_15.xlsx:Calculation!D43 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 44 | solusi_15.xlsx:Calculation!D44 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 45 | solusi_15.xlsx:Calculation!D45 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 46 | solusi_15.xlsx:Calculation!D46 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 47 | solusi_15.xlsx:Calculation!D47 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 48 | solusi_15.xlsx:Calculation!D48 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 49 | solusi_15.xlsx:Calculation!D49 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 50 | solusi_15.xlsx:Calculation!D50 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 51 | solusi_15.xlsx:Calculation!D51 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 52 | solusi_15.xlsx:Calculation!D52 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 53 | solusi_15.xlsx:Calculation!D53 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 54 | solusi_15.xlsx:Calculation!D54 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 55 | solusi_15.xlsx:Calculation!D55 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 56 | solusi_15.xlsx:Calculation!D56 | 0.0 | 0 | 0.0 | PASS |
| Saldo invoice 57 | solusi_15.xlsx:Calculation!D57 | 680000.0 | 680000 | 0.0 | PASS |
| Saldo invoice 58 | solusi_15.xlsx:Calculation!D58 | 39000.0 | 39000 | 0.0 | PASS |
| Saldo invoice 59 | solusi_15.xlsx:Calculation!D59 | 380000.0 | 380000 | 0.0 | PASS |
| Saldo invoice 60 | solusi_15.xlsx:Calculation!D60 | 44000.0 | 44000 | 0.0 | PASS |
| Saldo invoice 61 | solusi_15.xlsx:Calculation!D61 | 294000.0 | 294000 | 0.0 | PASS |
| Ringkasan 4 | template_15.xlsx:Latihan!C4 | 16303000.0 | 16303000 | 0.0 | PASS |
| Ringkasan 5 | template_15.xlsx:Latihan!C5 | 14861250.0 | 14861250 | 0.0 | PASS |
| Ringkasan 6 | template_15.xlsx:Latihan!C6 | 1491750.0 | 1491750 | 0.0 | PASS |
| Ringkasan 7 | template_15.xlsx:Latihan!C7 | 1.0 | 1 | 0.0 | PASS |
| Ringkasan 8 | template_15.xlsx:Latihan!C8 | 90000.0 | 90000 | 0.0 | PASS |
| Ringkasan 9 | template_15.xlsx:Latihan!C9 | 0.0 | 0 | 0.0 | PASS |
| Ringkasan 10 | template_15.xlsx:Latihan!C10 | 6.0 | 6 | 0.0 | PASS |

Ringkasan: 3 file, total error = 0, rekonsiliasi gagal = 0.

## Catatan

- Sel kuning pada latihan sengaja kosong. Semua perhitungan pada solusi memakai rumus hidup.
- Pembanding dibaca dari CSV dan dihitung dengan pandas/scipy; statistik dibandingkan sampai enam desimal, selisih harus nol.
- Preview berasal dari render nyata sheet Report ke PDF dengan LibreOffice, kemudian PNG melalui pdftoppm.
- Fungsi diperiksa terhadap daftar yang diperbolehkan; tidak ada rumus array dinamis atau tautan workbook eksternal.
- Interaksi manual di Excel dan Google Sheets belum diuji langsung; fitur yang memerlukan langkah manual dijelaskan jujur di README.
