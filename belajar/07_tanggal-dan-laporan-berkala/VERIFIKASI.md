# Verifikasi modul `07_tanggal-dan-laporan-berkala`

Tanggal verifikasi: 2026-11-05 (Asia/Jakarta). Mesin: LibreOffice headless (recalc) + openpyxl + pandas/scipy.
Skrip: `python tools/course_runtime.py 07_tanggal-dan-laporan-berkala` dari folder `belajar/`.

## Hasil recalc per file

| File | Sheet | Jumlah rumus | Jumlah error | Status recalc | Ukuran (KB) |
|---|---|---:|---:|---|---:|
| `latihan_07.xlsx` | Instruksi, Data, Latihan, Libur, Master, Laporan_Bulanan | 240 | 0 | success | 29.8 |
| `solusi_07.xlsx` | Instruksi, Data, Latihan, Libur, Master, Laporan_Bulanan | 253 | 0 | success | 30.1 |

## Rekonsiliasi pandas vs sel hasil recalc

| Pemeriksaan | Lokasi sel | Nilai pandas | Nilai sel | Selisih | Status |
|---|---|---:|---:|---:|---|
| Ringkasan 4 | solusi_07.xlsx:Latihan!C4 | 121630000 | 121630000 | 0.0 | PASS |
| Ringkasan 5 | solusi_07.xlsx:Latihan!C5 | 136220000 | 136220000 | 0.0 | PASS |
| Ringkasan 6 | solusi_07.xlsx:Latihan!C6 | 0.11995395872728767 | 0.119953958727288 | 0.0 | PASS |
| Ringkasan 7 | solusi_07.xlsx:Latihan!C7 | 22 | 22 | 0.0 | PASS |
| Ringkasan 8 | solusi_07.xlsx:Latihan!C8 | 2024-02-29 | 2024-02-29 00:00:00 | 0 | PASS |
| Ringkasan 9 | solusi_07.xlsx:Latihan!C9 | 2024-02-29 | 2024-02-29 00:00:00 | 0 | PASS |
| Ringkasan 10 | solusi_07.xlsx:Latihan!C10 | 2024 | 2024 | 0.0 | PASS |
| Ringkasan 11 | solusi_07.xlsx:Latihan!C11 | 12 | 12 | 0.0 | PASS |
| Ringkasan 12 | solusi_07.xlsx:Latihan!C12 | 31 | 31 | 0.0 | PASS |
| Ringkasan 13 | solusi_07.xlsx:Latihan!C13 | 2026-11-05 | 2026-11-05 00:00:00 | 0 | PASS |
| Ringkasan 14 | solusi_07.xlsx:Latihan!C14 | 22 | 22 | 0.0 | PASS |
| Ringkasan 15 | solusi_07.xlsx:Latihan!C15 | 11 | 11 | 0.0 | PASS |
| Ringkasan 16 | solusi_07.xlsx:Latihan!C16 | PEMBANDING NOL | PEMBANDING NOL | 0 | PASS |
| Nilai 2023-01 | solusi_07.xlsx:Laporan_Bulanan!C2 | 10290000 | 10290000 | 0.0 | PASS |
| Nilai 2023-02 | solusi_07.xlsx:Laporan_Bulanan!C3 | 0 | 0 | 0.0 | PASS |
| MoM 2023-02 | solusi_07.xlsx:Laporan_Bulanan!D3 | -1.0 | -1 | 0.0 | PASS |
| Nilai 2023-03 | solusi_07.xlsx:Laporan_Bulanan!C4 | 10820000 | 10820000 | 0.0 | PASS |
| MoM 2023-03 | solusi_07.xlsx:Laporan_Bulanan!D4 | PEMBANDING NOL | PEMBANDING NOL | 0 | PASS |
| Nilai 2023-04 | solusi_07.xlsx:Laporan_Bulanan!C5 | 12170000 | 12170000 | 0.0 | PASS |
| MoM 2023-04 | solusi_07.xlsx:Laporan_Bulanan!D5 | 0.12476894639556377 | 0.124768946395564 | 0.0 | PASS |
| Nilai 2023-05 | solusi_07.xlsx:Laporan_Bulanan!C6 | 9020000 | 9020000 | 0.0 | PASS |
| MoM 2023-05 | solusi_07.xlsx:Laporan_Bulanan!D6 | -0.25883319638455216 | -0.258833196384552 | 0.0 | PASS |
| Nilai 2023-06 | solusi_07.xlsx:Laporan_Bulanan!C7 | 13070000 | 13070000 | 0.0 | PASS |
| MoM 2023-06 | solusi_07.xlsx:Laporan_Bulanan!D7 | 0.4490022172949002 | 0.4490022172949 | -0.0 | PASS |
| Nilai 2023-07 | solusi_07.xlsx:Laporan_Bulanan!C8 | 11520000 | 11520000 | 0.0 | PASS |
| MoM 2023-07 | solusi_07.xlsx:Laporan_Bulanan!D8 | -0.11859219586840092 | -0.118592195868401 | -0.0 | PASS |
| Nilai 2023-08 | solusi_07.xlsx:Laporan_Bulanan!C9 | 10440000 | 10440000 | 0.0 | PASS |
| MoM 2023-08 | solusi_07.xlsx:Laporan_Bulanan!D9 | -0.09375 | -0.09375 | 0.0 | PASS |
| Nilai 2023-09 | solusi_07.xlsx:Laporan_Bulanan!C10 | 10930000 | 10930000 | 0.0 | PASS |
| MoM 2023-09 | solusi_07.xlsx:Laporan_Bulanan!D10 | 0.04693486590038314 | 0.0469348659003831 | -0.0 | PASS |
| Nilai 2023-10 | solusi_07.xlsx:Laporan_Bulanan!C11 | 11060000 | 11060000 | 0.0 | PASS |
| MoM 2023-10 | solusi_07.xlsx:Laporan_Bulanan!D11 | 0.011893870082342177 | 0.0118938700823422 | 0.0 | PASS |
| Nilai 2023-11 | solusi_07.xlsx:Laporan_Bulanan!C12 | 10200000 | 10200000 | 0.0 | PASS |
| MoM 2023-11 | solusi_07.xlsx:Laporan_Bulanan!D12 | -0.07775768535262206 | -0.0777576853526221 | -0.0 | PASS |
| Nilai 2023-12 | solusi_07.xlsx:Laporan_Bulanan!C13 | 12110000 | 12110000 | 0.0 | PASS |
| MoM 2023-12 | solusi_07.xlsx:Laporan_Bulanan!D13 | 0.18725490196078431 | 0.187254901960784 | -0.0 | PASS |
| Nilai 2024-01 | solusi_07.xlsx:Laporan_Bulanan!C14 | 9840000 | 9840000 | 0.0 | PASS |
| MoM 2024-01 | solusi_07.xlsx:Laporan_Bulanan!D14 | -0.1874483897605285 | -0.187448389760529 | -0.0 | PASS |
| YoY 2024-01 | solusi_07.xlsx:Laporan_Bulanan!E14 | -0.043731778425655975 | -0.043731778425656 | -0.0 | PASS |
| Nilai 2024-02 | solusi_07.xlsx:Laporan_Bulanan!C15 | 10750000 | 10750000 | 0.0 | PASS |
| MoM 2024-02 | solusi_07.xlsx:Laporan_Bulanan!D15 | 0.09247967479674797 | 0.092479674796748 | 0.0 | PASS |
| YoY 2024-02 | solusi_07.xlsx:Laporan_Bulanan!E15 | PEMBANDING NOL | PEMBANDING NOL | 0 | PASS |
| Nilai 2024-03 | solusi_07.xlsx:Laporan_Bulanan!C16 | 11840000 | 11840000 | 0.0 | PASS |
| MoM 2024-03 | solusi_07.xlsx:Laporan_Bulanan!D16 | 0.1013953488372093 | 0.101395348837209 | -0.0 | PASS |
| YoY 2024-03 | solusi_07.xlsx:Laporan_Bulanan!E16 | 0.09426987060998152 | 0.0942698706099815 | -0.0 | PASS |
| Nilai 2024-04 | solusi_07.xlsx:Laporan_Bulanan!C17 | 12520000 | 12520000 | 0.0 | PASS |
| MoM 2024-04 | solusi_07.xlsx:Laporan_Bulanan!D17 | 0.057432432432432436 | 0.0574324324324324 | -0.0 | PASS |
| YoY 2024-04 | solusi_07.xlsx:Laporan_Bulanan!E17 | 0.02875924404272802 | 0.028759244042728 | -0.0 | PASS |
| Nilai 2024-05 | solusi_07.xlsx:Laporan_Bulanan!C18 | 11320000 | 11320000 | 0.0 | PASS |
| MoM 2024-05 | solusi_07.xlsx:Laporan_Bulanan!D18 | -0.09584664536741214 | -0.0958466453674121 | 0.0 | PASS |
| YoY 2024-05 | solusi_07.xlsx:Laporan_Bulanan!E18 | 0.2549889135254989 | 0.254988913525499 | 0.0 | PASS |
| Nilai 2024-06 | solusi_07.xlsx:Laporan_Bulanan!C19 | 12660000 | 12660000 | 0.0 | PASS |
| MoM 2024-06 | solusi_07.xlsx:Laporan_Bulanan!D19 | 0.11837455830388692 | 0.118374558303887 | 0.0 | PASS |
| YoY 2024-06 | solusi_07.xlsx:Laporan_Bulanan!E19 | -0.03136954858454476 | -0.0313695485845448 | -0.0 | PASS |
| Nilai 2024-07 | solusi_07.xlsx:Laporan_Bulanan!C20 | 10540000 | 10540000 | 0.0 | PASS |
| MoM 2024-07 | solusi_07.xlsx:Laporan_Bulanan!D20 | -0.1674565560821485 | -0.167456556082149 | -0.0 | PASS |
| YoY 2024-07 | solusi_07.xlsx:Laporan_Bulanan!E20 | -0.08506944444444445 | -0.0850694444444445 | -0.0 | PASS |
| Nilai 2024-08 | solusi_07.xlsx:Laporan_Bulanan!C21 | 10630000 | 10630000 | 0.0 | PASS |
| MoM 2024-08 | solusi_07.xlsx:Laporan_Bulanan!D21 | 0.008538899430740038 | 0.00853889943074004 | 0.0 | PASS |
| YoY 2024-08 | solusi_07.xlsx:Laporan_Bulanan!E21 | 0.018199233716475097 | 0.0181992337164751 | 0.0 | PASS |
| Nilai 2024-09 | solusi_07.xlsx:Laporan_Bulanan!C22 | 10710000 | 10710000 | 0.0 | PASS |
| MoM 2024-09 | solusi_07.xlsx:Laporan_Bulanan!D22 | 0.007525870178739417 | 0.00752587017873942 | 0.0 | PASS |
| YoY 2024-09 | solusi_07.xlsx:Laporan_Bulanan!E22 | -0.020128087831655993 | -0.020128087831656 | -0.0 | PASS |
| Nilai 2024-10 | solusi_07.xlsx:Laporan_Bulanan!C23 | 12000000 | 12000000 | 0.0 | PASS |
| MoM 2024-10 | solusi_07.xlsx:Laporan_Bulanan!D23 | 0.12044817927170869 | 0.120448179271709 | 0.0 | PASS |
| YoY 2024-10 | solusi_07.xlsx:Laporan_Bulanan!E23 | 0.08499095840867993 | 0.0849909584086799 | -0.0 | PASS |
| Nilai 2024-11 | solusi_07.xlsx:Laporan_Bulanan!C24 | 12890000 | 12890000 | 0.0 | PASS |
| MoM 2024-11 | solusi_07.xlsx:Laporan_Bulanan!D24 | 0.07416666666666667 | 0.0741666666666667 | 0.0 | PASS |
| YoY 2024-11 | solusi_07.xlsx:Laporan_Bulanan!E24 | 0.26372549019607844 | 0.263725490196078 | -0.0 | PASS |
| Nilai 2024-12 | solusi_07.xlsx:Laporan_Bulanan!C25 | 10520000 | 10520000 | 0.0 | PASS |
| MoM 2024-12 | solusi_07.xlsx:Laporan_Bulanan!D25 | -0.18386346004654772 | -0.183863460046548 | -0.0 | PASS |
| YoY 2024-12 | solusi_07.xlsx:Laporan_Bulanan!E25 | -0.13129644921552436 | -0.131296449215524 | 0.0 | PASS |

Ringkasan: 2 file, total error = 0, rekonsiliasi gagal = 0.

## Catatan

- Sel kuning pada latihan sengaja kosong. Semua perhitungan pada solusi memakai rumus hidup.
- Pembanding dibaca dari CSV dan dihitung dengan pandas/scipy; statistik dibandingkan sampai enam desimal, selisih harus nol.
- Preview berasal dari render nyata sheet Latihan ke PDF dengan LibreOffice, kemudian PNG melalui pdftoppm.
- Fungsi diperiksa terhadap daftar yang diperbolehkan; tidak ada rumus array dinamis atau tautan workbook eksternal.
- Interaksi manual di Excel dan Google Sheets belum diuji langsung; fitur yang memerlukan langkah manual dijelaskan jujur di README.
- NETWORKDAYS.INTL disimpan dengan nama fungsi langsung tanpa awalan _xlfn, karena pengujian mesin ini menunjukkan awalan tersebut menghasilkan nama tidak dikenal. Interaksi Excel belum diuji langsung.
