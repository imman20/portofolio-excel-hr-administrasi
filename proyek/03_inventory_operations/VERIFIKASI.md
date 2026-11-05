# Verifikasi proyek `03_inventory_operations`

Tanggal verifikasi: 2026-11-05 (Asia/Jakarta). Mesin: LibreOffice headless (recalc) + openpyxl + pandas/scipy.
Skrip: `python tools/verify.py` dari folder proyek.

## Hasil recalc per file

| File | Sheet | Jumlah rumus | Jumlah error | Status recalc | Ukuran (KB) |
|---|---|---:|---:|---|---:|
| `latihan_inventory_operations.xlsx` | Readme, Raw_Data, Master, Permintaan_Harian, Calculation, Report, Checks, Purchase_Order | 74199 | 0 | success | 1164.9 |
| `solusi_inventory_operations.xlsx` | Readme, Raw_Data, Master, Permintaan_Harian, Calculation, Report, Checks, Purchase_Order | 74376 | 0 | success | 1168.2 |

## Rekonsiliasi pandas vs sel hasil recalc

| Pemeriksaan | Lokasi sel | Nilai pandas | Nilai sel | Selisih | Status |
|---|---|---:|---:|---:|---|
| SKU001 stok_akhir | Calculation!I2 | 123.0 | 123 | 0 | PASS |
| SKU002 stok_akhir | Calculation!I3 | 344.0 | 344 | 0 | PASS |
| SKU003 stok_akhir | Calculation!I4 | 394.0 | 394 | 0 | PASS |
| SKU004 stok_akhir | Calculation!I5 | 738.0 | 738 | 0 | PASS |
| SKU005 stok_akhir | Calculation!I6 | 1467.0 | 1467 | 0 | PASS |
| SKU006 stok_akhir | Calculation!I7 | 57.0 | 57 | 0 | PASS |
| SKU007 stok_akhir | Calculation!I8 | 187.0 | 187 | 0 | PASS |
| SKU008 stok_akhir | Calculation!I9 | 406.0 | 406 | 0 | PASS |
| SKU009 stok_akhir | Calculation!I10 | 1634.0 | 1634 | 0 | PASS |
| SKU010 stok_akhir | Calculation!I11 | 4977.0 | 4977 | 0 | PASS |
| SKU011 stok_akhir | Calculation!I12 | 163.0 | 163 | 0 | PASS |
| SKU012 stok_akhir | Calculation!I13 | 144.0 | 144 | 0 | PASS |
| SKU013 stok_akhir | Calculation!I14 | 657.0 | 657 | 0 | PASS |
| SKU014 stok_akhir | Calculation!I15 | 1286.0 | 1286 | 0 | PASS |
| SKU015 stok_akhir | Calculation!I16 | 2299.0 | 2299 | 0 | PASS |
| SKU016 stok_akhir | Calculation!I17 | 252.0 | 252 | 0 | PASS |
| SKU017 stok_akhir | Calculation!I18 | 334.0 | 334 | 0 | PASS |
| SKU018 stok_akhir | Calculation!I19 | 351.0 | 351 | 0 | PASS |
| SKU019 stok_akhir | Calculation!I20 | 2585.0 | 2585 | 0 | PASS |
| SKU020 stok_akhir | Calculation!I21 | 4602.0 | 4602 | 0 | PASS |
| SKU021 stok_akhir | Calculation!I22 | 106.0 | 106 | 0 | PASS |
| SKU022 stok_akhir | Calculation!I23 | 136.0 | 136 | 0 | PASS |
| SKU023 stok_akhir | Calculation!I24 | 554.0 | 554 | 0 | PASS |
| SKU024 stok_akhir | Calculation!I25 | 1185.0 | 1185 | 0 | PASS |
| SKU025 stok_akhir | Calculation!I26 | 4681.0 | 4681 | 0 | PASS |
| SKU026 stok_akhir | Calculation!I27 | 254.0 | 254 | 0 | PASS |
| SKU027 stok_akhir | Calculation!I28 | 135.0 | 135 | 0 | PASS |
| SKU028 stok_akhir | Calculation!I29 | 554.0 | 554 | 0 | PASS |
| SKU029 stok_akhir | Calculation!I30 | 1753.0 | 1753 | 0 | PASS |
| SKU030 stok_akhir | Calculation!I31 | 5084.0 | 5084 | 0 | PASS |
| SKU031 stok_akhir | Calculation!I32 | 121.0 | 121 | 0 | PASS |
| SKU032 stok_akhir | Calculation!I33 | 131.0 | 131 | 0 | PASS |
| SKU033 stok_akhir | Calculation!I34 | 633.0 | 633 | 0 | PASS |
| SKU034 stok_akhir | Calculation!I35 | 700.0 | 700 | 0 | PASS |
| SKU035 stok_akhir | Calculation!I36 | 2337.0 | 2337 | 0 | PASS |
| SKU036 stok_akhir | Calculation!I37 | 243.0 | 243 | 0 | PASS |
| SKU037 stok_akhir | Calculation!I38 | 262.0 | 262 | 0 | PASS |
| SKU038 stok_akhir | Calculation!I39 | 389.0 | 389 | 0 | PASS |
| SKU039 stok_akhir | Calculation!I40 | 1215.0 | 1215 | 0 | PASS |
| SKU040 stok_akhir | Calculation!I41 | 5676.0 | 5676 | 0 | PASS |
| SKU041 stok_akhir | Calculation!I42 | 95.0 | 95 | 0 | PASS |
| SKU042 stok_akhir | Calculation!I43 | 306.0 | 306 | 0 | PASS |
| SKU043 stok_akhir | Calculation!I44 | 687.0 | 687 | 0 | PASS |
| SKU044 stok_akhir | Calculation!I45 | 1396.0 | 1396 | 0 | PASS |
| SKU045 stok_akhir | Calculation!I46 | 1562.0 | 1562 | 0 | PASS |
| SKU046 stok_akhir | Calculation!I47 | 221.0 | 221 | 0 | PASS |
| SKU047 stok_akhir | Calculation!I48 | 530.0 | 530 | 0 | PASS |
| SKU048 stok_akhir | Calculation!I49 | 918.0 | 918 | 0 | PASS |
| SKU049 stok_akhir | Calculation!I50 | 1342.0 | 1342 | 0 | PASS |
| SKU050 stok_akhir | Calculation!I51 | 5742.0 | 5742 | 0 | PASS |
| SKU051 stok_akhir | Calculation!I52 | 57.0 | 57 | 0 | PASS |
| SKU052 stok_akhir | Calculation!I53 | 370.0 | 370 | 0 | PASS |
| SKU053 stok_akhir | Calculation!I54 | 861.0 | 861 | 0 | PASS |
| SKU054 stok_akhir | Calculation!I55 | 1183.0 | 1183 | 0 | PASS |
| SKU055 stok_akhir | Calculation!I56 | 2287.0 | 2287 | 0 | PASS |
| SKU056 stok_akhir | Calculation!I57 | 69.0 | 69 | 0 | PASS |
| SKU057 stok_akhir | Calculation!I58 | 343.0 | 343 | 0 | PASS |
| SKU058 stok_akhir | Calculation!I59 | 298.0 | 298 | 0 | PASS |
| SKU059 stok_akhir | Calculation!I60 | 818.0 | 818 | 0 | PASS |
| SKU060 stok_akhir | Calculation!I61 | 2753.0 | 2753 | 0 | PASS |
| SKU001 hari_stockout | Calculation!V2 | 117 | 117 | 0 | PASS |
| SKU002 hari_stockout | Calculation!V3 | 21 | 21 | 0 | PASS |
| SKU003 hari_stockout | Calculation!V4 | 0 | 0 | 0 | PASS |
| SKU004 hari_stockout | Calculation!V5 | 0 | 0 | 0 | PASS |
| SKU005 hari_stockout | Calculation!V6 | 6 | 6 | 0 | PASS |
| SKU006 hari_stockout | Calculation!V7 | 114 | 114 | 0 | PASS |
| SKU007 hari_stockout | Calculation!V8 | 43 | 43 | 0 | PASS |
| SKU008 hari_stockout | Calculation!V9 | 7 | 7 | 0 | PASS |
| SKU009 hari_stockout | Calculation!V10 | 6 | 6 | 0 | PASS |
| SKU010 hari_stockout | Calculation!V11 | 0 | 0 | 0 | PASS |
| SKU011 hari_stockout | Calculation!V12 | 122 | 122 | 0 | PASS |
| SKU012 hari_stockout | Calculation!V13 | 29 | 29 | 0 | PASS |
| SKU013 hari_stockout | Calculation!V14 | 0 | 0 | 0 | PASS |
| SKU014 hari_stockout | Calculation!V15 | 0 | 0 | 0 | PASS |
| SKU015 hari_stockout | Calculation!V16 | 7 | 7 | 0 | PASS |
| SKU016 hari_stockout | Calculation!V17 | 80 | 80 | 0 | PASS |
| SKU017 hari_stockout | Calculation!V18 | 34 | 34 | 0 | PASS |
| SKU018 hari_stockout | Calculation!V19 | 1 | 1 | 0 | PASS |
| SKU019 hari_stockout | Calculation!V20 | 0 | 0 | 0 | PASS |
| SKU020 hari_stockout | Calculation!V21 | 11 | 11 | 0 | PASS |
| SKU021 hari_stockout | Calculation!V22 | 88 | 88 | 0 | PASS |
| SKU022 hari_stockout | Calculation!V23 | 37 | 37 | 0 | PASS |
| SKU023 hari_stockout | Calculation!V24 | 0 | 0 | 0 | PASS |
| SKU024 hari_stockout | Calculation!V25 | 7 | 7 | 0 | PASS |
| SKU025 hari_stockout | Calculation!V26 | 1 | 1 | 0 | PASS |
| SKU026 hari_stockout | Calculation!V27 | 79 | 79 | 0 | PASS |
| SKU027 hari_stockout | Calculation!V28 | 47 | 47 | 0 | PASS |
| SKU028 hari_stockout | Calculation!V29 | 17 | 17 | 0 | PASS |
| SKU029 hari_stockout | Calculation!V30 | 4 | 4 | 0 | PASS |
| SKU030 hari_stockout | Calculation!V31 | 0 | 0 | 0 | PASS |
| SKU031 hari_stockout | Calculation!V32 | 72 | 72 | 0 | PASS |
| SKU032 hari_stockout | Calculation!V33 | 5 | 5 | 0 | PASS |
| SKU033 hari_stockout | Calculation!V34 | 16 | 16 | 0 | PASS |
| SKU034 hari_stockout | Calculation!V35 | 2 | 2 | 0 | PASS |
| SKU035 hari_stockout | Calculation!V36 | 0 | 0 | 0 | PASS |
| SKU036 hari_stockout | Calculation!V37 | 93 | 93 | 0 | PASS |
| SKU037 hari_stockout | Calculation!V38 | 38 | 38 | 0 | PASS |
| SKU038 hari_stockout | Calculation!V39 | 18 | 18 | 0 | PASS |
| SKU039 hari_stockout | Calculation!V40 | 0 | 0 | 0 | PASS |
| SKU040 hari_stockout | Calculation!V41 | 3 | 3 | 0 | PASS |
| SKU041 hari_stockout | Calculation!V42 | 94 | 94 | 0 | PASS |
| SKU042 hari_stockout | Calculation!V43 | 45 | 45 | 0 | PASS |
| SKU043 hari_stockout | Calculation!V44 | 2 | 2 | 0 | PASS |
| SKU044 hari_stockout | Calculation!V45 | 19 | 19 | 0 | PASS |
| SKU045 hari_stockout | Calculation!V46 | 12 | 12 | 0 | PASS |
| SKU046 hari_stockout | Calculation!V47 | 105 | 105 | 0 | PASS |
| SKU047 hari_stockout | Calculation!V48 | 61 | 61 | 0 | PASS |
| SKU048 hari_stockout | Calculation!V49 | 12 | 12 | 0 | PASS |
| SKU049 hari_stockout | Calculation!V50 | 0 | 0 | 0 | PASS |
| SKU050 hari_stockout | Calculation!V51 | 0 | 0 | 0 | PASS |
| SKU051 hari_stockout | Calculation!V52 | 101 | 101 | 0 | PASS |
| SKU052 hari_stockout | Calculation!V53 | 25 | 25 | 0 | PASS |
| SKU053 hari_stockout | Calculation!V54 | 17 | 17 | 0 | PASS |
| SKU054 hari_stockout | Calculation!V55 | 14 | 14 | 0 | PASS |
| SKU055 hari_stockout | Calculation!V56 | 13 | 13 | 0 | PASS |
| SKU056 hari_stockout | Calculation!V57 | 95 | 95 | 0 | PASS |
| SKU057 hari_stockout | Calculation!V58 | 28 | 28 | 0 | PASS |
| SKU058 hari_stockout | Calculation!V59 | 0 | 0 | 0 | PASS |
| SKU059 hari_stockout | Calculation!V60 | 0 | 0 | 0 | PASS |
| SKU060 hari_stockout | Calculation!V61 | 2 | 2 | 0 | PASS |
| SKU001 abc | Calculation!Z2 | B | B | 0 | PASS |
| SKU002 abc | Calculation!Z3 | A | A | 0 | PASS |
| SKU003 abc | Calculation!Z4 | A | A | 0 | PASS |
| SKU004 abc | Calculation!Z5 | B | B | 0 | PASS |
| SKU005 abc | Calculation!Z6 | B | B | 0 | PASS |
| SKU006 abc | Calculation!Z7 | B | B | 0 | PASS |
| SKU007 abc | Calculation!Z8 | A | A | 0 | PASS |
| SKU008 abc | Calculation!Z9 | A | A | 0 | PASS |
| SKU009 abc | Calculation!Z10 | B | B | 0 | PASS |
| SKU010 abc | Calculation!Z11 | B | B | 0 | PASS |
| SKU011 abc | Calculation!Z12 | B | B | 0 | PASS |
| SKU012 abc | Calculation!Z13 | A | A | 0 | PASS |
| SKU013 abc | Calculation!Z14 | A | A | 0 | PASS |
| SKU014 abc | Calculation!Z15 | A | A | 0 | PASS |
| SKU015 abc | Calculation!Z16 | A | A | 0 | PASS |
| SKU016 abc | Calculation!Z17 | B | B | 0 | PASS |
| SKU017 abc | Calculation!Z18 | A | A | 0 | PASS |
| SKU018 abc | Calculation!Z19 | C | C | 0 | PASS |
| SKU019 abc | Calculation!Z20 | A | A | 0 | PASS |
| SKU020 abc | Calculation!Z21 | A | A | 0 | PASS |
| SKU021 abc | Calculation!Z22 | A | A | 0 | PASS |
| SKU022 abc | Calculation!Z23 | C | C | 0 | PASS |
| SKU023 abc | Calculation!Z24 | A | A | 0 | PASS |
| SKU024 abc | Calculation!Z25 | C | C | 0 | PASS |
| SKU025 abc | Calculation!Z26 | B | B | 0 | PASS |
| SKU026 abc | Calculation!Z27 | A | A | 0 | PASS |
| SKU027 abc | Calculation!Z28 | B | B | 0 | PASS |
| SKU028 abc | Calculation!Z29 | B | B | 0 | PASS |
| SKU029 abc | Calculation!Z30 | A | A | 0 | PASS |
| SKU030 abc | Calculation!Z31 | A | A | 0 | PASS |
| SKU031 abc | Calculation!Z32 | B | B | 0 | PASS |
| SKU032 abc | Calculation!Z33 | C | C | 0 | PASS |
| SKU033 abc | Calculation!Z34 | A | A | 0 | PASS |
| SKU034 abc | Calculation!Z35 | C | C | 0 | PASS |
| SKU035 abc | Calculation!Z36 | A | A | 0 | PASS |
| SKU036 abc | Calculation!Z37 | A | A | 0 | PASS |
| SKU037 abc | Calculation!Z38 | A | A | 0 | PASS |
| SKU038 abc | Calculation!Z39 | C | C | 0 | PASS |
| SKU039 abc | Calculation!Z40 | C | C | 0 | PASS |
| SKU040 abc | Calculation!Z41 | A | A | 0 | PASS |
| SKU041 abc | Calculation!Z42 | B | B | 0 | PASS |
| SKU042 abc | Calculation!Z43 | A | A | 0 | PASS |
| SKU043 abc | Calculation!Z44 | A | A | 0 | PASS |
| SKU044 abc | Calculation!Z45 | A | A | 0 | PASS |
| SKU045 abc | Calculation!Z46 | B | B | 0 | PASS |
| SKU046 abc | Calculation!Z47 | A | A | 0 | PASS |
| SKU047 abc | Calculation!Z48 | A | A | 0 | PASS |
| SKU048 abc | Calculation!Z49 | C | C | 0 | PASS |
| SKU049 abc | Calculation!Z50 | A | A | 0 | PASS |
| SKU050 abc | Calculation!Z51 | A | A | 0 | PASS |
| SKU051 abc | Calculation!Z52 | C | C | 0 | PASS |
| SKU052 abc | Calculation!Z53 | B | B | 0 | PASS |
| SKU053 abc | Calculation!Z54 | A | A | 0 | PASS |
| SKU054 abc | Calculation!Z55 | B | B | 0 | PASS |
| SKU055 abc | Calculation!Z56 | C | C | 0 | PASS |
| SKU056 abc | Calculation!Z57 | C | C | 0 | PASS |
| SKU057 abc | Calculation!Z58 | A | A | 0 | PASS |
| SKU058 abc | Calculation!Z59 | B | B | 0 | PASS |
| SKU059 abc | Calculation!Z60 | A | A | 0 | PASS |
| SKU060 abc | Calculation!Z61 | B | B | 0 | PASS |
| Nilai stok akhir | Report!C4 | 6470217000.0 | 6470217000 | 0.0 | PASS |
| Rata-rata turnover | Report!C5 | 9.83061602348104 | 9.83061602348104 | 0.0 | PASS |
| Total hari stockout | Report!C6 | 1770 | 1770 | 0 | PASS |
| Total usulan pesan | Report!C7 | 1282.062638458292 | 1282.06263845829 | -0.0 | PASS |

Ringkasan: 2 file, total error = 0, rekonsiliasi gagal = 0. 184 pembanding utama dan 3063 pembanding rinci PASS.

## Catatan

- Sebanyak 87840 sel permintaan dan stok harian cocok tepat dengan pandas pada kedua workbook.
- Saldo, stockout dan ABC diperiksa per SKU, bukan hanya total. KPI lain, peringkat usulan dan urutan lima SKU terbesar ikut dicocokkan.
- Selisih numerik wajib nol setelah pembulatan enam desimal untuk representasi biner; unit, rupiah, hari stockout, dan label ABC cocok tepat.
- CSV stok_harian dari generator dicocokkan kembali dengan agregasi buku gerak independen, termasuk saldo awal hari, masuk, keluar dan penyesuaian.
- Relasi SKU/pemasok/PO, penerimaan satu per PO, tanggal dan stok nonnegatif diperiksa sebelum agregasi.
- Latihan menyisakan sel kuning; pemeriksaan kelengkapan BELUM DIISI adalah keadaan yang disengaja. Status dan peringkat usulan latihan belum bermakna sebelum diisi.
- Preview berasal dari Report pada salinan render: LibreOffice ke PDF satu halaman lalu pdftoppm ke PNG.
- Tidak ada rumus dinamis atau tautan eksternal. Tampilan/interaksi Excel dan Google Sheets belum diuji.
- Hasil lengkap: [rekonsiliasi](data/processed/rekonsiliasi.csv), [detail](data/processed/rekonsiliasi_detail.csv), [KPI referensi](data/processed/kpi_referensi.csv), [ringkasan dan hash](data/processed/ringkasan.json).
