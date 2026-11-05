# Verifikasi proyek `01_sales_profitability`

Tanggal verifikasi: 2026-11-05 (Asia/Jakarta). Mesin: LibreOffice headless (recalc) + openpyxl + pandas.
Skrip: `python tools/verify.py` dari folder proyek.

## Hasil recalc per file

| File | Sheet | Jumlah rumus | Jumlah error | Status recalc | Ukuran (KB) |
|---|---|---:|---:|---|---:|
| `latihan_sales_profitability.xlsx` | Instruksi, Raw_Data, Retur, Master, Calculation, Report, Checks | 133 | 0 | success (LibreOffice headless) | 502.8 |
| `solusi_sales_profitability.xlsx` | Readme, Raw_Data, Retur, Master, Calculation, Report, Checks | 92656 | 0 | success (LibreOffice headless) | 1405.0 |

## Rekonsiliasi pandas vs sel hasil recalc

| Pemeriksaan | Lokasi sel | Nilai pandas | Nilai sel | Selisih | Status |
|---|---|---:|---:|---:|---|
| TAHUN omzet_bruto | solusi_sales_profitability.xlsx:Calculation!B17 | 14074330000.0 | 14074330000.0 | 0.0 | PASS |
| TAHUN penjualan_bersih | solusi_sales_profitability.xlsx:Calculation!C17 | 12253701095.0 | 12253701095.0 | 0.0 | PASS |
| TAHUN hpp | solusi_sales_profitability.xlsx:Calculation!D17 | 9512929400.0 | 9512929400.0 | 0.0 | PASS |
| TAHUN laba_kotor | solusi_sales_profitability.xlsx:Calculation!E17 | 2740771695.0 | 2740771695.0 | 0.0 | PASS |
| 2024-01 omzet_bruto | solusi_sales_profitability.xlsx:Calculation!B5 | 1131879600.0 | 1131879600.0 | 0.0 | PASS |
| 2024-01 penjualan_bersih | solusi_sales_profitability.xlsx:Calculation!C5 | 988845650.0 | 988845650.0 | 0.0 | PASS |
| 2024-01 hpp | solusi_sales_profitability.xlsx:Calculation!D5 | 751777500.0 | 751777500.0 | 0.0 | PASS |
| 2024-01 laba_kotor | solusi_sales_profitability.xlsx:Calculation!E5 | 237068150.0 | 237068150.0 | 0.0 | PASS |
| 2024-02 omzet_bruto | solusi_sales_profitability.xlsx:Calculation!B6 | 741399900.0 | 741399900.0 | 0.0 | PASS |
| 2024-02 penjualan_bersih | solusi_sales_profitability.xlsx:Calculation!C6 | 653063065.0 | 653063065.0 | 0.0 | PASS |
| 2024-02 hpp | solusi_sales_profitability.xlsx:Calculation!D6 | 503538000.0 | 503538000.0 | 0.0 | PASS |
| 2024-02 laba_kotor | solusi_sales_profitability.xlsx:Calculation!E6 | 149525065.0 | 149525065.0 | 0.0 | PASS |
| 2024-03 omzet_bruto | solusi_sales_profitability.xlsx:Calculation!B7 | 1015005000.0 | 1015005000.0 | 0.0 | PASS |
| 2024-03 penjualan_bersih | solusi_sales_profitability.xlsx:Calculation!C7 | 898123795.0 | 898123795.0 | 0.0 | PASS |
| 2024-03 hpp | solusi_sales_profitability.xlsx:Calculation!D7 | 705034700.0 | 705034700.0 | 0.0 | PASS |
| 2024-03 laba_kotor | solusi_sales_profitability.xlsx:Calculation!E7 | 193089095.0 | 193089095.0 | 0.0 | PASS |
| 2024-04 omzet_bruto | solusi_sales_profitability.xlsx:Calculation!B8 | 1144565600.0 | 1144565600.0 | 0.0 | PASS |
| 2024-04 penjualan_bersih | solusi_sales_profitability.xlsx:Calculation!C8 | 1008987480.0 | 1008987480.0 | 0.0 | PASS |
| 2024-04 hpp | solusi_sales_profitability.xlsx:Calculation!D8 | 771114800.0 | 771114800.0 | 0.0 | PASS |
| 2024-04 laba_kotor | solusi_sales_profitability.xlsx:Calculation!E8 | 237872680.0 | 237872680.0 | 0.0 | PASS |
| 2024-05 omzet_bruto | solusi_sales_profitability.xlsx:Calculation!B9 | 1109334700.0 | 1109334700.0 | 0.0 | PASS |
| 2024-05 penjualan_bersih | solusi_sales_profitability.xlsx:Calculation!C9 | 970410550.0 | 970410550.0 | 0.0 | PASS |
| 2024-05 hpp | solusi_sales_profitability.xlsx:Calculation!D9 | 756101700.0 | 756101700.0 | 0.0 | PASS |
| 2024-05 laba_kotor | solusi_sales_profitability.xlsx:Calculation!E9 | 214308850.0 | 214308850.0 | 0.0 | PASS |
| 2024-06 omzet_bruto | solusi_sales_profitability.xlsx:Calculation!B10 | 836434100.0 | 836434100.0 | 0.0 | PASS |
| 2024-06 penjualan_bersih | solusi_sales_profitability.xlsx:Calculation!C10 | 730987190.0 | 730987190.0 | 0.0 | PASS |
| 2024-06 hpp | solusi_sales_profitability.xlsx:Calculation!D10 | 571430600.0 | 571430600.0 | 0.0 | PASS |
| 2024-06 laba_kotor | solusi_sales_profitability.xlsx:Calculation!E10 | 159556590.0 | 159556590.0 | 0.0 | PASS |
| 2024-07 omzet_bruto | solusi_sales_profitability.xlsx:Calculation!B11 | 1096532300.0 | 1096532300.0 | 0.0 | PASS |
| 2024-07 penjualan_bersih | solusi_sales_profitability.xlsx:Calculation!C11 | 969339160.0 | 969339160.0 | 0.0 | PASS |
| 2024-07 hpp | solusi_sales_profitability.xlsx:Calculation!D11 | 761161800.0 | 761161800.0 | 0.0 | PASS |
| 2024-07 laba_kotor | solusi_sales_profitability.xlsx:Calculation!E11 | 208177360.0 | 208177360.0 | 0.0 | PASS |
| 2024-08 omzet_bruto | solusi_sales_profitability.xlsx:Calculation!B12 | 1086715100.0 | 1086715100.0 | 0.0 | PASS |
| 2024-08 penjualan_bersih | solusi_sales_profitability.xlsx:Calculation!C12 | 950875950.0 | 950875950.0 | 0.0 | PASS |
| 2024-08 hpp | solusi_sales_profitability.xlsx:Calculation!D12 | 728246700.0 | 728246700.0 | 0.0 | PASS |
| 2024-08 laba_kotor | solusi_sales_profitability.xlsx:Calculation!E12 | 222629250.0 | 222629250.0 | 0.0 | PASS |
| 2024-09 omzet_bruto | solusi_sales_profitability.xlsx:Calculation!B13 | 1239283800.0 | 1239283800.0 | 0.0 | PASS |
| 2024-09 penjualan_bersih | solusi_sales_profitability.xlsx:Calculation!C13 | 1078893710.0 | 1078893710.0 | 0.0 | PASS |
| 2024-09 hpp | solusi_sales_profitability.xlsx:Calculation!D13 | 827900700.0 | 827900700.0 | 0.0 | PASS |
| 2024-09 laba_kotor | solusi_sales_profitability.xlsx:Calculation!E13 | 250993010.0 | 250993010.0 | 0.0 | PASS |
| 2024-10 omzet_bruto | solusi_sales_profitability.xlsx:Calculation!B14 | 1420897600.0 | 1420897600.0 | 0.0 | PASS |
| 2024-10 penjualan_bersih | solusi_sales_profitability.xlsx:Calculation!C14 | 1229245335.0 | 1229245335.0 | 0.0 | PASS |
| 2024-10 hpp | solusi_sales_profitability.xlsx:Calculation!D14 | 951832700.0 | 951832700.0 | 0.0 | PASS |
| 2024-10 laba_kotor | solusi_sales_profitability.xlsx:Calculation!E14 | 277412635.0 | 277412635.0 | 0.0 | PASS |
| 2024-11 omzet_bruto | solusi_sales_profitability.xlsx:Calculation!B15 | 1560391800.0 | 1560391800.0 | 0.0 | PASS |
| 2024-11 penjualan_bersih | solusi_sales_profitability.xlsx:Calculation!C15 | 1332580630.0 | 1332580630.0 | 0.0 | PASS |
| 2024-11 hpp | solusi_sales_profitability.xlsx:Calculation!D15 | 1052967400.0 | 1052967400.0 | 0.0 | PASS |
| 2024-11 laba_kotor | solusi_sales_profitability.xlsx:Calculation!E15 | 279613230.0 | 279613230.0 | 0.0 | PASS |
| 2024-12 omzet_bruto | solusi_sales_profitability.xlsx:Calculation!B16 | 1691890500.0 | 1691890500.0 | 0.0 | PASS |
| 2024-12 penjualan_bersih | solusi_sales_profitability.xlsx:Calculation!C16 | 1442348580.0 | 1442348580.0 | 0.0 | PASS |
| 2024-12 hpp | solusi_sales_profitability.xlsx:Calculation!D16 | 1131822800.0 | 1131822800.0 | 0.0 | PASS |
| 2024-12 laba_kotor | solusi_sales_profitability.xlsx:Calculation!E16 | 310525780.0 | 310525780.0 | 0.0 | PASS |
| Margin tahunan | solusi_sales_profitability.xlsx:Calculation!F17 | 0.223669 | 0.223669 | -0.0 | PASS |
| Pencapaian tahunan | solusi_sales_profitability.xlsx:Calculation!H17 | 1.00399 | 1.00399 | 0.0 | PASS |
| R01 penjualan_bersih | solusi_sales_profitability.xlsx:Calculation!C21 | 4758516825.0 | 4758516825.0 | 0.0 | PASS |
| R01 hpp | solusi_sales_profitability.xlsx:Calculation!D21 | 3743734300.0 | 3743734300.0 | 0.0 | PASS |
| R01 laba_kotor | solusi_sales_profitability.xlsx:Calculation!E21 | 1014782525.0 | 1014782525.0 | 0.0 | PASS |
| R01 pencapaian | solusi_sales_profitability.xlsx:Calculation!H21 | 1.068369 | 1.068369 | 0.0 | PASS |
| R02 penjualan_bersih | solusi_sales_profitability.xlsx:Calculation!C22 | 1763168180.0 | 1763168180.0 | 0.0 | PASS |
| R02 hpp | solusi_sales_profitability.xlsx:Calculation!D22 | 1362434600.0 | 1362434600.0 | 0.0 | PASS |
| R02 laba_kotor | solusi_sales_profitability.xlsx:Calculation!E22 | 400733580.0 | 400733580.0 | 0.0 | PASS |
| R02 pencapaian | solusi_sales_profitability.xlsx:Calculation!H22 | 1.017408 | 1.017408 | 0.0 | PASS |
| R03 penjualan_bersih | solusi_sales_profitability.xlsx:Calculation!C23 | 1298545720.0 | 1298545720.0 | 0.0 | PASS |
| R03 hpp | solusi_sales_profitability.xlsx:Calculation!D23 | 1010557500.0 | 1010557500.0 | 0.0 | PASS |
| R03 laba_kotor | solusi_sales_profitability.xlsx:Calculation!E23 | 287988220.0 | 287988220.0 | 0.0 | PASS |
| R03 pencapaian | solusi_sales_profitability.xlsx:Calculation!H23 | 0.942341 | 0.942341 | 0.0 | PASS |
| R04 penjualan_bersih | solusi_sales_profitability.xlsx:Calculation!C24 | 1511206095.0 | 1511206095.0 | 0.0 | PASS |
| R04 hpp | solusi_sales_profitability.xlsx:Calculation!D24 | 1139850500.0 | 1139850500.0 | 0.0 | PASS |
| R04 laba_kotor | solusi_sales_profitability.xlsx:Calculation!E24 | 371355595.0 | 371355595.0 | 0.0 | PASS |
| R04 pencapaian | solusi_sales_profitability.xlsx:Calculation!H24 | 1.00949 | 1.00949 | 0.0 | PASS |
| R05 penjualan_bersih | solusi_sales_profitability.xlsx:Calculation!C25 | 1871749140.0 | 1871749140.0 | 0.0 | PASS |
| R05 hpp | solusi_sales_profitability.xlsx:Calculation!D25 | 1449447800.0 | 1449447800.0 | 0.0 | PASS |
| R05 laba_kotor | solusi_sales_profitability.xlsx:Calculation!E25 | 422301340.0 | 422301340.0 | 0.0 | PASS |
| R05 pencapaian | solusi_sales_profitability.xlsx:Calculation!H25 | 0.925692 | 0.925692 | -0.0 | PASS |
| R06 penjualan_bersih | solusi_sales_profitability.xlsx:Calculation!C26 | 1050515135.0 | 1050515135.0 | 0.0 | PASS |
| R06 hpp | solusi_sales_profitability.xlsx:Calculation!D26 | 806904700.0 | 806904700.0 | 0.0 | PASS |
| R06 laba_kotor | solusi_sales_profitability.xlsx:Calculation!E26 | 243610435.0 | 243610435.0 | 0.0 | PASS |
| R06 pencapaian | solusi_sales_profitability.xlsx:Calculation!H26 | 0.937123 | 0.937123 | -0.0 | PASS |
| Alat Tulis Kantor penjualan | solusi_sales_profitability.xlsx:Calculation!B31 | 1281540780.0 | 1281540780.0 | 0.0 | PASS |
| Alat Tulis Kantor margin | solusi_sales_profitability.xlsx:Calculation!E31 | 0.277643 | 0.277643 | -0.0 | PASS |
| Kertas & Cetak penjualan | solusi_sales_profitability.xlsx:Calculation!B32 | 2801169230.0 | 2801169230.0 | 0.0 | PASS |
| Kertas & Cetak margin | solusi_sales_profitability.xlsx:Calculation!E32 | 0.067302 | 0.067302 | -0.0 | PASS |
| Elektronik Kantor penjualan | solusi_sales_profitability.xlsx:Calculation!B33 | 3406565790.0 | 3406565790.0 | 0.0 | PASS |
| Elektronik Kantor margin | solusi_sales_profitability.xlsx:Calculation!E33 | 0.161605 | 0.161605 | -0.0 | PASS |
| Furnitur Kantor penjualan | solusi_sales_profitability.xlsx:Calculation!B34 | 3150905750.0 | 3150905750.0 | 0.0 | PASS |
| Furnitur Kantor margin | solusi_sales_profitability.xlsx:Calculation!E34 | 0.326528 | 0.326528 | -0.0 | PASS |
| Perlengkapan Kebersihan penjualan | solusi_sales_profitability.xlsx:Calculation!B35 | 1613519545.0 | 1613519545.0 | 0.0 | PASS |
| Perlengkapan Kebersihan margin | solusi_sales_profitability.xlsx:Calculation!E35 | 0.382432 | 0.382432 | 0.0 | PASS |
| Kontribusi 20% pelanggan | solusi_sales_profitability.xlsx:Calculation!AU7 | 0.715613 | 0.715613 | 0.0 | PASS |
| Pelanggan C001 | solusi_sales_profitability.xlsx:Calculation!AO5 | 1936705485.0 | 1936705485.0 | 0.0 | PASS |
| Pelanggan C053 | solusi_sales_profitability.xlsx:Calculation!AO6 | 534101115.0 | 534101115.0 | 0.0 | PASS |
| Pelanggan C112 | solusi_sales_profitability.xlsx:Calculation!AO7 | 489033400.0 | 489033400.0 | 0.0 | PASS |
| Pelanggan C066 | solusi_sales_profitability.xlsx:Calculation!AO8 | 469652905.0 | 469652905.0 | 0.0 | PASS |
| Pelanggan C041 | solusi_sales_profitability.xlsx:Calculation!AO9 | 418159370.0 | 418159370.0 | 0.0 | PASS |
| Pelanggan C016 | solusi_sales_profitability.xlsx:Calculation!AO10 | 389437420.0 | 389437420.0 | 0.0 | PASS |
| Pelanggan C002 | solusi_sales_profitability.xlsx:Calculation!AO11 | 371512365.0 | 371512365.0 | 0.0 | PASS |
| Pelanggan C038 | solusi_sales_profitability.xlsx:Calculation!AO12 | 370274310.0 | 370274310.0 | 0.0 | PASS |
| Pelanggan C055 | solusi_sales_profitability.xlsx:Calculation!AO13 | 313983940.0 | 313983940.0 | 0.0 | PASS |
| Pelanggan C014 | solusi_sales_profitability.xlsx:Calculation!AO14 | 312760710.0 | 312760710.0 | 0.0 | PASS |
| Pelanggan C087 | solusi_sales_profitability.xlsx:Calculation!AO15 | 305131680.0 | 305131680.0 | 0.0 | PASS |
| Pelanggan C004 | solusi_sales_profitability.xlsx:Calculation!AO16 | 292630360.0 | 292630360.0 | 0.0 | PASS |
| Pelanggan C086 | solusi_sales_profitability.xlsx:Calculation!AO17 | 279451745.0 | 279451745.0 | 0.0 | PASS |
| Pelanggan C088 | solusi_sales_profitability.xlsx:Calculation!AO18 | 271980340.0 | 271980340.0 | 0.0 | PASS |
| Pelanggan C015 | solusi_sales_profitability.xlsx:Calculation!AO19 | 268550470.0 | 268550470.0 | 0.0 | PASS |
| Pelanggan C021 | solusi_sales_profitability.xlsx:Calculation!AO20 | 263674410.0 | 263674410.0 | 0.0 | PASS |
| Pelanggan C045 | solusi_sales_profitability.xlsx:Calculation!AO21 | 257696705.0 | 257696705.0 | 0.0 | PASS |
| Pelanggan C034 | solusi_sales_profitability.xlsx:Calculation!AO22 | 228607080.0 | 228607080.0 | 0.0 | PASS |
| Pelanggan C042 | solusi_sales_profitability.xlsx:Calculation!AO23 | 186341565.0 | 186341565.0 | 0.0 | PASS |
| Pelanggan C078 | solusi_sales_profitability.xlsx:Calculation!AO24 | 184225270.0 | 184225270.0 | 0.0 | PASS |
| Pelanggan C056 | solusi_sales_profitability.xlsx:Calculation!AO25 | 174391960.0 | 174391960.0 | 0.0 | PASS |
| Pelanggan C094 | solusi_sales_profitability.xlsx:Calculation!AO26 | 153808365.0 | 153808365.0 | 0.0 | PASS |
| Pelanggan C099 | solusi_sales_profitability.xlsx:Calculation!AO27 | 149025070.0 | 149025070.0 | 0.0 | PASS |
| Pelanggan C092 | solusi_sales_profitability.xlsx:Calculation!AO28 | 147768365.0 | 147768365.0 | 0.0 | PASS |
| Pelanggan C020 | solusi_sales_profitability.xlsx:Calculation!AO29 | 137091300.0 | 137091300.0 | 0.0 | PASS |
| Pelanggan C072 | solusi_sales_profitability.xlsx:Calculation!AO30 | 134795165.0 | 134795165.0 | 0.0 | PASS |
| Pelanggan C077 | solusi_sales_profitability.xlsx:Calculation!AO31 | 131220815.0 | 131220815.0 | 0.0 | PASS |
| Pelanggan C101 | solusi_sales_profitability.xlsx:Calculation!AO32 | 129915350.0 | 129915350.0 | 0.0 | PASS |
| Pelanggan C068 | solusi_sales_profitability.xlsx:Calculation!AO33 | 128160155.0 | 128160155.0 | 0.0 | PASS |
| Pelanggan C012 | solusi_sales_profitability.xlsx:Calculation!AO34 | 115532195.0 | 115532195.0 | 0.0 | PASS |
| Pelanggan C082 | solusi_sales_profitability.xlsx:Calculation!AO35 | 113662610.0 | 113662610.0 | 0.0 | PASS |
| Pelanggan C103 | solusi_sales_profitability.xlsx:Calculation!AO36 | 109917960.0 | 109917960.0 | 0.0 | PASS |
| Pelanggan C050 | solusi_sales_profitability.xlsx:Calculation!AO37 | 107483810.0 | 107483810.0 | 0.0 | PASS |
| Pelanggan C116 | solusi_sales_profitability.xlsx:Calculation!AO38 | 105116820.0 | 105116820.0 | 0.0 | PASS |
| Pelanggan C105 | solusi_sales_profitability.xlsx:Calculation!AO39 | 102120510.0 | 102120510.0 | 0.0 | PASS |
| Pelanggan C107 | solusi_sales_profitability.xlsx:Calculation!AO40 | 99850385.0 | 99850385.0 | 0.0 | PASS |
| Pelanggan C046 | solusi_sales_profitability.xlsx:Calculation!AO41 | 97639090.0 | 97639090.0 | 0.0 | PASS |
| Pelanggan C032 | solusi_sales_profitability.xlsx:Calculation!AO42 | 96144770.0 | 96144770.0 | 0.0 | PASS |
| Pelanggan C023 | solusi_sales_profitability.xlsx:Calculation!AO43 | 94243570.0 | 94243570.0 | 0.0 | PASS |
| Pelanggan C063 | solusi_sales_profitability.xlsx:Calculation!AO44 | 92520035.0 | 92520035.0 | 0.0 | PASS |
| Pelanggan C069 | solusi_sales_profitability.xlsx:Calculation!AO45 | 92427715.0 | 92427715.0 | 0.0 | PASS |
| Pelanggan C111 | solusi_sales_profitability.xlsx:Calculation!AO46 | 90289745.0 | 90289745.0 | 0.0 | PASS |
| Pelanggan C018 | solusi_sales_profitability.xlsx:Calculation!AO47 | 83637315.0 | 83637315.0 | 0.0 | PASS |
| Pelanggan C057 | solusi_sales_profitability.xlsx:Calculation!AO48 | 75262030.0 | 75262030.0 | 0.0 | PASS |
| Pelanggan C109 | solusi_sales_profitability.xlsx:Calculation!AO49 | 73642370.0 | 73642370.0 | 0.0 | PASS |
| Pelanggan C079 | solusi_sales_profitability.xlsx:Calculation!AO50 | 64488120.0 | 64488120.0 | 0.0 | PASS |
| Pelanggan C119 | solusi_sales_profitability.xlsx:Calculation!AO51 | 61902215.0 | 61902215.0 | 0.0 | PASS |
| Pelanggan C102 | solusi_sales_profitability.xlsx:Calculation!AO52 | 61243290.0 | 61243290.0 | 0.0 | PASS |
| Pelanggan C009 | solusi_sales_profitability.xlsx:Calculation!AO53 | 56139385.0 | 56139385.0 | 0.0 | PASS |
| Pelanggan C025 | solusi_sales_profitability.xlsx:Calculation!AO54 | 55882535.0 | 55882535.0 | 0.0 | PASS |
| Pelanggan C104 | solusi_sales_profitability.xlsx:Calculation!AO55 | 54672780.0 | 54672780.0 | 0.0 | PASS |
| Pelanggan C003 | solusi_sales_profitability.xlsx:Calculation!AO56 | 49093595.0 | 49093595.0 | 0.0 | PASS |
| Pelanggan C070 | solusi_sales_profitability.xlsx:Calculation!AO57 | 45238440.0 | 45238440.0 | 0.0 | PASS |
| Pelanggan C008 | solusi_sales_profitability.xlsx:Calculation!AO58 | 43267865.0 | 43267865.0 | 0.0 | PASS |
| Pelanggan C017 | solusi_sales_profitability.xlsx:Calculation!AO59 | 36013035.0 | 36013035.0 | 0.0 | PASS |
| Pelanggan C093 | solusi_sales_profitability.xlsx:Calculation!AO60 | 35122735.0 | 35122735.0 | 0.0 | PASS |
| Pelanggan C006 | solusi_sales_profitability.xlsx:Calculation!AO61 | 30540480.0 | 30540480.0 | 0.0 | PASS |
| Pelanggan C114 | solusi_sales_profitability.xlsx:Calculation!AO62 | 27082250.0 | 27082250.0 | 0.0 | PASS |
| Pelanggan C073 | solusi_sales_profitability.xlsx:Calculation!AO63 | 26382400.0 | 26382400.0 | 0.0 | PASS |
| Pelanggan C075 | solusi_sales_profitability.xlsx:Calculation!AO64 | 25248650.0 | 25248650.0 | 0.0 | PASS |
| Pelanggan C029 | solusi_sales_profitability.xlsx:Calculation!AO65 | 21670010.0 | 21670010.0 | 0.0 | PASS |
| Pelanggan C117 | solusi_sales_profitability.xlsx:Calculation!AO66 | 20036550.0 | 20036550.0 | 0.0 | PASS |
| Pelanggan C033 | solusi_sales_profitability.xlsx:Calculation!AO67 | 19372350.0 | 19372350.0 | 0.0 | PASS |
| Pelanggan C067 | solusi_sales_profitability.xlsx:Calculation!AO68 | 18225600.0 | 18225600.0 | 0.0 | PASS |
| Pelanggan C005 | solusi_sales_profitability.xlsx:Calculation!AO69 | 17564700.0 | 17564700.0 | 0.0 | PASS |
| Pelanggan C051 | solusi_sales_profitability.xlsx:Calculation!AO70 | 17266950.0 | 17266950.0 | 0.0 | PASS |
| Pelanggan C095 | solusi_sales_profitability.xlsx:Calculation!AO71 | 16899550.0 | 16899550.0 | 0.0 | PASS |
| Pelanggan C110 | solusi_sales_profitability.xlsx:Calculation!AO72 | 15344450.0 | 15344450.0 | 0.0 | PASS |
| Pelanggan C065 | solusi_sales_profitability.xlsx:Calculation!AO73 | 15306050.0 | 15306050.0 | 0.0 | PASS |
| Pelanggan C106 | solusi_sales_profitability.xlsx:Calculation!AO74 | 14907750.0 | 14907750.0 | 0.0 | PASS |
| Pelanggan C052 | solusi_sales_profitability.xlsx:Calculation!AO75 | 14788890.0 | 14788890.0 | 0.0 | PASS |
| Pelanggan C085 | solusi_sales_profitability.xlsx:Calculation!AO76 | 14486550.0 | 14486550.0 | 0.0 | PASS |
| Pelanggan C081 | solusi_sales_profitability.xlsx:Calculation!AO77 | 14274500.0 | 14274500.0 | 0.0 | PASS |
| Pelanggan C058 | solusi_sales_profitability.xlsx:Calculation!AO78 | 14206600.0 | 14206600.0 | 0.0 | PASS |
| Pelanggan C118 | solusi_sales_profitability.xlsx:Calculation!AO79 | 14082250.0 | 14082250.0 | 0.0 | PASS |
| Pelanggan C035 | solusi_sales_profitability.xlsx:Calculation!AO80 | 13742200.0 | 13742200.0 | 0.0 | PASS |
| Pelanggan C097 | solusi_sales_profitability.xlsx:Calculation!AO81 | 13583750.0 | 13583750.0 | 0.0 | PASS |
| Pelanggan C064 | solusi_sales_profitability.xlsx:Calculation!AO82 | 13087450.0 | 13087450.0 | 0.0 | PASS |
| Pelanggan C120 | solusi_sales_profitability.xlsx:Calculation!AO83 | 12533000.0 | 12533000.0 | 0.0 | PASS |
| Pelanggan C007 | solusi_sales_profitability.xlsx:Calculation!AO84 | 12118000.0 | 12118000.0 | 0.0 | PASS |
| Pelanggan C010 | solusi_sales_profitability.xlsx:Calculation!AO85 | 11936900.0 | 11936900.0 | 0.0 | PASS |
| Pelanggan C047 | solusi_sales_profitability.xlsx:Calculation!AO86 | 11776200.0 | 11776200.0 | 0.0 | PASS |
| Pelanggan C108 | solusi_sales_profitability.xlsx:Calculation!AO87 | 11760700.0 | 11760700.0 | 0.0 | PASS |
| Pelanggan C098 | solusi_sales_profitability.xlsx:Calculation!AO88 | 11730750.0 | 11730750.0 | 0.0 | PASS |
| Pelanggan C060 | solusi_sales_profitability.xlsx:Calculation!AO89 | 11505750.0 | 11505750.0 | 0.0 | PASS |
| Pelanggan C083 | solusi_sales_profitability.xlsx:Calculation!AO90 | 11138750.0 | 11138750.0 | 0.0 | PASS |
| Pelanggan C115 | solusi_sales_profitability.xlsx:Calculation!AO91 | 11100750.0 | 11100750.0 | 0.0 | PASS |
| Pelanggan C036 | solusi_sales_profitability.xlsx:Calculation!AO92 | 10427000.0 | 10427000.0 | 0.0 | PASS |
| Pelanggan C076 | solusi_sales_profitability.xlsx:Calculation!AO93 | 10263100.0 | 10263100.0 | 0.0 | PASS |
| Pelanggan C013 | solusi_sales_profitability.xlsx:Calculation!AO94 | 10149600.0 | 10149600.0 | 0.0 | PASS |
| Pelanggan C040 | solusi_sales_profitability.xlsx:Calculation!AO95 | 10055100.0 | 10055100.0 | 0.0 | PASS |
| Pelanggan C071 | solusi_sales_profitability.xlsx:Calculation!AO96 | 8894650.0 | 8894650.0 | 0.0 | PASS |
| Pelanggan C084 | solusi_sales_profitability.xlsx:Calculation!AO97 | 8624000.0 | 8624000.0 | 0.0 | PASS |
| Pelanggan C030 | solusi_sales_profitability.xlsx:Calculation!AO98 | 8441900.0 | 8441900.0 | 0.0 | PASS |
| Pelanggan C022 | solusi_sales_profitability.xlsx:Calculation!AO99 | 8177550.0 | 8177550.0 | 0.0 | PASS |
| Pelanggan C062 | solusi_sales_profitability.xlsx:Calculation!AO100 | 8089700.0 | 8089700.0 | 0.0 | PASS |
| Pelanggan C043 | solusi_sales_profitability.xlsx:Calculation!AO101 | 8083750.0 | 8083750.0 | 0.0 | PASS |
| Pelanggan C019 | solusi_sales_profitability.xlsx:Calculation!AO102 | 7533200.0 | 7533200.0 | 0.0 | PASS |
| Pelanggan C089 | solusi_sales_profitability.xlsx:Calculation!AO103 | 7202700.0 | 7202700.0 | 0.0 | PASS |
| Pelanggan C059 | solusi_sales_profitability.xlsx:Calculation!AO104 | 6916900.0 | 6916900.0 | 0.0 | PASS |
| Pelanggan C039 | solusi_sales_profitability.xlsx:Calculation!AO105 | 6847500.0 | 6847500.0 | 0.0 | PASS |
| Pelanggan C096 | solusi_sales_profitability.xlsx:Calculation!AO106 | 6841000.0 | 6841000.0 | 0.0 | PASS |
| Pelanggan C091 | solusi_sales_profitability.xlsx:Calculation!AO107 | 6548050.0 | 6548050.0 | 0.0 | PASS |
| Pelanggan C031 | solusi_sales_profitability.xlsx:Calculation!AO108 | 5725450.0 | 5725450.0 | 0.0 | PASS |
| Pelanggan C044 | solusi_sales_profitability.xlsx:Calculation!AO109 | 5601000.0 | 5601000.0 | 0.0 | PASS |
| Pelanggan C028 | solusi_sales_profitability.xlsx:Calculation!AO110 | 5547600.0 | 5547600.0 | 0.0 | PASS |
| Pelanggan C027 | solusi_sales_profitability.xlsx:Calculation!AO111 | 5528300.0 | 5528300.0 | 0.0 | PASS |
| Pelanggan C080 | solusi_sales_profitability.xlsx:Calculation!AO112 | 5459000.0 | 5459000.0 | 0.0 | PASS |
| Pelanggan C090 | solusi_sales_profitability.xlsx:Calculation!AO113 | 5281250.0 | 5281250.0 | 0.0 | PASS |
| Pelanggan C048 | solusi_sales_profitability.xlsx:Calculation!AO114 | 4790200.0 | 4790200.0 | 0.0 | PASS |
| Pelanggan C054 | solusi_sales_profitability.xlsx:Calculation!AO115 | 4675150.0 | 4675150.0 | 0.0 | PASS |
| Pelanggan C074 | solusi_sales_profitability.xlsx:Calculation!AO116 | 4337700.0 | 4337700.0 | 0.0 | PASS |
| Pelanggan C024 | solusi_sales_profitability.xlsx:Calculation!AO117 | 4279950.0 | 4279950.0 | 0.0 | PASS |
| Pelanggan C026 | solusi_sales_profitability.xlsx:Calculation!AO118 | 4044000.0 | 4044000.0 | 0.0 | PASS |
| Pelanggan C113 | solusi_sales_profitability.xlsx:Calculation!AO119 | 3832500.0 | 3832500.0 | 0.0 | PASS |
| Pelanggan C100 | solusi_sales_profitability.xlsx:Calculation!AO120 | 3503900.0 | 3503900.0 | 0.0 | PASS |
| Pelanggan C049 | solusi_sales_profitability.xlsx:Calculation!AO121 | 3359000.0 | 3359000.0 | 0.0 | PASS |
| Pelanggan C011 | solusi_sales_profitability.xlsx:Calculation!AO122 | 2942400.0 | 2942400.0 | 0.0 | PASS |
| Pelanggan C037 | solusi_sales_profitability.xlsx:Calculation!AO123 | 2911950.0 | 2911950.0 | 0.0 | PASS |
| Pelanggan C061 | solusi_sales_profitability.xlsx:Calculation!AO124 | 2443200.0 | 2443200.0 | 0.0 | PASS |
| Contoh baris 2 omzet_bruto | latihan_sales_profitability.xlsx:Raw_Data!J2 | 362000.0 | 362000.0 | 0.0 | PASS |
| Contoh baris 2 penjualan_bersih | latihan_sales_profitability.xlsx:Raw_Data!O2 | 307700.0 | 307700.0 | 0.0 | PASS |
| Contoh baris 2 hpp | latihan_sales_profitability.xlsx:Raw_Data!Q2 | 200000.0 | 200000.0 | 0.0 | PASS |
| Contoh baris 2 laba_kotor | latihan_sales_profitability.xlsx:Raw_Data!R2 | 107700.0 | 107700.0 | 0.0 | PASS |

Ringkasan: 2 file, total error = 0, rekonsiliasi gagal = 0. 213 pembanding agregat/contoh PASS.

## Catatan

- Empat KPI direkonsiliasi untuk setahun dan setiap bulan. Region, kategori, urutan Pareto lengkap, dan baris contoh latihan ikut diperiksa.
- Seluruh baris solusi juga dicocokkan dengan pandas untuk omzet, diskon, neto, retur, HPP, serta laba; relasi dan tanggal diperiksa sebelum agregasi.
- Selisih moneter harus nol setelah pembulatan dua desimal; rasio dibandingkan sampai dua belas desimal untuk mengatasi representasi biner. Angka data ini menghasilkan rupiah bulat.
- Latihan sengaja menyisakan sel kuning kosong. Status Checks adalah BELUM DIISI; ini tidak dihitung sebagai error rumus.
- Fungsi diperiksa terhadap daftar yang diizinkan; tidak ada rumus array dinamis atau tautan workbook eksternal.
- Preview berasal dari Report pada salinan sementara workbook, diekspor LibreOffice ke PDF lalu pdftoppm ke PNG. Sheet lain disembunyikan hanya pada salinan render.
- Pemeriksaan ini memakai LibreOfficeDev; tampilan dan interaksi Excel tetap perlu ditinjau saat dibuka di Excel. PivotTable dan Power Query belum tertanam.
- Hasil terstruktur: [rekonsiliasi CSV](data/processed/rekonsiliasi.csv) dan [ringkasan JSON](data/processed/ringkasan.json).
