# Verifikasi modul `10_power-query`

Tanggal verifikasi: 2026-11-05 (Asia/Jakarta). Mesin: LibreOffice headless (recalc) + openpyxl + pandas/scipy.
Skrip: `python tools/course_runtime.py 10_power-query` dari folder `belajar/`.

## Hasil recalc per file

| File | Sheet | Jumlah rumus | Jumlah error | Status recalc | Ukuran (KB) |
|---|---|---:|---:|---|---:|
| `latihan_10.xlsx` | Instruksi, Data, Latihan, Master, Target_Wide, Hasil_Query_Diharapkan | 624 | 0 | success | 25.4 |
| `solusi_10.xlsx` | Instruksi, Data, Latihan, Master, Target_Wide, Hasil_Query_Diharapkan | 631 | 0 | success | 25.5 |

## Rekonsiliasi pandas vs sel hasil recalc

| Pemeriksaan | Lokasi sel | Nilai pandas | Nilai sel | Selisih | Status |
|---|---|---:|---:|---:|---|
| Ringkasan 4 | solusi_10.xlsx:Latihan!C4 | 200 | 200 | 0.0 | PASS |
| Ringkasan 5 | solusi_10.xlsx:Latihan!C5 | 70996000 | 70996000 | 0.0 | PASS |
| Ringkasan 6 | solusi_10.xlsx:Latihan!C6 | 37798000 | 37798000 | 0.0 | PASS |
| Ringkasan 7 | solusi_10.xlsx:Latihan!C7 | 33198000 | 33198000 | 0.0 | PASS |
| Ringkasan 8 | solusi_10.xlsx:Latihan!C8 | 6 | 6 | 0.0 | PASS |
| Ringkasan 9 | solusi_10.xlsx:Latihan!C9 | 43500000 | 43500000 | 0.0 | PASS |
| Ringkasan 10 | solusi_10.xlsx:Latihan!C10 | 0 | 0 | 0.0 | PASS |
| Nilai gabungan 2 | solusi_10.xlsx:Data!F2 | 513000 | 513000 | 0.0 | PASS |
| Nilai gabungan 3 | solusi_10.xlsx:Data!F3 | 136000 | 136000 | 0.0 | PASS |
| Nilai gabungan 4 | solusi_10.xlsx:Data!F4 | 170000 | 170000 | 0.0 | PASS |
| Nilai gabungan 5 | solusi_10.xlsx:Data!F5 | 87000 | 87000 | 0.0 | PASS |
| Nilai gabungan 6 | solusi_10.xlsx:Data!F6 | 585000 | 585000 | 0.0 | PASS |
| Nilai gabungan 7 | solusi_10.xlsx:Data!F7 | 238000 | 238000 | 0.0 | PASS |
| Nilai gabungan 8 | solusi_10.xlsx:Data!F8 | 544000 | 544000 | 0.0 | PASS |
| Nilai gabungan 9 | solusi_10.xlsx:Data!F9 | 426000 | 426000 | 0.0 | PASS |
| Nilai gabungan 10 | solusi_10.xlsx:Data!F10 | 627000 | 627000 | 0.0 | PASS |
| Nilai gabungan 11 | solusi_10.xlsx:Data!F11 | 1020000 | 1020000 | 0.0 | PASS |
| Nilai gabungan 12 | solusi_10.xlsx:Data!F12 | 136000 | 136000 | 0.0 | PASS |
| Nilai gabungan 13 | solusi_10.xlsx:Data!F13 | 342000 | 342000 | 0.0 | PASS |
| Nilai gabungan 14 | solusi_10.xlsx:Data!F14 | 315000 | 315000 | 0.0 | PASS |
| Nilai gabungan 15 | solusi_10.xlsx:Data!F15 | 136000 | 136000 | 0.0 | PASS |
| Nilai gabungan 16 | solusi_10.xlsx:Data!F16 | 630000 | 630000 | 0.0 | PASS |
| Nilai gabungan 17 | solusi_10.xlsx:Data!F17 | 812000 | 812000 | 0.0 | PASS |
| Nilai gabungan 18 | solusi_10.xlsx:Data!F18 | 58000 | 58000 | 0.0 | PASS |
| Nilai gabungan 19 | solusi_10.xlsx:Data!F19 | 798000 | 798000 | 0.0 | PASS |
| Nilai gabungan 20 | solusi_10.xlsx:Data!F20 | 119000 | 119000 | 0.0 | PASS |
| Nilai gabungan 21 | solusi_10.xlsx:Data!F21 | 855000 | 855000 | 0.0 | PASS |
| Nilai gabungan 22 | solusi_10.xlsx:Data!F22 | 348000 | 348000 | 0.0 | PASS |
| Nilai gabungan 23 | solusi_10.xlsx:Data!F23 | 435000 | 435000 | 0.0 | PASS |
| Nilai gabungan 24 | solusi_10.xlsx:Data!F24 | 510000 | 510000 | 0.0 | PASS |
| Nilai gabungan 25 | solusi_10.xlsx:Data!F25 | 255000 | 255000 | 0.0 | PASS |
| Nilai gabungan 26 | solusi_10.xlsx:Data!F26 | 570000 | 570000 | 0.0 | PASS |
| Nilai gabungan 27 | solusi_10.xlsx:Data!F27 | 145000 | 145000 | 0.0 | PASS |
| Nilai gabungan 28 | solusi_10.xlsx:Data!F28 | 255000 | 255000 | 0.0 | PASS |
| Nilai gabungan 29 | solusi_10.xlsx:Data!F29 | 142000 | 142000 | 0.0 | PASS |
| Nilai gabungan 30 | solusi_10.xlsx:Data!F30 | 510000 | 510000 | 0.0 | PASS |
| Nilai gabungan 31 | solusi_10.xlsx:Data!F31 | 46000 | 46000 | 0.0 | PASS |
| Nilai gabungan 32 | solusi_10.xlsx:Data!F32 | 135000 | 135000 | 0.0 | PASS |
| Nilai gabungan 33 | solusi_10.xlsx:Data!F33 | 675000 | 675000 | 0.0 | PASS |
| Nilai gabungan 34 | solusi_10.xlsx:Data!F34 | 116000 | 116000 | 0.0 | PASS |
| Nilai gabungan 35 | solusi_10.xlsx:Data!F35 | 798000 | 798000 | 0.0 | PASS |
| Nilai gabungan 36 | solusi_10.xlsx:Data!F36 | 870000 | 870000 | 0.0 | PASS |
| Nilai gabungan 37 | solusi_10.xlsx:Data!F37 | 255000 | 255000 | 0.0 | PASS |
| Nilai gabungan 38 | solusi_10.xlsx:Data!F38 | 456000 | 456000 | 0.0 | PASS |
| Nilai gabungan 39 | solusi_10.xlsx:Data!F39 | 497000 | 497000 | 0.0 | PASS |
| Nilai gabungan 40 | solusi_10.xlsx:Data!F40 | 265000 | 265000 | 0.0 | PASS |
| Nilai gabungan 41 | solusi_10.xlsx:Data!F41 | 238000 | 238000 | 0.0 | PASS |
| Nilai gabungan 42 | solusi_10.xlsx:Data!F42 | 348000 | 348000 | 0.0 | PASS |
| Nilai gabungan 43 | solusi_10.xlsx:Data!F43 | 371000 | 371000 | 0.0 | PASS |
| Nilai gabungan 44 | solusi_10.xlsx:Data!F44 | 406000 | 406000 | 0.0 | PASS |
| Nilai gabungan 45 | solusi_10.xlsx:Data!F45 | 318000 | 318000 | 0.0 | PASS |
| Nilai gabungan 46 | solusi_10.xlsx:Data!F46 | 68000 | 68000 | 0.0 | PASS |
| Nilai gabungan 47 | solusi_10.xlsx:Data!F47 | 638000 | 638000 | 0.0 | PASS |
| Nilai gabungan 48 | solusi_10.xlsx:Data!F48 | 377000 | 377000 | 0.0 | PASS |
| Nilai gabungan 49 | solusi_10.xlsx:Data!F49 | 255000 | 255000 | 0.0 | PASS |
| Nilai gabungan 50 | solusi_10.xlsx:Data!F50 | 322000 | 322000 | 0.0 | PASS |
| Nilai gabungan 51 | solusi_10.xlsx:Data!F51 | 513000 | 513000 | 0.0 | PASS |
| Nilai gabungan 52 | solusi_10.xlsx:Data!F52 | 1065000 | 1065000 | 0.0 | PASS |
| Nilai gabungan 53 | solusi_10.xlsx:Data!F53 | 952000 | 952000 | 0.0 | PASS |
| Nilai gabungan 54 | solusi_10.xlsx:Data!F54 | 696000 | 696000 | 0.0 | PASS |
| Nilai gabungan 55 | solusi_10.xlsx:Data!F55 | 710000 | 710000 | 0.0 | PASS |
| Nilai gabungan 56 | solusi_10.xlsx:Data!F56 | 377000 | 377000 | 0.0 | PASS |
| Nilai gabungan 57 | solusi_10.xlsx:Data!F57 | 476000 | 476000 | 0.0 | PASS |
| Nilai gabungan 58 | solusi_10.xlsx:Data!F58 | 255000 | 255000 | 0.0 | PASS |
| Nilai gabungan 59 | solusi_10.xlsx:Data!F59 | 371000 | 371000 | 0.0 | PASS |
| Nilai gabungan 60 | solusi_10.xlsx:Data!F60 | 368000 | 368000 | 0.0 | PASS |
| Nilai gabungan 61 | solusi_10.xlsx:Data!F61 | 232000 | 232000 | 0.0 | PASS |
| Nilai gabungan 62 | solusi_10.xlsx:Data!F62 | 435000 | 435000 | 0.0 | PASS |
| Nilai gabungan 63 | solusi_10.xlsx:Data!F63 | 203000 | 203000 | 0.0 | PASS |
| Nilai gabungan 64 | solusi_10.xlsx:Data!F64 | 153000 | 153000 | 0.0 | PASS |
| Nilai gabungan 65 | solusi_10.xlsx:Data!F65 | 530000 | 530000 | 0.0 | PASS |
| Nilai gabungan 66 | solusi_10.xlsx:Data!F66 | 29000 | 29000 | 0.0 | PASS |
| Nilai gabungan 67 | solusi_10.xlsx:Data!F67 | 663000 | 663000 | 0.0 | PASS |
| Nilai gabungan 68 | solusi_10.xlsx:Data!F68 | 51000 | 51000 | 0.0 | PASS |
| Nilai gabungan 69 | solusi_10.xlsx:Data!F69 | 92000 | 92000 | 0.0 | PASS |
| Nilai gabungan 70 | solusi_10.xlsx:Data!F70 | 29000 | 29000 | 0.0 | PASS |
| Nilai gabungan 71 | solusi_10.xlsx:Data!F71 | 994000 | 994000 | 0.0 | PASS |
| Nilai gabungan 72 | solusi_10.xlsx:Data!F72 | 261000 | 261000 | 0.0 | PASS |
| Nilai gabungan 73 | solusi_10.xlsx:Data!F73 | 315000 | 315000 | 0.0 | PASS |
| Nilai gabungan 74 | solusi_10.xlsx:Data!F74 | 57000 | 57000 | 0.0 | PASS |
| Nilai gabungan 75 | solusi_10.xlsx:Data!F75 | 315000 | 315000 | 0.0 | PASS |
| Nilai gabungan 76 | solusi_10.xlsx:Data!F76 | 230000 | 230000 | 0.0 | PASS |
| Nilai gabungan 77 | solusi_10.xlsx:Data!F77 | 46000 | 46000 | 0.0 | PASS |
| Nilai gabungan 78 | solusi_10.xlsx:Data!F78 | 261000 | 261000 | 0.0 | PASS |
| Nilai gabungan 79 | solusi_10.xlsx:Data!F79 | 377000 | 377000 | 0.0 | PASS |
| Nilai gabungan 80 | solusi_10.xlsx:Data!F80 | 106000 | 106000 | 0.0 | PASS |
| Nilai gabungan 81 | solusi_10.xlsx:Data!F81 | 285000 | 285000 | 0.0 | PASS |
| Nilai gabungan 82 | solusi_10.xlsx:Data!F82 | 544000 | 544000 | 0.0 | PASS |
| Nilai gabungan 83 | solusi_10.xlsx:Data!F83 | 51000 | 51000 | 0.0 | PASS |
| Nilai gabungan 84 | solusi_10.xlsx:Data!F84 | 203000 | 203000 | 0.0 | PASS |
| Nilai gabungan 85 | solusi_10.xlsx:Data!F85 | 435000 | 435000 | 0.0 | PASS |
| Nilai gabungan 86 | solusi_10.xlsx:Data!F86 | 232000 | 232000 | 0.0 | PASS |
| Nilai gabungan 87 | solusi_10.xlsx:Data!F87 | 923000 | 923000 | 0.0 | PASS |
| Nilai gabungan 88 | solusi_10.xlsx:Data!F88 | 174000 | 174000 | 0.0 | PASS |
| Nilai gabungan 89 | solusi_10.xlsx:Data!F89 | 552000 | 552000 | 0.0 | PASS |
| Nilai gabungan 90 | solusi_10.xlsx:Data!F90 | 405000 | 405000 | 0.0 | PASS |
| Nilai gabungan 91 | solusi_10.xlsx:Data!F91 | 212000 | 212000 | 0.0 | PASS |
| Nilai gabungan 92 | solusi_10.xlsx:Data!F92 | 184000 | 184000 | 0.0 | PASS |
| Nilai gabungan 93 | solusi_10.xlsx:Data!F93 | 612000 | 612000 | 0.0 | PASS |
| Nilai gabungan 94 | solusi_10.xlsx:Data!F94 | 92000 | 92000 | 0.0 | PASS |
| Nilai gabungan 95 | solusi_10.xlsx:Data!F95 | 203000 | 203000 | 0.0 | PASS |
| Nilai gabungan 96 | solusi_10.xlsx:Data!F96 | 17000 | 17000 | 0.0 | PASS |
| Nilai gabungan 97 | solusi_10.xlsx:Data!F97 | 142000 | 142000 | 0.0 | PASS |
| Nilai gabungan 98 | solusi_10.xlsx:Data!F98 | 261000 | 261000 | 0.0 | PASS |
| Nilai gabungan 99 | solusi_10.xlsx:Data!F99 | 1065000 | 1065000 | 0.0 | PASS |
| Nilai gabungan 100 | solusi_10.xlsx:Data!F100 | 184000 | 184000 | 0.0 | PASS |
| Nilai gabungan 101 | solusi_10.xlsx:Data!F101 | 689000 | 689000 | 0.0 | PASS |
| Nilai gabungan 102 | solusi_10.xlsx:Data!F102 | 261000 | 261000 | 0.0 | PASS |
| Nilai gabungan 103 | solusi_10.xlsx:Data!F103 | 187000 | 187000 | 0.0 | PASS |
| Nilai gabungan 104 | solusi_10.xlsx:Data!F104 | 71000 | 71000 | 0.0 | PASS |
| Nilai gabungan 105 | solusi_10.xlsx:Data!F105 | 159000 | 159000 | 0.0 | PASS |
| Nilai gabungan 106 | solusi_10.xlsx:Data!F106 | 135000 | 135000 | 0.0 | PASS |
| Nilai gabungan 107 | solusi_10.xlsx:Data!F107 | 357000 | 357000 | 0.0 | PASS |
| Nilai gabungan 108 | solusi_10.xlsx:Data!F108 | 680000 | 680000 | 0.0 | PASS |
| Nilai gabungan 109 | solusi_10.xlsx:Data!F109 | 852000 | 852000 | 0.0 | PASS |
| Nilai gabungan 110 | solusi_10.xlsx:Data!F110 | 213000 | 213000 | 0.0 | PASS |
| Nilai gabungan 111 | solusi_10.xlsx:Data!F111 | 58000 | 58000 | 0.0 | PASS |
| Nilai gabungan 112 | solusi_10.xlsx:Data!F112 | 106000 | 106000 | 0.0 | PASS |
| Nilai gabungan 113 | solusi_10.xlsx:Data!F113 | 92000 | 92000 | 0.0 | PASS |
| Nilai gabungan 114 | solusi_10.xlsx:Data!F114 | 408000 | 408000 | 0.0 | PASS |
| Nilai gabungan 115 | solusi_10.xlsx:Data!F115 | 145000 | 145000 | 0.0 | PASS |
| Nilai gabungan 116 | solusi_10.xlsx:Data!F116 | 377000 | 377000 | 0.0 | PASS |
| Nilai gabungan 117 | solusi_10.xlsx:Data!F117 | 290000 | 290000 | 0.0 | PASS |
| Nilai gabungan 118 | solusi_10.xlsx:Data!F118 | 408000 | 408000 | 0.0 | PASS |
| Nilai gabungan 119 | solusi_10.xlsx:Data!F119 | 284000 | 284000 | 0.0 | PASS |
| Nilai gabungan 120 | solusi_10.xlsx:Data!F120 | 639000 | 639000 | 0.0 | PASS |
| Nilai gabungan 121 | solusi_10.xlsx:Data!F121 | 318000 | 318000 | 0.0 | PASS |
| Nilai gabungan 122 | solusi_10.xlsx:Data!F122 | 355000 | 355000 | 0.0 | PASS |
| Nilai gabungan 123 | solusi_10.xlsx:Data!F123 | 561000 | 561000 | 0.0 | PASS |
| Nilai gabungan 124 | solusi_10.xlsx:Data!F124 | 212000 | 212000 | 0.0 | PASS |
| Nilai gabungan 125 | solusi_10.xlsx:Data!F125 | 424000 | 424000 | 0.0 | PASS |
| Nilai gabungan 126 | solusi_10.xlsx:Data!F126 | 348000 | 348000 | 0.0 | PASS |
| Nilai gabungan 127 | solusi_10.xlsx:Data!F127 | 377000 | 377000 | 0.0 | PASS |
| Nilai gabungan 128 | solusi_10.xlsx:Data!F128 | 399000 | 399000 | 0.0 | PASS |
| Nilai gabungan 129 | solusi_10.xlsx:Data!F129 | 795000 | 795000 | 0.0 | PASS |
| Nilai gabungan 130 | solusi_10.xlsx:Data!F130 | 174000 | 174000 | 0.0 | PASS |
| Nilai gabungan 131 | solusi_10.xlsx:Data!F131 | 203000 | 203000 | 0.0 | PASS |
| Nilai gabungan 132 | solusi_10.xlsx:Data!F132 | 170000 | 170000 | 0.0 | PASS |
| Nilai gabungan 133 | solusi_10.xlsx:Data!F133 | 174000 | 174000 | 0.0 | PASS |
| Nilai gabungan 134 | solusi_10.xlsx:Data!F134 | 377000 | 377000 | 0.0 | PASS |
| Nilai gabungan 135 | solusi_10.xlsx:Data!F135 | 171000 | 171000 | 0.0 | PASS |
| Nilai gabungan 136 | solusi_10.xlsx:Data!F136 | 952000 | 952000 | 0.0 | PASS |
| Nilai gabungan 137 | solusi_10.xlsx:Data!F137 | 612000 | 612000 | 0.0 | PASS |
| Nilai gabungan 138 | solusi_10.xlsx:Data!F138 | 644000 | 644000 | 0.0 | PASS |
| Nilai gabungan 139 | solusi_10.xlsx:Data!F139 | 29000 | 29000 | 0.0 | PASS |
| Nilai gabungan 140 | solusi_10.xlsx:Data!F140 | 285000 | 285000 | 0.0 | PASS |
| Nilai gabungan 141 | solusi_10.xlsx:Data!F141 | 360000 | 360000 | 0.0 | PASS |
| Nilai gabungan 142 | solusi_10.xlsx:Data!F142 | 136000 | 136000 | 0.0 | PASS |
| Nilai gabungan 143 | solusi_10.xlsx:Data!F143 | 232000 | 232000 | 0.0 | PASS |
| Nilai gabungan 144 | solusi_10.xlsx:Data!F144 | 306000 | 306000 | 0.0 | PASS |
| Nilai gabungan 145 | solusi_10.xlsx:Data!F145 | 690000 | 690000 | 0.0 | PASS |
| Nilai gabungan 146 | solusi_10.xlsx:Data!F146 | 627000 | 627000 | 0.0 | PASS |
| Nilai gabungan 147 | solusi_10.xlsx:Data!F147 | 87000 | 87000 | 0.0 | PASS |
| Nilai gabungan 148 | solusi_10.xlsx:Data!F148 | 238000 | 238000 | 0.0 | PASS |
| Nilai gabungan 149 | solusi_10.xlsx:Data!F149 | 318000 | 318000 | 0.0 | PASS |
| Nilai gabungan 150 | solusi_10.xlsx:Data!F150 | 795000 | 795000 | 0.0 | PASS |
| Nilai gabungan 151 | solusi_10.xlsx:Data!F151 | 689000 | 689000 | 0.0 | PASS |
| Nilai gabungan 152 | solusi_10.xlsx:Data!F152 | 638000 | 638000 | 0.0 | PASS |
| Nilai gabungan 153 | solusi_10.xlsx:Data!F153 | 116000 | 116000 | 0.0 | PASS |
| Nilai gabungan 154 | solusi_10.xlsx:Data!F154 | 270000 | 270000 | 0.0 | PASS |
| Nilai gabungan 155 | solusi_10.xlsx:Data!F155 | 513000 | 513000 | 0.0 | PASS |
| Nilai gabungan 156 | solusi_10.xlsx:Data!F156 | 477000 | 477000 | 0.0 | PASS |
| Nilai gabungan 157 | solusi_10.xlsx:Data!F157 | 106000 | 106000 | 0.0 | PASS |
| Nilai gabungan 158 | solusi_10.xlsx:Data!F158 | 1020000 | 1020000 | 0.0 | PASS |
| Nilai gabungan 159 | solusi_10.xlsx:Data!F159 | 348000 | 348000 | 0.0 | PASS |
| Nilai gabungan 160 | solusi_10.xlsx:Data!F160 | 102000 | 102000 | 0.0 | PASS |
| Nilai gabungan 161 | solusi_10.xlsx:Data!F161 | 348000 | 348000 | 0.0 | PASS |
| Nilai gabungan 162 | solusi_10.xlsx:Data!F162 | 102000 | 102000 | 0.0 | PASS |
| Nilai gabungan 163 | solusi_10.xlsx:Data!F163 | 424000 | 424000 | 0.0 | PASS |
| Nilai gabungan 164 | solusi_10.xlsx:Data!F164 | 476000 | 476000 | 0.0 | PASS |
| Nilai gabungan 165 | solusi_10.xlsx:Data!F165 | 58000 | 58000 | 0.0 | PASS |
| Nilai gabungan 166 | solusi_10.xlsx:Data!F166 | 261000 | 261000 | 0.0 | PASS |
| Nilai gabungan 167 | solusi_10.xlsx:Data!F167 | 414000 | 414000 | 0.0 | PASS |
| Nilai gabungan 168 | solusi_10.xlsx:Data!F168 | 424000 | 424000 | 0.0 | PASS |
| Nilai gabungan 169 | solusi_10.xlsx:Data!F169 | 212000 | 212000 | 0.0 | PASS |
| Nilai gabungan 170 | solusi_10.xlsx:Data!F170 | 583000 | 583000 | 0.0 | PASS |
| Nilai gabungan 171 | solusi_10.xlsx:Data!F171 | 136000 | 136000 | 0.0 | PASS |
| Nilai gabungan 172 | solusi_10.xlsx:Data!F172 | 29000 | 29000 | 0.0 | PASS |
| Nilai gabungan 173 | solusi_10.xlsx:Data!F173 | 261000 | 261000 | 0.0 | PASS |
| Nilai gabungan 174 | solusi_10.xlsx:Data!F174 | 136000 | 136000 | 0.0 | PASS |
| Nilai gabungan 175 | solusi_10.xlsx:Data!F175 | 53000 | 53000 | 0.0 | PASS |
| Nilai gabungan 176 | solusi_10.xlsx:Data!F176 | 238000 | 238000 | 0.0 | PASS |
| Nilai gabungan 177 | solusi_10.xlsx:Data!F177 | 684000 | 684000 | 0.0 | PASS |
| Nilai gabungan 178 | solusi_10.xlsx:Data!F178 | 135000 | 135000 | 0.0 | PASS |
| Nilai gabungan 179 | solusi_10.xlsx:Data!F179 | 748000 | 748000 | 0.0 | PASS |
| Nilai gabungan 180 | solusi_10.xlsx:Data!F180 | 29000 | 29000 | 0.0 | PASS |
| Nilai gabungan 181 | solusi_10.xlsx:Data!F181 | 51000 | 51000 | 0.0 | PASS |
| Nilai gabungan 182 | solusi_10.xlsx:Data!F182 | 53000 | 53000 | 0.0 | PASS |
| Nilai gabungan 183 | solusi_10.xlsx:Data!F183 | 406000 | 406000 | 0.0 | PASS |
| Nilai gabungan 184 | solusi_10.xlsx:Data!F184 | 884000 | 884000 | 0.0 | PASS |
| Nilai gabungan 185 | solusi_10.xlsx:Data!F185 | 51000 | 51000 | 0.0 | PASS |
| Nilai gabungan 186 | solusi_10.xlsx:Data!F186 | 230000 | 230000 | 0.0 | PASS |
| Nilai gabungan 187 | solusi_10.xlsx:Data!F187 | 58000 | 58000 | 0.0 | PASS |
| Nilai gabungan 188 | solusi_10.xlsx:Data!F188 | 85000 | 85000 | 0.0 | PASS |
| Nilai gabungan 189 | solusi_10.xlsx:Data!F189 | 368000 | 368000 | 0.0 | PASS |
| Nilai gabungan 190 | solusi_10.xlsx:Data!F190 | 570000 | 570000 | 0.0 | PASS |
| Nilai gabungan 191 | solusi_10.xlsx:Data!F191 | 203000 | 203000 | 0.0 | PASS |
| Nilai gabungan 192 | solusi_10.xlsx:Data!F192 | 204000 | 204000 | 0.0 | PASS |
| Nilai gabungan 193 | solusi_10.xlsx:Data!F193 | 116000 | 116000 | 0.0 | PASS |
| Nilai gabungan 194 | solusi_10.xlsx:Data!F194 | 406000 | 406000 | 0.0 | PASS |
| Nilai gabungan 195 | solusi_10.xlsx:Data!F195 | 290000 | 290000 | 0.0 | PASS |
| Nilai gabungan 196 | solusi_10.xlsx:Data!F196 | 522000 | 522000 | 0.0 | PASS |
| Nilai gabungan 197 | solusi_10.xlsx:Data!F197 | 781000 | 781000 | 0.0 | PASS |
| Nilai gabungan 198 | solusi_10.xlsx:Data!F198 | 342000 | 342000 | 0.0 | PASS |
| Nilai gabungan 199 | solusi_10.xlsx:Data!F199 | 348000 | 348000 | 0.0 | PASS |
| Nilai gabungan 200 | solusi_10.xlsx:Data!F200 | 51000 | 51000 | 0.0 | PASS |
| Nilai gabungan 201 | solusi_10.xlsx:Data!F201 | 58000 | 58000 | 0.0 | PASS |

Ringkasan: 2 file, total error = 0, rekonsiliasi gagal = 0.

## Catatan

- Sel kuning pada latihan sengaja kosong. Semua perhitungan pada solusi memakai rumus hidup.
- Pembanding dibaca dari CSV dan dihitung dengan pandas/scipy; statistik dibandingkan sampai enam desimal, selisih harus nol.
- Preview berasal dari render nyata sheet Hasil_Query_Diharapkan ke PDF dengan LibreOffice, kemudian PNG melalui pdftoppm.
- Fungsi diperiksa terhadap daftar yang diperbolehkan; tidak ada rumus array dinamis atau tautan workbook eksternal.
- Interaksi manual di Excel dan Google Sheets belum diuji langsung; fitur yang memerlukan langkah manual dijelaskan jujur di README.
