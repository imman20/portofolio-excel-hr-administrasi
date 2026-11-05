# Verifikasi modul `03_formula-dasar`

Tanggal verifikasi: 2026-11-05 (Asia/Jakarta). Mesin: LibreOffice headless (recalc) + openpyxl + pandas/scipy.
Skrip: `python tools/course_runtime.py 03_formula-dasar` dari folder `belajar/`.

## Hasil recalc per file

| File | Sheet | Jumlah rumus | Jumlah error | Status recalc | Ukuran (KB) |
|---|---|---:|---:|---|---:|
| `latihan_03.xlsx` | Instruksi, Data, Latihan, Master | 720 | 0 | success | 27.0 |
| `solusi_03.xlsx` | Instruksi, Data, Latihan, Master | 749 | 0 | success | 27.3 |

## Rekonsiliasi pandas vs sel hasil recalc

| Pemeriksaan | Lokasi sel | Nilai pandas | Nilai sel | Selisih | Status |
|---|---|---:|---:|---:|---|
| Total bruto | solusi_03.xlsx:Latihan!C4 | 60540000.0 | 60540000 | 0.0 | PASS |
| Total neto | solusi_03.xlsx:Latihan!C5 | 57416875.0 | 57416875 | 0.0 | PASS |
| Pajak | solusi_03.xlsx:Latihan!C6 | 6315878.0 | 6315878 | 0.0 | PASS |
| Tagihan | solusi_03.xlsx:Latihan!C7 | 63732753.0 | 63732753 | 0.0 | PASS |
| Mean neto | solusi_03.xlsx:Latihan!C8 | 318982.6388888889 | 318982.638888889 | 0.0 | PASS |
| Minimum | solusi_03.xlsx:Latihan!C9 | 11875.0 | 11875 | 0.0 | PASS |
| Maksimum | solusi_03.xlsx:Latihan!C10 | 1211250.0 | 1211250 | 0.0 | PASS |
| COUNT | solusi_03.xlsx:Latihan!C11 | 180.0 | 180 | 0.0 | PASS |
| COUNTA | solusi_03.xlsx:Latihan!C12 | 180.0 | 180 | 0.0 | PASS |
| COUNTBLANK | solusi_03.xlsx:Latihan!C13 | 144.0 | 144 | 0.0 | PASS |
| ROUND | solusi_03.xlsx:Latihan!C14 | 318983.0 | 318983 | 0.0 | PASS |
| Prioritas kali | solusi_03.xlsx:Latihan!C15 | 14.0 | 14 | 0.0 | PASS |
| Prioritas kurung | solusi_03.xlsx:Latihan!C16 | 20.0 | 20 | 0.0 | PASS |
| Neto baris 2 | solusi_03.xlsx:Data!I2 | 45000.0 | 45000 | 0.0 | PASS |
| Neto baris 3 | solusi_03.xlsx:Data!I3 | 330000.0 | 330000 | 0.0 | PASS |
| Neto baris 4 | solusi_03.xlsx:Data!I4 | 306000.0 | 306000 | 0.0 | PASS |
| Neto baris 5 | solusi_03.xlsx:Data!I5 | 630000.0 | 630000 | 0.0 | PASS |
| Neto baris 6 | solusi_03.xlsx:Data!I6 | 154375.0 | 154375 | 0.0 | PASS |
| Neto baris 7 | solusi_03.xlsx:Data!I7 | 112500.0 | 112500 | 0.0 | PASS |
| Neto baris 8 | solusi_03.xlsx:Data!I8 | 324000.0 | 324000 | 0.0 | PASS |
| Neto baris 9 | solusi_03.xlsx:Data!I9 | 49500.0 | 49500 | 0.0 | PASS |
| Neto baris 10 | solusi_03.xlsx:Data!I10 | 153000.0 | 153000 | 0.0 | PASS |
| Neto baris 11 | solusi_03.xlsx:Data!I11 | 112500.0 | 112500 | 0.0 | PASS |
| Neto baris 12 | solusi_03.xlsx:Data!I12 | 33750.0 | 33750 | 0.0 | PASS |
| Neto baris 13 | solusi_03.xlsx:Data!I13 | 87500.0 | 87500 | 0.0 | PASS |
| Neto baris 14 | solusi_03.xlsx:Data!I14 | 272250.0 | 272250 | 0.0 | PASS |
| Neto baris 15 | solusi_03.xlsx:Data!I15 | 567000.0 | 567000 | 0.0 | PASS |
| Neto baris 16 | solusi_03.xlsx:Data!I16 | 1211250.0 | 1211250 | 0.0 | PASS |
| Neto baris 17 | solusi_03.xlsx:Data!I17 | 484500.0 | 484500 | 0.0 | PASS |
| Neto baris 18 | solusi_03.xlsx:Data!I18 | 360000.0 | 360000 | 0.0 | PASS |
| Neto baris 19 | solusi_03.xlsx:Data!I19 | 850000.0 | 850000 | 0.0 | PASS |
| Neto baris 20 | solusi_03.xlsx:Data!I20 | 22500.0 | 22500 | 0.0 | PASS |
| Neto baris 21 | solusi_03.xlsx:Data!I21 | 384750.0 | 384750 | 0.0 | PASS |
| Neto baris 22 | solusi_03.xlsx:Data!I22 | 459000.0 | 459000 | 0.0 | PASS |
| Neto baris 23 | solusi_03.xlsx:Data!I23 | 52250.0 | 52250 | 0.0 | PASS |
| Neto baris 24 | solusi_03.xlsx:Data!I24 | 161500.0 | 161500 | 0.0 | PASS |
| Neto baris 25 | solusi_03.xlsx:Data!I25 | 137500.0 | 137500 | 0.0 | PASS |
| Neto baris 26 | solusi_03.xlsx:Data!I26 | 56250.0 | 56250 | 0.0 | PASS |
| Neto baris 27 | solusi_03.xlsx:Data!I27 | 535500.0 | 535500 | 0.0 | PASS |
| Neto baris 28 | solusi_03.xlsx:Data!I28 | 384750.0 | 384750 | 0.0 | PASS |
| Neto baris 29 | solusi_03.xlsx:Data!I29 | 324000.0 | 324000 | 0.0 | PASS |
| Neto baris 30 | solusi_03.xlsx:Data!I30 | 153000.0 | 153000 | 0.0 | PASS |
| Neto baris 31 | solusi_03.xlsx:Data!I31 | 283500.0 | 283500 | 0.0 | PASS |
| Neto baris 32 | solusi_03.xlsx:Data!I32 | 74250.0 | 74250 | 0.0 | PASS |
| Neto baris 33 | solusi_03.xlsx:Data!I33 | 11875.0 | 11875 | 0.0 | PASS |
| Neto baris 34 | solusi_03.xlsx:Data!I34 | 83125.0 | 83125 | 0.0 | PASS |
| Neto baris 35 | solusi_03.xlsx:Data!I35 | 49500.0 | 49500 | 0.0 | PASS |
| Neto baris 36 | solusi_03.xlsx:Data!I36 | 198000.0 | 198000 | 0.0 | PASS |
| Neto baris 37 | solusi_03.xlsx:Data!I37 | 209000.0 | 209000 | 0.0 | PASS |
| Neto baris 38 | solusi_03.xlsx:Data!I38 | 726750.0 | 726750 | 0.0 | PASS |
| Neto baris 39 | solusi_03.xlsx:Data!I39 | 192500.0 | 192500 | 0.0 | PASS |
| Neto baris 40 | solusi_03.xlsx:Data!I40 | 567000.0 | 567000 | 0.0 | PASS |
| Neto baris 41 | solusi_03.xlsx:Data!I41 | 315000.0 | 315000 | 0.0 | PASS |
| Neto baris 42 | solusi_03.xlsx:Data!I42 | 130625.0 | 130625 | 0.0 | PASS |
| Neto baris 43 | solusi_03.xlsx:Data!I43 | 171000.0 | 171000 | 0.0 | PASS |
| Neto baris 44 | solusi_03.xlsx:Data!I44 | 45000.0 | 45000 | 0.0 | PASS |
| Neto baris 45 | solusi_03.xlsx:Data!I45 | 235125.0 | 235125 | 0.0 | PASS |
| Neto baris 46 | solusi_03.xlsx:Data!I46 | 130625.0 | 130625 | 0.0 | PASS |
| Neto baris 47 | solusi_03.xlsx:Data!I47 | 315000.0 | 315000 | 0.0 | PASS |
| Neto baris 48 | solusi_03.xlsx:Data!I48 | 135000.0 | 135000 | 0.0 | PASS |
| Neto baris 49 | solusi_03.xlsx:Data!I49 | 95000.0 | 95000 | 0.0 | PASS |
| Neto baris 50 | solusi_03.xlsx:Data!I50 | 213750.0 | 213750 | 0.0 | PASS |
| Neto baris 51 | solusi_03.xlsx:Data!I51 | 765000.0 | 765000 | 0.0 | PASS |
| Neto baris 52 | solusi_03.xlsx:Data!I52 | 161500.0 | 161500 | 0.0 | PASS |
| Neto baris 53 | solusi_03.xlsx:Data!I53 | 607500.0 | 607500 | 0.0 | PASS |
| Neto baris 54 | solusi_03.xlsx:Data!I54 | 71250.0 | 71250 | 0.0 | PASS |
| Neto baris 55 | solusi_03.xlsx:Data!I55 | 130625.0 | 130625 | 0.0 | PASS |
| Neto baris 56 | solusi_03.xlsx:Data!I56 | 78750.0 | 78750 | 0.0 | PASS |
| Neto baris 57 | solusi_03.xlsx:Data!I57 | 275000.0 | 275000 | 0.0 | PASS |
| Neto baris 58 | solusi_03.xlsx:Data!I58 | 49500.0 | 49500 | 0.0 | PASS |
| Neto baris 59 | solusi_03.xlsx:Data!I59 | 302500.0 | 302500 | 0.0 | PASS |
| Neto baris 60 | solusi_03.xlsx:Data!I60 | 313500.0 | 313500 | 0.0 | PASS |
| Neto baris 61 | solusi_03.xlsx:Data!I61 | 130625.0 | 130625 | 0.0 | PASS |
| Neto baris 62 | solusi_03.xlsx:Data!I62 | 1049750.0 | 1049750 | 0.0 | PASS |
| Neto baris 63 | solusi_03.xlsx:Data!I63 | 513000.0 | 513000 | 0.0 | PASS |
| Neto baris 64 | solusi_03.xlsx:Data!I64 | 1211250.0 | 1211250 | 0.0 | PASS |
| Neto baris 65 | solusi_03.xlsx:Data!I65 | 162000.0 | 162000 | 0.0 | PASS |
| Neto baris 66 | solusi_03.xlsx:Data!I66 | 162500.0 | 162500 | 0.0 | PASS |
| Neto baris 67 | solusi_03.xlsx:Data!I67 | 272250.0 | 272250 | 0.0 | PASS |
| Neto baris 68 | solusi_03.xlsx:Data!I68 | 62500.0 | 62500 | 0.0 | PASS |
| Neto baris 69 | solusi_03.xlsx:Data!I69 | 171000.0 | 171000 | 0.0 | PASS |
| Neto baris 70 | solusi_03.xlsx:Data!I70 | 726750.0 | 726750 | 0.0 | PASS |
| Neto baris 71 | solusi_03.xlsx:Data!I71 | 162500.0 | 162500 | 0.0 | PASS |
| Neto baris 72 | solusi_03.xlsx:Data!I72 | 180000.0 | 180000 | 0.0 | PASS |
| Neto baris 73 | solusi_03.xlsx:Data!I73 | 67500.0 | 67500 | 0.0 | PASS |
| Neto baris 74 | solusi_03.xlsx:Data!I74 | 918000.0 | 918000 | 0.0 | PASS |
| Neto baris 75 | solusi_03.xlsx:Data!I75 | 1190000.0 | 1190000 | 0.0 | PASS |
| Neto baris 76 | solusi_03.xlsx:Data!I76 | 112500.0 | 112500 | 0.0 | PASS |
| Neto baris 77 | solusi_03.xlsx:Data!I77 | 330000.0 | 330000 | 0.0 | PASS |
| Neto baris 78 | solusi_03.xlsx:Data!I78 | 128250.0 | 128250 | 0.0 | PASS |
| Neto baris 79 | solusi_03.xlsx:Data!I79 | 1147500.0 | 1147500 | 0.0 | PASS |
| Neto baris 80 | solusi_03.xlsx:Data!I80 | 67500.0 | 67500 | 0.0 | PASS |
| Neto baris 81 | solusi_03.xlsx:Data!I81 | 918000.0 | 918000 | 0.0 | PASS |
| Neto baris 82 | solusi_03.xlsx:Data!I82 | 242250.0 | 242250 | 0.0 | PASS |
| Neto baris 83 | solusi_03.xlsx:Data!I83 | 510000.0 | 510000 | 0.0 | PASS |
| Neto baris 84 | solusi_03.xlsx:Data!I84 | 24750.0 | 24750 | 0.0 | PASS |
| Neto baris 85 | solusi_03.xlsx:Data!I85 | 37500.0 | 37500 | 0.0 | PASS |
| Neto baris 86 | solusi_03.xlsx:Data!I86 | 71250.0 | 71250 | 0.0 | PASS |
| Neto baris 87 | solusi_03.xlsx:Data!I87 | 346500.0 | 346500 | 0.0 | PASS |
| Neto baris 88 | solusi_03.xlsx:Data!I88 | 371250.0 | 371250 | 0.0 | PASS |
| Neto baris 89 | solusi_03.xlsx:Data!I89 | 80750.0 | 80750 | 0.0 | PASS |
| Neto baris 90 | solusi_03.xlsx:Data!I90 | 470250.0 | 470250 | 0.0 | PASS |
| Neto baris 91 | solusi_03.xlsx:Data!I91 | 168750.0 | 168750 | 0.0 | PASS |
| Neto baris 92 | solusi_03.xlsx:Data!I92 | 299250.0 | 299250 | 0.0 | PASS |
| Neto baris 93 | solusi_03.xlsx:Data!I93 | 229500.0 | 229500 | 0.0 | PASS |
| Neto baris 94 | solusi_03.xlsx:Data!I94 | 330000.0 | 330000 | 0.0 | PASS |
| Neto baris 95 | solusi_03.xlsx:Data!I95 | 112500.0 | 112500 | 0.0 | PASS |
| Neto baris 96 | solusi_03.xlsx:Data!I96 | 364500.0 | 364500 | 0.0 | PASS |
| Neto baris 97 | solusi_03.xlsx:Data!I97 | 646000.0 | 646000 | 0.0 | PASS |
| Neto baris 98 | solusi_03.xlsx:Data!I98 | 306000.0 | 306000 | 0.0 | PASS |
| Neto baris 99 | solusi_03.xlsx:Data!I99 | 272250.0 | 272250 | 0.0 | PASS |
| Neto baris 100 | solusi_03.xlsx:Data!I100 | 118750.0 | 118750 | 0.0 | PASS |
| Neto baris 101 | solusi_03.xlsx:Data!I101 | 78375.0 | 78375 | 0.0 | PASS |
| Neto baris 102 | solusi_03.xlsx:Data!I102 | 148500.0 | 148500 | 0.0 | PASS |
| Neto baris 103 | solusi_03.xlsx:Data!I103 | 256500.0 | 256500 | 0.0 | PASS |
| Neto baris 104 | solusi_03.xlsx:Data!I104 | 67500.0 | 67500 | 0.0 | PASS |
| Neto baris 105 | solusi_03.xlsx:Data!I105 | 106875.0 | 106875 | 0.0 | PASS |
| Neto baris 106 | solusi_03.xlsx:Data!I106 | 641250.0 | 641250 | 0.0 | PASS |
| Neto baris 107 | solusi_03.xlsx:Data!I107 | 513000.0 | 513000 | 0.0 | PASS |
| Neto baris 108 | solusi_03.xlsx:Data!I108 | 125000.0 | 125000 | 0.0 | PASS |
| Neto baris 109 | solusi_03.xlsx:Data!I109 | 384750.0 | 384750 | 0.0 | PASS |
| Neto baris 110 | solusi_03.xlsx:Data!I110 | 161500.0 | 161500 | 0.0 | PASS |
| Neto baris 111 | solusi_03.xlsx:Data!I111 | 935000.0 | 935000 | 0.0 | PASS |
| Neto baris 112 | solusi_03.xlsx:Data!I112 | 850000.0 | 850000 | 0.0 | PASS |
| Neto baris 113 | solusi_03.xlsx:Data!I113 | 99000.0 | 99000 | 0.0 | PASS |
| Neto baris 114 | solusi_03.xlsx:Data!I114 | 1190000.0 | 1190000 | 0.0 | PASS |
| Neto baris 115 | solusi_03.xlsx:Data!I115 | 565250.0 | 565250 | 0.0 | PASS |
| Neto baris 116 | solusi_03.xlsx:Data!I116 | 55000.0 | 55000 | 0.0 | PASS |
| Neto baris 117 | solusi_03.xlsx:Data!I117 | 125000.0 | 125000 | 0.0 | PASS |
| Neto baris 118 | solusi_03.xlsx:Data!I118 | 261250.0 | 261250 | 0.0 | PASS |
| Neto baris 119 | solusi_03.xlsx:Data!I119 | 484500.0 | 484500 | 0.0 | PASS |
| Neto baris 120 | solusi_03.xlsx:Data!I120 | 888250.0 | 888250 | 0.0 | PASS |
| Neto baris 121 | solusi_03.xlsx:Data!I121 | 969000.0 | 969000 | 0.0 | PASS |
| Neto baris 122 | solusi_03.xlsx:Data!I122 | 243000.0 | 243000 | 0.0 | PASS |
| Neto baris 123 | solusi_03.xlsx:Data!I123 | 357500.0 | 357500 | 0.0 | PASS |
| Neto baris 124 | solusi_03.xlsx:Data!I124 | 146250.0 | 146250 | 0.0 | PASS |
| Neto baris 125 | solusi_03.xlsx:Data!I125 | 80750.0 | 80750 | 0.0 | PASS |
| Neto baris 126 | solusi_03.xlsx:Data!I126 | 62500.0 | 62500 | 0.0 | PASS |
| Neto baris 127 | solusi_03.xlsx:Data!I127 | 182875.0 | 182875 | 0.0 | PASS |
| Neto baris 128 | solusi_03.xlsx:Data!I128 | 297000.0 | 297000 | 0.0 | PASS |
| Neto baris 129 | solusi_03.xlsx:Data!I129 | 99000.0 | 99000 | 0.0 | PASS |
| Neto baris 130 | solusi_03.xlsx:Data!I130 | 45000.0 | 45000 | 0.0 | PASS |
| Neto baris 131 | solusi_03.xlsx:Data!I131 | 275000.0 | 275000 | 0.0 | PASS |
| Neto baris 132 | solusi_03.xlsx:Data!I132 | 346500.0 | 346500 | 0.0 | PASS |
| Neto baris 133 | solusi_03.xlsx:Data!I133 | 1130500.0 | 1130500 | 0.0 | PASS |
| Neto baris 134 | solusi_03.xlsx:Data!I134 | 23750.0 | 23750 | 0.0 | PASS |
| Neto baris 135 | solusi_03.xlsx:Data!I135 | 62500.0 | 62500 | 0.0 | PASS |
| Neto baris 136 | solusi_03.xlsx:Data!I136 | 680000.0 | 680000 | 0.0 | PASS |
| Neto baris 137 | solusi_03.xlsx:Data!I137 | 67500.0 | 67500 | 0.0 | PASS |
| Neto baris 138 | solusi_03.xlsx:Data!I138 | 166250.0 | 166250 | 0.0 | PASS |
| Neto baris 139 | solusi_03.xlsx:Data!I139 | 78750.0 | 78750 | 0.0 | PASS |
| Neto baris 140 | solusi_03.xlsx:Data!I140 | 340000.0 | 340000 | 0.0 | PASS |
| Neto baris 141 | solusi_03.xlsx:Data!I141 | 324000.0 | 324000 | 0.0 | PASS |
| Neto baris 142 | solusi_03.xlsx:Data!I142 | 11875.0 | 11875 | 0.0 | PASS |
| Neto baris 143 | solusi_03.xlsx:Data!I143 | 646000.0 | 646000 | 0.0 | PASS |
| Neto baris 144 | solusi_03.xlsx:Data!I144 | 45000.0 | 45000 | 0.0 | PASS |
| Neto baris 145 | solusi_03.xlsx:Data!I145 | 198000.0 | 198000 | 0.0 | PASS |
| Neto baris 146 | solusi_03.xlsx:Data!I146 | 382500.0 | 382500 | 0.0 | PASS |
| Neto baris 147 | solusi_03.xlsx:Data!I147 | 162500.0 | 162500 | 0.0 | PASS |
| Neto baris 148 | solusi_03.xlsx:Data!I148 | 87500.0 | 87500 | 0.0 | PASS |
| Neto baris 149 | solusi_03.xlsx:Data!I149 | 425000.0 | 425000 | 0.0 | PASS |
| Neto baris 150 | solusi_03.xlsx:Data!I150 | 213750.0 | 213750 | 0.0 | PASS |
| Neto baris 151 | solusi_03.xlsx:Data!I151 | 680000.0 | 680000 | 0.0 | PASS |
| Neto baris 152 | solusi_03.xlsx:Data!I152 | 173250.0 | 173250 | 0.0 | PASS |
| Neto baris 153 | solusi_03.xlsx:Data!I153 | 37500.0 | 37500 | 0.0 | PASS |
| Neto baris 154 | solusi_03.xlsx:Data!I154 | 364500.0 | 364500 | 0.0 | PASS |
| Neto baris 155 | solusi_03.xlsx:Data!I155 | 71250.0 | 71250 | 0.0 | PASS |
| Neto baris 156 | solusi_03.xlsx:Data!I156 | 540000.0 | 540000 | 0.0 | PASS |
| Neto baris 157 | solusi_03.xlsx:Data!I157 | 315000.0 | 315000 | 0.0 | PASS |
| Neto baris 158 | solusi_03.xlsx:Data!I158 | 81000.0 | 81000 | 0.0 | PASS |
| Neto baris 159 | solusi_03.xlsx:Data!I159 | 371250.0 | 371250 | 0.0 | PASS |
| Neto baris 160 | solusi_03.xlsx:Data!I160 | 42750.0 | 42750 | 0.0 | PASS |
| Neto baris 161 | solusi_03.xlsx:Data!I161 | 135000.0 | 135000 | 0.0 | PASS |
| Neto baris 162 | solusi_03.xlsx:Data!I162 | 47500.0 | 47500 | 0.0 | PASS |
| Neto baris 163 | solusi_03.xlsx:Data!I163 | 171000.0 | 171000 | 0.0 | PASS |
| Neto baris 164 | solusi_03.xlsx:Data!I164 | 567000.0 | 567000 | 0.0 | PASS |
| Neto baris 165 | solusi_03.xlsx:Data!I165 | 1071000.0 | 1071000 | 0.0 | PASS |
| Neto baris 166 | solusi_03.xlsx:Data!I166 | 382500.0 | 382500 | 0.0 | PASS |
| Neto baris 167 | solusi_03.xlsx:Data!I167 | 357500.0 | 357500 | 0.0 | PASS |
| Neto baris 168 | solusi_03.xlsx:Data!I168 | 450000.0 | 450000 | 0.0 | PASS |
| Neto baris 169 | solusi_03.xlsx:Data!I169 | 598500.0 | 598500 | 0.0 | PASS |
| Neto baris 170 | solusi_03.xlsx:Data!I170 | 427500.0 | 427500 | 0.0 | PASS |
| Neto baris 171 | solusi_03.xlsx:Data!I171 | 1020000.0 | 1020000 | 0.0 | PASS |
| Neto baris 172 | solusi_03.xlsx:Data!I172 | 247500.0 | 247500 | 0.0 | PASS |
| Neto baris 173 | solusi_03.xlsx:Data!I173 | 121500.0 | 121500 | 0.0 | PASS |
| Neto baris 174 | solusi_03.xlsx:Data!I174 | 646000.0 | 646000 | 0.0 | PASS |
| Neto baris 175 | solusi_03.xlsx:Data!I175 | 137500.0 | 137500 | 0.0 | PASS |
| Neto baris 176 | solusi_03.xlsx:Data!I176 | 156750.0 | 156750 | 0.0 | PASS |
| Neto baris 177 | solusi_03.xlsx:Data!I177 | 85500.0 | 85500 | 0.0 | PASS |
| Neto baris 178 | solusi_03.xlsx:Data!I178 | 106875.0 | 106875 | 0.0 | PASS |
| Neto baris 179 | solusi_03.xlsx:Data!I179 | 125000.0 | 125000 | 0.0 | PASS |
| Neto baris 180 | solusi_03.xlsx:Data!I180 | 403750.0 | 403750 | 0.0 | PASS |
| Neto baris 181 | solusi_03.xlsx:Data!I181 | 1020000.0 | 1020000 | 0.0 | PASS |
| Referensi campuran | solusi_03.xlsx:Latihan!E32 | 405000 | 405000 | 0.0 | PASS |
| Komponen konsisten | solusi_03.xlsx:Latihan!C35 | 0 | 0 | 0.0 | PASS |

Ringkasan: 2 file, total error = 0, rekonsiliasi gagal = 0.

## Catatan

- Sel kuning pada latihan sengaja kosong. Semua perhitungan pada solusi memakai rumus hidup.
- Pembanding dibaca dari CSV dan dihitung dengan pandas/scipy; statistik dibandingkan sampai enam desimal, selisih harus nol.
- Preview berasal dari render nyata sheet Latihan ke PDF dengan LibreOffice, kemudian PNG melalui pdftoppm.
- Fungsi diperiksa terhadap daftar yang diperbolehkan; tidak ada rumus array dinamis atau tautan workbook eksternal.
- Interaksi manual di Excel dan Google Sheets belum diuji langsung; fitur yang memerlukan langkah manual dijelaskan jujur di README.
