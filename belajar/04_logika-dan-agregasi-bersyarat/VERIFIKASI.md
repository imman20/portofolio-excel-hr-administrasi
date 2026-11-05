# Verifikasi modul `04_logika-dan-agregasi-bersyarat`

Tanggal verifikasi: 2026-11-05 (Asia/Jakarta). Mesin: LibreOffice headless (recalc) + openpyxl + pandas/scipy.
Skrip: `python tools/course_runtime.py 04_logika-dan-agregasi-bersyarat` dari folder `belajar/`.

## Hasil recalc per file

| File | Sheet | Jumlah rumus | Jumlah error | Status recalc | Ukuran (KB) |
|---|---|---:|---:|---|---:|
| `latihan_04.xlsx` | Instruksi, Data, Latihan, Master | 880 | 0 | success | 31.6 |
| `solusi_04.xlsx` | Instruksi, Data, Latihan, Master | 893 | 0 | success | 32.0 |

## Rekonsiliasi pandas vs sel hasil recalc

| Pemeriksaan | Lokasi sel | Nilai pandas | Nilai sel | Selisih | Status |
|---|---|---:|---:|---:|---|
| Ringkasan 4 | solusi_04.xlsx:Latihan!C4 | 67 | 67 | 0.0 | PASS |
| Ringkasan 5 | solusi_04.xlsx:Latihan!C5 | 23520000 | 23520000 | 0.0 | PASS |
| Ringkasan 6 | solusi_04.xlsx:Latihan!C6 | 36.47761194029851 | 36.4776119402985 | -0.0 | PASS |
| Ringkasan 7 | solusi_04.xlsx:Latihan!C7 | 145 | 145 | 0.0 | PASS |
| Ringkasan 8 | solusi_04.xlsx:Latihan!C8 | 67 | 67 | 0.0 | PASS |
| Ringkasan 9 | solusi_04.xlsx:Latihan!C9 | 66 | 66 | 0.0 | PASS |
| Ringkasan 10 | solusi_04.xlsx:Latihan!C10 | 120 | 120 | 0.0 | PASS |
| Ringkasan 11 | solusi_04.xlsx:Latihan!C11 | 0 | 0 | 0.0 | PASS |
| Ringkasan 12 | solusi_04.xlsx:Latihan!C12 | 94520000 | 94520000 | 0.0 | PASS |
| Ringkasan 13 | solusi_04.xlsx:Latihan!C13 | 0.36464088397790057 | 0.364640883977901 | 0.0 | PASS |
| Ringkasan 14 | solusi_04.xlsx:Latihan!C14 | TIDAK ADA DATA | TIDAK ADA DATA | 0 | PASS |
| Aturan baris 2 | solusi_04.xlsx:Data!I2 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 3 | solusi_04.xlsx:Data!I3 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 4 | solusi_04.xlsx:Data!I4 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 5 | solusi_04.xlsx:Data!I5 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 6 | solusi_04.xlsx:Data!I6 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 7 | solusi_04.xlsx:Data!I7 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 8 | solusi_04.xlsx:Data!I8 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 9 | solusi_04.xlsx:Data!I9 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 10 | solusi_04.xlsx:Data!I10 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 11 | solusi_04.xlsx:Data!I11 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 12 | solusi_04.xlsx:Data!I12 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 13 | solusi_04.xlsx:Data!I13 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 14 | solusi_04.xlsx:Data!I14 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 15 | solusi_04.xlsx:Data!I15 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 16 | solusi_04.xlsx:Data!I16 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 17 | solusi_04.xlsx:Data!I17 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 18 | solusi_04.xlsx:Data!I18 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 19 | solusi_04.xlsx:Data!I19 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 20 | solusi_04.xlsx:Data!I20 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 21 | solusi_04.xlsx:Data!I21 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 22 | solusi_04.xlsx:Data!I22 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 23 | solusi_04.xlsx:Data!I23 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 24 | solusi_04.xlsx:Data!I24 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 25 | solusi_04.xlsx:Data!I25 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 26 | solusi_04.xlsx:Data!I26 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 27 | solusi_04.xlsx:Data!I27 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 28 | solusi_04.xlsx:Data!I28 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 29 | solusi_04.xlsx:Data!I29 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 30 | solusi_04.xlsx:Data!I30 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 31 | solusi_04.xlsx:Data!I31 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 32 | solusi_04.xlsx:Data!I32 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 33 | solusi_04.xlsx:Data!I33 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 34 | solusi_04.xlsx:Data!I34 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 35 | solusi_04.xlsx:Data!I35 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 36 | solusi_04.xlsx:Data!I36 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 37 | solusi_04.xlsx:Data!I37 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 38 | solusi_04.xlsx:Data!I38 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 39 | solusi_04.xlsx:Data!I39 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 40 | solusi_04.xlsx:Data!I40 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 41 | solusi_04.xlsx:Data!I41 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 42 | solusi_04.xlsx:Data!I42 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 43 | solusi_04.xlsx:Data!I43 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 44 | solusi_04.xlsx:Data!I44 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 45 | solusi_04.xlsx:Data!I45 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 46 | solusi_04.xlsx:Data!I46 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 47 | solusi_04.xlsx:Data!I47 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 48 | solusi_04.xlsx:Data!I48 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 49 | solusi_04.xlsx:Data!I49 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 50 | solusi_04.xlsx:Data!I50 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 51 | solusi_04.xlsx:Data!I51 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 52 | solusi_04.xlsx:Data!I52 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 53 | solusi_04.xlsx:Data!I53 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 54 | solusi_04.xlsx:Data!I54 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 55 | solusi_04.xlsx:Data!I55 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 56 | solusi_04.xlsx:Data!I56 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 57 | solusi_04.xlsx:Data!I57 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 58 | solusi_04.xlsx:Data!I58 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 59 | solusi_04.xlsx:Data!I59 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 60 | solusi_04.xlsx:Data!I60 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 61 | solusi_04.xlsx:Data!I61 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 62 | solusi_04.xlsx:Data!I62 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 63 | solusi_04.xlsx:Data!I63 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 64 | solusi_04.xlsx:Data!I64 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 65 | solusi_04.xlsx:Data!I65 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 66 | solusi_04.xlsx:Data!I66 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 67 | solusi_04.xlsx:Data!I67 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 68 | solusi_04.xlsx:Data!I68 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 69 | solusi_04.xlsx:Data!I69 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 70 | solusi_04.xlsx:Data!I70 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 71 | solusi_04.xlsx:Data!I71 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 72 | solusi_04.xlsx:Data!I72 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 73 | solusi_04.xlsx:Data!I73 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 74 | solusi_04.xlsx:Data!I74 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 75 | solusi_04.xlsx:Data!I75 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 76 | solusi_04.xlsx:Data!I76 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 77 | solusi_04.xlsx:Data!I77 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 78 | solusi_04.xlsx:Data!I78 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 79 | solusi_04.xlsx:Data!I79 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 80 | solusi_04.xlsx:Data!I80 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 81 | solusi_04.xlsx:Data!I81 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 82 | solusi_04.xlsx:Data!I82 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 83 | solusi_04.xlsx:Data!I83 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 84 | solusi_04.xlsx:Data!I84 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 85 | solusi_04.xlsx:Data!I85 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 86 | solusi_04.xlsx:Data!I86 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 87 | solusi_04.xlsx:Data!I87 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 88 | solusi_04.xlsx:Data!I88 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 89 | solusi_04.xlsx:Data!I89 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 90 | solusi_04.xlsx:Data!I90 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 91 | solusi_04.xlsx:Data!I91 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 92 | solusi_04.xlsx:Data!I92 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 93 | solusi_04.xlsx:Data!I93 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 94 | solusi_04.xlsx:Data!I94 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 95 | solusi_04.xlsx:Data!I95 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 96 | solusi_04.xlsx:Data!I96 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 97 | solusi_04.xlsx:Data!I97 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 98 | solusi_04.xlsx:Data!I98 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 99 | solusi_04.xlsx:Data!I99 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 100 | solusi_04.xlsx:Data!I100 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 101 | solusi_04.xlsx:Data!I101 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 102 | solusi_04.xlsx:Data!I102 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 103 | solusi_04.xlsx:Data!I103 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 104 | solusi_04.xlsx:Data!I104 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 105 | solusi_04.xlsx:Data!I105 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 106 | solusi_04.xlsx:Data!I106 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 107 | solusi_04.xlsx:Data!I107 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 108 | solusi_04.xlsx:Data!I108 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 109 | solusi_04.xlsx:Data!I109 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 110 | solusi_04.xlsx:Data!I110 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 111 | solusi_04.xlsx:Data!I111 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 112 | solusi_04.xlsx:Data!I112 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 113 | solusi_04.xlsx:Data!I113 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 114 | solusi_04.xlsx:Data!I114 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 115 | solusi_04.xlsx:Data!I115 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 116 | solusi_04.xlsx:Data!I116 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 117 | solusi_04.xlsx:Data!I117 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 118 | solusi_04.xlsx:Data!I118 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 119 | solusi_04.xlsx:Data!I119 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 120 | solusi_04.xlsx:Data!I120 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 121 | solusi_04.xlsx:Data!I121 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 122 | solusi_04.xlsx:Data!I122 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 123 | solusi_04.xlsx:Data!I123 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 124 | solusi_04.xlsx:Data!I124 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 125 | solusi_04.xlsx:Data!I125 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 126 | solusi_04.xlsx:Data!I126 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 127 | solusi_04.xlsx:Data!I127 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 128 | solusi_04.xlsx:Data!I128 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 129 | solusi_04.xlsx:Data!I129 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 130 | solusi_04.xlsx:Data!I130 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 131 | solusi_04.xlsx:Data!I131 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 132 | solusi_04.xlsx:Data!I132 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 133 | solusi_04.xlsx:Data!I133 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 134 | solusi_04.xlsx:Data!I134 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 135 | solusi_04.xlsx:Data!I135 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 136 | solusi_04.xlsx:Data!I136 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 137 | solusi_04.xlsx:Data!I137 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 138 | solusi_04.xlsx:Data!I138 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 139 | solusi_04.xlsx:Data!I139 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 140 | solusi_04.xlsx:Data!I140 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 141 | solusi_04.xlsx:Data!I141 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 142 | solusi_04.xlsx:Data!I142 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 143 | solusi_04.xlsx:Data!I143 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 144 | solusi_04.xlsx:Data!I144 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 145 | solusi_04.xlsx:Data!I145 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 146 | solusi_04.xlsx:Data!I146 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 147 | solusi_04.xlsx:Data!I147 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 148 | solusi_04.xlsx:Data!I148 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 149 | solusi_04.xlsx:Data!I149 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 150 | solusi_04.xlsx:Data!I150 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 151 | solusi_04.xlsx:Data!I151 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 152 | solusi_04.xlsx:Data!I152 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 153 | solusi_04.xlsx:Data!I153 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 154 | solusi_04.xlsx:Data!I154 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 155 | solusi_04.xlsx:Data!I155 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 156 | solusi_04.xlsx:Data!I156 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 157 | solusi_04.xlsx:Data!I157 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 158 | solusi_04.xlsx:Data!I158 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 159 | solusi_04.xlsx:Data!I159 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 160 | solusi_04.xlsx:Data!I160 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 161 | solusi_04.xlsx:Data!I161 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 162 | solusi_04.xlsx:Data!I162 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 163 | solusi_04.xlsx:Data!I163 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 164 | solusi_04.xlsx:Data!I164 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 165 | solusi_04.xlsx:Data!I165 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 166 | solusi_04.xlsx:Data!I166 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 167 | solusi_04.xlsx:Data!I167 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 168 | solusi_04.xlsx:Data!I168 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 169 | solusi_04.xlsx:Data!I169 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 170 | solusi_04.xlsx:Data!I170 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 171 | solusi_04.xlsx:Data!I171 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 172 | solusi_04.xlsx:Data!I172 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 173 | solusi_04.xlsx:Data!I173 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 174 | solusi_04.xlsx:Data!I174 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 175 | solusi_04.xlsx:Data!I175 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 176 | solusi_04.xlsx:Data!I176 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 177 | solusi_04.xlsx:Data!I177 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 178 | solusi_04.xlsx:Data!I178 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 179 | solusi_04.xlsx:Data!I179 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 180 | solusi_04.xlsx:Data!I180 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 181 | solusi_04.xlsx:Data!I181 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 182 | solusi_04.xlsx:Data!I182 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 183 | solusi_04.xlsx:Data!I183 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 184 | solusi_04.xlsx:Data!I184 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 185 | solusi_04.xlsx:Data!I185 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 186 | solusi_04.xlsx:Data!I186 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 187 | solusi_04.xlsx:Data!I187 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 188 | solusi_04.xlsx:Data!I188 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 189 | solusi_04.xlsx:Data!I189 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 190 | solusi_04.xlsx:Data!I190 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 191 | solusi_04.xlsx:Data!I191 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 192 | solusi_04.xlsx:Data!I192 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 193 | solusi_04.xlsx:Data!I193 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 194 | solusi_04.xlsx:Data!I194 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 195 | solusi_04.xlsx:Data!I195 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 196 | solusi_04.xlsx:Data!I196 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 197 | solusi_04.xlsx:Data!I197 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 198 | solusi_04.xlsx:Data!I198 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 199 | solusi_04.xlsx:Data!I199 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 200 | solusi_04.xlsx:Data!I200 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 201 | solusi_04.xlsx:Data!I201 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 202 | solusi_04.xlsx:Data!I202 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 203 | solusi_04.xlsx:Data!I203 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 204 | solusi_04.xlsx:Data!I204 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 205 | solusi_04.xlsx:Data!I205 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 206 | solusi_04.xlsx:Data!I206 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 207 | solusi_04.xlsx:Data!I207 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 208 | solusi_04.xlsx:Data!I208 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 209 | solusi_04.xlsx:Data!I209 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 210 | solusi_04.xlsx:Data!I210 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 211 | solusi_04.xlsx:Data!I211 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 212 | solusi_04.xlsx:Data!I212 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 213 | solusi_04.xlsx:Data!I213 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 214 | solusi_04.xlsx:Data!I214 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 215 | solusi_04.xlsx:Data!I215 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 216 | solusi_04.xlsx:Data!I216 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 217 | solusi_04.xlsx:Data!I217 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 218 | solusi_04.xlsx:Data!I218 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 219 | solusi_04.xlsx:Data!I219 | TINJAU | TINJAU | 0 | PASS |
| Aturan baris 220 | solusi_04.xlsx:Data!I220 | TEPAT | TEPAT | 0 | PASS |
| Aturan baris 221 | solusi_04.xlsx:Data!I221 | TINJAU | TINJAU | 0 | PASS |

Ringkasan: 2 file, total error = 0, rekonsiliasi gagal = 0.

## Catatan

- Sel kuning pada latihan sengaja kosong. Semua perhitungan pada solusi memakai rumus hidup.
- Pembanding dibaca dari CSV dan dihitung dengan pandas/scipy; statistik dibandingkan sampai enam desimal, selisih harus nol.
- Preview berasal dari render nyata sheet Latihan ke PDF dengan LibreOffice, kemudian PNG melalui pdftoppm.
- Fungsi diperiksa terhadap daftar yang diperbolehkan; tidak ada rumus array dinamis atau tautan workbook eksternal.
- Interaksi manual di Excel dan Google Sheets belum diuji langsung; fitur yang memerlukan langkah manual dijelaskan jujur di README.
