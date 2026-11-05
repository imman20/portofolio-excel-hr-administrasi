# Verifikasi modul `06_cleaning-dan-validasi`

Tanggal verifikasi: 2026-11-05 (Asia/Jakarta). Mesin: LibreOffice headless (recalc) + openpyxl + pandas/scipy.
Skrip: `python tools/course_runtime.py 06_cleaning-dan-validasi` dari folder `belajar/`.

## Hasil recalc per file

| File | Sheet | Jumlah rumus | Jumlah error | Status recalc | Ukuran (KB) |
|---|---|---:|---:|---|---:|
| `latihan_06.xlsx` | Instruksi, Data, Latihan, Input_Validasi | 1005 | 0 | success | 30.3 |
| `solusi_06.xlsx` | Instruksi, Data, Latihan, Input_Validasi | 1015 | 0 | success | 30.5 |

## Rekonsiliasi pandas vs sel hasil recalc

| Pemeriksaan | Lokasi sel | Nilai pandas | Nilai sel | Selisih | Status |
|---|---|---:|---:|---:|---|
| Pemeriksaan 4 | solusi_06.xlsx:Latihan!C4 | 119000.0 | 119000 | 0.0 | PASS |
| Pemeriksaan 5 | solusi_06.xlsx:Latihan!C5 | 171000.0 | 171000 | 0.0 | PASS |
| Pemeriksaan 6 | solusi_06.xlsx:Latihan!C6 | 52000.0 | 52000 | 0.0 | PASS |
| Pemeriksaan 7 | solusi_06.xlsx:Latihan!C7 | 41000.0 | 41000 | 0.0 | PASS |
| Pemeriksaan 8 | solusi_06.xlsx:Latihan!C8 | 249000.0 | 249000 | 0.0 | PASS |
| Pemeriksaan 9 | solusi_06.xlsx:Latihan!C9 | 15.0 | 15 | 0.0 | PASS |
| Pemeriksaan 10 | solusi_06.xlsx:Latihan!C10 | 13498000.0 | 13498000 | 0.0 | PASS |
| Pemeriksaan 11 | solusi_06.xlsx:Latihan!C11 | 85.0 | 85 | 0.0 | PASS |
| Pemeriksaan 12 | solusi_06.xlsx:Latihan!C12 | 4.0 | 4 | 0.0 | PASS |
| Pemeriksaan 13 | solusi_06.xlsx:Latihan!C13 | 1.0 | 1 | 0.0 | PASS |
| Pemeriksaan 14 | solusi_06.xlsx:Latihan!C14 | 100.0 | 100 | 0.0 | PASS |
| Pemeriksaan 15 | solusi_06.xlsx:Latihan!C15 | 158800.0 | 158800 | 0.0 | PASS |
| Pemeriksaan 16 | solusi_06.xlsx:Latihan!C16 | 98.0 | 98 | 0.0 | PASS |
| DATEVALUE ISO | solusi_06.xlsx:Latihan!C20 | 2024-01-15 | 2024-01-15 00:00:00 | 0 | PASS |
| Nama 2 | solusi_06.xlsx:Data!G2 | Pegawai 001 | Pegawai 001 | 0 | PASS |
| Tanggal 2 | solusi_06.xlsx:Data!I2 | 2024-01-01 | 2024-01-01 00:00:00 | 0 | PASS |
| Nilai 2 | solusi_06.xlsx:Data!H2 | KOSONG | KOSONG | 0 | PASS |
| Nama 3 | solusi_06.xlsx:Data!G3 | Pegawai 002 | Pegawai 002 | 0 | PASS |
| Tanggal 3 | solusi_06.xlsx:Data!I3 | 2024-01-02 | 2024-01-02 00:00:00 | 0 | PASS |
| Nilai 3 | solusi_06.xlsx:Data!H3 | 199000.0 | 199000 | 0.0 | PASS |
| Nama 4 | solusi_06.xlsx:Data!G4 | Pegawai 003 | Pegawai 003 | 0 | PASS |
| Tanggal 4 | solusi_06.xlsx:Data!I4 | 2024-01-03 | 2024-01-03 00:00:00 | 0 | PASS |
| Nilai 4 | solusi_06.xlsx:Data!H4 | 143000.0 | 143000 | 0.0 | PASS |
| Nama 5 | solusi_06.xlsx:Data!G5 | Pegawai 004 | Pegawai 004 | 0 | PASS |
| Tanggal 5 | solusi_06.xlsx:Data!I5 | 2024-01-04 | 2024-01-04 00:00:00 | 0 | PASS |
| Nilai 5 | solusi_06.xlsx:Data!H5 | 104000.0 | 104000 | 0.0 | PASS |
| Nama 6 | solusi_06.xlsx:Data!G6 | Pegawai 005 | Pegawai 005 | 0 | PASS |
| Tanggal 6 | solusi_06.xlsx:Data!I6 | 2024-01-05 | 2024-01-05 00:00:00 | 0 | PASS |
| Nilai 6 | solusi_06.xlsx:Data!H6 | 121000.0 | 121000 | 0.0 | PASS |
| Nama 7 | solusi_06.xlsx:Data!G7 | Pegawai 006 | Pegawai 006 | 0 | PASS |
| Tanggal 7 | solusi_06.xlsx:Data!I7 | 2024-01-06 | 2024-01-06 00:00:00 | 0 | PASS |
| Nilai 7 | solusi_06.xlsx:Data!H7 | 1500000.0 | 1500000 | 0.0 | PASS |
| Nama 8 | solusi_06.xlsx:Data!G8 | Pegawai 007 | Pegawai 007 | 0 | PASS |
| Tanggal 8 | solusi_06.xlsx:Data!I8 | 2024-01-07 | 2024-01-07 00:00:00 | 0 | PASS |
| Nilai 8 | solusi_06.xlsx:Data!H8 | 180000.0 | 180000 | 0.0 | PASS |
| Nama 9 | solusi_06.xlsx:Data!G9 | Pegawai 008 | Pegawai 008 | 0 | PASS |
| Tanggal 9 | solusi_06.xlsx:Data!I9 | 2024-01-08 | 2024-01-08 00:00:00 | 0 | PASS |
| Nilai 9 | solusi_06.xlsx:Data!H9 | 146000.0 | 146000 | 0.0 | PASS |
| Nama 10 | solusi_06.xlsx:Data!G10 | Pegawai 009 | Pegawai 009 | 0 | PASS |
| Tanggal 10 | solusi_06.xlsx:Data!I10 | 2024-01-09 | 2024-01-09 00:00:00 | 0 | PASS |
| Nilai 10 | solusi_06.xlsx:Data!H10 | 120000.0 | 120000 | 0.0 | PASS |
| Nama 11 | solusi_06.xlsx:Data!G11 | Pegawai 010 | Pegawai 010 | 0 | PASS |
| Tanggal 11 | solusi_06.xlsx:Data!I11 | 2024-01-10 | 2024-01-10 00:00:00 | 0 | PASS |
| Nilai 11 | solusi_06.xlsx:Data!H11 | 197000.0 | 197000 | 0.0 | PASS |
| Nama 12 | solusi_06.xlsx:Data!G12 | Pegawai 011 | Pegawai 011 | 0 | PASS |
| Tanggal 12 | solusi_06.xlsx:Data!I12 | 2024-01-11 | 2024-01-11 00:00:00 | 0 | PASS |
| Nilai 12 | solusi_06.xlsx:Data!H12 | 124000.0 | 124000 | 0.0 | PASS |
| Nama 13 | solusi_06.xlsx:Data!G13 | Pegawai 012 | Pegawai 012 | 0 | PASS |
| Tanggal 13 | solusi_06.xlsx:Data!I13 | 2024-01-12 | 2024-01-12 00:00:00 | 0 | PASS |
| Nilai 13 | solusi_06.xlsx:Data!H13 | 114000.0 | 114000 | 0.0 | PASS |
| Nama 14 | solusi_06.xlsx:Data!G14 | Pegawai 013 | Pegawai 013 | 0 | PASS |
| Tanggal 14 | solusi_06.xlsx:Data!I14 | 2024-01-13 | 2024-01-13 00:00:00 | 0 | PASS |
| Nilai 14 | solusi_06.xlsx:Data!H14 | 109000.0 | 109000 | 0.0 | PASS |
| Nama 15 | solusi_06.xlsx:Data!G15 | Pegawai 014 | Pegawai 014 | 0 | PASS |
| Tanggal 15 | solusi_06.xlsx:Data!I15 | 2024-01-14 | 2024-01-14 00:00:00 | 0 | PASS |
| Nilai 15 | solusi_06.xlsx:Data!H15 | 132000.0 | 132000 | 0.0 | PASS |
| Nama 16 | solusi_06.xlsx:Data!G16 | Pegawai 015 | Pegawai 015 | 0 | PASS |
| Tanggal 16 | solusi_06.xlsx:Data!I16 | 2024-01-15 | 2024-01-15 00:00:00 | 0 | PASS |
| Nilai 16 | solusi_06.xlsx:Data!H16 | 126000.0 | 126000 | 0.0 | PASS |
| Nama 17 | solusi_06.xlsx:Data!G17 | Pegawai 016 | Pegawai 016 | 0 | PASS |
| Tanggal 17 | solusi_06.xlsx:Data!I17 | 2024-01-16 | 2024-01-16 00:00:00 | 0 | PASS |
| Nilai 17 | solusi_06.xlsx:Data!H17 | 119000.0 | 119000 | 0.0 | PASS |
| Nama 18 | solusi_06.xlsx:Data!G18 | Pegawai 017 | Pegawai 017 | 0 | PASS |
| Tanggal 18 | solusi_06.xlsx:Data!I18 | 2024-01-17 | 2024-01-17 00:00:00 | 0 | PASS |
| Nilai 18 | solusi_06.xlsx:Data!H18 | 171000.0 | 171000 | 0.0 | PASS |
| Nama 19 | solusi_06.xlsx:Data!G19 | Pegawai 018 | Pegawai 018 | 0 | PASS |
| Tanggal 19 | solusi_06.xlsx:Data!I19 | 2024-01-18 | 2024-01-18 00:00:00 | 0 | PASS |
| Nilai 19 | solusi_06.xlsx:Data!H19 | KOSONG | KOSONG | 0 | PASS |
| Nama 20 | solusi_06.xlsx:Data!G20 | Pegawai 019 | Pegawai 019 | 0 | PASS |
| Tanggal 20 | solusi_06.xlsx:Data!I20 | 2024-01-19 | 2024-01-19 00:00:00 | 0 | PASS |
| Nilai 20 | solusi_06.xlsx:Data!H20 | 174000.0 | 174000 | 0.0 | PASS |
| Nama 21 | solusi_06.xlsx:Data!G21 | Pegawai 020 | Pegawai 020 | 0 | PASS |
| Tanggal 21 | solusi_06.xlsx:Data!I21 | 2024-01-20 | 2024-01-20 00:00:00 | 0 | PASS |
| Nilai 21 | solusi_06.xlsx:Data!H21 | KOSONG | KOSONG | 0 | PASS |
| Nama 22 | solusi_06.xlsx:Data!G22 | Pegawai 021 | Pegawai 021 | 0 | PASS |
| Tanggal 22 | solusi_06.xlsx:Data!I22 | 2024-01-21 | 2024-01-21 00:00:00 | 0 | PASS |
| Nilai 22 | solusi_06.xlsx:Data!H22 | 197000.0 | 197000 | 0.0 | PASS |
| Nama 23 | solusi_06.xlsx:Data!G23 | Pegawai 022 | Pegawai 022 | 0 | PASS |
| Tanggal 23 | solusi_06.xlsx:Data!I23 | 2024-01-22 | 2024-01-22 00:00:00 | 0 | PASS |
| Nilai 23 | solusi_06.xlsx:Data!H23 | 164000.0 | 164000 | 0.0 | PASS |
| Nama 24 | solusi_06.xlsx:Data!G24 | Pegawai 023 | Pegawai 023 | 0 | PASS |
| Tanggal 24 | solusi_06.xlsx:Data!I24 | 2024-01-23 | 2024-01-23 00:00:00 | 0 | PASS |
| Nilai 24 | solusi_06.xlsx:Data!H24 | 124000.0 | 124000 | 0.0 | PASS |
| Nama 25 | solusi_06.xlsx:Data!G25 | Pegawai 024 | Pegawai 024 | 0 | PASS |
| Tanggal 25 | solusi_06.xlsx:Data!I25 | 2024-01-24 | 2024-01-24 00:00:00 | 0 | PASS |
| Nilai 25 | solusi_06.xlsx:Data!H25 | KOSONG | KOSONG | 0 | PASS |
| Nama 26 | solusi_06.xlsx:Data!G26 | Pegawai 025 | Pegawai 025 | 0 | PASS |
| Tanggal 26 | solusi_06.xlsx:Data!I26 | 2024-01-25 | 2024-01-25 00:00:00 | 0 | PASS |
| Nilai 26 | solusi_06.xlsx:Data!H26 | 120000.0 | 120000 | 0.0 | PASS |
| Nama 27 | solusi_06.xlsx:Data!G27 | Pegawai 026 | Pegawai 026 | 0 | PASS |
| Tanggal 27 | solusi_06.xlsx:Data!I27 | 2024-01-26 | 2024-01-26 00:00:00 | 0 | PASS |
| Nilai 27 | solusi_06.xlsx:Data!H27 | 198000.0 | 198000 | 0.0 | PASS |
| Nama 28 | solusi_06.xlsx:Data!G28 | Pegawai 027 | Pegawai 027 | 0 | PASS |
| Tanggal 28 | solusi_06.xlsx:Data!I28 | 2024-01-27 | 2024-01-27 00:00:00 | 0 | PASS |
| Nilai 28 | solusi_06.xlsx:Data!H28 | 122000.0 | 122000 | 0.0 | PASS |
| Nama 29 | solusi_06.xlsx:Data!G29 | Pegawai 028 | Pegawai 028 | 0 | PASS |
| Tanggal 29 | solusi_06.xlsx:Data!I29 | 2024-01-28 | 2024-01-28 00:00:00 | 0 | PASS |
| Nilai 29 | solusi_06.xlsx:Data!H29 | 115000.0 | 115000 | 0.0 | PASS |
| Nama 30 | solusi_06.xlsx:Data!G30 | Pegawai 029 | Pegawai 029 | 0 | PASS |
| Tanggal 30 | solusi_06.xlsx:Data!I30 | 2024-01-29 | 2024-01-29 00:00:00 | 0 | PASS |
| Nilai 30 | solusi_06.xlsx:Data!H30 | 126000.0 | 126000 | 0.0 | PASS |
| Nama 31 | solusi_06.xlsx:Data!G31 | Pegawai 030 | Pegawai 030 | 0 | PASS |
| Tanggal 31 | solusi_06.xlsx:Data!I31 | 2024-01-30 | 2024-01-30 00:00:00 | 0 | PASS |
| Nilai 31 | solusi_06.xlsx:Data!H31 | 129000.0 | 129000 | 0.0 | PASS |
| Nama 32 | solusi_06.xlsx:Data!G32 | Pegawai 031 | Pegawai 031 | 0 | PASS |
| Tanggal 32 | solusi_06.xlsx:Data!I32 | 2024-01-31 | 2024-01-31 00:00:00 | 0 | PASS |
| Nilai 32 | solusi_06.xlsx:Data!H32 | 165000.0 | 165000 | 0.0 | PASS |
| Nama 33 | solusi_06.xlsx:Data!G33 | Pegawai 032 | Pegawai 032 | 0 | PASS |
| Tanggal 33 | solusi_06.xlsx:Data!I33 | 2024-02-01 | 2024-02-01 00:00:00 | 0 | PASS |
| Nilai 33 | solusi_06.xlsx:Data!H33 | 105000.0 | 105000 | 0.0 | PASS |
| Nama 34 | solusi_06.xlsx:Data!G34 | Pegawai 033 | Pegawai 033 | 0 | PASS |
| Tanggal 34 | solusi_06.xlsx:Data!I34 | 2024-02-02 | 2024-02-02 00:00:00 | 0 | PASS |
| Nilai 34 | solusi_06.xlsx:Data!H34 | 176000.0 | 176000 | 0.0 | PASS |
| Nama 35 | solusi_06.xlsx:Data!G35 | Pegawai 034 | Pegawai 034 | 0 | PASS |
| Tanggal 35 | solusi_06.xlsx:Data!I35 | 2024-02-03 | 2024-02-03 00:00:00 | 0 | PASS |
| Nilai 35 | solusi_06.xlsx:Data!H35 | 172000.0 | 172000 | 0.0 | PASS |
| Nama 36 | solusi_06.xlsx:Data!G36 | Pegawai 035 | Pegawai 035 | 0 | PASS |
| Tanggal 36 | solusi_06.xlsx:Data!I36 | 2024-02-04 | 2024-02-04 00:00:00 | 0 | PASS |
| Nilai 36 | solusi_06.xlsx:Data!H36 | KOSONG | KOSONG | 0 | PASS |
| Nama 37 | solusi_06.xlsx:Data!G37 | Pegawai 036 | Pegawai 036 | 0 | PASS |
| Tanggal 37 | solusi_06.xlsx:Data!I37 | 2024-02-05 | 2024-02-05 00:00:00 | 0 | PASS |
| Nilai 37 | solusi_06.xlsx:Data!H37 | 140000.0 | 140000 | 0.0 | PASS |
| Nama 38 | solusi_06.xlsx:Data!G38 | Pegawai 037 | Pegawai 037 | 0 | PASS |
| Tanggal 38 | solusi_06.xlsx:Data!I38 | 2024-02-06 | 2024-02-06 00:00:00 | 0 | PASS |
| Nilai 38 | solusi_06.xlsx:Data!H38 | 109000.0 | 109000 | 0.0 | PASS |
| Nama 39 | solusi_06.xlsx:Data!G39 | Pegawai 038 | Pegawai 038 | 0 | PASS |
| Tanggal 39 | solusi_06.xlsx:Data!I39 | 2024-02-07 | 2024-02-07 00:00:00 | 0 | PASS |
| Nilai 39 | solusi_06.xlsx:Data!H39 | 152000.0 | 152000 | 0.0 | PASS |
| Nama 40 | solusi_06.xlsx:Data!G40 | Pegawai 039 | Pegawai 039 | 0 | PASS |
| Tanggal 40 | solusi_06.xlsx:Data!I40 | 2024-02-08 | 2024-02-08 00:00:00 | 0 | PASS |
| Nilai 40 | solusi_06.xlsx:Data!H40 | KOSONG | KOSONG | 0 | PASS |
| Nama 41 | solusi_06.xlsx:Data!G41 | Pegawai 040 | Pegawai 040 | 0 | PASS |
| Tanggal 41 | solusi_06.xlsx:Data!I41 | 2024-02-09 | 2024-02-09 00:00:00 | 0 | PASS |
| Nilai 41 | solusi_06.xlsx:Data!H41 | 169000.0 | 169000 | 0.0 | PASS |
| Nama 42 | solusi_06.xlsx:Data!G42 | Pegawai 041 | Pegawai 041 | 0 | PASS |
| Tanggal 42 | solusi_06.xlsx:Data!I42 | 2024-02-10 | 2024-02-10 00:00:00 | 0 | PASS |
| Nilai 42 | solusi_06.xlsx:Data!H42 | 101000.0 | 101000 | 0.0 | PASS |
| Nama 43 | solusi_06.xlsx:Data!G43 | Pegawai 042 | Pegawai 042 | 0 | PASS |
| Tanggal 43 | solusi_06.xlsx:Data!I43 | 2024-02-11 | 2024-02-11 00:00:00 | 0 | PASS |
| Nilai 43 | solusi_06.xlsx:Data!H43 | 104000.0 | 104000 | 0.0 | PASS |
| Nama 44 | solusi_06.xlsx:Data!G44 | Pegawai 043 | Pegawai 043 | 0 | PASS |
| Tanggal 44 | solusi_06.xlsx:Data!I44 | 2024-02-12 | 2024-02-12 00:00:00 | 0 | PASS |
| Nilai 44 | solusi_06.xlsx:Data!H44 | 138000.0 | 138000 | 0.0 | PASS |
| Nama 45 | solusi_06.xlsx:Data!G45 | Pegawai 044 | Pegawai 044 | 0 | PASS |
| Tanggal 45 | solusi_06.xlsx:Data!I45 | 2024-02-13 | 2024-02-13 00:00:00 | 0 | PASS |
| Nilai 45 | solusi_06.xlsx:Data!H45 | 126000.0 | 126000 | 0.0 | PASS |
| Nama 46 | solusi_06.xlsx:Data!G46 | Pegawai 045 | Pegawai 045 | 0 | PASS |
| Tanggal 46 | solusi_06.xlsx:Data!I46 | 2024-02-14 | 2024-02-14 00:00:00 | 0 | PASS |
| Nilai 46 | solusi_06.xlsx:Data!H46 | 111000.0 | 111000 | 0.0 | PASS |
| Nama 47 | solusi_06.xlsx:Data!G47 | Pegawai 046 | Pegawai 046 | 0 | PASS |
| Tanggal 47 | solusi_06.xlsx:Data!I47 | 2024-02-15 | 2024-02-15 00:00:00 | 0 | PASS |
| Nilai 47 | solusi_06.xlsx:Data!H47 | 142000.0 | 142000 | 0.0 | PASS |
| Nama 48 | solusi_06.xlsx:Data!G48 | Pegawai 047 | Pegawai 047 | 0 | PASS |
| Tanggal 48 | solusi_06.xlsx:Data!I48 | 2024-02-16 | 2024-02-16 00:00:00 | 0 | PASS |
| Nilai 48 | solusi_06.xlsx:Data!H48 | KOSONG | KOSONG | 0 | PASS |
| Nama 49 | solusi_06.xlsx:Data!G49 | Pegawai 048 | Pegawai 048 | 0 | PASS |
| Tanggal 49 | solusi_06.xlsx:Data!I49 | 2024-02-17 | 2024-02-17 00:00:00 | 0 | PASS |
| Nilai 49 | solusi_06.xlsx:Data!H49 | 109000.0 | 109000 | 0.0 | PASS |
| Nama 50 | solusi_06.xlsx:Data!G50 | Pegawai 049 | Pegawai 049 | 0 | PASS |
| Tanggal 50 | solusi_06.xlsx:Data!I50 | 2024-02-18 | 2024-02-18 00:00:00 | 0 | PASS |
| Nilai 50 | solusi_06.xlsx:Data!H50 | 170000.0 | 170000 | 0.0 | PASS |
| Nama 51 | solusi_06.xlsx:Data!G51 | Pegawai 050 | Pegawai 050 | 0 | PASS |
| Tanggal 51 | solusi_06.xlsx:Data!I51 | 2024-02-19 | 2024-02-19 00:00:00 | 0 | PASS |
| Nilai 51 | solusi_06.xlsx:Data!H51 | 108000.0 | 108000 | 0.0 | PASS |
| Nama 52 | solusi_06.xlsx:Data!G52 | Pegawai 051 | Pegawai 051 | 0 | PASS |
| Tanggal 52 | solusi_06.xlsx:Data!I52 | 2024-02-20 | 2024-02-20 00:00:00 | 0 | PASS |
| Nilai 52 | solusi_06.xlsx:Data!H52 | 188000.0 | 188000 | 0.0 | PASS |
| Nama 53 | solusi_06.xlsx:Data!G53 | Pegawai 052 | Pegawai 052 | 0 | PASS |
| Tanggal 53 | solusi_06.xlsx:Data!I53 | 2024-02-21 | 2024-02-21 00:00:00 | 0 | PASS |
| Nilai 53 | solusi_06.xlsx:Data!H53 | KOSONG | KOSONG | 0 | PASS |
| Nama 54 | solusi_06.xlsx:Data!G54 | Pegawai 053 | Pegawai 053 | 0 | PASS |
| Tanggal 54 | solusi_06.xlsx:Data!I54 | 2024-02-22 | 2024-02-22 00:00:00 | 0 | PASS |
| Nilai 54 | solusi_06.xlsx:Data!H54 | 116000.0 | 116000 | 0.0 | PASS |
| Nama 55 | solusi_06.xlsx:Data!G55 | Pegawai 054 | Pegawai 054 | 0 | PASS |
| Tanggal 55 | solusi_06.xlsx:Data!I55 | 2024-02-23 | 2024-02-23 00:00:00 | 0 | PASS |
| Nilai 55 | solusi_06.xlsx:Data!H55 | 170000.0 | 170000 | 0.0 | PASS |
| Nama 56 | solusi_06.xlsx:Data!G56 | Pegawai 055 | Pegawai 055 | 0 | PASS |
| Tanggal 56 | solusi_06.xlsx:Data!I56 | 2024-02-24 | 2024-02-24 00:00:00 | 0 | PASS |
| Nilai 56 | solusi_06.xlsx:Data!H56 | 104000.0 | 104000 | 0.0 | PASS |
| Nama 57 | solusi_06.xlsx:Data!G57 | Pegawai 056 | Pegawai 056 | 0 | PASS |
| Tanggal 57 | solusi_06.xlsx:Data!I57 | 2024-02-25 | 2024-02-25 00:00:00 | 0 | PASS |
| Nilai 57 | solusi_06.xlsx:Data!H57 | 143000.0 | 143000 | 0.0 | PASS |
| Nama 58 | solusi_06.xlsx:Data!G58 | Pegawai 057 | Pegawai 057 | 0 | PASS |
| Tanggal 58 | solusi_06.xlsx:Data!I58 | 2024-02-26 | 2024-02-26 00:00:00 | 0 | PASS |
| Nilai 58 | solusi_06.xlsx:Data!H58 | 189000.0 | 189000 | 0.0 | PASS |
| Nama 59 | solusi_06.xlsx:Data!G59 | Pegawai 058 | Pegawai 058 | 0 | PASS |
| Tanggal 59 | solusi_06.xlsx:Data!I59 | 2024-02-27 | 2024-02-27 00:00:00 | 0 | PASS |
| Nilai 59 | solusi_06.xlsx:Data!H59 | KOSONG | KOSONG | 0 | PASS |
| Nama 60 | solusi_06.xlsx:Data!G60 | Pegawai 059 | Pegawai 059 | 0 | PASS |
| Tanggal 60 | solusi_06.xlsx:Data!I60 | 2024-02-28 | 2024-02-28 00:00:00 | 0 | PASS |
| Nilai 60 | solusi_06.xlsx:Data!H60 | 106000.0 | 106000 | 0.0 | PASS |
| Nama 61 | solusi_06.xlsx:Data!G61 | Pegawai 060 | Pegawai 060 | 0 | PASS |
| Tanggal 61 | solusi_06.xlsx:Data!I61 | 2024-02-29 | 2024-02-29 00:00:00 | 0 | PASS |
| Nilai 61 | solusi_06.xlsx:Data!H61 | 114000.0 | 114000 | 0.0 | PASS |
| Nama 62 | solusi_06.xlsx:Data!G62 | Pegawai 061 | Pegawai 061 | 0 | PASS |
| Tanggal 62 | solusi_06.xlsx:Data!I62 | 2024-03-01 | 2024-03-01 00:00:00 | 0 | PASS |
| Nilai 62 | solusi_06.xlsx:Data!H62 | 200000.0 | 200000 | 0.0 | PASS |
| Nama 63 | solusi_06.xlsx:Data!G63 | Pegawai 062 | Pegawai 062 | 0 | PASS |
| Tanggal 63 | solusi_06.xlsx:Data!I63 | 2024-03-02 | 2024-03-02 00:00:00 | 0 | PASS |
| Nilai 63 | solusi_06.xlsx:Data!H63 | 139000.0 | 139000 | 0.0 | PASS |
| Nama 64 | solusi_06.xlsx:Data!G64 | Pegawai 063 | Pegawai 063 | 0 | PASS |
| Tanggal 64 | solusi_06.xlsx:Data!I64 | 2024-03-03 | 2024-03-03 00:00:00 | 0 | PASS |
| Nilai 64 | solusi_06.xlsx:Data!H64 | 115000.0 | 115000 | 0.0 | PASS |
| Nama 65 | solusi_06.xlsx:Data!G65 | Pegawai 064 | Pegawai 064 | 0 | PASS |
| Tanggal 65 | solusi_06.xlsx:Data!I65 | 2024-03-04 | 2024-03-04 00:00:00 | 0 | PASS |
| Nilai 65 | solusi_06.xlsx:Data!H65 | 199000.0 | 199000 | 0.0 | PASS |
| Nama 66 | solusi_06.xlsx:Data!G66 | Pegawai 065 | Pegawai 065 | 0 | PASS |
| Tanggal 66 | solusi_06.xlsx:Data!I66 | 2024-03-05 | 2024-03-05 00:00:00 | 0 | PASS |
| Nilai 66 | solusi_06.xlsx:Data!H66 | 133000.0 | 133000 | 0.0 | PASS |
| Nama 67 | solusi_06.xlsx:Data!G67 | Pegawai 066 | Pegawai 066 | 0 | PASS |
| Tanggal 67 | solusi_06.xlsx:Data!I67 | 2024-03-06 | 2024-03-06 00:00:00 | 0 | PASS |
| Nilai 67 | solusi_06.xlsx:Data!H67 | 119000.0 | 119000 | 0.0 | PASS |
| Nama 68 | solusi_06.xlsx:Data!G68 | Pegawai 067 | Pegawai 067 | 0 | PASS |
| Tanggal 68 | solusi_06.xlsx:Data!I68 | 2024-03-07 | 2024-03-07 00:00:00 | 0 | PASS |
| Nilai 68 | solusi_06.xlsx:Data!H68 | 130000.0 | 130000 | 0.0 | PASS |
| Nama 69 | solusi_06.xlsx:Data!G69 | Pegawai 068 | Pegawai 068 | 0 | PASS |
| Tanggal 69 | solusi_06.xlsx:Data!I69 | 2024-03-08 | 2024-03-08 00:00:00 | 0 | PASS |
| Nilai 69 | solusi_06.xlsx:Data!H69 | 132000.0 | 132000 | 0.0 | PASS |
| Nama 70 | solusi_06.xlsx:Data!G70 | Pegawai 069 | Pegawai 069 | 0 | PASS |
| Tanggal 70 | solusi_06.xlsx:Data!I70 | 2024-03-09 | 2024-03-09 00:00:00 | 0 | PASS |
| Nilai 70 | solusi_06.xlsx:Data!H70 | KOSONG | KOSONG | 0 | PASS |
| Nama 71 | solusi_06.xlsx:Data!G71 | Pegawai 070 | Pegawai 070 | 0 | PASS |
| Tanggal 71 | solusi_06.xlsx:Data!I71 | 2024-03-10 | 2024-03-10 00:00:00 | 0 | PASS |
| Nilai 71 | solusi_06.xlsx:Data!H71 | KOSONG | KOSONG | 0 | PASS |
| Nama 72 | solusi_06.xlsx:Data!G72 | Pegawai 071 | Pegawai 071 | 0 | PASS |
| Tanggal 72 | solusi_06.xlsx:Data!I72 | 2024-03-11 | 2024-03-11 00:00:00 | 0 | PASS |
| Nilai 72 | solusi_06.xlsx:Data!H72 | 139000.0 | 139000 | 0.0 | PASS |
| Nama 73 | solusi_06.xlsx:Data!G73 | Pegawai 072 | Pegawai 072 | 0 | PASS |
| Tanggal 73 | solusi_06.xlsx:Data!I73 | 2024-03-12 | 2024-03-12 00:00:00 | 0 | PASS |
| Nilai 73 | solusi_06.xlsx:Data!H73 | 189000.0 | 189000 | 0.0 | PASS |
| Nama 74 | solusi_06.xlsx:Data!G74 | Pegawai 073 | Pegawai 073 | 0 | PASS |
| Tanggal 74 | solusi_06.xlsx:Data!I74 | 2024-03-13 | 2024-03-13 00:00:00 | 0 | PASS |
| Nilai 74 | solusi_06.xlsx:Data!H74 | 110000.0 | 110000 | 0.0 | PASS |
| Nama 75 | solusi_06.xlsx:Data!G75 | Pegawai 074 | Pegawai 074 | 0 | PASS |
| Tanggal 75 | solusi_06.xlsx:Data!I75 | 2024-03-14 | 2024-03-14 00:00:00 | 0 | PASS |
| Nilai 75 | solusi_06.xlsx:Data!H75 | 189000.0 | 189000 | 0.0 | PASS |
| Nama 76 | solusi_06.xlsx:Data!G76 | Pegawai 075 | Pegawai 075 | 0 | PASS |
| Tanggal 76 | solusi_06.xlsx:Data!I76 | 2024-03-15 | 2024-03-15 00:00:00 | 0 | PASS |
| Nilai 76 | solusi_06.xlsx:Data!H76 | 179000.0 | 179000 | 0.0 | PASS |
| Nama 77 | solusi_06.xlsx:Data!G77 | Pegawai 076 | Pegawai 076 | 0 | PASS |
| Tanggal 77 | solusi_06.xlsx:Data!I77 | 2024-03-16 | 2024-03-16 00:00:00 | 0 | PASS |
| Nilai 77 | solusi_06.xlsx:Data!H77 | 125000.0 | 125000 | 0.0 | PASS |
| Nama 78 | solusi_06.xlsx:Data!G78 | Pegawai 077 | Pegawai 077 | 0 | PASS |
| Tanggal 78 | solusi_06.xlsx:Data!I78 | 2024-03-17 | 2024-03-17 00:00:00 | 0 | PASS |
| Nilai 78 | solusi_06.xlsx:Data!H78 | KOSONG | KOSONG | 0 | PASS |
| Nama 79 | solusi_06.xlsx:Data!G79 | Pegawai 078 | Pegawai 078 | 0 | PASS |
| Tanggal 79 | solusi_06.xlsx:Data!I79 | 2024-03-18 | 2024-03-18 00:00:00 | 0 | PASS |
| Nilai 79 | solusi_06.xlsx:Data!H79 | 191000.0 | 191000 | 0.0 | PASS |
| Nama 80 | solusi_06.xlsx:Data!G80 | Pegawai 079 | Pegawai 079 | 0 | PASS |
| Tanggal 80 | solusi_06.xlsx:Data!I80 | 2024-03-19 | 2024-03-19 00:00:00 | 0 | PASS |
| Nilai 80 | solusi_06.xlsx:Data!H80 | 134000.0 | 134000 | 0.0 | PASS |
| Nama 81 | solusi_06.xlsx:Data!G81 | Pegawai 080 | Pegawai 080 | 0 | PASS |
| Tanggal 81 | solusi_06.xlsx:Data!I81 | 2024-03-20 | 2024-03-20 00:00:00 | 0 | PASS |
| Nilai 81 | solusi_06.xlsx:Data!H81 | 127000.0 | 127000 | 0.0 | PASS |
| Nama 82 | solusi_06.xlsx:Data!G82 | Pegawai 081 | Pegawai 081 | 0 | PASS |
| Tanggal 82 | solusi_06.xlsx:Data!I82 | 2024-03-21 | 2024-03-21 00:00:00 | 0 | PASS |
| Nilai 82 | solusi_06.xlsx:Data!H82 | 169000.0 | 169000 | 0.0 | PASS |
| Nama 83 | solusi_06.xlsx:Data!G83 | Pegawai 082 | Pegawai 082 | 0 | PASS |
| Tanggal 83 | solusi_06.xlsx:Data!I83 | 2024-03-22 | 2024-03-22 00:00:00 | 0 | PASS |
| Nilai 83 | solusi_06.xlsx:Data!H83 | 154000.0 | 154000 | 0.0 | PASS |
| Nama 84 | solusi_06.xlsx:Data!G84 | Pegawai 083 | Pegawai 083 | 0 | PASS |
| Tanggal 84 | solusi_06.xlsx:Data!I84 | 2024-03-23 | 2024-03-23 00:00:00 | 0 | PASS |
| Nilai 84 | solusi_06.xlsx:Data!H84 | 142000.0 | 142000 | 0.0 | PASS |
| Nama 85 | solusi_06.xlsx:Data!G85 | Pegawai 084 | Pegawai 084 | 0 | PASS |
| Tanggal 85 | solusi_06.xlsx:Data!I85 | 2024-03-24 | 2024-03-24 00:00:00 | 0 | PASS |
| Nilai 85 | solusi_06.xlsx:Data!H85 | 169000.0 | 169000 | 0.0 | PASS |
| Nama 86 | solusi_06.xlsx:Data!G86 | Pegawai 085 | Pegawai 085 | 0 | PASS |
| Tanggal 86 | solusi_06.xlsx:Data!I86 | 2024-03-25 | 2024-03-25 00:00:00 | 0 | PASS |
| Nilai 86 | solusi_06.xlsx:Data!H86 | 178000.0 | 178000 | 0.0 | PASS |
| Nama 87 | solusi_06.xlsx:Data!G87 | Pegawai 086 | Pegawai 086 | 0 | PASS |
| Tanggal 87 | solusi_06.xlsx:Data!I87 | 2024-03-26 | 2024-03-26 00:00:00 | 0 | PASS |
| Nilai 87 | solusi_06.xlsx:Data!H87 | KOSONG | KOSONG | 0 | PASS |
| Nama 88 | solusi_06.xlsx:Data!G88 | Pegawai 087 | Pegawai 087 | 0 | PASS |
| Tanggal 88 | solusi_06.xlsx:Data!I88 | 2024-03-27 | 2024-03-27 00:00:00 | 0 | PASS |
| Nilai 88 | solusi_06.xlsx:Data!H88 | 141000.0 | 141000 | 0.0 | PASS |
| Nama 89 | solusi_06.xlsx:Data!G89 | Pegawai 088 | Pegawai 088 | 0 | PASS |
| Tanggal 89 | solusi_06.xlsx:Data!I89 | 2024-03-28 | 2024-03-28 00:00:00 | 0 | PASS |
| Nilai 89 | solusi_06.xlsx:Data!H89 | 115000.0 | 115000 | 0.0 | PASS |
| Nama 90 | solusi_06.xlsx:Data!G90 | Pegawai 089 | Pegawai 089 | 0 | PASS |
| Tanggal 90 | solusi_06.xlsx:Data!I90 | 2024-03-29 | 2024-03-29 00:00:00 | 0 | PASS |
| Nilai 90 | solusi_06.xlsx:Data!H90 | 128000.0 | 128000 | 0.0 | PASS |
| Nama 91 | solusi_06.xlsx:Data!G91 | Pegawai 090 | Pegawai 090 | 0 | PASS |
| Tanggal 91 | solusi_06.xlsx:Data!I91 | 2024-03-30 | 2024-03-30 00:00:00 | 0 | PASS |
| Nilai 91 | solusi_06.xlsx:Data!H91 | 187000.0 | 187000 | 0.0 | PASS |
| Nama 92 | solusi_06.xlsx:Data!G92 | Pegawai 091 | Pegawai 091 | 0 | PASS |
| Tanggal 92 | solusi_06.xlsx:Data!I92 | 2024-03-31 | 2024-03-31 00:00:00 | 0 | PASS |
| Nilai 92 | solusi_06.xlsx:Data!H92 | 131000.0 | 131000 | 0.0 | PASS |
| Nama 93 | solusi_06.xlsx:Data!G93 | Pegawai 092 | Pegawai 092 | 0 | PASS |
| Tanggal 93 | solusi_06.xlsx:Data!I93 | 2024-04-01 | 2024-04-01 00:00:00 | 0 | PASS |
| Nilai 93 | solusi_06.xlsx:Data!H93 | 110000.0 | 110000 | 0.0 | PASS |
| Nama 94 | solusi_06.xlsx:Data!G94 | Pegawai 093 | Pegawai 093 | 0 | PASS |
| Tanggal 94 | solusi_06.xlsx:Data!I94 | 2024-04-02 | 2024-04-02 00:00:00 | 0 | PASS |
| Nilai 94 | solusi_06.xlsx:Data!H94 | KOSONG | KOSONG | 0 | PASS |
| Nama 95 | solusi_06.xlsx:Data!G95 | Pegawai 094 | Pegawai 094 | 0 | PASS |
| Tanggal 95 | solusi_06.xlsx:Data!I95 | 2024-04-03 | 2024-04-03 00:00:00 | 0 | PASS |
| Nilai 95 | solusi_06.xlsx:Data!H95 | 171000.0 | 171000 | 0.0 | PASS |
| Nama 96 | solusi_06.xlsx:Data!G96 | Pegawai 095 | Pegawai 095 | 0 | PASS |
| Tanggal 96 | solusi_06.xlsx:Data!I96 | 2024-04-04 | 2024-04-04 00:00:00 | 0 | PASS |
| Nilai 96 | solusi_06.xlsx:Data!H96 | 134000.0 | 134000 | 0.0 | PASS |
| Nama 97 | solusi_06.xlsx:Data!G97 | Pegawai 096 | Pegawai 096 | 0 | PASS |
| Tanggal 97 | solusi_06.xlsx:Data!I97 | 2024-04-05 | 2024-04-05 00:00:00 | 0 | PASS |
| Nilai 97 | solusi_06.xlsx:Data!H97 | KOSONG | KOSONG | 0 | PASS |
| Nama 98 | solusi_06.xlsx:Data!G98 | Pegawai 097 | Pegawai 097 | 0 | PASS |
| Tanggal 98 | solusi_06.xlsx:Data!I98 | 2024-04-06 | 2024-04-06 00:00:00 | 0 | PASS |
| Nilai 98 | solusi_06.xlsx:Data!H98 | 101000.0 | 101000 | 0.0 | PASS |
| Nama 99 | solusi_06.xlsx:Data!G99 | Pegawai 098 | Pegawai 098 | 0 | PASS |
| Tanggal 99 | solusi_06.xlsx:Data!I99 | 2024-04-07 | 2024-04-07 00:00:00 | 0 | PASS |
| Nilai 99 | solusi_06.xlsx:Data!H99 | 102000.0 | 102000 | 0.0 | PASS |
| Nama 100 | solusi_06.xlsx:Data!G100 | Pegawai 099 | Pegawai 099 | 0 | PASS |
| Tanggal 100 | solusi_06.xlsx:Data!I100 | 2024-04-08 | 2024-04-08 00:00:00 | 0 | PASS |
| Nilai 100 | solusi_06.xlsx:Data!H100 | 178000.0 | 178000 | 0.0 | PASS |
| Nama 101 | solusi_06.xlsx:Data!G101 | Pegawai 100 | Pegawai 100 | 0 | PASS |
| Tanggal 101 | solusi_06.xlsx:Data!I101 | 2024-04-09 | 2024-04-09 00:00:00 | 0 | PASS |
| Nilai 101 | solusi_06.xlsx:Data!H101 | 137000.0 | 137000 | 0.0 | PASS |

Ringkasan: 2 file, total error = 0, rekonsiliasi gagal = 0.

## Catatan

- Sel kuning pada latihan sengaja kosong. Semua perhitungan pada solusi memakai rumus hidup.
- Pembanding dibaca dari CSV dan dihitung dengan pandas/scipy; statistik dibandingkan sampai enam desimal, selisih harus nol.
- Preview berasal dari render nyata sheet Latihan ke PDF dengan LibreOffice, kemudian PNG melalui pdftoppm.
- Fungsi diperiksa terhadap daftar yang diperbolehkan; tidak ada rumus array dinamis atau tautan workbook eksternal.
- Interaksi manual di Excel dan Google Sheets belum diuji langsung; fitur yang memerlukan langkah manual dijelaskan jujur di README.
