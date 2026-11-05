# Latihan 01: Database Karyawan

Buka `01_database_karyawan.xlsx` dan kerjakan di sheet baru bernama `Latihan`. Rumus memakai pemisah titik koma (;).

## Soal

1. Hitung jumlah karyawan tetap di departemen Produksi.
2. Hitung rata-rata usia karyawan perempuan.
3. Tampilkan nama karyawan dengan gaji tertinggi.
4. Buat kolom Kategori Masa Kerja: < 2 thn = Baru, 2-5 = Madya, > 5 = Senior.
5. Hitung berapa karyawan yang masuk pada tahun 2023.

## Kunci jawaban

1. `=COUNTIFS('Data Karyawan'!G2:G81;"Produksi";'Data Karyawan'!I2:I81;"Tetap")`
2. `=AVERAGEIF('Data Karyawan'!D2:D81;"P";'Data Karyawan'!F2:F81)`
3. `=INDEX('Data Karyawan'!C2:C81;MATCH(MAX('Data Karyawan'!M2:M81);'Data Karyawan'!M2:M81;0))`
4. `=IF(K2<2;"Baru";IF(K2<=5;"Madya";"Senior"))`
5. `=COUNTIFS('Data Karyawan'!J2:J81;">="&DATE(2023;1;1);'Data Karyawan'!J2:J81;"<="&DATE(2023;12;31))`
