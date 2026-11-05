# Latihan 05: Pengelolaan Cuti

Buka `05_pengelolaan_cuti.xlsx` dan kerjakan di sheet baru bernama `Latihan`. Rumus memakai pemisah titik koma (;).

## Soal

1. Berapa hari kerja antara 1 April dan 30 April 2024 (dengan libur Idulfitri)?
2. Hitung jumlah pengajuan cuti sakit yang disetujui.
3. Nama karyawan dengan sisa saldo terbanyak.
4. Tandai pengajuan yang jatuh pada bulan Desember.
5. Berapa karyawan yang belum berhak cuti?

## Kunci jawaban

1. `=NETWORKDAYS.INTL(DATE(2024;4;1);DATE(2024;4;30);1;'Hari Libur 2024'!A5:A31)`
2. `=COUNTIFS('Pengajuan Cuti'!E5:E500;"Sakit";'Pengajuan Cuti'!J5:J500;"Disetujui")`
3. `=INDEX('Saldo Cuti'!B5:B64;MATCH(MAX('Saldo Cuti'!I5:I64);'Saldo Cuti'!I5:I64;0))`
4. Aturan CF `=MONTH($F5)=12`
5. `=COUNTIF('Saldo Cuti'!J5:J64;"Belum berhak")`
