# Latihan 04: Penggajian dan Slip Gaji

Buka `04_penggajian_dan_slip_gaji.xlsx` dan kerjakan di sheet baru bernama `Latihan`. Rumus memakai pemisah titik koma (;).

## Soal

1. Total PPh 21 departemen IT.
2. Berapa karyawan yang gaji bersihnya di atas Rp10 juta?
3. Ubah tunjangan makan menjadi Rp30.000 dan amati sel mana saja yang berubah.
4. Tulis ulang rumus tunjangan jabatan dengan INDEX/MATCH.
5. Hitung rata-rata jam lembur karyawan Produksi.

## Kunci jawaban

1. `=SUMIFS('Perhitungan Gaji'!P5:P74;'Perhitungan Gaji'!C5:C74;"IT")`
2. `=COUNTIF('Perhitungan Gaji'!R5:R74;">10000000")`
3. Hanya Parameter!B7 yang diubah; kolom G, J, P, Q, R serta Rekap ikut berubah.
4. `=INDEX(Parameter!$B$20:$B$23;MATCH(E5;Parameter!$A$20:$A$23;0))`
5. `=AVERAGEIFS('Kehadiran & Lembur'!E5:E74;'Perhitungan Gaji'!C5:C74;"Produksi")`
