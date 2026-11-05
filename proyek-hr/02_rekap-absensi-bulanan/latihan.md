# Latihan 02: Rekap Absensi Bulanan

Buka `02_rekap_absensi_bulanan.xlsx` dan kerjakan di sheet baru bernama `Latihan`. Rumus memakai pemisah titik koma (;).

## Soal

1. Hitung total hari sakit seluruh perusahaan.
2. Berapa karyawan dengan kehadiran 100%?
3. Hitung persentase alpa departemen Produksi.
4. Tandai baris karyawan yang alpa pada hari Senin pertama (kolom F) dengan Conditional Formatting.
5. Tampilkan nama karyawan dengan hari sakit terbanyak.

## Kunci jawaban

1. `=SUM(Absensi!X5:X64)`
2. `=COUNTIF(Absensi!AC5:AC64;1)`
3. `=SUMIF(Absensi!C5:C64;"Produksi";Absensi!Z5:Z64)/(COUNTIF(Absensi!C5:C64;"Produksi")*Absensi!AB5)`
4. Aturan rumus pada A5:AE64: `=$F5="A"`
5. `=INDEX(Absensi!B5:B64;MATCH(MAX(Absensi!X5:X64);Absensi!X5:X64;0))`
