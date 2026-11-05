# Latihan 08: HR Dashboard Analitik

Buka `08_hr_dashboard_analitik.xlsx` dan kerjakan di sheet baru bernama `Latihan`. Rumus memakai pemisah titik koma (;).

## Soal

1. Headcount aktif level Manager.
2. Turnover rate departemen Produksi 2024.
3. Rata-rata gaji karyawan aktif usia 25-34.
4. Bulan dengan karyawan keluar terbanyak (2023-2024).
5. Dengan Excel 365, daftar nama karyawan HR aktif terurut.

## Kunci jawaban

1. `=COUNTIFS('Data Karyawan'!H5:H254;"Manager";'Data Karyawan'!L5:L254;"Aktif")`
2. `=Perhitungan!L10` atau `=K10/((B10+K10+B10)/2)`
3. `=AVERAGEIFS('Data Karyawan'!O5:O254;'Data Karyawan'!Q5:Q254;"25-34";'Data Karyawan'!L5:L254;"Aktif")`
4. `=INDEX(Perhitungan!A23:A46;MATCH(MAX(Perhitungan!F23:F46);Perhitungan!F23:F46;0))`
5. `=SORT(FILTER(tblKaryawan[Nama];(tblKaryawan[Departemen]="HR")*(tblKaryawan[Status Aktif]="Aktif")))`
