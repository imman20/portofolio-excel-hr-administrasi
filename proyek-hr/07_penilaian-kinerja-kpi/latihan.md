# Latihan 07: Penilaian Kinerja KPI

Buka `07_penilaian_kinerja_kpi.xlsx` dan kerjakan di sheet baru bernama `Latihan`. Rumus memakai pemisah titik koma (;).

## Soal

1. Berapa karyawan berkuadran Bintang (Star)?
2. Rata-rata nilai akhir level Supervisor.
3. Ubah bobot KPI 1 Staff menjadi 40% dan KPI 5 menjadi 0%; cek kolom Cek lalu lihat perubahan peringkat.
4. Buat rating dengan VLOOKUP approximate memakai tabel batas di Master KPI!D12:E15.
5. Persentase karyawan berperforma Tinggi di departemen IT.

## Kunci jawaban

1. `=COUNTIF(Penilaian!V5:V84;"Bintang (Star)")`
2. `=AVERAGEIF(Penilaian!E5:E84;"Supervisor";Penilaian!P5:P84)`
3. Master KPI!C5 = 0,4 dan K5 = 0; L5 tetap 100% sehingga Cek = OK.
4. `=VLOOKUP(P5;'Master KPI'!$D$12:$E$15;2;TRUE)`
5. `=COUNTIFS(Penilaian!C5:C84;"IT";Penilaian!T5:T84;"Tinggi")/COUNTIF(Penilaian!C5:C84;"IT")`
