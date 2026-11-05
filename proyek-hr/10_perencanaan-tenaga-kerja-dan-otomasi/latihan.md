# Latihan 10: Perencanaan Tenaga Kerja dan Otomasi

Buka `10_perencanaan_tenaga_kerja.xlsx` dan kerjakan di sheet baru bernama `Latihan`. Rumus memakai pemisah titik koma (;).

## Soal

1. Proyeksi headcount IT tahun 2028 dengan FORECAST.LINEAR.
2. Berapa total anggaran 2026 pada skenario Optimis?
3. Kenaikan gaji maksimal jika batas anggaran Rp52 miliar dan pertumbuhan headcount 4%.
4. Hitung kebutuhan rekrutmen total 2025 jika turnover 15%.
5. Modifikasi makro ValidasiData agar juga menandai tanggal masuk di masa depan.

## Kunci jawaban

1. `=FORECAST.LINEAR(2028;Historis!B9:G9;Historis!B13:G13)`
2. Ubah C6 ke Optimis; baca G25 (atau `=E25*((1+9%)*(1+8%))^2`).
3. `=52000000000/E25/(1+4%)-1`
4. Ubah Proyeksi!B13 = 15%; baca E22.
5. Tambahkan di loop: `If ws.Cells(r, "F").Value > Date Then Call Catat(wsHasil, n, r, ws, "Tanggal masuk di masa depan", ws.Cells(r, "F").Value)`
