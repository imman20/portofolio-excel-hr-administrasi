# Latihan 06: Rekrutmen Tracker

Buka `06_rekrutmen_tracker.xlsx` dan kerjakan di sheet baru bernama `Latihan`. Rumus memakai pemisah titik koma (;).

## Soal

1. Jumlah kandidat dari LinkedIn yang mencapai Interview User.
2. Time-to-hire tercepat untuk posisi Programmer.
3. Persentase kandidat yang ditolak di tahap Screening (punya tanggal screening tetapi tidak punya Interview HR dan status Rejected).
4. Rekruter dengan kandidat hired terbanyak.
5. Jumlah kandidat yang melamar pada kuartal 2 2024.

## Kunci jawaban

1. `=COUNTIFS(Kandidat!E5:E154;"LinkedIn";Kandidat!I5:I154;"<>")`
2. `=MINIFS(Kandidat!O5:O154;Kandidat!C5:C154;"Programmer")` (Excel 2019+) atau `=MIN(IF(Kandidat!C5:C154="Programmer";Kandidat!O5:O154))` sebagai rumus array.
3. `=COUNTIFS(Kandidat!G5:G154;"<>";Kandidat!H5:H154;"";Kandidat!N5:N154;"Rejected")/Funnel!B6`
4. `=INDEX('Ringkasan Rumus'!A34:A36;MATCH(MAX('Ringkasan Rumus'!C34:C36);'Ringkasan Rumus'!C34:C36;0))`
5. `=COUNTIFS(Kandidat!F5:F154;">="&DATE(2024;4;1);Kandidat!F5:F154;"<"&DATE(2024;7;1))`
