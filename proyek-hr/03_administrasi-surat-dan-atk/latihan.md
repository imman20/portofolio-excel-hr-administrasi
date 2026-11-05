# Latihan 03: Administrasi Surat dan ATK

Buka `03_administrasi_surat_dan_atk.xlsx` dan kerjakan di sheet baru bernama `Latihan`. Rumus memakai pemisah titik koma (;).

## Soal

1. Hitung jumlah surat masuk jenis Undangan yang belum selesai.
2. Total nilai pembelian ATK (jenis Masuk) bulan Februari 2024.
3. Nama barang dengan nilai persediaan terbesar.
4. Buat nomor surat keluar dengan kode departemen dari dropdown di kolom K (misal KEU).
5. Jumlah barang yang statusnya PESAN.

## Kunci jawaban

1. `=COUNTIFS('Surat Masuk'!F5:F64;"Undangan";'Surat Masuk'!H5:H64;"<>Selesai")`
2. `=SUMIFS('Transaksi ATK'!H5:H500;'Transaksi ATK'!D5:D500;"Masuk";'Transaksi ATK'!A5:A500;">="&DATE(2024;2;1);'Transaksi ATK'!A5:A500;"<"&DATE(2024;3;1))`
3. `=INDEX('Kartu Stok'!B5:B19;MATCH(MAX('Kartu Stok'!K5:K19);'Kartu Stok'!K5:K19;0))`
4. `=TEXT(A5;"000")&"/"&K5&"/"&CHOOSE(MONTH(B5);"I";...;"XII")&"/"&YEAR(B5)`
5. `=COUNTIF('Kartu Stok'!I5:I19;"PESAN")`
