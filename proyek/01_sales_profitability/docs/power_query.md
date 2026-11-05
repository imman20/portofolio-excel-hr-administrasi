# Power Query dari CSV

Saya menyediakan [skrip M transaksi](../pq/transaksi.pq) sebagai latihan impor dan transformasi. Query belum tertanam di workbook. Rumus pada workbook tetap menjadi hasil pembanding yang bisa dibuka tanpa Power Query.

## Menjalankan di Excel Windows

1. Unduh folder proyek secara lengkap. Catat lokasi absolut `data/raw` pada komputer, bukan URL halaman GitHub.
2. Buka salinan workbook solusi, pilih **Data > Get Data > From Other Sources > Blank Query**. Di Excel 2016, menu awal dapat bernama **New Query**.
3. Di Power Query Editor pilih **Home > Advanced Editor**. Ganti seluruh isi dengan skrip M, lalu ubah `FolderData` pada baris awal menjadi lokasi `data/raw`, misalnya `C:\portofolio\proyek\01_sales_profitability\data\raw`. Gunakan satu backslash per pemisah direktori pada nilai sebenarnya; teks M tidak memakai backslash sebagai escape.
4. Beri nama query `Transaksi_PQ`. Jika diminta izin sumber, gunakan pengaturan privasi yang sesuai lokasi lokal. Jangan mengaktifkan akses ke sumber lain yang tidak diperlukan.
5. Pilih **Home > Close & Load To... > Table > New worksheet**. Jangan menimpa tabel `Transaksi`, karena tabel itu berisi rumus latihan dan sumber chart.
6. Pilih **Data > Refresh All** saat CSV berubah. Periksa langkah terakhir dan jumlah baris yang dimuat. Refresh query tidak otomatis memperluas range absolut pada Calculation; jalankan ulang builder jika ukuran data berubah.

## Apa yang dikerjakan query

Query membaca transaksi, retur, produk, pelanggan, wilayah, sales, dan target dengan encoding UTF-8. Angka dibaca memakai locale `en-US`, sesuai desimal bertitik pada CSV; tanggal dibaca dari format ISO. ID tetap teks. Saya memeriksa kunci unik sebelum join, relasi region, range diskon/qty, tanggal order dan retur, serta grain target.

Retur di-Group By order_id sebelum di-Merge Left Outer ke transaksi. Nilai retur yang tidak ada menjadi nol. Produk dan pelanggan di-Merge hanya setelah keunikan master diperiksa. HPP dan kategori di-Expand dari produk; segmen dari pelanggan. Query menambahkan KPI yang sama dengan workbook dan mengembalikan kolom dalam urutan Raw_Data!A:X. Kolom order_ditemukan pada Retur adalah kontrol workbook terpisah.

## Pemeriksaan hasil

1. Pada hasil query, pastikan jumlah baris sama dengan `Raw_Data`. Jika join meningkatkan jumlah baris, jangan menghapus baris hasil secara acak; periksa kunci master dan agregasi retur.
2. Pilih **Transform > Group By > Advanced**, group `periode`, lalu jumlahkan omzet_bruto, penjualan_bersih, hpp, dan laba_kotor. Bandingkan setiap bulan dengan `Calculation!A4:E17` dan [VERIFIKASI.md](../VERIFIKASI.md).
3. Bandingkan jumlah total, nilai retur, dan hasil per customer_id. Margin dihitung dari total laba/total penjualan, bukan dijumlahkan.
4. Jika query berhenti dengan pesan relasi atau nilai tidak valid, perbaiki CSV sumber atau generator. Saya tidak memakai `try ... otherwise 0` untuk menutup kesalahan bisnis.

## Batasan

Skrip M sudah ditinjau terhadap skema CSV dan definisi KPI, tetapi tidak dieksekusi oleh mesin Power Query dalam lingkungan verifikasi ini. Eksekusi dan refresh harus diperiksa di Excel Windows. Dukungan Excel Mac bergantung versi; LibreOffice dan Google Sheets tidak mengeksekusi query M ini. File XLSX tetap dapat dipakai melalui rumus yang telah dihitung ulang.
