# Cleaning dan transformasi

Saya mempertahankan CSV di `data/raw/` sebagai input yang dapat dibangun ulang. Data Sales sudah berbentuk tabel rapi; saya tidak menyisipkan masalah kualitas secara sengaja pada proyek ini. Pemeriksaan tetap diperlukan karena CSV rapi belum membuktikan kunci, relasi, dan definisi bisnisnya benar.

## Urutan pengolahan

1. Saya membaca ID sebagai teks dan tanggal ISO sebagai tanggal asli, lalu memeriksa nilai kosong, kunci master unik, order unik, serta pasangan bulan dan region unik pada target.
2. Saya memastikan product_id, customer_id, region_id, dan salesperson_id tersedia di master. Region pelanggan dan region salesperson harus sama dengan region order. Transaksi tidak boleh mendahului tanggal bergabung pelanggan.
3. Saya memeriksa qty positif dan bulat, harga positif, diskon dalam rentang yang dirancang, serta tanggal di dalam periode analisis. Retur harus positif, memiliki order asal, dan tidak mendahului order.
4. Saya mengagregasi retur menurut order sebelum menggabungkannya ke transaksi. Order tanpa retur memperoleh qty_retur nol. Ini menjaga grain satu baris per order dan mencegah penjualan terduplikasi bila satu order mempunyai beberapa retur.
5. Saya mengambil HPP dan kategori dari master produk, segmen dari master pelanggan, lalu membangun omzet, diskon, neto, nilai retur, penjualan bersih, HPP neto, dan laba. Retur kumulatif tidak boleh melebihi qty order.
6. Saya membuat periode `yyyy-mm` dari tanggal order, kunci gabungan order dan produk, serta penanda duplikat. Keunikan order juga diperiksa terpisah karena model ini satu produk per order.
7. Saya merekap transaksi dengan SUMIFS dan menjumlahkan target dari master pada grain terpisah. Saya menghitung rasio sesudah agregasi.

## Jejak transformasi

`Raw_Data!A:I` berasal dari transaksi.csv; `J:X` adalah rumus hidup. `Retur!A:E` berasal dari retur.csv; kolom F adalah pemeriksaan relasi. Master memuat tabel produk, pelanggan, wilayah, sales, dan target secara berdampingan. Saya tidak menghapus transaksi rugi, retur penuh, atau diskon tinggi untuk memperbaiki hasil laporan.

Power Query yang disediakan mengulangi alur yang sama dari CSV. Saya tidak memakai Remove Duplicates untuk menyembunyikan konflik data: query berhenti bila kunci utama duplikat atau relasi tidak valid. Retur tetap diagregasi, bukan dihapus. Hasil query harus mempunyai jumlah baris sama dengan transaksi.csv dan total sama dengan Calculation. Lihat [panduan query](power_query.md).

## Batasan pemeliharaan

Range rumus pada workbook disusun oleh builder sesuai ukuran input simulasi. Menambah baris melalui Excel Table membantu filter, tetapi rekap memakai range absolut yang tidak dijamin meluas otomatis. Setelah ukuran input berubah, saya menjalankan kembali builder dan verifier; saya tidak hanya menempel baris lalu menganggap laporan sudah lengkap. Builder saat ini mengharapkan struktur master dan target simulasi yang sama.
