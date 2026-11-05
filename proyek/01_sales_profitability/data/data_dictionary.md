# Kamus data

Saya memakai data sintetis dengan seed tetap dari generator. ID disimpan sebagai teks; CSV memakai UTF-8, delimiter koma, desimal titik, dan tanggal ISO. Tidak ada nama atau transaksi dari sumber perusahaan nyata.

## Grain dan relasi

Transaksi mempunyai satu baris per order dan satu produk. Retur mempunyai satu baris per kejadian retur. Produk, pelanggan, wilayah, dan salesperson mempunyai kunci tunggal unik. Target mempunyai satu baris per bulan dan region. Retur diatribusi kembali ke order asal sebelum agregasi; target direkap terpisah dari order.

## transaksi.csv

Jumlah baris: 6000. Kunci: `order_id`.

| Kolom | Tipe/satuan | Arti dan aturan |
|---|---|---|
| order_id | teks | Kunci order; satu order memuat satu produk |
| tanggal | tanggal ISO | Tanggal order; dasar atribusi periode |
| customer_id | teks | Relasi master_pelanggan |
| product_id | teks | Relasi master_produk |
| region_id | teks | Relasi master_wilayah |
| salesperson_id | teks | Relasi master_sales |
| qty | bilangan bulat | Jumlah unit sebelum retur; positif |
| harga_satuan | rupiah per unit | Harga transaksi sebelum diskon; dapat berbeda dari harga_list |
| diskon_pct | rasio | Diskon sebagai pecahan, bukan angka persen tampilan |

## retur.csv

Jumlah baris: 120. Kunci: `retur_id`.

| Kolom | Tipe/satuan | Arti dan aturan |
|---|---|---|
| retur_id | teks | Kunci retur |
| order_id | teks | Relasi transaksi; beberapa retur dapat mengacu satu order |
| tanggal_retur | tanggal ISO | Tanggal kejadian retur; tidak menentukan bulan laporan laba |
| qty_retur | bilangan bulat | Unit dikembalikan; total per order tidak melebihi qty |
| alasan | teks | Alasan sintetis |

## master_produk.csv

Jumlah baris: 40. Kunci: `product_id`.

| Kolom | Tipe/satuan | Arti dan aturan |
|---|---|---|
| product_id | teks | Kunci produk |
| nama_produk | teks | Nama produk sintetis |
| kategori | teks | Kelompok analisis |
| sub_kategori | teks | Rincian kategori |
| harga_list | rupiah per unit | Harga daftar; bukan harga transaksi |
| hpp_satuan | rupiah per unit | Biaya konstan selama periode |

## master_pelanggan.csv

Jumlah baris: 120. Kunci: `customer_id`.

| Kolom | Tipe/satuan | Arti dan aturan |
|---|---|---|
| customer_id | teks | Kunci pelanggan |
| nama_pelanggan | teks | Nama sintetis; kemiripan dengan nama nyata tidak dimaksudkan |
| segmen | teks | Retail, Grosir, atau Korporat |
| kota | teks | Label lokasi |
| region_id | teks | Relasi wilayah |
| tanggal_bergabung | tanggal ISO | Batas awal pelanggan dapat bertransaksi |

## master_wilayah.csv

Jumlah baris: 6. Kunci: `region_id`.

| Kolom | Tipe/satuan | Arti dan aturan |
|---|---|---|
| region_id | teks | Kunci wilayah |
| nama_region | teks | Nama tampilan region |
| pulau | teks | Kelompok geografis |

## master_sales.csv

Jumlah baris: 12. Kunci: `salesperson_id`.

| Kolom | Tipe/satuan | Arti dan aturan |
|---|---|---|
| salesperson_id | teks | Kunci salesperson |
| nama | teks | Nama sintetis |
| region_id | teks | Wilayah penugasan yang harus cocok dengan order |

## target.csv

Jumlah baris: 72. Kunci: `bulan + region_id`.

| Kolom | Tipe/satuan | Arti dan aturan |
|---|---|---|
| bulan | teks yyyy-mm | Periode target |
| region_id | teks | Wilayah target |
| target_penjualan_bersih | rupiah | Target sintetis; positif dan tidak boleh digandakan lewat join order |

## Kolom turunan dan output

Raw_Data!J:X berisi omzet_bruto, nilai_diskon, omzet_neto, qty_retur, nilai_retur, penjualan_bersih, hpp_satuan, hpp, laba_kotor, margin_kotor, periode, kategori, segmen, kunci_unik, dan duplikat. Definisi aritmetika ada di [definisi KPI](../docs/definisi_kpi.md). `duplikat` bernilai satu jika kunci order dan produk muncul berulang; order unik diperiksa juga secara terpisah.

`processed/rekonsiliasi.csv` menyimpan pemeriksaan, lokasi sel, nilai pandas, nilai workbook, selisih, dan status. `processed/ringkasan.json` menyimpan total, rasio, ringkasan region/kategori, hitungan rumus/error, dan SHA-256 tiap CSV input. Keduanya adalah hasil pemeriksaan, bukan input workbook.
