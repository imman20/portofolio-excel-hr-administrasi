# DAX untuk model manual

Saya membuat tabel model bernama Fakt, Produk, dan Kalender. Rumus berikut memakai titik koma mengikuti regional Indonesia; bila editor DAX memakai koma, saya menyesuaikan pemisahnya. Rumus ini belum dijalankan pada mesin DAX dan tidak tertanam di XLSX.

Calculated column pada Fakt:

```dax
Nilai = Fakt[qty] * Fakt[harga_satuan]
```

Saya menghitung satu nilai per baris; ini berbeda dari measure yang merespons filter laporan.

```dax
Penjualan := SUM(Fakt[Nilai])
Qty := SUM(Fakt[qty])
Jumlah Pesanan := DISTINCTCOUNT(Fakt[order_id])
Nilai per Pesanan := DIVIDE([Penjualan]; [Jumlah Pesanan]; 0)
Penjualan ATK := CALCULATE([Penjualan]; Produk[kategori] = "ATK")
Penjualan 2024 := CALCULATE([Penjualan]; Kalender[tahun] = 2024)
```

Pada dialog New Measure, nama diisi pada kotak Measure Name dan bagian Formula dimulai dengan `=`; sintaks `:=` di atas adalah notasi daftar measure. SUM menjumlahkan kolom; DISTINCTCOUNT menghitung ID berbeda dalam filter aktif; DIVIDE menyediakan hasil alternatif saat pembagi nol. CALCULATE mengevaluasi measure dengan filter yang dinyatakan. Filter kategori ATK mengganti filter yang ada pada kolom kategori tersebut; saya tidak menafsirkannya otomatis sebagai irisan semua pilihan.

Relasi yang diperlukan: Produk[product_id] satu ke Fakt[product_id] banyak, dan Kalender[tanggal] satu ke Fakt[tanggal] banyak. Filter mengalir dari dimensi ke fakta. Saya memeriksa kunci unik serta kunci fakta yang tidak dikenal sebelum membuat relasi.

Jumlah Pesanan tidak aditif antarkategori karena satu pesanan dapat memuat beberapa kategori. Grand total DISTINCTCOUNT dihitung ulang dalam konteks keseluruhan, bukan menjumlahkan nilai kategori. Pembanding workbook menggunakan flag kemunculan pertama yang sesuai konteks; flag global tidak dipakai untuk menghitung distinct per kategori.
