# Modul 10: Power Query

> Saya memakai ekspor pesanan sintetis sebagai latihan transformasi. Rumus Excel pada dokumen memakai titik koma untuk regional Indonesia; regional Inggris memakai koma. Skrip M mengikuti sintaks bahasanya sendiri. Query belum tertanam dalam workbook dan belum dijalankan pada mesin Power Query.

## 1. Tujuan dan prasyarat

Saya belajar membangun alur impor yang dapat diulang: membaca CSV dari folder, menetapkan tipe, menggabungkan baris dengan Append, menambah atribut dengan Merge, dan mengubah tabel lebar menjadi panjang melalui Unpivot. Saya juga memeriksa jumlah baris agar master duplikat tidak menggandakan hasil tanpa disadari.

Prasyaratnya adalah tabel yang rapi, kunci unik, tanggal, serta rekonsiliasi. Saya menggunakan [latihan](latihan/latihan_10.xlsx) sebagai pembanding dan menjalankan langkah Power Query pada salinan di Excel yang mendukungnya. [Solusi](latihan/solusi_10.xlsx) menyediakan hasil rumus biasa, bukan koneksi query aktif. Tujuannya agar hasil manual mempunyai angka pembanding yang jelas.

## 2. Kegunaan di pekerjaan

Ekspor berkala sering datang sebagai beberapa CSV dengan struktur sama. Menyalin semuanya manual mudah menimbulkan baris tertinggal atau tergandakan. Power Query menyimpan urutan transformasi sehingga sumber baru dapat diproses melalui Refresh, selama struktur dan aturan sumber tetap sesuai.

Otomatisasi transformasi tidak menjamin isi sumber benar. Jika master memuat kode duplikat, Merge dapat memperbanyak baris setelah hasil diperluas. Saya mengukur jumlah baris sebelum dan sesudah, memeriksa kunci, serta mencocokkan total. Latihan ini menunjukkan disiplin proses, bukan klaim pernah mengelola integrasi perusahaan nyata.

## 3. Penjelasan dari dasar

Query adalah definisi langkah transformasi yang membaca sumber dan menghasilkan tabel. Applied Steps menyimpan perubahan seperti tipe data, pilihan kolom, dan join. Hasil dapat dimuat ke worksheet atau hanya menjadi koneksi. Saya membedakan definisi query, data sumber, dan tabel hasil agar tidak mengedit keluaran yang akan ditimpa saat Refresh.

Append menumpuk baris dari tabel yang sejenis. Nama kolom harus konsisten karena penggabungan mengikuti kolom, bukan posisi visual semata. Kolom hilang atau nama berbeda dapat menghasilkan null dan struktur yang tidak diharapkan. Saya memeriksa skema file sebelum mempercayai tabel gabungan.

Merge menghubungkan tabel berdasarkan kunci. Pada simulasi ini banyak pesanan mengarah ke satu produk master, sehingga Left Outer menjaga seluruh pesanan. Expand mengambil kategori dan harga dari tabel hasil join. Kunci yang tidak ditemukan harus diperiksa; saya tidak mengisi harga tebakan untuk menghilangkan null.

Relasi many-to-one memerlukan kunci unik pada master. Jika satu kode cocok dengan beberapa baris, Expand dapat menggandakan pesanan. Total mungkin naik walau tidak ada error tipe. Skrip yang disediakan menghitung frekuensi master dan menolak duplikat, lalu memeriksa jumlah baris sesudah Expand.

Unpivot memindahkan nama kolom pengukuran menjadi nilai pada kolom atribut. Target dengan kolom Januari dan Februari diubah menjadi wilayah, periode, dan target. Saya menjaga kolom identifier tetap, kemudian mencocokkan total sebelum dan sesudah. Tabel panjang lebih mudah digabung untuk periode baru.

Impor folder memerlukan aturan seleksi file. Folder dapat memuat master, target, keluaran, atau file sementara yang tidak boleh di-Append sebagai pesanan. Skrip hanya memilih nama berawalan `pesanan_` dengan ekstensi CSV. Saya menyimpan hasil ekspor dengan pola nama berbeda agar tidak terimpor kembali pada Refresh berikutnya.

Tipe data ditetapkan setelah header dipromosikan. ID dan kode menjadi teks, qty serta harga menjadi bilangan bulat, dan tanggal ISO diparse dengan format eksplisit. Saya tidak mengandalkan tebakan regional untuk tanggal karena alur yang dapat diulang perlu aturan yang sama setiap kali dijalankan.

Refresh menjalankan kembali langkah pada sumber saat ini. Perubahan nama kolom, lokasi, izin akses, atau tipe dapat membuat langkah gagal. Saya menyimpan parameter folder secara terpisah agar pemindahan lokasi tidak memerlukan perubahan di banyak ekspresi. Kredensial tidak dimasukkan ke skrip atau workbook latihan.

## 4. Dataset latihan

Sumber terdiri atas [pesanan Januari](latihan/pesanan_januari.csv), [pesanan Februari](latihan/pesanan_februari.csv), [master produk](latihan/master_produk.csv), dan [target lebar](latihan/target_wide.csv). Masing-masing file pesanan mempunyai 100 baris dengan ID, tanggal, kode, dan qty. Master berisi 12 produk berkunci unik.

[Data gabungan](latihan/data_10.csv) mempertahankan input kedua periode, sedangkan [hasil referensi](latihan/hasil_referensi.csv) merupakan keluaran pandas sebagai pembanding terpisah. Workbook menghitung harga dan nilai dengan formula. Sheet Hasil_Query_Diharapkan menyediakan ringkasan bulan serta bentuk target panjang.

Skrip lengkap tersedia sebagai [parameter folder](pq/pFolder.pq), [Append dan Merge](pq/append_merge.pq), serta [Unpivot target](pq/unpivot_target.pq). [Kunci jawaban](latihan/kunci_jawaban.md) dan [VERIFIKASI.md](VERIFIKASI.md) mencatat hasil pandas serta recalc formula, bukan bukti eksekusi mesin M.

## 5. Langkah praktik bernomor dengan instruksi klik

1. Salin CSV sumber ke satu folder lokal. Saya mempertahankan pola nama pesanan dan memastikan file master serta target tidak memakai awalan yang sama.
2. Di Excel, pilih **Data > Get Data > From Other Sources > Blank Query**. Buka **Home > Advanced Editor**, tempel isi pFolder.pq, ubah lokasi sesuai folder lokal, lalu beri nama query `pFolder` pada **Query Settings > Name**.
3. Buat Blank Query kedua, buka Advanced Editor, dan tempel append_merge.pq lengkap. Query ini memakai pFolder, memilih CSV pesanan, mempromosikan header, serta menggabungkan seluruh tabel yang dipilih.
4. Periksa Applied Steps sampai Appended dan Dates. Cocokkan nama kolom, tipe, serta jumlah baris dengan gabungan sumber. Saya tidak hanya meninjau beberapa baris pratinjau.
5. Pilih langkah CheckedMaster dan Joined. Skrip memakai Left Outer dengan kode sebagai kunci. Pada Expanded, periksa kategori dan harga; kunci yang tidak tersedia harus menjadi masalah yang ditindaklanjuti.
6. Bandingkan jumlah baris sebelum dan setelah Expand. Jika bertambah, periksa frekuensi kunci pada master sebelum menghapus duplikat hasil. Menghapus hasil ganda tanpa memahami sumber dapat menyembunyikan masalah relasi.
7. Pilih **Home > Close & Load > Close & Load To**, lalu **Table** pada worksheet baru. Saya memberi nama keluaran yang berbeda dari sumber supaya Refresh tidak membacanya sebagai CSV masukan.
8. Buat Blank Query berikutnya, tempel unpivot_target.pq, dan beri nama `Target_Panjang`. Periksa bahwa wilayah tetap identifier sedangkan nama periode berpindah menjadi nilai. Muat hasil ke sheet baru.
9. Untuk mempraktikkan antarmuka Unpivot, pada salinan query target pilih kolom wilayah, lalu **Transform > Unpivot Columns > Unpivot Other Columns**. Bandingkan hasil dengan skrip yang disediakan.
10. Tambahkan file pesanan berstruktur sama pada salinan folder, lalu **Data > Refresh All**. Periksa kenaikan jumlah baris dan total, bukan hanya keberhasilan pesan Refresh.
11. Kembalikan sumber awal untuk membandingkan dengan bukti bawaan. Isi ringkasan kuning, cocokkan nilai setiap bulan, dan pastikan total target long sama dengan wide.

## 6. Contoh formula dan arti tiap argumen

Dalam workbook, `=INDEX(Master!$C$2:$C$13;MATCH(C2;Master!$A$2:$A$13;0))` mengambil harga sebagai alternatif lookup biasa. Nilai pesanan dihitung dari qty dan harga. Rumus ini merupakan pembanding terhadap hasil Merge dan custom column M, bukan query yang disembunyikan dalam sel.

Pada M, `Table.NestedJoin` menerima tabel kiri, daftar kunci kiri, tabel kanan, daftar kunci kanan, nama kolom hasil, dan jenis join. `Table.ExpandTableColumn` memilih atribut dari tabel bersarang. Saya memeriksa jumlah baris setelah perluasan karena penggandaan sering baru terlihat pada tahap ini.

`Table.UnpivotOtherColumns` mempertahankan daftar identifier dan mengubah kolom lain menjadi pasangan atribut serta nilai. `Table.RowCount` dipakai untuk audit volume. Skrip memunculkan error yang jelas ketika master duplikat atau harga tidak ditemukan; saya tidak mengganti masalah itu dengan nol agar alur tampak berhasil.

## 7. Latihan mandiri minimal lima soal bertingkat

1. Dasar: jelaskan perbedaan Append dan Merge dengan kedua file pesanan serta master.
2. Dasar: gabungkan CSV dan cocokkan jumlah baris tanpa menghitung header sebagai data.
3. Menengah: lakukan Merge Left Outer dan jelaskan mengapa master harus unik.
4. Menengah: ubah target wide menjadi long dan buktikan totalnya tetap sama.
5. Menengah: jelaskan aturan seleksi file folder dan risiko memasukkan keluaran kembali sebagai sumber.
6. Lanjut: tambahkan kunci master duplikat pada salinan dan jelaskan pemeriksaan yang harus menghentikan proses.
7. Lanjut: tambahkan file periode baru pada salinan, lakukan Refresh, dan rancang rekonsiliasi volume serta nilai.

## 8. Kunci jawaban

[Kunci lengkap](latihan/kunci_jawaban.md) memuat jumlah gabungan, nilai per bulan, dan target panjang. Append menjumlahkan baris sumber; Merge menambah atribut sambil menjaga jumlah pesanan untuk relasi yang diharapkan. Unpivot mengubah bentuk, bukan menambah total target.

Saya menerima keluaran manual jika kunci, jumlah baris, nilai, dan periode sesuai pembanding. Jika query gagal karena sumber berubah, saya menyebut langkah serta penyebabnya. Keberhasilan Refresh saja belum membuktikan cakupan file sudah benar.

## 9. Kesalahan umum dan cara memeriksa hasil

Mengimpor seluruh CSV tanpa pola nama dapat mencampurkan master dan keluaran. Menetapkan kode sebagai angka dapat menghilangkan nol depan. Merge pada kunci yang tidak unik dapat menggandakan nilai tanpa error. Saya memeriksa daftar file, tipe, serta frekuensi kunci sebelum membandingkan total.

Mengedit tabel hasil secara manual berisiko kehilangan perubahan saat Refresh. Saya memperbaiki sumber atau langkah query, lalu menjalankan ulang. Parameter folder harus diperbarui ketika lokasi dipindah; saya tidak menyisipkan kredensial pada teks query.

## 10. Checklist kompetensi

- [ ] Saya dapat menjelaskan urutan impor dan Applied Steps.
- [ ] Saya dapat membedakan Append, Merge, serta Unpivot.
- [ ] Saya dapat menjaga tipe dan keunikan kunci sebelum join.
- [ ] Saya dapat membuktikan jumlah baris serta total tidak berubah tanpa alasan.
- [ ] Saya dapat menjalankan Refresh dengan cakupan sumber yang dinyatakan.

## 11. Catatan kompatibilitas Excel 2016/2019/365, LibreOffice, dan Google Sheets

Power Query tersedia pada Excel desktop Windows yang mendukung Get & Transform; kemampuan Mac bergantung versi. Workbook ini tidak mempunyai koneksi Power Query tertanam. Skrip M harus ditempel dan dijalankan manual; mesin M belum diuji di lingkungan ini.

LibreOffice hanya menghitung rumus pembanding dan menghasilkan [preview nyata](assets/preview.png); ia tidak menjalankan alur Power Query yang dijelaskan. Google Sheets tidak menjalankan skrip M secara langsung. Pembangunan ulang memakai `python tools/build.py`, lalu `python tools/course_runtime.py 10_power-query` dari folder belajar. Saya menyimpan hasil manual pada salinan dan mencatat sumber sebelum Refresh.
