# Modul 02: Struktur data yang benar

> Materi latihan dengan data sintetis yang saya bangkitkan melalui skrip, bukan data perusahaan atau pengalaman kerja. Rumus dalam dokumen memakai titik koma untuk Excel regional Indonesia; regional Inggris memakai koma. Saya mempertahankan nama fungsi Inggris dalam file.

## 1. Tujuan dan prasyarat

Saya belajar menyusun data agar satu tabel dapat dipakai berulang untuk perhitungan, pemeriksaan, dan laporan. Setelah modul ini, saya dapat menentukan arti satu baris, memilih nama kolom yang konsisten, menjaga tipe data, membedakan data transaksi dari ringkasan, serta membuat Excel Table dengan structured references. Saya juga dapat menjelaskan mengapa kode cabang dengan nol di depan harus diperlakukan sebagai teks.

Prasyarat saya adalah modul dasar: memilih rentang, membaca Formula Bar, mengubah format, menggunakan Sort dan Filter, serta membedakan nilai dengan tampilan. Saya memakai [workbook latihan](latihan/latihan_02.xlsx), menyimpan salinan pekerjaan, lalu membandingkannya dengan [solusi](latihan/solusi_02.xlsx). Saya tidak perlu menguasai Power Query terlebih dahulu. Transformasi wide ke long di sini dapat dilakukan manual pada tabel kecil sebelum dipelajari otomatis pada modul berikutnya.

## 2. Kegunaan di pekerjaan

Dalam simulasi rekap cabang, laporan yang terlihat rapi belum tentu cocok sebagai sumber analisis. Judul gabungan, subtotal di tengah transaksi, dan bulan yang tersebar sebagai kolom membuat penambahan periode sulit diperiksa. Jika satu rumus menghitung subtotal bersama transaksi, hasil dapat terhitung ganda tanpa pesan kesalahan.

Saya memisahkan penyimpanan data dari penyajian laporan. Tabel sumber mempertahankan detail, sedangkan laporan merangkum detail tersebut. Ketika muncul pertanyaan baru tentang produk, kategori, atau waktu, saya cukup membuat agregasi lain tanpa mengetik ulang angka. Struktur ini berguna untuk administrasi, penjualan, pengadaan, dan analisis, tetapi latihan ini tidak menyatakan saya pernah menangani sistem perusahaan nyata.

Kualitas struktur juga memengaruhi serah terima. Orang lain perlu mengetahui apakah satu baris mewakili transaksi, produk dalam transaksi, atau total cabang. Pernyataan tersebut disebut grain. Saya menulis grain sebelum memilih rumus sehingga jumlah baris tidak keliru ditafsirkan sebagai jumlah pelanggan atau jumlah barang.

## 3. Penjelasan dari dasar

Tidy data mempunyai satu variabel per kolom dan satu pengamatan per baris. Pada dataset ini satu baris adalah satu catatan transaksi produk pada satu tanggal dan cabang. Kolom `qty` berisi banyak barang, sedangkan `harga_satuan` berisi rupiah per barang. Hasil perkaliannya adalah nilai transaksi, bukan jenis data baru yang perlu diketik manual.

Header harus tunggal dan jelas. Saya memakai nama seperti `cabang_id`, bukan tiga kolom dengan nama sama atau header bertingkat. Baris kosong tidak dipakai sebagai pemisah kelompok. Catatan dan judul ditempatkan pada sheet Instruksi atau di luar tabel, sehingga area data tetap berbentuk persegi yang dapat difilter.

Kunci dan label mempunyai fungsi berbeda. `cabang_id` mengenali cabang secara stabil; `nama_cabang` membantu pembaca. ID `001` adalah teks, bukan bilangan yang perlu dijumlahkan. Mengubah angka menjadi format `000` hanya mengubah tampilannya. Fungsi TEXT menghasilkan teks baru, sedangkan VALUE mengubah teks numerik menjadi angka. Saya memilih tipe berdasarkan arti bisnis, bukan keseragaman visual.

Tanggal Excel adalah nilai tanggal yang dapat diurutkan dan dipakai sebagai batas periode. CSV menyimpan tanggal dalam bentuk ISO; saat dibuka, aplikasi dapat menebak tipe secara berbeda. Saya memeriksa beberapa tanggal melalui Formula Bar dan COUNT, lalu memastikan ID tetap Text saat impor. Mengganti format saja tidak memperbaiki tanggal yang sudah salah ditafsirkan.

Excel Table adalah objek tabel, berbeda dari range yang hanya diberi warna. Table mempunyai nama, header, filter, perluasan baris, dan structured references. `tblPenjualan[qty]` berarti isi kolom qty pada tabel bernama tblPenjualan. `[@qty]` berarti qty pada baris yang sama dalam calculated column. Referensi ini membantu keterbacaan, tetapi nama tabel dan kolom harus tepat.

Wide data menempatkan periode pada banyak kolom. Bentuk ini sesuai untuk membaca laporan kecil, tetapi kurang praktis sebagai sumber gabungan banyak periode. Long data menyimpan periode sebagai nilai pada satu kolom dan nilai pengukuran pada kolom lain. Saya tidak menjumlahkan ulang total laporan lebar yang sudah mencakup cabang sama dua kali.

Merged cells tetap berguna untuk judul laporan, tetapi saya menghindarinya di area transaksi. Sel gabungan mempersulit sort, filter, pemilihan baris, dan ekspor. Sheet Laporan_Lama sengaja menunjukkan anti-pola tersebut sebagai bahan pemeriksaan; tampilannya tidak dijadikan standar sumber data.

## 4. Dataset latihan

[CSV sumber](latihan/data_02.csv) berisi 240 catatan dari Januari sampai April 2024. Kolom mentahnya adalah nomor, tanggal, ID dan nama cabang, kode produk, kategori, qty, dan harga satuan. Dataset menggunakan label cabang sintetis dan daftar produk simulasi. Kolom total dalam workbook dihitung melalui rumus dari qty dan harga satuan.

Sheet `Laporan_Lama` adalah snapshot input laporan dengan beberapa angka tersimpan sebagai teks. `Data` adalah sumber tidy; pada latihan masih range biasa, sedangkan solusi sudah mempunyai Table. `ID_Rusak` menunjukkan angka yang kehilangan nol di depan. `Latihan` menyediakan area pemeriksaan tipe, agregasi, rekonsiliasi laporan, serta perubahan wide ke long. Snapshot laporan lama dipertahankan untuk pembanding, sedangkan subtotal dan pemeriksaan dihitung dengan rumus.

Saya tidak menghapus snapshot hanya karena strukturnya buruk. Saya menjaga jejak asal, memperbaiki sumber yang dipakai analisis, lalu mencocokkan totalnya. [Kunci jawaban](latihan/kunci_jawaban.md) memuat hasil pandas dari CSV; [VERIFIKASI.md](VERIFIKASI.md) menunjukkan hasil penghitungan ulang workbook.

## 5. Langkah praktik bernomor dengan instruksi klik

1. Buka latihan, baca Instruksi, lalu pilih **Formulas > Calculation Options > Automatic**. Saya menyimpan salinan sebelum mulai agar file awal tetap tersedia untuk perbandingan.
2. Pada Data, klik A1, tekan **Ctrl+T**, pastikan rentang mencakup seluruh kolom dan baris, lalu centang **My table has headers**. Pilih **Table Design > Table Name** dan beri nama `tblPenjualan`.
3. Pilih kolom tanggal, tekan **Ctrl+1 > Custom**, lalu gunakan `dd/mm/yyyy`. Untuk ID, periksa isi Formula Bar dan pastikan nol di depan tetap bagian dari teks, bukan sekadar tampilan.
4. Pada ID_Rusak, isi kolom TEXT berdasarkan angka sumber. Pada kolom pembanding gunakan referensi ke angka dengan format `000`. Saya membandingkan tampilan keduanya sambil membaca nilai sebenarnya.
5. Pada Latihan bagian A, isi COUNT dan COUNTA sesuai rentang yang tertulis. Kurangkan keduanya untuk menemukan sel terisi yang bukan angka. Saya tidak langsung menyebut semua teks sebagai kesalahan karena ID memang seharusnya teks.
6. Isi bagian structured references. Bila muncul pesan nama tidak dikenal, saya memeriksa nama tabel dan header melalui **Table Design**, bukan mengganti hasil dengan angka.
7. Isi matriks cabang dan bulan memakai SUMIFS dengan batas tanggal awal inklusif dan awal bulan berikutnya eksklusif. Setelah setiap baris terisi, jumlahkan bulan dan bandingkan dengan snapshot laporan lama.
8. Bentuk tabel long dari matriks kecil: ulangi nama cabang untuk setiap bulan dan masukkan bulan sebagai satu kolom. Nilai hasil tetap merujuk tabel sumber, sehingga perubahan sumber dapat ditelusuri.
9. Gunakan **Data > Sort**, pilih **Expand the selection**, dan periksa nomor transaksi setelah pengurutan. Saya tidak mengurutkan satu kolom sendirian karena hubungan antarkolom akan rusak.
10. Tekan **Ctrl+Alt+F9**, periksa selisih rekonsiliasi, lalu bandingkan dengan solusi dan catatan verifikasi. Saya menyimpan kesimpulan tentang struktur, bukan hanya hasil total.

## 6. Contoh formula dan arti tiap argumen

`=TEXT(A4;"000")` mengubah nilai A4 menjadi teks tiga digit. A4 adalah nilai sumber; format string menentukan jumlah digit. Untuk identifier, hasil teks dapat menjadi kunci. `=VALUE(C4)` melakukan arah sebaliknya dan akan menghilangkan nol depan yang tidak memiliki arti numerik.

`=COUNT(Data!C2:C241)` menghitung nilai numerik pada rentang ID; hasil nol merupakan keadaan yang benar karena kolom tersebut teks. `=COUNTA(Data!C2:C241)` menghitung seluruh sel terisi. Selisih COUNTA dan COUNT membantu membaca tipe, tetapi bukan validator lengkap terhadap duplikat atau karakter tersembunyi.

`=SUM(tblPenjualan[total])` menjumlahkan seluruh nilai kolom total. Pada calculated column, `=[@qty]*[@harga_satuan]` mengalikan nilai pada baris yang sama. Alternatif range biasa adalah `=G2*H2`, lalu isi ke bawah. Saya memakai nama kolom yang sama persis dengan header untuk menghindari salah referensi.

`=SUMIFS(tblPenjualan[total];tblPenjualan[nama_cabang];"Surabaya")` mempunyai argumen rentang nilai, rentang kriteria, dan nilai kriteria. Untuk periode saya menambah pasangan rentang tanggal dengan kriteria `">="&DATE(2024;1;1)` serta `"<"&DATE(2024;2;1)`. Kedua batas mencegah transaksi di luar bulan ikut terjumlah.

## 7. Latihan mandiri minimal lima soal bertingkat

1. Dasar: tuliskan grain tabel Data dan jelaskan mengapa subtotal tidak ditempatkan di antara transaksi.
2. Dasar: perbaiki ID pada ID_Rusak dengan TEXT, lalu jelaskan perbedaannya dengan custom number format.
3. Menengah: hitung total penjualan melalui structured references dan cocokkan dengan penjumlahan range biasa.
4. Menengah: hitung penjualan Surabaya pada Januari menggunakan batas tanggal, tanpa mengetik hasil snapshot.
5. Menengah: bentuk tabel long cabang dan bulan, lalu buktikan totalnya sama dengan matriks sumber.
6. Lanjut: identifikasi risiko menghitung subtotal bersama detail pada Laporan_Lama; tunjukkan baris mana yang harus dikecualikan dari penggabungan data transaksi.
7. Lanjut: tambahkan satu baris pada salinan Table, amati perluasan kolom total, lalu batalkan perubahan sebelum membandingkan hasil dengan bukti awal.

## 8. Kunci jawaban

[Kunci lengkap](latihan/kunci_jawaban.md) menghubungkan pertanyaan dengan formula, lokasi sel, dan nilai pembanding pandas. Grain adalah transaksi produk per tanggal dan cabang; subtotal merupakan hasil agregasi sehingga tidak menjadi transaksi baru. ID resmi dipertahankan sebagai teks, sedangkan format angka hanya membantu tampilan.

Saya menerima transformasi long jika setiap kombinasi cabang dan bulan tercatat sekali, nama kolom konsisten, serta selisih total long terhadap wide nol. Total yang sama adalah pemeriksaan penting, tetapi tetap perlu memeriksa jumlah kombinasi agar kehilangan dan duplikasi tidak saling menutup.

## 9. Kesalahan umum dan cara memeriksa hasil

Kesalahan pertama adalah mengira range berwarna otomatis merupakan Table. Saya memeriksa apakah tab Table Design muncul ketika salah satu sel dipilih. Kesalahan berikutnya adalah menamai tabel berbeda dengan formula; saya memperbaiki sumber nama melalui dialog, bukan menutupi pesan kesalahan.

ID yang sudah berubah menjadi angka mungkin tidak dapat dipulihkan tanpa aturan panjang atau master resmi. Di latihan ini panjangnya diketahui, tetapi saya tidak menerapkan padding sembarang pada data nyata. Saya juga memeriksa tanggal, karena urutan teks ISO dapat terlihat benar walau tipe belum sesuai untuk rumus.

Jika rekonsiliasi tidak nol, saya memeriksa batas bulan, cakupan range, subtotal ganda, dan sel angka sebagai teks. Saya membandingkan cabang per bulan terlebih dahulu; pemeriksaan detail lebih membantu daripada mengganti total akhir agar terlihat sama.

## 10. Checklist kompetensi

- [ ] Saya dapat menjelaskan grain dan membedakan transaksi dengan agregasi.
- [ ] Saya dapat membuat serta menamai Excel Table dan membaca structured references.
- [ ] Saya dapat menjaga ID teks dan menguji angka versus tampilan.
- [ ] Saya dapat mengubah wide menjadi long tanpa menggandakan total.
- [ ] Saya dapat merekonsiliasi hasil per cabang dan periode serta menjelaskan selisih.

## 11. Catatan kompatibilitas Excel 2016/2019/365, LibreOffice, dan Google Sheets

Excel 2016, 2019, dan 365 mendukung Table, SUMIFS, TEXT, serta referensi terstruktur yang dipakai. File tidak memerlukan dynamic array. LibreOffice dipakai untuk recalc dan render; representasi Table dapat berubah setelah penyimpanan ulang, sehingga saya memeriksa hasil angka dan nama tabel ketika membukanya kembali.

Google Sheets mempunyai mekanisme tabel dan referensi yang dapat berbeda dari Excel; alternatifnya adalah rentang biasa dengan referensi absolut. Saya belum menguji interaksi langsung di Excel atau Google Sheets. Preview berasal dari [render workbook](assets/preview.png), bukan gambar yang digambar ulang. Pembangunan ulang memakai `python tools/build.py` dari folder modul; verifikasi memakai `python tools/course_runtime.py 02_struktur-data-yang-benar` dari folder belajar. Saya tidak mengklaim PivotTable atau Power Query sudah tertanam pada file ini.
