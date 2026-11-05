# Modul 04: Logika dan agregasi bersyarat

> Saya memakai simulasi tiket layanan untuk latihan, bukan catatan pelanggan nyata. Rumus dokumen memakai titik koma pada Excel regional Indonesia; regional Inggris memakai koma. Semua hasil contoh dapat ditelusuri ke CSV dan workbook.

## 1. Tujuan dan prasyarat

Saya belajar menerjemahkan aturan layanan menjadi syarat yang jelas, kemudian menggabungkannya tanpa kehilangan arti. Fungsi yang saya gunakan adalah IF, IFS, IF bersarang, AND, OR, SUMIFS, COUNTIFS, dan AVERAGEIFS. Saya juga mempelajari wildcard serta batas tanggal agar ringkasan periode tidak memasukkan catatan yang salah.

Prasyaratnya adalah formula dasar dan referensi absolut. Saya sudah dapat menjelaskan mengapa parameter harus dikunci saat rumus disalin. Pada [latihan](latihan/latihan_04.xlsx), jawaban berwarna kuning sengaja kosong; [solusi](latihan/solusi_04.xlsx) memperlihatkan rumus lengkap. Saya menyimpan pekerjaan sendiri sebagai salinan agar perbandingan dengan keadaan awal tetap mungkin.

## 2. Kegunaan di pekerjaan

Dalam simulasi pusat layanan, tiket selesai belum tentu memenuhi batas waktu. Sebaliknya, tiket yang masih terbuka belum dapat disebut selesai tepat waktu hanya karena jam tercatat masih rendah. Aturan harus menyebut kedua kondisi sekaligus. Saya memisahkan status administratif dari ukuran kecepatan agar label tidak menyesatkan.

Agregasi bersyarat membantu menjawab pertanyaan yang lebih khusus daripada total keseluruhan: berapa biaya Email yang sudah selesai, bagaimana rata-rata durasi kanal tertentu, atau berapa tiket kategori dengan awalan tertentu. Saya mendefinisikan kelompok dan penyebut terlebih dahulu, lalu memilih fungsi. Ringkasan latihan ini tidak dipakai untuk menilai kinerja staf nyata.

## 3. Penjelasan dari dasar

Ekspresi logika menghasilkan TRUE atau FALSE. Pembanding seperti lebih kecil, sama dengan, dan lebih besar mempunyai makna batas yang harus ditulis eksplisit. Tiket pada batas jam layanan termasuk tepat waktu karena aturan memakai kurang dari atau sama dengan, bukan hanya kurang dari.

IF mempunyai tiga bagian: uji logika, hasil bila benar, dan hasil bila salah. Kedua hasil dapat berupa angka, teks, atau rumus lain. Teks keluaran ditulis dalam tanda petik. Saya menjaga agar satu kolom hasil mempunyai arti konsisten; mencampur durasi numerik dengan label bebas akan mempersulit agregasi berikutnya.

AND berarti semua syarat harus benar. Dalam latihan ini hasil TEPAT memerlukan status selesai dan jam tidak melebihi batas. OR berarti paling sedikit satu syarat benar. Tinjauan PRIORITAS diberikan bila tiket berlabel Urgent atau biayanya melewati ambang. Tiket yang memenuhi keduanya tetap satu tiket, bukan dua kejadian prioritas.

IFS memeriksa pasangan kondisi dan hasil secara berurutan, lalu mengambil hasil pada kondisi benar pertama. Saya menempatkan kelas durasi paling singkat sebelum kelas yang lebih luas. Jika kondisi luas ditempatkan lebih dahulu, kelompok singkat tidak pernah tercapai. Kondisi terakhir memeriksa jam di atas batas sehingga seluruh durasi numerik yang valid tercakup.

Untuk Excel 2016 yang tidak menyediakan IFS pada edisi tertentu, saya memakai IF bersarang. Fungsi ini membaca aturan yang sama, tetapi hasil salah dari IF pertama berisi IF berikutnya. Terlalu banyak tingkat membuat peninjauan sulit; jika aturan berkembang besar, saya mempertimbangkan tabel referensi yang diaudit, bukan menambah kurung tanpa batas.

COUNTIFS menghitung baris yang memenuhi seluruh pasangan syarat. SUMIFS menjumlahkan nilai pada baris yang memenuhi syarat tersebut. AVERAGEIFS menghitung rata-rata pada kelompok terpilih. Semua rentang kriteria harus sejajar dengan rentang nilai, sehingga baris yang dibandingkan tetap mewakili tiket sama.

Wildcard bintang mewakili rangkaian karakter dan tanda tanya mewakili satu karakter. Kriteria `Peng*` mencakup beberapa kategori dengan awalan sama. Jika karakter bintang harus dicari sebagai teks literal, saya memakai tilde untuk meloloskannya. Saya memeriksa kategori yang benar-benar masuk kelompok sebelum memberi nama ringkasan.

Tanggal adalah nilai numerik di Excel. Batas periode saya tulis sebagai awal bulan inklusif dan awal bulan berikutnya eksklusif. Pola tersebut tetap jelas jika sumber kelak mempunyai komponen jam. Saya tidak menulis tanggal ambigu sebagai teks dalam kriteria ketika DATE dapat menyatakan tahun, bulan, dan hari secara eksplisit.

## 4. Dataset latihan

[Data CSV](latihan/data_04.csv) memuat 220 tiket sintetis pada kuartal pertama 2024. Kolom mentah adalah ID, tanggal, kanal, kategori, jam, biaya, prioritas, dan status selesai. Kanal berisi Email, Telepon, atau Portal. Kategori Pengiriman dan Pengembalian sengaja tersedia untuk latihan wildcard.

Data menambahkan empat kolom berumus: hasil layanan, tinjauan prioritas, kelas IFS, dan kelas IF bersarang. Master menyimpan batas jam serta biaya prioritas sebagai input biru. Jam pada tiket terbuka adalah waktu tercatat dalam simulasi; saya tidak menafsirkannya sebagai durasi penyelesaian final.

Latihan memuat ringkasan kondisi, pembanding IFS terhadap IF, dan satu kelompok tanpa pengamatan. [Kunci jawaban](latihan/kunci_jawaban.md) memakai mask pandas dari CSV. [VERIFIKASI.md](VERIFIKASI.md) mencatat hasil penghitungan ulang serta perbandingan baris aturan layanan.

## 5. Langkah praktik bernomor dengan instruksi klik

1. Buka Instruksi, aktifkan **Formulas > Calculation Options > Automatic**, lalu baca parameter Master. Saya mencatat bahwa ambang ini adalah asumsi simulasi yang dapat diubah pada salinan.
2. Pada Data, pilih I2 dan tekan **F2**. Telusuri status selesai, durasi, dan batas jam. Saya memastikan label TEPAT tidak hanya bergantung pada salah satu kondisi.
3. Pilih J2 dan baca OR. Gunakan **Data > Filter** untuk melihat tiket yang memenuhi kedua syarat prioritas sekaligus. Saya menghitung setiap baris sekali ketika menilai gabungan.
4. Pada K2 dan L2, gunakan **Formulas > Evaluate Formula** untuk mengikuti IFS dan IF bersarang. Saya memeriksa urutan kondisi serta hasil penampung terakhir.
5. Isi jumlah Email pada Latihan dengan COUNTIFS. Untuk biaya Email selesai, gunakan SUMIFS dengan dua pasangan kriteria. Saya memilih rentang yang mempunyai baris awal dan akhir sama.
6. Hitung rata-rata jam Email dan bandingkan dengan hasil **Data > Filter** pada kanal sama. Saya membedakan rata-rata seluruh tiket kanal dari rata-rata tiket selesai saja.
7. Isi hitungan `Peng*`, kemudian filter daftar kategori untuk memeriksa cakupannya. Saya tidak menyimpulkan wildcard hanya memilih satu nama kategori.
8. Isi jumlah Januari menggunakan DATE. Saya menguji batas pada salinan dengan satu tanggal awal bulan berikutnya, kemudian membatalkan perubahan agar bukti awal tetap sesuai.
9. Hitung prioritas menggunakan hitungan syarat pertama ditambah kedua dikurangi irisan keduanya. Bandingkan dengan kolom J, bukan dengan menjumlahkan dua hitungan mentah.
10. Tekan **Ctrl+Alt+F9**, periksa selisih IFS terhadap IF dan kelompok tanpa data, lalu bandingkan dengan solusi. Saya menuliskan penyebut proporsi layanan sebagai tiket selesai.

## 6. Contoh formula dan arti tiap argumen

`=IF(AND(H2="YA";E2<=Master!$B$2);"TEPAT";"TINJAU")` menggabungkan status dan jam. AND menerima dua uji; IF mengubah hasil logika menjadi label. Referensi absolut mempertahankan ambang ketika formula disalin.

`=IF(OR(G2="Urgent";F2>Master!$B$3);"PRIORITAS";"BIASA")` menggunakan gabungan syarat. OR tidak menambahkan jumlah kondisi benar, melainkan menghasilkan satu nilai logika untuk tiket tersebut.

`=IFS(E2<=8;"CEPAT";E2<=Master!$B$2;"SEDANG";E2>Master!$B$2;"LAMA")` mempunyai pasangan kondisi dan label. Alternatifnya `=IF(E2<=8;"CEPAT";IF(E2<=Master!$B$2;"SEDANG";"LAMA"))`. Batas kelas pertama merupakan aturan latihan, bukan standar industri.

`=SUMIFS(Data!F2:F221;Data!C2:C221;"Email";Data!H2:H221;"YA")` menjumlahkan biaya sebagai argumen pertama, lalu memakai pasangan rentang dan nilai kriteria. `=COUNTIFS(Data!D2:D221;"Peng*")` tidak membutuhkan rentang nilai terpisah karena hanya menghitung baris.

Kelompok kosong harus diperiksa sebelum AVERAGEIFS. Saya memakai COUNTIFS untuk menentukan apakah ada pengamatan, lalu mengembalikan pesan TIDAK ADA DATA bila tidak ada. Mengganti kelompok kosong dengan rata-rata nol akan menciptakan angka yang tidak pernah diamati.

## 7. Latihan mandiri minimal lima soal bertingkat

1. Dasar: hitung tiket Email dan biaya Email yang selesai; jelaskan mengapa penyaringnya berbeda.
2. Dasar: jelaskan syarat AND untuk label TEPAT pada satu tiket terbuka dengan jam rendah.
3. Menengah: hitung rata-rata jam Email dan sebutkan populasi baris yang masuk pembagi.
4. Menengah: hitung kategori berawalan Peng dan sebutkan nama kategori yang tercakup.
5. Menengah: hitung tiket Januari dengan batas tanggal dan jelaskan pemakaian batas eksklusif.
6. Lanjut: hitung gabungan prioritas tanpa menggandakan tiket yang memenuhi dua syarat.
7. Lanjut: buktikan IFS dan IF bersarang memberi kelas sama serta jelaskan hasil untuk kelompok tanpa data.

## 8. Kunci jawaban

[Kunci lengkap](latihan/kunci_jawaban.md) memuat ringkasan yang dihitung dari CSV. Untuk gabungan OR, jawaban harus mengurangi irisan agar tiket tidak terhitung dua kali. Untuk AND, seluruh syarat diterapkan bersamaan pada baris yang sama.

Perbedaan kelas IFS dan IF harus nol. Kelompok tanpa pengamatan menghasilkan pesan yang jelas, sedangkan proporsi tepat memakai jumlah tiket selesai sebagai penyebut. Saya tidak membandingkan proporsi dengan ringkasan lain tanpa memeriksa definisi penyebutnya.

## 9. Kesalahan umum dan cara memeriksa hasil

COUNTIFS yang terlalu kecil dapat berasal dari syarat teks berbeda, spasi tersembunyi, atau rentang terpotong. Saya memeriksa nilai melalui Filter dan Formula Bar. SUMIFS yang tampak benar tetap perlu dibandingkan dengan jumlah baris supaya penggandaan nilai tidak luput.

Kesalahan batas terjadi ketika sama dengan seharusnya disertakan tetapi formula memakai lebih kecil saja. Saya menguji nilai tepat pada ambang. IFS yang menempatkan kondisi luas lebih dulu dapat membuat kelompok sempit selalu kosong; saya menelusuri urutan evaluasinya.

## 10. Checklist kompetensi

- [ ] Saya dapat menerjemahkan syarat menjadi IF, AND, dan OR.
- [ ] Saya dapat mengurutkan IFS dan menyediakan alternatif IF bersarang.
- [ ] Saya dapat menjaga keselarasan rentang agregasi.
- [ ] Saya dapat memakai wildcard serta batas tanggal secara tepat.
- [ ] Saya dapat membedakan nol pengamatan dari nilai rata-rata nol.

## 11. Catatan kompatibilitas Excel 2016/2019/365, LibreOffice, dan Google Sheets

Excel 2019/365 mendukung IFS; pada Excel 2016 saya memakai alternatif IF bersarang yang tersedia dalam file. Fungsi agregasi bersyarat inti tersedia pada ketiga versi. Tidak ada dynamic array atau fungsi lookup modern yang dibutuhkan di sini.

LibreOffice dipakai untuk recalc serta [preview nyata](assets/preview.png). Interaksi Excel dan impor Google Sheets belum diuji langsung; saya memeriksa ulang pemisah rumus, tanggal, dan format ketika berpindah aplikasi. Pembangunan ulang memakai `python tools/build.py` dari folder modul, lalu `python tools/course_runtime.py 04_logika-dan-agregasi-bersyarat` dari folder belajar. Saya tidak menyebut fitur manual seperti Slicer sebagai fitur yang sudah tertanam.
