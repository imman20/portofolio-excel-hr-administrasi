# Modul 05: Lookup dan penggabungan referensi

> Dataset produk dan pesanan adalah data sintetis untuk latihan. Saya memakai titik koma pada formula dokumen untuk Excel regional Indonesia; regional Inggris memakai koma. Saya tidak menyatakan data ini berasal dari pelanggan atau perusahaan nyata.

## 1. Tujuan dan prasyarat

Saya belajar mengambil atribut produk dari master berdasarkan kode, kemudian membuktikan bahwa hubungan tersebut tidak menggandakan atau menghilangkan pesanan. Saya memakai INDEX/MATCH dan VLOOKUP exact pada workbook, serta mempelajari XLOOKUP sebagai pilihan versi baru. Saya membedakan kunci tidak ditemukan, kunci duplikat, dan kesalahan formula yang seharusnya tetap diselidiki.

Prasyaratnya adalah struktur tabel, referensi absolut, dan logika IF. Saya membuka [latihan](latihan/latihan_05.xlsx), membaca Instruksi, lalu mengisi ringkasan kuning. [Solusi](latihan/solusi_05.xlsx) menyimpan formula detail dan jawaban lengkap. Tujuan saya bukan hanya mendapat harga, melainkan mengetahui apakah harga itu diambil dari baris yang benar.

## 2. Kegunaan di pekerjaan

Simulasi ini memisahkan pesanan dari master produk. Harga cukup dikelola pada satu tabel, sehingga pengetikan ulang pada setiap pesanan dapat dihindari. Namun manfaat tersebut bergantung pada master yang benar: kode harus unik, tipe data sesuai, dan aturan harga berlaku jelas.

Lookup bukan pengganti pemeriksaan hubungan data. Jika master memuat dua harga untuk kode sama, fungsi pencarian biasanya mengambil kecocokan pertama tanpa menyatakan bahwa sumber ambigu. Saya menambahkan pemeriksaan frekuensi kunci sebelum mempercayai hasil. Pendekatan ini relevan untuk daftar barang, akun, pegawai, atau cabang, tetapi kasus di sini tetap latihan sintetis.

## 3. Penjelasan dari dasar

Kunci adalah nilai yang menghubungkan tabel. Pada master satu kode mempunyai satu produk dan satu harga simulasi. Banyak pesanan boleh memakai kode sama; hubungan ini disebut many-to-one dari pesanan ke master. Saya tidak menganggap seluruh tabel harus bebas pengulangan kode karena pengulangan pada sisi pesanan justru wajar.

MATCH dengan argumen terakhir nol mencari kecocokan tepat dan menghasilkan posisi relatif dalam rentang. Posisi tersebut diberikan kepada INDEX untuk mengambil nilai pada rentang harga. Kedua rentang harus mempunyai susunan dan panjang yang sejajar. Menggeser salah satunya dapat menghasilkan harga dari produk lain tanpa pesan kesalahan.

VLOOKUP mencari kunci pada kolom pertama area tabel, kemudian mengambil kolom berdasarkan nomor indeks. Argumen terakhir nol menyatakan exact match. Jika argumen ini dihilangkan, perilaku approximate dapat memberi hasil yang tidak sesuai maksud. Saya tidak memakai approximate untuk identifier produk.

Kelemahan VLOOKUP adalah keharusan kunci berada pada sisi kiri dan penggunaan nomor kolom yang mudah rusak ketika struktur disisipkan atau diubah. INDEX/MATCH memisahkan rentang kunci dari rentang hasil sehingga lebih fleksibel. Fleksibilitas ini tetap memerlukan disiplin menjaga kesejajaran baris.

XLOOKUP dapat mencari pada rentang kunci dan mengembalikan rentang hasil secara langsung. Fungsi ini nyaman pada versi Excel yang mendukungnya, tetapi saya tidak menanamkannya dalam workbook agar jalur Excel 2016 tetap tersedia. Pilihan fungsi modern tidak menghilangkan kebutuhan memeriksa duplikat, tipe data, dan cakupan master.

Kunci yang tidak ditemukan menghasilkan kondisi berbeda dari kerusakan rumus. IFNA menangani hasil tidak tersedia yang lazim muncul pada pencarian exact. IFERROR lebih luas: ia juga dapat menutupi pembagi nol, referensi rusak, atau tipe data keliru. Karena itu saya tidak memasang IFERROR di seluruh workbook hanya agar layar terlihat bersih.

Pada latihan, harga tak dikenal ditampilkan sebagai teks TIDAK ADA. Nilai pesanan yang belum dapat dihitung diberi nol pada kolom nilai_dikenal, disertai status PERIKSA KUNCI. Saya menyebut totalnya sebagai nilai pesanan dengan kode dikenal, bukan nilai seluruh pesanan. Penamaan ini mencegah angka parsial dipakai sebagai total yang seolah lengkap.

Kode dianggap sebagai teks. Spasi, tanda baca, dan nol depan dapat memengaruhi kecocokan. Saya tidak langsung melakukan konversi angka atau pemotongan karakter sebelum mengetahui definisi kunci resmi. Pembersihan yang menggabungkan dua kode berbeda dapat sama berbahayanya dengan kegagalan lookup.

## 4. Dataset latihan

[Pesanan](latihan/data_05.csv) berisi 160 baris sintetis dengan ID pesanan, kode, qty, dan tanggal. [Master](latihan/master_05.csv) berisi 20 produk berkunci unik. Generator sengaja memasukkan tiga kode tidak dikenal sebagai bahan pemeriksaan, bukan membiarkannya menjadi error formula.

Data memuat kolom harga INDEX/MATCH, nilai yang dapat dihitung, status kunci, harga VLOOKUP, dan frekuensi kunci. Sheet Audit_Master berisi contoh duplikat yang terpisah dari master yang dipakai perhitungan. Saya menjaga pemisahan ini agar contoh buruk tidak diam-diam menjadi sumber hasil.

Saya memeriksa semua harga dan nilai per baris terhadap merge pandas dari CSV. [VERIFIKASI.md](VERIFIKASI.md) mencatat hasilnya; [kunci jawaban](latihan/kunci_jawaban.md) menyediakan ringkasan serta penjelasan kondisi lookup.

## 5. Langkah praktik bernomor dengan instruksi klik

1. Buka workbook, aktifkan **Formulas > Calculation Options > Automatic**, lalu lihat Master. Gunakan **Data > Filter** untuk mengenali kolom kunci dan harga, tanpa mengurutkan satu kolom sendirian.
2. Pada Data, pilih E2 dan tekan **F2**. Saya menelusuri kode pesanan, rentang kunci, dan rentang harga yang dipakai INDEX/MATCH.
3. Pilih satu baris berkode dikenal dan gunakan **Formulas > Evaluate Formula**. Perhatikan bahwa MATCH menghasilkan posisi, bukan harga; INDEX memakai posisi tersebut untuk mengambil harga.
4. Bandingkan dengan H pada baris sama. Saya memastikan VLOOKUP memakai argumen exact dan nomor kolom yang benar. Dua metode harus menghasilkan harga sama untuk seluruh pesanan.
5. Filter status PERIKSA KUNCI pada kolom G. Saya melihat kode yang tidak ditemukan dan qty yang belum dapat dinilai, lalu mengembalikan filter sebelum menghitung ringkasan seluruh data.
6. Pada Audit_Master, pilih kode berulang dan periksa COUNTIFS. Saya membedakan jumlah baris berkunci duplikat dari jumlah kelompok kode duplikat.
7. Isi total nilai dikenal, jumlah valid, jumlah perlu pemeriksaan, dan qty tidak dikenal pada Latihan. Saya membaca label total dengan cermat agar hasil parsial tidak diganti namanya menjadi total penuh.
8. Isi pembanding metode serta frekuensi kunci. Jika harga tidak sama, saya memeriksa area VLOOKUP, indeks kolom, dan batas MATCH terlebih dahulu.
9. Isi contoh IFNA dan IFERROR. Contoh pembagi nol sengaja ditangkap IFERROR untuk memperlihatkan luasnya penanganan, bukan untuk dijadikan pola perhitungan bisnis.
10. Tekan **Ctrl+Alt+F9**, bandingkan dengan solusi dan kunci, lalu tulis tindak lanjut untuk kode tak dikenal: pemeriksaan sumber atau master, bukan pembuatan harga rekaan.

## 6. Contoh formula dan arti tiap argumen

`=INDEX(Master!$C$2:$C$21;MATCH(B2;Master!$A$2:$A$21;0))` mengambil harga. INDEX menerima rentang hasil dan posisi baris; MATCH menerima nilai yang dicari, rentang kunci, serta mode exact. Semua rentang master dikunci agar tidak bergeser saat disalin.

`=VLOOKUP(B2;Master!$A$2:$C$21;3;0)` memakai kode B2, area master, kolom harga ketiga, dan exact match. Saya tidak mengganti argumen terakhir dengan perkiraan karena kode bukan interval numerik.

`=IFNA(INDEX(Master!$C$2:$C$21;MATCH(B2;Master!$A$2:$A$21;0));"TIDAK ADA")` memberi pesan ketika kunci tidak tersedia. `=COUNTIFS(Master!$A$2:$A$21;B2)` memeriksa frekuensi: hubungan yang diharapkan memerlukan satu kecocokan.

Untuk Excel yang mendukungnya, alternatif dokumen adalah `=XLOOKUP(B2;Master!A2:A21;Master!C2:C21;"TIDAK ADA";0)`. Rentang hasil terpisah dari rentang kunci; parameter pesan dan mode tetap dinyatakan. Formula modern ini dijelaskan saja, sedangkan file memakai alternatif Excel 2016.

## 7. Latihan mandiri minimal lima soal bertingkat

1. Dasar: ambil harga satu kode dikenal dengan INDEX/MATCH dan jelaskan dua tahapnya.
2. Dasar: bandingkan harga yang sama dengan VLOOKUP exact; sebutkan arti indeks kolom.
3. Menengah: hitung pesanan valid dan kode tidak dikenal, lalu jelaskan batas total nilai_dikenal.
4. Menengah: hitung qty yang belum dapat dinilai tanpa mengarang harga pengganti.
5. Menengah: temukan baris duplikat pada Audit_Master dan jelaskan mengapa mengambil kecocokan pertama belum cukup.
6. Lanjut: jelaskan perbedaan IFNA dan IFERROR menggunakan contoh yang disediakan.
7. Lanjut: terangkan cara mengganti lookup dengan XLOOKUP pada versi yang mendukung, serta alternatif untuk Excel 2016.

## 8. Kunci jawaban

[Kunci lengkap](latihan/kunci_jawaban.md) memuat jumlah dan nilai hasil merge many-to-one. Harga dari dua metode harus cocok. Kode yang belum tersedia tetap menjadi kasus pemeriksaan, sehingga jumlah valid ditambah jumlah perlu pemeriksaan harus sama dengan jumlah pesanan.

Baris audit yang berulang tidak otomatis berarti sebanyak itu kode unik bermasalah. Saya menyatakan apakah hitungan mewakili baris atau kelompok. IFERROR pada contoh menghasilkan pesan walau masalahnya pembagi nol; ini menunjukkan mengapa pemakaiannya perlu dibatasi.

## 9. Kesalahan umum dan cara memeriksa hasil

Harga keliru tanpa error dapat muncul dari approximate match atau indeks kolom salah. Saya memeriksa satu kode dengan harga yang diketahui dari master, lalu membandingkan seluruh baris. Harga benar pada satu baris tidak membuktikan rentang tetap benar setelah disalin.

Kunci tidak cocok dapat disebabkan tipe, spasi, atau master tidak lengkap. Saya menyelidiki penyebabnya sebelum membersihkan. Mengganti semua kunci tak dikenal dengan nol menghilangkan informasi; karena itu pesan dan hitungan kasus tetap ditampilkan.

## 10. Checklist kompetensi

- [ ] Saya dapat menjelaskan hubungan many-to-one dan keunikan master.
- [ ] Saya dapat menulis INDEX/MATCH dan VLOOKUP exact dengan rentang terkunci.
- [ ] Saya dapat membedakan kunci tidak dikenal dengan kunci duplikat.
- [ ] Saya dapat menjelaskan risiko IFERROR yang terlalu luas.
- [ ] Saya dapat menyatakan batas total parsial dan alternatif lintas versi.

## 11. Catatan kompatibilitas Excel 2016/2019/365, LibreOffice, dan Google Sheets

INDEX/MATCH, VLOOKUP, IFNA, dan COUNTIFS tersedia pada Excel 2016/2019/365. XLOOKUP tidak tersedia pada seluruh versi tersebut dan tidak ditanamkan di file. LibreOffice dipakai untuk recalc serta [preview nyata](assets/preview.png); fungsi modern tetap dijelaskan sebagai pilihan terpisah.

Interaksi Excel dan Google Sheets belum diuji langsung. Saya meninjau kembali tipe ID, rentang, dan format setelah impor. Pembangunan ulang memakai `python tools/build.py` dari folder modul, lalu `python tools/course_runtime.py 05_lookup-dan-penggabungan-referensi` dari folder belajar. Tidak ada klaim Power Query atau relasi Data Model sudah tertanam.
