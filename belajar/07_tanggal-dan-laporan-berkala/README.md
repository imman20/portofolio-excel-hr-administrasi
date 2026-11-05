# Modul 07 — Tanggal dan laporan berkala

> Saya memakai transaksi dan kalender libur sintetis untuk latihan. Rumus dokumen memakai titik koma untuk Excel regional Indonesia; regional Inggris memakai koma. Kalender ini bukan daftar libur resmi atau aturan kerja perusahaan.

## 1. Tujuan dan prasyarat

Saya belajar membentuk periode laporan yang konsisten dari tanggal transaksi. Fungsi yang dipakai mencakup TODAY, DATE, YEAR, MONTH, DAY, EOMONTH, EDATE, NETWORKDAYS.INTL, DATEDIF, TEXT, dan SUMIFS. Saya menghitung pertumbuhan bulanan serta tahunan sambil menangani pembanding nol secara terbuka.

Prasyaratnya adalah tanggal sebagai nilai numerik, referensi absolut, dan agregasi bersyarat. Saya membuka [latihan](latihan/latihan_07.xlsx), membaca asumsi kalender, lalu mengisi sel kuning. [Solusi](latihan/solusi_07.xlsx) memperlihatkan laporan bulanan lengkap. Saya tidak mengubah tanggal acuan historis menjadi tanggal berjalan hanya agar laporan terlihat terbaru.

## 2. Kegunaan di pekerjaan

Laporan berkala perlu konsisten tentang kapan suatu transaksi masuk periode. Kesalahan satu batas dapat membuat transaksi akhir bulan hilang atau terhitung pada dua bulan. Hal ini sulit terlihat jika hanya memeriksa total tahunan, sehingga saya merekonsiliasi setiap bulan dan membandingkan cakupan periode.

Hari kerja juga berbeda dari hari kalender. Waktu layanan dapat mengecualikan akhir pekan dan libur, tetapi pola libur setiap organisasi tidak selalu sama. Saya menyatakan pola yang dipakai sebelum menghitung, serta membedakan ukuran tanggal inklusif dengan selisih durasi. Semua keputusan di modul ini adalah asumsi simulasi, bukan standar yang berlaku untuk setiap tempat kerja.

## 3. Penjelasan dari dasar

DATE membentuk tanggal dari tahun, bulan, dan hari. YEAR, MONTH, dan DAY mengambil komponennya. Fungsi ini lebih aman untuk menyatakan batas laporan daripada menulis tanggal teks yang dapat ditafsirkan berbeda oleh regional aplikasi. Saya tetap memeriksa tipe sumber setelah impor CSV.

TODAY menghasilkan tanggal saat workbook dihitung ulang. Nilai tersebut sesuai untuk pengingat yang berjalan, tetapi membuat laporan historis berubah jika dipakai sebagai acuan tunggal. Saya menyediakan contoh TODAY tersendiri dan menyimpan tanggal acuan historis sebagai input Master. Keduanya menjawab kebutuhan berbeda dan tidak boleh tertukar tanpa penjelasan.

EOMONTH menghasilkan akhir bulan relatif terhadap tanggal sumber. EDATE menggeser tanggal sejumlah bulan dengan memperhatikan panjang bulan. Menambah jumlah hari tetap bukan cara umum untuk melompat bulan karena panjang bulan berbeda. Tanggal akhir Januari yang digeser ke Februari dapat menyesuaikan ke tanggal terakhir yang tersedia.

Kunci periode `yyyy-mm` adalah label teks untuk membaca dan mengelompokkan bulan. Saya mempertahankan tanggal awal bulan sebagai nilai tanggal terpisah untuk perhitungan batas. Teks periode membantu tampilan, tetapi tidak menggantikan kebutuhan tanggal valid ketika menghitung rentang.

SUMIFS bulanan memakai tanggal lebih besar atau sama dengan awal bulan dan kurang dari awal bulan berikutnya. Pola ini menghindari ketergantungan pada jam akhir hari terakhir. Saya tidak memakai kondisi sampai tanggal terakhir tanpa memikirkan kemungkinan sumber kelak menyimpan waktu.

NETWORKDAYS.INTL menghitung hari kerja dengan pola akhir pekan yang dapat ditentukan. Mask terdiri dari urutan Senin sampai Minggu; nol berarti hari kerja dan satu berarti bukan hari kerja. Argumen libur menambahkan tanggal yang dikecualikan. Tanggal libur yang sudah jatuh pada akhir pekan tidak dikurangkan dua kali.

Fungsi hari kerja menghitung tanggal awal dan akhir secara inklusif bila keduanya merupakan hari kerja. Sebaliknya, pengurangan tanggal biasa menyatakan jarak hari di antara dua titik. DATEDIF menyediakan selisih dalam unit tertentu, termasuk bulan penuh. Saya memilih unit sesuai pertanyaan dan tidak menafsirkan bulan penuh sebagai hasil membagi hari dengan angka tetap.

Pertumbuhan MoM membandingkan nilai bulan berjalan dengan bulan sebelumnya. YoY membandingkan bulan yang sama tahun sebelumnya. Keduanya memakai selisih dibagi nilai pembanding. Jika pembanding nol, persentase tidak terdefinisi menurut rumus ini; label PEMBANDING NOL lebih jujur daripada memberi nol persen atau angka sangat besar buatan.

Periode pertama yang belum memiliki pembanding juga berbeda dari pembanding bernilai nol. Laporan menampilkan BELUM ADA PEMBANDING untuk keadaan pertama. Saya mempertahankan perbedaan tersebut agar pembaca tidak mengira sejarah yang belum tersedia merupakan penurunan atau ketiadaan aktivitas.

## 4. Dataset latihan

[CSV transaksi](latihan/data_07.csv) berisi 480 catatan pada dua tahun, 2023 dan 2024. Kolomnya ID, tanggal, kanal, dan nilai. Februari pada tahun pertama sengaja memiliki nilai nol untuk menguji pembagi nol, sementara catatan tetap tersedia sebagai bagian simulasi kalender.

[CSV libur](latihan/libur_07.csv) berisi tanggal pengecualian simulasi. Sheet Laporan_Bulanan memuat awal bulan, label periode, nilai, MoM, YoY, hari kerja, komponen tanggal, dan jumlah hari kalender. Master menyimpan tanggal acuan akhir tahun agar laporan historis tetap dapat direproduksi.

[Kunci jawaban](latihan/kunci_jawaban.md) dibangun dari agregasi pandas dan kalender hari kerja. [VERIFIKASI.md](VERIFIKASI.md) mencocokkan nilai bulanan serta pertumbuhan yang dapat dihitung, termasuk label ketika pembanding tidak tersedia atau nol.

## 5. Langkah praktik bernomor dengan instruksi klik

1. Buka latihan dan periksa tanggal pada Data melalui Formula Bar. Pilih **Ctrl+1 > Custom** bila perlu menampilkan `dd/mm/yyyy`, tanpa mengetik ulang nilai sumber.
2. Buka Master dan bedakan tanggal acuan dengan contoh TODAY pada Latihan. Saya mencatat bahwa hasil TODAY dapat berubah pada penghitungan ulang berikutnya.
3. Pada Laporan_Bulanan, pilih A2 lalu A3. Gunakan **Formulas > Evaluate Formula** untuk membaca DATE dan EDATE yang membentuk kalender bulanan.
4. Pilih B2 dan periksa TEXT. Saya membandingkan label periode dengan tanggal awal bulan agar tidak mengira kedua kolom mempunyai tipe sama.
5. Pilih C2 dan telusuri SUMIFS. Pastikan batas akhir memakai awal bulan berikutnya secara eksklusif. Saya memeriksa beberapa tanggal di perbatasan bulan melalui **Data > Filter**.
6. Buka Libur dan baca daftar tanggal simulasi. Pilih F pada laporan, kemudian periksa mask akhir pekan serta rentang tanggal libur yang dikunci.
7. Isi total tahunan dan pertumbuhan pada Latihan. Saya memakai nilai tahun pembanding sebagai penyebut, bukan nilai tahun berjalan.
8. Isi contoh EOMONTH dan EDATE. Bandingkan tanggal akhir Februari kabisat dengan hasil pergeseran dari akhir Januari; saya menjelaskan mengapa keduanya mengikuti kalender.
9. Baca MoM dan YoY ketika pembanding nol. Saya tidak mengganti label menjadi angka untuk memaksa grafik atau rata-rata persentase.
10. Tekan **Ctrl+Alt+F9**, cocokkan setiap bulan dengan kunci, dan pastikan jumlah bulan kembali ke total tahunan. Saya menyimpan skenario kalender berbeda pada salinan terpisah.

## 6. Contoh formula dan arti tiap argumen

`=DATE(2024;2;1)` menyatakan awal Februari secara eksplisit. `=EOMONTH(A2;0)` mengambil akhir bulan yang sama; argumen terakhir menunjukkan pergeseran bulan. `=EDATE(A2;1)` menggeser tanggal satu bulan dan menyesuaikan tanggal yang tidak tersedia.

`=TEXT(A2;"yyyy-mm")` mengubah tanggal menjadi label tahun dan bulan. YEAR, MONTH, dan DAY mengambil komponen numerik tanpa perlu memotong teks tampilan. Saya tidak memakai LEFT pada tanggal numerik untuk mengambil tahun.

`=NETWORKDAYS.INTL(A2;EOMONTH(A2;0);"0000011";Libur!$A$2:$A$7)` menerima awal, akhir, mask akhir pekan Sabtu–Minggu, dan rentang libur. Pada contoh kedua mask diubah untuk Jumat–Sabtu. Saya memeriksa urutan karakter mask sebelum menafsirkan hasil.

`=IF(C2=0;"PEMBANDING NOL";(C3-C2)/C2)` menjaga pembagi nol. Formula tidak membuktikan sebab pertumbuhan; ia hanya menghitung perubahan relatif. `=DATEDIF(A2;EDATE(A2;1);"d")` menghitung jarak hari dari awal bulan ke awal bulan berikutnya, sehingga mencerminkan panjang bulan yang sebenarnya.

## 7. Latihan mandiri minimal lima soal bertingkat

1. Dasar: ambil tahun, bulan, dan hari dari tanggal acuan tanpa memotong teks.
2. Dasar: bentuk akhir Februari serta pergeseran akhir Januari dengan fungsi kalender.
3. Menengah: hitung total tahunan dan buktikan jumlah bulannya cocok.
4. Menengah: hitung hari kerja Januari dengan libur simulasi dan jelaskan sifat inklusifnya.
5. Menengah: ubah mask akhir pekan pada salinan, lalu bandingkan hari kerja yang diperoleh.
6. Lanjut: jelaskan MoM dan YoY ketika pembanding nol atau belum tersedia.
7. Lanjut: jelaskan kapan TODAY tepat digunakan dan kapan acuan historis tetap diperlukan.

## 8. Kunci jawaban

[Kunci lengkap](latihan/kunci_jawaban.md) memuat hasil agregasi dan contoh tanggal. Jawaban harus mengikuti kalender, mask, serta daftar libur yang sama. Hasil hari kerja berbeda belum berarti salah jika asumsi akhir pekannya memang berbeda; asumsi tersebut perlu dinyatakan.

Pertumbuhan dengan pembanding nol tetap berupa label. Saya membandingkan angka pertumbuhan hanya ketika penyebut tersedia dan bukan nol. Hasil TODAY mengikuti tanggal penghitungan ulang, sehingga bukan angka historis yang dipertahankan selamanya.

## 9. Kesalahan umum dan cara memeriksa hasil

Menambah hari tetap untuk berpindah bulan dapat melewati tanggal yang dimaksud. Mask akhir pekan terbalik dapat mengubah hari kerja menjadi hari libur. Saya memeriksa kalender kecil secara manual sebelum menggunakan hasil pada rentang panjang.

Tanggal teks dapat terlihat sama dengan tanggal valid tetapi gagal pada perbandingan numerik. Saya menguji tipe melalui COUNT dan membaca nilai sebenarnya. Saya juga memeriksa cakupan rentang SUMIFS supaya transaksi terakhir tidak terlewat.

## 10. Checklist kompetensi

- [ ] Saya dapat membentuk dan memecah tanggal melalui fungsi kalender.
- [ ] Saya dapat memisahkan tanggal berjalan dari acuan historis.
- [ ] Saya dapat menghitung hari kerja dengan mask dan libur yang dinyatakan.
- [ ] Saya dapat menyusun agregasi bulanan dengan batas yang konsisten.
- [ ] Saya dapat menjelaskan MoM/YoY serta pembanding nol.

## 11. Catatan kompatibilitas Excel 2016/2019/365, LibreOffice, dan Google Sheets

Fungsi kalender utama tersedia di Excel 2016/2019/365. DATEDIF digunakan dengan unit hari atau bulan penuh; saya menghindari unit gabungan yang dapat menimbulkan interpretasi tidak jelas. Tidak ada dynamic array dalam file.

LibreOffice dipakai untuk recalc dan [preview nyata](assets/preview.png), dengan zona waktu pengujian Asia/Jakarta. Interaksi Excel serta Google Sheets belum diuji langsung; tanggal berjalan dan regional perlu diperiksa ulang setelah impor. Pembangunan ulang memakai `python tools/build.py`, lalu `python tools/course_runtime.py 07_tanggal-dan-laporan-berkala` dari folder belajar. Saya tidak mengklaim Timeline sudah tertanam pada workbook ini.
