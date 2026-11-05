# Modul 06 — Cleaning dan validasi

> Saya memakai data sintetis untuk melatih pembersihan yang dapat ditelusuri. Tidak ada identitas pribadi nyata. Rumus dokumen memakai titik koma pada Excel regional Indonesia; regional Inggris memakai koma.

## 1. Tujuan dan prasyarat

Saya belajar membersihkan teks, angka, tanggal, dan kode tanpa menghapus jejak sumber. Fungsi utamanya TRIM, CLEAN, SUBSTITUTE, LEFT, RIGHT, MID, LEN, VALUE, DATE, dan DATEVALUE. Saya membahas TEXTSPLIT dengan alternatif Text to Columns, missing values, duplikasi, outlier IQR, Data Validation, serta Conditional Formatting.

Prasyarat saya adalah struktur data, logika, dan lookup. Saya dapat membedakan tipe data dari tampilan serta memahami bahwa satu ID berulang belum otomatis merupakan kesalahan. Saya membuka [latihan](latihan/latihan_06.xlsx), mempertahankan kolom mentah, lalu mengisi ringkasan kuning. [Solusi](latihan/solusi_06.xlsx) menjadi pembanding formula dan keputusan pembersihan.

## 2. Kegunaan di pekerjaan

Ekspor administrasi dapat membawa spasi tersembunyi, angka sebagai teks, atau tanggal dengan format campuran. Jika langsung dijumlahkan atau digabung dengan master, hasil dapat kurang lengkap tanpa tanda yang jelas. Saya mengidentifikasi jenis masalah terlebih dahulu, menetapkan aturan, lalu menghitung dampaknya terhadap ringkasan.

Cleaning yang baik tidak hanya menghasilkan tabel rapi. Orang lain harus dapat melihat apa yang berubah, mengapa perubahan dilakukan, dan apakah ada nilai yang sengaja belum diputuskan. Dalam simulasi ini saya memakai penanda KOSONG dan TINJAU agar kekurangan data tidak menghilang dari laporan. Saya tidak menyebut hasil pembersihan sebagai data benar hanya karena semua error formula sudah hilang.

## 3. Penjelasan dari dasar

TRIM mengurangi spasi biasa di awal dan akhir serta rangkaian spasi di tengah. CLEAN membuang sebagian karakter kontrol yang dapat muncul dari penyalinan. Keduanya tidak membersihkan seluruh karakter Unicode. Karena dataset memuat nonbreaking space, saya menggantinya dengan spasi biasa melalui SUBSTITUTE sebelum menjalankan TRIM.

SUBSTITUTE menerima teks, karakter lama, dan karakter pengganti. Saya mengganti karakter yang sudah diketahui, bukan menghapus seluruh tanda baca secara membabi buta. Tanda hubung pada kode dapat menjadi bagian struktur yang penting; menghilangkannya sebelum pemisahan dapat menghapus informasi batas segmen.

LEFT dan RIGHT mengambil karakter dari sisi kiri atau kanan. MID mengambil karakter dari posisi tertentu, sedangkan LEN menghitung panjang teks. Pada kode berpola tetap, fungsi tersebut cukup untuk memisahkan kota dan nomor. Jika panjang atau pola berubah, saya perlu memeriksa panjang serta pemisah terlebih dahulu; formula posisi tetap bukan parser universal.

VALUE mengubah teks numerik menjadi angka ketika pola teks dapat dikenali pengaturan aplikasi. Penanda N/A, tanda hubung, dan sel kosong tidak saya paksa menjadi nol. Nol adalah nilai terukur, sementara missing berarti nilai tidak tersedia. Mengubah keduanya menjadi nol dapat menurunkan mean dan menciptakan kesimpulan palsu.

Tanggal campuran memerlukan aturan eksplisit. Saya membedakan format ISO dengan posisi tanda hubung, lalu membangun tanggal melalui DATE dari bagian tahun, bulan, dan hari. DATEVALUE berguna untuk teks tanggal yang dikenali, tetapi pengenalan bergantung regional. Saya menyediakan contoh ISO terpisah dan tidak mengandalkannya untuk menebak seluruh format campuran.

Duplikasi harus dinilai berdasarkan kunci dan grain. Dataset memuat ID berulang sebagai masalah yang harus ditinjau. Saya menghitung frekuensi setiap ID dan menandai baris berulang; saya tidak langsung membuang baris karena dua catatan dapat mengandung informasi berbeda. Kebijakan memilih baris pertama, terakhir, atau menggabungkan harus disepakati berdasarkan konteks.

Outlier IQR memakai kuartil pertama dan ketiga dari nilai numerik yang tersedia. Rentang interkuartil adalah selisih keduanya; batas tinjauan berada pada satu setengah IQR di bawah dan di atas kuartil. Nilai di luar batas bukan otomatis salah input. Saya menandainya, memeriksa bukti, dan mempertahankannya sampai ada alasan koreksi yang dapat dipertanggungjawabkan.

Data Validation mengendalikan ketikan baru, sedangkan Conditional Formatting membantu melihat kondisi. Keduanya bukan pengganti pemeriksaan data yang sudah ada. Penempelan atau impor dapat melewati validasi, dan warna tidak mengubah nilai. Saya tetap menyediakan kolom flag serta hitungan yang dapat diperiksa tanpa mengandalkan warna layar.

## 4. Dataset latihan

[CSV](latihan/data_06.csv) memuat 100 catatan sintetis. Kolom mentahnya ID, nama, email, nilai, tanggal, dan kode. Masalah yang sengaja dimasukkan mencakup spasi nonbreaking, tab pada email, penanda kosong, dua format tanggal, ID berulang, serta nilai ekstrem.

Workbook mempertahankan seluruh kolom mentah dan menambahkan nama, email, nilai, tanggal, segmen kode, frekuensi ID, dan flag IQR. Sheet Input_Validasi menyediakan area percobaan aturan qty dan kategori. Kuartil serta batas IQR dihitung pada kedua file agar flag detail tidak bergantung pada jawaban ringkasan yang masih kosong.

[Kunci jawaban](latihan/kunci_jawaban.md) berasal dari parsing pandas dengan penanda kosong dipertahankan. [VERIFIKASI.md](VERIFIKASI.md) membandingkan setiap nama, tanggal, dan nilai bersih dengan workbook, bukan hanya total yang dapat saling menutupi kesalahan.

## 5. Langkah praktik bernomor dengan instruksi klik

1. Buka Data dan bandingkan kolom mentah dengan kolom bersih. Saya membaca Formula Bar agar karakter yang tidak terlihat tidak dinilai dari tampilan saja.
2. Pilih G2 dan tekan **F2**. Telusuri SUBSTITUTE, CLEAN, dan TRIM dari bagian terdalam ke luar. Saya memastikan karakter nonbreaking diganti sebelum spasi dirapikan.
3. Pada H, filter KOSONG. Saya membedakan penanda yang tidak dapat dipakai sebagai angka dari nilai numerik nol. Seluruh baris sumber tetap dipertahankan.
4. Pilih I2 dan I3, lalu gunakan **Formulas > Evaluate Formula**. Periksa pola tanggal yang berbeda dan bagian tahun, bulan, serta hari yang diberikan ke DATE.
5. Pilih J sampai M untuk melihat LEFT, RIGHT, MID, dan LEN. Cocokkan segmen hasil dengan kode mentah tanpa mengubah nomor menjadi bilangan jika nol depannya bermakna.
6. Pada salinan kolom kode, gunakan **Data > Text to Columns > Delimited**, pilih tanda hubung sebagai delimiter, lalu tentukan **Destination** di area kosong. Saya tidak menimpa kolom sumber.
7. Filter frekuensi ID lebih besar dari satu. Catat baris yang terlibat dan bedakan jumlah baris berulang dari jumlah ID unik yang berulang.
8. Baca kuartil dan ambang IQR pada Latihan, lalu filter flag TINJAU pada Data. Saya menulis alasan pemeriksaan sebelum memutuskan koreksi atau penghapusan.
9. Pada Input_Validasi, pilih B2 dan buka **Data > Data Validation** untuk melihat batas integer. Coba ketikan di luar batas pada salinan, lalu batalkan. Periksa daftar kategori pada kolom C.
10. Buka **Home > Conditional Formatting > Manage Rules** dan lihat aturan duplikat. Isi ringkasan kuning, tekan **Ctrl+Alt+F9**, lalu cocokkan hasil dengan solusi.

## 6. Contoh formula dan arti tiap argumen

`=TRIM(CLEAN(SUBSTITUTE(B2;" ";" ")))` menerima teks sumber, mengganti nonbreaking space dengan spasi biasa, membuang karakter kontrol yang dikenali, lalu merapikan spasi. Karakter pada argumen kedua SUBSTITUTE terlihat mirip spasi tetapi berbeda kode.

`=LEFT(F2;3)`, `=RIGHT(F2;3)`, dan `=MID(F2;5;3)` memisahkan kode berpola tetap. Argumen MID adalah teks, posisi awal, dan banyak karakter. `=LEN(F2)` membantu memeriksa apakah panjangnya sesuai pola sebelum pemotongan dipercaya.

`=QUARTILE.INC(Data!H2:H101;1)` memilih kuartil pertama dari nilai numerik; teks penanda kosong tidak diperlakukan sebagai nol. Kuartil ketiga memakai argumen terakhir tiga. Batas bawah dan atas dibentuk dari kuartil serta IQR, kemudian dibandingkan dengan nilai setiap baris.

Pada Excel yang mendukungnya, `=TEXTSPLIT(F2;"-")` dapat membagi kode berdasarkan delimiter. Fungsi ini hanya dijelaskan di dokumen dan tidak berada dalam file. Alternatif Excel 2016 adalah Text to Columns pada salinan atau kombinasi fungsi teks untuk pola yang sudah divalidasi.

## 7. Latihan mandiri minimal lima soal bertingkat

1. Dasar: bersihkan satu nama dan jelaskan mengapa TRIM saja tidak cukup untuk nonbreaking space.
2. Dasar: pisahkan kota dan nomor kode sambil mempertahankan nol di depan.
3. Menengah: hitung nilai missing dan bedakan dari jumlah nilai numerik yang tersedia.
4. Menengah: jelaskan parsing dua format tanggal dan risiko DATEVALUE pada regional berbeda.
5. Menengah: tandai ID berulang tanpa langsung menghapus baris; bedakan baris dengan kelompok ID.
6. Lanjut: hitung kuartil, IQR, serta kandidat outlier dan jelaskan keputusan yang belum dapat diambil dari flag saja.
7. Lanjut: jelaskan batas Data Validation dan perbedaan fungsi warna dengan kolom pemeriksaan.

## 8. Kunci jawaban

[Kunci lengkap](latihan/kunci_jawaban.md) memuat hasil hitungan, total, mean valid, kuartil, dan jumlah ID unik. Saya menerima hasil cleaning bila teks, angka, dan tanggal cocok dengan aturan yang dinyatakan, bukan hanya terlihat seragam.

Outlier dipertahankan sebagai kandidat tinjauan. ID berulang juga tetap ada sampai kebijakan deduplikasi ditetapkan. Penanda KOSONG dikecualikan dari mean numerik; penyebutnya harus sesuai jumlah nilai yang tersedia.

## 9. Kesalahan umum dan cara memeriksa hasil

Menghapus semua spasi dapat menyatukan kata nama. Memaksa penanda missing menjadi nol mengubah statistik. Menebak tanggal dari tampilan dapat menukar hari dan bulan. Saya membandingkan hasil per baris dengan sumber dan menguji dua pola tanggal secara terpisah.

Remove Duplicates langsung pada sumber dapat menghilangkan bukti sebelum ditinjau. Saya membuat salinan dan memakai flag terlebih dahulu. Format warna membantu fokus, tetapi saya memeriksa nilai flag melalui Filter agar temuan tetap dapat direproduksi.

## 10. Checklist kompetensi

- [ ] Saya dapat merapikan teks tanpa menghapus informasi penting.
- [ ] Saya dapat memisahkan missing dari nol dan membaca penyebut statistik.
- [ ] Saya dapat mengonversi tanggal dengan aturan yang eksplisit.
- [ ] Saya dapat menandai duplikat dan outlier tanpa keputusan otomatis yang tidak didukung.
- [ ] Saya dapat menerapkan validasi serta menjelaskan batasnya.

## 11. Catatan kompatibilitas Excel 2016/2019/365, LibreOffice, dan Google Sheets

Fungsi inti tersedia di Excel 2016/2019/365; TEXTSPLIT memerlukan versi yang mendukungnya dan tidak ditanamkan pada file. Data Validation dan Conditional Formatting dibuat dengan openpyxl, tetapi tampilan dialog dapat berbeda di aplikasi lain.

LibreOffice dipakai untuk recalc dan [preview nyata](assets/preview.png). Interaksi Excel serta Google Sheets belum diuji langsung; saya meninjau parsing tanggal dan karakter Unicode setelah impor. Pembangunan ulang memakai `python tools/build.py`, lalu verifikasi dari folder belajar melalui `python tools/course_runtime.py 06_cleaning-dan-validasi`. Saya mempertahankan sumber agar setiap perubahan dapat ditelusuri.
