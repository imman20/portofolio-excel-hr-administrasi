# Modul 08 — PivotTable dan PivotChart

> Saya memakai data distribusi sintetis sebagai latihan. Rumus dokumen memakai titik koma untuk Excel regional Indonesia; regional Inggris memakai koma. Workbook menyediakan sumber dan hasil pembanding, sedangkan objek PivotTable, PivotChart, Slicer, serta Timeline harus saya buat manual.

## 1. Tujuan dan prasyarat

Saya belajar merangkum detail transaksi melalui penempatan field pada Rows, Columns, Values, dan Filters. Saya juga mempelajari Show Values As, pengelompokan tanggal, filter interaktif, serta Refresh. Target saya adalah dapat menjelaskan setiap angka pivot dan membuktikan hasilnya sama dengan agregasi sumber.

Prasyaratnya adalah Excel Table, tipe data, SUMIFS, serta tanggal periode. Saya mengunduh [latihan](latihan/latihan_08.xlsx), menyimpan salinan, dan membaca Instruksi. [Solusi](latihan/solusi_08.xlsx) berisi jawaban formula dan sheet `Hasil_Pivot_Diharapkan`; ia tidak berisi PivotTable yang sudah dibuat. Pemisahan ini menjaga klaim fitur tetap sesuai isi file.

## 2. Kegunaan di pekerjaan

PivotTable membantu menjawab pertanyaan ringkasan dengan cepat ketika struktur sumber sudah benar. Saya dapat membandingkan wilayah dan kategori, membaca kontribusi, atau mengubah tampilan menjadi bulanan tanpa menulis ulang setiap rumus. Kecepatan tersebut tetap membutuhkan definisi grain dan agregasi yang jelas.

Dalam simulasi distribusi ini, jumlah transaksi berbeda dari total qty dan nilai rupiah. Menyeret field yang salah ke Values dapat menghasilkan hitungan yang benar secara mekanis tetapi menjawab pertanyaan berbeda. Saya menyebut ukuran beserta satuannya, bukan hanya menyalin angka dari pivot ke laporan.

## 3. Penjelasan dari dasar

PivotTable membaca tabel sumber dan membuat ringkasan menurut kategori yang dipilih. Rows menyusun kategori ke bawah, Columns menyusun kategori mendatar, Values menentukan ukuran yang diagregasi, dan Filters membatasi cakupan keseluruhan. Penempatan field mengubah pertanyaan, bukan sekadar tata letak.

Jika wilayah berada pada Rows dan kategori pada Columns, satu sel memuat agregasi untuk irisan wilayah dan kategori. Grand Total merangkum seluruh irisan yang sedang terlihat sesuai filter. Subtotal muncul ketika beberapa tingkat kategori disusun bertingkat. Saya tidak menambahkan Grand Total kembali ke detail sebagai transaksi baru.

Value Field Settings menentukan Sum, Count, Average, dan agregasi lain. Excel dapat memilih Count jika kolom nilai terbaca sebagai teks atau mengandung tipe campuran. Saya memeriksa sumber dan tipe sebelum memaksa Sum, karena perubahan pengaturan tidak memperbaiki angka yang sebenarnya tersimpan sebagai teks.

Show Values As mengubah cara hasil agregasi disajikan. Persentase Grand Total memakai total keseluruhan yang sedang berlaku sebagai penyebut; persentase Row Total memakai total baris. Kedua persentase dapat berbeda walau nilai rupiahnya sama. Saya menamai ukuran agar penyebutnya terbaca, misalnya Kontribusi grand total.

Tanggal valid dapat dikelompokkan menjadi bulan dan tahun. Jika data mencakup beberapa tahun, memilih Months saja dapat menggabungkan Januari dari tahun berbeda. Saya memilih Years dan Months ketika konteks memerlukannya, lalu memeriksa kembali total. Teks tanggal atau nilai kosong bermasalah dapat menghambat grouping.

Slicer menampilkan pilihan filter sebagai tombol, sedangkan Timeline memfilter rentang waktu. Keduanya memengaruhi cakupan yang terlihat, sehingga pemeriksaan terhadap hasil awal harus dilakukan setelah filter dibersihkan. Filter tersembunyi adalah salah satu penyebab perbedaan angka yang tampak seperti kesalahan formula.

Refresh membaca ulang sumber ke cache pivot. Mengubah data tidak selalu langsung memperbarui PivotTable. Sumber Excel Table membantu perluasan rentang ketika baris ditambahkan, tetapi saya tetap memeriksa apakah kolom formula pada baris baru terisi dan melakukan Refresh setelah perubahan.

PivotChart mengikuti ringkasan pivot dan filter yang terkait. Ini berbeda dari chart biasa yang merujuk rentang sel. Saya membuat PivotChart melalui menu pada PivotTable manual; file pembanding tidak menyamar sebagai PivotChart hanya karena tampilannya mirip.

## 4. Dataset latihan

[CSV](latihan/data_08.csv) berisi 240 transaksi sintetis pada Januari sampai Juni 2024. Kolomnya ID, tanggal, wilayah, kategori, qty, dan harga satuan. Workbook menghitung nilai sebagai qty dikalikan harga satuan. Table bernama `tblDistribusi` mencakup kolom nilai tersebut.

Wilayah adalah Jawa, Sumatra, dan Bali; kategori adalah ATK, Kemasan, dan Peralatan. Sheet Hasil_Pivot_Diharapkan berisi matriks nilai per wilayah dan kategori, kontribusi wilayah, serta ringkasan bulanan nilai, qty, dan jumlah transaksi. Seluruh hasil dibangun dengan rumus biasa.

[Kunci jawaban](latihan/kunci_jawaban.md) memuat angka pembanding dari pandas. [VERIFIKASI.md](VERIFIKASI.md) mencatat recalc dan rekonsiliasi matriks serta bulan. Saya membedakan bukti konsistensi formula dari pengujian interaksi PivotTable manual yang belum dilakukan langsung di Excel.

## 5. Langkah praktik bernomor dengan instruksi klik

1. Pada Data, klik salah satu sel dan periksa **Table Design > Table Name**. Nama sumber harus `tblDistribusi`; range mencakup seluruh transaksi dan kolom nilai.
2. Pilih **Insert > PivotTable > From Table/Range**. Pastikan sumber tabel benar, pilih **New Worksheet**, lalu **OK**. Beri nama sheet manual yang mudah dibedakan dari pembanding.
3. Di **PivotTable Fields**, tarik wilayah ke **Rows**, kategori ke **Columns**, dan nilai ke **Values**. Klik panah field nilai, pilih **Value Field Settings**, lalu **Sum**. Gunakan **Number Format** untuk format ribuan tanpa desimal.
4. Bandingkan setiap irisan dengan Hasil_Pivot_Diharapkan, bukan hanya Grand Total. Jika selisih muncul, periksa filter, tipe nilai, dan sumber yang dipakai sebelum mengubah rumus pembanding.
5. Buat pivot kedua pada sheet lain dengan wilayah pada Rows dan nilai dua kali pada Values. Pada field nilai kedua, pilih **Value Field Settings > Show Values As > % of Grand Total**; ubah nama menjadi Kontribusi grand total.
6. Cocokkan kontribusi wilayah dengan bagian kedua pembanding. Coba **% of Row Total** pada salinan dan jelaskan perubahan penyebut, kemudian kembalikan sebelum rekonsiliasi.
7. Buat pivot waktu dengan tanggal pada Rows dan nilai pada Values. Klik kanan salah satu tanggal, pilih **Group**, centang **Years** dan **Months**, lalu **OK**. Bandingkan nilai per bulan dengan bagian bulanan pembanding.
8. Klik pivot, pilih **PivotTable Analyze > Insert Slicer**, centang wilayah, lalu **OK**. Pilih satu wilayah dan amati perubahan. Gunakan tombol **Clear Filter** pada Slicer sebelum membandingkan hasil seluruh data.
9. Pilih **PivotTable Analyze > Insert Timeline**, centang tanggal, lalu **OK**. Pilih tingkat **Months** dan rentang yang diinginkan. Saya mencatat filter aktif dan membersihkannya sebelum pemeriksaan awal.
10. Pilih **PivotTable Analyze > PivotChart**, gunakan **Clustered Column**, lalu beri judul dan satuan. Saya menghindari grafik tiga dimensi karena perspektif dapat menyulitkan perbandingan.
11. Pada salinan sumber, tambahkan baris percobaan, periksa perluasan Table dan formula nilai, lalu pilih **Data > Refresh All** atau klik kanan pivot **Refresh**. Setelah pengamatan, batalkan perubahan atau tutup salinan tanpa menimpa bukti awal.
12. Isi ringkasan kuning pada Latihan dan cocokkan dengan solusi. Saya menyimpan pivot manual pada salinan terpisah karena menjalankan skrip pembangun akan mengganti workbook bawaan.

## 6. Contoh formula dan arti tiap argumen

`=SUMIFS(Data!$G$2:$G$241;Data!$C$2:$C$241;$A5;Data!$D$2:$D$241;B$4)` adalah pembanding satu sel matriks. Rentang G berisi nilai, C berisi wilayah, dan D berisi kategori. Referensi campuran membuat label wilayah berubah ke bawah dan kategori berubah ke kanan.

`=IF($E$8=0;0;E5/$E$8)` menghitung kontribusi wilayah terhadap grand total pada keadaan sumber awal. Penyebut dikunci. Jika filter pivot berubah, pembanding ini tidak otomatis mengikuti filter manual; saya harus menyamakan cakupan atau menghitung pembanding dengan syarat filter yang sama.

Ringkasan bulanan memakai SUMIFS dengan batas awal inklusif dan awal bulan berikutnya eksklusif. COUNTIFS menghitung catatan, sementara SUM pada qty menghitung unit barang. Saya tidak menukar kedua ukuran hanya karena keduanya berformat angka bulat.

## 7. Latihan mandiri minimal lima soal bertingkat

1. Dasar: buat pivot wilayah dan kategori dengan Sum of nilai serta cocokkan Grand Total.
2. Dasar: ambil satu irisan wilayah dan kategori, lalu tunjukkan transaksi sumber yang membentuknya.
3. Menengah: buat kontribusi wilayah dengan Show Values As dan sebutkan penyebutnya.
4. Menengah: kelompokkan tanggal menurut tahun dan bulan, kemudian cocokkan bulan Januari.
5. Menengah: gunakan Slicer dan Timeline, lalu jelaskan mengapa filter harus dibersihkan saat rekonsiliasi awal.
6. Lanjut: tambahkan baris pada salinan dan buktikan kapan Refresh diperlukan.
7. Lanjut: jelaskan perbedaan PivotChart dengan chart biasa dan bedakan nilai, qty, serta jumlah transaksi.

## 8. Kunci jawaban

[Kunci lengkap](latihan/kunci_jawaban.md) memuat grand total, kontribusi, irisan kategori, dan ringkasan bulanan. Tata letak visual boleh berbeda selama field, agregasi, filter, serta angkanya sama. Saya menerima jawaban yang menyatakan sumber dan penyebut dengan jelas.

Selisih total wilayah terhadap sumber harus nol dan total kontribusi harus utuh pada cakupan awal. Ketika ada filter, hasil dibandingkan dengan sumber yang difilter sama; saya tidak menyimpulkan salah hanya karena berbeda dari grand total tanpa filter.

## 9. Kesalahan umum dan cara memeriksa hasil

Count of nilai sering menunjukkan tipe sumber tidak seluruhnya numerik. Grand Total yang berbeda dapat berasal dari filter aktif atau sumber belum diperluas. Saya memeriksa field list, source range, dan Refresh sebelum menyelidiki hal lain.

Grouping hanya Months dapat menyatukan tahun berbeda. Persentase tanpa nama penyebut dapat disalahartikan. Saya menambahkan label yang jelas dan memeriksa satu kelompok melalui detail sumber, sambil menjaga data sintetis tetap menjadi konteks latihan.

## 10. Checklist kompetensi

- [ ] Saya dapat memilih Rows, Columns, Values, dan Filters sesuai pertanyaan.
- [ ] Saya dapat memilih agregasi serta format nilai yang tepat.
- [ ] Saya dapat menjelaskan Show Values As dan penyebutnya.
- [ ] Saya dapat memakai Group, Slicer, Timeline, serta Refresh secara sadar.
- [ ] Saya dapat merekonsiliasi pivot manual dengan hasil rumus biasa.

## 11. Catatan kompatibilitas Excel 2016/2019/365, LibreOffice, dan Google Sheets

Langkah ditulis untuk Excel desktop Windows; letak menu dapat berbeda menurut versi. Slicer dan Timeline memerlukan dukungan PivotTable dan field tanggal yang sesuai. Workbook bawaan hanya mempunyai Excel Table serta pembanding SUMIFS; objek pivot dan kontrol interaktif belum tertanam.

LibreOffice dipakai untuk recalc dan [preview nyata pembanding](assets/preview.png). Google Sheets mempunyai pivot dan filter sendiri dengan langkah berbeda. Interaksi manual pada aplikasi tersebut belum diuji langsung. Pembangunan ulang memakai `python tools/build.py`, lalu verifikasi dari folder belajar melalui `python tools/course_runtime.py 08_pivottable-dan-pivotchart`. Saya menyimpan hasil klik manual pada salinan agar tidak hilang saat pembangunan ulang.
