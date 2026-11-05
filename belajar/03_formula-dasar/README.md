# Modul 03: Formula dasar

> Saya memakai data sintetis sebagai materi latihan, bukan pengalaman kerja. Contoh dokumen memakai titik koma untuk Excel regional Indonesia; regional Inggris memakai koma. File menyimpan nama fungsi Inggris dan koma sesuai format XLSX.

## 1. Tujuan dan prasyarat

Saya mempelajari cara membangun perhitungan yang dapat diperiksa dari input sampai total. Target saya adalah menguasai operator, urutan operasi, referensi relatif, absolut, dan campuran, serta SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, COUNTBLANK, dan ROUND. Saya tidak hanya menghafal bentuk fungsi: saya harus dapat menjelaskan sumber nilai, satuan, dan perilakunya ketika rumus disalin.

Prasyaratnya adalah navigasi dasar dan struktur tabel dari modul sebelumnya. Saya sudah dapat memilih rentang, mengenali angka dan teks, serta menjaga hubungan antarkolom saat mengurutkan data. Saya mengerjakan [latihan](latihan/latihan_03.xlsx) pada salinan, lalu membandingkannya dengan [solusi](latihan/solusi_03.xlsx). Saya membaca label sebelum mengisi sel kuning supaya jawaban tidak ditempatkan pada rentang yang berbeda.

## 2. Kegunaan di pekerjaan

Dalam simulasi administrasi jasa, tagihan dihitung dari kuantitas, tarif, potongan, dan pajak. Kesalahan referensi dapat membuat baris pertama benar tetapi baris berikutnya memakai tarif pajak yang bergeser ke sel kosong. Karena tidak selalu muncul pesan kesalahan, pemeriksaan satu contoh saja belum memadai.

Saya menempatkan parameter bersama pada Master dan data baris pada Data. Pemisahan ini membuat perubahan asumsi mudah ditelusuri. Tarif pajak di sini merupakan parameter simulasi, bukan panduan perpajakan. Tujuan latihannya adalah menjaga konsistensi rumus, bukan menentukan aturan hukum suatu transaksi.

Fungsi agregasi dasar membantu pemeriksaan cepat: total menunjukkan besaran keseluruhan, minimum dan maksimum menemukan kandidat pemeriksaan, sedangkan hitungan mengungkap rentang yang terpotong. Saya membaca semuanya bersama konteks agar angka terbesar tidak otomatis dianggap kesalahan.

## 3. Penjelasan dari dasar

Rumus dimulai dengan tanda sama dengan. Excel membaca ekspresi setelah tanda tersebut sebagai perintah perhitungan. Operator perkalian dan pembagian didahulukan dari penjumlahan dan pengurangan; tanda kurung mengubah kelompok yang dihitung lebih dahulu. Ketika menulis potongan harga, saya mengalikan nilai bruto dengan satu dikurangi persentase diskon. Kurung tersebut menyatakan bahwa faktor diskon harus dibentuk terlebih dahulu.

Persentase disimpan sebagai pecahan. Diskon yang terlihat sebagai persen pada layar tidak perlu dibagi seratus lagi bila nilai sel memang sudah berupa pecahan. Saya melihat Formula Bar dan format sel sebelum mengubah formula. Mengalikan persentase dengan seratus kemudian memformatnya sebagai persen akan memperbesar tampilan secara keliru.

Referensi relatif seperti D2 mengikuti perpindahan rumus. Ketika formula D2 dikalikan E2 disalin turun, sumber berubah menjadi D3 dan E3. Inilah perilaku yang diperlukan untuk menghitung tiap tagihan. Referensi absolut seperti `$B$2` mengunci kolom dan baris; saya menggunakannya untuk parameter bersama yang tidak boleh bergeser.

Referensi campuran mengunci salah satu arah. `$B30` mempertahankan kolom B ketika formula disalin ke kanan, tetapi baris dapat berubah ketika disalin turun. `C$29` mempertahankan baris tarif ketika disalin turun, tetapi kolom bergerak saat disalin ke kanan. Keduanya dapat digabung untuk mengisi matriks kuantitas dan tarif tanpa mengetik seluruh hasil.

SUM menjumlahkan angka dalam rentang, sedangkan AVERAGE membagi jumlah dengan banyaknya nilai numerik. Sel kosong tidak dihitung sebagai pengamatan oleh AVERAGE, tetapi angka nol tetap dihitung. Saya tidak mengganti nol dengan kosong hanya untuk menaikkan rata-rata. MIN dan MAX memilih nilai ekstrem; keduanya tidak menjelaskan mengapa nilai tersebut muncul.

COUNT menghitung sel numerik. COUNTA menghitung sel terisi, termasuk teks, sementara COUNTBLANK menghitung kekosongan menurut aturan Excel. Rumus yang menghasilkan teks kosong dapat diperlakukan berbeda oleh COUNTA dan COUNTBLANK, sehingga jumlah keduanya tidak selalu menjadi alat audit sempurna. Dataset ini memakai catatan kosong sungguhan untuk menjaga contoh dasar jelas.

ROUND membulatkan nilai yang digunakan perhitungan berikutnya. Format angka hanya mengubah tampilan. Pada simulasi ini neto dibulatkan per tagihan, lalu pajak dibulatkan dari neto tersebut. Menjumlahkan komponen yang sudah dibulatkan per baris bisa berbeda dari membulatkan satu total di akhir. Saya menetapkan aturan pembulatan terlebih dahulu dan mempertahankannya pada pembanding pandas.

## 4. Dataset latihan

[CSV sumber](latihan/data_03.csv) berisi 180 catatan tagihan jasa sintetis. Kolomnya adalah ID, tanggal, unit layanan, qty, tarif, diskon, dan catatan. Catatan sebagian kosong agar saya dapat membandingkan fungsi hitung. ID bersifat unik dan bukan angka yang perlu dijumlahkan.

Workbook menambahkan kolom bruto, neto, pajak, dan tagihan dengan rumus hidup. Master menyimpan tarif pajak simulasi berwarna biru. Latihan menyediakan agregasi, contoh prioritas operasi, perbandingan referensi, matriks tarif, dan pemeriksaan komponen. Perhitungan detail tersedia pada kedua file; jawaban ringkasan pada latihan sengaja kosong.

Saya tidak menganggap kolom turunan sebagai data mentah pada CSV. Ketika membangun ulang, Python membaca input dan workbook menghitung hasil. [Kunci jawaban](latihan/kunci_jawaban.md) berasal dari perhitungan pandas atas input yang sama. Bukti penghitungan ulang dan rekonsiliasi tersedia pada [VERIFIKASI.md](VERIFIKASI.md).

## 5. Langkah praktik bernomor dengan instruksi klik

1. Buka latihan, baca Instruksi, lalu pilih **Formulas > Calculation Options > Automatic**. Saya menekan **Ctrl+Alt+F9** untuk memastikan hasil tidak tertinggal dari input.
2. Pada Data, pilih H2 dan baca Formula Bar. Tekan **F2** untuk melihat sumber qty dan tarif yang berwarna. Saya membatalkan edit dengan **Esc** setelah menelusuri sumbernya.
3. Pilih I2. Periksa kurung pada faktor diskon dan posisi ROUND. Saya membandingkan rumus ini dengan satu perhitungan manual pada salinan, tanpa mengganti formula menjadi nilai.
4. Pada J2, pilih referensi parameter pajak di Formula Bar dan tekan **F4** untuk melihat pergantian jenis referensi. Saya mengembalikannya menjadi absolut sebelum menyimpan.
5. Pada Latihan, isi C4 sampai C7 dengan SUM untuk komponen yang sesuai. Saya memilih rentang dari baris data pertama sampai terakhir dan tidak memasukkan header.
6. Isi C8 sampai C14 memakai fungsi yang tertulis pada label. Untuk COUNTBLANK, pilih kolom catatan, bukan kolom ID yang selalu terisi. Saya mencatat perbedaan arti hitungan.
7. Isi contoh operator dengan dan tanpa kurung. Gunakan **Formulas > Evaluate Formula** untuk mengikuti urutan perhitungan bila hasil tidak sesuai dugaan.
8. Isi perbandingan referensi relatif dan absolut. Salin formula turun melalui fill handle, lalu klik baris terakhir dan periksa apakah sumber tarif berubah sesuai maksud.
9. Isi matriks mulai C30 dengan `=$B30*C$29`. Salin ke kanan dan ke bawah. Saya memeriksa satu sudut matriks selain sel awal agar kesalahan penguncian dapat terlihat.
10. Isi pemeriksaan komponen pada C35. Selisih total tagihan terhadap neto dan pajak harus nol. Saya membandingkan hasil dengan solusi dan kunci setelah seluruh sel kuning terisi.

## 6. Contoh formula dan arti tiap argumen

`=D2*E2` mengalikan kuantitas dengan tarif per unit. Kedua referensi relatif karena setiap baris mempunyai pasangan inputnya sendiri. `=ROUND(H2*(1-F2);0)` memakai ekspresi neto sebagai argumen pertama dan jumlah digit desimal sebagai argumen kedua. Nilai nol pada argumen digit berarti pembulatan ke rupiah bulat.

`=ROUND(I2*Master!$B$2;0)` mengambil tarif pajak dari sheet lain. Tanda seru memisahkan nama sheet dari alamat; tanda dolar mengunci parameter. Bila sheet mempunyai spasi, nama sheet perlu tanda petik tunggal. Saya tidak mengunci I2 karena neto harus berubah mengikuti baris tagihan.

`=SUM(Data!I2:I181)` menjumlahkan neto. `=AVERAGE(Data!I2:I181)`, `=MIN(Data!I2:I181)`, dan `=MAX(Data!I2:I181)` memakai rentang yang sama untuk pertanyaan berbeda. Saya menjaga cakupan rentang konsisten agar ringkasan dapat dibandingkan.

`=COUNT(Data!I2:I181)` menghitung nilai numerik neto, `=COUNTA(Data!A2:A181)` menghitung ID terisi, dan `=COUNTBLANK(Data!G2:G181)` menghitung catatan kosong. Pada matriks, `=$B30*C$29` menggunakan dua referensi campuran dengan arah penguncian berbeda; saya menjelaskan masing-masing sebelum menyalin formula.

## 7. Latihan mandiri minimal lima soal bertingkat

1. Dasar: hitung total bruto dan neto, lalu jelaskan mengapa keduanya berbeda.
2. Dasar: bandingkan COUNT pada neto dengan COUNTA pada ID serta COUNTBLANK pada catatan.
3. Menengah: hitung rata-rata neto dengan dan tanpa pembulatan nilai, lalu bedakan dari format tampilan tanpa desimal.
4. Menengah: isi contoh operator tanpa kurung dan dengan kurung; jelaskan bagian yang dihitung lebih dahulu.
5. Menengah: salin formula relatif serta absolut dan tunjukkan sumber tarif pada baris terakhir.
6. Lanjut: lengkapi matriks kuantitas dan tarif dengan satu pola formula campuran, lalu periksa sudut kanan bawah.
7. Lanjut: buktikan bahwa total tagihan sama dengan total neto ditambah total pajak dan jelaskan aturan pembulatan yang dipakai.

## 8. Kunci jawaban

[Kunci lengkap](latihan/kunci_jawaban.md) menyediakan nilai agregasi dari CSV, hasil contoh operator, dan pemeriksaan matriks. Saya menilai jawaban dengan dua syarat: angka cocok dan referensi memiliki perilaku yang benar ketika disalin. Mengetik angka yang cocok sekali belum menunjukkan bahwa formula dapat digunakan ulang.

Selisih komponen harus nol dengan aturan pembulatan per baris yang sama. Jika rata-rata hanya diformat tanpa desimal, Formula Bar masih mempertahankan pecahannya; jika memakai ROUND, nilai yang diteruskan berubah. Saya menyebutkan perbedaan ini saat menjelaskan jawaban.

## 9. Kesalahan umum dan cara memeriksa hasil

Kesalahan umum adalah referensi parameter bergeser, tanda kurung diskon hilang, atau rentang agregasi berhenti sebelum baris terakhir. Saya memeriksa formula pada baris awal, tengah, dan akhir. **Formulas > Show Formulas** membantu melihat pola secara serentak, tetapi saya tetap mengembalikan tampilan hasil setelah pemeriksaan.

SUM yang tidak berubah ketika angka baru ditambahkan dapat menunjukkan cakupan rentang belum diperluas. Nilai yang tampak bulat namun menghasilkan selisih kecil bisa berasal dari format tampilan saja. Saya meninjau Formula Bar serta aturan pembulatan sebelum menyimpulkan ada kerusakan data.

Saya tidak membungkus seluruh rumus dengan IFERROR untuk menyembunyikan masalah dasar. Referensi yang salah perlu diperbaiki pada sumbernya. Pesan kesalahan dan selisih rekonsiliasi adalah petunjuk pemeriksaan, bukan gangguan yang harus dihapus dari layar.

## 10. Checklist kompetensi

- [ ] Saya dapat menjelaskan urutan operator dan fungsi tanda kurung.
- [ ] Saya dapat memilih referensi relatif, absolut, dan campuran sesuai arah salin.
- [ ] Saya dapat membedakan fungsi hitung angka, sel terisi, dan sel kosong.
- [ ] Saya dapat menetapkan aturan pembulatan nilai secara konsisten.
- [ ] Saya dapat merekonsiliasi komponen dan menemukan rentang yang terpotong.

## 11. Catatan kompatibilitas Excel 2016/2019/365, LibreOffice, dan Google Sheets

Seluruh fungsi inti pada modul ini tersedia di Excel 2016, 2019, dan 365 tanpa dynamic array. LibreOffice dipakai untuk recalc dan [preview nyata](assets/preview.png); pemisah tampilan dan menu dapat berbeda. Google Sheets mendukung agregasi dasar dan referensi dolar, tetapi impor format serta dialog perlu ditinjau kembali.

Interaksi langsung pada Excel dan Google Sheets belum diuji. Saya membangun ulang dengan `python tools/build.py` dari folder modul dan memverifikasi dengan `python tools/course_runtime.py 03_formula-dasar` dari folder belajar. Saya menyimpan perubahan parameter sebagai skenario terpisah agar hasilnya tidak tercampur dengan bukti awal. Modul ini tidak memerlukan PivotTable, makro, atau fitur optimasi manual.
