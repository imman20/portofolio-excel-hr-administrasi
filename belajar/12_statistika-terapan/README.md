# Modul 12: Statistika terapan

> Saya memakai data skor dan deret unit sintetis untuk latihan, bukan hasil penelitian atau eksperimen nyata. Formula dokumen memakai titik koma untuk Excel regional Indonesia; regional Inggris memakai koma. Angka inferensial merupakan ilustrasi bersyarat pada asumsi, bukan bukti dampak nyata.

## 1. Tujuan dan prasyarat

Saya belajar membaca statistik deskriptif, histogram, interval kepercayaan, korelasi, regresi, uji t, ukuran efek, dan prediksi linear. Fungsi utama adalah CONFIDENCE.T, CORREL, LINEST, T.TEST, FORECAST.LINEAR, dan TREND, disertai agregasi serta moving average. Target saya adalah menjelaskan arti angka, desain pembanding, asumsi, dan batas kesimpulan sebelum memberi rekomendasi.

Prasyaratnya adalah agregasi, referensi, serta visualisasi. Saya mengerjakan [latihan](latihan/latihan_12.xlsx), memeriksa sumber, dan membandingkan dengan [solusi](latihan/solusi_12.xlsx). Saya tidak memilih uji hanya karena menghasilkan p-value kecil. Definisi kelompok, pasangan, dan arah pengujian harus ditentukan dari pertanyaan serta struktur pengamatan.

## 2. Kegunaan di pekerjaan

Statistik membantu merangkum variasi, mengukur ketidakpastian, dan membandingkan kelompok. Mean tanpa distribusi dapat menyembunyikan nilai ekstrem, sedangkan perbedaan mean tanpa ketidakpastian dapat terlalu mudah dianggap penting. Saya membaca ukuran pusat, sebaran, dan cakupan data bersama-sama.

Dalam simulasi ini, skor awal dan akhir merujuk ID sama, sementara kelompok A dan B berisi pengamatan independen yang berbeda. Perbedaan desain menentukan uji yang dipakai. Saya tidak menyebut korelasi jam dan skor sebagai efek pelatihan, karena skor juga dibentuk oleh faktor lain dalam generator dan tidak mewakili penugasan kausal nyata.

## 3. Penjelasan dari dasar

Mean adalah rata-rata aritmetika, median adalah titik tengah urutan, dan simpangan baku sampel merangkum sebaran dengan pembagi n dikurangi satu. Varians mempunyai satuan kuadrat, sedangkan simpangan baku kembali pada satuan skor. MIN dan MAX membantu membaca rentang, tetapi tidak menentukan apakah nilai ekstrem salah.

Histogram membagi nilai menjadi interval dan menghitung frekuensi. Saya memakai interval sama lebar, batas bawah inklusif, batas atas eksklusif, kecuali interval terakhir yang menyertakan batas tertinggi. Aturan tersebut mencegah nilai pada batas terhitung dua kali. Bentuk histogram dipengaruhi lebar bin, sehingga saya tidak menganggap satu tampilan sebagai satu-satunya gambaran distribusi.

CONFIDENCE.T menghasilkan margin interval mean berdasarkan distribusi t, simpangan baku sampel, dan jumlah pengamatan. Batas interval adalah mean ditambah atau dikurangi margin tersebut. Interpretasi frequentist mengacu pada prosedur: dalam pengulangan sampel dengan asumsi yang sesuai, proporsi interval yang dibentuk akan mencakup mean populasi pada tingkat yang dinyatakan. Saya tidak mengatakan probabilitas parameter tetap berada dalam satu interval ini adalah tingkat kepercayaan tersebut.

CI mengasumsikan pengamatan independen dan model sampling yang memadai; pendekatan t mengandalkan normalitas mean atau ukuran sampel yang mendukung pendekatan tersebut. CI tidak mengoreksi sampel non-acak, bias seleksi, kesalahan pengukuran, atau data hilang yang tidak acak. Pada data sintetis ini, CI adalah latihan mekanisme, bukan generalisasi ke populasi nyata.

CORREL mengukur hubungan linear antara dua variabel numerik. Nilai dekat nol tidak membuktikan tidak ada hubungan nonlinier. Outlier dapat memengaruhi korelasi, dan hubungan statistik tidak membuktikan sebab-akibat. Saya memeriksa scatter plot serta definisi pasangan sebelum membaca koefisien.

LINEST mengestimasi regresi linear. Pada satu prediktor, keluaran memuat slope, intercept, serta statistik tambahan ketika diminta. Saya mengambil unsur tertentu melalui INDEX agar sel tetap berupa formula biasa tanpa dynamic array. Slope menyatakan perubahan prediksi rata-rata skor untuk perubahan satu unit prediktor dalam model; ia bukan otomatis efek sebab-akibat.

Regresi memerlukan bentuk hubungan yang sesuai. Untuk inferensi klasik diperlukan asumsi galat yang relevan, termasuk independensi dan homogenitas varians, sedangkan normalitas galat berkaitan dengan inferensi tertentu. Saya meninjau residual serta rentang prediktor, tidak hanya R kuadrat. R kuadrat tinggi tidak menjamin model benar atau prediksi di luar sampel akurat.

T.TEST mempunyai argumen ekor dan tipe. Dua ekor menguji perbedaan tanpa arah yang dipilih sesudah melihat hasil. Tipe satu dipakai untuk pasangan pada unit sama; tipe dua untuk kelompok independen dengan asumsi varians sama; tipe tiga adalah Welch yang mengizinkan varians berbeda. Welch tetap membutuhkan pengamatan independen dan tidak memperbaiki confounding atau bias desain.

Saya memakai Welch sebagai contoh utama kelompok A terhadap B, serta paired untuk awal terhadap akhir pada ID sama. Uji equal variance disediakan sebagai pembanding pembelajaran. P-value menyatakan seberapa tidak selaras data dengan model nol menurut statistik uji, bukan probabilitas hipotesis nol benar atau ukuran besarnya efek. Banyak uji eksploratif juga memerlukan pertimbangan pengujian berganda.

Cohen’s d pada latihan adalah selisih mean B dikurangi A dibagi simpangan baku pooled. Ini ukuran efek terstandar kelompok independen, bukan ukuran efek paired. Besaran praktis perlu dibaca bersama skala skor, konteks, dan ketidakpastian; ambang kategori umum tidak menggantikan pertimbangan substantif.

FORECAST.LINEAR dan TREND memberi prediksi dari hubungan linear yang sama pada contoh satu prediktor. Moving average merangkum beberapa observasi terakhir. Keduanya bukan bukti bahwa pola akan berlanjut. Saya membedakan smoothing yang memakai observasi berjalan dari pengujian ramalan pada periode masa depan yang belum digunakan saat pemasangan model.

## 4. Dataset latihan

[CSV skor](latihan/data_12.csv) memuat 60 ID sintetis, kelompok, jam latihan, skor awal, dan skor akhir. Kelompok tersusun agar dua rentang kelompok jelas; pasangan awal dan akhir tetap pada baris ID yang sama. Skor dibatasi pada rentang simulasi dan tidak merupakan penilaian orang nyata.

[CSV tren](latihan/tren_12.csv) memuat 24 periode bulanan dengan indeks dan unit. Sheet Tren menghitung moving average tiga observasi, sedangkan Histogram menghitung frekuensi skor dan mempunyai chart native pada solusi. Latihan menyediakan statistik serta perbandingan metode.

[Kunci jawaban](latihan/kunci_jawaban.md) dihitung dengan pandas/scipy. [VERIFIKASI.md](VERIFIKASI.md) mencocokkan statistik, p-value, interval, prediksi, bin histogram, dan moving average. Deret unit sengaja memuat tren serta komponen musiman; garis linear dan moving average belum memodelkan musiman secara lengkap. Kesetaraan perhitungan tidak menjadi bukti validitas asumsi atau desain kausal.

## 5. Langkah praktik bernomor dengan instruksi klik

1. Buka Data dan baca arti setiap kolom. Saya memastikan kelompok A dan B berbeda ID, sedangkan skor awal dan akhir benar-benar berpasangan.
2. Isi n, mean, median, simpangan baku, varians, minimum, dan maksimum. Gunakan **Formulas > Trace Precedents** untuk memeriksa cakupan rentang.
3. Buka Histogram dan periksa batas bin serta frekuensinya. Jika membuat chart manual, pilih label dan frekuensi lalu **Insert > Column Chart**; saya menjaga interval sama lebar dan sumbu frekuensi mulai nol.
4. Isi margin CI dan kedua batasnya. Saya mencatat tingkat kepercayaan serta asumsi independensi, bukan hanya menyalin angka interval.
5. Pilih kolom jam dan skor akhir pada salinan, lalu **Insert > Scatter**. Amati bentuk hubungan sebelum membaca CORREL.
6. Isi slope, intercept, dan R kuadrat melalui INDEX atas LINEST. Gunakan **Formulas > Evaluate Formula** untuk menelusuri keluaran yang diambil.
7. Hitung T.TEST Welch dua sisi untuk kelompok dan paired untuk awal/akhir. Saya memeriksa argumen tipe, ekor, serta pasangan data sebelum menafsirkan p-value.
8. Hitung Cohen’s d dan jelaskan arah B minus A. Saya tidak menyamakan ukuran efek dengan signifikansi statistik atau memakai rumus tersebut untuk pasangan tanpa penyesuaian.
9. Bandingkan FORECAST.LINEAR dan TREND pada input jam sama. Saya memastikan input berada pada rentang yang dapat ditafsirkan dan tidak menganggap prediksi sebagai efek intervensi.
10. Buka Tren dan telusuri moving average. Amati bahwa baris awal belum mempunyai cukup observasi; label tersebut tidak diganti nol. Periksa ramalan periode berikutnya secara terpisah.
11. Tekan **Ctrl+Alt+F9**, cocokkan dengan kunci, dan tulis interpretasi yang menyebut ukuran, ketidakpastian, asumsi, serta batas kesimpulan.

## 6. Contoh formula dan arti tiap argumen

`=CONFIDENCE.T(0,05;STDEV.S(Data!E2:E61);COUNT(Data!E2:E61))` menerima alpha, simpangan baku sampel, dan jumlah pengamatan. Hasilnya margin, bukan batas bawah atau atas. Saya menambahkannya pada mean sesuai arah batas.

`=CORREL(Data!C2:C61;Data!E2:E61)` memakai pasangan jam dan skor pada baris sama. `=INDEX(LINEST(Data!E2:E61;Data!C2:C61;1;1);1;1)` mengambil slope. Argumen satu pada LINEST mempertahankan intercept dan meminta statistik tambahan; INDEX memilih baris serta kolom keluaran.

`=T.TEST(Data!E2:E31;Data!E32:E61;2;3)` adalah Welch dua sisi. Tipe satu dipakai pada rentang skor awal dan akhir yang sejajar. Pengujian satu sisi hanya tepat bila arah hipotesis ditetapkan sebelumnya dan sesuai pertanyaan; saya tidak memilihnya setelah melihat hasil.

`=FORECAST.LINEAR(8;Data!E2:E61;Data!C2:C61)` memberikan prediksi pada jam input. TREND memakai known_y, known_x, dan new_x dalam urutan yang sesuai. `=AVERAGE(Tren!C23:C25)` merangkum tiga nilai terakhir; ia tidak mengukur akurasi ramalan di luar sampel.

## 7. Latihan mandiri minimal lima soal bertingkat

1. Dasar: hitung mean, median, SD, dan varians serta jelaskan satuannya.
2. Dasar: buktikan frekuensi histogram mencakup seluruh skor sekali.
3. Menengah: bentuk CI mean dan jelaskan interpretasi proseduralnya.
4. Menengah: jelaskan korelasi serta slope tanpa klaim sebab-akibat.
5. Menengah: pilih uji Welch atau paired sesuai struktur pengamatan dan jelaskan ekornya.
6. Lanjut: hitung Cohen’s d, sebutkan arah serta penyebutnya, dan bedakan dari p-value.
7. Lanjut: bandingkan prediksi linear dengan moving average serta rancang pemisahan periode untuk uji ramalan yang jujur.

## 8. Kunci jawaban

[Kunci lengkap](latihan/kunci_jawaban.md) memuat hasil pandas/scipy sampai ketelitian yang dapat dibandingkan dengan workbook. Saya menerima interpretasi yang menyatakan ukuran dan asumsi, bukan hanya keputusan signifikan atau tidak signifikan.

Welch, paired, dan equal variance mempunyai pertanyaan serta asumsi berbeda. Prediksi FORECAST dan TREND pada input sama harus cocok, tetapi kecocokan itu tidak menunjukkan kualitas ramalan masa depan. Moving average harus memakai jendela yang disebutkan tanpa memasukkan observasi masa depan.

## 9. Kesalahan umum dan cara memeriksa hasil

Memilih paired untuk kelompok berbeda mengasumsikan pasangan yang tidak ada. Mengacak urutan salah satu kolom merusak pasangan korelasi dan uji berpasangan. Saya memeriksa ID serta kesejajaran sebelum menghitung.

Menyebut CI sebagai rentang seluruh individu mencampur ketidakpastian mean dengan sebaran pengamatan. Menganggap p-value kecil sebagai efek besar juga keliru. Saya melaporkan ukuran efek serta keterbatasan bersama hasil inferensial dan tidak membuat klaim populasi dari simulasi.

## 10. Checklist kompetensi

- [ ] Saya dapat membaca pusat, sebaran, dan histogram.
- [ ] Saya dapat membentuk CI serta menyatakan asumsi dan interpretasinya.
- [ ] Saya dapat membedakan korelasi, regresi, serta sebab-akibat.
- [ ] Saya dapat memilih tipe dan ekor uji serta membaca ukuran efek.
- [ ] Saya dapat membedakan smoothing, prediksi, dan validasi di luar sampel.

## 11. Catatan kompatibilitas Excel 2016/2019/365, LibreOffice, dan Google Sheets

Fungsi yang dipakai mempunyai padanan pada Excel 2016/2019/365. LINEST dan TREND diambil secara skalar melalui INDEX agar file tidak memerlukan dynamic array. Analysis ToolPak tidak tertanam atau dijalankan; formula worksheet menjadi jalur utama.

LibreOffice dipakai untuk recalc dan [preview nyata](assets/preview.png). Interaksi Excel dan impor Google Sheets belum diuji langsung; saya memeriksa kembali nama fungsi statistik, rentang, dan regional. Pembangunan ulang memakai `python tools/build.py`, lalu `python tools/course_runtime.py 12_statistika-terapan` dari folder belajar. Data sintetis tidak membuktikan efektivitas pelatihan atau performa bisnis nyata.
