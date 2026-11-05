# Modul 09: Visualisasi dan dashboard

> Saya membangun dashboard dari data sintetis untuk latihan komunikasi analisis. Rumus dokumen memakai titik koma untuk Excel regional Indonesia; regional Inggris memakai koma. Angka bukan hasil bisnis atau pencapaian kerja nyata.

## 1. Tujuan dan prasyarat

Saya belajar memilih KPI dan grafik berdasarkan keputusan yang perlu dibantu. Saya menghubungkan angka ringkasan, tren terhadap target, kontribusi kategori, sumber, satuan, dan tanggal pembaruan dalam satu laporan yang dapat ditelusuri. File solusi memakai chart native openpyxl yang merujuk sel workbook, bukan gambar grafik statis.

Prasyaratnya adalah agregasi bersyarat, laporan bulanan, dan pemahaman perbedaan nilai dengan rasio. Saya memakai [latihan](latihan/latihan_09.xlsx) untuk menghitung ringkasan dan mencoba membuat chart sendiri. [Solusi](latihan/solusi_09.xlsx) menyediakan contoh dashboard dengan grafik garis serta batang. Saya memeriksa hitungan sebelum menilai tampilannya.

## 2. Kegunaan di pekerjaan

Dashboard membantu pembaca menemukan pertanyaan penting dengan cepat: apakah penjualan mencapai target, bagaimana margin, dan bagian mana perlu ditinjau. Jika hanya menampilkan angka terbesar tanpa definisi, pembaca dapat menarik kesimpulan yang tidak didukung data. Saya menempatkan konteks dekat dengan ukuran agar laporan tidak bergantung pada penjelasan lisan.

Simulasi ini memisahkan penjualan dari laba kotor. Kategori dengan penjualan tinggi tidak otomatis memberikan margin tinggi, dan pencapaian target tidak menjelaskan penyebab perubahan. Saya memakai dashboard sebagai awal pemeriksaan, lalu mengarahkan pembaca ke tabel detail untuk menelusuri perbedaan.

## 3. Penjelasan dari dasar

KPI adalah ukuran yang dipilih karena relevan terhadap tujuan. Banyaknya angka bukan ukuran kualitas dashboard. Saya memilih penjualan bersih, laba kotor, margin kotor, dan pencapaian target karena keempatnya menghubungkan volume nilai dengan hasil serta rencana. Masing-masing harus mempunyai definisi, periode, dan satuan.

Penjualan bersih dihitung dari qty, harga, dan diskon pada baris transaksi. HPP memakai qty dikalikan HPP satuan. Laba kotor adalah selisih keduanya. Margin kotor pada dashboard merupakan total laba dibagi total penjualan, bukan rata-rata sederhana margin transaksi. Rata-rata rasio dapat memberi bobot yang tidak dimaksudkan pada transaksi kecil.

Pencapaian target membandingkan penjualan dengan target pada periode yang sama. Target adalah input perencanaan simulasi, bukan hasil yang dihitung dari penjualan. Saya tidak mengubah target agar pencapaian tampak baik. Bila target nol, rasio memerlukan penanganan serta penjelasan; dataset ini memakai target positif yang diperiksa.

Grafik garis cocok untuk urutan waktu dan perbandingan terhadap target sepanjang periode. Batang cocok untuk membandingkan kategori diskret. Scatter cocok untuk hubungan dua ukuran numerik, sedangkan histogram memperlihatkan distribusi satu ukuran. Saya tidak memilih grafik hanya karena bentuknya menarik atau tersedia pada menu pertama.

Grafik batang umumnya memerlukan sumbu nol agar panjang batang tetap mewakili besaran secara wajar. Grafik garis dapat memakai skala yang ditinjau menurut kebutuhan, tetapi perubahan skala harus jelas dan tidak dipakai untuk melebih-lebihkan fluktuasi. Saya menjaga satuan seri yang dibandingkan sama.

Warna berfungsi sebagai penanda, bukan dekorasi untuk setiap titik. Saya memakai sedikit warna dengan kontras cukup, label terbaca, dan legenda yang jelas. Efek tiga dimensi, gradient berlebihan, serta dua sumbu tanpa kebutuhan yang kuat dapat mengalihkan perhatian atau mengaburkan perbandingan.

Sumber data dan tanggal acuan berbeda dari tanggal penghitungan ulang. Data pada latihan berakhir pada periode historis tertentu, sedangkan file dapat dihitung ulang hari ini. Saya menampilkan keduanya agar pembaca tidak mengira tanggal pembaruan berarti seluruh transaksi sudah mencakup hari tersebut.

Insight menghubungkan pola dengan pertanyaan lanjutan yang masuk akal. Saya menyatakan observasi, kemungkinan penjelasan yang belum diuji, dan pemeriksaan berikutnya secara terpisah. Grafik penjualan tidak membuktikan kampanye menyebabkan kenaikan; data sintetis juga tidak mendukung klaim dampak nyata.

## 4. Dataset latihan

[CSV transaksi](latihan/data_09.csv) berisi 360 catatan pada Januari sampai Juni 2024. Input mencakup kategori, wilayah, qty, harga, HPP satuan, dan diskon. [CSV target](latihan/target_09.csv) menyediakan rencana bulanan. Kolom penjualan, HPP, dan laba dihitung dengan rumus dalam workbook.

Calculation berisi ringkasan bulanan serta kategori. Report merujuk ringkasan tersebut untuk KPI, tabel, chart, dan catatan. Latihan menyediakan pertanyaan tentang margin, pencapaian, kategori, dan rekonsiliasi. Pada file latihan chart belum dipasang agar saya dapat membuatnya melalui langkah manual.

[Kunci jawaban](latihan/kunci_jawaban.md) memuat hasil pandas dari input. [VERIFIKASI.md](VERIFIKASI.md) mencocokkan KPI dan bulan serta memeriksa bahwa solusi mempunyai dua chart native. Preview berasal dari render Report yang sebenarnya.

## 5. Langkah praktik bernomor dengan instruksi klik

1. Buka Instruksi, lalu Data dan Master. Saya membaca definisi penjualan, HPP, diskon, target, serta tanggal acuan sebelum merancang tampilan.
2. Pada Calculation, pilih satu nilai bulanan dan gunakan **Formulas > Trace Precedents**. Pastikan target dan realisasi berasal dari bulan yang sama.
3. Isi KPI kuning pada Latihan. Saya menghitung margin dari total laba dan total penjualan, kemudian membandingkannya dengan solusi.
4. Pilih Calculation A1:C7, lalu **Insert > Line or Area Chart > Line with Markers**. Gunakan **Chart Design > Select Data** untuk memastikan hanya penjualan dan target yang menjadi seri.
5. Pilih **Chart Design > Add Chart Element > Axis Titles**. Tulis satuan rupiah pada sumbu nilai dan bulan pada sumbu waktu. Ubah judul agar menyatakan perbandingan yang dibaca.
6. Pilih A11:B14 pada Calculation, lalu **Insert > Column or Bar Chart > Clustered Column**. Saya menampilkan kategori sebagai label dan penjualan sebagai nilai, tanpa memasukkan baris header sebagai data.
7. Pindahkan chart ke Report dengan **Cut** lalu **Paste**, atur ukuran supaya tidak menutupi KPI, dan gunakan **Format > Align** bila diperlukan. Saya menjaga jarak serta ukuran font agar pembaca tidak perlu memperbesar terus-menerus.
8. Tambahkan sumber, acuan data, dan tanggal penghitungan ulang. Saya tidak mengetik angka KPI ulang sebagai teks karena hubungan ke perhitungan akan terputus.
9. Buka **View > Page Break Preview**, kemudian **Page Layout > Print Area > Set Print Area** pada area laporan. Periksa **File > Print** agar label dan grafik tidak terpotong.
10. Tekan **Ctrl+Alt+F9**, cocokkan KPI dan tabel, lalu tulis satu observasi serta tindak lanjut bersyarat. Saya membandingkan dengan solusi dan preview tanpa menganggap desain sebagai bukti perhitungan.

## 6. Contoh formula dan arti tiap argumen

`=ROUND(E2*F2*(1-H2);0)` membentuk penjualan bersih per baris dengan qty, harga, diskon, dan pembulatan rupiah. `=E2*G2` membentuk HPP, lalu laba adalah selisih penjualan dan HPP. Saya menyatakan bahwa biaya operasi, pajak, dan retur tidak dimodelkan di latihan ini.

`=IF(C4=0;0;C5/C4)` menghitung margin dari KPI total pada Report. Nilai nol pada pembagi perlu ditinjau sebagai kondisi khusus; ia tidak membuktikan margin bisnis sebenarnya nol. Untuk data awal, penjualan positif sehingga rasio dapat dibandingkan langsung.

`=IF(SUM(Calculation!C2:C7)=0;0;C4/SUM(Calculation!C2:C7))` menghitung pencapaian keseluruhan dengan total target sebagai penyebut. Saya tidak merata-ratakan pencapaian bulanan secara sederhana karena target bulan dapat berbeda.

Chart merujuk Calculation melalui seri dan kategori, sehingga tidak memerlukan formula baru untuk setiap titik. Saya memeriksa rentang melalui Select Data; penambahan bulan pada sumber tidak selalu berarti rentang chart sudah diperluas.

## 7. Latihan mandiri minimal lima soal bertingkat

1. Dasar: hitung penjualan bersih dan laba kotor serta jelaskan satuannya.
2. Dasar: hitung margin dari total dan bedakan dari rata-rata margin transaksi.
3. Menengah: buat tren bulanan terhadap target dengan judul serta sumbu yang jelas.
4. Menengah: buat grafik kategori dan identifikasi kategori penjualan terbesar serta margin terendah.
5. Menengah: hitung pencapaian keseluruhan dengan penyebut yang benar.
6. Lanjut: temukan tiga anti-pola visual yang dapat menyesatkan pembaca dan jelaskan perbaikannya.
7. Lanjut: tulis observasi dan tindak lanjut tanpa mengarang sebab perubahan atau dampak bisnis nyata.

## 8. Kunci jawaban

[Kunci lengkap](latihan/kunci_jawaban.md) memuat nilai KPI dan identifikasi kategori dari pandas. Jawaban desain tidak harus memiliki posisi chart persis sama, tetapi harus mempertahankan definisi, satuan, sumber, periode, dan keterbacaan. Saya menilai apakah pembaca dapat menelusuri angka kembali ke Calculation.

Selisih tren terhadap sumber harus nol. Observasi kategori tidak boleh menyamakan nilai penjualan terbesar dengan margin tertinggi. Pernyataan target juga harus menyebut periode, karena pencapaian satu bulan dapat berbeda dari keseluruhan periode.

## 9. Kesalahan umum dan cara memeriksa hasil

Grafik dapat menampilkan baris total bersama detail sehingga nilai terhitung dua kali secara visual. Seri target dapat tertukar dengan label bulan. Saya memeriksa Select Data, tabel sumber, dan satu titik pada chart sebelum menilai pola.

Margin yang dirata-ratakan tanpa bobot dapat berbeda dari rasio total. Judul tanpa satuan dapat membuat rupiah disalahartikan sebagai unit barang. Saya memeriksa definisi ukuran dan label, lalu membandingkan jumlah kategori dengan total sumber.

## 10. Checklist kompetensi

- [ ] Saya dapat memilih KPI berdasarkan pertanyaan keputusan.
- [ ] Saya dapat membedakan nilai, rasio, target, dan pencapaian.
- [ ] Saya dapat memilih chart serta menjaga skala dan label.
- [ ] Saya dapat menyatakan sumber, acuan, serta tanggal pembaruan.
- [ ] Saya dapat menulis insight yang dibatasi bukti dan merekonsiliasi angka.

## 11. Catatan kompatibilitas Excel 2016/2019/365, LibreOffice, dan Google Sheets

Chart garis dan batang native serta fungsi inti tersedia pada Excel 2016/2019/365. Workbook tidak memakai dynamic array atau PivotChart. Chart solusi dibuat dengan openpyxl; fitur interaktif seperti Slicer belum tertanam dan tidak diperlukan untuk laporan ini.

LibreOffice dipakai untuk recalc dan [preview nyata](assets/preview.png). Interaksi Excel dan impor Google Sheets belum diuji langsung; posisi chart, format tanggal, dan legenda perlu diperiksa kembali. Pembangunan ulang memakai `python tools/build.py`, lalu verifikasi dari folder belajar melalui `python tools/course_runtime.py 09_visualisasi-dan-dashboard`. Saya menyimpan perubahan desain manual pada salinan terpisah.
