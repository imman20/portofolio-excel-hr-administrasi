# Modul 13: What-if dan optimasi

> Saya memakai kasus produksi dan laba sintetis untuk latihan. Formula dokumen memakai titik koma pada Excel regional Indonesia; regional Inggris memakai koma. Goal Seek, Data Table analisis, Scenario Manager, dan Solver belum tertanam atau dijalankan melalui antarmuka Excel.

## 1. Tujuan dan prasyarat

Saya belajar membedakan pencarian input untuk mencapai target, perhitungan sensitivitas, penyimpanan skenario, dan optimasi dengan kendala. Saya menggunakan model formula yang sama untuk menunjukkan kapan Goal Seek, Data Table satu serta dua variabel, Scenario Manager, atau Solver sesuai dengan pertanyaan.

Prasyaratnya adalah referensi, SUMPRODUCT, logika, dan kemampuan membaca satuan. Saya mengerjakan [latihan](latihan/latihan_13.xlsx) pada salinan, kemudian memeriksa [solusi](latihan/solusi_13.xlsx). Solusi optimal dihitung terpisah dengan scipy.optimize.linprog dan ditulis pada kunci, sehingga hasil Solver manual mempunyai pembanding independen.

## 2. Kegunaan di pekerjaan

Pertanyaan “berapa qty agar laba mencapai target” berbeda dari “kombinasi produk mana memberi laba terbesar dengan kapasitas terbatas”. Pertanyaan pertama mempunyai satu input yang dicari, sedangkan kedua memerlukan keputusan bersama serta kendala. Saya memilih alat berdasarkan struktur masalah, bukan berdasarkan menu yang paling dikenal.

Sensitivitas membantu melihat bagaimana hasil berubah jika harga atau volume berbeda. Ia tidak memberikan probabilitas setiap skenario dan tidak otomatis menyatakan skenario mana paling mungkin. Model produksi ini juga tidak menyatakan rencana perusahaan nyata; semua kapasitas, kontribusi, dan asumsi merupakan simulasi.

## 3. Penjelasan dari dasar

Model memisahkan input, keputusan, dan hasil. Harga, biaya variabel, serta biaya tetap adalah input kasus laba. Qty adalah input yang dapat dicari Goal Seek. Laba merupakan hasil formula. Saya tidak mengetik laba target pada sel hasil untuk membuat model terlihat memenuhi tujuan.

Goal Seek mengubah satu sel input agar satu sel formula mendekati nilai target. Ia tidak mengelola banyak kendala seperti kapasitas atau nonnegativitas secara umum. Jika model mempunyai beberapa solusi atau hubungan tidak sesuai, hasil perlu diperiksa. Pada contoh linear ini, solusi aljabar menjadi pembanding yang jelas.

Data Table dalam What-If Analysis berbeda dari Excel Table untuk struktur sumber. Data Table mengevaluasi formula untuk daftar nilai input. Satu variabel memakai satu input sel; dua variabel memakai satu daftar di baris dan satu di kolom. Saya menentukan arah input berdasarkan penempatan daftar agar harga tidak tertukar dengan qty.

Scenario Manager menyimpan beberapa kelompok input bernama. Ia membantu membandingkan kombinasi yang dipilih, tetapi tidak mencari optimum dari seluruh kemungkinan. Saya menyebut asumsi yang berubah dan menjaga input lain tetap agar perbedaan dapat dijelaskan.

Solver mempunyai sel tujuan, sel keputusan, dan kendala. Kasus produksi memaksimalkan total margin dua produk dengan batas jam serta bahan. Laba pada LP berarti total margin kontribusi; biaya lain tidak dimodelkan pada kasus produksi tersebut. Total pemakaian dihitung sebagai kombinasi linear qty dan koefisien per unit. Karena model linear dan kontinu, metode Simplex LP menjadi pilihan yang sesuai pada Excel.

Asumsi linear berarti margin per unit dan kebutuhan sumber daya per unit tetap pada seluruh rentang keputusan. Tidak ada diskon volume, biaya setup, ukuran batch, atau ketergantungan antarkeputusan. Jika asumsi berubah, model ini tidak otomatis tetap berlaku dan pilihan metode perlu ditinjau.

Qty pada model LP adalah variabel kontinu. Solusi dapat pecahan, yang sesuai untuk demonstrasi matematika tetapi belum tentu untuk unit barang yang wajib utuh. Membulatkan solusi dapat melanggar kendala atau kehilangan optimalitas. Saya tidak menyebut hasil pembulatan sebagai solusi integer optimal tanpa menyelesaikan masalah integer yang sesuai.

Untuk pembanding worksheet, saya mengevaluasi titik asal, titik sumbu, dan perpotongan kendala pada kasus dua variabel ini. Titik yang tidak layak tidak dipilih. Ini membantu visualisasi solusi, tetapi bukan pengganti algoritme umum untuk model berdimensi lebih besar atau matriks kendala yang berbeda.

Toleransi numerik memisahkan galat representasi pecahan dari pelanggaran substantif. Sisa kapasitas yang sangat kecil dinormalisasi sesuai ambang yang dinyatakan. Saya tetap memeriksa nilai tujuan, keputusan, serta seluruh kendala terhadap hasil linprog, bukan hanya status berhasil.

## 4. Dataset latihan

[Data produk](latihan/data_13.csv) berisi margin serta kebutuhan jam dan bahan untuk produk A serta B. [Kapasitas](latihan/kapasitas_13.csv) menetapkan sumber daya tersedia. [Skenario](latihan/skenario_13.csv) berisi kelompok harga dan qty sintetis dengan seed tetap.

Sheet Kandidat menghitung titik pembanding untuk LP, Model memuat keputusan serta pemakaian sumber daya, dan What_If menyediakan model laba, sensitivitas satu serta dua variabel, dan tabel skenario biasa. Sel input Model pada latihan dimulai dari nol; solusi memakai formula kandidat optimal.

[Hasil optimal JSON](latihan/solusi_optimal.json) dan [kunci jawaban](latihan/kunci_jawaban.md) berasal dari linprog yang dijalankan. [VERIFIKASI.md](VERIFIKASI.md) mencocokkan keputusan, laba, pemakaian, serta sisa kapasitas. Fitur analisis manual tidak disamarkan sebagai objek yang sudah aktif di file.

## 5. Langkah praktik bernomor dengan instruksi klik

1. Buka What_If dan baca harga, biaya, qty, laba, serta target. Saya memastikan B6 merupakan formula dan B5 merupakan input yang boleh diubah.
2. Pilih **Data > What-If Analysis > Goal Seek**. Isi **Set cell** dengan B6, **To value** dengan target laba yang tercantum pada B7, dan **By changing cell** dengan B5. Jalankan pada salinan dan bandingkan qty dengan B8.
3. Untuk Data Table satu variabel, buat salinan area sensitivitas. Letakkan referensi ke B6 di pojok atas kolom hasil dan daftar qty di kolom kiri. Blok seluruh area, pilih **Data > What-If Analysis > Data Table**, lalu isi **Column input cell** dengan B5; Row input dibiarkan kosong.
4. Pada tabel dua variabel, G2 merujuk laba, H2:J2 berisi harga, dan G3:G5 berisi qty. Pada salinan area G2:J5, pilih Data Table; isi **Row input cell** dengan B2 untuk harga serta **Column input cell** dengan B5 untuk qty. Cocokkan dengan grid formula pembanding.
5. Pilih **Data > What-If Analysis > Scenario Manager > Add**. Gunakan changing cells B2 dan B5, beri nama skenario, lalu masukkan harga serta qty dari tabel skenario. Ulangi untuk kombinasi lain dan gunakan **Show** atau **Summary** untuk membandingkan.
6. Aktifkan Solver melalui **File > Options > Add-ins > Manage: Excel Add-ins > Go**, centang **Solver Add-in**, lalu **OK**. Ketersediaan serta menu dapat berbeda menurut aplikasi.
7. Buka Model pada workbook latihan. Pilih **Data > Solver**, tetapkan tujuan B4 dengan **Max**, dan changing variable cells B2:B3. Tambahkan B5 <= Master!B2 serta B6 <= Master!B3.
8. Tetapkan nonnegativitas B2:B3 dan pilih **Simplex LP**. Jangan tambahkan kendala integer bila sedang membandingkan dengan solusi kontinu pada kunci. Pilih **Solve** dan tinjau hasil.
9. Periksa sisa jam serta bahan, lalu cocokkan qty dan laba dengan linprog. Jika menggunakan workbook solusi sebagai salinan latihan, ganti formula keputusan dengan input awal terlebih dahulu sambil mempertahankan sheet Kandidat sebagai pembanding.
10. Isi ringkasan kuning, tekan **Ctrl+Alt+F9**, dan tulis asumsi yang membatasi penerapan. Saya tidak menyimpan skenario percobaan di atas bukti awal tanpa penjelasan.

## 6. Contoh formula dan arti tiap argumen

`=(B2-B3)*B5-B4` pada What_If menghitung laba dari margin unit dikalikan qty lalu dikurangi biaya tetap. `=(B7+B4)/(B2-B3)` memberi qty pembanding untuk target pada model linear, selama margin unit positif.

`=SUMPRODUCT(B2:B3;Data!B2:B3)` pada Model menjumlahkan hasil kali keputusan dan margin. Dengan koefisien jam atau bahan, pola yang sama menghitung pemakaian sumber daya. Saya menjaga urutan produk sama agar koefisien tidak tertukar.

Grid dua variabel memakai referensi campuran: harga berubah ke kanan dan qty berubah ke bawah, sementara biaya tetap serta biaya variabel dikunci. Tabel tersebut merupakan formula biasa yang dapat dibandingkan dengan Data Table analisis manual.

Linprog meminimalkan fungsi tujuan. Skrip memberikan negatif margin agar hasilnya setara dengan memaksimalkan laba, memakai matriks batas atas untuk kapasitas, dan menetapkan batas keputusan nonnegatif. Status sukses, keputusan, nilai tujuan, serta slack disimpan sebagai bukti.

## 7. Latihan mandiri minimal lima soal bertingkat

1. Dasar: cari qty untuk mencapai target laba dan bandingkan Goal Seek dengan solusi aljabar.
2. Dasar: jelaskan perbedaan Data Table analisis dengan Excel Table sumber data.
3. Menengah: buat tabel sensitivitas satu variabel dan sebutkan input selnya.
4. Menengah: buat tabel dua variabel tanpa menukar row input dan column input.
5. Menengah: simpan skenario harga serta qty dan jelaskan bahwa skenario bukan hasil optimasi.
6. Lanjut: jalankan Solver kontinu dan cocokkan keputusan, laba, serta semua kendala dengan linprog.
7. Lanjut: jelaskan mengapa membulatkan qty pecahan belum membuktikan optimalitas integer.

## 8. Kunci jawaban

[Kunci lengkap](latihan/kunci_jawaban.md) menyebut keputusan dan laba maksimum dari linprog beserta pemakaian kapasitas. Qty Goal Seek diperiksa terhadap rumus linear, sedangkan grid sensitivitas diperiksa pada kombinasi input yang sama.

Solusi dianggap cocok bila nilai tujuan dan keputusan sesuai pembanding serta kendala tidak dilanggar melampaui toleransi numerik. Saya tidak menerima label sukses saja sebagai bukti. Jika kebutuhan nyata mengharuskan unit utuh, saya menyatakan perlunya model integer terpisah.

## 9. Kesalahan umum dan cara memeriksa hasil

Mengisi sel hasil dengan angka target memutus model. Menukar input baris dan kolom pada Data Table menghasilkan kombinasi yang berbeda. Menggunakan GRG Nonlinear untuk kasus linear tidak memberi alasan untuk mengabaikan struktur model. Saya memeriksa alamat, satuan, serta metode sebelum menyimpulkan.

Kendala yang ditulis dengan arah salah dapat memperluas wilayah layak. Tidak menetapkan nonnegativitas dapat memberi qty negatif yang tidak masuk akal. Saya menghitung ulang pemakaian dari keputusan dan membandingkannya dengan kapasitas satu per satu.

## 10. Checklist kompetensi

- [ ] Saya dapat membedakan target, sensitivitas, skenario, dan optimasi.
- [ ] Saya dapat menyiapkan Goal Seek serta Data Table dengan input tepat.
- [ ] Saya dapat menyatakan tujuan, keputusan, dan kendala Solver.
- [ ] Saya dapat membandingkan hasil dengan linprog dan membaca slack.
- [ ] Saya dapat menjelaskan batas linearitas serta variabel kontinu.

## 11. Catatan kompatibilitas Excel 2016/2019/365, LibreOffice, dan Google Sheets

Langkah manual ditulis untuk Excel desktop yang menyediakan What-If Analysis dan Solver Add-in. Workbook berisi data serta formula pembanding; konfigurasi Goal Seek, Data Table analisis, Scenario Manager, dan Solver belum tertanam. Eksekusi antarmuka tersebut belum diuji langsung.

LibreOffice dipakai untuk recalc dan [preview nyata](assets/preview.png), sedangkan scipy.optimize.linprog benar-benar dijalankan sebagai pembanding optimasi. Google Sheets mempunyai pilihan alat berbeda dan tidak otomatis menjalankan konfigurasi Excel. Pembangunan ulang memakai `python tools/build.py`, lalu `python tools/course_runtime.py 13_what-if-dan-optimasi` dari folder belajar. Saya menyimpan percobaan manual pada salinan.
