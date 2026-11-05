# Analisis statistik dan interpretasi

Rumus dalam dokumen memakai titik koma untuk Excel regional Indonesia; regional Inggris memakai koma. Saya memakai fungsi kompatibel Excel 2016, dengan prefiks format internal yang diperlukan pada XLSX.

## Populasi sasaran dan observasi

Saya mensimulasikan survei sukarela pada pelanggan Toko, Online, dan Call Center. Data bukan hasil pengambilan sampel acak dan tidak memiliki bobot desain. Unit independensi diasumsikan responden; salinan ID dibuang. Self-report dan skala ordinal membatasi interpretasi jarak skor. Tidak ada hubungan sebab-akibat yang dapat disimpulkan dari analisis berikut.

## Skor, CI, dan reliabilitas

Saya menghitung mean tiga item hanya pada responden lengkap per dimensi. Pendekatan ini menghindari skor dengan jumlah item berbeda, tetapi dapat mengubah komposisi sampel bila missing tidak acak. Tidak ada imputasi atau asumsi bahwa jawaban hilang sama dengan netral.

| Dimensi | n valid | Mean (1–5) | CI 95% | Alpha |
|---|---:|---:|---|---:|
| Produk | 369 | 4,012 | 3,940–4,083 | 0,813 |
| Layanan | 375 | 3,718 | 3,638–3,798 | 0,835 |
| Harga | 372 | 3,246 | 3,169–3,322 | 0,800 |
| Pengiriman | 366 | 3,462 | 3,376–3,548 | 0,850 |

Margin CI adalah `=CONFIDENCE.T(0,05;SD;n)`, setara t(0,975;n−1) × SD / √n. Saya memakai sample SD dan df n−1. CI tidak menyatakan probabilitas 95% bahwa mean tetap berada pada rentang tertentu; prosedur mempunyai cakupan nominal di bawah asumsi yang sesuai. CI ini tidak mengatasi bias seleksi dan tidak membenarkan generalisasi pelanggan nyata.

Cronbach's alpha memakai k=3 serta varians sampel item dan skor jumlah pada responden yang sama. Alpha Harga yang dibulatkan menjadi 0,800 sebenarnya 0,799962501; saya tidak mengubah kategori berdasarkan pembulatan tampilan. Alpha tinggi bukan bukti validitas, unidimensionalitas, atau ketiadaan redundansi item. Interpretasi alpha sebagai reliabilitas juga memerlukan asumsi seperti kesetaraan pengukuran yang tidak diuji di sini.

## Uji t Welch dua sisi

Saya membandingkan kepuasan keseluruhan Toko dengan Online; hipotesis nol mean sama, alternatif mean berbeda. Formula adalah `=T.TEST(Clean_Data!AS2:AS421;Clean_Data!AT2:AT421;2;3)`. Saya tidak memakai paired test dan tidak mengasumsikan varians sama. Nilai kosong berupa teks diabaikan oleh fungsi statistik.

| Ukuran | Hasil |
|---|---:|
| n Toko / Online | 186 / 132 |
| Mean Toko / Online | 4,032258 / 3,750000 |
| Selisih Toko − Online | 0,282258 |
| t | 3,448126 |
| df Welch–Satterthwaite | 266,113460 |
| p dua sisi | 0,0006560274 |
| Cohen's d (SD gabungan) | 0,398472 |

Pada taraf 0,05 saya menolak H0, bersyarat pada independensi dan asumsi pendekatan mean. P-value adalah peluang memperoleh statistik setidaknya seekstrem yang diamati bila H0 dan asumsi uji berlaku; bukan peluang H0 benar dan bukan ukuran manfaat bisnis. Perbedaan profil kanal serta pola generator dapat menjelaskan asosiasi. Cohen's d memberikan skala efek deskriptif, tanpa label manfaat operasional otomatis.

Welch memakai df pecahan untuk p-value T.TEST. Fungsi T.INV.2T pada Excel memotong argumen df menjadi bilangan bulat; nilai kritis tampil sebagai ilustrasi dan tidak digunakan untuk menghitung p-value atau keputusan utama. Jika banyak pasangan/variabel diuji, saya menetapkan hipotesis lebih dahulu dan mempertimbangkan koreksi multipel. Uji Toko–Online di sini adalah perbandingan yang ditetapkan; tabel lain bersifat eksploratif.

## CORREL dan LINEST deskriptif

Saya memakai 248 kasus lengkap seluruh empat dimensi dan kepuasan keseluruhan. Saya tidak melakukan pairwise deletion terpisah untuk setiap korelasi, agar sampel korelasi dan regresi sama. Model deskriptif adalah kepuasan = intersep + bProduk × Produk + bLayanan × Layanan + bHarga × Harga + bPengiriman × Pengiriman.

| Prediktor | Koefisien | Galat baku | r dengan kepuasan |
|---|---:|---:|---:|
| Intersep | 0,570548 | 0,188367 | - |
| Produk | 0,257719 | 0,048758 | 0,576619 |
| Layanan | 0,270052 | 0,044360 | 0,602314 |
| Harga | 0,195873 | 0,046612 | 0,573209 |
| Pengiriman | 0,189917 | 0,041966 | 0,571425 |

R² = 0,583745, df sisa = 243, dan galat baku residual = 0,452775. LINEST mengembalikan urutan koefisien terbalik; saya mengambil kolom melalui INDEX agar label Produk dan Pengiriman tidak tertukar. Saya membandingkan koefisien dan galat baku dengan least squares numpy.

Prediktor saling berkorelasi; koefisien menggambarkan asosiasi bersyarat, bukan kontribusi kausal. Intersep pada semua skor nol berada di luar skala dan tidak ditafsirkan secara bisnis. Galat baku standar LINEST mengasumsikan bentuk linear, error independen dan homoskedastis; normalitas residual relevan untuk inferensi kecil-sampel. Saya belum melakukan validasi prediktif atau diagnostik lengkap, sehingga model ini hanya latihan deskriptif. Membandingkan besarnya koefisien tidak membuktikan prioritas perbaikan.

## NPS dan eksplorasi tambahan

NPS = (119 − 107) / 403 × 100 = 2,977667. Jawaban 0 tetap valid; missing tidak dihitung sebagai detraktor. NPS per kota/segmen mempunyai pembagi validnya sendiri dan tidak selalu stabil pada kelompok kecil. Saya tidak menyediakan CI NPS atau mengklaim tren dari satu gelombang.

Calculation juga menyediakan F ANOVA klasik sebagai perhitungan deskriptif tiga kanal, tanpa keputusan memakai tabel kritis pendekatan. Untuk menjalankan ANOVA manual di Excel: **File > Options > Add-ins > Manage: Excel Add-ins > Go > Analysis ToolPak**, lalu **Data > Data Analysis > Anova: Single Factor**; siapkan tiga kolom kepuasan valid per kanal, pilih **Grouped By: Columns**, centang **Labels in First Row**, dan alpha 0,05. Cocokkan F/df dengan Calculation. ANOVA klasik mengasumsikan varians homogen; bila tidak sesuai, gunakan Welch ANOVA di perangkat statistik. Langkah manual ini belum dijalankan dalam verifikasi workbook.
