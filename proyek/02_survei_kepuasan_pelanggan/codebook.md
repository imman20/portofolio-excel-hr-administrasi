# Codebook survei

Saya menggunakan satu ID per responden unik, empat dimensi dengan tiga item per dimensi, dan dua pertanyaan hasil yang terpisah. Item q01–q12 memakai kode 1 = Sangat tidak setuju, 2 = Tidak setuju, 3 = Netral, 4 = Setuju, dan 5 = Sangat setuju. Kepuasan keseluruhan memakai 1 = Sangat tidak puas sampai 5 = Sangat puas. Tidak ada item berkunci terbalik.

| Variabel | Pertanyaan / arti | Tipe dan nilai valid | Aturan yang saya gunakan |
|---|---|---|---|
| respondent_id | Kode unik catatan responden | Teks, R-0001 … R-0414 | Saya mempertahankan kemunculan pertama; salinan identik berikutnya ditandai dibuang. |
| tanggal_survei | Tanggal pengisian | Tanggal 2024; ISO atau DD/MM/YYYY | Saya membaca setiap pola dengan format eksplisit. |
| kanal | Kanal transaksi | Toko / Online / Call Center | Saya memangkas spasi dan mencocokkan label ke Master tanpa membedakan kapitalisasi. |
| kota | Kota responden | Surabaya, Jakarta, Bandung, Semarang, Medan, Makassar | Saya memangkas spasi. |
| segmen | Segmen pelanggan | Retail / Grosir / Korporat | Saya memangkas spasi. |
| usia | Usia dalam tahun | Bilangan bulat 17–80 | Saya mengosongkan 150 dan 3 tanpa membuang responden. |
| jenis_kelamin | Kode jenis kelamin | L / P | Saya memangkas spasi. |
| lama_pelanggan_bulan | Lama berlangganan | Bilangan bulat 1–120 | Saya mengonversi angka teks. |
| q01 | [Produk] Kualitas produk sesuai harapan saya | Likert 1–5 | Saya mengosongkan missing, teks, dan skor di luar rentang; spasi dipangkas. |
| q02 | [Produk] Produk awet dan dapat diandalkan | Likert 1–5 | Saya mengosongkan missing, teks, dan skor di luar rentang; spasi dipangkas. |
| q03 | [Produk] Pilihan/variasi produk memadai | Likert 1–5 | Saya mengosongkan missing, teks, dan skor di luar rentang; spasi dipangkas. |
| q04 | [Layanan] Staf melayani dengan ramah dan sopan | Likert 1–5 | Saya mengosongkan missing, teks, dan skor di luar rentang; spasi dipangkas. |
| q05 | [Layanan] Staf tanggap menyelesaikan masalah saya | Likert 1–5 | Saya mengosongkan missing, teks, dan skor di luar rentang; spasi dipangkas. |
| q06 | [Layanan] Informasi yang diberikan jelas dan akurat | Likert 1–5 | Saya mengosongkan missing, teks, dan skor di luar rentang; spasi dipangkas. |
| q07 | [Harga] Harga sebanding dengan kualitas | Likert 1–5 | Saya mengosongkan missing, teks, dan skor di luar rentang; spasi dipangkas. |
| q08 | [Harga] Harga bersaing dibanding tempat lain | Likert 1–5 | Saya mengosongkan missing, teks, dan skor di luar rentang; spasi dipangkas. |
| q09 | [Harga] Promo dan diskon menarik | Likert 1–5 | Saya mengosongkan missing, teks, dan skor di luar rentang; spasi dipangkas. |
| q10 | [Pengiriman] Pesanan tiba tepat waktu | Likert 1–5 | Saya mengosongkan missing, teks, dan skor di luar rentang; spasi dipangkas. |
| q11 | [Pengiriman] Barang tiba dalam kondisi baik | Likert 1–5 | Saya mengosongkan missing, teks, dan skor di luar rentang; spasi dipangkas. |
| q12 | [Pengiriman] Biaya pengiriman wajar | Likert 1–5 | Saya mengosongkan missing, teks, dan skor di luar rentang; spasi dipangkas. |
| kepuasan_keseluruhan | Secara keseluruhan, seberapa puas Anda dengan kami? | 1–5 | Saya memakai jawaban valid untuk mean/CI/uji kanal. |
| nps | Seberapa besar kemungkinan Anda merekomendasikan kami kepada teman/rekan? (0-10) | Bilangan bulat 0–10 | Saya mempertahankan nol valid; kosong bukan nol. |
| komentar | Komentar bebas opsional | Teks; boleh kosong | Saya mempertahankan teks apa adanya; tidak ada analisis sentimen otomatis. |
| dim_produk … dim_pengiriman | Mean tiga item tiap dimensi | 1–5 | Saya menghitung hanya jika ketiga item valid pada baris dipakai. |

Saya mengelompokkan NPS 9–10 sebagai promotor, 7–8 sebagai pasif, dan 0–6 sebagai detraktor. Pembagi adalah semua jawaban rekomendasi valid, bukan seluruh 414 responden. Kelengkapan dimensi tidak mensyaratkan NPS terisi.

Saya menyediakan Codebook di workbook untuk navigasi dan [kamus data](data/data_dictionary.md) untuk posisi kolom, grain, file, serta kolom bantu. Data dan komentar seluruhnya sintetis.
