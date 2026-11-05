# Masalah kualitas data yang saya sisipkan dan periksa

Saya mengaudit nilai mentah sebagai teks agar N/A, spasi, dan bidang kosong tidak lebih dahulu diubah oleh pembaca CSV. Catatan awal generator menghitung penyisipan pada 414 responden, kemudian enam baris identik ditambahkan. Dengan demikian hitungan sebelum duplikasi dan hitungan CSV akhir sama-sama diperlukan.

| Masalah / kunci JSON | 414 ID unik | Seluruh 420 baris CSV |
|---|---:|---:|
| sel_kosong_skor | 174 | 175 |
| teks_na | 12 | 12 |
| teks_strip | 10 | 10 |
| angka_berspasi | 15 | 16 |
| luar_rentang_0 | 5 | 5 |
| luar_rentang_6 | 5 | 5 |
| luar_rentang_7 | 3 | 3 |
| usia_outlier | 2 | 2 |
| tanggal_format_dmy | 62 | 62 |
| kanal_kapitalisasi | 34 | 34 |

174 dari 5.796 sel skor pada ID unik kosong karena penyisipan, sekitar 3,00%. Pada CSV akhir ada 175 dari 5.880 sel skor kosong, sekitar 2,98%. Hitungan ini mencakup q01–q12, kepuasan keseluruhan, dan nps; komentar opsional tidak termasuk. N/A dan tanda minus adalah kategori tambahan, bukan bagian hitungan sel kosong. Angka 0, 6, dan 7 di atas hanya dihitung pada q01–q12 dan kepuasan; pada nps ketiganya valid.

Saya menemukan enam ID bersalin identik: R-0094, R-0097, R-0179, R-0297, R-0377, R-0390. Saya mempertahankan kemunculan pertama, menandai salinan berikutnya `dipakai=0`, dan mempertahankan semua baris sebagai jejak audit. Tidak ada penggabungan jawaban atau pemilihan responden berdasarkan skor.

## Penanganan yang saya gunakan

Saya mengonversi angka berspasi dengan TRIM dan VALUE; sel kosong, N/A, minus, dan skor salah menjadi kosong. Usia 150 dan 3 menjadi kosong tanpa menghapus jawaban survei lain. Tanggal dibaca dengan dua format eksplisit; kanal dipetakan ke daftar standar. Saya tidak mengisi missing dengan mean, tidak mengganti skor salah dengan batas terdekat, dan tidak mengubah nol NPS menjadi missing.

## Kesalahan yang saya temukan pada bahan

Parser referensi awal memakai format campuran dengan preferensi hari di depan. Pada versi pandas yang dijalankan, 133 tanggal ISO dengan hari dan bulan sama-sama ≤12 tertukar. Saya memperbaiki generator dengan format `%Y-%m-%d` untuk ISO dan `%d/%m/%Y` untuk pola garis miring, kemudian membangkitkan ulang referensi. [Daftar koreksi](../data/processed/koreksi_tanggal_referensi.csv) menunjukkan ID, nilai lama, dan nilai baru. Saya tidak mengubah CSV mentah; statistik skor tidak bergantung pada koreksi ini.

Saya memperluas JSON dengan `cakupan_hitungan` dan `hitungan_csv_akhir`; kunci hitungan penyisipan awal tetap tersedia. Verifier menghitung ulang kedua cakupan langsung dari CSV dan menolak ketidakcocokan. Referensi juga dicocokkan dengan implementasi pembersihan independen sebelum workbook dibangun.
