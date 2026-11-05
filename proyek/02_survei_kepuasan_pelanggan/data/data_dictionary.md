# Kamus data

Saya menyimpan satu baris per catatan pengisian pada `raw/survei_raw.csv`, lalu satu baris per ID unik pada `processed/survei_bersih_referensi.csv`. Kunci analisis adalah respondent_id; tidak ada tabel pelanggan atau join operasional lain. Tanggal pengisian berada pada tahun 2024. Seluruh skala, label pertanyaan, kategori, dan satuan ada di [codebook](../codebook.md).

| Kolom CSV mentah | Posisi Raw_Data | Tipe penyimpanan workbook | Satuan / makna |
|---|---|---|---|
| respondent_id | A | Teks | Kunci responden; enam ID muncul dua kali |
| tanggal_survei | B | Teks | YYYY-MM-DD atau DD/MM/YYYY |
| kanal, kota, segmen | C:E | Teks | Label kategori |
| usia, jenis_kelamin, lama_pelanggan_bulan | F:H | Teks | Tahun, kode L/P, bulan |
| q01–q12 | I:T | Teks | Jawaban Likert 1–5, termasuk gangguan yang disengaja |
| kepuasan_keseluruhan | U | Teks | Skala kepuasan 1–5 |
| nps | V | Teks | Skala rekomendasi 0–10 |
| komentar | W | Teks opsional | Komentar sintetis; boleh kosong |

Saya mempertahankan seluruh input sebagai teks, termasuk angka berspasi dan literal N/A. Pada Clean_Data A:V, hasil formula mengonversi tanggal dan angka, menstandarkan kanal, dan mempertahankan missing sebagai string kosong. W adalah status duplikat, X penanda dipakai, Y:AB skor dimensi, AC kepuasan untuk analisis, AD rekomendasi untuk analisis, AE penanda kasus lengkap, AF nomor urut kasus lengkap, AG:AR item tersaring untuk alpha, AS:AU kepuasan per kanal, dan AV komentar. AW:BF adalah hitungan masalah mentah per baris. Saya tidak menjumlahkan A:V tanpa filter X=1.

File referensi mempunyai 23 kolom mentah yang sudah dibersihkan ditambah empat skor dimensi (27 kolom). Nilai hilang ditulis sebagai bidang kosong CSV. Pada Reference, isi file tersebut menjadi input pembanding yang dapat ditinjau; Check_Detail mencari kemunculan pertama ID dan membandingkan semua 27 kolom.

`catatan_masalah_disengaja.json` memisahkan jumlah gangguan pada 414 ID unik dari hitungan CSV akhir 420 baris. `koreksi_tanggal_referensi.csv` mencatat 133 tanggal pada referensi bahan yang diperbaiki tanpa mengubah tanggal mentah. `rekonsiliasi.csv` berisi lokasi sel dan pasangan nilai Python/workbook. `ringkasan.json` merekam statistik, versi mesin, jumlah rumus, dan hash SHA-256 input serta workbook. [Provenans sumber](processed/provenans.json) merekam hash blob bahan awal.
