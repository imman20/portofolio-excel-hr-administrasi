# Panduan saya menjelaskan proyek

## Pembukaan singkat

“Saya membuat proyek simulasi survei kepuasan untuk menunjukkan alur dari CSV bermasalah ke hasil yang bisa diaudit di Excel. Saya mempertahankan data mentah, memakai rumus pembersihan, dan membandingkan setiap sel hasil serta statistik dengan pandas/scipy. Hasilnya mencakup 414 responden unik dan NPS 2,98, dengan batas interpretasi yang saya tulis.”

Saya menyebut proyek ini sebagai latihan mandiri dengan data sintetis. Saya tidak mengklaim menangani perusahaan, klien, atau dampak operasional nyata.

## Demonstrasi yang saya pilih

1. Saya membuka satu contoh angka berspasi, tanggal campuran, dan duplikat pada Raw_Data, lalu menelusuri formula Clean_Data serta penanda X.
2. Saya menjelaskan pemilihan tiga item lengkap per dimensi dan perbedaan pembagi dimensi, NPS, serta regresi. Missing bukan nol.
3. Saya membuka tabel CI dan alpha, lalu menjelaskan mengapa alpha tinggi tidak membuktikan validitas. Untuk Welch saya menunjukkan dua sisi, tipe 3, t, df pecahan, p, dan ukuran efek.
4. Saya menunjukkan Check_Detail dan VERIFIKASI.md, kemudian mengubah satu skor pada salinan agar pemeriksaan gagal. Kesamaan total saja dapat menyembunyikan kesalahan berlawanan arah.
5. Saya menunjukkan parser referensi yang saya koreksi: 133 tanggal ISO sebelumnya tertukar karena interpretasi format. Saya mempertahankan CSV mentah dan daftar koreksi.

## Pertanyaan teknis yang saya siapkan

- **Mengapa Excel dan Python?** Saya ingin workbook dapat dipakai dan ditelusuri oleh pembaca Excel, sedangkan perhitungan independen mendeteksi kesalahan implementasi yang bisa lolos dari kontrol internal.
- **Mengapa tidak mengimputasi?** Saya memilih aturan sederhana yang transparan untuk latihan. Complete case mengurangi sampel dan dapat bias bila missing tidak acak; itu tetap keterbatasan, bukan solusi universal.
- **Apa arti p-value?** Saya membahasnya bersyarat pada H0 serta asumsi, bukan probabilitas H0 benar atau bukti manfaat bisnis.
- **Apa yang belum diuji?** Representativitas, validitas konstruk, diagnostik lengkap regresi, dampak intervensi, dan interaksi lintas versi Excel. Tidak ada hubungan sebab-akibat yang dapat disimpulkan.
- **Fitur apa yang sungguh tersedia?** Rumus hidup dan chart native. PivotTable, Slicer, Power Query, dan Analysis ToolPak belum tertanam; saya menjelaskan langkah manual secara jujur.

Saya menutup penjelasan dengan keputusan yang dapat ditinjau: aturan pembersihan, pembagi yang konsisten, bukti rekonsiliasi, serta rekomendasi yang masih harus diuji.
