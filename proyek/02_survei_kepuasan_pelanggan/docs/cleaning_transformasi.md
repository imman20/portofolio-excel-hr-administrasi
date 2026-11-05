# Cleaning dan transformasi

Rumus dalam dokumen memakai titik koma untuk Excel regional Indonesia; regional Inggris memakai koma. Di file, fungsi Inggris dan koma disimpan sesuai format XLSX.

Saya membaca CSV dengan `dtype=str, keep_default_na=False` dan menyimpan Raw_Data sebagai teks. Nilai mentah tidak dibersihkan saat impor. Transformasi dimulai pada Clean_Data, satu baris untuk setiap baris mentah.

1. Saya mengambil ID dan kategori, lalu memangkas spasi pada kanal, kota, segmen, serta kode jenis kelamin. Kanal memakai INDEX/MATCH terhadap Master; pencocokan tidak peka kapitalisasi. Saya tetap mengaudit ketidakbakuan kapitalisasi dengan SUBSTITUTE yang peka huruf.
2. Saya membaca pola tanggal sebelum konversi: ISO menggunakan LEFT tahun, MID bulan, dan RIGHT hari; DD/MM/YYYY menggunakan LEFT hari, MID bulan, dan RIGHT tahun. DATE menghasilkan nilai tanggal Excel, dengan tampilan dd/mm/yyyy.
3. Saya memeriksa kosong sebelum VALUE(TRIM()). Skor harus bilangan bulat dalam rentang Master. IFERROR mengembalikan kosong hanya untuk input yang memang tidak dapat dikonversi; audit mentah dan rekonsiliasi mencatat efeknya.
4. Saya menandai kemunculan pertama dengan `=IF(COUNTIFS(Raw_Data!$A$2:$A2;Raw_Data!A2)=1;1;0)`. Dua salinan dengan ID sama harus identik; verifier menolak duplikat yang bertentangan.
5. Saya menghitung skor dimensi hanya jika X=1 dan COUNT tiga item=3. Contoh Produk: `=IF(AND($X2=1;COUNT(I2:K2)=3);AVERAGE(I2:K2);"")`. Item pembentuk alpha disaring memakai responden yang sama.
6. Saya membuat kolom khusus kepuasan dan NPS yang hanya berisi baris dipakai. Kolom NPS tetap mempertahankan nol valid. COUNT menentukan jumlah valid; COUNTIFS untuk detraktor juga mensyaratkan tidak kosong.
7. Saya menumpuk kasus lengkap ke Analisis_Set memakai nomor AF dan INDEX/MATCH. CORREL/LINEST menggunakan batas akhir INDEX sesuai jumlah lengkap, tanpa FILTER atau dynamic array.

Saya memakai `Reference` sebagai snapshot input pembanding, bukan pengganti rumus pembersihan. Check_Detail mencari ID, membandingkan tanggal sebagai ISO, angka sampai 12 desimal, teks secara sama, dan missing sebagai kosong. Ubah satu nilai mentah pada salinan untuk melihat pemeriksaan gagal; jangan menyunting reference agar selisih hilang.

## Batas kontrak input

Saya memvalidasi pola tanggal, jumlah baris, label kategori, ID, dan salinan identik sebelum build. Formula hanya dirancang untuk dua pola tanggal dan kategori yang didokumentasikan, bukan parser universal. Saya menjalankan builder/verifier ulang setelah mengganti dataset atau menambah baris; rentang 2:421 tidak berkembang otomatis.
