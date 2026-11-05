# Instruksi pengerjaan

Rumus dokumen memakai titik koma untuk Excel regional Indonesia; regional Inggris memakai koma. Saya mengerjakan latihan pada salinan, kemudian membandingkan dengan solusi.

## Menyiapkan dan mengenali data

1. Unduh workbook latihan, buka Readme dan Instruksi, lalu **File > Save As**. Pilih **Formulas > Calculation Options > Automatic**.
2. Baca Codebook. Pada Raw_Data A1:W421 pilih **Data > Filter**. Gunakan **Home > Conditional Formatting > Highlight Cells Rules > Text that Contains** untuk N/A dan minus. Jangan mengganti atau menghapus data di sheet ini.
3. Hitung masalah pada q01–nps, bedakan hitungan seluruh 420 baris dari hitungan 414 ID unik, lalu cocokkan dengan masalah_kualitas_data.md. Enam salinan identik harus ditandai, bukan digabungkan.

## Mengisi Clean_Data

1. Baris 2 pada A:AV berisi contoh formula lengkap. Pilih A2:AV421, lalu **Home > Fill > Down** atau **Ctrl+D**. Periksa bahwa ID dan referensi Raw_Data bergeser sesuai baris, sedangkan referensi Master tetap absolut.
2. Baca B untuk DATE yang membedakan ISO dan DMY; F untuk usia; I:V untuk skor. VALUE dan TRIM mengubah angka teks, tetapi missing dan invalid tetap kosong. Uji satu NPS nol valid pada salinan dan pastikan tidak dianggap missing.
3. Baca W:X untuk duplikat dan Y:AB untuk dimensi. Gunakan **Data > Filter** pada X, pilih 1, dan periksa bahwa mean dimensi kosong bila salah satu item tidak valid.
4. AW:BF disediakan sebagai kolom audit per baris. Salin rumus audit dari solusi, atau susun sendiri memakai COUNTBLANK, COUNTIFS, TRIM, VALUE, dan SUBSTITUTE. Audit kapitalisasi harus peka huruf; COUNTIFS saja tidak membedakan Toko/toko.
5. Pada Analisis_Set isi nomor urut 1 sampai 420, lalu INDEX/MATCH dari AF untuk menumpuk kasus lengkap. Gunakan IFERROR untuk baris sesudah jumlah lengkap, bukan untuk menyembunyikan skor yang salah.

## Mengisi Calculation dan Report

1. Mulai dari ukuran sampel dan distribusi item. Gunakan COUNTIFS dengan X=1 dan skor 1–5. Jumlah frekuensi harus sama dengan n valid item. Mean total tidak sama dengan mean tanpa bobot seluruh kanal.
2. Isi tabel dimensi dengan COUNT, AVERAGE, STDEV.S, dan CONFIDENCE.T. Pastikan setiap varians item alpha hanya memakai responden yang memiliki ketiga item valid.
3. Isi ringkasan kanal, kota, dan segmen. Kolom bantu AS:AU pada Clean_Data menyediakan jawaban valid untuk T.TEST. Gunakan dua sisi dan tipe 3; bandingkan t, df, p, serta Cohen's d dengan analisis_statistik.md.
4. Isi CORREL dan INDEX(LINEST(...;1;1);baris;kolom) pada kasus lengkap. Perhatikan urutan koefisien LINEST yang terbalik. Gunakan 1 sebagai argumen const/stats, tanpa formula array atau fungsi dynamic array.
5. Isi jumlah promotor, pasif, detraktor, dan pembagi NPS; pastikan missing tidak ikut kriteria detraktor. NPS bukan rata-rata skor rekomendasi.
6. Pada Report buat sel ringkas yang merujuk Calculation. Pilih **Insert > Bar Chart** untuk dimensi dan **Insert > Column Chart** untuk NPS kanal. Tampilkan skala, pembagi, CI, sumber, temuan, rekomendasi bersyarat, dan batas inferensi.

## Memeriksa hasil dan latihan mandiri

1. Pada Checks isi kolom C dengan formula dari bagian yang sesuai; kolom D adalah input pembanding independen. E dan F menghitung selisih serta status. Semua status harus PASS, sedangkan Check_Detail harus menunjukkan nol perbedaan pada setiap ID. Tekan **Ctrl+Alt+F9**.
2. Bandingkan [solusi](../workbook/solusi_survei_kepuasan_pelanggan.xlsx) dan [VERIFIKASI.md](../VERIFIKASI.md). Jangan mengganti input Reference untuk memaksa hasil cocok.
3. Pada salinan, ubah satu item valid, satu tanggal ISO ambigu, dan satu NPS menjadi nol. Jelaskan perubahan mean/alpha/NPS dan alasan Check_Detail harus gagal terhadap snapshot lama.
4. Jelaskan mengapa n dimensi tidak selalu 414, mengapa alpha membutuhkan kasus yang sama, mengapa p kecil tidak menunjukkan sebab-akibat, dan mengapa rata-rata kota tanpa bobot dapat salah. Jawaban ada pada definisi_kpi.md dan analisis_statistik.md.
5. Untuk ringkasan pivot manual: filter X=1, salin kolom hasil sebagai nilai pada sheet baru, pilih **Insert > PivotTable**, tempatkan kanal pada Rows dan kepuasan pada Values dengan **Value Field Settings > Average**. Tambahkan Count untuk pembagi dan cocokkan dengan Calculation. PivotTable belum tertanam pada file.
