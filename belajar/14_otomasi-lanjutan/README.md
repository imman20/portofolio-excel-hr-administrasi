# Modul 14: Otomasi lanjutan

> Saya memakai permintaan sintetis untuk latihan otomasi. Rumus Excel di dokumen memakai titik koma pada regional Indonesia; regional Inggris memakai koma. Kode VBA dan Office Scripts disediakan sebagai berkas terpisah; keduanya belum dieksekusi pada host Excel, dan flow Power Automate belum dibuat atau diaktifkan.

## 1. Tujuan dan prasyarat

Saya belajar memilih sarana otomasi berdasarkan lingkungan, cakupan tindakan, dan cara menjalankannya kembali dengan aman. Saya membahas VBA desktop, Office Scripts pada Microsoft 365, serta Power Automate untuk mengatur trigger dan urutan proses. Tujuan saya bukan membuat tombol yang sekadar terlihat bekerja, melainkan menjaga validasi, rekonsiliasi, serta catatan ketika proses gagal.

Prasyaratnya adalah tabel, logika, lookup, dan pemeriksaan hasil. Saya membuka [latihan](latihan/latihan_14.xlsx) pada salinan dan membaca sumber kode sebelum menggunakannya. [Solusi](latihan/solusi_14.xlsx) menyediakan hasil formula pembanding. Format XLSX tersebut tidak mempunyai makro VBA tertanam.

## 2. Kegunaan di pekerjaan

Otomasi dapat mengurangi langkah berulang seperti mengisi formula atau menjalankan pemeriksaan standar. Namun proses yang salah juga dapat diulang lebih cepat. Saya menetapkan syarat masuk dan hasil yang harus cocok sebelum memilih trigger otomatis. Kasus ini merupakan simulasi permintaan internal, bukan sistem operasional nyata.

Pilihan alat dipengaruhi lokasi workbook dan izin. VBA cocok untuk Excel desktop serta proses lokal yang disetujui. Office Scripts bekerja pada lingkungan Microsoft 365 yang mendukung Automate. Power Automate mengatur pemicu, konektor, serta penanganan eksekusi; ia tidak memperbaiki definisi data yang belum jelas.

## 3. Penjelasan dari dasar

Makro VBA adalah kode yang berjalan dalam aplikasi desktop. Ia dapat mempunyai akses luas sesuai lingkungan pengguna, sehingga saya memeriksa kode dan sumbernya sebelum mengaktifkan. Workbook XLSX tidak menyimpan proyek VBA; setelah impor kode pada salinan, format yang mendukung makro diperlukan jika kode ingin disimpan bersama file.

Office Scripts memakai TypeScript dan objek ExcelScript untuk mengakses workbook. Script dapat mengubah range, tabel, dan formula melalui API yang tersedia. Ia tidak sama dengan VBA yang dipindahkan bahasanya secara langsung; objek, izin, cara pemicu, dan batas host berbeda. Saya memeriksa dukungan akun serta kebijakan organisasi terlebih dahulu.

Power Automate menghubungkan trigger dengan tindakan. Contoh rancangan dimulai dari trigger manual, lalu menjalankan script dan merangkum hasil. Jadwal atau trigger otomatis baru dipertimbangkan setelah alur dasar diperiksa. Saya tidak menganggap keberhasilan suatu tindakan berarti semua data memenuhi aturan bisnis.

Validasi dilakukan sebelum penulisan. Kode contoh memeriksa nama tabel dan kolom, ID kosong atau duplikat, qty positif, harga nonnegatif, serta status dikenal. Bila masalah ditemukan, proses berhenti dengan pesan yang menjelaskan alasan. Saya tidak mengubah status permintaan atau menghapus sumber sebagai bagian dari contoh ini.

Idempotensi berarti menjalankan proses ulang pada masukan dan versi sama tidak membuat dampak ganda yang tidak diinginkan. Kode contoh hanya mengisi formula nilai yang sama, sehingga tidak menambah baris atau transaksi. Jika kelak proses membuat dokumen atau mengubah status, saya memerlukan kunci request_id dan versi serta jejak keberhasilan sebelum mengulangi tindakan.

Penanganan kegagalan perlu mempertahankan keadaan yang dapat ditinjau. Pesan sukses tidak boleh ditampilkan jika validasi gagal. Pada flow, kondisi error atau timeout diarahkan ke jalur pencatatan, bukan dilanjutkan seolah semua catatan telah selesai. Saya membedakan kesalahan isi data dari kegagalan akses atau layanan.

Concurrency berarti beberapa eksekusi dapat berjalan pada waktu sama. Dua proses yang menulis workbook bersamaan dapat bertabrakan. Rancangan awal membatasi concurrency dan mencatat run ID agar pemeriksaan tidak hanya mengandalkan waktu file berubah. Keputusan tersebut perlu disesuaikan ketika proses berkembang.

Keamanan makro dimulai dari prinsip menggunakan sumber yang dikenal dan membaca tindakan yang dilakukan kode. Saya tidak menurunkan pengaturan keamanan global atau mengaktifkan semua makro tanpa peninjauan. Tanda tangan, lokasi tepercaya, serta kebijakan akses merupakan bagian pengelolaan organisasi, bukan pengganti audit isi kode.

## 4. Dataset latihan

[CSV permintaan](latihan/data_14.csv) memuat 80 catatan sintetis dengan ID, tanggal, bagian, jenis, qty, harga, dan status. Table `tblPermintaan` pada Data menambahkan nilai, validasi, serta frekuensi ID. Status yang dikenal adalah BARU, DIPROSES, dan SELESAI.

Sheet Ringkasan menghitung jumlah serta nilai per status dengan rumus, sedangkan Latihan menyediakan pemeriksaan agregat. [VBA](vba/periksa_permintaan.bas), [Office Scripts](office_scripts/periksa_permintaan.ts), dan [rancangan flow](docs/alur_power_automate.md) disimpan terpisah agar dapat dibaca sebelum dijalankan manual.

[Kunci jawaban](latihan/kunci_jawaban.md) berasal dari CSV, dan [VERIFIKASI.md](VERIFIKASI.md) membuktikan recalc formula pembanding. Bukti tersebut tidak menyatakan runtime VBA, Office Scripts, atau konektor Power Automate sudah diuji. Saya menyimpan hasil pengujian host secara terpisah jika kelak dilakukan.

## 5. Langkah praktik bernomor dengan instruksi klik

1. Buka salinan latihan dan tinjau Data. Pastikan nama tabel melalui **Table Design > Table Name** adalah tblPermintaan, lalu cocokkan nama kolom dengan kode.
2. Baca berkas VBA terlebih dahulu. Kode memeriksa data sebelum mengisi formula dan tidak mempunyai tindakan penghapusan, pengiriman pesan, atau perubahan status. Saya tetap memakai salinan sintetis untuk percobaan pertama.
3. Pada Excel desktop yang diizinkan, tekan **Alt+F11**, pilih **File > Import File**, lalu pilih berkas BAS. Kembali ke Excel dan simpan salinan melalui **File > Save As** ke format yang mendukung makro bila ingin mempertahankan kode.
4. Gunakan pengaturan keamanan sesuai kebijakan yang berlaku. Jalankan **Developer > Macros > PeriksaDanHitung > Run** hanya setelah kode ditinjau dan sumber dipercaya. Saya tidak menggunakan instruksi ini untuk menurunkan keamanan global.
5. Bandingkan jumlah serta total sebelum dan sesudah. Jalankan lagi pada masukan sama dan periksa bahwa tidak ada catatan baru atau status yang berubah.
6. Untuk Office Scripts, buka salinan workbook pada lingkungan yang mendukung **Automate > New Script**. Tempel isi TypeScript, beri nama yang jelas, lalu tinjau nama tabel serta kolom sebelum menjalankan.
7. Jalankan script pada salinan valid dan baca hasil status serta jumlah baris. Cocokkan total worksheet dengan pembanding. Perubahan formula harus tetap menghitung nilai dari qty dan harga, bukan menempel hasil statis.
8. Pada salinan pengujian lain, coba ID duplikat atau qty bertipe teks. Proses seharusnya berhenti dengan alasan, bukan menampilkan sukses. Kembalikan sumber valid sebelum perbandingan awal.
9. Ikuti rancangan Power Automate terpisah bila memiliki lingkungan dan izin yang sesuai. Mulai dari trigger manual, Run script, Condition, dan Compose, dengan concurrency yang dibatasi serta jalur kegagalan.
10. Isi sel kuning pada Latihan, hitung ulang, dan bandingkan dengan solusi. Saya mencatat bahwa eksekusi host manual memerlukan bukti sendiri dan tidak diwarisi dari verifikasi formula.

## 6. Contoh formula dan arti tiap argumen

`=E2*F2` menghitung nilai permintaan. Kode contoh menulis pola calculated column `=[@qty]*[@harga_satuan]` agar formula tetap hidup dan mudah dibaca di Table. Saya tidak menggantinya dengan hasil numerik sekali pakai.

`=IF(AND(COUNTA(A2)=1;E2>0;F2>=0;OR(G2="BARU";G2="DIPROSES";G2="SELESAI"));"VALID";"PERIKSA")` adalah pemeriksaan worksheet. Frekuensi ID diperiksa terpisah melalui COUNTIFS agar duplikasi tidak tertutup oleh validasi nilai per baris.

VBA menggunakan ListObject dan nama ListColumns untuk mengurangi ketergantungan pada posisi visual. Office Scripts memeriksa header sebelum membaca indeks kolom. Kedua kode mengembalikan atau menampilkan kegagalan secara jelas; saya tidak memakai penanganan error untuk melanjutkan seluruh proses tanpa jejak.

Pada rancangan flow, hasil script menjadi masukan Condition. Status LULUS berarti pemeriksaan contoh selesai, bukan persetujuan bisnis terhadap seluruh permintaan. Perubahan status operasional memerlukan aturan tambahan yang belum diimplementasikan pada contoh ini.

## 7. Latihan mandiri minimal lima soal bertingkat

1. Dasar: hitung jumlah dan nilai seluruh permintaan sebagai pembanding sebelum otomasi.
2. Dasar: jelaskan perbedaan lokasi eksekusi VBA, Office Scripts, dan Power Automate.
3. Menengah: identifikasi validasi yang dilakukan kode sebelum menulis formula.
4. Menengah: jelaskan mengapa menjalankan ulang contoh tidak boleh menambah catatan.
5. Menengah: rancang pengujian ID duplikat dan qty salah tipe pada salinan, beserta hasil yang diharapkan.
6. Lanjut: jelaskan peran concurrency, run ID, serta jalur kegagalan pada flow.
7. Lanjut: sebutkan batas keamanan makro dan bukti tambahan sebelum alur diaktifkan otomatis.

## 8. Kunci jawaban

[Kunci lengkap](latihan/kunci_jawaban.md) memuat jumlah, nilai, status, dan rekonsiliasi dari CSV. Hasil setelah otomasi manual seharusnya tetap sama untuk sumber valid karena kode hanya menuliskan kembali formula yang setara. Masukan tidak valid harus menghasilkan penghentian yang jelas.

Jawaban desain harus menyatakan tindakan apa yang boleh berubah, kunci idempotensi, serta cara membedakan sukses dengan gagal. Saya tidak menerima “tidak ada pesan error” sebagai satu-satunya bukti, dan tidak mengklaim flow aktif tanpa hasil eksekusi host.

## 9. Kesalahan umum dan cara memeriksa hasil

Menjalankan script pada workbook salah dapat mengubah salinan yang tidak dimaksud. Nama tabel berubah dapat membuat proses berhenti. Saya memeriksa lokasi dan identitas sumber sebelum Run, kemudian membandingkan total serta jumlah baris.

Membungkus semua error dan selalu menampilkan sukses menyembunyikan kegagalan. Trigger paralel juga dapat menimbulkan penulisan ganda pada pengembangan yang membuat keluaran. Saya meninjau riwayat eksekusi dan menjaga pemisahan sumber dengan hasil.

## 10. Checklist kompetensi

- [ ] Saya dapat memilih alat sesuai host dan izin.
- [ ] Saya dapat membaca validasi serta tindakan kode sebelum menjalankan.
- [ ] Saya dapat memeriksa hasil sebelum dan sesudah eksekusi.
- [ ] Saya dapat menjelaskan idempotensi, concurrency, dan pencatatan gagal.
- [ ] Saya dapat menyatakan batas pengujian serta keamanan makro.

## 11. Catatan kompatibilitas Excel 2016/2019/365, LibreOffice, dan Google Sheets

VBA tersedia pada Excel desktop yang mendukungnya, dengan perbedaan fitur antar platform; penggunaan Scripting.Dictionary pada contoh perlu diuji pada host tujuan. Office Scripts serta integrasi Power Automate memerlukan lingkungan Microsoft 365 dan izin yang sesuai, bukan Excel 2016/2019 secara umum.

XLSX bawaan tidak memuat VBA atau Office Scripts tertanam. LibreOffice hanya dipakai untuk recalc formula dan [preview nyata](assets/preview.png); runtime kode host belum diuji. Google Sheets memakai sistem otomasi berbeda. Pembangunan ulang memakai `python tools/build.py`, lalu `python tools/course_runtime.py 14_otomasi-lanjutan` dari folder belajar. Saya tidak mengaktifkan flow atau mengirim pesan sebagai bagian pembuatan modul ini.
