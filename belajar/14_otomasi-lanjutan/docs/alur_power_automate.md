# Rancangan alur Power Automate

Saya menyediakan rancangan, bukan flow yang sudah dibuat atau diaktifkan. Kode Office Scripts juga belum dieksekusi pada host Microsoft 365. Saya memakai salinan data sintetis saat pengujian manual.

1. Simpan salinan XLSX pada OneDrive for Business atau SharePoint yang diizinkan organisasi. Tetapkan lokasi workbook dan script secara eksplisit; jangan menaruh token atau kata sandi di sel.
2. Buat Instant cloud flow dengan **Manually trigger a flow** sebagai tahap awal. Parameter input mencatat periode dan versi sumber; jadwal otomatis baru dipertimbangkan setelah pengujian serta otorisasi proses selesai.
3. Pada pengaturan trigger, batasi concurrency menjadi satu agar dua eksekusi tidak menulis workbook bersamaan. Catat run ID dan waktu dari riwayat flow.
4. Tambahkan **Excel Online (Business) > Run script**. Pilih lokasi file, workbook salinan, dan script PeriksaPermintaan. Script memeriksa schema, ID unik, nilai numerik, serta status sebelum menulis ulang formula nilai.
5. Tambahkan Condition terhadap hasil `status`. Jika LULUS, gunakan **Compose** untuk merangkum periode, versi, jumlah baris, dan run ID. Jika KOSONG, tandai bahwa tidak ada data; jangan menyebutnya pemrosesan sukses atas catatan yang tidak ada.
6. Gunakan Scope pemeriksaan dan Scope kegagalan dengan **Configure run after** untuk mencatat error/timeout. Hasil gagal perlu ditinjau; jangan meneruskan status seolah semua catatan telah diproses.
7. Sebelum mengubah status atau membuat keluaran bisnis pada pengembangan berikutnya, tetapkan kunci idempotensi gabungan request_id dan versi. Periksa apakah versi tersebut sudah diproses, lalu simpan jejak setelah operasi berhasil. Rancangan contoh ini belum melakukan perubahan status atau pengiriman pesan.
8. Uji salinan valid, ID duplikat, tipe qty salah, status tidak dikenal, tabel kosong, serta eksekusi ulang sumber yang sama. Bandingkan jumlah baris dan total dengan workbook pembanding. Simpan hasil pengujian host sebelum mengaktifkan trigger otomatis.

VBA sesuai untuk proses desktop dengan akses lokal yang disetujui. Office Scripts sesuai untuk workbook pada lingkungan Microsoft 365 yang mendukung Automate. Power Automate mengatur trigger, urutan tindakan, izin, serta pencatatan; ia tidak menggantikan validasi isi data.
