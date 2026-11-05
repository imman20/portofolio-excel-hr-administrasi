# Panduan menjelaskan proyek

## Pembukaan

“Saya membangun proyek portofolio simulasi inventori untuk menunjukkan alur dari transaksi stok sampai usulan pemesanan yang dapat diperiksa. Saya memakai Python untuk data sintetis dan pembangunan file, Excel untuk rumus hidup dan laporan, serta perhitungan pandas independen untuk rekonsiliasi. Ini latihan analisis, bukan pengalaman kerja pada perusahaan atau klien.”

Saya menunjukkan Report terlebih dahulu, kemudian menelusuri satu SKU ke Calculation, Permintaan_Harian, Raw_Data, dan PO. Saya menjelaskan hubungan angka dengan keputusan: stok habis membutuhkan tinjauan pemenuhan segera; overstock membutuhkan pemeriksaan kebutuhan dan kemungkinan pemindahan sebelum pembelian baru.

## Keputusan teknis

Saya menggunakan kalender penuh agar hari tanpa transaksi tetap berpengaruh pada mean dan sigma. Menghitung hanya hari aktif akan menjawab pertanyaan yang berbeda. Saldo harian juga memberi penyebut turnover yang lebih representatif daripada dua titik awal/akhir tahun, dengan biaya lebih banyak rumus dan waktu recalc.

Saya memisahkan stok fisik dari on order. PO terbuka mengurangi kebutuhan pesan tambahan, tetapi belum menambah barang di gudang. Status risiko memakai stok fisik dibanding ROP; saya tetap perlu memeriksa jadwal kedatangan dan reliabilitas pemasok sebelum keputusan nyata.

Saya memakai HPP tetap per SKU untuk menjaga buku nilai dapat ditelusuri. Ini tidak mendemonstrasikan FIFO atau moving average cost. Saya menyatakan aturan ABC, termasuk perlakuan item pelintas batas dan tie, supaya orang lain bisa mereproduksi kelas yang sama.

Saya membangun rumus dengan INDEX/MATCH dan kolom bantu untuk kompatibilitas Excel 2016. SUMIFS pada kisi kalender membuat alur terlihat, tetapi lebih lambat daripada arsitektur agregasi khusus untuk data produksi besar. PivotTable dan Power Query tidak tertanam; jika ditanyakan, saya menunjukkan langkah pembanding manual dan batas tool.

## Pertanyaan yang perlu dijawab jujur

**Mengapa service level tidak otomatis tercapai?** Rumus safety stock memakai pendekatan normal, permintaan independen, dan lead time tetap. Simulasi juga membatasi pengeluaran oleh stok, sehingga data keluar bisa tersensor. Saya tidak memiliki bukti service level aktual dari perhitungan tersebut.

**Apakah usulan pesan optimal?** Belum. Definisi usulan menambah kebutuhan selama periode review, tetapi safety stock hanya melindungi variasi lead time. MOQ, kemasan, anggaran, ruang, masa simpan, dan lead time bervariasi belum masuk. Saya memakainya sebagai daftar tinjauan, bukan instruksi beli otomatis.

**Apa arti hari stockout?** Hari dengan saldo akhir nol per SKU. Saya tidak menyebutnya lost sales dan tidak menganggapnya seluruh durasi kehabisan intrahari. Total adalah SKU-hari, bukan tanggal unik.

**Bagaimana saya tahu angka benar?** Saya menghitung ulang file dengan LibreOffice dan mencocokkan seluruh kisi harian serta KPI per SKU dengan pandas dari CSV mentah. Saya menunjukkan VERIFIKASI.md, rekonsiliasi, dan skrip; saya tidak mengandalkan screenshot sebagai bukti perhitungan.

**Apa langkah berikutnya pada penerapan nyata?** Saya meminta data kebutuhan yang tidak terpenuhi, jadwal pemasok, biaya dan batas pemesanan, menguji pola permintaan, lalu menguji kebijakan pada data terpisah. Saya tidak mengarang penghematan, peningkatan layanan, atau dampak operasional dari proyek simulasi ini.

Angka hasil yang saya sebut harus diambil dari [temuan dan rekomendasi](temuan_rekomendasi.md) serta bukti verifikasi. Jika parameter berubah, saya menyebutnya skenario baru dan tidak mencampurnya dengan hasil awal.
