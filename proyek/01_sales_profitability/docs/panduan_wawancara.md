# Panduan saya menjelaskan proyek

## Pembuka singkat

“Saya membuat proyek portofolio dengan data penjualan sintetis untuk melatih analisis profitabilitas. Saya memisahkan transaksi, retur, master, perhitungan, dashboard, dan pemeriksaan. Saya menelusuri penjualan sampai laba kotor, membandingkan target per region, lalu memeriksa hasil Excel dengan perhitungan pandas yang independen.”

Saya menyebutnya simulasi dan latihan, bukan pengalaman kerja, proyek klien, atau pencapaian bisnis. Angka yang saya kutip berasal dari [ringkasan hasil](temuan_rekomendasi.md), bukan angka yang saya ingat atau perkirakan.

## Urutan demonstrasi

1. Saya membuka Report untuk menjelaskan pertanyaan bisnis, periode, dan satuan.
2. Saya menelusuri satu kartu KPI ke Calculation, lalu ke Raw_Data. Saya menunjukkan bahwa penjualan sudah mengurangi diskon dan retur.
3. Saya membuka Retur dan Master untuk memperlihatkan grain, kunci, exact match, serta alasan HPP barang retur dikurangi.
4. Saya membandingkan margin antar kategori dan pencapaian region tanpa menyimpulkan sebab-akibat dari pola yang dibangkitkan generator.
5. Saya membuka Checks dan VERIFIKASI untuk menunjukkan perbedaan kontrol rumus internal dengan pembanding independen dari CSV.

## Keputusan teknis yang bisa saya jelaskan

- **Mengapa SUMIFS dan INDEX/MATCH?** Saya ingin workbook dapat dipakai di Excel 2016. SUMIFS mempertahankan rumus hidup; lookup exact match menjaga hubungan produk dan pelanggan. Fungsi baru seperti XLOOKUP tidak diperlukan untuk hasil yang sama.
- **Mengapa agregasi retur sebelum join?** Satu order dapat mempunyai beberapa retur. Saya menjumlahkan qty retur menurut order agar transaksi tidak tergandakan saat digabung.
- **Mengapa margin bukan rata-rata?** Ukuran transaksi berbeda. Saya membagi total laba dengan total penjualan; rata-rata margin baris memberi bobot sama pada transaksi kecil dan besar.
- **Mengapa target tidak ada pada setiap order?** Grain target adalah bulan × region. Menempel dan menjumlahkannya di semua order akan menggandakan target. Saya mengagregasi dari tabel target terpisah.
- **Bagaimana menangani pembagi nol?** Retur penuh ditampilkan dengan margin nol sebagai konvensi teknis; pencapaian tanpa target dan MoM tanpa pembanding ditampilkan `-`. Saya menjelaskan maknanya, bukan mengklaim semua rasio nol bermakna bisnis.
- **Bagaimana memutus seri Pareto?** Saya memakai customer_id menaik setelah mengurutkan penjualan menurun. Saya memakai peringkat bulat, sehingga nilai uang tidak diubah untuk membedakan pelanggan.
- **Mengapa tidak memakai PivotTable langsung?** File yang dibuat openpyxl menyediakan chart dan rumus, sedangkan PivotTable dan Power Query dibuat manual memakai panduan. Saya tidak mengklaim fitur itu sudah tertanam.

## Keterbatasan yang saya nyatakan

Data, nama pelanggan, aktivitas, dan target adalah sintetis. Target dibentuk dari realisasi simulasi, jadi pencapaian bukan evaluasi independen. HPP konstan, retur diasumsikan pulih seluruhnya, dan belum ada biaya operasional, pajak, arus kas, atau persediaan. Pola musim dan margin memang dirancang; rekomendasi hanya ilustrasi keputusan yang harus diuji lagi pada data nyata.

Recalculation dan kontrol numerik dijalankan di LibreOffice. Saya belum menguji interaksi file pada semua versi Excel atau menjalankan skrip M di Power Query. Saya dapat menunjukkan bukti yang tersedia dan menyebutkan apa yang belum diuji.

## Bila diminta mengembangkan proyek

Saya akan meminta definisi order multi-produk, kebijakan cutoff retur, biaya logistik, dan target yang disepakati sebelum periode berjalan. Saya akan menetapkan kunci order_line_id, memisahkan tanggal order dan tanggal pencatatan retur sesuai kebutuhan, lalu menambahkan pengujian agar join tidak menggandakan nilai. Saya tidak akan menyebut hasil hipotetis ini sebagai peningkatan laba yang sudah terjadi.
