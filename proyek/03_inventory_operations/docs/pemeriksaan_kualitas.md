# Pemeriksaan kualitas

Saya memisahkan pemeriksaan data, rumus, dan interpretasi. [Verifier](../tools/verify.py) membangun ulang perhitungan pandas langsung dari CSV mentah, menghitung ulang kedua workbook dengan LibreOffice headless, membaca cache melalui openpyxl `data_only=True`, lalu menghentikan proses bila satu pemeriksaan gagal.

## Data dan hubungan antartabel

Saya memeriksa keunikan movement_id, po_id, sku, dan supplier_id; relasi supplier pada master dan PO; semua SKU gerak/PO dikenal master; qty integer; transaksi masuk/keluar positif; jenis transaksi dikenal; serta tanggal berada pada periode simulasi. PO selesai harus diterima penuh dengan tanggal terima tidak mendahului tanggal pesan. PO terbuka tidak punya tanggal terima dan qty terima nol.

Setiap penerimaan MASUK dipasangkan satu-ke-satu dengan PO selesai, lalu SKU, qty, dan tanggal harus sama. Transaksi selain MASUK tidak memakai po_id. Penyesuaian dapat bertanda negatif tetapi seluruh saldo akhir harian harus nonnegatif. Saya tidak menganggap rekonsiliasi total cukup: selisih antar-SKU dapat saling menutupi jika hanya total diperiksa.

## Buku harian dan KPI

CSV stok_harian dari generator dibandingkan dengan pivot/cumulative sum independen untuk saldo awal hari, masuk, keluar, penyesuaian, saldo akhir, dan indikator stockout. Kedua workbook kemudian dibandingkan per sel untuk seluruh kisi keluar dan saldo harian. Setiap KPI solusi dicocokkan per SKU, termasuk sigma, lead time, safety stock, on_order, usulan, turnover, status, dan ABC. Urutan lima usulan terbesar juga harus sama, termasuk aturan tie.

Selisih integer/rupiah dan kisi harian harus nol tepat. Rasio dan statistik harus berselisih nol setelah pembulatan enam desimal untuk mengakomodasi representasi biner. Toleransi ini adalah ketelitian numerik, bukan izin mengabaikan selisih unit. Hasil utama ditulis ke [rekonsiliasi](../data/processed/rekonsiliasi.csv); seluruh KPI tersedia pada [detail](../data/processed/rekonsiliasi_detail.csv).

## Workbook

Saya menghitung seluruh rumus dan memeriksa semua sel bertipe error, termasuk #REF!, #NAME?, #DIV/0!, #VALUE!, #N/A, #NUM!, dan error lain. Cache formula tidak boleh kosong. Nama fungsi diperiksa terhadap daftar yang diizinkan; rumus array dan referensi workbook eksternal ditolak. Ukuran setiap XLSX harus di bawah 5 MB. Freeze Panes, struktur sheet, chart native, dan sel latihan kosong ikut diperiksa.

Checks melakukan rekonsiliasi internal saldo akhir, saldo harian terakhir, jumlah SKU, kelas/status, HPP, stok negatif harian, panjang kalender, penerimaan PO, urutan ambang ABC, serta kelengkapan jawaban. Semua harus PASS pada solusi. Latihan memiliki satu BELUM DIISI yang disengaja. Checks melengkapi pembanding pandas; ia bukan pengganti audit data dan tidak memeriksa seluruh kemungkinan edit manual.

## Bukti dan batas pemeriksaan

[VERIFIKASI.md](../VERIFIKASI.md) mencatat jumlah rumus, error, ukuran file, pembanding per SKU, dan tanggal pengujian. [Ringkasan JSON](../data/processed/ringkasan.json) menyimpan hash CSV mentah. Preview berasal dari Report yang diekspor ke PDF satu halaman oleh LibreOffice, lalu dikonversi ke PNG dengan pdftoppm. Saya menyembunyikan sheet lain hanya pada salinan render, tanpa menggambar ulang dashboard.

Pemeriksaan ini membuktikan konsistensi implementasi pada data dan parameter awal. Ia tidak membuktikan service level aktual, validitas distribusi normal, optimalitas kebijakan, atau kecukupan stok untuk bisnis nyata. Parameter yang berubah harus diikuti recalc dan pembanding skenario yang sesuai. Tanggal simulasi dibatasi agar kalender tidak diam-diam tidak cocok dengan data.
