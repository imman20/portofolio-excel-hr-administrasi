# Cleaning dan transformasi

Saya membangun data sintetis dari nol dengan `numpy.random.default_rng(202403)` di [generator](../tools/generate_data.py). Data tidak memuat identitas pribadi atau perusahaan nyata. Tujuannya simulasi stok yang dapat direkonsiliasi; saya tidak menyisipkan masalah kualitas seperti pada proyek survei.

Saya menetapkan kunci SKU dan supplier sebagai teks, tanggal CSV ISO, HPP tetap, serta saldo awal sebelum transaksi hari pertama. Generator membuat PO selesai dan PO terbuka, lalu menerima PO selesai sebagai MASUK dengan referensi po_id. Satu PO hanya satu SKU dan satu penerimaan penuh. Lead time berasal dari selisih tanggal, bukan angka acak yang ditulis sebagai hasil jadi di workbook.

Saya menjalankan simulasi kalender per SKU. Pada hari yang sama penerimaan dicatat sebelum penyesuaian, lalu pengeluaran dibatasi stok tersedia. Penyesuaian negatif juga dibatasi stok tersedia. Urutan ini mencegah stok negatif pada generator, tetapi CSV hanya menyimpan tanggal; workbook merekonsiliasi saldo akhir harian dan tidak mengaudit waktu intrahari. Transaksi KELUAR yang dicatat adalah pemenuhan, sehingga kebutuhan yang tidak terpenuhi tidak muncul dalam dataset.

Saya menyimpan `data/processed/stok_harian.csv` sebagai referensi pembangkit, lalu verifier membangun ulang saldo dari CSV mentah dengan pivot dan cumulative sum pandas. Verifier tidak memakai saldo referensi sebagai masukan untuk menghitung KPI. Ia membandingkan masuk, keluar, penyesuaian, saldo awal hari, saldo akhir hari, dan indikator stockout dengan CSV referensi tersebut.

Di workbook, Raw_Data mempertahankan kolom CSV dan menambahkan qty bertanda, lookup HPP, serta HPP keluar dalam rumus. Master memuat parameter, master SKU, dan supplier. Purchase_Order mempertahankan PO dan menambahkan lead time serta outstanding berumus. Permintaan_Harian memuat kalender penuh, SUMIFS pengeluaran, dan saldo berantai. Calculation membangun KPI dengan rumus, Report merujuk Calculation, dan Checks melakukan rekonsiliasi internal. Nilai agregat tidak ditempel sebagai angka statis di workbook.

Sebelum analisis, saya memeriksa kunci unik, relasi SKU dan supplier, jenis transaksi, qty integer, tanggal, penerimaan yang cocok dengan PO, serta saldo nonnegatif. Verifier berhenti saat pemeriksaan gagal; tidak ada penghapusan baris atau IFERROR menyeluruh yang menutupi pelanggaran. CSV menggunakan koma sebagai delimiter; menu impor perlu mempertahankan ID sebagai Text dan tanggal sebagai tanggal.

Untuk membuka CSV manual: **Data > From Text/CSV**, pilih UTF-8 dan delimiter comma, lalu periksa tipe kolom di pratinjau sebelum **Load**. Power Query dapat dipakai sebagai jalur impor manual, tetapi query tidak tertanam dan tidak diperlukan untuk workbook yang sudah berisi data. Saya tidak mengklaim alur refresh otomatis tersedia.
