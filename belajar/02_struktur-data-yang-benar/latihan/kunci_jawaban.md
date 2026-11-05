# Kunci jawaban modul 02

Saya memakai CSV sumber dan pandas untuk pembanding berikut. Rumus tetap hidup pada solusi.

Grain adalah satu transaksi produk per tanggal dan cabang. ID teks berbeda dari angka berformat; Table mempunyai nama dan perluasan baris.

| Pemeriksaan | Sel solusi | Nilai pandas |
|---|---|---:|
| COUNT tanggal | Latihan!C5 | 240 |
| COUNT cabang_id (teks) = 0 | Latihan!C6 | 0 |
| Total penjualan (structured ref) | Latihan!C18 | 51749000 |
| Penjualan Surabaya | Latihan!C20 | 23882500 |
| Surabaya Jan (SUMIFS tanggal) | Latihan!C26 | 3779000 |
| Selisih total vs Laporan_Lama = 0 | Latihan!I29 | 0 |
| Selisih long vs wide = 0 | Latihan!D52 | 0 |
| ID 001 ditemukan di Data | ID_Rusak!E4 | 106 |

Tabel long memuat satu baris per cabang dan bulan; totalnya harus sama dengan matriks sumber. Saya tidak memasukkan subtotal sebagai transaksi. Percobaan menambah baris dilakukan pada salinan dan dibatalkan sebelum rekonsiliasi awal.
