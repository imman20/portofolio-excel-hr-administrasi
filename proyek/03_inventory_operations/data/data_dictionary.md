# Kamus data inventori

Saya memakai tanggal ISO `YYYY-MM-DD` pada CSV, ID bertipe teks, qty integer dalam unit barang, dan nilai rupiah tanpa desimal pada data mentah. Semua tabel adalah simulasi sintetis. Nilai hasil di workbook selalu dihitung dengan rumus.

| File | Grain dan kunci | Kolom dan makna |
|---|---|---|
| raw/master_sku.csv | Satu SKU, kunci sku | nama_sku: label simulasi; supplier_id: pemasok utama; kategori: kelompok barang; hpp_satuan: HPP tetap rupiah/unit; stok_awal: unit sebelum transaksi awal tahun; tanggal_stok_awal: tanggal saldo pembuka |
| raw/supplier.csv | Satu supplier, kunci supplier_id | nama_supplier: label sintetis; lead_time_nominal: hari kalender untuk fallback jika SKU belum punya PO selesai |
| raw/pergerakan.csv | Satu gerakan SKU, kunci movement_id | tanggal: hari transaksi; sku: relasi master; jenis: MASUK/KELUAR/PENYESUAIAN; qty: unit positif untuk masuk/keluar, bertanda untuk penyesuaian; po_id: wajib untuk MASUK, kosong untuk jenis lain |
| raw/purchase_order.csv | Satu PO satu SKU, kunci po_id | sku dan supplier_id: relasi master; tanggal_pesan: awal tunggu; tanggal_terima: selesai diterima, kosong jika terbuka; qty_pesan: unit pesanan; qty_terima: unit diterima; status: SELESAI/TERBUKA |
| processed/stok_harian.csv | Satu tanggal × SKU, kunci gabungan | stok_awal_hari, masuk, keluar, penyesuaian, stok_akhir: unit; stockout: 1 jika stok akhir nol, selain itu 0 |
| processed/kpi_referensi.csv | Satu SKU | Hasil perhitungan pandas independen; nama KPI sama dengan Calculation, dengan supplier_id untuk kolom supplier |
| processed/rekonsiliasi.csv | Satu pembanding utama | pemeriksaan, lokasi sel, pandas, workbook, selisih, status |
| processed/rekonsiliasi_detail.csv | Satu file × SKU × KPI | pandas dan workbook: nilai/label yang dicocokkan; selisih numerik atau nol untuk label cocok; status PASS |
| processed/ringkasan.json | Satu hasil verifikasi | Tanggal, volume, status/ABC, lima usulan, turnover, stockout, jumlah rumus/error, dan hash CSV |

Tidak ada penerimaan parsial, pesanan dibatalkan, reservasi, backorder, transfer antargudang, nomor lot, atau kedaluwarsa. KELUAR adalah barang yang benar-benar dikeluarkan, bukan seluruh permintaan sebelum pembatasan stok. PO terbuka mengurangi usulan tetapi tidak mengubah saldo. Penilaian memakai HPP tetap; perubahan harga tidak dimodelkan.

CSV processed adalah bukti referensi dan keluaran verifikasi; workbook tidak mengimpor KPI jadi dari sana. [Definisi KPI](../docs/definisi_kpi.md) menjelaskan periode, satuan, pembagi nol, status, dan asumsi statistik.
