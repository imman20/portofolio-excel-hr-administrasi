# Pemetaan Karier: Kemampuan Excel x Bidang Kerja

Tabel ini menjawab pertanyaan "modul mana yang harus saya kuasai dulu untuk bidang X?" dan menghubungkan setiap bidang dengan proyek portofolio di repo ini. Semua proyek yang dirujuk memakai data simulasi berlabel jelas; ini materi latihan, bukan pengalaman kerja.

Skala prioritas: **3** = wajib dikuasai sebelum melamar, **2** = sangat membantu, **1** = pelengkap, **-** = jarang dipakai.

## Matriks prioritas modul per bidang

| Modul | Data Analyst | Finance / Akuntansi | Marketing | Operations / Supply Chain | HR / Administrasi | Risk / Audit Internal | Riset / Akademik | Quality Control |
|---|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|
| 01 Dasar Excel | 3 | 3 | 3 | 3 | 3 | 3 | 3 | 3 |
| 02 Struktur data yang benar | 3 | 3 | 3 | 3 | 3 | 3 | 3 | 3 |
| 03 Formula dasar | 3 | 3 | 3 | 3 | 3 | 3 | 3 | 3 |
| 04 Logika dan agregasi bersyarat | 3 | 3 | 3 | 3 | 3 | 3 | 2 | 3 |
| 05 Lookup dan penggabungan referensi | 3 | 3 | 2 | 3 | 3 | 3 | 2 | 2 |
| 06 Cleaning dan validasi | 3 | 2 | 2 | 2 | 3 | 3 | 3 | 3 |
| 07 Tanggal dan laporan berkala | 3 | 3 | 2 | 3 | 3 | 2 | 1 | 2 |
| 08 PivotTable dan PivotChart | 3 | 3 | 3 | 3 | 3 | 3 | 2 | 2 |
| 09 Visualisasi dan dashboard | 3 | 2 | 3 | 2 | 2 | 2 | 2 | 2 |
| 10 Power Query | 3 | 2 | 2 | 2 | 2 | 2 | 1 | 1 |
| 11 Data Model dan Power Pivot | 2 | 2 | 1 | 2 | 1 | 1 | - | - |
| 12 Statistika terapan | 2 | 1 | 2 | 1 | 1 | 2 | 3 | 3 |
| 13 What-if dan optimasi | 1 | 3 | 1 | 3 | 1 | 2 | 1 | 1 |
| 14 Otomasi lanjutan | 1 | 2 | 1 | 1 | 2 | 1 | - | 1 |
| 15 Praktik kerja profesional | 3 | 3 | 2 | 2 | 3 | 3 | 2 | 2 |

## Rincian per bidang

### Data Analyst

- **Modul paling penting**: 02, 04, 05, 06, 08, 10, 09, 15.
- **Contoh tugas nyata**: menggabungkan ekspor transaksi dari beberapa cabang, membersihkan ID yang tidak konsisten, membuat rekap mingguan dengan PivotTable, menyusun dashboard KPI yang di-refresh dari Power Query, mendokumentasikan definisi setiap metrik.
- **Proyek portofolio relevan**: `../proyek/01_sales_profitability` (KPI penjualan, margin, Pareto pelanggan, dashboard), `../proyek/03_inventory_operations` (KPI operasional dari ±9.000 baris pergerakan stok), `../proyek-hr/08_hr-dashboard-analitik`.
- **Yang membedakan kandidat**: bisa menjelaskan *mengapa* angka di dashboard benar (rekonsiliasi di sheet `Checks`), bukan hanya bisa membuatnya.

### Finance / Akuntansi

- **Modul paling penting**: 03, 04, 05, 07, 08, 13, 15.
- **Contoh tugas nyata**: rekonsiliasi bank vs buku besar, laporan laba rugi bulanan dari jurnal (SUMIFS per akun dan periode), analisis varians anggaran vs realisasi, proyeksi arus kas dengan skenario (Scenario Manager), titik impas dengan Goal Seek, jadwal amortisasi.
- **Proyek portofolio relevan**: `../proyek/01_sales_profitability` (laba kotor dan margin per kategori, pencapaian target), modul 13 (titik impas dan optimasi), `../proyek-hr/04_penggajian-dan-slip-gaji`.
- **Yang membedakan kandidat**: sel input terpisah dari rumus, proteksi sheet, dan jejak audit (Trace Precedents) yang rapi; auditor akan membukanya.

### Marketing

- **Modul paling penting**: 04, 08, 09, 12, 06.
- **Contoh tugas nyata**: rekap performa kampanye per kanal, funnel konversi, segmentasi pelanggan sederhana (RFM dengan COUNTIFS/SUMIFS dan PERCENTILE), analisis survei kepuasan dan NPS, uji A/B sederhana dengan T.TEST, dashboard bulanan untuk manajemen.
- **Proyek portofolio relevan**: `../proyek/02_survei_kepuasan_pelanggan` (NPS, skor per dimensi, uji beda antar kanal, korelasi dengan kepuasan), `../proyek/01_sales_profitability` (musiman dan segmen pelanggan).
- **Yang membedakan kandidat**: tidak membaca korelasi sebagai sebab-akibat, dan mengerti batas sampel survei non-acak (modul 12 bagian keterbatasan).

### Operations / Supply Chain

- **Modul paling penting**: 04, 05, 07, 08, 13, 11.
- **Contoh tugas nyata**: kartu stok dan stok akhir per SKU, perhitungan hari persediaan, reorder point dan safety stock, lead time supplier, klasifikasi ABC, alokasi produksi dengan Solver, penjadwalan staf minimal, laporan OTIF (on time in full).
- **Proyek portofolio relevan**: `../proyek/03_inventory_operations` (seluruh KPI persediaan, ABC, usulan pesan), modul 13 (alokasi produksi dua produk dengan kendala jam mesin dan bahan baku), modul 11 (tabel fakta pergerakan stok dengan dimensi SKU dan supplier).
- **Yang membedakan kandidat**: menuliskan asumsi (misalnya permintaan normal dan independen pada safety stock) dan periode setiap KPI.

### HR / Administrasi

- **Modul paling penting**: 01, 02, 05, 06, 07, 08, 14, 15.
- **Contoh tugas nyata**: database karyawan dengan validasi input (NIK, tanggal masuk, departemen), rekap absensi dan lembur dari mesin absensi, perhitungan cuti dan masa kerja (DATEDIF, NETWORKDAYS), penggajian dengan lookup golongan, tracker rekrutmen, dashboard turnover, surat massal dan arsip, proteksi data pribadi sebelum dibagikan.
- **Proyek portofolio relevan**: `../proyek-hr/01_database-karyawan`, `02_rekap-absensi-bulanan`, `03_administrasi-surat-dan-atk`, `04_penggajian-dan-slip-gaji`, `05_pengelolaan-cuti`, `06_rekrutmen-tracker`, `07_penilaian-kinerja-kpi`, `08_hr-dashboard-analitik`, `09_analisis-survei-engagement-statistik`, `10_perencanaan-tenaga-kerja-dan-otomasi`.
- **Yang membedakan kandidat**: Data Validation yang mencegah salah ketik sejak awal (modul 06), dan kebiasaan menghapus data pribadi serta memeriksa Document Inspector sebelum mengirim file (modul 15).

### Risk / Audit Internal

- **Modul paling penting**: 04, 05, 06, 15, 08, 12.
- **Contoh tugas nyata**: rekonsiliasi dua sumber (sistem vs laporan manual) dan klasifikasi penyebab selisih, pengujian duplikat transaksi dan nilai di luar batas (outlier IQR), sampling transaksi untuk pengujian, audit formula pada workbook yang diwarisi (Trace Precedents, Error Checking), penilaian risiko dengan skor dan skenario.
- **Proyek portofolio relevan**: modul 15 (rekonsiliasi dan audit formula), modul 06 (duplikat dan outlier), sheet `Checks` pada ketiga proyek di `../proyek/`.
- **Yang membedakan kandidat**: tidak memakai IFERROR untuk menyembunyikan kesalahan data; setiap `#N/A` dicari penyebabnya (modul 05).

### Riset / Akademik

- **Modul paling penting**: 02, 06, 12, 09, 03.
- **Contoh tugas nyata**: tabulasi kuesioner dari format lebar ke panjang, statistik deskriptif per kelompok, interval kepercayaan, uji beda dua kelompok (uji t Welch), korelasi dan regresi sederhana dengan LINEST, ukuran efek (Cohen's d), visualisasi distribusi (histogram, scatter) untuk laporan.
- **Proyek portofolio relevan**: `../proyek/02_survei_kepuasan_pelanggan` (codebook, reliabilitas Cronbach's alpha, uji t, regresi dengan peringatan), modul 12, `../proyek-hr/09_analisis-survei-engagement-statistik`.
- **Yang membedakan kandidat**: menuliskan tipe uji dan ekor yang dipakai, asumsi, dan keterbatasan sampel; tidak melaporkan p-value tanpa ukuran efek.

### Quality Control / Produksi

- **Modul paling penting**: 04, 06, 12, 03, 09.
- **Contoh tugas nyata**: pencatatan hasil inspeksi per shift, proporsi cacat per lini (COUNTIFS), batas kendali sederhana (rata-rata ± 3 sigma dengan AVERAGE dan STDEV.S), peta kendali dengan line chart, Pareto penyebab cacat, Conditional Formatting untuk nilai di luar spesifikasi, uji beda rata-rata sebelum dan sesudah perbaikan.
- **Proyek portofolio relevan**: modul 12 (distribusi, uji t sebelum-sesudah, ukuran efek), modul 06 (Conditional Formatting dan Data Validation), modul 09 (chart dan dashboard).
- **Yang membedakan kandidat**: memahami bahwa batas kendali dan batas spesifikasi berbeda, serta mengerti arti variasi acak sebelum menyimpulkan ada masalah.

## Urutan belajar yang saya sarankan per bidang

| Bidang | Urutan modul minimal (sebelum melamar) | Lanjutan |
|---|---|---|
| Data Analyst | 01 02 03 04 05 06 07 08 09 10 15 | 11 12 14 13 |
| Finance / Akuntansi | 01 02 03 04 05 07 08 13 15 | 09 10 14 12 |
| Marketing | 01 02 03 04 08 09 12 | 05 06 10 15 |
| Operations / Supply Chain | 01 02 03 04 05 07 08 13 | 11 09 15 10 |
| HR / Administrasi | 01 02 03 04 05 06 07 08 15 | 14 09 12 |
| Risk / Audit Internal | 01 02 03 04 05 06 08 15 | 12 07 13 |
| Riset / Akademik | 01 02 03 06 12 09 | 04 05 08 |
| Quality Control | 01 02 03 04 06 12 09 | 08 07 |

## Cara memakai dokumen ini saat melamar

1. Pilih bidang; kerjakan urutan modul minimalnya.
2. Buka proyek portofolio yang dirujuk, kerjakan versi `latihan_*.xlsx`-nya, lalu bandingkan dengan solusi.
3. Siapkan penjelasan 3 menit per proyek: masalah bisnis, pendekatan, keputusan teknis, cara memverifikasi, keterbatasan. Panduannya ada di `docs/panduan_wawancara.md` tiap proyek.
4. Saat ditanya "pengalaman Excel", jawab jujur: ini latihan terstruktur dengan data simulasi yang saya bangun dan verifikasi sendiri; tunjukkan file dan `VERIFIKASI.md`-nya.
