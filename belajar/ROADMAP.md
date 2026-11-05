# Roadmap 10 Minggu: Excel dari Pemula sampai Siap Kerja

Jadwal ini saya susun untuk sekitar 7 sampai 9 jam belajar per minggu (1 sampai 1,5 jam per hari kerja, atau dua sesi panjang di akhir pekan). Setiap minggu punya keluaran konkret berupa file yang harus selesai dan kriteria lulus yang bisa diperiksa sendiri dengan membandingkan terhadap `solusi_NN.xlsx` dan `kunci_jawaban.md`. Ini materi latihan dengan data sintetis, bukan pengalaman kerja.

Aturan umum:

- Jangan lanjut ke minggu berikutnya sebelum semua kriteria lulus terpenuhi; lebih baik mundur satu minggu daripada menumpuk lubang.
- Simpan hasil latihan sendiri dengan nama `latihan_NN_saya.xlsx` di folder lokal (bukan di repo) agar bisa dibandingkan dengan solusi tanpa menimpa.
- Setiap akhir minggu, tulis 5 baris catatan: apa yang dipelajari, apa yang masih bingung, satu kesalahan yang dilakukan, cara memeriksanya, dan rencana minggu depan.

| Minggu | Modul | Jam | Keluaran konkret (file yang harus jadi) | Kriteria lulus |
|---|---|---|---|---|
| 1 | 01 Dasar Excel, 02 Struktur data yang benar | 6 | `latihan_01_saya.xlsx` terisi penuh; `latihan_02_saya.xlsx` dengan Excel Table bernama dan tabel "tidy" hasil perombakan | Semua sel kuning modul 01 dan 02 cocok dengan solusi; bisa menjelaskan angka serial tanggal; tabel rapi lolos 5 aturan tidy data di README 02 |
| 2 | 03 Formula dasar, 04 Logika dan agregasi bersyarat | 9 | `latihan_03_saya.xlsx` (tabel harga dengan referensi campuran), `latihan_04_saya.xlsx` (rekap SUMIFS/COUNTIFS per kategori dan bulan) | Rumus bisa disalin ke seluruh tabel tanpa diedit satu per satu; total rekap SUMIFS sama dengan SUM kolom sumber; tidak ada `#VALUE!` |
| 3 | 05 Lookup, 06 Cleaning dan validasi | 11 | `latihan_05_saya.xlsx` (gabungan transaksi + master produk dengan INDEX/MATCH), `latihan_06_saya.xlsx` (data kotor menjadi bersih, duplikat dan outlier teridentifikasi, Data Validation dan Conditional Formatting terpasang) | Tidak ada `#N/A` yang disembunyikan; daftar kode tak ditemukan terlapor eksplisit; jumlah baris bersih sama dengan kunci jawaban; validasi menolak input salah saat diuji |
| 4 | 07 Tanggal dan laporan berkala, 08 PivotTable | 10 | `latihan_07_saya.xlsx` (laporan bulanan 12 baris dengan MoM, YoY, dan hari kerja), PivotTable di `latihan_08_saya.xlsx` dengan Slicer dan Timeline | Angka PivotTable identik dengan sheet `Hasil_Pivot_Diharapkan`; MoM pada bulan pertama kosong (bukan error); NETWORKDAYS memperhitungkan hari libur |
| 5 | 09 Visualisasi dan dashboard | 8 | `latihan_09_saya.xlsx` berisi dashboard satu layar: 4 KPI, 3 chart (line, bar, scatter), sumber dan tanggal pembaruan | Dashboard bisa dibaca dalam 30 detik oleh orang lain; tidak ada chart 3D atau pie beririsan banyak; setiap angka KPI bisa ditelusuri ke rumusnya |
| 6 | 10 Power Query | 7 | `latihan_10_saya.xlsx` dengan query Append 3 CSV cabang, Merge ke master, Unpivot target, dan Refresh berhasil | Jumlah baris sesudah Merge sama dengan sebelum (tidak ada penggandaan); hasil query sama dengan sheet hasil yang diharapkan; skrip M tersimpan di Advanced Editor |
| 7 | 11 Data Model dan Power Pivot, 12 Statistika terapan (bagian deskriptif, distribusi, interval kepercayaan) | 9 | `latihan_11_saya.xlsx` dengan Data Model (fakta + 3 dimensi + kalender) dan 5 measure DAX; `latihan_12_saya.xlsx` bagian 1 sampai 3 | Measure DAX menghasilkan angka yang sama dengan sheet pembanding; CI 95 % dihitung dengan CONFIDENCE.T dan ditafsirkan benar |
| 8 | 12 Statistika terapan (korelasi, regresi, uji t, peramalan, ukuran efek), 13 What-if dan optimasi | 10 | `latihan_12_saya.xlsx` lengkap; `latihan_13_saya.xlsx` dengan Goal Seek, Data Table 1 dan 2 variabel, Scenario Manager, dan Solver | Interpretasi koefisien regresi dan R² ditulis dengan keterbatasannya; solusi Solver sama dengan kunci jawaban (dihitung terpisah dengan program linear) |
| 9 | 14 Otomasi lanjutan, 15 Praktik kerja profesional | 10 | Makro `RapikanLaporan` berjalan pada file latihan 14; Office Script dicoba di Excel web; `latihan_15_saya.xlsx` rekonsiliasi dua sumber dan template Input-Proses-Output terproteksi | Semua selisih rekonsiliasi punya penyebab yang tercatat; sheet terproteksi hanya membuka sel input; Document Inspector bersih; checklist handover terisi |
| 10 | Proyek portofolio | 8 | Mengerjakan `../proyek/01_sales_profitability/workbook/latihan_sales_profitability.xlsx` dari data mentah sampai dashboard, lalu menuliskan temuan sendiri sebelum membaca `docs/temuan_rekomendasi.md` | Angka KPI saya sama dengan workbook solusi proyek; bisa menjelaskan proyek dalam 3 menit (masalah, pendekatan, keputusan teknis, keterbatasan) tanpa mengarang pengalaman |

Total sekitar 88 jam termasuk pengulangan.

## Setelah minggu ke-10

- Kerjakan `../proyek/02_survei_kepuasan_pelanggan` (menguatkan modul 06 dan 12) dan `../proyek/03_inventory_operations` (menguatkan modul 07, 11, 13).
- Pilih satu bidang di `PEMETAAN-KARIER.md` dan ulangi modul yang ditandai paling penting untuk bidang itu dengan dataset buatan sendiri.
- Ulangi modul 15 setiap kali akan menyerahkan workbook kepada orang lain.

## Jika waktu terbatas (jalur 4 minggu)

Minggu 1: modul 01, 02, 03. Minggu 2: modul 04, 05, 06. Minggu 3: modul 07, 08, 09. Minggu 4: modul 15 lalu proyek 01. Modul 10 sampai 14 dikerjakan setelahnya sesuai kebutuhan bidang.
