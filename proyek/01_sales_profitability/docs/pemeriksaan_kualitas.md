# Pemeriksaan kualitas

Saya menggunakan dua lapis pemeriksaan: kontrol hidup di sheet Checks dan perhitungan independen pandas dari CSV. Status PASS pada kontrol internal saja tidak membuktikan definisi KPI benar; rumus yang sama-sama salah masih dapat menghasilkan total yang cocok.

## Sumber, kunci, dan relasi

Verifier memeriksa keunikan order_id, retur_id, serta seluruh kunci master sebelum join. Pasangan bulan × region pada target harus unik, lengkap, dan positif. Semua ID transaksi wajib ditemukan tepat di master. Region order harus cocok dengan region pelanggan dan salesperson. Tanggal order tidak boleh mendahului tanggal bergabung; tanggal retur tidak boleh mendahului order asal.

Saya memeriksa nilai kosong pada input, qty bulat positif, harga positif, diskon dalam rentang generator, kuantitas retur positif dan tidak melampaui qty order. Periode transaksi dan retur harus berada di tahun analisis. Tidak ada imputasi atau penghapusan baris untuk membuat pemeriksaan lolos.

## Workbook

Verifier membuka setiap file dengan `data_only=False` untuk menghitung rumus dan memeriksa daftar fungsi yang diizinkan, lalu dengan `data_only=True` untuk memeriksa cache hasil recalc. Semua sel error harus nol. Ukuran setiap file harus di bawah batas yang ditentukan, sheet wajib dan Freeze Panes harus ada, dan solusi harus memiliki chart laporan yang lengkap.

Sheet Checks membandingkan agregasi transaksi dengan SUMPRODUCT independen dan rekap bulanan, region, kategori, pelanggan, serta memeriksa duplikat, relasi, rentang input, kontribusi total, dan kumulatif Pareto. Selisih dibulatkan dua desimal dan harus nol. Pada latihan, kolom yang sengaja belum diisi menghasilkan status BELUM DIISI, bukan PASS palsu.

## Rekonsiliasi independen

Saya menghitung omzet bruto, penjualan bersih, HPP, dan laba kotor dari CSV dan master melalui pandas. Keempat nilai dicocokkan untuk setahun dan setiap bulan, kemudian diperluas ke region, kategori, seluruh pelanggan berperingkat, dan contoh baris latihan. Seluruh baris transaksi solusi juga dicocokkan pada komponen omzet, diskon, neto, retur, HPP, dan laba.

Selisih uang harus nol pada dua desimal; selisih rasio harus nol pada dua belas desimal agar sisa representasi biner tidak dianggap kesalahan bisnis. Saya tidak memakai toleransi persentase yang dapat menyembunyikan selisih uang besar. Hasil perbandingan dan hash input disimpan di `data/processed/`, sedangkan tabel yang dapat dibaca ada di [VERIFIKASI.md](../VERIFIKASI.md).

## Reproduksi

```bash
python -m pip install -r requirements.txt
python tools/generate_data.py
python tools/build_workbook.py
python tools/verify.py
```

LibreOffice headless dan pdftoppm harus tersedia pada PATH. Preview dibentuk dari render Report ke PDF lalu PNG. Menjalankan builder menulis ulang workbook; verifikasi wajib diulang setelah setiap pembangunan. Bila LibreOffice tidak tersedia, proses berhenti dan file baru belum boleh disebut terverifikasi.

Verifier selalu menghitung ulang workbook. Sebelum membuat ulang hasil, saya memeriksa tanggal laporan lama agar tidak menganggap dokumen verifikasi sebelumnya berlaku untuk file baru.

Saya juga memeriksa tautan Markdown yang diubah dan memastikan folder proyek HR serta modul 01 tidak berubah. Pemeriksaan tampilan dilakukan pada preview; eksekusi PivotTable/Power Query dan interaksi lintas versi Excel tetap langkah manual.

Saya menguji verifier pada salinan input yang sengaja dirusak: kunci produk duplikat, qty retur melebihi pembelian, serta pasangan target duplikat. Ketiganya ditolak sebelum hasil dapat dinyatakan lolos. CSV proyek tetap dipertahankan dan hasil generator ulang identik dengan sumber.
