# Changelog

Semua perubahan penting pada repositori ini dicatat di sini. Format mengikuti Keep a Changelog dan versi mengikuti Semantic Versioning.

## [1.3.0] - 2026-11-05

### Ditambahkan
- Proyek simulasi `proyek/03_inventory_operations/`: 60 SKU, 8 supplier, 9.087 pergerakan, 500 PO, dan 21.960 baris stok harian dari generator berseed tetap.
- Workbook latihan dan solusi dengan kisi SUMIFS harian, saldo harian, lead time PO, safety stock, reorder point, usulan pesan, turnover, ABC, dan Checks.
- Dokumentasi KPI beserta asumsi/periode, kamus data, instruksi latihan, temuan, pemeriksaan kualitas, panduan wawancara, serta preview nyata dari PDF/PNG.
- Recalculation LibreOffice dan rekonsiliasi pandas/scipy per SKU serta seluruh kisi harian pada kedua workbook.

### Diperbarui
- Status dan tautan inventori pada README utama dan README belajar. Kedua proyek analisis sebelumnya, proyek HR, dan modul dasar dipertahankan.

## [1.2.0] - 2026-11-05

### Ditambahkan
- Proyek simulasi `proyek/02_survei_kepuasan_pelanggan/`: CSV mentah 420 baris, 414 ID unik, workbook latihan dan solusi dengan rumus pembersih, distribusi item, CI, alpha, NPS, Welch, CORREL, dan LINEST deskriptif.
- Codebook, kamus data, dokumentasi kualitas dan statistik, panduan latihan/wawancara, dashboard native, serta preview dari render Report.
- Rekonsiliasi 11.178 sel bersih terhadap CSV referensi dan pembanding statistik pandas/scipy, audit hitungan masalah sebelum/sesudah duplikasi, serta pengujian negatif pada salinan.

### Diperbaiki
- Parser tanggal bahan survei yang menukar hari/bulan pada 133 tanggal ISO di referensi; CSV mentah dipertahankan. Daftar koreksi tersedia pada proyek survei.
- Cakupan catatan masalah dibuat eksplisit: 174 sel kosong dan 15 angka berspasi pada ID unik; 175 dan 16 pada CSV akhir yang mencakup duplikat.

### Diperbarui
- Status dan tautan proyek survei pada README utama dan README belajar. Proyek Sales, HR, serta modul dasar dipertahankan.

## [1.1.0] - 2026-11-05

### Ditambahkan
- Proyek simulasi `proyek/01_sales_profitability/`: generator berseed tetap, CSV, workbook latihan dan solusi dengan rumus hidup, dashboard, serta Checks.
- Dokumentasi KPI, cleaning, temuan, latihan, pemeriksaan kualitas, panduan wawancara, PivotTable manual, dan skrip Power Query.
- Recalculation LibreOffice, rekonsiliasi pandas tahunan dan bulanan, pemeriksaan setiap transaksi, serta preview asli dari PDF/PNG.

### Diperbarui
- Tabel status dan tautan proyek Sales pada README utama serta hubungan proyek di README belajar.

## [1.0.0] - 2025-11-05

### Ditambahkan
- Jalur belajar Excel di `belajar/`: README, roadmap 10 minggu, pemetaan kemampuan ke bidang kerja, dan modul 01 Dasar Excel lengkap dengan dataset, latihan, solusi, dan verifikasi.
- Sepuluh proyek HR dan administrasi dipindahkan ke `proyek-hr/` beserta README daftar proyek.
- Sepuluh proyek Excel bertingkat (pemula, menengah, lanjut) untuk HR, administrasi, dan analisis data, masing-masing dengan `TUTORIAL.md`, `latihan.md`, dan screenshot.
- Skrip pembangun `proyek-hr/tools/build_all.py` berbasis openpyxl dengan data sintetis berseed tetap, penghitungan ulang LibreOffice, verifikasi nol error, dan pembuatan screenshot.
- Dokumentasi pendukung: `proyek-hr/docs/glosarium_rumus.md`, `proyek-hr/docs/standar-data-hr.md`, `proyek-hr/tools/verifikasi.md`.
- Tiga modul VBA (`10_perencanaan-tenaga-kerja-dan-otomasi/vba/`) beserta panduan pemasangan.
