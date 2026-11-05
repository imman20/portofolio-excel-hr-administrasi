# Portofolio Excel: Jalur Belajar, Proyek Analisis, dan Proyek HR/Administrasi

Satu repositori untuk tiga hal yang saya pisahkan dengan jelas: **materi belajar Excel** dari nol sampai siap kerja, **proyek portofolio analisis** dengan data simulasi yang bisa dibuka dan diperiksa, dan **proyek tematik HR/administrasi** yang menunjukkan penerapan Excel di pekerjaan kantor.

![Excel](https://img.shields.io/badge/Microsoft%20Excel-2016%2B%20%7C%20365-217346?logo=microsoftexcel&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green)
![Rumus](https://img.shields.io/badge/rumus-hidup%20dan%20terverifikasi-blue)
![Data](https://img.shields.io/badge/data-simulasi%20berlabel-orange)

> **Penting untuk pembaca dan recruiter.** Semua data di repositori ini adalah data sintetis yang saya bangkitkan dengan skrip Python (seed tetap, bisa dibangun ulang). Tidak ada data perusahaan, klien, atau pribadi nyata. Materi di folder `belajar/` adalah latihan, bukan pengalaman kerja. Proyek di `proyek/` dan `proyek-hr/` adalah proyek portofolio simulasi yang saya kerjakan untuk menunjukkan cara saya bekerja: dari data mentah, pembersihan, rumus, pemeriksaan, sampai dashboard dan rekomendasi.

## Daftar Isi

- [Untuk siapa repositori ini](#untuk-siapa-repositori-ini)
- [Tiga bagian repositori](#tiga-bagian-repositori)
- [Cara mulai belajar dari nol](#cara-mulai-belajar-dari-nol)
- [Cara recruiter menilai portofolio ini](#cara-recruiter-menilai-portofolio-ini)
- [Peta belajar](#peta-belajar)
- [Proyek portofolio](#proyek-portofolio)
- [Proyek HR dan administrasi](#proyek-hr-dan-administrasi)
- [Standar file dan cara verifikasi](#standar-file-dan-cara-verifikasi)
- [Batasan tool dan langkah manual](#batasan-tool-dan-langkah-manual)
- [Pengaturan Excel yang perlu diketahui](#pengaturan-excel-yang-perlu-diketahui)
- [Struktur folder](#struktur-folder)
- [Lisensi](#lisensi)
- [Kontak](#kontak)

## Untuk siapa repositori ini

- **Saya sendiri**, sebagai jalur belajar yang terstruktur: lulusan Statistika terbiasa dengan R dan Python, tetapi pekerjaan analis, HR, finance, dan operasional sehari-hari masih berjalan di Excel.
- **Pembelajar lain** yang ingin mulai dari nol dengan dataset, latihan, dan kunci jawaban yang saling konsisten.
- **Recruiter dan hiring manager** yang ingin melihat bukti kemampuan Excel yang bisa dibuka, ditelusuri rumusnya, dan diperiksa hasilnya, bukan sekadar daftar skill di CV.

## Status pengerjaan

| Bagian | Status |
|---|---|
| Jalur belajar: README, roadmap 10 minggu, pemetaan karier | Selesai |
| Modul belajar 01 Dasar Excel | Selesai: README, dataset, workbook latihan dan solusi, kunci jawaban, verifikasi |
| Modul belajar 02 sampai 15 | Menyusul. Cakupan tiap modul sudah tercantum di peta belajar di bawah |
| Proyek HR dan administrasi 01 sampai 10 | Selesai: 10 workbook, tutorial, latihan, preview, verifikasi |
| Proyek portofolio analisis 01 Sales & Profitability | Selesai: data, latihan dan solusi, dokumentasi, preview, verifikasi |
| Proyek portofolio analisis 02 dan 03 | Menyusul |

Bagian bertanda "menyusul" belum ada di repositori dan akan ditambahkan dengan standar yang sama. Saya sengaja tidak memasang file kosong atau setengah jadi.

## Tiga bagian repositori

| Bagian | Folder | Sifat | Isi |
|---|---|---|---|
| Jalur belajar | [`belajar/`](belajar/) | Materi latihan | 15 modul berurutan dari dasar Excel sampai praktik kerja profesional, roadmap 10 minggu, pemetaan kemampuan ke bidang kerja |
| Proyek portofolio analisis | [`proyek/`](proyek/) | Proyek simulasi | Sales & Profitability, Survei Kepuasan Pelanggan, Inventory & Operations |
| Proyek HR dan administrasi | [`proyek-hr/`](proyek-hr/) | Proyek simulasi bertingkat | 10 proyek dari database karyawan sampai workforce planning dan makro VBA, masing-masing dengan tutorial |

## Cara mulai belajar dari nol

1. Pastikan Excel 2016 atau lebih baru (Excel 365 ideal). Baca [`belajar/README.md`](belajar/README.md) untuk prasyarat dan pengaturan regional.
2. Ikuti [`belajar/ROADMAP.md`](belajar/ROADMAP.md): 10 minggu, tiap minggu ada modul yang dipelajari dan file keluaran yang harus jadi.
3. Di setiap modul: baca `README.md`, buka `latihan/latihan_NN.xlsx`, kerjakan, lalu bandingkan dengan `latihan/solusi_NN.xlsx` dan `latihan/kunci_jawaban.md`. Jangan lanjut sebelum checklist kompetensi di akhir modul terpenuhi.
4. Setelah modul 09, kerjakan proyek pertama di `proyek/01_sales_profitability` memakai `workbook/latihan_*.xlsx`; cocokkan dengan workbook solusi dan sheet `Checks`.
5. Modul 10 sampai 15 dan proyek 02, 03 melengkapi kemampuan ke tingkat kerja profesional.

## Cara recruiter menilai portofolio ini

- Buka `proyek/<nama>/README.md` untuk masalah bisnis, pertanyaan analisis, dan hasil utama dalam 2 menit.
- Buka `workbook/solusi_*.xlsx`: sheet `Readme` menjelaskan navigasi; `Raw_Data` dan `Master` adalah input; `Calculation` berisi rumus; `Report` adalah dashboard; `Checks` berisi rekonsiliasi dengan status PASS/FAIL.
- `VERIFIKASI.md` di tiap proyek mencatat hasil recalculation dan pencocokan total antara Python (pandas) dan sel Excel.
- `docs/panduan_wawancara.md` berisi cara saya menjelaskan keputusan teknis dan keterbatasan proyek, tanpa klaim pengalaman yang tidak saya miliki.

## Peta belajar

Lihat tabel lengkap di [`belajar/README.md`](belajar/README.md) dan pemetaan ke bidang kerja di [`belajar/PEMETAAN-KARIER.md`](belajar/PEMETAAN-KARIER.md).

| No | Modul | Inti |
|---|---|---|
| 01 | Dasar Excel | Workbook, sel, range, format, Sort, Filter, Freeze Panes, nilai vs tampilan |
| 02 | Struktur data yang benar | Tidy data, Excel Table, structured references, ID dengan nol di depan |
| 03 | Formula dasar | Referensi relatif/absolut, SUM, AVERAGE, COUNT, ROUND |
| 04 | Logika dan agregasi bersyarat | IF, IFS, AND, OR, SUMIFS, COUNTIFS, AVERAGEIFS |
| 05 | Lookup | XLOOKUP, VLOOKUP, INDEX-MATCH, IFNA vs IFERROR |
| 06 | Cleaning dan validasi | TRIM, SUBSTITUTE, teks ke angka, duplikasi, outlier, Data Validation, Conditional Formatting |
| 07 | Tanggal dan laporan berkala | EOMONTH, NETWORKDAYS, periode, MoM/YoY, pembagi nol |
| 08 | PivotTable dan PivotChart | Agregasi, % of total, Group tanggal, Slicer, Timeline |
| 09 | Visualisasi dan dashboard | KPI, chart yang tepat, target, insight |
| 10 | Power Query | Impor, cleaning, Append, Merge, Unpivot, folder, Refresh |
| 11 | Data Model dan Power Pivot | Fakta/dimensi, relasi, kalender, DAX dasar |
| 12 | Statistika terapan | Deskriptif, CI, korelasi, regresi, uji beda, peramalan, ukuran efek |
| 13 | What-if dan optimasi | Goal Seek, Data Table, Scenario Manager, Solver |
| 14 | Otomasi lanjutan | VBA, Office Scripts, Power Automate |
| 15 | Praktik kerja profesional | Rekonsiliasi, audit formula, proteksi, serah terima |

## Proyek portofolio

| Proyek | Pertanyaan utama | Folder |
|---|---|---|
| Sales & Profitability Analysis | Bagaimana tren penjualan bersih dan laba kotor 2024, pencapaian target per wilayah, dan kontribusi produk serta pelanggan? | [Sales & Profitability](proyek/01_sales_profitability/) |
| Customer Satisfaction Survey | Bagaimana distribusi kepuasan per dimensi dan kanal, seberapa reliabel kuesionernya, dan apa yang boleh serta tidak boleh disimpulkan? | `proyek/02_survei_kepuasan_pelanggan` *(menyusul)* |
| Inventory & Operations | SKU mana yang berisiko habis atau overstock, dan berapa usulan pemesanan berdasarkan asumsi yang ditulis? | `proyek/03_inventory_operations` *(menyusul)* |

## Proyek HR dan administrasi

Sepuluh proyek bertingkat (pemula sampai lanjut) dengan tutorial langkah demi langkah dan latihan: database karyawan, rekap absensi, agenda surat dan stok ATK, penggajian dan slip gaji, pengelolaan cuti, rekrutmen tracker, penilaian kinerja KPI, HR dashboard, analisis survei engagement, serta perencanaan tenaga kerja dengan What-If dan VBA. Daftar lengkap di [`proyek-hr/README.md`](proyek-hr/README.md).

## Standar file dan cara verifikasi

- Semua `.xlsx` dibangun dengan skrip Python (openpyxl) yang ada di folder `tools/` tiap unit, sehingga bisa dibangun ulang: `python tools/build.py` atau `python tools/build_workbook.py`.
- Perhitungan memakai rumus hidup; nilai statis hanya untuk data mentah.
- Setiap file di-recalculate dengan LibreOffice headless dan diperiksa: nol error (`#REF!`, `#NAME?`, `#DIV/0!`, `#VALUE!`, `#N/A` yang tidak disengaja). Total kunci dicocokkan dengan perhitungan pandas. Hasil dicatat di `VERIFIKASI.md` tiap unit.
- Preview PNG berasal dari render nyata file (LibreOffice ke PDF ke PNG), bukan gambar buatan.
- Untuk membangun ulang semuanya dibutuhkan Python 3.10+, `openpyxl`, `pandas`, `numpy`, dan LibreOffice (opsional, untuk recalc).

## Batasan tool dan langkah manual

File dibuat secara terprogram, sehingga fitur berikut **tidak tertanam** di dalam `.xlsx` dan harus dibuat manual mengikuti langkah di dokumen masing-masing: PivotTable, PivotChart, Slicer, Timeline, Power Query, Data Model/Power Pivot, Solver, Goal Seek, dan makro. Untuk setiap fitur itu saya sediakan data siap pakai, skrip M/DAX/VBA bila relevan, langkah klik yang presisi, dan sheet pembanding yang dihitung dengan rumus biasa agar hasil manual bisa dicocokkan. Fungsi Excel 365 (XLOOKUP, FILTER, UNIQUE, TEXTSPLIT, LET) diajarkan di materi dengan alternatif untuk Excel 2016/2019; di dalam file dipakai fungsi yang kompatibel lintas versi.

## Pengaturan Excel yang perlu diketahui

- **Pemisah argumen.** Excel dengan pengaturan regional Indonesia memakai titik koma: `=SUMIFS(C:C;A:A;"Jawa")`. Pengaturan Inggris memakai koma: `=SUMIFS(C:C,A:A,"Jawa")`. Dokumen di repo ini menulis rumus dengan titik koma; file `.xlsx` menyimpan rumus dalam format internal dan menyesuaikan otomatis saat dibuka.
- **Pemisah desimal.** Regional Indonesia memakai koma untuk desimal (`0,25`), Inggris memakai titik (`0.25`).
- **Tanggal.** Simpan tanggal sebagai tanggal asli (bukan teks). Data CSV di repo memakai format ISO `YYYY-MM-DD`.

## Struktur folder

```
portofolio-excel-hr-administrasi/
├── README.md
├── belajar/              # materi latihan: 15 modul, ROADMAP.md, PEMETAAN-KARIER.md
├── proyek-hr/            # 10 proyek HR/administrasi bertingkat
│   ├── 01_database-karyawan/ ... 10_perencanaan-tenaga-kerja-dan-otomasi/
│   ├── docs/             # glosarium rumus, standar data HR
│   └── tools/            # skrip pembangun, generator data, log verifikasi
├── proyek/               # proyek analisis: Sales selesai; survei dan inventori menyusul
├── LICENSE, CHANGELOG.md, CONTRIBUTING.md, .gitignore
```

Tiap unit (modul atau proyek) memiliki `README.md`, folder data atau `latihan/`, `tools/` pembangun, dan `VERIFIKASI.md`.

## Lisensi

[MIT License](LICENSE). Data sintetis dan materi boleh dipakai ulang dengan menyebut sumber.

## Kontak

Immantaka Juniarsa, S1 Statistika Universitas Airlangga. GitHub [@imman20](https://github.com/imman20), email [immantakajuniarsa30@gmail.com](mailto:immantakajuniarsa30@gmail.com).
