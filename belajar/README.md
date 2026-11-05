# Jalur Belajar Excel: dari Pemula sampai Siap Kerja

**Ini materi latihan yang saya susun dan kerjakan sendiri, bukan pengalaman kerja.** Semua dataset di folder ini sintetis (dibangkitkan skrip Python dengan seed tetap); tidak ada data perusahaan atau pribadi nyata. Tujuannya dua: (1) saya mempelajari Excel secara sistematis dengan latihan yang bisa diperiksa, dan (2) recruiter dapat melihat cakupan kemampuan Excel saya beserta bukti file yang terverifikasi.

Latar belakang saya S1 Statistika (Universitas Airlangga, 2024), sehingga modul statistika terapan (12) dan praktik profesional (15) saya tulis dengan ketelitian yang sama seperti modul dasar.

## Cara memakai jalur belajar ini

1. Kerjakan modul berurutan dari 01. Setiap modul punya `README.md` (teori, langkah klik, latihan), folder `latihan/` (dataset, workbook latihan dengan sel kuning yang harus diisi, workbook solusi, kunci jawaban), `tools/build.py` (skrip pembangkit dataset dan workbook), dan `VERIFIKASI.md` (hasil pemeriksaan file solusi).
2. Buka `latihan/latihan_NN.xlsx`, baca sheet `Instruksi`, isi sel kuning, lalu bandingkan dengan `solusi_NN.xlsx` dan `kunci_jawaban.md`.
3. Centang checklist kompetensi di akhir README modul sebelum lanjut.
4. Ikuti `ROADMAP.md` (10 minggu) dan lihat `PEMETAAN-KARIER.md` untuk tahu modul mana yang paling penting bagi bidang kerja yang dituju.

Konvensi warna sel di semua workbook: font biru = input/data mentah, font hitam = rumus, font hijau = tautan antar-sheet, fill kuning = sel yang harus diisi.

## Prasyarat

- Excel 2016 atau lebih baru. Beberapa fungsi (XLOOKUP, TEXTSPLIT, IFS) hanya ada di Excel 2019/365; setiap modul menyediakan alternatifnya.
- Pengaturan regional: contoh di README ditulis untuk Excel regional Indonesia (pemisah argumen `;`, desimal `,`). Pada regional Inggris, ganti `;` dengan `,` dan desimal dengan `.`. Di dalam file `.xlsx`, rumus selalu tersimpan dengan koma dan nama fungsi Inggris; Excel menyesuaikan tampilannya otomatis.
- Power Query dan Power Pivot (modul 10 dan 11): Excel 2016+ untuk Windows. Pada Mac, Power Query tersedia terbatas di Microsoft 365 dan Power Pivot tidak tersedia.
- Solver (modul 13): add-in bawaan, aktifkan lewat `File > Options > Add-ins > Go > Solver Add-in`.
- Untuk menjalankan ulang skrip pembangun: Python 3.10+, `pandas`, `numpy`, `openpyxl`, `scipy` (modul 13). Verifikasi memakai LibreOffice headless.

## Daftar modul

| No | Modul | Fitur utama | Durasi perkiraan | Prasyarat |
|---|---|---|---|---|
| 01 | [Dasar Excel](01_dasar-excel/) | cell/range, shortcut, format angka-tanggal-persen, Paste Special, Sort, Filter, Freeze Panes, nilai vs tampilan | 3 jam | - |
| 02 | [Struktur data yang benar](02_struktur-data-yang-benar/) | tidy data, Excel Table (Ctrl+T), structured references, tipe data, ID berawalan nol, anti-pola | 3 jam | 01 |
| 03 | [Formula dasar](03_formula-dasar/) | operator, referensi relatif/absolut/campuran, SUM..COUNTBLANK, ROUND/ROUNDUP/ROUNDDOWN | 4 jam | 01 |
| 04 | [Logika dan agregasi bersyarat](04_logika-dan-agregasi-bersyarat/) | IF, IFS, AND, OR, SUMIFS, COUNTIFS, AVERAGEIFS, wildcard, kriteria tanggal | 5 jam | 03 |
| 05 | [Lookup dan penggabungan referensi](05_lookup-dan-penggabungan-referensi/) | XLOOKUP, VLOOKUP, INDEX-MATCH, kunci unik, IFNA vs IFERROR | 5 jam | 03, 04 |
| 06 | [Cleaning dan validasi](06_cleaning-dan-validasi/) | TRIM/CLEAN/SUBSTITUTE, teks, konversi angka, duplikat, outlier IQR, Data Validation, Conditional Formatting | 6 jam | 02, 05 |
| 07 | [Tanggal dan laporan berkala](07_tanggal-dan-laporan-berkala/) | fungsi tanggal, NETWORKDAYS, kunci periode YYYY-MM, MoM/YoY, laporan bulanan SUMIFS | 5 jam | 04 |
| 08 | [PivotTable dan PivotChart](08_pivottable-dan-pivotchart/) | Rows/Columns/Values/Filters, Show Values As, Group tanggal, Slicer, Timeline, Refresh | 5 jam | 02, 04 |
| 09 | [Visualisasi dan dashboard](09_visualisasi-dan-dashboard/) | pemilihan KPI, chart line/bar/scatter/histogram, target, dashboard satu layar, anti-pola | 6 jam | 07, 08 |
| 10 | [Power Query](10_power-query/) | Get Data, cleaning, Append, Merge (join), Unpivot, impor folder, parameter, Refresh, skrip M | 6 jam | 02, 06 |
| 11 | [Data Model dan Power Pivot](11_data-model-dan-power-pivot/) | fakta/dimensi, relasi, kalender, calculated column vs measure, DAX dasar | 6 jam | 08, 10 |
| 12 | [Statistika terapan](12_statistika-terapan/) | deskriptif, distribusi, CI, korelasi, regresi LINEST, uji t, peramalan, ukuran efek, keterbatasan | 8 jam | 03, 09 |
| 13 | [What-if dan optimasi](13_what-if-dan-optimasi/) | Goal Seek, Data Table 1 dan 2 variabel, Scenario Manager, Solver (program linear) | 5 jam | 03, 04 |
| 14 | Otomasi lanjutan *(menyusul)* | VBA (.bas), Office Scripts (.ts), Power Automate, keamanan makro; materi lanjutan | 5 jam | 05, 09 (opsional) |
| 15 | Praktik kerja profesional *(menyusul)* | rekonsiliasi, audit formula, dokumentasi KPI, Input-Proses-Output, proteksi, keamanan data, handover | 6 jam | 01 sampai 09 |

Total sekitar 78 jam belajar efektif, dijadwalkan dalam 10 minggu di `ROADMAP.md`.

## Apa yang bisa dan tidak bisa dibuat otomatis di file

File `.xlsx` dibangun dengan openpyxl dan dihitung ulang dengan LibreOffice untuk verifikasi. Konsekuensinya:

- Rumus hidup (SUM, SUMIFS, INDEX/MATCH, LINEST, dll.) ada di file dan sudah diverifikasi nol error.
- Fungsi array dinamis (XLOOKUP, FILTER, UNIQUE, TEXTSPLIT, LET) **dijelaskan di README dengan contoh**, tetapi di file dipakai alternatifnya (INDEX/MATCH, SUMIFS, kolom bantu) agar kompatibel dengan Excel 2016 dan mesin verifikasi.
- PivotTable, Slicer, Timeline, Power Query, Data Model, Solver, Goal Seek, dan makro **tidak bisa disisipkan otomatis**. Untuk fitur-fitur itu saya menyediakan dataset siap pakai, langkah klik presisi, skrip (M, DAX, VBA, Office Script) yang bisa ditempel, dan sheet "hasil yang diharapkan" yang dihitung dengan rumus biasa sebagai pembanding. Setiap README modul menuliskan batasan ini pada bagian kompatibilitas.

## Struktur folder

```
belajar/
├── README.md               # file ini
├── ROADMAP.md              # jadwal 10 minggu
├── PEMETAAN-KARIER.md      # kemampuan Excel x bidang kerja
├── tools/
│   ├── belajar_lib.py      # pustaka bersama pembangun workbook
│   └── verify.py           # recalc LibreOffice + rekonsiliasi pandas -> VERIFIKASI.md
└── NN_slug/
    ├── README.md
    ├── VERIFIKASI.md
    ├── tools/build.py
    └── latihan/ (data_NN.csv, latihan_NN.xlsx, solusi_NN.xlsx, kunci_jawaban.md, ...)
```

Membangun ulang satu modul: `cd belajar/NN_slug && python tools/build.py`. Memverifikasi: `cd belajar && python tools/verify.py NN_slug`.

## Hubungan dengan proyek portofolio

Setelah modul 09, saya menerapkan semuanya pada [Sales & Profitability](../proyek/01_sales_profitability/) dan proyek tematik HR di `../proyek-hr/`. [Survei Kepuasan Pelanggan](../proyek/02_survei_kepuasan_pelanggan/) sudah tersedia untuk latihan cleaning, CI, uji t Welch, alpha, dan NPS; [Inventory & Operations](../proyek/03_inventory_operations/) sudah tersedia untuk latihan stok harian, safety stock, ABC, dan usulan pemesanan. Proyek-proyek itu juga memakai data simulasi berlabel jelas; pemetaannya ada di `PEMETAAN-KARIER.md`.
