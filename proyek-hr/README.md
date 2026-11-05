# Proyek HR dan Administrasi dengan Excel

Sepuluh proyek bertingkat yang menerapkan Excel pada pekerjaan HR, human capital, dan administrasi kantor. Setiap proyek berisi satu workbook dengan rumus hidup, `TUTORIAL.md` langkah demi langkah, `latihan.md` dengan kunci jawaban, dan preview di folder `assets/`.

> **Data simulasi.** Seluruh data karyawan, gaji, absensi, kandidat, dan survei di folder ini adalah data sintetis yang saya bangkitkan dengan skrip (`tools/data_gen.py`, seed tetap). Tidak ada data perusahaan atau pribadi nyata. Ini proyek latihan portofolio, bukan pengalaman kerja.

## Daftar proyek

| No | Proyek | Tingkat | Fitur utama | Workbook | Jumlah rumus |
|---|---|---|---|---|---|
| 01 | [Database Karyawan](01_database-karyawan/TUTORIAL.md) | Pemula | Format as Table, Freeze Panes, Sort & Filter, Data Validation tipe List, Format tanggal dd/mm/yyyy dan rupiah "Rp"#,##0 | [`01_database_karyawan.xlsx`](01_database-karyawan/01_database_karyawan.xlsx) | 297 |
| 02 | [Rekap Absensi Bulanan](02_rekap-absensi-bulanan/TUTORIAL.md) | Pemula | COUNTIF pada rentang horizontal, COUNT untuk menghitung hari kerja dari header tanggal, Persentase dan format 0.0%, Conditional Formatting: Cell Is less than, Equal To | [`02_rekap_absensi_bulanan.xlsx`](02_rekap-absensi-bulanan/02_rekap_absensi_bulanan.xlsx) | 655 |
| 03 | [Administrasi Surat dan ATK](03_administrasi-surat-dan-atk/TUTORIAL.md) | Pemula | TEXT + ROW untuk nomor agenda otomatis, CHOOSE + MONTH untuk bulan romawi, INDEX/MATCH untuk menarik nama dan harga barang, SUMIFS dengan 2-4 kriteria | [`03_administrasi_surat_dan_atk.xlsx`](03_administrasi-surat-dan-atk/03_administrasi_surat_dan_atk.xlsx) | 1110 |
| 04 | [Penggajian dan Slip Gaji](04_penggajian-dan-slip-gaji/TUTORIAL.md) | Menengah | Sheet Parameter sebagai satu sumber asumsi, VLOOKUP exact, INDEX/MATCH lintas sheet, SUMIF untuk lembur, SUMIFS untuk rekap departemen | [`04_penggajian_dan_slip_gaji.xlsx`](04_penggajian-dan-slip-gaji/04_penggajian_dan_slip_gaji.xlsx) | 1853 |
| 05 | [Pengelolaan Cuti](05_pengelolaan-cuti/TUTORIAL.md) | Menengah | NETWORKDAYS.INTL dengan argumen hari libur, DATEDIF untuk masa kerja dalam bulan, SUMIFS dengan rentang yang memanjang, IF bertingkat untuk validasi | [`05_pengelolaan_cuti.xlsx`](05_pengelolaan-cuti/05_pengelolaan_cuti.xlsx) | 1401 |
| 06 | [Rekrutmen Tracker](06_rekrutmen-tracker/TUTORIAL.md) | Menengah | Satu kolom tanggal per tahap, COUNTIFS dengan kriteria "<>", IF bertingkat untuk tahap terakhir, DATEDIF "D" untuk time-to-hire | [`06_rekrutmen_tracker.xlsx`](06_rekrutmen-tracker/06_rekrutmen_tracker.xlsx) | 627 |
| 07 | [Penilaian Kinerja KPI](07_penilaian-kinerja-kpi/TUTORIAL.md) | Menengah | SUMPRODUCT untuk rata-rata berbobot, IFS, RANK.EQ dan peringkat unik, INDEX/MATCH dua arah | [`07_penilaian_kinerja_kpi.xlsx`](07_penilaian-kinerja-kpi/07_penilaian_kinerja_kpi.xlsx) | 1097 |
| 08 | [HR Dashboard Analitik](08_hr-dashboard-analitik/TUTORIAL.md) | Lanjut | Pemisahan Data - Perhitungan - Dashboard, COUNTIFS/AVERAGEIFS dengan kriteria wildcard *, Turnover rate = keluar / rata-rata headcount, Headcount bulanan dari tanggal masuk dan keluar | [`08_hr_dashboard_analitik.xlsx`](08_hr-dashboard-analitik/08_hr_dashboard_analitik.xlsx) | 1636 |
| 09 | [Analisis Survei Engagement dengan Statistika](09_analisis-survei-engagement-statistik/TUTORIAL.md) | Lanjut | AVERAGE item -> skor dimensi, COUNT, MEDIAN, MODE.SNGL, STDEV.S, VAR.S, QUARTILE, SKEW, KURT, Standard error dan CI 95% dengan T.INV.2T, CORREL dan uji signifikansi r | [`09_analisis_survei_engagement.xlsx`](09_analisis-survei-engagement-statistik/09_analisis_survei_engagement.xlsx) | 3098 |
| 10 | [Perencanaan Tenaga Kerja dan Otomasi](10_perencanaan-tenaga-kerja-dan-otomasi/TUTORIAL.md) | Lanjut | FORECAST.LINEAR dan TREND, CAGR dengan pangkat pecahan, MAX untuk kebutuhan rekrutmen non-negatif, Model anggaran dengan parameter terpusat dan skenario INDEX/MATCH | [`10_perencanaan_tenaga_kerja.xlsx`](10_perencanaan-tenaga-kerja-dan-otomasi/10_perencanaan_tenaga_kerja.xlsx) | 196 |

Urutan yang saya sarankan: kerjakan 01 sampai 03 setelah modul belajar 01 sampai 04, proyek 04 sampai 07 setelah modul 05 sampai 07, dan proyek 08 sampai 10 setelah modul 08 sampai 13. Lihat [`../belajar/ROADMAP.md`](../belajar/ROADMAP.md).

## Cara memakai

1. Buka `TUTORIAL.md` proyek yang dipilih dan ikuti langkahnya dari workbook kosong, atau buka workbook jadi untuk menelusuri rumusnya (Formulas > Show Formulas, atau `Ctrl+``).
2. Kerjakan `latihan.md`, lalu cocokkan dengan kunci jawaban di file yang sama.
3. Rumus di dokumen memakai pemisah titik koma (`;`) untuk Excel regional Indonesia. Excel regional Inggris memakai koma (`,`).

## Verifikasi

Semua workbook dibangun ulang dengan `python tools/build_all.py` (membutuhkan Python 3.10+, openpyxl, dan LibreOffice untuk recalculation). Setiap file dihitung ulang dengan LibreOffice headless dan seluruh sel diperiksa terhadap nilai error. Hasil lengkap per sheet ada di [`tools/verifikasi.md`](tools/verifikasi.md): **11.970 rumus hidup, 0 sel error**.

## Batasan dan langkah manual

- PivotTable, Slicer, Timeline, Power Query, Analysis ToolPak, Goal Seek, Scenario Manager, dan Data Table tidak tertanam di file karena workbook dibuat secara terprogram. Tutorial proyek terkait menjelaskan langkah kliknya, dan workbook menyediakan hasil pembanding yang dihitung dengan rumus biasa.
- Makro VBA proyek 10 disediakan sebagai modul teks di folder `vba/`. Untuk memakainya, impor modul lewat Developer > Visual Basic > File > Import File, lalu simpan workbook sebagai `.xlsm`. File `.xlsx` di repo sengaja tanpa makro.
- Perhitungan PPh 21 dan BPJS di proyek 04 memakai penyederhanaan yang dijelaskan di tutorialnya; jangan dipakai sebagai dasar penggajian sungguhan tanpa menyesuaikan regulasi terbaru.

## Dokumen pendukung

- [`docs/glosarium_rumus.md`](docs/glosarium_rumus.md): daftar rumus yang dipakai beserta sintaks dan contoh.
- [`docs/standar-data-hr.md`](docs/standar-data-hr.md): konvensi kolom, format NIK, kode absensi, dan aturan data.
