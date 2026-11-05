# Log Verifikasi Workbook

Dibuat otomatis oleh `tools/build_all.py`. Setiap file dibangun dengan openpyxl, dihitung ulang (recalculate) dengan LibreOffice headless, lalu seluruh sel diperiksa terhadap nilai error (#REF!, #NAME?, #DIV/0!, #VALUE!, #N/A, #NUM!, #NULL!).

| Proyek | File | Sheet | Jumlah rumus | Sel error | Ukuran (KB) | Hasil |
|---|---|---|---|---|---|---|
| 01 | `01_database-karyawan/01_database_karyawan.xlsx` | 4 | 297 | 0 | 28.4 | LOLOS |
| 02 | `02_rekap-absensi-bulanan/02_rekap_absensi_bulanan.xlsx` | 4 | 655 | 0 | 28.6 | LOLOS |
| 03 | `03_administrasi-surat-dan-atk/03_administrasi_surat_dan_atk.xlsx` | 9 | 1110 | 0 | 53.8 | LOLOS |
| 04 | `04_penggajian-dan-slip-gaji/04_penggajian_dan_slip_gaji.xlsx` | 7 | 1853 | 0 | 54.1 | LOLOS |
| 05 | `05_pengelolaan-cuti/05_pengelolaan_cuti.xlsx` | 7 | 1401 | 0 | 49.1 | LOLOS |
| 06 | `06_rekrutmen-tracker/06_rekrutmen_tracker.xlsx` | 5 | 627 | 0 | 43.5 | LOLOS |
| 07 | `07_penilaian-kinerja-kpi/07_penilaian_kinerja_kpi.xlsx` | 5 | 1097 | 0 | 37.9 | LOLOS |
| 08 | `08_hr-dashboard-analitik/08_hr_dashboard_analitik.xlsx` | 5 | 1636 | 0 | 67.7 | LOLOS |
| 09 | `09_analisis-survei-engagement-statistik/09_analisis_survei_engagement.xlsx` | 8 | 3098 | 0 | 82.1 | LOLOS |
| 10 | `10_perencanaan-tenaga-kerja-dan-otomasi/10_perencanaan_tenaga_kerja.xlsx` | 5 | 196 | 0 | 31.1 | LOLOS |

**Total rumus hidup: 11970. Total sel error: 0.**

## Rincian rumus per sheet

### Proyek 01: `01_database-karyawan/01_database_karyawan.xlsx`

| Sheet | Jumlah rumus |
|---|---|
| Data Karyawan | 240 |
| Daftar | 0 |
| Ringkasan | 57 |
| Petunjuk | 0 |

### Proyek 02: `02_rekap-absensi-bulanan/02_rekap_absensi_bulanan.xlsx`

| Sheet | Jumlah rumus |
|---|---|
| Absensi | 558 |
| Rekap Departemen | 97 |
| Kode & Libur | 0 |
| Petunjuk | 0 |

### Proyek 03: `03_administrasi-surat-dan-atk/03_administrasi_surat_dan_atk.xlsx`

| Sheet | Jumlah rumus |
|---|---|
| Surat Masuk | 120 |
| Surat Keluar | 90 |
| Ringkasan Surat | 91 |
| Master ATK | 0 |
| Transaksi ATK | 480 |
| Kartu Stok | 171 |
| Laporan Bulanan ATK | 158 |
| Daftar | 0 |
| Petunjuk | 0 |

### Proyek 04: `04_penggajian-dan-slip-gaji/04_penggajian_dan_slip_gaji.xlsx`

| Sheet | Jumlah rumus |
|---|---|
| Parameter | 0 |
| Master Karyawan | 70 |
| Kehadiran & Lembur | 420 |
| Perhitungan Gaji | 1272 |
| Rekap Departemen | 70 |
| Slip Gaji | 21 |
| Petunjuk | 0 |

### Proyek 05: `05_pengelolaan-cuti/05_pengelolaan_cuti.xlsx`

| Sheet | Jumlah rumus |
|---|---|
| Hari Libur 2024 | 57 |
| Saldo Cuti | 360 |
| Pengajuan Cuti | 720 |
| Rekap Cuti | 102 |
| Kalender | 162 |
| Daftar | 0 |
| Petunjuk | 0 |

### Proyek 06: `06_rekrutmen-tracker/06_rekrutmen_tracker.xlsx`

| Sheet | Jumlah rumus |
|---|---|
| Kandidat | 450 |
| Funnel | 31 |
| Ringkasan Rumus | 146 |
| Panduan Pivot | 0 |
| Daftar | 0 |

### Proyek 07: `07_penilaian-kinerja-kpi/07_penilaian_kinerja_kpi.xlsx`

| Sheet | Jumlah rumus |
|---|---|
| Master KPI | 8 |
| Penilaian | 960 |
| 9-Box | 16 |
| Rekap | 113 |
| Petunjuk | 0 |

### Proyek 08: `08_hr-dashboard-analitik/08_hr_dashboard_analitik.xlsx`

| Sheet | Jumlah rumus |
|---|---|
| Dashboard | 45 |
| Data Karyawan | 1250 |
| Perhitungan | 341 |
| Panduan Power Query | 0 |
| Panduan Slicer | 0 |

Sel error: {"#VALUE!": ["Perhitungan!N25", "Perhitungan!N26", "Perhitungan!C40", "Perhitungan!E40", "Perhitungan!F40", "Perhitungan!G40", "Perhitungan!C54"]}

### Proyek 09: `09_analisis-survei-engagement-statistik/09_analisis_survei_engagement.xlsx`

| Sheet | Jumlah rumus |
|---|---|
| Data Survei | 2800 |
| Statistik Deskriptif | 105 |
| Korelasi | 40 |
| Uji t | 18 |
| ANOVA | 41 |
| Regresi | 67 |
| Histogram | 27 |
| Interpretasi | 0 |

### Proyek 10: `10_perencanaan-tenaga-kerja-dan-otomasi/10_perencanaan_tenaga_kerja.xlsx`

| Sheet | Jumlah rumus |
|---|---|
| Historis | 20 |
| Proyeksi | 91 |
| Anggaran Gaji | 82 |
| Data Karyawan | 3 |
| Panduan What-If & VBA | 0 |

## Catatan metode

- Fungsi pasca-2007 (IFS, RANK.EQ, STDEV.S, T.TEST, F.DIST.RT, FORECAST.LINEAR, dsb.) ditulis dengan awalan internal `_xlfn.` sesuai format file Excel; di Excel awalan ini tidak terlihat.
- Fungsi array dinamis Excel 365 (XLOOKUP, FILTER, UNIQUE, SORT) tidak dipakai di sel hidup agar file tetap terbuka sempurna di Excel 2016/2019 dan LibreOffice; alternatif INDEX/MATCH dipakai dan versi 365 dijelaskan di tutorial.
- Screenshot di folder `assets/` dihasilkan dari file yang sama lewat konversi PDF -> PNG LibreOffice.
- Waktu verifikasi: 2026-11-05 07:59.

Catatan perbaikan: pada build awal proyek 08, dua sel bantu KPI menimpa baris Juni dan Juli 2024 di tabel headcount bulanan sehingga terjadi referensi melingkar (7 sel #VALUE!). Sel bantu dipindahkan ke C64:C65 dan file dibangun ulang dengan hasil 0 error.
