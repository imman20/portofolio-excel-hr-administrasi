# Verifikasi proyek `02_survei_kepuasan_pelanggan`

Tanggal verifikasi: 2026-11-05 (Asia/Jakarta). Mesin: LibreOfficeDev 26.8.0.0.alpha0 2c87e51eeaa2b413ff4ae097b2705eea1995d8e5; pandas 2.2.3, scipy 1.17.0.
Skrip: `python tools/verify.py` dari folder proyek.

## Hasil recalc per file

| File | Sheet | Jumlah rumus | Jumlah error | Status recalc | Ukuran (KB) |
|---|---|---:|---:|---|---:|
| `latihan_survei_kepuasan_pelanggan.xlsx` | Readme, Instruksi, Codebook, Raw_Data, Master, Clean_Data, Analisis_Set, Calculation, Report, Reference, Check_Detail, Checks | 13047 | 0 | success (LibreOffice headless) | 515.9 |
| `solusi_survei_kepuasan_pelanggan.xlsx` | Readme, Codebook, Raw_Data, Master, Clean_Data, Analisis_Set, Calculation, Report, Reference, Check_Detail, Checks | 41313 | 0 | success (LibreOffice headless) | 993.1 |

## Rekonsiliasi pandas/scipy vs sel hasil recalc

| Pemeriksaan | Lokasi sel | Nilai Python | Nilai sel | Selisih | Status |
|---|---|---:|---:|---:|---|
| Baris mentah | Checks!C4 | 420 | 420 | 0 | PASS |
| Responden bersih | Checks!C5 | 414 | 414 | 0 | PASS |
| Duplikat dibuang | Checks!C6 | 6 | 6 | 0 | PASS |
| Sel referensi berbeda | Checks!C7 | 0 | 0 | 0 | PASS |
| Sel pembersih terisi | Checks!C8 | 20160 | 20160 | 0 | PASS |
| CSV akhir: sel_kosong_skor | Checks!C9 | 175 | 175 | 0 | PASS |
| CSV akhir: teks_na | Checks!C10 | 12 | 12 | 0 | PASS |
| CSV akhir: teks_strip | Checks!C11 | 10 | 10 | 0 | PASS |
| CSV akhir: angka_berspasi | Checks!C12 | 16 | 16 | 0 | PASS |
| CSV akhir: luar_rentang_0 | Checks!C13 | 5 | 5 | 0 | PASS |
| CSV akhir: luar_rentang_6 | Checks!C14 | 5 | 5 | 0 | PASS |
| CSV akhir: luar_rentang_7 | Checks!C15 | 3 | 3 | 0 | PASS |
| CSV akhir: usia_outlier | Checks!C16 | 2 | 2 | 0 | PASS |
| CSV akhir: tanggal_format_dmy | Checks!C17 | 62 | 62 | 0 | PASS |
| CSV akhir: kanal_kapitalisasi | Checks!C18 | 34 | 34 | 0 | PASS |
| Produk: n | Checks!C19 | 369 | 369 | 0 | PASS |
| Produk: mean | Checks!C20 | 4.01174345077 | 4.01174345077 | 0.0 | PASS |
| Produk: sd | Checks!C21 | 0.696900922195 | 0.696900922195 | -0.0 | PASS |
| Produk: margin | Checks!C22 | 0.0713405785797 | 0.0713405785797 | 0.0 | PASS |
| Produk: ci_low | Checks!C23 | 3.94040287219 | 3.94040287219 | 0.0 | PASS |
| Produk: ci_high | Checks!C24 | 4.08308402935 | 4.08308402935 | 0.0 | PASS |
| Produk: alpha | Checks!C25 | 0.813482222282 | 0.813482222282 | -0.0 | PASS |
| Layanan: n | Checks!C26 | 375 | 375 | 0 | PASS |
| Layanan: mean | Checks!C27 | 3.71822222222 | 3.71822222222 | 0.0 | PASS |
| Layanan: sd | Checks!C28 | 0.786666465251 | 0.786666465251 | -0.0 | PASS |
| Layanan: margin | Checks!C29 | 0.0798786628553 | 0.0798786628553 | 0.0 | PASS |
| Layanan: ci_low | Checks!C30 | 3.63834355937 | 3.63834355937 | 0.0 | PASS |
| Layanan: ci_high | Checks!C31 | 3.79810088508 | 3.79810088508 | 0.0 | PASS |
| Layanan: alpha | Checks!C32 | 0.835164427193 | 0.835164427193 | -0.0 | PASS |
| Harga: n | Checks!C33 | 372 | 372 | 0 | PASS |
| Harga: mean | Checks!C34 | 3.24551971326 | 3.24551971326 | 0.0 | PASS |
| Harga: sd | Checks!C35 | 0.749999060736 | 0.749999060736 | 0.0 | PASS |
| Harga: margin | Checks!C36 | 0.0764638979491 | 0.0764638979491 | -0.0 | PASS |
| Harga: ci_low | Checks!C37 | 3.16905581531 | 3.16905581531 | 0.0 | PASS |
| Harga: ci_high | Checks!C38 | 3.32198361121 | 3.32198361121 | 0.0 | PASS |
| Harga: alpha | Checks!C39 | 0.79996250093 | 0.79996250093 | 0.0 | PASS |
| Pengiriman: n | Checks!C40 | 366 | 366 | 0 | PASS |
| Pengiriman: mean | Checks!C41 | 3.46174863388 | 3.46174863388 | 0.0 | PASS |
| Pengiriman: sd | Checks!C42 | 0.835418455809 | 0.835418455809 | -0.0 | PASS |
| Pengiriman: margin | Checks!C43 | 0.0858724930536 | 0.0858724930536 | -0.0 | PASS |
| Pengiriman: ci_low | Checks!C44 | 3.37587614083 | 3.37587614083 | 0.0 | PASS |
| Pengiriman: ci_high | Checks!C45 | 3.54762112693 | 3.54762112693 | 0.0 | PASS |
| Pengiriman: alpha | Checks!C46 | 0.849610485258 | 0.849610485258 | -0.0 | PASS |
| Keseluruhan: n | Checks!C47 | 395 | 395 | 0 | PASS |
| Keseluruhan: mean | Checks!C48 | 3.90126582278 | 3.90126582278 | 0.0 | PASS |
| Keseluruhan: sd | Checks!C49 | 0.707374797604 | 0.707374797604 | 0.0 | PASS |
| Keseluruhan: margin | Checks!C50 | 0.0699737671611 | 0.0699737671611 | 0.0 | PASS |
| Keseluruhan: ci_low | Checks!C51 | 3.83129205562 | 3.83129205562 | 0.0 | PASS |
| Keseluruhan: ci_high | Checks!C52 | 3.97123958995 | 3.97123958995 | 0.0 | PASS |
| NPS: n | Checks!C53 | 403 | 403 | 0 | PASS |
| NPS: promotor | Checks!C54 | 119 | 119 | 0 | PASS |
| NPS: pasif | Checks!C55 | 177 | 177 | 0 | PASS |
| NPS: detraktor | Checks!C56 | 107 | 107 | 0 | PASS |
| NPS: value | Checks!C57 | 2.9776674938 | 2.9776674938 | 0.0 | PASS |
| Welch: n_toko | Checks!C58 | 186 | 186 | 0 | PASS |
| Welch: n_online | Checks!C59 | 132 | 132 | 0 | PASS |
| Welch: mean_toko | Checks!C60 | 4.03225806452 | 4.03225806452 | 0.0 | PASS |
| Welch: mean_online | Checks!C61 | 3.75 | 3.75 | 0.0 | PASS |
| Welch: diff | Checks!C62 | 0.282258064516 | 0.282258064516 | 0.0 | PASS |
| Welch: t | Checks!C63 | 3.44812613669 | 3.44812613669 | 0.0 | PASS |
| Welch: df | Checks!C64 | 266.113459828 | 266.113459828 | 0.0 | PASS |
| Welch: p | Checks!C65 | 0.000656027390291 | 0.000656027390291 | 0.0 | PASS |
| Welch: cohen_d | Checks!C66 | 0.39847212099 | 0.39847212099 | 0.0 | PASS |
| CORREL Produk | Checks!C67 | 0.576619177376 | 0.576619177376 | 0.0 | PASS |
| CORREL Layanan | Checks!C68 | 0.602313977591 | 0.602313977591 | 0.0 | PASS |
| CORREL Harga | Checks!C69 | 0.573208895013 | 0.573208895013 | 0.0 | PASS |
| CORREL Pengiriman | Checks!C70 | 0.571425273858 | 0.571425273858 | 0.0 | PASS |
| Kasus lengkap regresi | Checks!C71 | 248 | 248 | 0 | PASS |
| LINEST koefisien 0 | Checks!C72 | 0.570548007209 | 0.570548007209 | 0.0 | PASS |
| LINEST koefisien 1 | Checks!C73 | 0.257719359711 | 0.257719359711 | 0.0 | PASS |
| LINEST koefisien 2 | Checks!C74 | 0.270051661261 | 0.270051661261 | 0.0 | PASS |
| LINEST koefisien 3 | Checks!C75 | 0.195872686509 | 0.195872686509 | -0.0 | PASS |
| LINEST koefisien 4 | Checks!C76 | 0.189916900778 | 0.189916900778 | -0.0 | PASS |
| LINEST galat baku 0 | Checks!C77 | 0.188366525206 | 0.188366525206 | 0.0 | PASS |
| LINEST galat baku 1 | Checks!C78 | 0.0487583603495 | 0.0487583603495 | 0.0 | PASS |
| LINEST galat baku 2 | Checks!C79 | 0.0443604551992 | 0.0443604551992 | 0.0 | PASS |
| LINEST galat baku 3 | Checks!C80 | 0.0466118822104 | 0.0466118822104 | 0.0 | PASS |
| LINEST galat baku 4 | Checks!C81 | 0.041966229606 | 0.041966229606 | 0.0 | PASS |
| LINEST r2 | Checks!C82 | 0.583744714015 | 0.583744714015 | 0.0 | PASS |
| LINEST sey | Checks!C83 | 0.452775430212 | 0.452775430212 | 0.0 | PASS |
| LINEST F | Checks!C84 | 85.1940925927 | 85.1940925927 | 0.0 | PASS |
| LINEST dfres | Checks!C85 | 243 | 243 | 0 | PASS |

## Cakupan dan batas verifikasi

- Saya mencocokkan 11,178 sel bersih (414 responden × 27 variabel), termasuk tanggal, kategori, komentar, skor, dan nilai kosong. Selisih numerik dibulatkan 12 desimal; teks harus sama. Checks memuat 414 status per baris pada Check_Detail.
- Seluruh 9,660 input Raw_Data dicocokkan dengan CSV, termasuk spasi dan literal N/A. Raw_Data menyimpan semua input sebagai teks.
- 82 pembanding agregat/statistik PASS, selisih nol setelah pembulatan 10 desimal. Ini toleransi representasi desimal, bukan persentase. Selisih asli tersedia pada rekonsiliasi CSV melalui kedua nilai pembanding.
- Mean, SD, margin dan kedua batas CI setiap dimensi serta kepuasan keseluruhan, alpha, jumlah NPS, statistik/p-value Welch, CORREL, dan koefisien/galat baku LINEST diperiksa.
- Latihan sengaja menyisakan sel kuning kosong dan berstatus BELUM DIISI; contoh baris pertama dihitung ulang. Tidak ada PASS palsu untuk pekerjaan yang belum diisi.
- Fungsi diperiksa terhadap daftar yang diizinkan sebelum dan sesudah recalc. Tidak ada rumus dynamic array, sel error, atau tautan workbook eksternal. Semua sheet memiliki Freeze Panes; setiap file di bawah 5 MB.
- Preview berasal dari Report pada salinan sementara XLSX melalui LibreOffice PDF dan pdftoppm; PDF memuat 2 halaman.
- Sampel non-acak, self-report, skala ordinal, dan tidak adanya hubungan sebab-akibat membatasi inferensi. Keberhasilan recalc tidak memvalidasi rancangan survei.
- PivotTable, Slicer, Power Query, dan Analysis ToolPak belum tertanam atau dijalankan. Interaksi Excel 2016/2019/365 dan impor Google Sheets belum diuji.
- [Rekonsiliasi CSV](data/processed/rekonsiliasi.csv), [ringkasan dan hash input](data/processed/ringkasan.json), [audit masalah data](docs/masalah_kualitas_data.md).

## Pengujian negatif

| Kerusakan pada salinan | Hasil |
|---|---|
| Duplikat ID berkonflik | DITOLAK |
| Hitungan JSON salah | DITOLAK |
| Tanggal ISO referensi tertukar | DITOLAK |
| Nilai sel hasil pembersihan berubah | DITOLAK |

Saya menjalankan pengujian ini pada objek salinan; file sumber dan workbook hasil tidak diubah.
