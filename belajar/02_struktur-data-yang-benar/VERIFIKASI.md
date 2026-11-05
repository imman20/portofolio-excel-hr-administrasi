# Verifikasi modul `02_struktur-data-yang-benar`

Tanggal verifikasi: 2026-11-05 (Asia/Jakarta). Mesin: LibreOffice headless (recalc) + openpyxl + pandas/scipy.
Skrip: `python tools/course_runtime.py 02_struktur-data-yang-benar` dari folder `belajar/`.

## Hasil recalc per file

| File | Sheet | Jumlah rumus | Jumlah error | Status recalc | Ukuran (KB) |
|---|---|---:|---:|---|---:|
| `latihan_02.xlsx` | Instruksi, Laporan_Lama, Data, ID_Rusak, Latihan | 269 | 0 | success | 27.6 |
| `solusi_02.xlsx` | Instruksi, Laporan_Lama, Data, ID_Rusak, Latihan | 332 | 0 | success | 29.6 |

## Rekonsiliasi pandas vs sel hasil recalc

| Pemeriksaan | Lokasi sel | Nilai pandas | Nilai sel | Selisih | Status |
|---|---|---:|---:|---:|---|
| COUNT tanggal | solusi_02.xlsx:Latihan!C5 | 240 | 240 | 0.0 | PASS |
| COUNT cabang_id (teks) = 0 | solusi_02.xlsx:Latihan!C6 | 0 | 0 | 0.0 | PASS |
| Total penjualan (structured ref) | solusi_02.xlsx:Latihan!C18 | 51749000 | 51749000 | 0.0 | PASS |
| Penjualan Surabaya | solusi_02.xlsx:Latihan!C20 | 23882500 | 23882500 | 0.0 | PASS |
| Surabaya Jan (SUMIFS tanggal) | solusi_02.xlsx:Latihan!C26 | 3779000 | 3779000 | 0.0 | PASS |
| Selisih total vs Laporan_Lama = 0 | solusi_02.xlsx:Latihan!I29 | 0 | 0 | 0.0 | PASS |
| Selisih long vs wide = 0 | solusi_02.xlsx:Latihan!D52 | 0 | 0 | 0.0 | PASS |
| ID 001 ditemukan di Data | solusi_02.xlsx:ID_Rusak!E4 | 106 | 106 | 0.0 | PASS |

Ringkasan: 2 file, total error = 0, rekonsiliasi gagal = 0.

## Catatan

- Sel kuning pada latihan sengaja kosong. Semua perhitungan pada solusi memakai rumus hidup.
- Pembanding dibaca dari CSV dan dihitung dengan pandas/scipy; statistik dibandingkan sampai enam desimal, selisih harus nol.
- Preview berasal dari render nyata sheet Latihan ke PDF dengan LibreOffice, kemudian PNG melalui pdftoppm.
- Fungsi diperiksa terhadap daftar yang diperbolehkan; tidak ada rumus array dinamis atau tautan workbook eksternal.
- Interaksi manual di Excel dan Google Sheets belum diuji langsung; fitur yang memerlukan langkah manual dijelaskan jujur di README.
- Di latihan_02.xlsx sheet Data masih range biasa (peserta membuat tabel tblPenjualan sendiri); rumus Bagian C di file solusi memakai structured references dan terverifikasi di LibreOffice.
- Sheet Laporan_Lama sengaja berisi anti-pola (merged cell, subtotal di tengah, angka sebagai teks) sebagai bahan diskusi, bukan untuk ditiru.
