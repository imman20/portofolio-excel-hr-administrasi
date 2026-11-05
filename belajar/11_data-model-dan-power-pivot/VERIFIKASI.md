# Verifikasi modul `11_data-model-dan-power-pivot`

Tanggal verifikasi: 2026-11-05 (Asia/Jakarta). Mesin: LibreOffice headless (recalc) + openpyxl + pandas/scipy.
Skrip: `python tools/course_runtime.py 11_data-model-dan-power-pivot` dari folder `belajar/`.

## Hasil recalc per file

| File | Sheet | Jumlah rumus | Jumlah error | Status recalc | Ukuran (KB) |
|---|---|---:|---:|---|---:|
| `latihan_11.xlsx` | Instruksi, Data, Latihan, Produk, Kalender, Hasil_Model_Diharapkan | 2112 | 0 | success | 45.8 |
| `solusi_11.xlsx` | Instruksi, Data, Latihan, Produk, Kalender, Hasil_Model_Diharapkan | 2121 | 0 | success | 45.9 |

## Rekonsiliasi pandas vs sel hasil recalc

| Pemeriksaan | Lokasi sel | Nilai pandas | Nilai sel | Selisih | Status |
|---|---|---:|---:|---:|---|
| Ringkasan 4 | solusi_11.xlsx:Latihan!C4 | 47820000.0 | 47820000 | 0.0 | PASS |
| Ringkasan 5 | solusi_11.xlsx:Latihan!C5 | 120.0 | 120 | 0.0 | PASS |
| Ringkasan 6 | solusi_11.xlsx:Latihan!C6 | 240.0 | 240 | 0.0 | PASS |
| Ringkasan 7 | solusi_11.xlsx:Latihan!C7 | 398500.0 | 398500 | 0.0 | PASS |
| Ringkasan 8 | solusi_11.xlsx:Latihan!C8 | 16895000.0 | 16895000 | 0.0 | PASS |
| Ringkasan 9 | solusi_11.xlsx:Latihan!C9 | 0.0 | 0 | 0.0 | PASS |
| Ringkasan 10 | solusi_11.xlsx:Latihan!C10 | 366.0 | 366 | 0.0 | PASS |
| Ringkasan 11 | solusi_11.xlsx:Latihan!C11 | 203.0 | 203 | 0.0 | PASS |
| Ringkasan 12 | solusi_11.xlsx:Latihan!C12 | 83.0 | 83 | 0.0 | PASS |
| Penjualan bulan 1 | solusi_11.xlsx:Hasil_Model_Diharapkan!B5 | 3807000 | 3807000 | 0.0 | PASS |
| Pesanan bulan 1 | solusi_11.xlsx:Hasil_Model_Diharapkan!C5 | 8 | 8 | 0.0 | PASS |
| Penjualan bulan 2 | solusi_11.xlsx:Hasil_Model_Diharapkan!B6 | 3641000 | 3641000 | 0.0 | PASS |
| Pesanan bulan 2 | solusi_11.xlsx:Hasil_Model_Diharapkan!C6 | 12 | 12 | 0.0 | PASS |
| Penjualan bulan 3 | solusi_11.xlsx:Hasil_Model_Diharapkan!B7 | 6397000 | 6397000 | 0.0 | PASS |
| Pesanan bulan 3 | solusi_11.xlsx:Hasil_Model_Diharapkan!C7 | 16 | 16 | 0.0 | PASS |
| Penjualan bulan 4 | solusi_11.xlsx:Hasil_Model_Diharapkan!B8 | 2103000 | 2103000 | 0.0 | PASS |
| Pesanan bulan 4 | solusi_11.xlsx:Hasil_Model_Diharapkan!C8 | 4 | 4 | 0.0 | PASS |
| Penjualan bulan 5 | solusi_11.xlsx:Hasil_Model_Diharapkan!B9 | 4805000 | 4805000 | 0.0 | PASS |
| Pesanan bulan 5 | solusi_11.xlsx:Hasil_Model_Diharapkan!C9 | 11 | 11 | 0.0 | PASS |
| Penjualan bulan 6 | solusi_11.xlsx:Hasil_Model_Diharapkan!B10 | 5074000 | 5074000 | 0.0 | PASS |
| Pesanan bulan 6 | solusi_11.xlsx:Hasil_Model_Diharapkan!C10 | 14 | 14 | 0.0 | PASS |
| Penjualan bulan 7 | solusi_11.xlsx:Hasil_Model_Diharapkan!B11 | 4057000 | 4057000 | 0.0 | PASS |
| Pesanan bulan 7 | solusi_11.xlsx:Hasil_Model_Diharapkan!C11 | 12 | 12 | 0.0 | PASS |
| Penjualan bulan 8 | solusi_11.xlsx:Hasil_Model_Diharapkan!B12 | 3654000 | 3654000 | 0.0 | PASS |
| Pesanan bulan 8 | solusi_11.xlsx:Hasil_Model_Diharapkan!C12 | 8 | 8 | 0.0 | PASS |
| Penjualan bulan 9 | solusi_11.xlsx:Hasil_Model_Diharapkan!B13 | 3535000 | 3535000 | 0.0 | PASS |
| Pesanan bulan 9 | solusi_11.xlsx:Hasil_Model_Diharapkan!C13 | 10 | 10 | 0.0 | PASS |
| Penjualan bulan 10 | solusi_11.xlsx:Hasil_Model_Diharapkan!B14 | 3804000 | 3804000 | 0.0 | PASS |
| Pesanan bulan 10 | solusi_11.xlsx:Hasil_Model_Diharapkan!C14 | 11 | 11 | 0.0 | PASS |
| Penjualan bulan 11 | solusi_11.xlsx:Hasil_Model_Diharapkan!B15 | 4316000 | 4316000 | 0.0 | PASS |
| Pesanan bulan 11 | solusi_11.xlsx:Hasil_Model_Diharapkan!C15 | 8 | 8 | 0.0 | PASS |
| Penjualan bulan 12 | solusi_11.xlsx:Hasil_Model_Diharapkan!B16 | 2627000 | 2627000 | 0.0 | PASS |
| Pesanan bulan 12 | solusi_11.xlsx:Hasil_Model_Diharapkan!C16 | 6 | 6 | 0.0 | PASS |
| Distinct ATK | solusi_11.xlsx:Hasil_Model_Diharapkan!C22 | 73 | 73 | 0.0 | PASS |
| Distinct Kemasan | solusi_11.xlsx:Hasil_Model_Diharapkan!C23 | 63 | 63 | 0.0 | PASS |
| Distinct Peralatan | solusi_11.xlsx:Hasil_Model_Diharapkan!C24 | 67 | 67 | 0.0 | PASS |

Ringkasan: 2 file, total error = 0, rekonsiliasi gagal = 0.

## Catatan

- Sel kuning pada latihan sengaja kosong. Semua perhitungan pada solusi memakai rumus hidup.
- Pembanding dibaca dari CSV dan dihitung dengan pandas/scipy; statistik dibandingkan sampai enam desimal, selisih harus nol.
- Preview berasal dari render nyata sheet Hasil_Model_Diharapkan ke PDF dengan LibreOffice, kemudian PNG melalui pdftoppm.
- Fungsi diperiksa terhadap daftar yang diperbolehkan; tidak ada rumus array dinamis atau tautan workbook eksternal.
- Interaksi manual di Excel dan Google Sheets belum diuji langsung; fitur yang memerlukan langkah manual dijelaskan jujur di README.
