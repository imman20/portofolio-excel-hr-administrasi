# Verifikasi modul `01_dasar-excel`

Tanggal verifikasi: 2026-11-05. Mesin: LibreOffice headless (recalc) + openpyxl + pandas.
Skrip: `python tools/verify.py 01_dasar-excel` dari folder `belajar/`.

## Hasil recalc per file

| File | Sheet | Jumlah rumus | Jumlah error | Status recalc | Ukuran (KB) |
|---|---|---:|---:|---|---:|
| `latihan_01.xlsx` | Instruksi, Data, Latihan | 5 | 0 | success | 19.3 |
| `solusi_01.xlsx` | Instruksi, Data, Latihan | 24 | 0 | success | 20.0 |

## Rekonsiliasi pandas vs sel hasil recalc

| Pemeriksaan | Lokasi sel | Nilai pandas | Nilai sel | Selisih | Status |
|---|---|---:|---:|---:|---|
| Jumlah transaksi | solusi_01.xlsx:Latihan!C4 | 150 | 150 | 0 | PASS |
| Total penjualan | solusi_01.xlsx:Latihan!C5 | 22461150 | 22461150 | 0 | PASS |
| Transaksi terbesar | solusi_01.xlsx:Latihan!C6 | 1120000 | 1120000 | 0 | PASS |
| Rata-rata transaksi | solusi_01.xlsx:Latihan!C8 | 149,741 | 149741 | 0 | PASS |
| Transaksi Sembako | solusi_01.xlsx:Latihan!C21 | 31 | 31 | 0 | PASS |
| Kasir Rina qty>=5 | solusi_01.xlsx:Latihan!C22 | 21 | 21 | 0 | PASS |
| Total Minuman | solusi_01.xlsx:Latihan!C23 | 4217450 | 4217450 | 0 | PASS |
| Selisih hari | solusi_01.xlsx:Latihan!C17 | 89 | 89 | 0 | PASS |

Ringkasan: 2 file, total error = 0, rekonsiliasi gagal = 0.

## Catatan

- latihan_01.xlsx sengaja berisi sedikit rumus (hanya pembanding Bagian E); sel kuning kosong diisi peserta.
- Sheet Data pada latihan_01.xlsx sengaja belum diformat (tanggal tampil sebagai angka serial) sebagai bahan latihan format.
