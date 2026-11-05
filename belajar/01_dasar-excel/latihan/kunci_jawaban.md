# Kunci jawaban modul 01 - Dasar Excel

Semua angka berasal dari `solusi_01.xlsx` yang telah dihitung ulang (lihat `../VERIFIKASI.md`). Rumus ditulis dengan pemisah `;` (regional Indonesia); regional Inggris memakai `,`.

## Sheet Latihan

### Bagian A - Statistik dasar (Data!J2:J151)

| Sel | Rumus | Hasil |
|---|---|---:|
| C4 | `=COUNTA(Data!A2:A151)` | 150 |
| C5 | `=SUM(Data!J2:J151)` | 22.461.150 |
| C6 | `=MAX(Data!J2:J151)` | 1.120.000 |
| C7 | `=MIN(Data!J2:J151)` | 4.500 |
| C8 | `=AVERAGE(Data!J2:J151)` | 149.741,00 |
| C9 | `=SUM(Data!G2:G151)` | 868 |

### Bagian B - Nilai vs tampilan

| Sel | Rumus / isi | Format | Tampilan |
|---|---|---|---|
| C12 | `=Data!B2` | `yyyy-mm-dd` | 2024-01-01 |
| C13 | `=C12*1` | General | 45292 |
| C14 | `0,1` | `0%` | 10% |
| C15 | `=C14*100` | `#,##0.00` | 10,00 |
| C16 | `=Data!B151` | `yyyy-mm-dd` | 2024-03-30 |
| C17 | `=C16-C12` | `#,##0` | 89 |
| C18 | `=MONTH(C12)` | General | 1 |

### Bagian C - Hasil Filter

| Sel | Kriteria filter | Angka yang dibaca di status bar | Pembanding rumus (solusi) |
|---|---|---:|---|
| C21 | kategori = Sembako | 31 | `=COUNTIFS(Data!$F$2:$F$151;"Sembako")` |
| C22 | kasir = Rina dan qty >= 5 | 21 | `=COUNTIFS(Data!$C$2:$C$151;"Rina";Data!$G$2:$G$151;">=5")` |
| C23 | kategori = Minuman, Sum kolom total | 4.217.450 | `=SUMIFS(Data!$J$2:$J$151;Data!$F$2:$F$151;"Minuman")` |
| C24 | diskon_pct > 0 | 73 | `=COUNTIFS(Data!$I$2:$I$151;">0")` |

### Bagian D - Sort 3 terbesar

| Peringkat | no_transaksi | total | Rumus kolom D |
|---|---|---:|---|
| 1 | TRX-0068 (Mi Instan Goreng dus, qty 10) | 1.120.000 | `=LARGE(Data!$J$2:$J$151;1)` |
| 2 | TRX-0080 (Mi Instan Goreng dus, qty 8) | 896.000 | `=LARGE(Data!$J$2:$J$151;2)` |
| 3 | TRX-0120 (Beras Premium 5 kg, qty 11) | 792.000 | `=LARGE(Data!$J$2:$J$151;3)` |

### Bagian E - Transpose

Hasil tempel Transpose (nilai statis) mulai `E34`:

| kategori | Sembako | Minuman | Makanan | Perawatan | Rumah Tangga |
|---|---:|---:|---:|---:|---:|
| jumlah_transaksi | 31 | 36 | 31 | 15 | 37 |

Jumlah lima kategori = 150 (sama dengan C4). Ini pemeriksaan silang yang selalu saya lakukan.

## Latihan mandiri (bagian 7 README)

1. Angka serial `2024-03-30` adalah **45381**. (`2024-01-01` = 45292; selisih 89 hari.)
2. Kasir Dedi: **50 transaksi**, total `qty` **287**, total penjualan 6.450.375.
3. Sort `kategori` A to Z lalu `total` Largest to Smallest: baris teratas **TRX-0068** (Makanan, Mi Instan Goreng dus, 1.120.000). Kategori "Makanan" berada di urutan abjad pertama setelah tidak ada kategori yang diawali A sampai L, dan di dalamnya total terbesar adalah 1.120.000.
4. Total per bulan: Januari **6.871.525**, Februari **8.976.750**, Maret **6.612.875**. Jumlah ketiganya 22.461.150 (cocok dengan C5).
5. `=Data!B2+30` menghasilkan **2024-01-31**. Sah karena tanggal adalah angka serial (45292 + 30 = 45322), dan 45322 ditampilkan sebagai 31 Januari 2024.
6. `=SUM(L2:L151)` tetap **22.461.150**. Format `"Rp"#,##0` hanya mengubah tampilan menjadi `Rp22.461.150`; nilai tidak berubah.
7. Transaksi dengan `qty >= 10`: **28**. Di antaranya yang mendapat diskon (`diskon_pct > 0`): **11**.

## Cara memeriksa sendiri

- Blok `Data!J2:J151` tanpa filter: status bar harus menampilkan Sum 22.461.150, Average 149.741, Count 150.
- Jumlah lima angka di Bagian E harus 150.
- Jika hasil berbeda, periksa: (a) filter masih aktif, (b) range rumus tidak sampai baris 151, (c) data terurut berbeda tetapi itu tidak mengubah total.
