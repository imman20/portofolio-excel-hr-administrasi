# Kunci jawaban modul 13

Saya menjalankan scipy.optimize.linprog dengan metode HiGHS. Produk A optimum = 26.666666667, produk B = 46.666666667, dan laba maksimum = 2466666.666666667 rupiah. Solusi ini kontinu; pembulatan unit belum merupakan solusi integer optimal.

| Pemeriksaan | Sel | Nilai |
|---|---|---:|
| Pembanding 4 | C4 | 26.66666666666667 |
| Pembanding 5 | C5 | 46.666666666666664 |
| Pembanding 6 | C6 | 2466666.666666667 |
| Pembanding 7 | C7 | 70.0 |
| Pembanding 8 | C8 | 1500000.0 |
| Pembanding 9 | C9 | 2000000.0 |
| Pembanding 10 | C10 | 100.0 |
| Pembanding 11 | C11 | 120.0 |
| Model produk A | B2 | 26.66666666666667 |
| Model produk B | B3 | 46.666666666666664 |
| Model laba | B4 | 2466666.666666667 |
| Sisa jam | B7 | 0 |
| Sisa bahan | B8 | 0 |
| Laba skenario Dasar | G12 | 2000000 |
| Laba skenario Alternatif Harga | G13 | 1200000 |
| Laba skenario Alternatif Volume | G14 | 1000000 |
| Sensitivitas qty 20 | E2 | -1000000 |
| Sensitivitas qty 40 | E3 | 0 |
| Sensitivitas qty 60 | E4 | 1000000 |
| Sensitivitas qty 80 | E5 | 2000000 |
| Sensitivitas qty 100 | E6 | 3000000 |
| Sensitivitas qty 120 | E7 | 4000000 |
| Sensitivitas 100000 40 | H3 | -800000 |
| Sensitivitas 120000 40 | I3 | 0 |
| Sensitivitas 140000 40 | J3 | 800000 |
| Sensitivitas 100000 80 | H4 | 400000 |
| Sensitivitas 120000 80 | I4 | 2000000 |
| Sensitivitas 140000 80 | J4 | 3600000 |
| Sensitivitas 100000 120 | H5 | 1600000 |
| Sensitivitas 120000 120 | I5 | 4000000 |
| Sensitivitas 140000 120 | J5 | 6400000 |

Kedua sumber daya aktif pada solusi kontinu. Goal Seek mengubah satu input untuk mencapai target; Data Table menghitung kombinasi skenario; Scenario Manager menyimpan kelompok input; Solver mengoptimalkan tujuan dengan kendala. File hanya menyediakan rumus pembanding dan model, bukan konfigurasi fitur manual yang sudah tertanam.
