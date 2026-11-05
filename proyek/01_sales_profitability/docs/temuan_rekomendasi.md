# Temuan dan rekomendasi

Saya mengutip angka yang sudah dicocokkan antara hasil recalculation workbook dan pandas. Semua data adalah simulasi; rekomendasi berikut adalah latihan penalaran bisnis, bukan hasil implementasi atau pencapaian kerja.

## Profitabilitas tahunan

Penjualan bersih adalah Rp12.253.701.095, HPP Rp9.512.929.400, dan laba kotor Rp2.740.771.695. Margin kotor 22,37%, dengan pencapaian target keseluruhan 100,40%. Diskon dan retur sudah dikurangkan dari penjualan, sementara HPP hanya dibebankan pada qty neto retur. Saya tidak menyamakan laba kotor dengan laba bersih.

## Region

| Region | Penjualan bersih (Rp) | Target (Rp) | Pencapaian |
|---|---:|---:|---:|
| Jabodetabek | 4.758.516.825 | 4.454.000.000 | 106,84% |
| Jawa Barat | 1.763.168.180 | 1.733.000.000 | 101,74% |
| Jawa Tengah & DIY | 1.298.545.720 | 1.378.000.000 | 94,23% |
| Jawa Timur | 1.511.206.095 | 1.497.000.000 | 100,95% |
| Sumatera | 1.871.749.140 | 2.022.000.000 | 92,57% |
| Kalimantan | 1.050.515.135 | 1.121.000.000 | 93,71% |

Saya meninjau region di bawah target melalui bauran kategori, pelanggan aktif, dan definisi target sebelum merekomendasikan tindakan. Target simulasi berasal dari realisasi yang diberi faktor, sehingga perbandingan ini terutama membuktikan cara menghitung dan menelusuri rasio. Pada penggunaan nyata, saya memerlukan target yang ditetapkan sebelum hasil periode diketahui.

## Kategori

| Kategori | Penjualan bersih (Rp) | Margin kotor |
|---|---:|---:|
| Alat Tulis Kantor | 1.281.540.780 | 27,76% |
| Kertas & Cetak | 2.801.169.230 | 6,73% |
| Elektronik Kantor | 3.406.565.790 | 16,16% |
| Furnitur Kantor | 3.150.905.750 | 32,65% |
| Perlengkapan Kebersihan | 1.613.519.545 | 38,24% |

Margin terendah terdapat pada Kertas & Cetak, yaitu 6,73%. Saya menyarankan peninjauan harga, diskon, serta biaya aktual kategori tersebut. Skenario kenaikan margin sebesar dua poin persentase pada penjualan tetap menghasilkan tambahan laba kotor sebesar Rp56.023.385, dibulatkan ke rupiah terdekat. Ini identitas aritmetika dengan asumsi volume dan penjualan tidak berubah, bukan prediksi respons pelanggan maupun bukti peningkatan laba.

## Konsentrasi pelanggan

Kontribusi 20% pelanggan teratas, sebanyak 24 pelanggan, adalah 71,56%. Saya menyarankan pemantauan retensi, keragaman pelanggan menengah, dan eksposur piutang pelanggan utama. Penjualan besar saja tidak cukup untuk memutuskan prioritas: margin, ketepatan pembayaran, dan biaya pelayanan juga diperlukan, sedangkan dua hal terakhir belum tersedia.

## Tindakan dan ukuran pemantauan

| Prioritas | Usulan | Yang perlu diperiksa sebelum tindakan | Ukuran pemantauan |
|---|---|---|---|
| Margin kategori tipis | Tinjau diskon dan harga menurut segmen | Biaya aktual, kontrak, sensitivitas pelanggan | Laba kotor total, margin tertimbang, volume |
| Region di bawah target | Bahas kewajaran target dan bauran penjualan | Target independen, cakupan pelanggan, tenaga penjualan | Pencapaian dan laba per region |
| Pelanggan terkonsentrasi | Susun retensi dan pengembangan pelanggan | Piutang, kontrak, biaya layanan | Kontribusi kelompok teratas dan margin pelanggan |
| Perubahan antarbulan | Tinjau pola kapasitas dan stok | Kalender promosi serta data lebih dari satu tahun | Penjualan, MoM, retur, service level |

Saya tidak menyimpulkan penyebab perubahan antarbulan dari dataset ini. Retur dibebankan ke bulan order, bukan tanggal penerimaan, sehingga analisis arus retur memerlukan rekap tambahan. Seluruh angka dapat ditelusuri melalui [VERIFIKASI.md](../VERIFIKASI.md) dan [definisi KPI](definisi_kpi.md).
