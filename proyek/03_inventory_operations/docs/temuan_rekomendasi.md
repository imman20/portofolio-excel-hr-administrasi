# Temuan dan rekomendasi

Saya memakai hasil parameter awal pada 31/12/2024. Seluruh angka berasal dari [ringkasan verifikasi](../data/processed/ringkasan.json), [KPI pandas](../data/processed/kpi_referensi.csv), dan workbook yang dihitung ulang. Ini simulasi; saya tidak mengklaim dampak finansial atau pengalaman operasional nyata.

## Prioritas stok

| Status | Jumlah SKU |
|---|---:|
| HABIS | 0 |
| RISIKO HABIS | 8 |
| NORMAL | 27 |
| OVERSTOCK | 25 |
| TANPA PERMINTAAN | 0 |

Saya meninjau SKU RISIKO HABIS terlebih dahulu. Status membandingkan stok fisik dengan reorder point, sedangkan usulan memperhitungkan on order dan review. Saya memeriksa apakah PO tiba sebelum perkiraan persediaan habis, apakah pemasok dapat mempercepat pengiriman, dan apakah substitusi tersedia. Tanpa jadwal penerimaan terbuka yang pasti, saya tidak menganggap seluruh outstanding sebagai stok siap pakai.

## Lima usulan terbesar

| Urutan | SKU | Usulan unit teoritis | Status |
|---|---|---:|---|
| 1 | SKU021 | 207,10 | RISIKO HABIS |
| 2 | SKU051 | 111,13 | RISIKO HABIS |
| 3 | SKU042 | 110,57 | NORMAL |
| 4 | SKU027 | 105,85 | NORMAL |
| 5 | SKU026 | 100,06 | NORMAL |

SKU NORMAL dapat memiliki usulan positif karena periode review menambah kebutuhan rata-rata. Saya meninjau pesanan terbuka lebih dahulu untuk mencegah duplikasi, lalu menerapkan kemasan/MOQ dan batas anggaran pada tahap keputusan terpisah. Angka pecahan di tabel tidak dibulatkan menjadi pesanan final.

## Overstock dan nilai konsumsi

Sebanyak 25 SKU memiliki cakupan di atas ambang. Saya mengusulkan penundaan pembelian tambahan atau peninjauan pemindahan setelah memeriksa criticality, rencana permintaan, dan biaya yang belum dimodelkan. Saya tidak merekomendasikan penghapusan stok hanya dari label ini.

ABC menghasilkan 31 SKU A, 18 SKU B, dan 11 SKU C. Saya dapat memulai pemeriksaan ketelitian stok pada kelas A karena nilai konsumsi lebih besar, tetapi barang murah yang kritis tetap perlu pengawasan. Ambang berbasis kontribusi HPP, bukan persentase jumlah SKU.

## Turnover dan riwayat stockout

Rata-rata turnover SKU **9,831 kali**, dan total stockout **1.770 SKU-hari**. Tidak ada SKU HABIS pada akhir tahun, tetapi itu tidak menghapus hari tanpa stok pada periode sebelumnya. Saya membaca turnover bersama stockout: stok rendah dapat menaikkan turnover sambil memperburuk layanan. Hari saldo nol tidak mengukur lost sales; data kebutuhan yang ditolak tidak tersedia.

## Tindak lanjut dan batas keputusan

Saya memerlukan kebutuhan yang tidak terpenuhi, jadwal PO terbuka, ketidakpastian lead time, MOQ, biaya pesan/simpan, kapasitas, serta masa simpan sebelum menerapkan kebijakan nyata. Saya akan memantau pemenuhan permintaan dan ketepatan pemasok pada periode terpisah; perubahan status saja bukan bukti perbaikan. Safety stock normal tidak memvalidasi sendiri asumsi atau service level. [Definisi KPI](definisi_kpi.md) menjelaskan kapan pendekatan ini tidak berlaku.
