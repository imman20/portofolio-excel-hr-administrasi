# Definisi KPI persediaan

Saya memakai pemisah titik koma pada contoh rumus untuk Excel regional Indonesia; Excel regional Inggris memakai koma. Format internal XLSX memakai nama fungsi Inggris dan koma. Periode analisis adalah 01/01/2024–31/12/2024, termasuk hari tanpa transaksi dan hari kabisat. `Master!B2` adalah tanggal acuan tetap; validasi tanggal membatasi analisis ini pada akhir tahun tersebut. Mengubah tahun memerlukan pembangunan ulang kalender, data, dan verifikasi.

## Parameter dan satuan

| Sel Master | Parameter | Makna |
|---|---|---|
| B2 | tanggal_acuan | Akhir tahun simulasi |
| B3 | service_level | 95%, probabilitas tidak stockout per siklus dalam pendekatan normal |
| B4 | periode_review | 14 hari kalender; input untuk usulan pesan |
| B5 | periode_permintaan | 90 hari kalender; rentang awal 03/10/2024–31/12/2024 |
| B6 | ambang_overstock_hari | 90 hari; status jika hari persediaan lebih besar dari ambang |
| B7 | faktor_risiko | 1; pengali reorder point untuk status risiko |
| B8 | ambang_abc_a | 80% kontribusi HPP tahunan |
| B9 | ambang_abc_b | 95% kontribusi HPP tahunan, yaitu A dan B |
| B10 | awal_tahun | Awal kalender simulasi |

Semua sel parameter berwarna biru. Saya memisahkan jendela statistik dari periode review; keduanya memiliki arti berbeda. Qty adalah unit barang, rupiah adalah nilai pada HPP tetap per SKU, waktu adalah hari kalender. Tidak ada kurs, pajak, diskon, retur penjualan, atau biaya simpan dalam simulasi ini.

## Saldo dan permintaan

**Stok akhir** (`Calculation!I`) = stok awal + MASUK − KELUAR + PENYESUAIAN sampai tanggal acuan. Qty MASUK dan KELUAR positif; PENYESUAIAN bertanda positif atau negatif. `Raw_Data!G` memberi tanda negatif hanya pada KELUAR. Stok awal berlaku sebelum seluruh transaksi 01/01/2024. Saldo akhir setiap hari dibentuk secara berantai pada `Permintaan_Harian!BK:DR`. Transaksi pada tanggal yang sama tidak memiliki jam; saya tidak menyimpulkan stockout intrahari.

**Permintaan harian rata-rata** (`J`) = jumlah unit KELUAR pada 90 hari terakhir / 90. `Permintaan_Harian!B:BI` memuat satu kolom per SKU dan satu baris per tanggal, berumus SUMIFS, termasuk nol. Contoh permintaan harian SKU pertama:

```excel
=SUMIFS(Raw_Data!$E$2:$E$9088;Raw_Data!$C$2:$C$9088;B$1;Raw_Data!$B$2:$B$9088;$A2;Raw_Data!$D$2:$D$9088;"KELUAR")
```

Batas baris mengikuti CSV saat pembangunan ulang. Statistik ini mengukur keluar tercatat, bukan seluruh permintaan pasar. Saat stockout, transaksi keluar tersensor oleh ketersediaan; tanpa catatan permintaan ditolak atau backorder, kebutuhan sesungguhnya tidak dapat dipulihkan.

**Simpangan baku permintaan harian** (`K`) memakai `STDEV.S` pada rentang kalender yang sama, dengan pembagi n−1. Hari nol wajib ikut agar rata-rata dan variasi tidak bias ke hari aktif saja. Parameter B5 minimal dua hari. Statistik diasumsikan cukup stabil untuk kebutuhan dekat; promosi, tren, hari libur, atau perubahan bauran dapat membuat periode lampau tidak representatif.

**Hari persediaan** (`L`) = stok akhir / rata-rata keluar harian. Saat rata-rata nol, workbook menulis nol sebagai penanda dan memberi status TANPA PERMINTAAN jika stok positif. Nol tersebut tidak berarti persediaan akan habis hari ini; cakupan waktunya tidak terdefinisi. Saya tidak mengubah pembagi nol menjadi kesimpulan bisnis.

## Lead time, safety stock, dan reorder point

**Lead time aktual** (`M`) adalah rata-rata selisih `tanggal_terima − tanggal_pesan` dari PO SELESAI per SKU selama periode analisis. Workbook menghitung selisih tiap PO pada `Purchase_Order!I` dan merata-ratakannya dengan AVERAGEIFS. PO TERBUKA dikecualikan, bukan dianggap lead time nol. Jika tidak ada PO selesai, formula memakai lead time nominal supplier dari tabel Master sebagai asumsi; data simulasi saat ini memiliki PO selesai untuk setiap SKU. Kedatangan parsial tidak dimodelkan.

**Safety stock** (`N`) mengikuti:

```excel
=NORM.S.INV(Master!$B$3)*K2*SQRT(M2)
```

Dengan parameter awal, `safety_stock = NORM.S.INV(0,95) × σ permintaan harian × √lead_time`. Saya mengasumsikan permintaan harian independen, mendekati normal pada agregasi lead time, dan lead time dianggap tetap pada rata-rata aktual. Service level 95% adalah sasaran probabilitas tidak kehabisan dalam satu siklus menurut pendekatan tersebut, bukan persentase unit permintaan yang terpenuhi dan bukan jaminan empiris.

Rumus ini tidak memasukkan varians lead time atau korelasi permintaan dengan keterlambatan pemasok. Permintaan intermiten, musiman, sangat miring, tren cepat, substitusi SKU, dan penjualan tersensor saat stockout melanggar asumsi. Untuk keputusan nyata saya memerlukan data kebutuhan yang tidak terpenuhi dan distribusi lead time, lalu melakukan pengujian kebijakan pada data terpisah. Saya tidak mengklaim service level aktual mencapai parameter hanya karena rumusnya menggunakan 95%.

**Reorder point** (`O`) = rata-rata permintaan harian × lead time + safety stock:

```excel
=J2*M2+N2
```

ROP adalah unit, bukan hari atau rupiah. Saya memakai stok fisik untuk memberi tanda risiko segera; keberadaan PO terbuka tidak otomatis menghilangkan risiko sebelum barang tiba.

## Usulan pemesanan

**On order** (`P`) adalah jumlah `MAX(0; qty_pesan − qty_terima)` pada PO TERBUKA yang sudah dipesan sampai tanggal acuan. PO selesai tidak masuk. Workbook menjumlahkan outstanding pada `Purchase_Order!J`. Tidak ada pembatalan, backorder pelanggan, atau reservasi stok dalam simulasi.

**Usulan pesan** (`Q`) mengikuti tepat:

```excel
=MAX(0;O2+J2*Master!$B$4-I2-P2)
```

Dengan demikian `usulan_pesan = MAX(0; reorder_point + permintaan harian × periode_review − stok_akhir − on_order)`, dengan periode review awal 14 hari sebagai input. MAX membatasi usulan agar tidak negatif. Angka keluaran adalah unit teoritis yang bisa pecahan; saya tidak membulatkan diam-diam dan belum menerapkan MOQ, kelipatan kemasan, kapasitas gudang, atau anggaran.

Penambahan permintaan periode review memperpanjang kebutuhan rata-rata, tetapi safety stock di sini tetap memakai √lead_time sesuai definisi proyek. Ini bukan kebijakan periodic review penuh yang biasanya memperhitungkan variasi selama lead time + periode review. Karena itu saya tidak menafsirkan usulan ini sebagai kebijakan optimal atau jaminan service level untuk interval perlindungan yang lebih panjang. On order diasumsikan dapat diandalkan; jadwal kedatangan perlu diperiksa sebelum mengeksekusi pesanan.

## Nilai persediaan dan turnover

**HPP keluar setahun** (`R`) = total unit KELUAR setahun × HPP satuan tetap. PENYESUAIAN tidak masuk HPP keluar; MASUK bukan penjualan. **Rata-rata stok harian** (`S`) adalah rata-rata 366 saldo akhir hari. **Rata-rata nilai persediaan** (`T`) = rata-rata stok harian × HPP satuan. **Inventory turnover** (`U`) = HPP keluar setahun / rata-rata nilai persediaan:

```excel
=IF(T2=0;0;R2/T2)
```

Penyebut memakai seluruh saldo akhir hari, bukan rata-rata dua titik awal/akhir tahun. Jika penyebut nol, nilai nol adalah penanda perlunya pemeriksaan. Pada data ini penyebut harus positif dan diperiksa verifier. Penilaian tidak memakai FIFO, moving average cost, atau perubahan harga; rasio tidak dapat dibandingkan langsung dengan perusahaan yang memakai konvensi penyebut lain.

Report menampilkan rata-rata aritmetika turnover SKU (setiap SKU berbobot sama), bukan rasio HPP total terhadap rata-rata nilai total. Keduanya menjawab pertanyaan berbeda. Tahun kabisat digunakan apa adanya, tanpa mengubah pembagi harian menjadi 365.

## Stockout, ABC, dan status

**Hari stockout** (`V`) = jumlah hari kalender dengan stok akhir hari tepat nol. Saya memakai COUNTIFS pada saldo harian. Total pada Report adalah penjumlahan SKU-hari; beberapa SKU yang habis pada tanggal sama dihitung masing-masing. Ini bukan jumlah tanggal unik, frekuensi kejadian, durasi jam, atau lost sales. Stockout intrahari yang pulih sebelum akhir hari tidak terlihat.

**ABC** (`Z`) memakai kontribusi HPP keluar tahunan, diurutkan menurun; jika sama, SKU menaik. Kolom W memberi peringkat unik tanpa dynamic array, X memberi kontribusi, dan Y menjumlah kontribusi SKU sebelum item tersebut. Item dengan kumulatif sebelum item <80% masuk A, berikutnya <95% masuk B, sisanya C. Item yang melintasi ambang tetap masuk kelas sebelumnya; hasil kontribusi kelas dapat melewati persentase sasaran karena SKU tidak dibagi. ABC mengukur nilai konsumsi, bukan profit, kritikalitas, atau stok akhir.

**Status** (`AA`) memakai prioritas berurutan: HABIS jika stok nol; TANPA PERMINTAAN jika stok positif dan rata-rata keluar nol; RISIKO HABIS jika stok ≤ ROP × faktor risiko; OVERSTOCK jika hari persediaan > ambang; selain itu NORMAL. Label bersifat saling lepas. Status tidak mempertimbangkan criticality, kedaluwarsa, ruang gudang, atau tanggal tiba PO. Overlap risiko dan overstock diselesaikan dengan prioritas risiko. Ambang ABC dan status dapat diubah di Master; Checks memeriksa urutan ambang ABC.
