# Definisi KPI dan asumsi

Saya menggunakan satu baris per order dengan satu produk. Semua nilai uang adalah rupiah; tampilan workbook tanpa desimal, tetapi perhitungan tidak dibulatkan pada setiap tahap. Contoh formula memakai `;`; regional Inggris memakai `,`.

| KPI | Definisi yang saya gunakan | Lokasi contoh |
|---|---|---|
| omzet_bruto | qty × harga_satuan | Raw_Data!J2 = `=G2*H2` |
| nilai_diskon | omzet_bruto × diskon_pct | K2 = `=J2*I2` |
| omzet_neto | omzet_bruto − nilai_diskon | L2 = `=J2-K2` |
| qty_retur | Total kuantitas retur menurut order_id | M2 memakai SUMIFS ke Retur |
| nilai_retur | qty_retur × harga_satuan × (1 − diskon_pct) | N2 = `=M2*H2*(1-I2)` |
| penjualan_bersih | omzet_neto − nilai_retur | O2 = `=L2-N2` |
| hpp_satuan | Biaya satuan dari master produk, exact match | P2 memakai INDEX/MATCH |
| hpp | (qty − qty_retur) × hpp_satuan | Q2 = `=(G2-M2)*P2` |
| laba_kotor | penjualan_bersih − hpp | R2 = `=O2-Q2` |
| margin_kotor | laba_kotor / penjualan_bersih | S2 = `=IF(O2=0;0;R2/O2)` |
| pencapaian_target | penjualan_bersih / target | Calculation!H5 = `=IF(G5=0;"-";C5/G5)` |
| pertumbuhan_mom | (bulan ini − bulan lalu) / bulan lalu | I6 = `=IF(C5=0;"-";(C6-C5)/C5)` |

## Periode dan retur

Saya mengatribusikan seluruh retur ke tanggal order asal, termasuk retur yang tanggalnya masuk bulan berikutnya. Karena itu laporan menunjukkan profitabilitas kelompok order, bukan jurnal pengembalian berdasarkan tanggal penerimaan retur. Tidak ada data sebelum atau sesudah periode dalam simulasi ini. Untuk pelaporan akuntansi bulanan, kebijakan cutoff perlu ditetapkan ulang.

HPP dianggap konstan per produk sepanjang periode. Barang retur dianggap pulih seluruh biayanya; biaya kerusakan, pengiriman, pajak, biaya operasional, piutang, dan kurs tidak dimodelkan. Laba kotor bukan laba bersih maupun arus kas. Jika retur rusak dan tidak dapat dijual ulang, asumsi pemulihan HPP harus diganti.

## Agregasi dan pembagi nol

Saya menjumlahkan penjualan serta HPP terlebih dahulu, lalu menghitung margin dari total laba dibagi total penjualan. Rata-rata sederhana margin transaksi tidak mewakili margin perusahaan karena ukuran transaksi berbeda. Target berada pada grain bulan × region; penjumlahan target dilakukan dari tabel target tersendiri supaya tidak tergandakan oleh join order.

Retur penuh menghasilkan penjualan dan HPP baris nol. Margin baris ditampilkan nol sebagai konvensi teknis, dengan arti rasio tidak terdefinisi. Margin rekap nol juga memakai konvensi tersebut bila tidak ada penjualan. Pencapaian dengan target nol ditampilkan `-`; target yang tidak positif tetap ditandai gagal di Checks. Januari tidak memiliki pembanding Desember sebelumnya sehingga MoM ditampilkan `-`; bulan dengan pembanding nol juga memakai `-`. Saya tidak membungkus lookup dalam IFERROR agar referensi yang rusak tetap terlihat.

## Kontribusi dan peringkat

Kontribusi region/kategori = penjualan bersih kelompok / total penjualan bersih. Pareto mengurutkan penjualan bersih pelanggan menurun; nilai seri mengikuti customer_id menaik. Kolom kunci peringkat berupa bilangan bulat unik, bukan nilai uang yang ditambah pecahan kecil. Kontribusi 20% pelanggan = total penjualan pelanggan peringkat teratas sebanyak ROUND(0,2 × jumlah pelanggan;0), dibagi penjualan seluruh pelanggan. Jumlah pelanggan dan hasilnya tersedia di Calculation!AU5:AU7. Untuk ukuran populasi lain, aturan pembulatan jumlah pelanggan harus dinyatakan kembali.

Target sintetis dibangkitkan dari realisasi simulasi dengan faktor ambisi dan gangguan acak. Pencapaian target di sini menunjukkan cara menghitung rasio dan menelusuri penyebab, bukan evaluasi independen performa tim nyata. Pola musiman dan margin kategori juga sengaja dimasukkan dalam generator.
