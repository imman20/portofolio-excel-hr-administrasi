# Membuat PivotTable dan mencocokkan Calculation

PivotTable, PivotChart, Slicer, dan Timeline belum tertanam. Saya menyediakan tabel `Transaksi` dan hasil pembanding berumus. Langkah ini ditujukan untuk Excel 2016+ di Windows; nama ribbon dapat sedikit berbeda menurut versi.

## Rekap bulanan

1. Buka workbook solusi, klik `Raw_Data!A1`, lalu **Table Design > Table Name**. Pastikan nama tabel `Transaksi` dan seluruh kolom perhitungan telah terisi.
2. Pilih **Insert > PivotTable > From Table/Range** (pada Excel 2016 langsung **Insert > PivotTable**). Table/Range harus `Transaksi`. Pilih **New Worksheet > OK**, lalu beri nama sheet `Pivot_Bulanan`.
3. Di **PivotTable Fields**, seret `periode` ke **Rows**. Urutkan **Sort A to Z**. Karena kunci berbentuk `yyyy-mm`, urutannya sama dengan kronologi.
4. Seret `omzet_bruto`, `penjualan_bersih`, `hpp`, dan `laba_kotor` ke **Values** dengan urutan tersebut. Klik dropdown setiap field > **Value Field Settings > Summarize Values By > Sum**. Jika muncul Count, periksa tipe data sumber.
5. Pada dialog yang sama pilih **Number Format > Number**, Decimal places nol, centang pemisah ribuan. Pilih **Design > Report Layout > Show in Tabular Form**, **Subtotals > Do Not Show Subtotals**, dan **Grand Totals > On for Rows and Columns**.
6. Bandingkan setiap periode dan grand total dengan `Calculation!A4:E17`, serta [rekonsiliasi CSV](../data/processed/rekonsiliasi.csv). Jangan hanya memeriksa grand total karena selisih antarbulan dapat saling menutup.
7. Tambahkan margin di sel di luar PivotTable sebagai total laba baris dibagi total penjualan baris. Jangan memasukkan `margin_kotor` ke Values lalu mengambil Average atau Sum. Rumus dokumen memakai `;`; regional Inggris memakai `,`.

## Region, kategori, pelanggan

1. Buat PivotTable baru dari `Transaksi`. Untuk region, tempatkan `region_id` pada Rows; penjualan, HPP, laba pada Values (Sum). Bandingkan dengan `Calculation!A20:E27`. Nama region dapat dicocokkan dari Master; Pivot memakai ID supaya tidak bergantung pada ejaan nama.
2. Untuk kontribusi, tambahkan penjualan bersih sekali lagi ke Values. Buka **Value Field Settings > Show Values As > % of Grand Total**. Cocokkan dengan `Calculation!I21:I26`.
3. Target bukan bagian dari tabel transaksi. Buat PivotTable tersendiri dari tabel `Target`, Rows `region_id`, Values Sum of target_penjualan_bersih. Cocokkan `Calculation!G21:G26`. Hitung pencapaian di luar Pivot sebagai penjualan region dibagi target region. Jangan join target ke semua order lalu menjumlahkannya.
4. Untuk kategori, pakai Rows `kategori`; Values penjualan bersih, HPP, dan laba. Cocokkan `Calculation!A30:D36`. Tambahkan kontribusi dengan % of Grand Total; hitung margin dari total laba/penjualan di luar Pivot.
5. Untuk pelanggan, Rows `customer_id`; Values Sum of penjualan_bersih. Klik kanan nilai > **Sort > Sort Largest to Smallest**. Tambahkan nilai yang sama > **Show Values As > % Running Total In**, Base field `customer_id`. Cocokkan dengan `Calculation!AL5:AQ124` melalui ID, bukan hanya posisi baris; bila nilai seri, gunakan customer_id menaik sebagai penentu urutan. Pivot biasa tidak menampilkan pelanggan tanpa order, sedangkan Calculation memakai seluruh master.
6. Kontribusi pelanggan teratas dihitung setelah pengurutan. Jangan mengasumsikan prinsip 80/20 pasti berlaku; gunakan hasil aktual simulasi pada `Calculation!AU7`.

## Grafik, filter, dan refresh

Klik Pivot bulanan > **PivotTable Analyze > PivotChart > Line**. Chart gabungan penjualan dan target pada Report tetap chart biasa karena sumber target terpisah. Tambahkan **PivotTable Analyze > Insert Slicer > region_id > OK** bila ingin filter region. Setelah memfilter, kembalikan **Clear Filter** sebelum membandingkan dengan rekap keseluruhan.

Untuk Timeline, buat Pivot lain dengan field `tanggal` pada Rows, lalu **PivotTable Analyze > Insert Timeline > tanggal > OK**. Jika tombol tidak aktif, periksa apakah sumber tanggal benar-benar tanggal, bukan teks. Group tanggal melalui klik kanan tanggal > **Group > Months dan Years**. Timeline hanya mengontrol Pivot yang terhubung; ia tidak mengubah SUMIFS pada Calculation.

Sesudah mengubah sumber, pilih **Data > Refresh All**, tekan **Ctrl+Alt+F9**, lalu bandingkan lagi. Simpan Pivot manual pada salinan karena menjalankan builder akan menulis ulang workbook. LibreOffice dan Google Sheets memiliki fitur pivot sendiri; instruksi ribbon ini tidak menjamin interaksi yang sama di kedua aplikasi tersebut.
