# Instruksi pengerjaan

Saya memakai latihan ini untuk menunjukkan alur input, perhitungan, laporan, dan kontrol pada data simulasi. Rumus dalam dokumen memakai titik koma untuk Excel regional Indonesia; regional Inggris memakai koma. Nama fungsi di file selalu disimpan dalam bahasa Inggris.

## Membuka latihan

1. Buka [workbook latihan](../workbook/latihan_sales_profitability.xlsx), pilih sheet `Instruksi`, lalu baca urutan pekerjaan dan legenda warna.
2. Simpan salinan melalui **File > Save As**. Font biru menunjukkan input, font hitam rumus lokal, font hijau referensi antar-sheet, dan latar kuning sel latihan.
3. Buka **Formulas > Calculation Options > Automatic**. Bahan sudah berupa Excel Table; klik di `Raw_Data`, lalu periksa **Table Design > Table Name**, yaitu `Transaksi`.
4. Di `Raw_Data`, periksa bahwa satu baris adalah satu order dengan satu produk. Jangan menambah subtotal atau merge di badan tabel. Gunakan **Data > Filter** pada header untuk meninjau data.

## Mengisi transaksi

Baris awal berisi contoh rumus. Pilih rentang yang ditentukan, lalu tekan **Ctrl+D** untuk menyalin baris paling atas ke bawah; referensi absolut harus tetap terkunci.

1. `J2:L6001`: omzet bruto, nilai diskon, dan omzet neto. Periksa `J2` dengan `=G2*H2`, `K2` dengan `=J2*I2`, dan `L2` dengan `=J2-K2`.
2. `M2:O6001`: total kuantitas retur menurut order, nilai retur setelah diskon, dan penjualan bersih. `M2` memakai `=SUMIFS(Retur!$D$2:$D$121;Retur!$B$2:$B$121;A2)`.
3. `P2:S6001`: ambil HPP satuan dari master produk, kurangi HPP untuk kuantitas yang diretur, hitung laba, lalu margin. Gunakan exact match: `=INDEX(Master!$F$3:$F$42;MATCH(D2;Master!$A$3:$A$42;0))`.
4. Untuk margin, gunakan `=IF(O2=0;0;R2/O2)`. Nilai nol pada retur penuh adalah konvensi tampilan, bukan bukti profitabilitas nol. Kesalahan lookup harus diperbaiki, jangan ditutup dengan IFERROR.
5. `T2:X6001`: periode, kategori, segmen, kunci order dan produk, serta penanda duplikasi. Kunci periode berbentuk `yyyy-mm`. Jangan menjumlahkan margin baris.
6. Pada `Retur`, salin contoh pemeriksaan `F2` sampai baris terakhir. Semua order retur harus ditemukan di transaksi.

## Mengisi rekap dan dashboard

1. Buka `Calculation`. Isi sel kuning blok bulanan `B5:H16`, pertumbuhan `I6:I16`, serta total baris 17. `I5` berupa tanda `-`, karena periode pembanding sebelumnya tidak tersedia.
2. Untuk penjualan Januari, gunakan `=SUMIFS(Raw_Data!$O$2:$O$6001;Raw_Data!$T$2:$T$6001;$A5)`. Sesuaikan sum range untuk omzet dan HPP. Laba adalah penjualan dikurangi HPP; margin agregat adalah total laba dibagi total penjualan.
3. Isi blok region, kategori, segmen, produk, dan pelanggan. Gunakan kunci pada baris yang sama sebagai kriteria. Target berada di `Master!W:Y`; jumlahkan per periode atau region sebelum membagi.
4. Beri peringkat pelanggan berdasarkan penjualan terbesar. Untuk nilai seri, urutan customer_id pada master menjadi penentu. Kolom kunci peringkat memakai COUNTIFS nilai lebih besar ditambah hitungan nilai yang sama sampai baris kini, sehingga setiap peringkat unik tanpa mengubah nilai penjualan.
5. Susun Pareto memakai SMALL atas peringkat dan INDEX/MATCH untuk mengambil pelanggan. Hitung kumulatif penjualan, lalu bagi dengan total seluruh pelanggan. Gunakan semua pelanggan, termasuk yang tidak memiliki transaksi.
6. Pada `Report`, buat kartu KPI yang merujuk `Calculation!C17`, `E17`, `F17`, dan `H17`; tambahkan jumlah order dan nilai retur. Pilih **Insert > Line or Area Chart > Line with Markers** untuk tren penjualan dan target.
7. Buat bar region dan kategori melalui **Insert > Bar Chart**. Untuk Pareto, pilih **Insert > Combo Chart > Custom Combination**: penjualan sebagai Clustered Column, persentase kumulatif sebagai Line pada **Secondary Axis** dengan batas 0–1. Ini chart biasa, bukan PivotChart.
8. Tambahkan sumber, periode, satuan, dan kalimat temuan dengan angka yang merujuk sel. Pisahkan rekomendasi dari kepastian hasil: skenario bukan ramalan.

## Menyelesaikan dan memeriksa

1. Isi kolom C kuning di `Checks`. Kolom D berisi pembanding berumus; status `BELUM DIISI` hilang setelah rumus diisi. Semua pemeriksaan pada solusi harus `PASS`.
2. Tekan **Ctrl+Alt+F9**. Bandingkan hasil dengan [workbook solusi](../workbook/solusi_sales_profitability.xlsx) dan [VERIFIKASI.md](../VERIFIKASI.md).
3. Kerjakan variasi: ubah satu diskon pada salinan; jelaskan mengapa penjualan dan laba berubah, sedangkan HPP satuan tidak berubah. Coba retur penuh dan pastikan penjualan serta HPP baris menjadi nol.
4. Ikuti [panduan PivotTable](pivot_table.md) dan [Power Query](power_query.md) untuk latihan manual. Kedua fitur tersebut belum tertanam dalam file.
5. Jawab pertanyaan: mengapa margin total bukan rata-rata margin transaksi, mengapa target tidak boleh digabung langsung ke semua baris order, dan mengapa retur diatribusikan ke bulan order? Cocokkan penjelasan dengan [definisi KPI](definisi_kpi.md).
