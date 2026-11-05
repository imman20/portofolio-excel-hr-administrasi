# Instruksi pengerjaan inventori

Saya menyediakan [latihan](../workbook/latihan_inventory_operations.xlsx) dan [solusi](../workbook/solusi_inventory_operations.xlsx). Latihan mempertahankan data dan alur perhitungan, tetapi safety stock, reorder point, dan usulan pesan pada baris setelah contoh sengaja kosong berwarna kuning. Saya menyimpan hasil manual sebagai salinan supaya pembangunan ulang tidak menimpanya.

## Menyiapkan file

1. Unduh file XLSX, buka di Excel, lalu baca Readme. Jika file berasal dari lokasi tepercaya, pilih **Enable Editing** agar sel latihan dapat diisi.
2. Pilih **Formulas > Calculation Options > Automatic** dan tekan **Ctrl+Alt+F9**. Jangan menyimpulkan hasil dari preview GitHub.
3. Buka Master. Kenali input biru B2:B10, tabel SKU, dan supplier. Tanggal acuan dan awal tahun dibatasi kalender simulasi; parameter kebijakan dapat diubah untuk skenario.
4. Buka Raw_Data. Pilih baris header lalu **Data > Filter** jika filter belum aktif. Filter SKU pertama, kemudian periksa transaksi MASUK/KELUAR/PENYESUAIAN beserta qty_sig.
5. Buka Purchase_Order dan filter SKU yang sama. Cocokkan PO selesai dengan MASUK melalui po_id; PO terbuka hanya memberi on_order.

## Menelusuri stok harian

6. Pada Permintaan_Harian, B:BI berisi keluar per SKU dan BK:DR berisi saldo akhir. Gunakan **View > Freeze Panes > Freeze Panes** bila mengubah tampilan. Setiap tanggal kalender harus ada, termasuk tanggal tanpa gerakan.
7. Pilih B2 dan baca Formula Bar: SUMIFS mencocokkan SKU, tanggal, dan jenis KELUAR. Pilih BK2: saldo awal ditambah gerakan bertanda hari pertama. BK3 menambah gerakan hari berikutnya ke saldo sebelumnya.
8. Pada Calculation, cocokkan stok akhir dengan baris terakhir saldo harian. Gunakan **Formulas > Trace Precedents** untuk menelusuri sumber. Saya tidak menyamakan jumlah baris transaksi dengan jumlah hari.
9. Periksa rata-rata permintaan dan sigma pada J/K. Rentang statistik harus mencakup hari nol dan berakhir pada tanggal acuan. Baca asumsi di [definisi KPI](definisi_kpi.md).

## Mengisi latihan

10. Baca contoh `Calculation!N2`: `=NORM.S.INV(Master!$B$3)*K2*SQRT(M2)`. Ketik pola yang sama pada N3 dengan nomor baris yang sesuai, lalu tarik fill handle sampai baris terakhir. Referensi parameter dibuat absolut; sigma dan lead time berubah per SKU.
11. Isi O3: `=J3*M3+N3`, lalu salin ke bawah. Periksa unit: unit/hari × hari menghasilkan unit.
12. Isi Q3: `=MAX(0;O3+J3*Master!$B$4-I3-P3)`, lalu salin ke bawah. On order dikurangkan satu kali; jangan menjumlahkan PO selesai lagi.
13. Tekan **Ctrl+Alt+F9**, lalu buka Checks. Pemeriksaan kelengkapan harus berubah dari BELUM DIISI menjadi PASS. Jika belum, gunakan **Home > Find & Select > Go To Special > Blanks** pada area latihan untuk menemukan kekosongan.
14. Buka Report. Bandingkan jumlah status, lima usulan terbesar, rata-rata turnover, dan stockout dengan solusi serta [hasil verifikasi](../VERIFIKASI.md). Angka status dan peringkat latihan baru bermakna setelah jawaban diisi lengkap.

## Latihan analisis dan kunci

- Telusuri satu SKU HABIS dari saldo awal sampai stok akhir. Kunci: gunakan E+F−G+H, cocokkan dengan saldo harian terakhir; PO terbuka tidak menambah saldo fisik.
- Jelaskan mengapa jumlah hari nol di kisi permintaan tidak sama dengan hari stockout. Kunci: hari tanpa pengeluaran bisa tetap memiliki stok; stockout memakai kisi saldo, bukan kisi keluar.
- Bandingkan satu SKU OVERSTOCK dan satu RISIKO HABIS. Kunci: periksa I/J/L/O/P, bukan nilai stok akhir saja. Ringkasan angka ada di [KPI referensi](../data/processed/kpi_referensi.csv).
- Naikkan periode review pada salinan file. Kunci: usulan tidak boleh turun ketika parameter lain tetap; status risiko berbasis ROP tidak berubah karena periode review tidak masuk ROP. Kembalikan parameter sebelum membandingkan dengan bukti awal.
- Naikkan service level pada salinan file. Kunci: safety stock dan ROP tidak menurun untuk sigma positif; usulan tidak menurun. Periksa asumsi normal dan ketergantungan harian sebelum memilih target lebih tinggi.
- Urutkan salinan Calculation berdasarkan peringkat_abc melalui **Data > Sort** dan pilih **Expand the selection**. Kunci: A mendahului B lalu C berdasarkan HPP keluar; kelas memakai kumulatif sebelum item, bukan jumlah SKU 80/15/5. Rumus tie mengikuti urutan SKU awal; jangan menyimpan pengurutan di file bukti.
- Jelaskan usulan pecahan. Kunci: kebutuhan teoritis belum menerapkan kelipatan kemasan, MOQ, anggaran, dan kapasitas. Pembulatan operasional dilakukan setelah tinjauan, bukan dengan mengganti definisi KPI.

Saya dapat memakai **Insert > PivotTable > From Table/Range** pada Raw_Data untuk pembanding manual: Rows = sku, Columns = jenis, Values = Sum of qty. Untuk stok akhir, gabungkan stok awal + MASUK − KELUAR + PENYESUAIAN; jangan menjumlahkan seluruh qty tanpa tanda. PivotTable harus dibuat manual dan tidak ada di file bawaan.

## Kompatibilitas dan pembangunan ulang

Rumus memakai alternatif Excel 2016 seperti INDEX/MATCH, SUMIFS, COUNTIFS, dan kolom bantu; tidak ada XLOOKUP atau dynamic array. Excel 2016/2019/365 mendukung fungsi inti, tetapi tampilan/interaksi belum diuji langsung di Excel atau Google Sheets. Bukti penghitungan memakai LibreOffice headless, bukan klaim semua aplikasi identik. Untuk membangun ulang dari folder proyek:

```bash
python -m pip install -r requirements.txt
python tools/generate_data.py
python tools/build_workbook.py
python tools/verify.py
```

LibreOffice (`soffice`), `pdfinfo`, dan `pdftoppm` harus ada di PATH. Verifier berhenti jika recalc gagal; file yang baru dibangun belum mewarisi bukti verifikasi lama. Setelah skenario atau perubahan data, jalankan pemeriksaan kembali dengan pembanding yang sesuai.
