# Glosarium Rumus Excel yang Dipakai

Semua rumus ditulis dengan pemisah titik koma (`;`) sesuai Excel regional Indonesia. Pada Excel regional Inggris ganti `;` menjadi `,`. Kolom "Proyek" merujuk ke nomor folder.

| Fungsi | Sintaks | Contoh | Kapan dipakai | Proyek |
|---|---|---|---|---|
| SUM | `SUM(rentang)` | `=SUM(E5:I5)` | Menjumlahkan komponen gaji | 01, 03, 04 |
| AVERAGE | `AVERAGE(rentang)` | `=AVERAGE(M2:M81)` | Rata-rata gaji/skor | 01, 02, 09 |
| MIN / MAX | `MAX(rentang)` | `=MAX(M2:M81)` | Nilai ekstrem | 01, 06, 09 |
| COUNT / COUNTA | `COUNTA(rentang)` | `=COUNTA(B2:B1000)` | Menghitung jumlah karyawan (sel terisi) | 01, 02, 05 |
| COUNTBLANK | `COUNTBLANK(rentang)` | `=COUNTBLANK(H5:H124)` | Mendeteksi data kosong | 10 |
| COUNTIF | `COUNTIF(rentang;kriteria)` | `=COUNTIF(E5:W5;"H")` | Menghitung kode hadir, jumlah per departemen | 01, 02, 03 |
| COUNTIFS | `COUNTIFS(r1;k1;r2;k2;...)` | `=COUNTIFS(dept;"IT";status;"Aktif")` | Hitung dengan banyak kriteria (funnel, headcount, kalender) | 01, 05, 06, 08 |
| SUMIF | `SUMIF(rentang;kriteria;rentang_jumlah)` | `=SUMIF(C5:C64;A10;X5:X64)` | Total hadir per departemen | 02 |
| SUMIFS | `SUMIFS(rentang_jumlah;r1;k1;...)` | `=SUMIFS(E:E;B:B;kode;D:D;"Masuk")` | Stok masuk/keluar, rekap gaji, cuti terpakai | 03, 04, 05 |
| AVERAGEIF / AVERAGEIFS | `AVERAGEIFS(rentang;r1;k1;...)` | `=AVERAGEIFS(usia;dept;"HR";status;"Aktif")` | Rata-rata bersyarat | 01, 06, 08 |
| MAXIFS | `MAXIFS(rentang;r1;k1)` | `=MAXIFS(nilai;dept;"IT")` | Nilai tertinggi per departemen (Excel 2019+) | 07 |
| SUMPRODUCT | `SUMPRODUCT(a;b)` | `=SUMPRODUCT(F5:J5;K5:O5)` | Nilai akhir berbobot, SS ANOVA, hitung duplikat | 07, 09, 10 |
| IF | `IF(uji;benar;salah)` | `=IF(G5<=H5;"PESAN";"Aman")` | Keputusan dua cabang | semua |
| IF bertingkat | `IF(a;x;IF(b;y;z))` | `=IF(P5>=4;"Tinggi";IF(P5>=3;"Sedang";"Rendah"))` | Kategori lebih dari dua | 02, 04, 07 |
| IFS | `IFS(uji1;hasil1;uji2;hasil2;...)` | `=IFS(P5>=4,5;"Istimewa";P5>=3,5;"Baik";TRUE;"Cukup")` | Pengganti IF bertingkat (Excel 2019+) | 07 |
| AND / OR | `AND(uji1;uji2)` | `=AND(H5<>"Selesai";J5>7)` | Gabungan syarat, Conditional Formatting | 03, 05 |
| IFERROR | `IFERROR(rumus;nilai_jika_error)` | `=IFERROR(C10/B10;0)` | Mencegah #DIV/0! dan #N/A | semua |
| VLOOKUP | `VLOOKUP(kunci;tabel;kolom;FALSE)` | `=VLOOKUP(E5;Parameter!$A$20:$B$23;2;FALSE)` | Ambil tunjangan dari level (pencarian tepat) | 04 |
| VLOOKUP approximate | `VLOOKUP(nilai;tabel;2;TRUE)` | `=VLOOKUP(J5;Parameter!$D$20:$E$44;2;TRUE)` | Tarif pajak berjenjang (tabel diurutkan naik) | 04 |
| INDEX + MATCH | `INDEX(hasil;MATCH(kunci;kolom_kunci;0))` | `=INDEX(Master!B:B;MATCH(A5;Master!A:A;0))` | Lookup ke kiri/kanan, slip gaji, nama dari NIK | 03-10 |
| INDEX dua arah | `INDEX(matriks;MATCH(baris);MATCH(kolom))` | `=INDEX('9-Box'!C6:E8;MATCH(T5;B6:B8;0);MATCH(U5;C5:E5;0))` | Kuadran 9-box | 07 |
| XLOOKUP (365) | `XLOOKUP(kunci;kolom_kunci;hasil;[jika_tidak_ada])` | `=XLOOKUP(A5;tblMaster[NIK];tblMaster[Nama];"-")` | Alternatif modern INDEX/MATCH | tutorial 04, 08 |
| CHOOSE | `CHOOSE(indeks;v1;v2;...)` | `=CHOOSE(MONTH(B5);"I";"II";...)` | Bulan romawi untuk nomor surat | 03 |
| TEXT | `TEXT(nilai;format)` | `=TEXT(ROW()-4;"000")` | Nomor urut tiga digit, nama bulan | 03, 06, 08 |
| ROW | `ROW()` | `=ROW()-4` | Nomor urut otomatis | 03, 05 |
| UPPER / TRIM / & | `UPPER(teks)` | `=UPPER(TEXT(C5;"mmmm yyyy"))` | Judul kalender, gabung teks | 05 |
| TODAY | `TODAY()` | `=DATEDIF(E2;TODAY();"Y")` | Usia dan masa kerja berjalan | 01, 03, 06 |
| DATE / YEAR / MONTH | `DATE(th;bl;tgl)` | `=DATE($E$4;$C$4;1)` | Awal bulan kalender, filter tahun | 05, 08 |
| EOMONTH | `EOMONTH(tanggal;0)` | `=EOMONTH(C5;0)` | Akhir bulan | 05 |
| WEEKDAY | `WEEKDAY(tanggal;2)` | `=C5-WEEKDAY(C5;2)+1` | Senin minggu pertama kalender | 05 |
| DATEDIF | `DATEDIF(awal;akhir;"Y"/"M"/"D"/"YM")` | `=DATEDIF(J2;TODAY();"Y")&" thn"` | Usia, masa kerja, time-to-hire | 01, 05, 06, 08 |
| NETWORKDAYS.INTL | `NETWORKDAYS.INTL(awal;akhir;1;libur)` | `=NETWORKDAYS.INTL(F5;G5;1;'Hari Libur 2024'!A5:A31)` | Hari kerja cuti tanpa akhir pekan dan libur | 05 |
| ROUND | `ROUND(angka;digit)` | `=ROUND(F5/173;0)` | Pembulatan rupiah | 04, 10 |
| RANK.EQ | `RANK.EQ(nilai;rentang;0/1)` | `=RANK.EQ(P5;$P$5:$P$84;0)` | Peringkat kinerja/kehadiran | 02, 07 |
| MEDIAN / MODE.SNGL / QUARTILE | `QUARTILE(rentang;1)` | `=MODE.SNGL(X5:X204)` | Statistik deskriptif | 09 |
| STDEV.S / VAR.S | `STDEV.S(rentang)` | `=STDEV.S(AC5:AC204)` | Simpangan baku dan varians sampel | 09 |
| SKEW / KURT | `SKEW(rentang)` | `=SKEW(AC5:AC204)` | Bentuk distribusi | 09 |
| CORREL | `CORREL(x;y)` | `=CORREL(X5:X204;AC5:AC204)` | Korelasi Pearson | 09 |
| T.TEST | `T.TEST(r1;r2;ekor;tipe)` | `=T.TEST(tetap;kontrak;2;2)` | Uji beda dua rata-rata | 09 |
| T.INV.2T / T.DIST.2T | `T.DIST.2T(t;df)` | `=T.DIST.2T(ABS(B13);B14)` | p-value dan t tabel | 09 |
| F.DIST.RT / F.INV.RT | `F.DIST.RT(F;df1;df2)` | `=F.DIST.RT(E18;C18;C19)` | p-value ANOVA dan regresi | 09 |
| LINEST | `INDEX(LINEST(y;x;TRUE;TRUE);baris;kolom)` | `=INDEX(LINEST(AC5:AC204;X5:AB204;TRUE;TRUE);1;1)` | Regresi linear berganda | 09 |
| FORECAST.LINEAR | `FORECAST.LINEAR(x_baru;y;x)` | `=FORECAST.LINEAR(2025;B5:G5;B13:G13)` | Proyeksi headcount | 10 |
| TREND | `TREND(y;x;x_baru)` | `=TREND(B5:G5;B13:G13;2026)` | Proyeksi linear (alternatif) | 10 |
| Pangkat (^) | `(akhir/awal)^(1/n)-1` | `=(G5/B5)^(1/5)-1` | CAGR | 10 |
| Conditional Formatting rumus | aturan `=rumus` | `=$G5<=$H5` | Sorot baris bermasalah | 02, 03, 05 |
| Data Validation List | sumber `=Daftar!$A$2:$A$7` | - | Dropdown departemen, status, kode | semua |
| Format as Table | Ctrl+T | `tblKaryawan` | Tabel terstruktur, filter, referensi otomatis | 01, 03-09 |
| Freeze Panes | View > Freeze Panes | - | Header tetap terlihat | semua |
| PivotTable / Slicer / Timeline | Insert > PivotTable | - | Ringkasan interaktif (panduan klik) | 06, 08 |
| Power Query | Data > Get Data | - | Menggabungkan CSV (panduan klik) | 08 |
| Data Table / Goal Seek / Scenario Manager | Data > What-If Analysis | - | Simulasi anggaran (panduan klik) | 10 |
| Analysis ToolPak | Data > Data Analysis | - | ANOVA dan regresi (panduan klik) | 09 |
| VBA (Alt+F11) | Developer > Visual Basic | `vba/*.bas` | Otomasi slip PDF, rekap, validasi | 10 |
