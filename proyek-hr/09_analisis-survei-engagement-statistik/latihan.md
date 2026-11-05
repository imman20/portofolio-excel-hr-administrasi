# Latihan 09: Analisis Survei Engagement dengan Statistika

Buka `09_analisis_survei_engagement.xlsx` dan kerjakan di sheet baru bernama `Latihan`. Rumus memakai pemisah titik koma (;).

## Soal

1. Hitung Cronbach's alpha dimensi Kepemimpinan (item F, G, H).
2. Uji t engagement laki-laki vs perempuan (data tidak terurut).
3. Korelasi Spearman sederhana antara Kompensasi dan Engagement.
4. Berapa responden dengan engagement < 2,5?
5. Prediksi engagement bila semua dimensi bernilai 4.

## Kunci jawaban

1. `=3/2*(1-(VAR.S(F5:F204)+VAR.S(G5:G204)+VAR.S(H5:H204))/VAR.S(X5:X204*3))` dengan total = F+G+H per responden pada kolom bantu.
2. Buat kolom bantu `=IF(C5="L";AC5;"")` dan `=IF(C5="P";AC5;"")`, lalu `=T.TEST(kolom_L;kolom_P;2;3)` tidak menerima teks: salin nilai angka saja ke dua kolom (Paste Values, hapus kosong) sebelum T.TEST.
3. Buat kolom peringkat `=RANK.AVG(Y5;$Y$5:$Y$204)` untuk kedua variabel lalu `=CORREL(rank1;rank2)`.
4. `=COUNTIF('Data Survei'!AC5:AC204;"<2,5")`
5. `=Regresi!G5+4*SUM(Regresi!B5:F5)`
