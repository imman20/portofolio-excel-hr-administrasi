# Fungsi modern dan alternatif Excel 2016

Saya menjelaskan fungsi berikut sebagai tambahan materi lintas versi. Seluruh contoh hanya berada pada dokumen; workbook tetap memakai fungsi yang diperbolehkan dan tidak mengandung dynamic array. Pemisah contoh adalah titik koma untuk regional Indonesia; regional Inggris memakai koma. Ketersediaan fungsi bergantung versi serta kanal pembaruan, bukan hanya nama Excel pada perangkat.

## XLOOKUP: mengambil atribut

Pada [modul lookup](../05_lookup-dan-penggabungan-referensi/), kode produk berada pada Data dan master harga terpisah.

```excel
=XLOOKUP(Data!B5;Master!A2:A21;Master!C2:C21;"TIDAK ADA";0)
```

Argumennya nilai dicari, rentang kunci, rentang hasil, pesan tidak ditemukan, dan mode exact. Alternatif Excel 2016 adalah INDEX/MATCH exact dengan IFNA. Saya tetap memeriksa kunci unik serta kode yang tidak dikenal; fungsi modern tidak menyelesaikan duplikasi master secara otomatis.

## FILTER: memilih baris

```excel
=FILTER(Data!A2:I161;Data!G2:G161="PERIKSA KUNCI";"TIDAK ADA")
```

Array sumber disaring menurut kondisi status pada baris yang sejajar. Argumen terakhir memberi pesan saat tidak ada hasil. Alternatif Excel 2016 adalah Data > Filter pada kolom status, lalu menyalin visible cells ke salinan laporan bila diperlukan. COUNTIFS dan SUMIFS menyediakan ringkasan tanpa harus membuat salinan baris. Saya mencatat kondisi filter agar hasil dapat direproduksi.

## UNIQUE: daftar nilai berbeda

```excel
=UNIQUE(Data!B2:B161)
```

Hasilnya daftar kode berbeda dari sumber lookup. Daftar berbeda bukan bukti bahwa setiap kode valid pada master. Alternatif Excel 2016 adalah menyalin kolom ke area terpisah, lalu Data > Remove Duplicates, atau Data > Advanced dengan Unique records only. Saya tidak menghapus pengulangan kode pada tabel transaksi karena banyak pesanan boleh memakai produk sama.

## SORT dan SORTBY: urutan keluaran

```excel
=SORT(Data!A2:I161;6;-1)
=SORTBY(Data!A2:I161;Data!F2:F161;-1;Data!A2:A161;1)
```

SORT mengurutkan array menurut indeks kolom; pada contoh kolom keenam adalah nilai dikenal. SORTBY menggunakan rentang kunci urut terpisah, kemudian ID sebagai pemecah nilai seri. Angka negatif menyatakan menurun dan positif menaik. Alternatif Excel 2016 adalah Data > Sort dengan Expand the selection serta Add Level untuk kunci berikutnya. Saya tidak mengurutkan satu kolom sendirian dan tidak mengubah data sumber hanya untuk membuat tampilan hasil.

## SEQUENCE: deret posisi

```excel
=SEQUENCE(COUNTA(Data!A2:A161);1;1;1)
```

Argumen menyatakan banyak baris, banyak kolom, awal, dan langkah. Contoh membentuk nomor posisi sebanyak ID terisi. Alternatif Excel 2016 adalah Home > Fill > Series atau mengisi dua nilai awal lalu menyeret fill handle. Untuk ID yang dipastikan terisi pada setiap baris, helper `=COUNTA($A$2:A2)` yang disalin ke bawah dapat menghitung urutan kumulatif. Nomor posisi bukan pengganti identifier permanen; urutan dapat berubah saat data disortir.

## TEXTSPLIT: memisahkan kode

Pada [modul cleaning](../06_cleaning-dan-validasi/), kode memakai tanda hubung sebagai delimiter.

```excel
=TEXTSPLIT(Data!F2;"-")
```

Hasil dapat menyebar ke beberapa sel. Alternatif Excel 2016 adalah Data > Text to Columns pada salinan kolom, dengan Destination di area kosong. LEFT, MID, dan RIGHT dapat dipakai untuk pola posisi tetap yang telah diperiksa. Saya meninjau panjang kode serta jumlah delimiter sebelum menganggap seluruh baris mempunyai struktur sama.

## LET: nama lokal dalam formula

Pada [modul formula dasar](../03_formula-dasar/), bruto dibentuk dari qty dan tarif.

```excel
=LET(bruto;Data!D2*Data!E2;ROUND(bruto*(1-Data!F2);0))
```

LET mengikat nama lokal pada suatu ekspresi lalu menggunakannya dalam hasil. Ini dapat membantu keterbacaan dan mengurangi pengulangan ekspresi pada formula yang lebih panjang. Alternatif Excel 2016 adalah kolom helper bruto dan neto yang tersedia pada workbook, atau formula langsung dengan referensi yang jelas. Saya tidak menyembunyikan langkah penting hanya agar seluruh perhitungan terlihat sebagai satu formula.

## LAMBDA: fungsi formula yang dapat digunakan ulang

```excel
=LAMBDA(q;tarif;diskon;ROUND(q*tarif*(1-diskon);0))(Data!D2;Data!E2;Data!F2)
```

Parameter q, tarif, dan diskon dipakai dalam ekspresi, lalu contoh langsung dipanggil dengan input baris. Pada versi yang mendukung, fungsi dapat diberi nama melalui Name Manager untuk penggunaan ulang. Alternatif Excel 2016 adalah calculated column atau helper dengan formula `=ROUND(D2*E2*(1-F2);0)` yang disalin dengan aturan referensi yang konsisten. Saya tidak menambahkan VBA hanya untuk meniru fungsi ini ketika formula biasa sudah cukup.

## Pemeriksaan kompatibilitas

Fungsi yang menghasilkan array memerlukan area keluaran kosong. Sel yang terhalang dapat menimbulkan pesan spill; saya tidak menutupi masalah itu dengan angka statis. Contoh modern belum diuji melalui interaksi Excel dalam lingkungan ini. Jalur workbook 2016 tetap menjadi pembanding utama, dengan recalc serta hasil yang dicatat pada setiap modul. Saya memilih pendekatan sesuai versi penerima dan menyebutkan kebutuhan fungsi modern sebelum serah terima.
