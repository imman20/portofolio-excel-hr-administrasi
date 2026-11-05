# Modul 11 — Data Model dan Power Pivot

> Saya memakai fakta pesanan, dimensi produk, dan kalender sintetis untuk latihan. Rumus dokumen memakai titik koma pada regional Indonesia; regional Inggris memakai koma. Data Model, relasi, serta measure DAX belum tertanam dalam file dan belum dieksekusi pada mesin DAX.

## 1. Tujuan dan prasyarat

Saya belajar memisahkan tabel fakta dan dimensi, menentukan grain, membuat relasi satu-ke-banyak, serta memahami calculated column dan measure. Saya mempelajari SUM, DIVIDE, DISTINCTCOUNT, dan CALCULATE dalam konteks DAX. Workbook menyediakan hasil worksheet yang dapat dipakai untuk memeriksa model manual.

Prasyaratnya adalah lookup, keunikan kunci, kalender, dan PivotTable. Saya membuka [latihan](latihan/latihan_11.xlsx), meninjau sumber, lalu mengerjakan model pada salinan Excel yang mendukung Power Pivot. [Solusi](latihan/solusi_11.xlsx) berisi formula pembanding; ia tidak menyatakan bahwa relasi atau mesin DAX sudah dijalankan.

## 2. Kegunaan di pekerjaan

Ketika beberapa tabel mempunyai peran berbeda, menggabungkan semuanya menjadi satu lembar lebar tidak selalu menjadi pilihan terbaik. Dimensi dapat menyimpan atribut produk sekali, sementara fakta mempertahankan detail pesanan. Model yang rapi membantu filter dan ukuran dipakai konsisten pada banyak laporan.

Simulasi ini menekankan pesanan unik. Satu pesanan mempunyai beberapa baris produk, sehingga menghitung baris bukan menghitung pesanan. Pesanan yang mencakup beberapa kategori juga membuat distinct count tidak aditif antarkategori. Saya memeriksa makna ukuran sebelum membandingkan grand total dengan subtotal.

## 3. Penjelasan dari dasar

Fakta menyimpan kejadian pada grain tertentu. Pada dataset ini satu baris adalah satu baris produk dalam pesanan, bukan satu pesanan lengkap. Dimensi menyimpan atribut yang menjelaskan kejadian, seperti nama dan kategori produk. Kalender adalah dimensi waktu dengan satu tanggal unik pada setiap baris.

Model bintang menghubungkan fakta di pusat dengan dimensi di sekelilingnya. Kunci produk pada dimensi harus unik, sementara kunci yang sama boleh muncul berkali-kali pada fakta. Saya memeriksa bahwa setiap kunci fakta dikenal dimensi, bukan hanya bahwa kedua kolom mempunyai nama mirip.

Relasi membuat filter dari dimensi menjangkau fakta. Jika kategori ATK dipilih pada dimensi Produk, ukuran penjualan dievaluasi untuk baris fakta yang terkait. Relasi tidak perlu menyalin seluruh nama produk ke setiap baris fakta. Saya menghindari hubungan ambigu dan jalur filter bolak-balik yang belum diperlukan pada latihan dasar.

Kalender harus mencakup tanggal transaksi dan sebaiknya tetap mempunyai hari tanpa transaksi. Tahun dan bulan menjadi atribut kalender, bukan ditambahkan ulang dengan definisi berbeda pada setiap laporan. Saya memakai tanggal tanpa komponen jam untuk kunci relasi; timestamp yang masih mempunyai jam dapat gagal cocok dengan tanggal kalender.

Calculated column dihitung per baris dan disimpan sebagai bagian tabel model. Contohnya nilai baris dari qty dikalikan harga satuan. Measure dievaluasi saat laporan meminta hasil dalam filter yang berlaku. Penjualan per kategori dan per bulan dapat memakai satu measure yang sama karena konteks filternya berbeda.

SUM menjumlahkan kolom numerik dalam konteks yang sedang berlaku. DISTINCTCOUNT menghitung nilai kunci berbeda dalam konteks tersebut. DIVIDE menyediakan penanganan pembagi nol dengan hasil alternatif yang dinyatakan. Saya tetap menjelaskan arti alternatif agar angka nol tidak menutupi ketidaktersediaan rasio.

CALCULATE mengevaluasi suatu ekspresi setelah konteks filter dimodifikasi. Pada contoh Penjualan ATK, filter kategori ditetapkan secara eksplisit. Saya tidak menganggapnya otomatis menambah irisan terhadap semua pilihan yang ada; perilaku penggantian filter pada kolom sama perlu dipahami sebelum dipakai dalam laporan.

Distinct count bukan ukuran yang selalu dapat dijumlahkan. Pesanan yang membeli ATK dan Kemasan muncul pada kedua hitungan kategori, tetapi hanya sekali pada grand total pesanan. Grand total DAX menghitung ulang ekspresi dalam konteks total, bukan selalu menambahkan angka yang terlihat di atasnya.

Workbook memperlihatkan pembanding dengan flag kemunculan pertama. Flag global cocok untuk distinct seluruh pesanan, sedangkan flag per pasangan pesanan dan kategori dipakai untuk distinct kategori. Saya tidak memakai flag global yang sama untuk semua konteks karena hasilnya dapat kurang hitung.

## 4. Dataset latihan

[Fakta](latihan/data_11.csv) memuat 240 baris produk dari 120 pesanan sintetis. Kolomnya line_id, order_id, tanggal, product_id, qty, dan harga satuan historis. [Produk](latihan/produk_11.csv) berisi 12 kunci unik beserta nama serta kategori. [Kalender](latihan/kalender_11.csv) mencakup seluruh tahun 2024, termasuk hari kabisat.

Workbook mempunyai Table Fakt, Produk, dan Kalender untuk memudahkan penambahan manual ke model. Kolom nilai_pembanding dan flag adalah alat pemeriksaan worksheet. Sheet Hasil_Model_Diharapkan menampilkan penjualan serta pesanan unik per bulan dan kategori.

[DAX lengkap](dax/measures.md) menjelaskan calculated column, measure, nama tabel, dan relasi. [Kunci jawaban](latihan/kunci_jawaban.md) berasal dari groupby serta nunique pandas. [VERIFIKASI.md](VERIFIKASI.md) membuktikan konsistensi formula pembanding, bukan eksekusi model DAX.

## 5. Langkah praktik bernomor dengan instruksi klik

1. Periksa Data, Produk, dan Kalender. Saya menulis grain fakta dan memeriksa kunci dimensi sebelum membuka menu model. Harga historis tetap pada fakta agar perubahan master tidak mengubah nilai transaksi lama.
2. Jika tab belum muncul, pilih **File > Options > Add-ins**, pada **Manage** pilih **COM Add-ins > Go**, centang **Microsoft Power Pivot for Excel**, lalu **OK**. Ketersediaannya bergantung edisi Excel.
3. Klik Table Fakt, pilih **Power Pivot > Add to Data Model**. Ulangi untuk Table Produk dan Kalender. Saya tidak menganggap nama worksheet otomatis sama dengan nama tabel model; saya memeriksanya pada jendela Power Pivot.
4. Pilih **Power Pivot > Manage > Diagram View**. Hubungkan Produk[product_id] ke Fakt[product_id], kemudian Kalender[tanggal] ke Fakt[tanggal]. Pastikan sisi dimensi merupakan satu dan fakta banyak.
5. Pada Kalender, pilih **Design > Mark as Date Table**, pilih kolom tanggal, lalu konfirmasi. Saya memeriksa tanggal unik, tanpa kosong, dan mencakup transaksi.
6. Pada tabel Fakt di Data View, tambahkan calculated column bernama Nilai dengan ekspresi dari dokumen DAX. Kolom nilai_pembanding pada worksheet membantu memeriksa hasil, bukan menggantikan penjelasan row context.
7. Pilih **Power Pivot > Measures > New Measure**. Buat Penjualan, Qty, Jumlah Pesanan, dan Nilai per Pesanan menggunakan formula pada dokumen. Nama diisi pada kotak nama; formula pada dialog dimulai dengan tanda sama dengan.
8. Buat measure Penjualan ATK serta Penjualan 2024. Saya menguji filter yang berubah dan membaca perbedaan antara measure umum dengan filter yang ditentukan secara eksplisit.
9. Pilih **Insert > PivotTable > From Data Model** atau pilihan model yang tersedia. Tarik Kalender[periode] ke Rows dan measure Penjualan serta Jumlah Pesanan ke Values. Bandingkan setiap bulan dengan Hasil_Model_Diharapkan.
10. Buat pivot kategori dan bandingkan DISTINCTCOUNT per kategori dengan grand total. Saya menjelaskan overlap pesanan, bukan memaksa grand total menjadi jumlah subtotal kategori.
11. Isi ringkasan kuning pada Latihan, hitung ulang, dan cocokkan dengan kunci. Saya menyimpan model manual pada salinan karena pembangunan ulang hanya menyediakan workbook pembanding.

## 6. Contoh formula dan arti tiap argumen

Pada worksheet, `=E2*F2` adalah nilai pembanding per baris. `=IF(COUNTIFS($B$2:B2;B2)=1;1;0)` menandai kemunculan pertama order_id. Untuk konteks kategori, COUNTIFS menambah pasangan rentang kategori dan kategori baris yang sama.

Dalam DAX, `Penjualan := SUM(Fakt[Nilai])` menjumlahkan nilai dalam filter aktif. `Jumlah Pesanan := DISTINCTCOUNT(Fakt[order_id])` menghitung ID berbeda, bukan banyak baris. Notasi daftar measure dijelaskan di dokumen DAX agar dapat dipindahkan dengan benar ke dialog New Measure.

`Nilai per Pesanan := DIVIDE([Penjualan];[Jumlah Pesanan];0)` menerima pembilang, penyebut, dan hasil alternatif. `Penjualan ATK := CALCULATE([Penjualan];Produk[kategori]="ATK")` mengevaluasi measure dengan filter kategori yang dinyatakan. Formula ini belum tertanam dan belum dieksekusi di workbook bawaan.

## 7. Latihan mandiri minimal lima soal bertingkat

1. Dasar: nyatakan grain fakta dan jelaskan perbedaan line_id dengan order_id.
2. Dasar: periksa keunikan produk dan tanggal kalender sebelum membuat relasi.
3. Menengah: buat calculated column Nilai dan measure Penjualan, lalu jelaskan perbedaan evaluasinya.
4. Menengah: hitung pesanan unik serta nilai per pesanan menggunakan measure.
5. Menengah: buat ringkasan bulan dari kalender dan cocokkan dengan pembanding.
6. Lanjut: jelaskan mengapa jumlah distinct kategori dapat melebihi distinct keseluruhan.
7. Lanjut: jelaskan perubahan filter pada CALCULATE dan risiko relasi yang ambigu.

## 8. Kunci jawaban

[Kunci lengkap](latihan/kunci_jawaban.md) memuat penjualan, pesanan unik, banyak baris, nilai per pesanan, dan overlap kategori. Selisih penjualan bulanan terhadap fakta harus nol. Hitungan distinct kategori tidak dipaksa menjumlah ke grand total karena konteksnya berbeda.

Jawaban model harus menyatakan nama tabel, arah relasi, grain, dan filter yang dipakai. Saya tidak menyimpulkan model benar hanya karena satu KPI total cocok; hasil per bulan dan kategori juga perlu diperiksa.

## 9. Kesalahan umum dan cara memeriksa hasil

Relasi gagal dapat berasal dari kunci dimensi duplikat, tipe berbeda, atau tanggal fakta yang masih mempunyai jam. Saya memeriksa data sebelum mengubah arah filter. Menambah relasi bolak-balik untuk memaksa angka cocok dapat membuat model sulit dijelaskan.

Menggunakan COUNT baris sebagai jumlah pesanan menghasilkan angka berlebih pada fakta baris produk. Menggunakan flag global untuk distinct per kategori dapat menghasilkan angka kurang. Saya memeriksa konteks serta menghitung ulang dengan nunique pandas.

## 10. Checklist kompetensi

- [ ] Saya dapat menjelaskan grain fakta dan peran dimensi.
- [ ] Saya dapat membuat relasi satu-ke-banyak dengan kunci valid.
- [ ] Saya dapat membedakan calculated column dari measure.
- [ ] Saya dapat membaca SUM, DIVIDE, DISTINCTCOUNT, dan CALCULATE.
- [ ] Saya dapat menjelaskan distinct count yang tidak aditif.

## 11. Catatan kompatibilitas Excel 2016/2019/365, LibreOffice, dan Google Sheets

Power Pivot dan Data Model memerlukan edisi Excel Windows yang mendukungnya; ketersediaan dapat berbeda antar lisensi dan platform. Workbook bawaan tidak mempunyai model, relasi, atau measure tertanam. DAX harus dibuat manual dan belum diuji pada mesin DAX di lingkungan ini.

LibreOffice hanya menghitung formula pembanding serta menghasilkan [preview nyata](assets/preview.png). Google Sheets tidak menjalankan model Power Pivot tersebut secara langsung. Interaksi Excel belum diuji langsung. Pembangunan ulang memakai `python tools/build.py`, lalu verifikasi dari folder belajar melalui `python tools/course_runtime.py 11_data-model-dan-power-pivot`. Saya menjaga bukti formula terpisah dari klaim fitur model manual.
