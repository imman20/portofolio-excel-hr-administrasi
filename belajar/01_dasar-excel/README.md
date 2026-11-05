# Modul 01 - Dasar Excel

> Materi latihan. Dataset sintetis (dibangkitkan skrip `tools/build.py`), bukan data perusahaan nyata.
> Rumus di dokumen ini ditulis dengan pemisah titik koma `;` (Excel regional Indonesia). Jika Excel Anda memakai regional Inggris, ganti `;` dengan `,`.

## 1. Tujuan dan prasyarat

Setelah menyelesaikan modul ini saya dapat:

- membedakan workbook, worksheet, cell, dan range, serta bergerak cepat di lembar kerja dengan keyboard;
- memformat angka, tanggal, dan persen tanpa mengubah nilai yang tersimpan;
- menjelaskan mengapa tanggal `2024-01-01` sebenarnya adalah angka `45292` dan mengapa `0,1` bisa tampil sebagai `10%`;
- memakai Paste Special (Values, Transpose), Sort, Filter, dan Freeze Panes pada tabel transaksi;
- menulis lima rumus agregasi pertama: `SUM`, `COUNTA`, `MAX`, `MIN`, `AVERAGE`.

Prasyarat: tidak ada. Modul ini mengasumsikan pembaca belum pernah membuka Excel. Yang dibutuhkan hanya Excel 2016 atau lebih baru (atau Microsoft 365) dan file di folder `latihan/`.

## 2. Kegunaan di pekerjaan

Hampir semua pekerjaan kantor di Indonesia menyentuh Excel: rekap penjualan harian toko, daftar hadir karyawan, laporan kas, hingga ekspor data dari sistem ERP. Masalah paling sering yang saya lihat pada file kantor bukan rumus yang rumit, melainkan dasar yang keliru: tanggal tersimpan sebagai teks sehingga tidak bisa diurutkan, angka "1.250.000" yang sebenarnya teks sehingga `SUM` menghasilkan 0, atau header yang hilang saat digulir karena Freeze Panes tidak dipasang. Modul ini menutup lubang-lubang itu dulu sebelum masuk ke rumus.

## 3. Penjelasan dari dasar

### 3.1 Workbook, worksheet, cell, range

- **Workbook** adalah file Excel itu sendiri (`.xlsx`). Satu workbook berisi satu atau lebih **worksheet** (sheet), yang tab-nya ada di bagian bawah layar.
- **Cell** (sel) adalah kotak tempat nilai disimpan. Alamatnya gabungan huruf kolom dan nomor baris: `B2` berarti kolom B baris 2. Alamat sel yang sedang aktif tampil di **Name Box** (kotak kecil di kiri atas, di sebelah kiri Formula Bar).
- **Range** adalah kumpulan sel, ditulis dengan titik dua: `A2:J151` berarti dari `A2` sampai `J151` (semua sel di dalam persegi itu). Range satu kolom penuh ditulis `J:J`.
- **Formula Bar** menampilkan isi sel yang sebenarnya (nilai atau rumus). Yang terlihat di sel adalah *tampilan*; yang ada di Formula Bar adalah *nilai*. Pembedaan ini penting dan dibahas di 3.4.

### 3.2 Navigasi dan shortcut penting

| Shortcut | Fungsi |
|---|---|
| `Ctrl + panah` | lompat ke tepi blok data (sangat cepat untuk tabel panjang) |
| `Ctrl + Shift + panah` | memilih dari sel aktif sampai tepi blok data |
| `Ctrl + Home` / `Ctrl + End` | ke sel A1 / ke sel terakhir yang terpakai |
| `Ctrl + A` | memilih seluruh blok data (tekan dua kali untuk seluruh sheet) |
| `Ctrl + Page Down` / `Page Up` | pindah sheet |
| `F2` | mengedit isi sel aktif |
| `F4` | saat mengedit rumus: mengunci referensi (`A1` menjadi `$A$1`); di luar edit: mengulang aksi terakhir |
| `Ctrl + ;` | menyisipkan tanggal hari ini sebagai nilai |
| `Ctrl + 1` | membuka dialog Format Cells |
| `Ctrl + Shift + L` | mengaktifkan/mematikan Filter |
| `Ctrl + T` | mengubah range menjadi Excel Table (dibahas modul 02) |
| `Ctrl + Z` / `Ctrl + Y` | undo / redo |
| `Alt + =` | menyisipkan `SUM` otomatis |
| `Ctrl + Shift + ~` | format General (menampilkan nilai mentah) |

Kebiasaan yang saya latih: tangan kiri di `Ctrl`, tangan kanan di panah. Untuk memilih seluruh kolom `total` pada 150 baris: klik `J2`, lalu `Ctrl + Shift + panah bawah`.

### 3.3 Format angka, tanggal, dan persen

Format mengubah **tampilan**, bukan nilai. Semua diatur dari `Home > Number` atau dialog `Ctrl + 1 > Number`.

| Kebutuhan | Kode format | Contoh nilai | Tampilan |
|---|---|---|---|
| Ribuan dengan pemisah | `#,##0` | 1250000 | 1.250.000 (regional Indonesia) |
| Dua desimal | `#,##0.00` | 149741 | 149.741,00 |
| Persen | `0%` atau `0.0%` | 0,1 | 10% atau 10,0% |
| Tanggal ISO | `yyyy-mm-dd` | 45292 | 2024-01-01 |
| Tanggal Indonesia | `dd/mm/yyyy` | 45292 | 01/01/2024 |
| Rupiah | `"Rp"#,##0` | 72000 | Rp72.000 |

Catatan regional: pada Excel regional Indonesia, pemisah desimal adalah koma dan pemisah ribuan adalah titik. Kode format tetap ditulis `#,##0` (notasi internasional) dan Excel menerjemahkannya ke tampilan lokal.

### 3.4 Nilai vs tampilan: tanggal adalah angka serial

Excel menyimpan tanggal sebagai **angka serial**: jumlah hari sejak 1 Januari 1900 (hari ke-1). `2024-01-01` adalah hari ke-45292. Karena tanggal hanyalah angka, saya bisa mengurangkan dua tanggal untuk mendapat selisih hari (`=C16-C12`), menambahkan 30 untuk 30 hari ke depan, atau mengurutkannya. Jam disimpan sebagai pecahan hari (`0,5` = pukul 12.00).

Konsekuensi praktis:

- Jika kolom tanggal tampil seperti `45292`, nilainya benar; hanya formatnya General. Ubah format, jangan ketik ulang.
- Jika tanggal tampil `2024-01-01` tetapi rata kiri dan tidak bisa dikurangkan, kemungkinan itu **teks**, bukan tanggal. Modul 06 membahas konversinya (`DATEVALUE`).
- Persen juga sama: `10%` yang tersimpan adalah `0,1`. Mengetik `10` lalu memformat persen menghasilkan `1000%`. Jadi ketik `0,1` atau ketik `10%` langsung.
- Pembulatan tampilan: nilai `149741,333` yang diformat `#,##0` tampil `149.741`, tetapi perhitungan tetap memakai `149741,333`. Untuk benar-benar membulatkan nilai, pakai `ROUND` (modul 03).

### 3.5 Paste Special

`Ctrl + C` lalu `Home > Paste > Paste Special` (atau `Ctrl + Alt + V`) memberi pilihan:

- **Values**: menempel hanya nilai hasil, tanpa rumus dan tanpa format. Dipakai saat mengirim laporan agar rumus tidak rusak, atau membekukan hasil perhitungan.
- **Formats**: menempel format saja.
- **Transpose**: menukar baris menjadi kolom. Berguna untuk mengubah tabel ringkasan vertikal menjadi horizontal.
- **Values + Transpose** bisa dikombinasikan dalam satu dialog.

### 3.6 Sort, Filter, Freeze Panes

- **Sort**: klik satu sel di dalam tabel, `Data > Sort`. Pilih kolom dan arah (A to Z / Smallest to Largest). Excel otomatis memperluas pilihan ke seluruh blok data selama tidak ada baris/kolom kosong yang memutus. Jangan pernah memblok satu kolom saja lalu mengurutkan; baris akan tercampur (Excel memperingatkan dengan dialog "Sort Warning", pilih *Expand the selection*).
- **Filter**: `Data > Filter` atau `Ctrl + Shift + L`. Panah muncul di header. Pilih nilai, atau pakai *Number Filters* (misalnya *Greater Than Or Equal To*). Jumlah baris hasil filter terlihat di **status bar** kiri bawah ("x of y records found"). Jika saya memblok kolom angka pada hasil filter, status bar menampilkan `Sum`, `Average`, dan `Count` tanpa rumus apa pun.
- **Freeze Panes**: `View > Freeze Panes > Freeze Top Row` agar header tetap terlihat saat menggulir. Untuk membekukan baris 1 dan kolom A sekaligus, klik `B2` lalu `View > Freeze Panes > Freeze Panes`.

## 4. Dataset latihan

Folder `latihan/`:

| File | Isi |
|---|---|
| `data_01.csv` | 150 transaksi Toko Sumber Rejeki, 1 Januari sampai 31 Maret 2024 |
| `latihan_01.xlsx` | Sheet `Instruksi`, `Data` (belum diformat), `Latihan` (sel kuning untuk diisi) |
| `solusi_01.xlsx` | Versi terisi dan terverifikasi |
| `kunci_jawaban.md` | Jawaban latihan mandiri dan angka hasil |

Kolom sheet `Data` (A sampai J, baris 2 sampai 151):

| Kolom | Nama | Tipe | Keterangan |
|---|---|---|---|
| A | `no_transaksi` | teks | `TRX-0001` sampai `TRX-0150`, unik |
| B | `tanggal` | tanggal | di file latihan sengaja berformat General (tampil 45292 dst.) |
| C | `kasir` | teks | Rina, Dedi, Sinta, Agus |
| D | `kode_produk` | teks | P001 sampai P024 |
| E | `nama_produk` | teks | |
| F | `kategori` | teks | Sembako, Minuman, Makanan, Perawatan, Rumah Tangga |
| G | `qty` | bilangan bulat | 1 sampai 11 |
| H | `harga_satuan` | rupiah | |
| I | `diskon_pct` | pecahan | 0; 0,05; 0,1; 0,15 |
| J | `total` | rupiah | `qty x harga_satuan x (1 - diskon_pct)` |

## 5. Langkah praktik

Buka `latihan/latihan_01.xlsx`. Jika muncul bilah kuning *Protected View*, klik *Enable Editing*.

1. Klik tab sheet `Data`. Klik sel `B2`. Perhatikan Formula Bar menampilkan `45292`. Ini angka serial tanggal.
2. Klik header kolom `B` (huruf B di atas) untuk memilih seluruh kolom. Tekan `Ctrl + 1`, pilih kategori `Custom`, ketik `yyyy-mm-dd` pada kotak *Type*, klik OK. Kolom kini tampil `2024-01-01`.
3. Pilih kolom `H` dan `J` (klik H, tahan `Ctrl`, klik J). `Ctrl + 1 > Number`, Decimal places 0, centang *Use 1000 Separator*, OK.
4. Pilih kolom `I`. `Ctrl + 1 > Percentage`, Decimal places 0, OK. Nilai `0,1` kini tampil `10%`.
5. Klik `A1`. `View > Freeze Panes > Freeze Top Row`. Gulir ke bawah; header tetap terlihat.
6. Pindah ke sheet `Latihan` (`Ctrl + Page Down`). Klik `C4`, ketik `=COUNTA(Data!A2:A151)` lalu Enter. Cara lebih cepat: ketik `=COUNTA(`, klik tab `Data`, klik `A2`, tekan `Ctrl + Shift + panah bawah`, ketik `)`, Enter.
7. Isi `C5` sampai `C9` dengan `SUM`, `MAX`, `MIN`, `AVERAGE`, dan `SUM` kolom `qty` (lihat bagian 6).
8. Bagian B: di `C12` ketik `=Data!B2` dan format `yyyy-mm-dd`. Di `C13` ketik `=C12*1` dan biarkan format General; tampil `45292`. (Jika mengetik `=C12` saja, Excel otomatis menyalin format tanggal; tekan `Ctrl + Shift + ~` untuk memaksa General.)
9. Di `C14` ketik `0,1` lalu format `0%`. Di `C15` ketik `=C14*100` untuk membuktikan nilai sebenarnya 0,1. Isi `C16`, `C17`, `C18` sesuai label.
10. Bagian C: kembali ke sheet `Data`, tekan `Ctrl + Shift + L`. Klik panah di header `kategori`, hilangkan centang *Select All*, centang `Sembako`, OK. Baca status bar: "31 of 150 records found". Ketik `31` di `Latihan!C21`. Ulangi untuk kriteria lain (untuk `qty >= 5`: panah di header `qty` > *Number Filters* > *Greater Than Or Equal To* > 5). Untuk `C23`, blok `J2:J151` pada hasil filter Minuman dan baca `Sum` di status bar. Hapus filter dengan `Data > Clear` sebelum lanjut.
11. Bagian D: klik satu sel di kolom `total`, `Data > Sort Largest to Smallest`. Salin tiga `no_transaksi` teratas: blok `A2:A4`, `Ctrl + C`, klik `Latihan!C28`, `Ctrl + Alt + V > Values`, OK. Di `D28:D30` isi `=LARGE(Data!$J$2:$J$151;1)` dan seterusnya. Kembalikan urutan Data dengan `Ctrl + Z` atau Sort berdasarkan `no_transaksi`.
12. Bagian E: blok `B34:C39` di sheet `Latihan`, `Ctrl + C`, klik `E34`, `Ctrl + Alt + V`, centang *Values* dan *Transpose*, OK.
13. Simpan dengan `Ctrl + S`. Bandingkan dengan `solusi_01.xlsx`.

## 6. Contoh formula dan penjelasan tiap argumen

```
=COUNTA(Data!A2:A151)
```
`COUNTA` menghitung sel yang **tidak kosong** (teks maupun angka). `Data!A2:A151` adalah range di sheet lain: nama sheet, tanda seru, lalu alamat. Hasil: 150.

```
=SUM(Data!J2:J151)
```
Menjumlahkan semua angka pada range. Sel teks atau kosong diabaikan. Hasil: 22.461.150.

```
=MAX(Data!J2:J151)   =MIN(Data!J2:J151)   =AVERAGE(Data!J2:J151)
```
Nilai terbesar, terkecil, dan rata-rata aritmetika (jumlah dibagi banyaknya angka). `AVERAGE` mengabaikan sel kosong, tetapi menghitung sel berisi 0. Hasil: 1.120.000; 4.500; 149.741.

```
=C16-C12
```
Karena tanggal adalah angka serial, pengurangan dua tanggal menghasilkan selisih hari: 89.

```
=MONTH(C12)
```
Mengambil nomor bulan (1 sampai 12) dari sebuah tanggal. Hasil: 1.

```
=LARGE(Data!$J$2:$J$151;2)
```
Nilai terbesar ke-2 dari range. Tanda `$` mengunci range agar tidak bergeser saat rumus disalin ke bawah (dijelaskan tuntas di modul 03). Argumen kedua adalah peringkat.

```
=COUNTIFS(Data!$F$2:$F$151;"Sembako")
```
Dipakai di file solusi hanya sebagai pembanding hasil Filter: menghitung baris yang kategorinya "Sembako". Fungsi ini dipelajari di modul 04; di sini cukup dipahami bahwa hasilnya harus sama dengan angka di status bar.

## 7. Latihan mandiri

Kerjakan pada sheet `Data` yang sudah diformat (langkah 1 sampai 5).

1. **Dasar.** Berapa angka serial tanggal `2024-03-30`? Cek dengan mengetik tanggal itu di sel kosong lalu menekan `Ctrl + Shift + ~`.
2. **Dasar.** Dengan Filter, berapa jumlah transaksi kasir Dedi? Berapa total `qty` yang dijual Dedi (baca `Sum` di status bar)?
3. **Menengah.** Urutkan data berdasarkan `kategori` (A to Z) lalu `total` (Largest to Smallest) sekaligus memakai `Data > Sort > Add Level`. Transaksi apa yang berada di baris paling atas?
4. **Menengah.** Dengan Filter tanggal (*Date Filters > Between*), berapa total penjualan bulan Februari 2024? Kemudian hapus filter dan hitung total Januari, Februari, Maret; pastikan ketiganya berjumlah 22.461.150.
5. **Menengah.** Di sel kosong ketik `=Data!B2+30` dan format sebagai tanggal. Tanggal apa yang muncul dan mengapa penjumlahan ini sah?
6. **Lanjut.** Salin kolom `total` ke kolom `L` dengan Paste Special > Values, lalu format `L` sebagai `"Rp"#,##0`. Buktikan dengan `=SUM(L2:L151)` bahwa nilainya tidak berubah meski tampilannya berbeda.
7. **Lanjut.** Berapa transaksi yang memiliki `qty >= 10`? Berapa di antaranya yang mendapat diskon? Gunakan dua filter bersamaan.

## 8. Kunci jawaban

Lihat `latihan/kunci_jawaban.md`. Angka utama: 150 transaksi, total penjualan 22.461.150, transaksi terbesar 1.120.000 (`TRX-0068`), rata-rata 149.741.

## 9. Kesalahan umum dan cara memeriksa hasil

| Kesalahan | Gejala | Cara memeriksa dan memperbaiki |
|---|---|---|
| Mengurutkan satu kolom saja | nilai `total` tidak lagi cocok dengan `no_transaksi` | selalu klik satu sel di dalam tabel, bukan memblok satu kolom; cek baris 1 sebelum dan sesudah sort |
| Mengetik ulang tanggal yang tampil sebagai angka | tanggal berubah menjadi teks atau salah tahun | cukup ubah format; tanggal yang benar rata kanan dan bisa dikurangkan |
| Mengetik `10` lalu format persen | tampil `1000%` | ketik `0,1` atau `10%`; cek nilai di Formula Bar |
| `SUM` menghasilkan 0 | angka tersimpan sebagai teks (rata kiri, ada segitiga hijau) | modul 06; sementara pakai `Data > Text to Columns > Finish` pada kolom itu |
| Filter masih aktif saat menghitung | `SUM` di status bar lebih kecil dari seharusnya | periksa ikon corong di header; `Data > Clear` |
| Lupa Freeze Panes | salah membaca kolom saat baris ke-100 | `View > Freeze Panes` |
| Paste biasa saat mengirim laporan | rumus `#REF!` di komputer penerima | gunakan Paste Special > Values untuk laporan final |

Cara memeriksa hasil: setelah mengisi `C4:C9`, blok `Data!J2:J151` dan bandingkan `Sum`, `Average`, `Count` di status bar dengan rumus saya. Angkanya harus identik.

## 10. Checklist kompetensi sebelum lanjut

- [ ] Saya bisa menyebut alamat sel dari Name Box dan memilih blok data dengan `Ctrl + Shift + panah`.
- [ ] Saya bisa mengubah format tanggal, ribuan, dan persen tanpa mengubah nilai.
- [ ] Saya bisa menjelaskan angka serial tanggal dan membuktikannya dengan `Ctrl + Shift + ~`.
- [ ] Saya bisa memakai Sort multi-level dan Filter angka/tanggal, serta membaca status bar.
- [ ] Saya bisa memakai Paste Special Values dan Transpose.
- [ ] Saya bisa menulis `SUM`, `COUNTA`, `MAX`, `MIN`, `AVERAGE` yang merujuk sheet lain.
- [ ] Hasil `latihan_01.xlsx` saya cocok dengan `solusi_01.xlsx`.

## 11. Catatan kompatibilitas

- **Excel 2016/2019/365**: semua fitur modul ini tersedia. Tampilan ribbon sedikit berbeda, tetapi menu `Data > Sort`, `Data > Filter`, `View > Freeze Panes` sama.
- **LibreOffice Calc**: Paste Special ada di `Edit > Paste Special`; Freeze Panes di `View > Freeze Cells`. Sistem tanggal sama (serial 45292 untuk 2024-01-01).
- **Google Sheets**: Freeze di `View > Freeze`; Paste Special di `Edit > Paste special`. Angka serial tanggal sama. Status bar ringkasan ada di kanan bawah.
- Pemisah argumen: regional Indonesia memakai `;`, regional Inggris memakai `,`. Di dalam file `.xlsx` rumus selalu tersimpan dengan koma; Excel menerjemahkannya saat ditampilkan.
