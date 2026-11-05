#!/usr/bin/env python3
"""Menulis TUTORIAL.md dan latihan.md untuk tiap proyek dari isi terstruktur di bawah."""

from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

CATATAN_REGIONAL = ("> Catatan: semua rumus ditulis dengan pemisah titik koma (`;`) untuk Excel regional Indonesia. "
                    "Pada Excel regional Inggris ganti `;` menjadi `,`.")

PROYEK = []


def tambah(**k):
    PROYEK.append(k)


tambah(
    folder="01_database-karyawan", judul="Database Karyawan", tingkat="Pemula", file="01_database_karyawan.xlsx",
    gambar="01_data_karyawan.png",
    tujuan="Saya membangun tabel master karyawan (80 orang) yang rapi: Format as Table, header beku, dropdown departemen "
           "dan status, kolom usia serta masa kerja yang dihitung otomatis, dan sheet Ringkasan berisi COUNTA, COUNTIF, "
           "COUNTIFS, AVERAGE, MAX, MIN, serta chart komposisi. Hasil akhirnya adalah database yang siap dipakai filter, "
           "sort, dan menjadi sumber proyek-proyek berikutnya.",
    konsep=["Format as Table (Ctrl+T) dan nama tabel `tblKaryawan`", "Freeze Panes, Sort & Filter",
            "Data Validation tipe List (dropdown)", "Format tanggal `dd/mm/yyyy` dan rupiah `\"Rp\"#,##0`",
            "DATEDIF untuk usia dan masa kerja (tahun dan bulan)", "COUNTA, COUNTIF, COUNTIFS, AVERAGE, AVERAGEIF, MAX, MIN, SUM",
            "Conditional Formatting berbasis rumus (gaji di bawah UMK)", "Chart kolom dan kolom bertumpuk"],
    struktur=[("Data Karyawan", "Tabel master 80 karyawan, 17 kolom", "`=DATEDIF(E2;TODAY();\"Y\")`, `=DATEDIF(J2;TODAY();\"Y\")&\" thn \"&DATEDIF(J2;TODAY();\"YM\")&\" bln\"`"),
              ("Daftar", "Sumber dropdown departemen, status, jenis kelamin", "-"),
              ("Ringkasan", "Statistik ringkas dan tabel per departemen + 2 chart", "`=COUNTA('Data Karyawan'!B2:B1000)`, `=COUNTIFS(...;$A18;...;C$17)`, `=AVERAGEIF(...)`"),
              ("Petunjuk", "Cara pakai singkat", "-")],
    langkah=[
        ("Siapkan header", "Di sheet baru ketik header di A1:Q1: No, NIK, Nama Lengkap, Jenis Kelamin, Tanggal Lahir, Usia (Thn), Departemen, Jabatan, Status, Tanggal Masuk, Masa Kerja (Thn), Masa Kerja (Teks), Gaji Pokok, Pendidikan, Kota, Email, No. HP. Isi 5 baris data contoh dulu."),
        ("Format kolom", "Blok kolom E dan J > klik kanan > Format Cells > Date > `dd/mm/yyyy`. Blok kolom M > Format Cells > Custom > ketik `\"Rp\"#,##0`. Blok kolom B dan Q > Format Cells > Text agar NIK dan nomor HP tidak kehilangan nol di depan."),
        ("Usia dengan DATEDIF", "Ketik di F2: `=DATEDIF(E2;TODAY();\"Y\")`. Argumen: tanggal awal (lahir), tanggal akhir (hari ini), dan satuan `\"Y\"` = tahun penuh. DATEDIF tidak muncul di daftar saran Excel tetapi tetap berfungsi."),
        ("Masa kerja", "Ketik di K2: `=DATEDIF(J2;TODAY();\"Y\")` dan di L2: `=DATEDIF(J2;TODAY();\"Y\")&\" thn \"&DATEDIF(J2;TODAY();\"YM\")&\" bln\"`. Satuan `\"YM\"` menghitung sisa bulan setelah tahun penuh; `&` menggabungkan teks."),
        ("Jadikan tabel", "Klik sel A1, tekan Ctrl+T, centang My table has headers, OK. Di Table Design ganti Table Name menjadi `tblKaryawan` dan pilih gaya Medium 2. Rumus di baris pertama otomatis disalin ke baris baru."),
        ("Bekukan header", "Klik sel D2 lalu View > Freeze Panes > Freeze Panes. Baris 1 dan kolom A-C tetap terlihat saat digulir."),
        ("Sheet Daftar dan dropdown", "Buat sheet `Daftar`; isi A2:A7 dengan enam departemen dan B2:B4 dengan Tetap, Kontrak, Magang. Kembali ke Data Karyawan, blok G2:G81 > Data > Data Validation > Allow: List > Source: `=Daftar!$A$2:$A$7`. Ulangi untuk I2:I81 dengan `=Daftar!$B$2:$B$4` dan D2:D81 dengan `L,P` (ketik langsung)."),
        ("Sheet Ringkasan", "Buat sheet `Ringkasan`. B6: `=COUNTA('Data Karyawan'!B2:B1000)` (menghitung sel terisi = jumlah karyawan). B7: `=COUNTIF('Data Karyawan'!D2:D1000;\"L\")`. B9: `=AVERAGE('Data Karyawan'!M2:M1000)`. B10: `=MAX(...)`, B11: `=MIN(...)`, B12: `=SUM(...)`."),
        ("Tabel per departemen", "Di A17:G17 ketik header Departemen, Jumlah, Tetap, Kontrak, Magang, Rata-rata Gaji, % dari Total. A18:A23 isi nama departemen. B18: `=COUNTIF('Data Karyawan'!$G$2:$G$1000;A18)`. C18: `=COUNTIFS('Data Karyawan'!$G$2:$G$1000;$A18;'Data Karyawan'!$I$2:$I$1000;C$17)` lalu tarik ke E23; perhatikan `$A18` (kolom terkunci) dan `C$17` (baris terkunci). F18: `=IFERROR(AVERAGEIF('Data Karyawan'!$G$2:$G$1000;A18;'Data Karyawan'!$M$2:$M$1000);0)`. G18: `=IFERROR(B18/$B$24;0)` dengan B24 = `=SUM(B18:B23)`."),
        ("Conditional Formatting", "Di Ringkasan!B5 ketik UMK 4725479. Blok M2:M81 > Home > Conditional Formatting > New Rule > Use a formula > `=M2<Ringkasan!$B$5` > Format isi merah muda. Gaji di bawah UMK langsung tersorot."),
        ("Chart", "Blok A17:B23 > Insert > Column > Clustered Column; beri judul. Untuk komposisi status blok A17:A23 dan C17:E23 (tahan Ctrl) > Insert > Stacked Column."),
        ("Uji", "Ganti departemen satu karyawan lewat dropdown dan pastikan angka di Ringkasan berubah. Urutkan tabel dengan panah filter pada header Gaji Pokok > Sort Largest to Smallest."),
    ],
    kesalahan=["`#NUM!` pada DATEDIF: tanggal akhir lebih kecil dari tanggal awal; periksa tanggal lahir/masuk.",
               "NIK berubah menjadi 2019004 (nol hilang): format kolom belum Text sebelum mengetik.",
               "COUNTIF menghasilkan 0 padahal data ada: ada spasi di akhir teks (`HR `); gunakan TRIM atau dropdown agar ejaan seragam.",
               "Rumus tarik ke bawah merujuk baris salah: kunci rentang dengan `$` (tekan F4 saat mengedit).",
               "Tanggal tampil sebagai angka 45321: format sel belum Date."],
    variasi=["Ganti `TODAY()` dengan sel tanggal acuan agar laporan tidak berubah tiap hari (dipakai di proyek 08).",
             "Tambahkan kolom Email otomatis: `=LOWER(SUBSTITUTE(C2;\" \";\".\"))&\"@perusahaan.co.id\"`.",
             "Lanjut ke proyek 02 untuk menghubungkan data ini dengan absensi."],
    latihan=[("Hitung jumlah karyawan tetap di departemen Produksi.", "`=COUNTIFS('Data Karyawan'!G2:G81;\"Produksi\";'Data Karyawan'!I2:I81;\"Tetap\")`"),
             ("Hitung rata-rata usia karyawan perempuan.", "`=AVERAGEIF('Data Karyawan'!D2:D81;\"P\";'Data Karyawan'!F2:F81)`"),
             ("Tampilkan nama karyawan dengan gaji tertinggi.", "`=INDEX('Data Karyawan'!C2:C81;MATCH(MAX('Data Karyawan'!M2:M81);'Data Karyawan'!M2:M81;0))`"),
             ("Buat kolom Kategori Masa Kerja: < 2 thn = Baru, 2-5 = Madya, > 5 = Senior.", "`=IF(K2<2;\"Baru\";IF(K2<=5;\"Madya\";\"Senior\"))`"),
             ("Hitung berapa karyawan yang masuk pada tahun 2023.", "`=COUNTIFS('Data Karyawan'!J2:J81;\">=\"&DATE(2023;1;1);'Data Karyawan'!J2:J81;\"<=\"&DATE(2023;12;31))`")],
)

tambah(
    folder="02_rekap-absensi-bulanan", judul="Rekap Absensi Bulanan", tingkat="Pemula", file="02_rekap_absensi_bulanan.xlsx",
    gambar="02_absensi.png",
    tujuan="Saya membuat lembar absensi harian Maret 2024 untuk 60 karyawan dengan kode H/S/I/A/C, rekap otomatis per karyawan "
           "(COUNTIF), persentase kehadiran yang disorot merah bila di bawah 90%, dan rekap per departemen dengan SUMIF beserta chart. "
           "Hasilnya adalah laporan absensi bulanan yang siap dikirim ke manajemen.",
    konsep=["COUNTIF pada rentang horizontal (per baris)", "COUNT untuk menghitung hari kerja dari header tanggal",
            "Persentase dan format `0.0%`", "Conditional Formatting: Cell Is less than, Equal To (warna per kode)",
            "Data Validation daftar kode", "SUMIF dan COUNTIF per departemen", "RANK.EQ + COUNTIF untuk peringkat unik",
            "INDEX/MATCH untuk daftar 10 kehadiran terendah", "Chart kolom dan bertumpuk", "TEXT(tanggal;\"ddd\") untuk nama hari"],
    struktur=[("Absensi", "60 baris x 18 hari kerja + kolom rekap", "`=COUNTIF($E5:$V5;\"H\")`, `=COUNT($E$4:$V$4)`, `=IFERROR(W5/AB5;0)`"),
              ("Rekap Departemen", "Total per departemen, 10 terendah, 2 chart", "`=SUMIF(Absensi!$C$5:$C$64;A10;Absensi!$W$5:$W$64)`"),
              ("Kode & Libur", "Arti kode dan daftar libur 2024", "-"), ("Petunjuk", "Cara pakai", "-")],
    langkah=[
        ("Header tanggal", "Baris 4: A4 NIK, B4 Nama, C4 Departemen, D4 Jabatan. Di E4 ketik 01/03/2024, lalu isi tanggal hari kerja berikutnya ke kanan (lewati Sabtu, Minggu, 11-12 Maret, 29 Maret). Format E4:V4 Custom `dd` agar hanya tanggal yang tampil. Di E3 ketik `=TEXT(E4;\"ddd\")` dan tarik ke kanan untuk nama hari."),
        ("Isi kode", "Blok E5:V64 > Data Validation > List > Source `H,S,I,A,C`. Isi kode per hari. Home > Alignment > Center."),
        ("Rekap per baris", "W5: `=COUNTIF($E5:$V5;\"H\")`; X5 untuk S, Y5 untuk I, Z5 untuk A, AA5 untuk C. `$E5:$V5` mengunci kolom sehingga rumus bisa disalin ke bawah tetapi tidak bergeser ke samping."),
        ("Hari kerja dan persentase", "AB5: `=COUNT($E$4:$V$4)` menghitung tanggal di header (angka). AC5: `=IFERROR(W5/AB5;0)` lalu format `0.0%`. IFERROR menjaga agar baris kosong tidak menampilkan #DIV/0!."),
        ("Keterangan otomatis", "AD5: `=IF(Z5>=2;\"Peringatan alpa\";IF(AC5<0,9;\"Kehadiran rendah\";\"Baik\"))`. IF bertingkat: cek alpa dulu, lalu persentase."),
        ("Peringkat terendah", "AE5: `=RANK.EQ(AC5;$AC$5:$AC$64;1)+COUNTIF($AC$5:AC5;AC5)-1`. Argumen ketiga 1 = urutan naik (terendah = peringkat 1); COUNTIF dengan rentang yang memanjang memecah nilai kembar."),
        ("Conditional Formatting", "Blok AC5:AC64 > Conditional Formatting > Highlight Cells Rules > Less Than > `0,9` > isi merah. Blok E5:V64 > Equal To > `A` > merah; ulangi untuk S (kuning), I (biru), C (hijau)."),
        ("Freeze dan filter", "Klik E5 > View > Freeze Panes. Blok A4:AE64 > Data > Filter."),
        ("Rekap departemen", "Sheet baru. B4: `=Absensi!AB5` (hari kerja). A10:A15 nama departemen. B10: `=COUNTIF(Absensi!$C$5:$C$64;A10)`. C10: `=SUMIF(Absensi!$C$5:$C$64;A10;Absensi!$W$5:$W$64)` (rentang kriteria, kriteria, rentang yang dijumlah). H10: `=B10*$B$4`; I10: `=IFERROR(C10/H10;0)`."),
        ("Daftar 10 terendah", "A38: 1 sampai 10. B38: `=INDEX(Absensi!$B$5:$B$64;MATCH(A38;Absensi!$AE$5:$AE$64;0))`. MATCH mencari peringkat ke-n di kolom bantu, INDEX mengambil nama pada posisi itu."),
        ("Chart", "Blok A9:A15 dan I9:I15 (Ctrl) > Insert > Clustered Column. Blok A9:A15 dan D9:G15 > Stacked Column untuk komposisi ketidakhadiran."),
    ],
    kesalahan=["COUNTIF menghitung 0: kode diketik huruf kecil `h` tanpa dropdown; COUNTIF tidak peka kapital, jadi lebih mungkin ada spasi.",
               "Hari kerja salah: COUNT hanya menghitung angka/tanggal; header yang diketik sebagai teks tidak ikut.",
               "Persentase menampilkan 94,4 bukan 94,4%: hasil pembagian belum diformat persen (jangan dikalikan 100).",
               "Conditional Formatting Less Than 90 bukan 0,9: nilai persen disimpan sebagai pecahan.",
               "Peringkat ganda: lupa menambah `COUNTIF($AC$5:AC5;AC5)-1`."],
    variasi=["Tambahkan kolom jam masuk/pulang dan hitung keterlambatan dengan `=IF(F5>TIME(8;0;0);\"Terlambat\";\"\")`.",
             "Buat satu sheet per bulan dan rekap tahunan dengan SUMIF lintas sheet.",
             "Gunakan PivotTable (proyek 06) untuk rekap kode per departemen."],
    latihan=[("Hitung total hari sakit seluruh perusahaan.", "`=SUM(Absensi!X5:X64)`"),
             ("Berapa karyawan dengan kehadiran 100%?", "`=COUNTIF(Absensi!AC5:AC64;1)`"),
             ("Hitung persentase alpa departemen Produksi.", "`=SUMIF(Absensi!C5:C64;\"Produksi\";Absensi!Z5:Z64)/(COUNTIF(Absensi!C5:C64;\"Produksi\")*Absensi!AB5)`"),
             ("Tandai baris karyawan yang alpa pada hari Senin pertama (kolom F) dengan Conditional Formatting.", "Aturan rumus pada A5:AE64: `=$F5=\"A\"`"),
             ("Tampilkan nama karyawan dengan hari sakit terbanyak.", "`=INDEX(Absensi!B5:B64;MATCH(MAX(Absensi!X5:X64);Absensi!X5:X64;0))`")],
)

tambah(
    folder="03_administrasi-surat-dan-atk", judul="Administrasi Surat dan ATK", tingkat="Pemula", file="03_administrasi_surat_dan_atk.xlsx",
    gambar="03_kartu_stok.png",
    tujuan="Saya menyusun dua modul administrasi dalam satu workbook: agenda surat masuk/keluar dengan nomor otomatis, status disposisi, "
           "dan filter; serta kartu stok ATK yang menghitung stok akhir dari transaksi (SUMIFS), memberi peringatan stok minimum, dan "
           "laporan pemakaian bulanan. Hasilnya adalah sistem administrasi sederhana yang biasa dipakai staf admin/HR.",
    konsep=["TEXT + ROW untuk nomor agenda otomatis", "CHOOSE + MONTH untuk bulan romawi", "INDEX/MATCH untuk menarik nama dan harga barang",
            "SUMIFS dengan 2-4 kriteria (kode, jenis, rentang tanggal)", "IF untuk status PESAN/Aman", "COUNTIF/COUNTIFS untuk ringkasan surat",
            "Data Validation dan Conditional Formatting status", "Format as Table, AutoFilter, Freeze Panes", "Chart batang horizontal"],
    struktur=[("Surat Masuk", "60 surat, nomor agenda otomatis, disposisi", "`=\"SM-\"&TEXT(B5;\"yyyy\")&\"-\"&TEXT(ROW()-4;\"000\")`, `=IF(H5=\"Selesai\";I5-B5;TODAY()-B5)`"),
              ("Surat Keluar", "45 surat, nomor NNN/HRD/ROMAWI/TTTT", "`=TEXT(A5;\"000\")&\"/HRD/\"&CHOOSE(MONTH(B5);\"I\";...)&\"/\"&YEAR(B5)`"),
              ("Ringkasan Surat", "COUNTIF status, per bulan, per disposisi + chart", "`=COUNTIFS('Surat Masuk'!$B$5:$B$500;\">=\"&G13;...;\"<\"&H13)`"),
              ("Master ATK", "15 barang, stok awal, minimum, harga", "-"),
              ("Transaksi ATK", "±120 transaksi masuk/keluar", "`=INDEX('Master ATK'!$B$5:$B$19;MATCH(B5;'Master ATK'!$A$5:$A$19;0))`"),
              ("Kartu Stok", "Stok akhir, status, nilai persediaan", "`=SUMIFS('Transaksi ATK'!$E$5:$E$500;'Transaksi ATK'!$B$5:$B$500;A5;'Transaksi ATK'!$D$5:$D$500;\"Masuk\")`"),
              ("Laporan Bulanan ATK", "Pemakaian per barang per bulan", "SUMIFS dengan rentang tanggal")],
    langkah=[
        ("Agenda surat masuk", "Baris 4 header: No. Agenda, Tanggal Terima, Nomor Surat Asal, Pengirim, Perihal, Jenis, Disposisi Ke, Status Disposisi, Tanggal Selesai, Lama Proses, Keterangan. A5: `=\"SM-\"&TEXT(B5;\"yyyy\")&\"-\"&TEXT(ROW()-4;\"000\")`. ROW() memberi nomor baris (5), dikurangi 4 menjadi 1, TEXT mengubahnya menjadi `001`."),
        ("Lama proses", "J5: `=IF(H5=\"Selesai\";I5-B5;IF(B5=\"\";\"\";TODAY()-B5))`. Selisih dua tanggal langsung menghasilkan jumlah hari."),
        ("Dropdown dan warna", "Buat sheet Daftar (jenis surat, departemen, status). Data Validation untuk F, G, H. Conditional Formatting > Equal To `Belum` merah, `Proses` kuning, `Selesai` hijau. Tambahkan aturan rumus `=AND($H5<>\"Selesai\";$J5>7)` pada J5:J64 untuk surat yang menggantung lebih dari 7 hari."),
        ("Jadikan tabel", "Ctrl+T pada A4:K64, nama `tblSuratMasuk`. Klik C5 > Freeze Panes."),
        ("Nomor surat keluar", "Di sheet Surat Keluar A5: `=ROW()-4`; C5: `=TEXT(A5;\"000\")&\"/HRD/\"&CHOOSE(MONTH(B5);\"I\";\"II\";\"III\";\"IV\";\"V\";\"VI\";\"VII\";\"VIII\";\"IX\";\"X\";\"XI\";\"XII\")&\"/\"&YEAR(B5)`. CHOOSE memilih item ke-n sesuai nomor bulan."),
        ("Master dan transaksi ATK", "Master ATK: Kode, Nama Barang, Satuan, Stok Awal, Stok Minimum, Harga Satuan. Transaksi ATK: pilih kode lewat dropdown `='Master ATK'!$A$5:$A$19`; C5: `=IFERROR(INDEX('Master ATK'!$B$5:$B$19;MATCH(B5;'Master ATK'!$A$5:$A$19;0));\"\")`. MATCH menemukan posisi kode, INDEX mengambil nama pada posisi itu. H5: `=E5*G5`."),
        ("Kartu stok", "E5: `=SUMIFS('Transaksi ATK'!$E$5:$E$500;'Transaksi ATK'!$B$5:$B$500;A5;'Transaksi ATK'!$D$5:$D$500;\"Masuk\")` (jumlahkan kolom Jumlah bila kode cocok DAN jenis Masuk). F5 sama dengan \"Keluar\". G5: `=D5+E5-F5`. I5: `=IF(G5<=H5;\"PESAN\";\"Aman\")`. K5: `=G5*J5`."),
        ("Laporan bulanan", "Baris 3 berisi tanggal awal tiap bulan (C3 = 01/01/2024, D3 = 01/02/2024, ...). C5: `=SUMIFS('Transaksi ATK'!$E$5:$E$500;'Transaksi ATK'!$B$5:$B$500;$A5;'Transaksi ATK'!$D$5:$D$500;\"Keluar\";'Transaksi ATK'!$A$5:$A$500;\">=\"&C$3;'Transaksi ATK'!$A$5:$A$500;\"<\"&D$3)`. Kriteria tanggal digabung dengan `&` karena operator harus berupa teks."),
        ("Ringkasan surat", "B5: `=COUNTIF('Surat Masuk'!H5:H500;\"Belum\")`. Tabel per bulan memakai COUNTIFS dengan dua batas tanggal seperti langkah 8."),
        ("Chart", "Kartu Stok: blok B4:B19 dan G4:H19 > Insert > Clustered Column (stok akhir vs minimum). Laporan: blok B4:B19 dan J4:J19 > Bar."),
    ],
    kesalahan=["Nomor agenda berubah saat baris dihapus: ROW() dinamis; untuk arsip final salin kolom lalu Paste Values.",
               "`#N/A` pada INDEX/MATCH: kode belum ada di Master atau ada spasi; bungkus IFERROR.",
               "SUMIFS menghasilkan 0 pada rentang tanggal: kriteria ditulis `\">=C3\"` (teks literal) bukan `\">=\"&C3`.",
               "Stok akhir negatif: transaksi keluar dicatat sebelum barang masuk; periksa urutan tanggal.",
               "Rentang SUMIFS tidak sama panjang: semua rentang kriteria dan jumlah harus berukuran sama."],
    variasi=["Tambahkan kolom lampiran (hyperlink ke file PDF) dengan `=HYPERLINK(\"arsip\\\"&A5&\".pdf\";\"Buka\")`.",
             "Buat nomor otomatis yang tidak bergeser dengan `=MAX($A$4:A4)+1`.",
             "Gunakan PivotTable untuk laporan pemakaian per departemen pemakai."],
    latihan=[("Hitung jumlah surat masuk jenis Undangan yang belum selesai.", "`=COUNTIFS('Surat Masuk'!F5:F64;\"Undangan\";'Surat Masuk'!H5:H64;\"<>Selesai\")`"),
             ("Total nilai pembelian ATK (jenis Masuk) bulan Februari 2024.", "`=SUMIFS('Transaksi ATK'!H5:H500;'Transaksi ATK'!D5:D500;\"Masuk\";'Transaksi ATK'!A5:A500;\">=\"&DATE(2024;2;1);'Transaksi ATK'!A5:A500;\"<\"&DATE(2024;3;1))`"),
             ("Nama barang dengan nilai persediaan terbesar.", "`=INDEX('Kartu Stok'!B5:B19;MATCH(MAX('Kartu Stok'!K5:K19);'Kartu Stok'!K5:K19;0))`"),
             ("Buat nomor surat keluar dengan kode departemen dari dropdown di kolom K (misal KEU).", "`=TEXT(A5;\"000\")&\"/\"&K5&\"/\"&CHOOSE(MONTH(B5);\"I\";...;\"XII\")&\"/\"&YEAR(B5)`"),
             ("Jumlah barang yang statusnya PESAN.", "`=COUNTIF('Kartu Stok'!I5:I19;\"PESAN\")`")],
)

tambah(
    folder="04_penggajian-dan-slip-gaji", judul="Penggajian dan Slip Gaji", tingkat="Menengah", file="04_penggajian_dan_slip_gaji.xlsx",
    gambar="04_slip_gaji.png",
    tujuan="Saya membangun sistem payroll bulanan (70 karyawan) dari parameter terpusat: gaji pokok, tunjangan jabatan (VLOOKUP), tunjangan "
           "makan/transport berbasis kehadiran, lembur (jam x gaji/173 x 1,5), potongan BPJS Kesehatan 1%, JHT 2%, JP 1%, PPh 21 metode TER "
           "yang disederhanakan, rekap per departemen dengan SUMIFS, dan slip gaji yang berubah lewat dropdown nama (INDEX/MATCH).",
    konsep=["Sheet Parameter sebagai satu sumber asumsi", "VLOOKUP exact (FALSE) dan approximate (TRUE) untuk tarif berjenjang",
            "INDEX/MATCH lintas sheet (alternatif XLOOKUP di Excel 365)", "SUMIF untuk lembur, SUMIFS untuk rekap departemen",
            "MIN untuk batas atas upah BPJS, ROUND untuk pembulatan rupiah", "IF bertingkat untuk memilih tabel TER A/B/C",
            "Data Validation dropdown nama pada slip", "Desain slip gaji siap cetak"],
    struktur=[("Parameter", "Hari kerja, tarif tunjangan, iuran BPJS, tunjangan per level, PTKP, tabel TER", "sel input kuning"),
              ("Master Karyawan", "70 karyawan: PTKP, NPWP, gaji pokok", "`=VLOOKUP(E5;Parameter!$A$20:$B$23;2;FALSE)`"),
              ("Kehadiran & Lembur", "Hari hadir, jam lembur, tarif lembur", "`=ROUND(F5/Parameter!$B$9;0)`, `=E5*G5*Parameter!$B$10`"),
              ("Perhitungan Gaji", "Bruto, potongan, PPh 21, THP", "`=ROUND(MIN(E5+F5;Parameter!$B$12)*Parameter!$B$11;0)`, `=IF(N5=\"A\";VLOOKUP(J5;Parameter!$D$20:$E$44;2;TRUE);IF(N5=\"B\";...))`"),
              ("Rekap Departemen", "SUMIFS per departemen + chart", "`=SUMIFS('Perhitungan Gaji'!$J$5:$J$74;'Perhitungan Gaji'!$C$5:$C$74;$A5)`"),
              ("Slip Gaji", "Slip dinamis berdasarkan dropdown C6", "`=INDEX('Perhitungan Gaji'!$E$5:$E$74;MATCH($C$6;'Perhitungan Gaji'!$B$5:$B$74;0))`")],
    langkah=[
        ("Parameter", "Buat sheet Parameter. Isi B5 hari kerja (18), B7 tunjangan makan/hari (25000), B8 transport (20000), B9 pembagi lembur (173), B10 pengali (1,5), B11 iuran BPJS Kesehatan (1%), B12 batas upah (12000000), B13 JHT (2%), B14 JP (1%), B15 batas JP (10042300), B16 tambahan tanpa NPWP (20%). Warnai kuning. Di A20:B23 tabel level -> tunjangan jabatan. Di A27:C34 status PTKP -> kategori TER. Di D20:E44, G20:H40, J20:K41 tabel TER A, B, C (batas bawah bruto, tarif) diurutkan naik."),
        ("Master Karyawan", "Kolom: NIK, Nama, Departemen, Jabatan, Level, Status PTKP, NPWP, Gaji Pokok, Tunjangan Jabatan, Bank, Rekening. I5: `=VLOOKUP(E5;Parameter!$A$20:$B$23;2;FALSE)` (cari level di kolom pertama tabel, ambil kolom ke-2, FALSE = harus persis). Dropdown Level dan PTKP merujuk Parameter."),
        ("Kehadiran & Lembur", "A5: `='Master Karyawan'!A5`. B5: `=INDEX('Master Karyawan'!$B$5:$B$74;MATCH(A5;'Master Karyawan'!$A$5:$A$74;0))`. D5 hari hadir dan E5 jam lembur diisi manual. F5 gaji pokok via INDEX/MATCH. G5: `=ROUND(F5/Parameter!$B$9;0)` (upah per jam = gaji/173). H5: `=E5*G5*Parameter!$B$10`."),
        ("Perhitungan Gaji - penerimaan", "E5 gaji pokok, F5 tunjangan jabatan (INDEX/MATCH ke Master). G5: `=INDEX('Kehadiran & Lembur'!$D$5:$D$74;MATCH($A5;'Kehadiran & Lembur'!$A$5:$A$74;0))*Parameter!$B$7`. H5 sama dengan `$B$8`. I5: `=SUMIF('Kehadiran & Lembur'!$A$5:$A$74;$A5;'Kehadiran & Lembur'!$H$5:$H$74)`. J5: `=SUM(E5:I5)`."),
        ("Potongan BPJS", "K5: `=ROUND(MIN(E5+F5;Parameter!$B$12)*Parameter!$B$11;0)`. MIN membatasi dasar upah pada batas atas. L5: `=ROUND((E5+F5)*Parameter!$B$13;0)`. M5: `=ROUND(MIN(E5+F5;Parameter!$B$15)*Parameter!$B$14;0)`."),
        ("PPh 21 TER", "N5 kategori: `=INDEX(Parameter!$B$27:$B$34;MATCH(INDEX('Master Karyawan'!$F$5:$F$74;MATCH($A5;'Master Karyawan'!$A$5:$A$74;0));Parameter!$A$27:$A$34;0))`. O5 tarif: `=IF(N5=\"A\";VLOOKUP(J5;Parameter!$D$20:$E$44;2;TRUE);IF(N5=\"B\";VLOOKUP(J5;Parameter!$G$20:$H$40;2;TRUE);VLOOKUP(J5;Parameter!$J$20:$K$41;2;TRUE)))`. TRUE = pencarian perkiraan: mengambil batas bawah terbesar yang <= bruto. P5: `=ROUND(J5*O5*IF(NPWP=\"Tidak\";1+Parameter!$B$16;1);0)`."),
        ("Penyederhanaan yang saya lakukan", "TER bulanan dipakai Januari-November; Desember seharusnya dihitung ulang setahun dengan tarif progresif Pasal 17 dikurangi PPh yang sudah dipotong. Tabel TER di sini cuplikan yang dibulatkan; lembur memakai pengali rata 1,5 (aturan resmi: 1,5x jam pertama, 2x jam berikutnya); iuran ditanggung perusahaan tidak dihitung."),
        ("THP dan rekap", "Q5: `=SUM(K5:M5)+P5`; R5: `=J5-Q5`. Sheet Rekap Departemen: C5: `=SUMIFS('Perhitungan Gaji'!$E$5:$E$74;'Perhitungan Gaji'!$C$5:$C$74;$A5)` dan seterusnya; chart bruto vs bersih."),
        ("Slip gaji", "Buat sheet Slip Gaji, gabungkan B2:F2 untuk nama perusahaan. C6 dropdown: Data Validation > List > `='Master Karyawan'!$B$5:$B$74`. C8: `=INDEX('Master Karyawan'!$A$5:$A$74;MATCH($C$6;'Master Karyawan'!$B$5:$B$74;0))`. Komponen: `=INDEX('Perhitungan Gaji'!$E$5:$E$74;MATCH($C$6;'Perhitungan Gaji'!$B$5:$B$74;0))` dan seterusnya per kolom. F22: `=C20-F20`. Di Excel 365 bisa ditulis `=XLOOKUP($C$6;'Perhitungan Gaji'!$B$5:$B$74;'Perhitungan Gaji'!$E$5:$E$74)`."),
        ("Cetak", "Page Layout > Print Area > Set Print Area pada B2:F32; Margins Narrow; Scale to Fit 1 page. Ctrl+P."),
    ],
    kesalahan=["VLOOKUP TRUE salah tarif: tabel TER tidak diurutkan naik pada kolom pertama.",
               "`#N/A` pada slip: nama di C6 diketik berbeda dari Master (gunakan dropdown).",
               "Potongan BPJS kebesaran: lupa MIN batas atas upah.",
               "Hasil THP desimal: lupa ROUND pada tiap komponen.",
               "Mengubah tarif di dalam rumus, bukan di Parameter, sehingga sebagian baris berbeda."],
    variasi=["Tambahkan perhitungan Desember (setahun) dengan tarif progresif Pasal 17 dan PTKP tahunan.",
             "Buat makro VBA untuk ekspor slip PDF massal (lihat proyek 10, `vba/SlipGajiPDF.bas`).",
             "Tambahkan sheet jurnal akuntansi (beban gaji, utang BPJS, utang PPh 21)."],
    latihan=[("Total PPh 21 departemen IT.", "`=SUMIFS('Perhitungan Gaji'!P5:P74;'Perhitungan Gaji'!C5:C74;\"IT\")`"),
             ("Berapa karyawan yang gaji bersihnya di atas Rp10 juta?", "`=COUNTIF('Perhitungan Gaji'!R5:R74;\">10000000\")`"),
             ("Ubah tunjangan makan menjadi Rp30.000 dan amati sel mana saja yang berubah.", "Hanya Parameter!B7 yang diubah; kolom G, J, P, Q, R serta Rekap ikut berubah."),
             ("Tulis ulang rumus tunjangan jabatan dengan INDEX/MATCH.", "`=INDEX(Parameter!$B$20:$B$23;MATCH(E5;Parameter!$A$20:$A$23;0))`"),
             ("Hitung rata-rata jam lembur karyawan Produksi.", "`=AVERAGEIFS('Kehadiran & Lembur'!E5:E74;'Perhitungan Gaji'!C5:C74;\"Produksi\")`")],
)

tambah(
    folder="05_pengelolaan-cuti", judul="Pengelolaan Cuti", tingkat="Menengah", file="05_pengelolaan_cuti.xlsx",
    gambar="05_pengajuan_cuti.png",
    tujuan="Saya membuat sistem cuti tahunan 2024 untuk 60 karyawan: hak cuti 12 hari, log pengajuan dengan hari kerja yang dihitung "
           "NETWORKDAYS.INTL memakai daftar libur nasional, saldo berjalan, validasi yang menolak pengajuan melebihi saldo, rekap per "
           "departemen/bulan, dan kalender ringkas yang menampilkan jumlah orang cuti per tanggal.",
    konsep=["NETWORKDAYS.INTL dengan argumen hari libur", "DATEDIF untuk masa kerja dalam bulan", "SUMIFS dengan rentang yang memanjang (saldo berjalan)",
            "IF bertingkat untuk validasi", "Data Validation tipe Date (selesai >= mulai)", "DATE, WEEKDAY, MONTH, EOMONTH untuk grid kalender",
            "COUNTIFS dengan tanggal (`\"<=\"&tgl`)", "Conditional Formatting rumus pada kalender", "TEXT(tanggal;\"dddd\")"],
    struktur=[("Hari Libur 2024", "27 tanggal libur/cuti bersama", "`=NETWORKDAYS.INTL(DATE(2024;1;1);DATE(2024;12;31);1;A5:A31)`"),
              ("Saldo Cuti", "Hak, terpakai, sisa per karyawan", "`=IF(E5>=12;$H$2;0)`, `=SUMIFS('Pengajuan Cuti'!$I$5:$I$500;...;\"Disetujui\")`"),
              ("Pengajuan Cuti", "±85 pengajuan dengan validasi", "`=NETWORKDAYS.INTL(F5;G5;1;'Hari Libur 2024'!$A$5:$A$31)`, `=IF(I5>K5;\"MELEBIHI SALDO\";\"OK\")`"),
              ("Rekap Cuti", "Per departemen dan per bulan + chart", "COUNTIFS/SUMIFS dengan rentang tanggal"),
              ("Kalender", "Grid bulan terpilih (C4) dan daftar yang cuti", "`=IF(MONTH($C$6+7)=$C$4;$C$6+7;\"\")`, `=COUNTIFS(mulai;\"<=\"&tgl;selesai;\">=\"&tgl;status;\"Disetujui\")`")],
    langkah=[
        ("Daftar libur", "Sheet Hari Libur 2024: A5:A31 tanggal, B5 `=TEXT(A5;\"dddd\")`, C keterangan, D5 `=IF(WEEKDAY(A5;2)<=5;\"Ya\";\"Tidak (akhir pekan)\")`. WEEKDAY tipe 2: Senin = 1 ... Minggu = 7."),
        ("Saldo cuti", "Sheet Saldo Cuti; H2 = 12 (hak penuh), H3 = 01/01/2024. E5: `=IF(D5>$H$3;0;DATEDIF(D5;$H$3;\"M\"))` (bulan masa kerja; IF mencegah #NUM! untuk karyawan masuk setelah tanggal acuan). F5: `=IF(E5>=12;$H$2;0)`."),
        ("Log pengajuan", "Sheet Pengajuan Cuti header: No, Tanggal Pengajuan, NIK, Nama, Jenis Cuti, Mulai, Selesai, Hari Kalender, Hari Kerja, Status, Saldo Tersedia, Validasi, Alasan, Disetujui Oleh. D5: `=IFERROR(INDEX('Saldo Cuti'!$B$5:$B$64;MATCH(C5;'Saldo Cuti'!$A$5:$A$64;0));\"\")`. H5: `=G5-F5+1`."),
        ("Hari kerja cuti", "I5: `=NETWORKDAYS.INTL(F5;G5;1;'Hari Libur 2024'!$A$5:$A$31)`. Argumen ke-3 = 1 berarti akhir pekan Sabtu-Minggu (gunakan 11 untuk hanya Minggu libur); argumen ke-4 daftar libur yang dikecualikan."),
        ("Saldo berjalan", "K5: `=IFERROR(INDEX('Saldo Cuti'!$F$5:$F$64;MATCH(C5;'Saldo Cuti'!$A$5:$A$64;0));0)`. K6: hak cuti dikurangi `SUMIFS($I$5:I5;$C$5:C5;C6;$E$5:E5;\"Tahunan\";$J$5:J5;\"Disetujui\")`. Rentang `$I$5:I5` memanjang saat disalin ke bawah sehingga hanya menghitung pengajuan sebelumnya."),
        ("Validasi", "L5: `=IF(E5<>\"Tahunan\";\"Tidak memakai saldo\";IF(I5>K5;\"MELEBIHI SALDO\";IF(F5<B5;\"Tanggal mulai sebelum pengajuan\";\"OK\")))`. Conditional Formatting Equal To `MELEBIHI SALDO` merah. Data Validation pada G5:G500: Allow Date, greater than or equal to `=F5`."),
        ("Cuti terpakai", "Kembali ke Saldo Cuti G5: `=SUMIFS('Pengajuan Cuti'!$I$5:$I$500;'Pengajuan Cuti'!$C$5:$C$500;A5;'Pengajuan Cuti'!$E$5:$E$500;\"Tahunan\";'Pengajuan Cuti'!$J$5:$J$500;\"Disetujui\")`. I5: `=F5-G5`. Conditional Formatting Less Than 3 pada I."),
        ("Kalender", "Sheet Kalender: C4 bulan (dropdown 1-12), E4 tahun. C5: `=DATE($E$4;$C$4;1)`. C6: `=C5-WEEKDAY(C5;2)+1` (Senin pada minggu pertama). B10: `=IF(MONTH($C$6+0)=$C$4;$C$6+0;\"\")`, C10 dengan +1, ... sampai H10 (+6); baris 12 mulai +7, dan seterusnya 6 minggu. Format `d`."),
        ("Jumlah orang cuti per hari", "B11: `=IF(B10=\"\";\"\";COUNTIFS('Pengajuan Cuti'!$F$5:$F$500;\"<=\"&B10;'Pengajuan Cuti'!$G$5:$G$500;\">=\"&B10;'Pengajuan Cuti'!$J$5:$J$500;\"Disetujui\"))`: hitung pengajuan yang mulai <= tanggal dan selesai >= tanggal. Conditional Formatting rumus `=AND(ISNUMBER(B11);B11>=3)` merah."),
        ("Rekap", "Sheet Rekap Cuti: per departemen SUMIF ke Saldo Cuti; per bulan COUNTIFS/SUMIFS dengan `\">=\"&H14` dan `\"<\"&I14` (awal bulan dan awal bulan berikut). Chart kolom hari cuti per bulan."),
    ],
    kesalahan=["NETWORKDAYS.INTL mengabaikan libur: rentang libur berisi teks, bukan tanggal.",
               "Saldo berjalan salah: rentang SUMIFS ditulis `$I$5:$I$500` (semua) bukan `$I$5:I5` (memanjang).",
               "Kalender menampilkan tanggal bulan lain: lupa membungkus dengan IF(MONTH(...)=$C$4).",
               "`#NUM!` DATEDIF untuk karyawan baru: tanggal masuk setelah tanggal acuan; tambahkan IF.",
               "COUNTIFS tanggal mengembalikan 0: tanggal disimpan sebagai teks."],
    variasi=["Tambahkan cuti yang dibawa dari tahun lalu (carry-over) maksimal 6 hari dengan MIN.",
             "Gunakan NETWORKDAYS.INTL dengan string akhir pekan `\"0000011\"` untuk pola shift.",
             "Buat formulir pengajuan terpisah yang menulis ke log lewat VBA (proyek 10)."],
    latihan=[("Berapa hari kerja antara 1 April dan 30 April 2024 (dengan libur Idulfitri)?", "`=NETWORKDAYS.INTL(DATE(2024;4;1);DATE(2024;4;30);1;'Hari Libur 2024'!A5:A31)`"),
             ("Hitung jumlah pengajuan cuti sakit yang disetujui.", "`=COUNTIFS('Pengajuan Cuti'!E5:E500;\"Sakit\";'Pengajuan Cuti'!J5:J500;\"Disetujui\")`"),
             ("Nama karyawan dengan sisa saldo terbanyak.", "`=INDEX('Saldo Cuti'!B5:B64;MATCH(MAX('Saldo Cuti'!I5:I64);'Saldo Cuti'!I5:I64;0))`"),
             ("Tandai pengajuan yang jatuh pada bulan Desember.", "Aturan CF `=MONTH($F5)=12`"),
             ("Berapa karyawan yang belum berhak cuti?", "`=COUNTIF('Saldo Cuti'!J5:J64;\"Belum berhak\")`")],
)

tambah(
    folder="06_rekrutmen-tracker", judul="Rekrutmen Tracker", tingkat="Menengah", file="06_rekrutmen_tracker.xlsx",
    gambar="06_funnel.png",
    tujuan="Saya mencatat 150 kandidat yang bergerak di pipeline Applied > Screening > Interview HR > Interview User > Offer > Hired, "
           "lalu menghitung funnel dengan COUNTIFS, conversion rate antar tahap, time-to-hire dengan DATEDIF, analisis sumber kandidat, "
           "chart funnel, dan panduan klik membuat PivotTable + slicer yang hasilnya dibandingkan dengan sheet ringkasan berbasis rumus.",
    konsep=["Satu kolom tanggal per tahap (desain data yang ramah rumus)", "COUNTIFS dengan kriteria `\"<>\"` (tidak kosong)",
            "IF bertingkat untuk tahap terakhir", "DATEDIF \"D\" untuk time-to-hire", "AVERAGEIFS, COUNTIFS multi-kriteria",
            "Conversion rate = tahap n / tahap n-1", "Chart batang horizontal terbalik (funnel)", "PivotTable, Slicer, Timeline, PivotChart (panduan)"],
    struktur=[("Kandidat", "150 kandidat, 7 kolom tanggal, status", "`=IF(K5<>\"\";\"Hired\";IF(J5<>\"\";\"Offer\";...))`, `=IF(N5=\"Hired\";DATEDIF(F5;K5;\"D\");\"\")`"),
              ("Funnel", "Jumlah per tahap, conversion, KPI", "`=COUNTIFS(Kandidat!$G$5:$G$154;\"<>\")`, `=IFERROR(B6/B5;0)`"),
              ("Ringkasan Rumus", "Per sumber, posisi x tahap, departemen, rekruter, bulan", "`=AVERAGEIFS(Kandidat!$O$5:$O$154;Kandidat!$E$5:$E$154;A5;Kandidat!$N$5:$N$154;\"Hired\")`"),
              ("Panduan Pivot", "Langkah klik PivotTable/slicer", "-")],
    langkah=[
        ("Desain tabel", "Header baris 4: ID Kandidat, Nama, Posisi, Departemen, Sumber, Tanggal Lamar, Tgl Screening, Tgl Interview HR, Tgl Interview User, Tgl Offer, Tgl Hired, Tgl Ditolak, Tahap Terakhir, Status, Time to Hire, Lama di Pipeline, Rekruter. Setiap tahap memiliki kolom tanggal sendiri sehingga funnel cukup menghitung sel terisi."),
        ("Tahap terakhir", "M5: `=IF(K5<>\"\";\"Hired\";IF(J5<>\"\";\"Offer\";IF(I5<>\"\";\"Interview User\";IF(H5<>\"\";\"Interview HR\";IF(G5<>\"\";\"Screening\";\"Applied\")))))`. Dicek dari tahap paling akhir ke awal."),
        ("Time to hire", "O5: `=IF(N5=\"Hired\";DATEDIF(F5;K5;\"D\");\"\")` (hari dari lamar sampai hired). P5: `=IF(N5=\"Hired\";K5-F5;IF(N5=\"Rejected\";L5-F5;TODAY()-F5))`."),
        ("Dropdown dan warna", "Status: list `Aktif,Hired,Rejected`; Posisi dan Sumber dari sheet Daftar. Conditional Formatting Equal To Hired hijau, Rejected merah, Aktif kuning. Ctrl+T -> `tblKandidat`."),
        ("Funnel", "Sheet Funnel A5:A10 nama tahap. B5: `=COUNTIFS(Kandidat!$F$5:$F$154;\"<>\")` (Applied = semua yang punya tanggal lamar); B6 kolom G, B7 kolom H, dan seterusnya. C5: `=IFERROR(B5/$B$5;0)`; D6: `=IFERROR(B6/B5;0)` (conversion dari tahap sebelumnya); E6: `=B5-B6` (drop-off)."),
        ("KPI", "B15: `=IFERROR(AVERAGE(Kandidat!$O$5:$O$154);0)` (AVERAGE mengabaikan teks kosong). B18 offer acceptance: `=IFERROR(B10/B9;0)`."),
        ("Chart funnel", "Blok A4:B10 > Insert > Bar > Clustered Bar. Klik sumbu vertikal > Format Axis > centang Categories in reverse order agar Applied di atas."),
        ("Ringkasan per sumber", "B5: `=COUNTIF(Kandidat!$E$5:$E$154;A5)`; C5: `=COUNTIFS(Kandidat!$E$5:$E$154;A5;Kandidat!$N$5:$N$154;\"Hired\")`; D5: `=IFERROR(C5/B5;0)`; G5: `=IFERROR(AVERAGEIFS(Kandidat!$O$5:$O$154;Kandidat!$E$5:$E$154;A5;Kandidat!$N$5:$N$154;\"Hired\");0)`."),
        ("Matriks posisi x tahap", "B14: `=COUNTIFS(Kandidat!$C$5:$C$154;$A14;Kandidat!$F$5:$F$154;\"<>\")` lalu ganti kolom F menjadi G, H, I, J, K untuk tahap berikutnya."),
        ("PivotTable", "Ikuti sheet Panduan Pivot: Insert > PivotTable dari tblKandidat; Rows Sumber, Columns Status, Values Count of ID; Insert Slicer Posisi; Insert Timeline Tanggal Lamar. Bandingkan angkanya dengan Ringkasan Rumus; harus sama."),
    ],
    kesalahan=["COUNTIFS `\"<>\"` menghitung sel berisi spasi: bersihkan dengan Find & Replace.",
               "DATEDIF `#NUM!`: tanggal hired sebelum tanggal lamar (salah ketik tahun).",
               "PivotTable tidak berubah setelah data ditambah: klik Refresh.",
               "Conversion > 100%: tanggal tahap lanjut terisi tetapi tahap sebelumnya kosong; isi berurutan.",
               "AVERAGE menghasilkan #DIV/0!: belum ada kandidat Hired; bungkus IFERROR."],
    variasi=["Tambahkan biaya per sumber (cost per hire) dengan SUMIFS/COUNTIFS.",
             "Hitung SLA tiap tahap (hari antar tahap) dan rata-ratanya dengan AVERAGE pada kolom bantu.",
             "Buat dashboard rekrutmen dengan slicer terhubung ke beberapa PivotTable (proyek 08)."],
    latihan=[("Jumlah kandidat dari LinkedIn yang mencapai Interview User.", "`=COUNTIFS(Kandidat!E5:E154;\"LinkedIn\";Kandidat!I5:I154;\"<>\")`"),
             ("Time-to-hire tercepat untuk posisi Programmer.", "`=MINIFS(Kandidat!O5:O154;Kandidat!C5:C154;\"Programmer\")` (Excel 2019+) atau `=MIN(IF(Kandidat!C5:C154=\"Programmer\";Kandidat!O5:O154))` sebagai rumus array."),
             ("Persentase kandidat yang ditolak di tahap Screening (punya tanggal screening tetapi tidak punya Interview HR dan status Rejected).", "`=COUNTIFS(Kandidat!G5:G154;\"<>\";Kandidat!H5:H154;\"\";Kandidat!N5:N154;\"Rejected\")/Funnel!B6`"),
             ("Rekruter dengan kandidat hired terbanyak.", "`=INDEX('Ringkasan Rumus'!A34:A36;MATCH(MAX('Ringkasan Rumus'!C34:C36);'Ringkasan Rumus'!C34:C36;0))`"),
             ("Jumlah kandidat yang melamar pada kuartal 2 2024.", "`=COUNTIFS(Kandidat!F5:F154;\">=\"&DATE(2024;4;1);Kandidat!F5:F154;\"<\"&DATE(2024;7;1))`")],
)

tambah(
    folder="07_penilaian-kinerja-kpi", judul="Penilaian Kinerja KPI", tingkat="Menengah", file="07_penilaian_kinerja_kpi.xlsx",
    gambar="07_penilaian.png",
    tujuan="Saya membangun sistem penilaian kinerja semester untuk 80 karyawan: lima KPI berbobot per level jabatan, skor 1-5, nilai akhir "
           "dengan SUMPRODUCT, rating dengan IFS, peringkat dengan RANK.EQ, kuadran 9-box performa vs potensi dengan INDEX/MATCH dua arah, "
           "Conditional Formatting skala warna, serta rekap distribusi rating per departemen dan 10 besar karyawan.",
    konsep=["SUMPRODUCT untuk rata-rata berbobot", "IFS (dan padanan IF bertingkat)", "RANK.EQ dan peringkat unik", "INDEX/MATCH dua arah (matriks)",
            "COUNTIFS untuk matriks 9-box dan distribusi", "MAXIFS untuk karyawan terbaik per departemen", "Data Validation whole number 1-5",
            "Conditional Formatting Color Scale", "Validasi bobot = 100% dengan ROUND"],
    struktur=[("Master KPI", "KPI dan bobot per level, skala skor, batas rating", "`=IF(ROUND(L5;4)=1;\"OK\";\"Bobot harus 100%\")`"),
              ("Penilaian", "80 karyawan: skor, bobot, nilai akhir, rating, peringkat, 9-box", "`=ROUND(SUMPRODUCT(F5:J5;K5:O5);2)`, `=IFS(P5>=4,5;\"Istimewa\";P5>=3,5;\"Baik\";P5>=2,5;\"Cukup\";TRUE;\"Perlu Perbaikan\")`, `=RANK.EQ(P5;$P$5:$P$84;0)`"),
              ("9-Box", "Label kuadran dan jumlah karyawan per kuadran", "`=COUNTIFS(Penilaian!$T$5:$T$84;$B13;Penilaian!$U$5:$U$84;C$12)`"),
              ("Rekap", "Distribusi rating per departemen, target distribusi, 10 besar", "`=MAXIFS(...)`, `=INDEX(...;MATCH(A22;Penilaian!$X$5:$X$84;0))`")],
    langkah=[
        ("Master KPI", "A5:A8 level (Staff, Senior Staff, Supervisor, Manager). Untuk tiap level isi KPI 1-5 (B, D, F, H, J) dan bobot (C, E, G, I, K) dalam persen. L5: `=C5+E5+G5+I5+K5`; M5: `=IF(ROUND(L5;4)=1;\"OK\";\"Bobot harus 100%\")`. ROUND menghindari selisih pembulatan biner."),
        ("Skor", "Sheet Penilaian: kolom A-E data karyawan, F-J skor 1-5. Blok F5:J84 > Data Validation > Whole number between 1 and 5. Warnai kuning sebagai sel input."),
        ("Bobot via INDEX/MATCH", "K5: `=INDEX('Master KPI'!$C$5:$C$8;MATCH($E5;'Master KPI'!$A$5:$A$8;0))`; L5 kolom E, M5 kolom G, N5 kolom I, O5 kolom K. Bobot mengikuti level sehingga Manager dan Staff dinilai berbeda."),
        ("Nilai akhir", "P5: `=ROUND(SUMPRODUCT(F5:J5;K5:O5);2)`. SUMPRODUCT mengalikan berpasangan (skor1 x bobot1 + ...) lalu menjumlahkan. Conditional Formatting > Color Scales pada P5:P84."),
        ("Rating", "Q5: `=IFS(P5>=4,5;\"Istimewa\";P5>=3,5;\"Baik\";P5>=2,5;\"Cukup\";TRUE;\"Perlu Perbaikan\")`. IFS mengevaluasi berurutan; `TRUE` di akhir sebagai default. Excel 2016 pakai `=IF(P5>=4,5;\"Istimewa\";IF(P5>=3,5;\"Baik\";IF(P5>=2,5;\"Cukup\";\"Perlu Perbaikan\")))`."),
        ("Peringkat", "R5: `=RANK.EQ(P5;$P$5:$P$84;0)` (0 = turun, nilai tertinggi peringkat 1). Kolom bantu X5: `=R5+COUNTIF($P$5:P5;P5)-1` agar peringkat unik untuk INDEX/MATCH."),
        ("Kategori dan 9-box", "T5: `=IF(P5>=4;\"Tinggi\";IF(P5>=3;\"Sedang\";\"Rendah\"))`; U5 sama untuk potensi S5. Sheet 9-Box: B6:B8 performa, C5:E5 potensi, C6:E8 label kuadran. V5: `=INDEX('9-Box'!$C$6:$E$8;MATCH(T5;'9-Box'!$B$6:$B$8;0);MATCH(U5;'9-Box'!$C$5:$E$5;0))`: MATCH pertama mencari baris, kedua mencari kolom."),
        ("Jumlah per kuadran", "9-Box C13: `=COUNTIFS(Penilaian!$T$5:$T$84;$B13;Penilaian!$U$5:$U$84;C$12)`, tarik ke E15. Color Scale pada C13:E15."),
        ("Rekap", "C5: `=COUNTIFS(Penilaian!$C$5:$C$84;$A5;Penilaian!$Q$5:$Q$84;C$4)` tarik ke F10. G5: `=AVERAGEIF(Penilaian!$C$5:$C$84;A5;Penilaian!$P$5:$P$84)`. I5: `=RANK.EQ(G5;$G$5:$G$10;0)`. J5 karyawan terbaik: `=INDEX(Penilaian!$B$5:$B$84;MATCH(MAXIFS(Penilaian!$P$5:$P$84;Penilaian!$C$5:$C$84;A5);Penilaian!$P$5:$P$84;0))`."),
        ("10 besar", "A22:A31 angka 1-10. B22: `=INDEX(Penilaian!$B$5:$B$84;MATCH(A22;Penilaian!$X$5:$X$84;0))`; salin untuk departemen, nilai, rating. Chart stacked column distribusi rating."),
    ],
    kesalahan=["SUMPRODUCT `#VALUE!`: rentang skor dan bobot tidak sama panjang.",
               "IFS `#NAME?` di Excel 2016: gunakan IF bertingkat.",
               "Rating salah karena pemisah desimal: tulis `4,5` di regional Indonesia, `4.5` di Inggris.",
               "Peringkat kembar membuat INDEX/MATCH mengambil orang yang sama: gunakan kolom peringkat unik.",
               "Bobot 0,3 diketik 30 (3000%): pastikan format persen dan nilai pecahan."],
    variasi=["Tambahkan penilaian 360 derajat (atasan, rekan, bawahan) dengan bobot penilai.",
             "Hitung bonus = gaji x faktor rating via VLOOKUP tabel faktor.",
             "Gunakan Scatter chart performa vs potensi dengan label nama (Excel 2013+: Data Labels from Cells)."],
    latihan=[("Berapa karyawan berkuadran Bintang (Star)?", "`=COUNTIF(Penilaian!V5:V84;\"Bintang (Star)\")`"),
             ("Rata-rata nilai akhir level Supervisor.", "`=AVERAGEIF(Penilaian!E5:E84;\"Supervisor\";Penilaian!P5:P84)`"),
             ("Ubah bobot KPI 1 Staff menjadi 40% dan KPI 5 menjadi 0%; cek kolom Cek lalu lihat perubahan peringkat.", "Master KPI!C5 = 0,4 dan K5 = 0; L5 tetap 100% sehingga Cek = OK."),
             ("Buat rating dengan VLOOKUP approximate memakai tabel batas di Master KPI!D12:E15.", "`=VLOOKUP(P5;'Master KPI'!$D$12:$E$15;2;TRUE)`"),
             ("Persentase karyawan berperforma Tinggi di departemen IT.", "`=COUNTIFS(Penilaian!C5:C84;\"IT\";Penilaian!T5:T84;\"Tinggi\")/COUNTIF(Penilaian!C5:C84;\"IT\")`")],
)

tambah(
    folder="08_hr-dashboard-analitik", judul="HR Dashboard Analitik", tingkat="Lanjut", file="08_hr_dashboard_analitik.xlsx",
    gambar="08_dashboard.png",
    tujuan="Saya merancang dashboard HR satu halaman dari dataset 250 karyawan dan data keluar 2023-2024: kartu KPI (headcount, turnover rate, "
           "rata-rata masa kerja, rasio gender, rata-rata usia), chart headcount per departemen, tren keluar bulanan, distribusi usia, "
           "filter departemen berbasis dropdown, panduan Power Query untuk menggabungkan dua CSV, panduan slicer/timeline, dan padanan rumus Excel 365.",
    konsep=["Pemisahan Data - Perhitungan - Dashboard", "COUNTIFS/AVERAGEIFS dengan kriteria wildcard `*` (filter Semua)",
            "Turnover rate = keluar / rata-rata headcount", "Headcount bulanan dari tanggal masuk dan keluar", "Kartu KPI dengan sel gabungan dan format besar",
            "Chart kolom, bertumpuk, pie, garis dari tabel bantu", "Power Query: import CSV, tipe data, Merge (Left Outer), Custom Column",
            "Slicer/Timeline untuk tabel dan PivotTable", "UNIQUE/FILTER/SORT/XLOOKUP/LET (Excel 365) dan padanannya di Excel 2016"],
    struktur=[("Dashboard", "KPI card, 6 chart, tabel ringkas, filter F3", "`=COUNTIFS('Data Karyawan'!$L$5:$L$254;\"Aktif\";'Data Karyawan'!$F$5:$F$254;$I$3)` dengan I3 = `=IF(F3=\"Semua\";\"*\";F3)`"),
              ("Data Karyawan", "250 baris, kolom usia/masa kerja/kelompok usia rumus", "`=DATEDIF(D5;Dashboard!$C$3;\"Y\")`, `=ROUND((IF(L5=\"Aktif\";Dashboard!$C$3;M5)-J5)/365,25;1)`"),
              ("Perhitungan", "Headcount per dept, usia, level, alasan keluar, turnover bulanan 24 bulan, ringkasan tahunan", "`=COUNTIFS(J;\"<\"&B23)-COUNTIFS(M;\"<\"&B23)`, `=IFERROR(F23/AVERAGE(D23;G23);0)`"),
              ("Panduan Power Query", "Langkah gabung karyawan.csv + turnover_2023_2024.csv", "-"), ("Panduan Slicer", "Slicer, timeline, rumus 365", "-")],
    langkah=[
        ("Dataset", "Sheet Data Karyawan: NIK, Nama, JK, Tanggal Lahir, Usia, Departemen, Jabatan, Level, Status Kepegawaian, Tanggal Masuk, Masa Kerja, Status Aktif, Tanggal Keluar, Alasan Keluar, Gaji Pokok, Kota, Kelompok Usia, Tahun Masuk, Tahun Keluar, Pendidikan. Tanggal acuan di Dashboard!C3 (31/12/2024). E5: `=DATEDIF(D5;Dashboard!$C$3;\"Y\")`. K5: `=ROUND((IF(L5=\"Aktif\";Dashboard!$C$3;M5)-J5)/365,25;1)` (masa kerja berhenti saat keluar). Q5: `=IF(E5<25;\"< 25\";IF(E5<35;\"25-34\";IF(E5<45;\"35-44\";IF(E5<55;\"45-54\";\">= 55\"))))`. S5: `=IF(M5=\"\";\"\";YEAR(M5))`. Ctrl+T -> tblKaryawan."),
        ("Filter departemen", "Dashboard!F3 dropdown `Semua,HR,Finance,Operasional,Marketing,IT,Produksi`. I3: `=IF(F3=\"Semua\";\"*\";F3)`. Tanda `*` adalah wildcard COUNTIFS yang cocok dengan teks apa pun, sehingga satu rumus melayani filter dan tanpa filter."),
        ("Kartu KPI", "Gabungkan B6:C6 (label) dan B7:C7 (angka), isi warna, font 18. B7: `=COUNTIFS('Data Karyawan'!$L$5:$L$254;\"Aktif\";'Data Karyawan'!$F$5:$F$254;$I$3)`. F7: `=AVERAGEIFS('Data Karyawan'!$K$5:$K$254;'Data Karyawan'!$L$5:$L$254;\"Aktif\";'Data Karyawan'!$F$5:$F$254;$I$3)`. Rasio gender: `=TEXT(COUNTIFS(JK;\"L\";...)/C7;\"0%\")&\" : \"&TEXT(COUNTIFS(JK;\"P\";...)/C7;\"0%\")`."),
        ("Tabel bantu per departemen", "Sheet Perhitungan A5:A10 departemen. B5: `=COUNTIFS('Data Karyawan'!$F$5:$F$254;A5;'Data Karyawan'!$L$5:$L$254;\"Aktif\")`; C/D gender; E/F/G status; H usia dan I masa kerja dengan AVERAGEIFS; K5 keluar 2024: `=COUNTIFS('Data Karyawan'!$F$5:$F$254;A5;'Data Karyawan'!$S$5:$S$254;2024)`."),
        ("Headcount bulanan", "Baris 23-46: B tanggal awal bulan, C awal bulan berikut. D23: `=COUNTIFS('Data Karyawan'!$J$5:$J$254;\"<\"&B23)-COUNTIFS('Data Karyawan'!$M$5:$M$254;\"<\"&B23)` (sudah masuk dikurangi sudah keluar). E23 masuk: `=COUNTIFS(J;\">=\"&B23;J;\"<\"&C23)`; F23 keluar sama pada kolom M. G23: `=D23+E23-F23`. H23: `=IFERROR(F23/AVERAGE(D23;G23);0)`."),
        ("Turnover tahunan", "M23: `=D23` (headcount awal 2023), M24: `=G34`, M26 keluar: `=SUM(F23:F34)`, M27: `=AVERAGE(M23;M24)`, M28: `=IFERROR(M26/M27;0)`. Kolom N untuk 2024 (baris 35-46). KPI turnover di Dashboard merujuk N26/N27."),
        ("Chart", "Dari Perhitungan: blok A4:B10 > Insert > Column (headcount). A4:A10 + C4:D10 > Stacked Column (gender). K14:L17 > Pie (status). A48:C60 > Line (keluar 2023 vs 2024). A14:B19 > Column (usia). H14:I19 > Bar (alasan keluar). Potong (Ctrl+X) tiap chart dan tempel di Dashboard, atur ukuran agar sejajar; matikan gridlines (View > Gridlines)."),
        ("Power Query", "Ikuti sheet Panduan Power Query: Data > Get Data > From Text/CSV `data/karyawan.csv` (delimiter titik koma), ulangi untuk `data/turnover_2023_2024.csv`, Merge Queries berdasarkan NIK dengan Left Outer, expand kolom TanggalKeluar dan AlasanKeluar, Custom Column StatusAktif `if [TanggalKeluar] = null then \"Aktif\" else \"Keluar\"`, Close & Load."),
        ("Slicer dan timeline", "Klik tblKaryawan > Table Design > Insert Slicer (Departemen, Status Aktif). Untuk PivotTable: Insert Timeline Tanggal Masuk. Report Connections untuk menyambungkan beberapa pivot."),
        ("Versi Excel 365", "Ganti tabel bantu dengan `=SORT(UNIQUE(tblKaryawan[Departemen]))` lalu COUNTIFS merujuk hasil tumpahan (`A5#`). Daftar karyawan IT aktif: `=FILTER(tblKaryawan;(tblKaryawan[Departemen]=\"IT\")*(tblKaryawan[Status Aktif]=\"Aktif\"))`. Untuk Excel 2016: daftar unik lewat Data > Remove Duplicates atau PivotTable; FILTER lewat Advanced Filter."),
    ],
    kesalahan=["Headcount awal negatif/salah: tanggal keluar kosong dianggap teks; COUNTIFS `\"<\"&tanggal` mengabaikan sel kosong, jadi aman, tetapi jangan isi `-`.",
               "Turnover 2024 berubah saat filter departemen: KPI ini sengaja dihitung total; gunakan kolom L di Perhitungan untuk per departemen.",
               "Chart kehilangan data setelah dipindah: sumber chart tetap di Perhitungan, jangan hapus sheet itu.",
               "Power Query gagal baca tanggal: pilih File Origin UTF-8 dan ubah tipe kolom ke Date dengan locale Indonesia.",
               "Wildcard `*` tidak bekerja untuk angka: untuk kolom numerik gunakan kriteria `\"<>\"`."],
    variasi=["Tambahkan KPI biaya gaji per departemen dan rasio span of control.",
             "Buat versi Power Pivot dengan DAX (`CALCULATE`, `DISTINCTCOUNT`).",
             "Sambungkan dashboard ke Power BI (lihat repositori analisis penjualan saya)."],
    latihan=[("Headcount aktif level Manager.", "`=COUNTIFS('Data Karyawan'!H5:H254;\"Manager\";'Data Karyawan'!L5:L254;\"Aktif\")`"),
             ("Turnover rate departemen Produksi 2024.", "`=Perhitungan!L10` atau `=K10/((B10+K10+B10)/2)`"),
             ("Rata-rata gaji karyawan aktif usia 25-34.", "`=AVERAGEIFS('Data Karyawan'!O5:O254;'Data Karyawan'!Q5:Q254;\"25-34\";'Data Karyawan'!L5:L254;\"Aktif\")`"),
             ("Bulan dengan karyawan keluar terbanyak (2023-2024).", "`=INDEX(Perhitungan!A23:A46;MATCH(MAX(Perhitungan!F23:F46);Perhitungan!F23:F46;0))`"),
             ("Dengan Excel 365, daftar nama karyawan HR aktif terurut.", "`=SORT(FILTER(tblKaryawan[Nama];(tblKaryawan[Departemen]=\"HR\")*(tblKaryawan[Status Aktif]=\"Aktif\")))`")],
)

tambah(
    folder="09_analisis-survei-engagement-statistik", judul="Analisis Survei Engagement dengan Statistika", tingkat="Lanjut", file="09_analisis_survei_engagement.xlsx",
    gambar="09_regresi.png",
    tujuan="Saya menganalisis survei engagement 200 responden (15 item pada 5 dimensi + 3 item engagement, skala 1-5) dengan alat statistika di Excel: "
           "statistik deskriptif lengkap dengan interval kepercayaan, matriks korelasi, uji t dua sampel (manual dan T.TEST), ANOVA satu arah "
           "(rumus manual dibandingkan Analysis ToolPak), regresi linear berganda dengan LINEST, histogram, dan interpretasi untuk HR. "
           "Proyek ini memperlihatkan latar belakang statistika saya.",
    konsep=["AVERAGE item -> skor dimensi", "COUNT, MEDIAN, MODE.SNGL, STDEV.S, VAR.S, QUARTILE, SKEW, KURT", "Standard error dan CI 95% dengan T.INV.2T",
            "CORREL dan uji signifikansi r", "T.TEST tipe 2 (varians sama) dan 3 (Welch), t hitung manual, Cohen's d",
            "ANOVA: SS, df, MS, F dengan SUMPRODUCT dan F.DIST.RT/F.INV.RT", "LINEST 5x6 via INDEX, t dan p koefisien, R2 disesuaikan",
            "Histogram dengan COUNTIFS (alternatif FREQUENCY)", "Analysis ToolPak (panduan klik)"],
    struktur=[("Data Survei", "200 responden, 18 item, skor dimensi, kolom bantu ANOVA, prediksi & residual", "`=AVERAGE(F5:H5)`, `=IF($B5=\"HR\";$AC5;\"\")`, `=Regresi!$G$5+Regresi!$F$5*X5+...`"),
              ("Statistik Deskriptif", "15 ukuran x 6 variabel + chart", "`=STDEV.S('Data Survei'!$X$5:$X$204)`, `=B6-T.INV.2T(0,05;B5-1)*B17`"),
              ("Korelasi", "Matriks 6x6 CORREL, uji r", "`=CORREL(x;y)`, `=T.DIST.2T(ABS(B15);198)`"),
              ("Uji t", "Tetap vs Kontrak", "`=T.TEST(tetap;kontrak;2;2)`, `=B11/SQRT(B12*(1/B5+1/C5))`"),
              ("ANOVA", "Engagement antar departemen", "`=SUMPRODUCT(B5:B10;(C5:C10-B15)^2)`, `=F.DIST.RT(E18;C18;C19)`"),
              ("Regresi", "LINEST, tabel koefisien, ringkasan model", "`=INDEX(LINEST(y;x;TRUE;TRUE);1;1)`"),
              ("Histogram", "8 kelas, frekuensi, chart", "`=COUNTIFS(y;\">=\"&B5;y;\"<\"&C5)`"), ("Interpretasi", "Bacaan untuk HR", "-")],
    langkah=[
        ("Skor dimensi", "Data Survei: F-W item 1-5 (3 item per dimensi, 3 item engagement). X5: `=AVERAGE(F5:H5)` (Kepemimpinan), Y5 `=AVERAGE(I5:K5)`, ..., AC5: `=AVERAGE(U5:W5)` (Engagement). Urutkan data berdasarkan Status (Tetap dulu) agar tiap kelompok menjadi satu rentang untuk T.TEST."),
        ("Kolom bantu ANOVA", "AD5: `=IF($B5=\"HR\";$AC5;\"\")`, AE5 Finance, ..., AI5 Produksi. Fungsi COUNT/AVERAGE/VAR.S mengabaikan teks kosong sehingga tiap kolom mewakili satu kelompok."),
        ("Statistik deskriptif", "B5: `=COUNT('Data Survei'!$X$5:$X$204)`, B6 AVERAGE, B7 MEDIAN, B8 `=MODE.SNGL(...)`, B9 `=STDEV.S(...)`, B10 `=VAR.S(...)`, B11 MIN, B12 MAX, B13 `=QUARTILE(...;1)`, B14 `=QUARTILE(...;3)`, B15 SKEW, B16 KURT, B17 `=B9/SQRT(B5)`, B18 `=B6-T.INV.2T(0,05;B5-1)*B17`, B19 `=B6+...`. Salin ke kolom C-G untuk dimensi lain. STDEV.S membagi n-1 (sampel)."),
        ("Korelasi", "B5: `=CORREL('Data Survei'!$X$5:$X$204;'Data Survei'!$X$5:$X$204)` dan seterusnya membentuk matriks 6x6; Color Scale. Uji r: t = `=B14*SQRT(200-2)/SQRT(1-B14^2)`, p = `=T.DIST.2T(ABS(B15);198)`."),
        ("Uji t", "Rentang Tetap = AC5:AC(n_tetap+4), Kontrak sisanya. B5 `=COUNT(tetap)`, B6 AVERAGE, B7 VAR.S. B12 pooled: `=((B5-1)*B7+(C5-1)*C7)/(B5+C5-2)`. B13: `=B11/SQRT(B12*(1/B5+1/C5))`. B16: `=T.DIST.2T(ABS(B13);B14)`. B17: `=T.TEST(tetap;kontrak;2;2)` (ekor = 2, tipe 2 = varians sama; tipe 3 = Welch). Keputusan: `=IF(B17<B10;\"Tolak H0\";\"Gagal tolak H0\")`."),
        ("ANOVA manual", "B5: `=COUNT('Data Survei'!$AD$5:$AD$204)`, C5 AVERAGE, D5 VAR.S, E5 SUM per departemen. B15 grand mean `=SUM(E5:E10)/B13`. SS antar: `=SUMPRODUCT(B5:B10;(C5:C10-B15)^2)`; SS dalam: `=SUMPRODUCT(B5:B10-1;D5:D10)`; df = k-1 dan N-k; MS = SS/df; F = MS antar/MS dalam; p = `=F.DIST.RT(E18;C18;C19)`; F tabel = `=F.INV.RT(0,05;C18;C19)`. Eta kuadrat = SS antar/SS total."),
        ("ANOVA ToolPak", "File > Options > Add-ins > Go > centang Analysis ToolPak. Data > Data Analysis > Anova: Single Factor; Input Range AD4:AI204 (Grouped by Columns, Labels in first row), Alpha 0,05, Output di sel kosong. Angkanya harus identik dengan tabel manual."),
        ("Regresi LINEST", "B5: `=INDEX(LINEST('Data Survei'!$AC$5:$AC$204;'Data Survei'!$X$5:$AB$204;TRUE;TRUE);1;1)` dan ubah indeks baris 1-5, kolom 1-6 untuk sel lain. Baris 1 koefisien (urutan terbalik: X5, X4, ..., X1, konstanta), baris 2 SE, baris 3 R2 dan SE y, baris 4 F dan df residual, baris 5 SS regresi dan SS residual. t = koefisien/SE, p = `=T.DIST.2T(ABS(t);df)`. R2 adj = `=1-(1-R2)*(n-1)/(n-k-1)`."),
        ("Prediksi dan residual", "Data Survei AJ5: `=Regresi!$G$5+Regresi!$F$5*X5+Regresi!$E$5*Y5+Regresi!$D$5*Z5+Regresi!$C$5*AA5+Regresi!$B$5*AB5`; AK5: `=AC5-AJ5`. Residual dapat diplot untuk memeriksa asumsi."),
        ("Histogram", "Batas kelas B5:C12 (1,0-1,5; ...; 4,5-5,0). D5: `=COUNTIFS('Data Survei'!$AC$5:$AC$204;\">=\"&B5;'Data Survei'!$AC$5:$AC$204;\"<\"&C5)`. E5: `=D5/SUM($D$5:$D$12)`. Chart Column, Gap Width 10%. Alternatif: `=FREQUENCY(data;batas_atas)` sebagai rumus array."),
    ],
    kesalahan=["T.TEST `#N/A`: rentang mengandung teks atau kosong; pastikan kelompok adalah rentang angka murni (data terurut).",
               "LINEST hanya menampilkan satu angka: di Excel lama blok 5x6 lalu Ctrl+Shift+Enter, atau pakai INDEX seperti di sini.",
               "Koefisien tertukar: LINEST mengembalikan koefisien dari variabel terakhir ke pertama.",
               "VAR.S vs VAR.P: gunakan .S untuk sampel survei.",
               "ANOVA manual beda dengan ToolPak: ada sel kosong yang terisi spasi sehingga COUNT berbeda."],
    variasi=["Uji asumsi: normalitas residual (histogram, Q-Q), homogenitas varians (uji Levene manual).",
             "Regresi logistik untuk memprediksi niat resign (butuh Solver atau R).",
             "Analisis reliabilitas Cronbach's alpha per dimensi: `=k/(k-1)*(1-SUM(var_item)/var_total)`."],
    latihan=[("Hitung Cronbach's alpha dimensi Kepemimpinan (item F, G, H).", "`=3/2*(1-(VAR.S(F5:F204)+VAR.S(G5:G204)+VAR.S(H5:H204))/VAR.S(X5:X204*3))` dengan total = F+G+H per responden pada kolom bantu."),
             ("Uji t engagement laki-laki vs perempuan (data tidak terurut).", "Buat kolom bantu `=IF(C5=\"L\";AC5;\"\")` dan `=IF(C5=\"P\";AC5;\"\")`, lalu `=T.TEST(kolom_L;kolom_P;2;3)` tidak menerima teks: salin nilai angka saja ke dua kolom (Paste Values, hapus kosong) sebelum T.TEST."),
             ("Korelasi Spearman sederhana antara Kompensasi dan Engagement.", "Buat kolom peringkat `=RANK.AVG(Y5;$Y$5:$Y$204)` untuk kedua variabel lalu `=CORREL(rank1;rank2)`."),
             ("Berapa responden dengan engagement < 2,5?", "`=COUNTIF('Data Survei'!AC5:AC204;\"<2,5\")`"),
             ("Prediksi engagement bila semua dimensi bernilai 4.", "`=Regresi!G5+4*SUM(Regresi!B5:F5)`")],
)

tambah(
    folder="10_perencanaan-tenaga-kerja-dan-otomasi", judul="Perencanaan Tenaga Kerja dan Otomasi", tingkat="Lanjut", file="10_perencanaan_tenaga_kerja.xlsx",
    gambar="10_anggaran_gaji.png",
    tujuan="Saya memproyeksikan headcount 2025-2027 per departemen dari data 2019-2024 dengan FORECAST.LINEAR, TREND, dan CAGR, menghitung kebutuhan "
           "rekrutmen, menyimulasikan anggaran gaji dengan skenario kenaikan persen, lalu melatih What-If Analysis (Data Table satu dan dua "
           "variabel, Goal Seek, Scenario Manager). Modul VBA di folder `vba/` mengotomatiskan slip gaji PDF massal, rekap per departemen, dan validasi data.",
    konsep=["FORECAST.LINEAR dan TREND (regresi linear sederhana)", "CAGR dengan pangkat pecahan", "MAX untuk kebutuhan rekrutmen non-negatif",
            "Model anggaran dengan parameter terpusat dan skenario INDEX/MATCH", "Data Table 1 dan 2 variabel (What-If)", "Goal Seek dan jawaban analitik",
            "Scenario Manager dan Summary", "SUMPRODUCT + COUNTIF untuk duplikat, COUNTBLANK", "VBA: Dictionary, loop, ExportAsFixedFormat, Worksheets.Add"],
    struktur=[("Historis", "Headcount 2019-2024 per departemen, CAGR, tahun numerik", "`=(G5/B5)^(1/5)-1`"),
              ("Proyeksi", "FORECAST.LINEAR, TREND, CAGR, kebutuhan rekrutmen", "`=ROUND(FORECAST.LINEAR(2025;Historis!$B$5:$G$5;Historis!$B$13:$G$13);0)`, `=ROUND(TREND(y;x;2026);0)`"),
              ("Anggaran Gaji", "Skenario, model 2025-2027, simulasi setara Data Table, Goal Seek", "`=INDEX(C14:E14;MATCH(C6;C13:E13;0))`, `=E19*(1+$C$4)*(1+$C$5)`, `=C8/E25/(1+C5)-1`"),
              ("Data Karyawan", "120 baris bahan uji makro (3 masalah disengaja)", "`=SUMPRODUCT((COUNTIF(A5:A124;A5:A124)>1)*1)`"),
              ("Panduan What-If & VBA", "Langkah klik", "-")],
    langkah=[
        ("Data historis", "Historis!A5:A10 departemen, B4:G4 tahun sebagai teks header, B13:G13 tahun sebagai angka (2019...2024) untuk sumbu x. H5 CAGR: `=(G5/B5)^(1/5)-1`."),
        ("FORECAST.LINEAR", "Proyeksi!C5: `=ROUND(FORECAST.LINEAR(2025;Historis!$B$5:$G$5;Historis!$B$13:$G$13);0)`. Argumen: x baru, known_y (headcount), known_x (tahun). Excel 2013 ke bawah: gunakan `FORECAST`. D5 untuk 2026, E5 untuk 2027."),
        ("TREND", "F5: `=ROUND(TREND(Historis!$B$5:$G$5;Historis!$B$13:$G$13;2025);0)`; hasil sama dengan FORECAST.LINEAR tetapi TREND dapat mengembalikan beberapa nilai sekaligus (array)."),
        ("CAGR dan kebutuhan rekrutmen", "I5: `=ROUND(B5*(1+Historis!H5)^3;0)`. Kebutuhan 2025 (baris 16-21): pengganti turnover `=ROUND(B16*$B$13;0)` (B13 = asumsi turnover 12%), penambahan baru `=MAX(C5-B5;0)`, total = keduanya."),
        ("Model anggaran", "Anggaran Gaji!C13:E15 tabel skenario (nama, kenaikan gaji, pertumbuhan headcount). C6 dropdown skenario; C4: `=INDEX(C14:E14;MATCH(C6;C13:E13;0))`, C5 serupa. E19: `=C19*D19*$C$7` (headcount x gaji rata-rata x 13 bulan). F19: `=E19*(1+$C$4)*(1+$C$5)`, G19 dari F19, H19 dari G19. Total baris 25."),
        ("Data Table satu variabel", "B32:B39 kenaikan 3%-10%. Versi rumus eksplisit: C32 `=$E$25*(1+B32)*(1+$C$5)`. Versi Excel asli: di C31 ketik `=F25`, blok B31:C39, Data > What-If Analysis > Data Table, Column input cell C4, OK; Excel mengisi `{=TABLE(;C4)}` yang menghitung ulang model untuk tiap nilai."),
        ("Data Table dua variabel", "F31 sel sudut berisi `=F25`; G31:K31 pertumbuhan headcount; F32:F36 kenaikan gaji. Blok F31:K36 > Data Table > Row input cell C5, Column input cell C4. Versi rumus: `=$E$25*(1+$F32)*(1+G$31)`."),
        ("Goal Seek", "Data > What-If Analysis > Goal Seek: Set cell F25, To value = nilai C8 (batas anggaran), By changing cell C4. Excel mencari kenaikan gaji yang memenuhi batas. Jawaban analitiknya: C28 `=C8/E25/(1+C5)-1`."),
        ("Scenario Manager", "Data > What-If Analysis > Scenario Manager > Add: nama Pesimis, Changing cells C4:C5, nilai 3% dan 0%; ulangi Moderat (6%, 4%) dan Optimis (9%, 8%). Show untuk menerapkan; Summary dengan Result cells F25:H25 menghasilkan laporan perbandingan."),
        ("Pasang makro VBA", "Alt+F11 > File > Import File > pilih `vba/ValidasiData.bas`, `vba/RekapPerDepartemen.bas` (untuk file ini) dan `vba/SlipGajiPDF.bas` (untuk file proyek 04). Save As tipe Excel Macro-Enabled Workbook (.xlsm). Developer > Macros > Run. ValidasiData menemukan 3 masalah yang sengaja ditanam (NIK ganda baris 40, email kosong baris 61, gaji di bawah UMK baris 85); RekapPerDepartemen membuat 6 sheet; SlipGajiPDF mengekspor 70 PDF ke folder slip_pdf."),
    ],
    kesalahan=["FORECAST.LINEAR `#NAME?` di Excel 2013: gunakan FORECAST.",
               "Proyeksi negatif untuk departemen menyusut: batasi dengan MAX(...;0).",
               "Data Table tidak berubah: Formulas > Calculation Options harus Automatic (bukan Automatic except for data tables).",
               "Goal Seek tidak konvergen: sel pengubah mengandung rumus (C4 dari INDEX); ketik angka dulu atau ubah ke sel input.",
               "Makro tidak muncul: file masih .xlsx atau makro dinonaktifkan di Trust Center."],
    variasi=["Tambahkan proyeksi musiman dengan FORECAST.ETS untuk data bulanan.",
             "Gunakan Solver untuk mengalokasikan tambahan headcount dengan batas anggaran.",
             "Kembangkan makro menjadi UserForm input karyawan baru."],
    latihan=[("Proyeksi headcount IT tahun 2028 dengan FORECAST.LINEAR.", "`=FORECAST.LINEAR(2028;Historis!B9:G9;Historis!B13:G13)`"),
             ("Berapa total anggaran 2026 pada skenario Optimis?", "Ubah C6 ke Optimis; baca G25 (atau `=E25*((1+9%)*(1+8%))^2`)."),
             ("Kenaikan gaji maksimal jika batas anggaran Rp52 miliar dan pertumbuhan headcount 4%.", "`=52000000000/E25/(1+4%)-1`"),
             ("Hitung kebutuhan rekrutmen total 2025 jika turnover 15%.", "Ubah Proyeksi!B13 = 15%; baca E22."),
             ("Modifikasi makro ValidasiData agar juga menandai tanggal masuk di masa depan.", "Tambahkan di loop: `If ws.Cells(r, \"F\").Value > Date Then Call Catat(wsHasil, n, r, ws, \"Tanggal masuk di masa depan\", ws.Cells(r, \"F\").Value)`")],
)


def tulis(p):
    folder = ROOT / p["folder"]
    t = [f"# Tutorial {p['folder'][:2]}: {p['judul']} (Tingkat {p['tingkat']})", "",
         f"File: `{p['file']}` | Latihan: [latihan.md](latihan.md) | Kembali ke [README utama](../README.md)", "",
         "## 1. Tujuan dan hasil akhir", "", p["tujuan"], "", f"![Hasil akhir](assets/{p['gambar']})", "",
         "## 2. Konsep Excel yang dipelajari", ""]
    t += [f"- {k}" for k in p["konsep"]]
    t += ["", "## 3. Struktur workbook", "", "| Sheet | Isi | Rumus kunci |", "|---|---|---|"]
    t += [f"| {s} | {i} | {r} |" for s, i, r in p["struktur"]]
    t += ["", "## 4. Langkah demi langkah dari nol", "", CATATAN_REGIONAL, ""]
    for n, (jud, isi) in enumerate(p["langkah"], start=1):
        t += [f"### Langkah {n}: {jud}", "", isi, ""]
    t += ["## 5. Kesalahan umum dan cara memperbaikinya", ""]
    t += [f"- {k}" for k in p["kesalahan"]]
    t += ["", "## 6. Latihan tambahan", "", "Kerjakan soal di [latihan.md](latihan.md); kunci jawaban rumus tersedia di bagian bawah file tersebut.", "",
          "## 7. Variasi untuk tingkat selanjutnya", ""]
    t += [f"- {v}" for v in p["variasi"]]
    (folder / "TUTORIAL.md").write_text("\n".join(t) + "\n", encoding="utf-8")

    l = [f"# Latihan {p['folder'][:2]}: {p['judul']}", "", f"Buka `{p['file']}` dan kerjakan di sheet baru bernama `Latihan`. "
         "Rumus memakai pemisah titik koma (;).", "", "## Soal", ""]
    l += [f"{i}. {s}" for i, (s, _) in enumerate(p["latihan"], start=1)]
    l += ["", "## Kunci jawaban", ""]
    l += [f"{i}. {j}" for i, (_, j) in enumerate(p["latihan"], start=1)]
    (folder / "latihan.md").write_text("\n".join(l) + "\n", encoding="utf-8")


if __name__ == "__main__":
    for p in PROYEK:
        tulis(p)
    print(f"{len(PROYEK)} tutorial dan latihan ditulis.")
