# Standar Data HR dan Administrasi

Konvensi ini saya pakai di semua workbook agar rumus, filter, dan PivotTable konsisten antar proyek.

## 1. Penamaan sheet dan tabel

| Jenis | Konvensi | Contoh |
|---|---|---|
| Sheet data mentah | Kata benda, huruf kapital awal | `Data Karyawan`, `Pengajuan Cuti` |
| Sheet parameter/input | `Parameter` atau `Daftar` (sumber dropdown) | `Parameter`, `Daftar` |
| Sheet hasil | `Ringkasan`, `Rekap ...`, `Dashboard` | `Rekap Departemen` |
| Nama tabel (Format as Table) | Awalan `tbl` + nama tunggal | `tblKaryawan`, `tblPengajuan` |
| Sheet bantu | Diberi label "(bantu)" pada header kolomnya | `Peringkat (bantu)` |

Baris 1-2 berisi judul dan keterangan, baris 4 header tabel, data mulai baris 5 (kecuali proyek 01 yang header di baris 1 agar lebih sederhana).

## 2. Kolom identitas

| Kolom | Format | Keterangan |
|---|---|---|
| NIK (Nomor Induk Karyawan) | 8 digit teks: `TTTTNNNN` | 4 digit tahun masuk + 4 digit nomor urut, contoh `20190042`. Disimpan sebagai teks (format `@`) agar nol di depan tidak hilang dan tidak dijumlahkan. |
| ID Kandidat | `KD-TTTT-NNN` | `KD-2024-017` |
| ID Responden survei | `RNNN` | `R001` |
| Nomor agenda surat masuk | `SM-TTTT-NNN` | Dibuat rumus `="SM-"&TEXT(B5;"yyyy")&"-"&TEXT(ROW()-4;"000")` |
| Nomor surat keluar | `NNN/HRD/BULAN_ROMAWI/TTTT` | `012/HRD/III/2024` |
| Kode barang ATK | `ATK-NNN` | `ATK-001` |

## 3. Nilai kategori baku

| Kolom | Nilai yang diizinkan |
|---|---|
| Departemen | `HR`, `Finance`, `Operasional`, `Marketing`, `IT`, `Produksi` |
| Level jabatan | `Staff`, `Senior Staff`, `Supervisor`, `Manager` |
| Status kepegawaian | `Tetap`, `Kontrak`, `Magang` |
| Status aktif | `Aktif`, `Keluar` |
| Jenis kelamin | `L`, `P` |
| Kode absensi | `H` hadir, `S` sakit, `I` izin, `A` alpa, `C` cuti |
| Status disposisi surat | `Belum`, `Proses`, `Selesai` |
| Status cuti | `Diajukan`, `Disetujui`, `Ditolak` |
| Tahap rekrutmen | `Applied`, `Screening`, `Interview HR`, `Interview User`, `Offer`, `Hired` (+ status `Rejected`/`Aktif`) |
| Status PTKP | `TK/0` ... `TK/3`, `K/0` ... `K/3` |

Semua kolom kategori memakai Data Validation tipe List yang merujuk ke sheet `Daftar`/`Parameter`, bukan daftar yang diketik langsung, supaya cukup diubah di satu tempat.

## 4. Format angka dan tanggal

| Jenis | Format Excel | Tampilan |
|---|---|---|
| Rupiah | `"Rp"#,##0` | Rp4.725.479 (regional Indonesia) |
| Persentase | `0.0%` (disimpan sebagai pecahan, 0,15 = 15,0%) | 15,0% |
| Tanggal | `dd/mm/yyyy`, disimpan sebagai tanggal asli bukan teks | 31/12/2024 |
| Skor Likert / KPI | bilangan bulat 1-5, validasi *whole number* | 4 |
| Jam lembur | angka desimal | 8 |

Tanggal acuan perhitungan usia dan masa kerja diletakkan di satu sel input (misalnya `Dashboard!C3`) supaya hasil dapat direproduksi; `TODAY()` hanya dipakai di proyek pemula.

## 5. Warna sel

| Warna | Arti |
|---|---|
| Isi biru tua, teks putih | Header |
| Kuning muda (`FFF2CC`) | Sel input yang boleh diubah pengguna |
| Hijau muda | Kartu KPI / hasil utama |
| Merah muda (Conditional Formatting) | Peringatan: di bawah target, melebihi saldo, stok minimum, duplikat |

## 6. Aturan rumus

1. Rujukan ke sheet lain memakai rentang absolut (`'Data Karyawan'!$F$5:$F$254`), bukan seluruh kolom, agar cepat dan mudah dibaca.
2. Parameter tidak pernah diketik di dalam rumus (`=E5*Parameter!$B$11`, bukan `=E5*1%`).
3. Pembagian selalu dibungkus `IFERROR(...;0)` bila penyebut bisa nol.
4. Fungsi Excel 365 (`XLOOKUP`, `FILTER`, `UNIQUE`, `SORT`) hanya ditulis di tutorial sebagai alternatif; sel hidup memakai `INDEX`/`MATCH`, `COUNTIFS`, dan `SUMIFS` agar kompatibel Excel 2016/2019 dan LibreOffice.
5. Pemisah argumen: tutorial memakai titik koma (`;`) sesuai regional Indonesia; Excel regional Inggris memakai koma (`,`).
