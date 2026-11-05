# Pemeriksaan kualitas

Saya memisahkan empat bukti: audit sumber mentah, pembersihan independen Python, pemeriksaan hidup workbook, dan pembacaan hasil recalc. Kesamaan agregat saja tidak membuktikan setiap sel benar.

## Input dan referensi

analysis_support.py membaca semua input sebagai teks. Saya menghitung ulang masalah pada 414 ID unik serta seluruh 420 baris, memeriksa enam salinan identik dan daftar ID, lalu mencocokkannya dengan JSON. Saya menolak duplikat yang bertentangan. Saya membaca ISO dan DMY dengan format eksplisit, membersihkan skor per rentang, dan membandingkan seluruh kolom hasil dengan CSV referensi yang diperbaiki.

Saya menguji bahwa CSV mentah hasil generator mempunyai hash blob sama dengan bahan. Koreksi tanggal referensi mempunyai daftar ID yang dapat ditinjau; tidak ada koreksi nilai mentah. JSON membedakan hitungan penyisipan dan hitungan akhir agar dua angka yang benar tidak tertukar cakupannya.

## Workbook

Raw_Data mempertahankan seluruh 9.660 nilai input, termasuk spasi. Reference adalah input pembanding snapshot dari CSV, dengan 414 × 27 nilai; Check_Detail mencocokkan ID dan semua kolom, termasuk missing serta tanggal. Checks merangkum jumlah perbedaan dan statistik independen. Status latihan tetap BELUM DIISI hingga formula selesai.

Verifier memeriksa semua sheet wajib, Freeze Panes, ukuran di bawah 5 MB, ketiadaan tautan workbook eksternal, serta daftar fungsi yang diizinkan sebelum dan sesudah recalc. Rumus array dan sel error menyebabkan kegagalan. Selisih pembersihan numerik dibulatkan 12 desimal; selisih statistik dibulatkan 10 desimal. Ini mengakomodasi representasi biner tanpa toleransi persentase. Nilai pembanding penuh tetap ditulis.

Mean, SD, CI, alpha, jumlah/proporsi NPS, t, df, p, d, korelasi, dan regresi dibandingkan dengan perhitungan dari CSV; perhitungan Python tidak memakai nilai workbook sebagai sumber. Verifier membaca data_only=True sesudah recalc. Saya tidak menyebut file baru terverifikasi jika LibreOffice gagal.

## Reproduksi

```bash
python -m pip install -r requirements.txt
python tools/generate_data.py
python tools/build_workbook.py
python tools/verify.py
```

LibreOffice headless (`soffice`), pdfinfo, dan pdftoppm diperlukan pada PATH. Verifier menulis VERIFIKASI.md, rekonsiliasi CSV, ringkasan versi dan hash, serta render PDF/PNG Report. Saya menjalankan ulang verifikasi setelah mengubah builder atau data, memeriksa preview secara visual, serta memeriksa tautan Markdown yang diubah. Proyek Sales, HR, dan modul dasar tidak berubah.

Saya juga menguji verifier pada salinan yang sengaja dirusak: tanggal referensi tertukar, duplikat berkonflik, hitungan JSON salah, dan skor formula berubah. Kegagalan harus terdeteksi, bukan ditutup dengan perubahan pembanding. Rincian hasil pengujian negatif ada pada [VERIFIKASI.md](../VERIFIKASI.md).
