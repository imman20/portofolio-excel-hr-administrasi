# Modul 15: Praktik kerja profesional

> Saya memakai tagihan dan pembayaran sintetis sebagai latihan serah terima, bukan pengalaman kerja. Rumus dokumen memakai titik koma pada Excel regional Indonesia; regional Inggris memakai koma. Nama vendor adalah label simulasi dan tidak menunjuk perusahaan nyata.

## 1. Tujuan dan prasyarat

Saya belajar menyelesaikan workbook sebagai berkas yang dapat dibaca, diperiksa, dan diteruskan kepada orang lain. Cakupannya rekonsiliasi, audit formula, definisi KPI, pemisahan input/proses/output, proteksi sel, keamanan data, kompatibilitas versi, dan checklist serah terima. Tujuan saya adalah membuat asumsi serta pengecualian tetap terlihat ketika pembuat file tidak berada di dekat penerima.

Prasyaratnya adalah modul dasar sampai visualisasi dan kebiasaan memeriksa hasil. Saya memakai [latihan](latihan/latihan_15.xlsx), [solusi](latihan/solusi_15.xlsx), dan [template terproteksi](latihan/template_15.xlsx). Saya tidak menganggap file selesai hanya karena tampilannya rapi atau semua rumus tidak menampilkan pesan error.

## 2. Kegunaan di pekerjaan

Workbook yang diteruskan tanpa definisi dan pemeriksaan mudah disalahgunakan. Total pembayaran dapat berbeda dari pembayaran yang benar-benar dapat dialokasikan, sementara saldo negatif dapat tersembunyi jika digabungkan dengan invoice belum lunas. Saya memisahkan ukuran dan menjaga jejak sumber agar keputusan tidak bergantung pada tebakan penerima.

Kasus rekonsiliasi ini memperlihatkan bahwa PASS matematis berbeda dari persetujuan bisnis. Pembayaran tanpa kunci dan lebih bayar tetap memerlukan tinjauan walau seluruh identitas penjumlahan cocok. Saya menjelaskan batas tersebut sebagai bagian serah terima, bukan menghapus pengecualian agar laporan tampak bersih.

## 3. Penjelasan dari dasar

Input adalah nilai sumber dan parameter yang boleh diperbarui. Proses adalah transformasi serta formula, sedangkan output adalah ukuran yang dipakai membaca hasil. Pada file ini Data dan Pembayaran adalah sumber, Master menyimpan acuan, Calculation serta Check_Pembayaran menjadi proses, dan Report menyajikan hasil. Saya menjaga alur tersebut agar perubahan dapat ditelusuri.

Grain tagihan adalah satu invoice, sedangkan grain pembayaran adalah satu kejadian pembayaran. Satu invoice boleh dibayar beberapa kali, tetapi payment_id harus tetap unik. Menganggap setiap pengulangan invoice_id pada pembayaran sebagai duplikat akan menghapus transaksi sah. Saya memeriksa kunci sesuai tabelnya.

Rekonsiliasi membandingkan besaran yang didefinisikan sama. Jumlah pembayaran yang dikenal harus sama dengan alokasi ke invoice. Pembayaran tanpa invoice dikenal dilaporkan terpisah. Saya tidak mencampur keduanya lalu menyebut total sebagai pembayaran yang sudah dialokasikan.

Saldo invoice adalah nilai tagihan dikurangi pembayaran dikenal. Saldo positif menunjukkan sisa pada invoice, nol menunjukkan lunas, dan negatif menunjukkan lebih bayar. Sisa positif tidak dikurangi lebih bayar invoice lain secara diam-diam karena keduanya memerlukan tindakan berbeda. Saldo neto tetap tersedia untuk identitas matematis.

Audit formula berarti menelusuri sumber, cakupan, dan pola, bukan hanya mencari teks error. Formula dapat menghitung tanpa error tetapi memakai rentang yang salah. Saya memeriksa baris awal, tengah, dan akhir, membaca parameter absolut, serta menelusuri satu angka laporan ke detail dan sumber.

Dokumentasi KPI menyebut nama, definisi, satuan, periode, aturan pembagi nol, serta pengecualian. Definisi lewat tempo pada latihan memakai jatuh tempo sebelum acuan dan saldo positif. Invoice tepat pada tanggal acuan belum masuk kelompok tersebut. Saya tidak mengubah definisi tanpa memperbarui label serta pembanding.

Proteksi sheet membantu mencegah perubahan formula yang tidak disengaja. Sel input dibuka untuk edit, formula tetap terkunci, dan formula tidak disembunyikan agar audit mungkin dilakukan. Proteksi ini bukan enkripsi; orang yang memperoleh file tetap dapat membaca isinya sesuai kemampuan aplikasi dan akses yang diberikan.

Keamanan data memerlukan pengelolaan sumber, akses, berbagi, retensi, serta cadangan. Saya tidak menaruh kredensial atau data nyata pada template latihan. Password contoh sengaja diketahui pembaca sehingga tidak boleh dipakai sebagai klaim kerahasiaan. Kebutuhan nyata memerlukan pengaturan penyimpanan dan keamanan yang sesuai konteks.

Serah terima harus menyebut apa yang sudah diuji dan apa yang belum. Recalc LibreOffice tidak membuktikan semua interaksi Excel atau Google Sheets sama. Fitur manual, proteksi, grafik, serta regional perlu ditinjau pada aplikasi tujuan. Saya membedakan keterbatasan tersebut dari kesalahan hitungan yang harus diperbaiki.

## 4. Dataset latihan

[CSV tagihan](latihan/data_15.csv) berisi 60 invoice sintetis dengan tanggal, vendor, qty, harga, dan jatuh tempo. [CSV pembayaran](latihan/pembayaran_15.csv) memuat kejadian pembayaran, termasuk pembayaran terpecah, lebih bayar, dan satu kunci invoice yang tidak dikenal. Masalah ini sengaja dipertahankan sebagai kasus audit.

Calculation menghitung nilai, alokasi, saldo, status, dan lewat tempo per invoice. Check_Pembayaran memisahkan nominal dikenal dengan tanpa kunci. Checks menguji identitas nilai serta keunikan kunci; Report menampilkan ukuran dan pengecualian tanpa menyatakan semua kasus siap diposting.

[Definisi KPI](docs/definisi_kpi.md), [keamanan dan proteksi](docs/keamanan_dan_proteksi.md), serta [checklist serah terima](docs/checklist_serah_terima.md) melengkapi file. [Kunci jawaban](latihan/kunci_jawaban.md) berasal dari pandas. [VERIFIKASI.md](VERIFIKASI.md) mencatat recalc ketiga workbook, pembanding saldo, serta pemeriksaan metadata proteksi.

## 5. Langkah praktik bernomor dengan instruksi klik

1. Buka latihan, baca Instruksi, lalu periksa pemisahan sheet. Saya memilih **Formulas > Calculation Options > Automatic** dan menekan **Ctrl+Alt+F9** sebelum membaca hasil.
2. Pada Data, cocokkan arti invoice_id dan nilai sumber. Pada Pembayaran, periksa payment_id sebagai kunci kejadian serta kemungkinan beberapa pembayaran untuk invoice sama.
3. Buka Calculation dan pilih satu saldo. Gunakan **Formulas > Trace Precedents** untuk menelusuri nilai tagihan dan SUMIFS pembayaran. Saya membandingkan dengan baris sumber, bukan hanya formula yang terlihat.
4. Pilih **Formulas > Show Formulas** untuk memeriksa pola referensi. Periksa beberapa posisi, lalu kembalikan tampilan hasil. Saya tidak mengganti formula yang sulit dipahami dengan angka agar audit tampak mudah.
5. Pada Check_Pembayaran, gunakan **Data > Filter** untuk menampilkan PERIKSA KUNCI. Catat nominal tanpa alokasi dan alasan belum dapat ditetapkan ke invoice tertentu.
6. Pada Calculation, filter LEBIH BAYAR dan SEBAGIAN. Bedakan saldo negatif dengan sisa positif; saya tidak saling menghapuskan keduanya tanpa aturan bisnis.
7. Buka Checks dan pastikan identitas matematis PASS. Kemudian baca Report dan jelaskan mengapa pengecualian bisnis tetap ada meski semua selisih pemeriksaan nol.
8. Buka template pada salinan. Coba memilih input Data dan parameter Master; formula Calculation harus tetap terkunci. Interaksi ini perlu diuji pada aplikasi tujuan dan tidak diasumsikan identik dari metadata saja.
9. Untuk meninjau pengaturan, pilih **Review > Unprotect Sheet**, masukkan password contoh `latihan`, lalu **Ctrl+1 > Protection** pada sel input dan formula. Setelah perubahan yang disetujui, gunakan **Review > Protect Sheet** kembali.
10. Isi jawaban kuning, cocokkan dengan solusi, lalu ikuti checklist serah terima. Saya mencatat nama file, versi, periode, batas, cara membangun ulang, serta pengecualian yang perlu ditinjau penerima.
11. Simpan sumber, salinan kerja, dan keluaran dengan nama berbeda. Setelah setiap perubahan input, saya menghitung ulang serta mengulang rekonsiliasi, bukan mengandalkan status PASS dari keadaan sebelumnya.

## 6. Contoh formula dan arti tiap argumen

`=Data!D2*Data!E2` menghitung nilai invoice dari qty dan harga. `=SUMIFS(Pembayaran!$C$2:$C$58;Pembayaran!$B$2:$B$58;A2)` mengalokasikan seluruh pembayaran pada invoice sama, termasuk pembayaran terpecah yang sah.

`=B2-C2` pada Calculation menghasilkan saldo. Status ditentukan dari tanda saldo serta keberadaan pembayaran. Saya menjaga saldo negatif sebagai angka yang dapat diperiksa, bukan membatasinya dengan MAX menjadi nol sehingga lebih bayar menghilang.

`=SUMIFS(Calculation!D2:D61;Calculation!D2:D61;">0")` menjumlahkan sisa positif. Negatif dari SUMIFS dengan syarat kurang dari nol menunjukkan lebih bayar terpisah. Definisi kedua ukuran tersebut berbeda dari SUM seluruh saldo neto.

`=IF(D2=0;"PASS";"FAIL")` pada Checks membaca selisih rekonsiliasi, bukan seluruh kelayakan bisnis. Saya tidak memakai satu status tersebut untuk menggantikan daftar pengecualian atau keputusan pihak yang berwenang.

## 7. Latihan mandiri minimal lima soal bertingkat

1. Dasar: petakan sheet menjadi input, proses, output, dan pemeriksaan.
2. Dasar: hitung nilai tagihan serta pembayaran dikenal dan jelaskan grain kedua sumber.
3. Menengah: telusuri satu pembayaran terpecah tanpa menyebutnya duplikat hanya karena invoice_id berulang.
4. Menengah: pisahkan sisa positif, lebih bayar, dan pembayaran tanpa kunci dari saldo neto.
5. Menengah: audit satu formula dan buktikan identitas saldo dengan pembanding.
6. Lanjut: jelaskan proteksi sel, enkripsi, serta pembatasan akses sebagai kebutuhan berbeda.
7. Lanjut: susun serah terima yang menyatakan hasil uji, pengecualian, batas kompatibilitas, dan langkah setelah perubahan.

## 8. Kunci jawaban

[Kunci lengkap](latihan/kunci_jawaban.md) memuat ukuran dari CSV dan penjelasan pengecualian. Identitas saldo harus cocok, tetapi invoice tanpa kunci dikenal tetap perlu ditinjau. Pembayaran terpecah dijumlahkan menurut invoice, sementara keunikan diperiksa pada payment_id.

Template membuka input yang dinyatakan dan mengunci formula. Password contoh bukan rahasia. Jawaban serah terima harus menyebut konteks data sintetis, aplikasi yang diuji, serta hal yang tidak dapat disimpulkan dari status PASS matematis.

## 9. Kesalahan umum dan cara memeriksa hasil

Menjumlahkan seluruh pembayaran sebagai alokasi dapat memasukkan invoice asing. Menghapus semua pengulangan invoice_id dapat menghilangkan pembayaran sah. Saya memeriksa relasi dan kunci sebelum mengoreksi hasil.

Mengunci seluruh sheet tanpa membuka input membuat template sulit dipakai, sedangkan membuka formula mengurangi perlindungan dari perubahan tak sengaja. Saya memeriksa metadata dan menguji aplikasi tujuan. Proteksi juga tidak mencegah salah nilai input, sehingga rekonsiliasi tetap wajib setelah edit.

## 10. Checklist kompetensi

- [ ] Saya dapat menyatakan grain, kunci, dan definisi KPI.
- [ ] Saya dapat memisahkan input, proses, serta output.
- [ ] Saya dapat merekonsiliasi nilai sambil menjaga pengecualian terlihat.
- [ ] Saya dapat mengaudit formula dan menjelaskan batas proteksi.
- [ ] Saya dapat menyiapkan serah terima serta pengujian setelah perubahan.

## 11. Catatan kompatibilitas Excel 2016/2019/365, LibreOffice, dan Google Sheets

Formula inti tersedia pada Excel 2016/2019/365 tanpa dynamic array. Template memuat proteksi sheet dan input terbuka yang diuji secara programatis; interaksi proteksi di Excel atau Google Sheets belum diuji langsung. Impor dapat mengubah perilaku atau pengaturan proteksi.

LibreOffice dipakai untuk recalc dan [preview nyata](assets/preview.png). Tidak ada enkripsi file yang diklaim. Pembangunan ulang memakai `python tools/build.py`, lalu `python tools/course_runtime.py 15_praktik-kerja-profesional` dari folder belajar. Saya menyimpan hasil penerima pada salinan agar proses pembangunan ulang tidak menimpa pekerjaan manual.
