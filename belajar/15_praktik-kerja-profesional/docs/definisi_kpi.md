# Definisi KPI rekonsiliasi

Saya memakai tanggal acuan tetap pada Master dan input sintetis. Satu baris tagihan adalah satu invoice; satu invoice dapat dibayar beberapa kali. payment_id harus unik, sedangkan invoice_id pada pembayaran boleh berulang.

| KPI | Definisi | Periode dan batas |
|---|---|---|
| Nilai tagihan | qty × harga satuan | Seluruh invoice pada sumber latihan |
| Pembayaran dikenal | SUMIFS pembayaran menurut invoice_id yang ada pada Data | Sampel pembayaran sumber yang ditinjau, bukan seluruh histori perusahaan |
| Saldo invoice | Nilai tagihan − pembayaran dikenal | Saldo negatif tetap terlihat sebagai lebih bayar |
| Sisa positif | Jumlah saldo invoice yang lebih besar dari nol | Tidak dikurangi lebih bayar invoice lain |
| Lebih bayar | Negatif dari jumlah saldo yang kurang dari nol | Dilaporkan terpisah untuk ditinjau |
| Pembayaran tanpa kunci | Jumlah pembayaran yang invoice_id tidak ada pada Data | Tidak dialokasikan ke invoice rekaan |
| Lewat tempo | Invoice dengan jatuh tempo sebelum tanggal acuan dan saldo positif | Jatuh tempo tepat pada acuan belum masuk kelompok lewat tempo |
| Rekonsiliasi nilai | Jumlah saldo = jumlah tagihan − pembayaran dikenal | Pemeriksaan matematis, bukan persetujuan bisnis |

Input awal berada sebelum tanggal acuan dan verifier memeriksa relasi serta tanggal. Jika sumber diperluas ke tanggal lain, saya menyaring periode dan memperbarui pembanding; tidak cukup hanya mengganti label laporan. Tidak ada pajak, biaya bank, kurs, diskon pembayaran, atau penyisihan piutang dalam simulasi.
