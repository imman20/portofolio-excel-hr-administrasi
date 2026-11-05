# Keamanan dan proteksi

Saya menyediakan `latihan/template_15.xlsx` dengan proteksi sheet untuk mengurangi perubahan formula yang tidak disengaja. Password contoh adalah `latihan`; password ini sengaja dipublikasikan sebagai bagian materi. Input Data, Pembayaran, dan Master!B2 dibuka untuk edit, sedangkan Calculation, pemeriksaan, serta output tetap terkunci. Formula tidak disembunyikan agar dapat diaudit.

Proteksi sheet bukan enkripsi dan tidak menjamin kerahasiaan. Saya tidak memasukkan data nyata, token, atau kredensial ke template. Pada penerapan sesungguhnya, akses penyimpanan, berbagi file, enkripsi, kebijakan retensi, dan pencadangan perlu ditetapkan sesuai kebutuhan serta kebijakan yang berlaku. Tidak ada penerapan enkripsi file yang diklaim pada contoh ini.

Untuk meninjau proteksi di Excel: pilih Review > Unprotect Sheet, masukkan password contoh, lalu periksa Format Cells > Protection pada sel input dan formula. Setelah perubahan yang disetujui, pilih Review > Protect Sheet dan tetapkan izin yang diperlukan. Jangan menyamakan kemampuan membuka proteksi dengan otorisasi mengubah data.

Saya menguji metadata proteksi serta hasil recalc secara programatis; interaksi proteksi di Excel/Google Sheets belum diuji langsung. Impor ke aplikasi lain dapat mengubah perilaku proteksi. Setelah perubahan input, saya menghitung ulang dan memeriksa seluruh rekonsiliasi; proteksi tidak mencegah kesalahan input atau membuktikan validitas bisnis.
