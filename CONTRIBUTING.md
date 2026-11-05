# Panduan Kontribusi

Terima kasih sudah tertarik membantu. Repositori ini terutama portofolio pribadi, tetapi perbaikan rumus, koreksi tutorial, dan soal latihan baru selalu diterima.

1. Buka *issue* lebih dulu untuk perubahan besar agar arahnya sejalan.
2. Fork, lalu buat branch dari `main` dengan nama deskriptif (`fix/rumus-pph21`, `docs/tutorial-cuti`).
3. Jangan mengedit file `.xlsx` langsung; ubah skrip di `proyek-hr/tools/` (atau `tools/build.py` modul belajar) lalu bangun ulang dengan `python proyek-hr/tools/build_all.py` sehingga semua file tetap dapat direproduksi.
4. Pastikan `proyek-hr/tools/verifikasi.md` dan `VERIFIKASI.md` modul menunjukkan nol sel error setelah pembangunan ulang.
5. Gunakan Bahasa Indonesia baku, pemisah rumus titik koma (;) di dokumen tutorial, dan pesan commit berformat Conventional Commits (`feat:`, `fix:`, `docs:`).
6. Kirim *pull request* dengan ringkasan perubahan dan screenshot bila menyangkut tampilan sheet.
