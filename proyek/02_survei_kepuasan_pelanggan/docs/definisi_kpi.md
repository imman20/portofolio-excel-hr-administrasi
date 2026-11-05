# Definisi KPI dan pembagi

Rumus dalam dokumen memakai titik koma; Excel regional Inggris memakai koma. Saya memakai baris `dipakai=1` dan hanya nilai valid.

| Ukuran | Definisi | Pembagi dan missing |
|---|---|---|
| n bersih | SUM penanda dipakai | Seluruh ID unik, meskipun jawaban sebagian kosong |
| Distribusi item | COUNTIFS setiap skor 1–5 | Jawaban valid pada item itu; total frekuensi sama dengan n |
| Top-2-box | (frekuensi skor 4 + 5) / n valid | Tidak memasukkan missing |
| Skor dimensi | Mean tiga item responden, lalu mean skor responden | Responden dengan ketiga item valid; tanpa imputasi |
| SD | STDEV.S atas skor dimensi | Penyebut varians n−1 |
| CI 95% | Mean ± CONFIDENCE.T(0,05;SD;n) | t dengan df n−1; masing-masing dimensi punya n berbeda |
| Alpha | k/(k−1) × (1 − jumlah varians item / varians jumlah item) | k=3; responden lengkap sama pada semua varians |
| NPS | 100 × (n promotor − n detraktor) / n rekomendasi valid | Promotor 9–10; pasif 7–8; detraktor 0–6; skala −100–100 |
| Welch | T.TEST(kepuasan Toko;kepuasan Online;2;3) | Dua sampel independen; missing dikeluarkan per sampel |
| Cohen's d | Selisih mean / SD gabungan berbobot df | Ukuran deskriptif; bukan t Welch dan bukan d berpasangan |
| Korelasi / regresi | CORREL dan LINEST | Kasus lengkap keempat dimensi serta kepuasan keseluruhan |

Saya tidak merata-ratakan mean kanal tanpa bobot untuk mendapatkan mean total. Distribusi respons item gabungan dan distribusi skor dimensi adalah dua objek berbeda. CI dimensi memakai rata-rata per responden, bukan seluruh item sebagai observasi independen.

Saya mensyaratkan minimal dua observasi dan varians positif untuk statistik yang memerlukannya. Dataset tetap memenuhi syarat; jika data baru gagal syarat atau menghasilkan pembagi nol, verifier berhenti. Saya tidak mengganti statistik yang tidak terdefinisi dengan nol. Korelasi membutuhkan variasi dan regresi membutuhkan matriks rancangan berpangkat penuh.
