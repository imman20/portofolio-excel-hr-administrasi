# Kunci jawaban modul 12

Saya memakai pandas/scipy; p-value adalah ilustrasi pada data sintetis dan tidak membuktikan efek nyata.

| Ukuran | Sel | Nilai |
|---|---|---:|
| n | C4 | 60.0 |
| Mean | C5 | 70.82633333333334 |
| Median | C6 | 71.72 |
| SD sampel | C7 | 6.797683046616685 |
| Varians sampel | C8 | 46.2084948022599 |
| Minimum | C9 | 54.55 |
| Maksimum | C10 | 85.98 |
| Margin CI | C11 | 1.7560277365133417 |
| CI bawah | C12 | 69.07030559681999 |
| CI atas | C13 | 72.58236106984668 |
| Pearson r | C14 | 0.8013895905942389 |
| Slope OLS | C15 | 2.3311880505283344 |
| Intercept OLS | C16 | 56.517501079190424 |
| R kuadrat | C17 | 0.642225275912802 |
| p Welch | C18 | 0.0011528540254506655 |
| p paired | C19 | 1.3743300264479845e-24 |
| Cohen d | C20 | 0.8851863034599317 |
| p equal variance | C21 | 0.001123660685938495 |
| FORECAST | C22 | 75.1670054834171 |
| TREND | C23 | 75.1670054834171 |
| Selisih prediksi | C24 | 0.0 |
| Ramalan periode 25 | C26 | 396.78985507246375 |
| Moving average terakhir | C27 | 384.0 |

Welch adalah uji utama perbandingan kelompok independen; paired memakai pengukuran awal/akhir pada ID sama. Equal-variance hanya pembanding metodologis. Cohen d memakai selisih mean B minus A dibagi SD pooled, bukan ukuran efek berpasangan. CI mengasumsikan pengamatan independen; korelasi dan regresi tidak membuktikan sebab-akibat. Histogram memakai interval sama lebar dan interval terakhir menyertakan batas atas. Moving average memakai tiga observasi berjalan, bukan validasi ramalan di luar sampel.
