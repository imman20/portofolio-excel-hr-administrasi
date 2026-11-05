# Kunci jawaban modul 10

Saya memakai concat, merge many-to-one, dan melt pandas untuk pembanding. Query M belum dijalankan pada mesin Power Query.

| Pemeriksaan | Sel | Nilai |
|---|---|---:|
| Ringkasan 4 | C4 | 200 |
| Ringkasan 5 | C5 | 70996000 |
| Ringkasan 6 | C6 | 37798000 |
| Ringkasan 7 | C7 | 33198000 |
| Ringkasan 8 | C8 | 6 |
| Ringkasan 9 | C9 | 43500000 |
| Ringkasan 10 | C10 | 0 |

Append menambah baris; Merge menambah atribut. Master unik menjaga jumlah baris. Unpivot menghasilkan pasangan wilayah dan periode dengan total tetap. Penambahan master duplikat harus ditolak sebelum Expand; skrip M mempunyai pemeriksaan eksplisit untuk kondisi tersebut.
