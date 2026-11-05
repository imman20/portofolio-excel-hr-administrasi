"""Pembangkit data sintetis survei kepuasan pelanggan.

Menghasilkan `data/raw/survei_raw.csv` (420 baris) yang berisi masalah kualitas
data yang DISENGAJA (lihat docs/masalah_kualitas_data.md) dan
`data/processed/survei_bersih_referensi.csv` (414 baris) sebagai pembanding
hasil pembersihan di Excel. Seluruh proses deterministik (seed tetap).

Jalankan: python tools/generate_data.py
"""
from __future__ import annotations

import json
from pathlib import Path

import numpy as np
import pandas as pd

SEED = 20240214
BASE = Path(__file__).resolve().parents[1]
RAW_DIR = BASE / "data" / "raw"
PROC_DIR = BASE / "data" / "processed"

N_UNIK = 414          # responden unik
N_DUPLIKAT = 6        # baris duplikat yang disisipkan -> total 420 baris

KANAL = ["Toko", "Online", "Call Center"]
P_KANAL = [0.45, 0.35, 0.20]
KOTA = ["Surabaya", "Jakarta", "Bandung", "Semarang", "Medan", "Makassar"]
P_KOTA = [0.28, 0.24, 0.15, 0.12, 0.11, 0.10]
SEGMEN = ["Retail", "Grosir", "Korporat"]
P_SEGMEN = [0.60, 0.25, 0.15]

DIMENSI = {
    "Produk": ["q01", "q02", "q03"],
    "Layanan": ["q04", "q05", "q06"],
    "Harga": ["q07", "q08", "q09"],
    "Pengiriman": ["q10", "q11", "q12"],
}
MEAN_DIM = {"Produk": 4.05, "Layanan": 3.75, "Harga": 3.35, "Pengiriman": 3.55}
EFEK_KANAL = {
    "Toko": {"Layanan": 0.20, "Pengiriman": 0.15},
    "Online": {"Layanan": -0.15, "Pengiriman": -0.55},
    "Call Center": {"Layanan": -0.45, "Produk": -0.05},
}
EFEK_SEGMEN = {"Korporat": {"Harga": -0.30}, "Grosir": {"Harga": -0.10}}

KOMENTAR_POSITIF = [
    "Produk sesuai deskripsi, pengiriman cepat.",
    "Pelayanan ramah dan jelas.",
    "Kualitas bagus untuk harganya.",
    "Proses pembelian mudah.",
    "Barang sampai dengan aman.",
    "Staf toko membantu memilih produk.",
]
KOMENTAR_NEGATIF = [
    "Pengiriman lebih lama dari perkiraan.",
    "Harga naik dibanding bulan lalu.",
    "Sulit menghubungi call center.",
    "Kemasan sedikit penyok saat tiba.",
    "Stok sering kosong untuk varian tertentu.",
    "Respons chat lambat.",
]
KOMENTAR_NETRAL = [
    "Cukup baik.",
    "Biasa saja.",
    "Tidak ada keluhan khusus.",
    "Semoga ada promo lagi.",
]


def bangkitkan_bersih(rng: np.random.Generator) -> pd.DataFrame:
    n = N_UNIK
    kanal = rng.choice(KANAL, size=n, p=P_KANAL)
    kota = rng.choice(KOTA, size=n, p=P_KOTA)
    segmen = rng.choice(SEGMEN, size=n, p=P_SEGMEN)
    usia = np.clip(np.round(rng.normal(36, 10, n)), 18, 68).astype(int)
    jk = rng.choice(["L", "P"], size=n, p=[0.52, 0.48])
    lama = np.clip(np.round(rng.gamma(2.0, 12.0, n)), 1, 120).astype(int)
    hari = rng.integers(0, 366, n)
    tanggal = pd.Timestamp("2024-01-01") + pd.to_timedelta(hari, unit="D")

    baseline = rng.normal(0, 0.55, n)  # kecenderungan umum responden
    df = pd.DataFrame(
        {
            "respondent_id": [f"R-{i:04d}" for i in range(1, n + 1)],
            "tanggal_survei": tanggal,
            "kanal": kanal,
            "kota": kota,
            "segmen": segmen,
            "usia": usia,
            "jenis_kelamin": jk,
            "lama_pelanggan_bulan": lama,
        }
    )
    laten = {}
    for dim, items in DIMENSI.items():
        efek = np.array([EFEK_KANAL.get(k, {}).get(dim, 0.0) for k in kanal])
        efek += np.array([EFEK_SEGMEN.get(s, {}).get(dim, 0.0) for s in segmen])
        lat = MEAN_DIM[dim] + efek + baseline + rng.normal(0, 0.45, n)
        laten[dim] = lat
        for it in items:
            skor = np.clip(np.round(lat + rng.normal(0, 0.50, n)), 1, 5)
            df[it] = skor.astype(int)
    kk_lat = (
        0.45
        + 0.28 * laten["Produk"]
        + 0.30 * laten["Layanan"]
        + 0.14 * laten["Harga"]
        + 0.22 * laten["Pengiriman"]
        + rng.normal(0, 0.35, n)
    )
    df["kepuasan_keseluruhan"] = np.clip(np.round(kk_lat), 1, 5).astype(int)
    nps_lat = 2.05 * kk_lat - 0.6 + rng.normal(0, 1.1, n)
    df["nps"] = np.clip(np.round(nps_lat), 0, 10).astype(int)

    komentar = []
    for kk in df["kepuasan_keseluruhan"]:
        u = rng.random()
        if u < 0.42:
            komentar.append("")
        elif kk >= 4:
            komentar.append(rng.choice(KOMENTAR_POSITIF if rng.random() < 0.8 else KOMENTAR_NETRAL))
        elif kk <= 2:
            komentar.append(rng.choice(KOMENTAR_NEGATIF if rng.random() < 0.8 else KOMENTAR_NETRAL))
        else:
            komentar.append(rng.choice(KOMENTAR_NETRAL + KOMENTAR_NEGATIF[:2]))
    df["komentar"] = komentar
    return df


def sisipkan_masalah(bersih: pd.DataFrame, rng: np.random.Generator):
    """Mengembalikan (df_raw_teks, catatan) dengan masalah yang disengaja."""
    df = bersih.copy()
    item_cols = list(DIMENSI["Produk"] + DIMENSI["Layanan"] + DIMENSI["Harga"] + DIMENSI["Pengiriman"])
    likert_cols = item_cols + ["kepuasan_keseluruhan"]
    semua_skor = likert_cols + ["nps"]

    # Semua kolom jadi teks agar variasi format bisa disimpan apa adanya
    raw = df.copy()
    raw["tanggal_survei"] = raw["tanggal_survei"].dt.strftime("%Y-%m-%d")
    raw = raw.astype(str)
    raw.loc[df["komentar"] == "", "komentar"] = ""

    catatan: dict[str, int] = {}
    n = len(raw)
    sel = [(r, c) for r in range(n) for c in semua_skor]
    rng.shuffle(sel)
    ptr = 0

    def ambil(k):
        nonlocal ptr
        out = sel[ptr : ptr + k]
        ptr += k
        return out

    # 1. sel kosong ±3 % dari sel skor (414 x 14 = 5.796 sel -> 174)
    k_kosong = int(round(0.03 * n * len(semua_skor)))
    for r, c in ambil(k_kosong):
        raw.iat[r, raw.columns.get_loc(c)] = ""
    catatan["sel_kosong_skor"] = k_kosong

    # 2. teks "N/A" dan "-"
    for r, c in ambil(12):
        raw.iat[r, raw.columns.get_loc(c)] = "N/A"
    catatan["teks_na"] = 12
    for r, c in ambil(10):
        raw.iat[r, raw.columns.get_loc(c)] = "-"
    catatan["teks_strip"] = 10

    # 3. angka dengan spasi (hanya pada item 1-5 agar tetap valid setelah TRIM)
    k_spasi = 15
    pola = [" {}", "{} ", "  {}", "{}  "]
    for i, (r, c) in enumerate(ambil(k_spasi)):
        v = raw.iat[r, raw.columns.get_loc(c)]
        raw.iat[r, raw.columns.get_loc(c)] = pola[i % 4].format(v)
    catatan["angka_berspasi"] = k_spasi

    # 4. nilai di luar rentang pada item Likert / kepuasan (0, 6, 7)
    luar = [0] * 5 + [6] * 5 + [7] * 3
    sel_likert = [(r, c) for (r, c) in sel[ptr:] if c in likert_cols][: len(luar)]
    for (r, c), v in zip(sel_likert, luar):
        raw.iat[r, raw.columns.get_loc(c)] = str(v)
    # geser pointer sejumlah yang dipakai (agar tidak tertimpa masalah lain)
    sel[ptr:] = [x for x in sel[ptr:] if x not in sel_likert]
    catatan["luar_rentang_0"] = 5
    catatan["luar_rentang_6"] = 5
    catatan["luar_rentang_7"] = 3

    # 5. usia outlier 150 dan 3
    idx_usia = rng.choice(n, size=2, replace=False)
    raw.iat[idx_usia[0], raw.columns.get_loc("usia")] = "150"
    raw.iat[idx_usia[1], raw.columns.get_loc("usia")] = "3"
    catatan["usia_outlier"] = 2

    # 6. format tanggal campuran: ~15 % jadi dd/mm/yyyy
    k_tgl = int(round(0.15 * n))
    idx_tgl = rng.choice(n, size=k_tgl, replace=False)
    for r in idx_tgl:
        raw.iat[r, raw.columns.get_loc("tanggal_survei")] = df["tanggal_survei"].iloc[r].strftime("%d/%m/%Y")
    catatan["tanggal_format_dmy"] = k_tgl

    # 7. kanal huruf besar-kecil tidak konsisten
    varian = {
        "Toko": ["toko", "TOKO"],
        "Online": ["online", "ONLINE", "OnLine"],
        "Call Center": ["call center", "CALL CENTER", "Call center"],
    }
    k_kanal = 34
    idx_kanal = rng.choice(n, size=k_kanal, replace=False)
    for r in idx_kanal:
        k = df["kanal"].iloc[r]
        raw.iat[r, raw.columns.get_loc("kanal")] = rng.choice(varian[k])
    catatan["kanal_kapitalisasi"] = k_kanal

    # 8. enam baris duplikat respondent_id (salinan persis) disisipkan acak
    idx_dup = rng.choice(n, size=N_DUPLIKAT, replace=False)
    dup_rows = raw.iloc[idx_dup].copy()
    posisi = sorted(rng.choice(np.arange(1, n), size=N_DUPLIKAT, replace=False), reverse=True)
    bagian = [raw]
    raw_list = raw.reset_index(drop=True)
    for p, (_, row) in zip(posisi, dup_rows.iterrows()):
        raw_list = pd.concat([raw_list.iloc[:p], row.to_frame().T, raw_list.iloc[p:]], ignore_index=True)
    raw = raw_list
    catatan["baris_duplikat"] = N_DUPLIKAT
    catatan["duplikat_id"] = sorted(dup_rows["respondent_id"].tolist())
    catatan["total_baris_raw"] = len(raw)
    catatan["cakupan_hitungan"] = "Kunci masalah di tingkat atas dihitung pada 414 ID unik sebelum duplikasi; hitungan_csv_akhir mencakup seluruh 420 baris."
    from analysis_support import counts
    catatan["hitungan_csv_akhir"] = counts(raw)
    return raw, catatan


def referensi_bersih(raw: pd.DataFrame) -> pd.DataFrame:
    """Meniru aturan pembersihan yang dipakai di workbook (untuk rekonsiliasi)."""
    df = raw.copy()
    df = df.drop_duplicates(subset="respondent_id", keep="first").reset_index(drop=True)

    def ke_angka(s, lo, hi):
        v = pd.to_numeric(s.str.strip(), errors="coerce")
        return v.where((v >= lo) & (v <= hi))

    items = [f"q{i:02d}" for i in range(1, 13)]
    for c in items + ["kepuasan_keseluruhan"]:
        df[c] = ke_angka(df[c], 1, 5)
    df["nps"] = ke_angka(df["nps"], 0, 10)
    df["usia"] = ke_angka(df["usia"], 17, 80)
    df["lama_pelanggan_bulan"] = pd.to_numeric(df["lama_pelanggan_bulan"], errors="coerce")
    df["kanal"] = df["kanal"].str.strip().str.lower().map({"toko": "Toko", "online": "Online", "call center": "Call Center"})
    df["tanggal_survei"] = df["tanggal_survei"].map(
        lambda x: pd.to_datetime(x, format="%Y-%m-%d" if "-" in x else "%d/%m/%Y").strftime("%Y-%m-%d")
    )
    for dim, cols in DIMENSI.items():
        sub = df[cols]
        df[f"dim_{dim.lower()}"] = sub.mean(axis=1).where(sub.notna().all(axis=1))
    return df


def main():
    rng = np.random.default_rng(SEED)
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    PROC_DIR.mkdir(parents=True, exist_ok=True)
    bersih = bangkitkan_bersih(rng)
    raw, catatan = sisipkan_masalah(bersih, rng)
    raw.to_csv(RAW_DIR / "survei_raw.csv", index=False)
    ref = referensi_bersih(raw)
    ref.to_csv(PROC_DIR / "survei_bersih_referensi.csv", index=False)
    with open(PROC_DIR / "catatan_masalah_disengaja.json", "w", encoding="utf-8") as f:
        json.dump(catatan, f, indent=2, ensure_ascii=False)
    print(f"survei_raw.csv: {len(raw)} baris; bersih referensi: {len(ref)} baris")
    print(json.dumps(catatan, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
