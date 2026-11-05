"""
Pembangkit data sintetis proyek "Sales & Profitability Analysis".

Semua data di sini adalah SIMULASI. Tidak ada data perusahaan atau pribadi nyata.
Skrip ini deterministik: seed tetap (numpy.random.default_rng(SEED)) sehingga
setiap kali dijalankan ulang menghasilkan CSV yang identik.

Cara menjalankan (dari folder proyek):
    python tools/generate_data.py

Keluaran (folder data/raw/):
    transaksi.csv, retur.csv, master_produk.csv, master_pelanggan.csv,
    master_wilayah.csv, master_sales.csv, target.csv
"""

from __future__ import annotations

import datetime as dt
from pathlib import Path

import numpy as np
import pandas as pd

SEED = 20240101
TAHUN = 2024
N_TRANSAKSI = 6000
N_PELANGGAN = 120
PROJECT_DIR = Path(__file__).resolve().parent.parent
RAW_DIR = PROJECT_DIR / "data" / "raw"

rng = np.random.default_rng(SEED)


# ---------------------------------------------------------------------------
# 1. Master wilayah dan sales
# ---------------------------------------------------------------------------
WILAYAH = [
    ("R01", "Jabodetabek", "Jawa"),
    ("R02", "Jawa Barat", "Jawa"),
    ("R03", "Jawa Tengah & DIY", "Jawa"),
    ("R04", "Jawa Timur", "Jawa"),
    ("R05", "Sumatera", "Sumatera"),
    ("R06", "Kalimantan", "Kalimantan"),
]
# Bobot aktivitas penjualan per region (Jawa dominan: 4 region pertama)
BOBOT_REGION = {"R01": 0.27, "R02": 0.15, "R03": 0.12, "R04": 0.17, "R05": 0.17, "R06": 0.12}
# Jumlah pelanggan per region (total 120)
PELANGGAN_PER_REGION = {"R01": 32, "R02": 18, "R03": 15, "R04": 20, "R05": 20, "R06": 15}

KOTA = {
    "R01": ["Jakarta Pusat", "Jakarta Selatan", "Jakarta Barat", "Jakarta Timur",
            "Jakarta Utara", "Bogor", "Depok", "Tangerang", "Bekasi"],
    "R02": ["Bandung", "Cimahi", "Cirebon", "Tasikmalaya", "Sukabumi", "Karawang"],
    "R03": ["Semarang", "Surakarta", "Yogyakarta", "Magelang", "Purwokerto", "Tegal"],
    "R04": ["Surabaya", "Malang", "Sidoarjo", "Kediri", "Jember", "Gresik"],
    "R05": ["Medan", "Palembang", "Pekanbaru", "Padang", "Bandar Lampung", "Jambi", "Batam"],
    "R06": ["Balikpapan", "Samarinda", "Banjarmasin", "Pontianak", "Palangka Raya"],
}

SALES = [
    ("S01", "Andi Prasetyo", "R01"), ("S02", "Dewi Lestari", "R01"),
    ("S03", "Budi Santoso", "R02"), ("S04", "Rina Wulandari", "R02"),
    ("S05", "Agus Setiawan", "R03"), ("S06", "Sri Rahayu", "R03"),
    ("S07", "Hendra Gunawan", "R04"), ("S08", "Fitri Handayani", "R04"),
    ("S09", "Rizky Ramadhan", "R05"), ("S10", "Nurul Hidayah", "R05"),
    ("S11", "Yoga Pratama", "R06"), ("S12", "Maya Puspita", "R06"),
]

# ---------------------------------------------------------------------------
# 2. Master produk (40 produk, 5 kategori x 8)
# ---------------------------------------------------------------------------
# (nama_produk, sub_kategori, harga_list). Margin list per kategori ditentukan di bawah.
PRODUK = {
    "Alat Tulis Kantor": [
        ("Pulpen Gel Hitam (box 12)", "Alat Tulis", 36000),
        ("Pensil 2B (box 12)", "Alat Tulis", 24000),
        ("Spidol Whiteboard (set 4)", "Alat Tulis", 32000),
        ("Stapler Sedang", "Alat Jilid", 45000),
        ("Isi Staples No.10 (box 20)", "Alat Jilid", 38000),
        ("Penghapus Karet (box 20)", "Alat Tulis", 30000),
        ("Penggaris Besi 30 cm", "Alat Ukur", 12000),
        ("Map Plastik Snelhecter (pak 12)", "Arsip", 42000),
    ],
    "Kertas & Cetak": [
        ("Kertas HVS A4 70 gsm (rim)", "Kertas", 52000),
        ("Kertas HVS A4 80 gsm (rim)", "Kertas", 60000),
        ("Kertas HVS F4 70 gsm (rim)", "Kertas", 56000),
        ("Kertas Foto A4 Glossy (pak 50)", "Kertas", 85000),
        ("Toner Laser Hitam", "Cetak", 650000),
        ("Tinta Printer Hitam 70 ml", "Cetak", 95000),
        ("Tinta Printer Warna (set 3)", "Cetak", 270000),
        ("Amplop Putih Polos (box 100)", "Kertas", 45000),
    ],
    "Elektronik Kantor": [
        ("Kalkulator 12 Digit", "Perangkat", 185000),
        ("Mesin Penghancur Kertas", "Perangkat", 1450000),
        ("Laminator A3", "Perangkat", 1250000),
        ("Mouse Nirkabel", "Aksesoris", 165000),
        ("Keyboard USB", "Aksesoris", 210000),
        ("Flashdisk 32 GB", "Aksesoris", 95000),
        ("Headset Kantor", "Aksesoris", 320000),
        ("Webcam HD", "Aksesoris", 480000),
    ],
    "Furnitur Kantor": [
        ("Kursi Kerja Staf", "Kursi", 950000),
        ("Kursi Direktur", "Kursi", 2650000),
        ("Meja Kerja 120 cm", "Meja", 1850000),
        ("Lemari Arsip 4 Laci", "Penyimpanan", 3200000),
        ("Rak Buku 5 Tingkat", "Penyimpanan", 1150000),
        ("Papan Tulis 120x90 cm", "Papan", 650000),
        ("Partisi Meja", "Meja", 780000),
        ("Filing Cabinet 2 Laci", "Penyimpanan", 1950000),
    ],
    "Perlengkapan Kebersihan": [
        ("Tisu Wajah (box 250)", "Tisu", 22000),
        ("Sabun Cuci Tangan 500 ml", "Sabun", 28000),
        ("Pembersih Lantai 1 L", "Pembersih", 24000),
        ("Kantong Sampah (pak 20)", "Sampah", 18000),
        ("Hand Sanitizer 1 L", "Sabun", 65000),
        ("Pel Lantai Putar", "Alat Bersih", 120000),
        ("Sapu Lantai", "Alat Bersih", 35000),
        ("Pengharum Ruangan", "Pembersih", 42000),
    ],
}
# Margin list (sebelum diskon) per kategori. Kertas & Cetak sengaja tipis.
MARGIN_LIST = {
    "Alat Tulis Kantor": 0.38,
    "Kertas & Cetak": 0.18,
    "Elektronik Kantor": 0.28,
    "Furnitur Kantor": 0.42,
    "Perlengkapan Kebersihan": 0.45,
}
# Bobot kategori dalam transaksi (Kertas paling sering dibeli)
BOBOT_KATEGORI = {
    "Alat Tulis Kantor": 0.27,
    "Kertas & Cetak": 0.24,
    "Elektronik Kantor": 0.13,
    "Furnitur Kantor": 0.10,
    "Perlengkapan Kebersihan": 0.26,
}

# ---------------------------------------------------------------------------
# 3. Musiman dan hari
# ---------------------------------------------------------------------------
BOBOT_BULAN = {1: 0.90, 2: 0.72, 3: 0.95, 4: 1.00, 5: 1.00, 6: 0.95,
               7: 1.00, 8: 1.00, 9: 1.05, 10: 1.10, 11: 1.32, 12: 1.48}
BOBOT_HARI = {0: 1.0, 1: 1.0, 2: 1.0, 3: 1.0, 4: 1.0, 5: 0.6, 6: 0.25}  # Senin..Minggu

KATA_TOKO = ["Sumber", "Maju", "Jaya", "Sinar", "Mitra", "Cahaya", "Berkah", "Abadi",
             "Sejahtera", "Makmur", "Mulia", "Harapan", "Karya", "Prima", "Sentosa",
             "Lestari", "Bahagia", "Gemilang", "Mandiri", "Utama"]
KATA_TOKO2 = ["Rejeki", "Bersama", "Baru", "Indah", "Agung", "Mulia", "Persada", "Niaga",
              "Sarana", "Perkasa", "Nusantara", "Pratama", "Sukses", "Cemerlang", "Santosa"]
KATA_PT = ["Nusantara", "Mitra", "Sinergi", "Prima", "Global", "Cipta", "Bumi", "Karya",
           "Mega", "Inti", "Surya", "Tirta", "Graha", "Dinamika", "Adi", "Sarana",
           "Bina", "Astra", "Citra", "Mandala"]
KATA_PT2 = ["Teknologi", "Logistik", "Energi", "Konstruksi", "Pangan", "Finansial",
            "Properti", "Digital", "Industri", "Manufaktur", "Medika", "Edukasi",
            "Transportasi", "Agro", "Telekomunikasi"]
KATA_PT3 = ["Sejahtera", "Abadi", "Mandiri", "Persada", "Perkasa", "Utama",
            "Indonesia", "Nusantara", "Makmur", "Lestari"]


def tanggal_str(d: dt.date) -> str:
    return d.isoformat()


def main() -> None:
    global rng
    rng = np.random.default_rng(SEED)
    RAW_DIR.mkdir(parents=True, exist_ok=True)

    # ---- master_wilayah
    df_wilayah = pd.DataFrame(WILAYAH, columns=["region_id", "nama_region", "pulau"])

    # ---- master_sales
    df_sales = pd.DataFrame(SALES, columns=["salesperson_id", "nama", "region_id"])

    # ---- master_produk
    rows = []
    i = 0
    for kategori, daftar in PRODUK.items():
        for nama, sub, harga in daftar:
            i += 1
            m = MARGIN_LIST[kategori] + rng.uniform(-0.025, 0.025)
            hpp = int(round(harga * (1 - m) / 100.0) * 100)
            rows.append((f"P{i:03d}", nama, kategori, sub, int(harga), hpp))
    df_produk = pd.DataFrame(rows, columns=["product_id", "nama_produk", "kategori",
                                            "sub_kategori", "harga_list", "hpp_satuan"])

    # Popularitas tiap produk di dalam kategorinya (Dirichlet supaya tidak seragam)
    bobot_produk = np.zeros(len(df_produk))
    for kategori in PRODUK:
        idx = np.where(df_produk["kategori"] == kategori)[0]
        w = rng.dirichlet(np.ones(len(idx)) * 3.0)
        bobot_produk[idx] = w * BOBOT_KATEGORI[kategori]
    bobot_produk /= bobot_produk.sum()

    # ---- master_pelanggan
    pelanggan = []
    segmen_pool = ["Retail"] * 60 + ["Grosir"] * 40 + ["Korporat"] * 20
    segmen_pool = list(rng.permutation(segmen_pool))
    nama_terpakai = set()
    c = 0
    for region_id, n in PELANGGAN_PER_REGION.items():
        for _ in range(n):
            c += 1
            segmen = segmen_pool[c - 1]
            # Pelanggan pertama dipaksa Korporat besar (untuk analisis Pareto)
            if c == 1:
                segmen = "Korporat"
            while True:
                if segmen == "Retail":
                    nama = f"Toko {rng.choice(KATA_TOKO)} {rng.choice(KATA_TOKO2)}"
                elif segmen == "Grosir":
                    prefix = rng.choice(["CV", "UD"])
                    nama = f"{prefix} {rng.choice(KATA_TOKO)} {rng.choice(KATA_TOKO2)}"
                else:
                    nama = f"PT {rng.choice(KATA_PT)} {rng.choice(KATA_PT2)} {rng.choice(KATA_PT3)}"
                if nama not in nama_terpakai:
                    nama_terpakai.add(nama)
                    break
            kota = rng.choice(KOTA[region_id])
            # Tanggal bergabung: 2019-01-01 s.d. 2024-06-30; mayoritas sebelum 2024
            if rng.random() < 0.92:
                hari = int(rng.integers(0, (dt.date(2023, 12, 31) - dt.date(2019, 1, 1)).days + 1))
                bergabung = dt.date(2019, 1, 1) + dt.timedelta(days=hari)
            else:
                hari = int(rng.integers(0, (dt.date(2024, 6, 30) - dt.date(2024, 1, 1)).days + 1))
                bergabung = dt.date(2024, 1, 1) + dt.timedelta(days=hari)
            # Bobot aktivitas pelanggan (lognormal) menurut segmen
            if segmen == "Retail":
                bobot = rng.lognormal(mean=0.0, sigma=0.45)
            elif segmen == "Grosir":
                bobot = 2.2 * rng.lognormal(mean=0.0, sigma=0.45)
            else:
                bobot = 4.0 * rng.lognormal(mean=0.0, sigma=0.45)
            if c == 1:
                bobot = 22.0  # pelanggan korporat besar
            pelanggan.append((f"C{c:03d}", nama, segmen, kota, region_id, bergabung, bobot))
    df_pelanggan_full = pd.DataFrame(pelanggan, columns=[
        "customer_id", "nama_pelanggan", "segmen", "kota", "region_id",
        "tanggal_bergabung", "bobot"])
    # Pelanggan C001 ditempatkan di Jabodetabek dan segmen Korporat (sudah via c == 1)
    df_pelanggan = df_pelanggan_full.drop(columns=["bobot"]).copy()
    df_pelanggan["tanggal_bergabung"] = df_pelanggan["tanggal_bergabung"].map(tanggal_str)

    # ---- tanggal transaksi: bobot per hari = bobot bulan x bobot hari kerja
    hari_list = []
    bobot_hari = []
    d = dt.date(TAHUN, 1, 1)
    while d.year == TAHUN:
        hari_list.append(d)
        bobot_hari.append(BOBOT_BULAN[d.month] * BOBOT_HARI[d.weekday()])
        d += dt.timedelta(days=1)
    bobot_hari = np.array(bobot_hari)
    bobot_hari /= bobot_hari.sum()
    idx_hari = rng.choice(len(hari_list), size=N_TRANSAKSI, p=bobot_hari)
    tanggal = sorted(hari_list[i] for i in idx_hari)

    # ---- bobot pelanggan efektif = bobot aktivitas x bobot region
    bobot_cust = df_pelanggan_full["bobot"].to_numpy().copy()
    for r, w in BOBOT_REGION.items():
        mask = (df_pelanggan_full["region_id"] == r).to_numpy()
        # Normalisasi supaya total bobot region mengikuti BOBOT_REGION
        bobot_cust[mask] = bobot_cust[mask] / bobot_cust[mask].sum() * w
    bobot_cust /= bobot_cust.sum()
    bergabung_arr = df_pelanggan_full["tanggal_bergabung"].to_numpy()
    cust_region = df_pelanggan_full["region_id"].to_numpy()
    cust_segmen = df_pelanggan_full["segmen"].to_numpy()
    sales_by_region = {r: [s for s, _, rr in SALES if rr == r] for r in BOBOT_REGION}

    harga_list_arr = df_produk["harga_list"].to_numpy()
    kategori_arr = df_produk["kategori"].to_numpy()
    sub_kategori_arr = df_produk["sub_kategori"].to_numpy()

    trans = []
    for n, tgl in enumerate(tanggal, start=1):
        # pelanggan yang sudah bergabung pada tanggal tersebut
        eligible = np.array([b <= tgl for b in bergabung_arr])
        p = bobot_cust * eligible
        p /= p.sum()
        ci = int(rng.choice(len(p), p=p))
        customer_id = f"C{ci + 1:03d}"
        region_id = cust_region[ci]
        segmen = cust_segmen[ci]
        salesperson_id = str(rng.choice(sales_by_region[region_id]))

        pi = int(rng.choice(len(bobot_produk), p=bobot_produk))
        product_id = f"P{pi + 1:03d}"
        harga_list = int(harga_list_arr[pi])
        kategori = kategori_arr[pi]

        # kuantitas menurut segmen; furnitur dan elektronik dibeli lebih sedikit
        if segmen == "Retail":
            qty = int(rng.integers(1, 7))
        elif segmen == "Grosir":
            qty = int(rng.integers(5, 41))
        else:
            qty = int(rng.integers(10, 81))
        if kategori == "Furnitur Kantor":
            qty = max(1, qty // 6)
        elif kategori == "Elektronik Kantor":
            qty = max(1, qty // 3)
        elif sub_kategori_arr[pi] == "Cetak":
            qty = max(1, qty // 3)
        if ci == 0:
            qty = int(round(qty * 1.4))

        # harga satuan: Retail memakai harga list; Grosir/Korporat sedikit negosiasi
        if segmen == "Retail":
            harga_satuan = harga_list
        else:
            harga_satuan = int(round(harga_list * rng.uniform(0.98, 1.02) / 100.0) * 100)

        # diskon menurut segmen, promo tambahan Nov-Des
        if segmen == "Retail":
            diskon = float(rng.choice([0.0, 0.05, 0.10], p=[0.70, 0.20, 0.10]))
        elif segmen == "Grosir":
            diskon = float(rng.choice([0.05, 0.10, 0.15], p=[0.40, 0.40, 0.20]))
        else:
            diskon = float(rng.choice([0.10, 0.15, 0.20, 0.25], p=[0.45, 0.35, 0.15, 0.05]))
        if tgl.month >= 11 and rng.random() < 0.30:
            diskon = min(0.25, diskon + 0.05)
        diskon = round(diskon, 2)

        trans.append((f"SO-{TAHUN}-{n:06d}", tanggal_str(tgl), customer_id, product_id,
                      region_id, salesperson_id, qty, harga_satuan, diskon))

    df_trans = pd.DataFrame(trans, columns=[
        "order_id", "tanggal", "customer_id", "product_id", "region_id",
        "salesperson_id", "qty", "harga_satuan", "diskon_pct"])

    # ---- retur (sekitar 2 % order)
    n_retur = int(round(N_TRANSAKSI * 0.02))
    kandidat = df_trans.index[pd.to_datetime(df_trans["tanggal"]) <= pd.Timestamp(TAHUN, 12, 17)]
    idx_retur = np.sort(rng.choice(kandidat, size=n_retur, replace=False))
    ALASAN = ["Rusak saat pengiriman", "Salah kirim produk", "Tidak sesuai pesanan",
              "Cacat produk", "Kelebihan kirim"]
    retur = []
    for k, i in enumerate(idx_retur, start=1):
        row = df_trans.loc[i]
        qty = int(row["qty"])
        if rng.random() < 0.15:
            qty_retur = qty  # retur penuh -> penjualan bersih baris = 0 (uji IFERROR)
        else:
            qty_retur = int(min(qty, 1 + rng.geometric(0.45) - 1))
            qty_retur = max(1, min(qty_retur, qty))
        tgl = dt.date.fromisoformat(row["tanggal"]) + dt.timedelta(days=int(rng.integers(2, 15)))
        retur.append((f"RT-{TAHUN}-{k:04d}", row["order_id"], tanggal_str(tgl), qty_retur,
                      str(rng.choice(ALASAN, p=[0.35, 0.20, 0.20, 0.15, 0.10]))))
    df_retur = pd.DataFrame(retur, columns=["retur_id", "order_id", "tanggal_retur",
                                            "qty_retur", "alasan"])

    # ---- target per bulan x region
    # Target disusun dari baseline simulasi: penjualan bersih aktual per region-bulan
    # dikalikan faktor ambisi per region (sebagian region ditarget lebih tinggi dari
    # realisasinya) plus gangguan acak, lalu dibulatkan ke jutaan rupiah.
    hitung = df_trans.merge(df_produk[["product_id", "hpp_satuan"]], on="product_id")
    qty_retur_map = df_retur.groupby("order_id")["qty_retur"].sum()
    hitung["qty_retur"] = hitung["order_id"].map(qty_retur_map).fillna(0).astype(int)
    hitung["omzet_bruto"] = hitung["qty"] * hitung["harga_satuan"]
    hitung["nilai_diskon"] = hitung["omzet_bruto"] * hitung["diskon_pct"]
    hitung["omzet_neto"] = hitung["omzet_bruto"] - hitung["nilai_diskon"]
    hitung["nilai_retur"] = hitung["qty_retur"] * hitung["harga_satuan"] * (1 - hitung["diskon_pct"])
    hitung["penjualan_bersih"] = hitung["omzet_neto"] - hitung["nilai_retur"]
    hitung["bulan"] = hitung["tanggal"].str[:7]
    FAKTOR_TARGET = {"R01": 0.94, "R02": 1.00, "R03": 1.06, "R04": 0.97, "R05": 1.12, "R06": 1.08}
    aktual = hitung.groupby(["bulan", "region_id"])["penjualan_bersih"].sum()
    target = []
    for bulan in [f"{TAHUN}-{m:02d}" for m in range(1, 13)]:
        for region_id, _, _ in WILAYAH:
            dasar = float(aktual.get((bulan, region_id), 0.0))
            nilai = dasar * FAKTOR_TARGET[region_id] * (1 + rng.normal(0, 0.05))
            target.append((bulan, region_id, int(round(nilai / 1_000_000.0) * 1_000_000)))
    df_target = pd.DataFrame(target, columns=["bulan", "region_id", "target_penjualan_bersih"])

    # ---- simpan
    df_trans.to_csv(RAW_DIR / "transaksi.csv", index=False)
    df_retur.to_csv(RAW_DIR / "retur.csv", index=False)
    df_produk.to_csv(RAW_DIR / "master_produk.csv", index=False)
    df_pelanggan.to_csv(RAW_DIR / "master_pelanggan.csv", index=False)
    df_wilayah.to_csv(RAW_DIR / "master_wilayah.csv", index=False)
    df_sales.to_csv(RAW_DIR / "master_sales.csv", index=False)
    df_target.to_csv(RAW_DIR / "target.csv", index=False)

    # ---- ringkasan ke layar
    print(f"transaksi      : {len(df_trans):>6} baris")
    print(f"retur          : {len(df_retur):>6} baris")
    print(f"master_produk  : {len(df_produk):>6} baris")
    print(f"master_pelanggan: {len(df_pelanggan):>5} baris")
    print(f"target         : {len(df_target):>6} baris")
    print(f"total omzet bruto     : {hitung['omzet_bruto'].sum():,.0f}")
    print(f"total penjualan bersih: {hitung['penjualan_bersih'].sum():,.2f}")
    hitung["hpp"] = (hitung["qty"] - hitung["qty_retur"]) * hitung["hpp_satuan"]
    hitung["laba_kotor"] = hitung["penjualan_bersih"] - hitung["hpp"]
    print(f"total laba kotor      : {hitung['laba_kotor'].sum():,.2f}")
    print(f"margin kotor          : {hitung['laba_kotor'].sum() / hitung['penjualan_bersih'].sum():.4f}")
    print(f"pencapaian target     : {hitung['penjualan_bersih'].sum() / df_target['target_penjualan_bersih'].sum():.4f}")
    print("per bulan (penjualan bersih, juta):")
    print((hitung.groupby("bulan")["penjualan_bersih"].sum() / 1e6).round(1).to_string())
    print("per region (juta):")
    print((hitung.groupby("region_id")["penjualan_bersih"].sum() / 1e6).round(1).to_string())
    kat = hitung.merge(df_produk[["product_id", "kategori"]], on="product_id").groupby("kategori")
    print("per kategori (penjualan bersih juta, margin):")
    print(pd.DataFrame({"pb_juta": (kat["penjualan_bersih"].sum() / 1e6).round(1),
                        "margin": (kat["laba_kotor"].sum() / kat["penjualan_bersih"].sum()).round(4)}).to_string())
    cust = hitung.groupby("customer_id")["penjualan_bersih"].sum().sort_values(ascending=False)
    kum = cust.cumsum() / cust.sum()
    print(f"20% pelanggan teratas (24) menyumbang: {kum.iloc[23]:.4f}")
    print(f"pelanggan teratas: {cust.index[0]} share {cust.iloc[0] / cust.sum():.4f}")


if __name__ == "__main__":
    main()
