"""Generator data sintetis (nama Indonesia, departemen, jabatan, gaji) dengan seed tetap."""

from __future__ import annotations

import datetime as dt
import random

from util import DEPARTEMEN

NAMA_DEPAN_L = [
    "Agus", "Bagus", "Budi", "Dimas", "Eko", "Fajar", "Gilang", "Hendra", "Irfan", "Joko",
    "Kurniawan", "Lukman", "Mahendra", "Nanda", "Oki", "Putra", "Rizky", "Satria", "Teguh", "Wahyu",
    "Yoga", "Zaki", "Adi", "Bayu", "Candra", "Dedi", "Erwin", "Farhan", "Galih", "Hafiz",
    "Ilham", "Johan", "Krisna", "Lutfi", "Miftah", "Naufal", "Panji", "Raka", "Surya", "Taufik",
    "Umar", "Vino", "Wisnu", "Yusuf", "Arif", "Bima", "Danang", "Fikri", "Hanif", "Rendi",
]
NAMA_DEPAN_P = [
    "Ayu", "Bella", "Citra", "Dewi", "Eka", "Fitri", "Gita", "Hana", "Indah", "Jihan",
    "Kartika", "Laras", "Maya", "Nadia", "Oktavia", "Putri", "Rani", "Sari", "Tiara", "Utami",
    "Vina", "Wulan", "Yuni", "Zahra", "Anisa", "Bunga", "Clara", "Dina", "Elsa", "Farah",
    "Galuh", "Hesti", "Intan", "Julia", "Kirana", "Lestari", "Mega", "Nabila", "Prita", "Ratih",
    "Salsabila", "Tasya", "Ulfa", "Vania", "Widya", "Yasmin", "Amelia", "Diah", "Ratna", "Sinta",
]
NAMA_BELAKANG = [
    "Pratama", "Saputra", "Wijaya", "Santoso", "Kusuma", "Hidayat", "Nugroho", "Setiawan", "Rahmawati",
    "Lestari", "Permata", "Anggraini", "Firmansyah", "Ramadhan", "Maulana", "Hakim", "Susanto", "Prasetyo",
    "Utomo", "Wibowo", "Handayani", "Puspita", "Kurnia", "Fauzi", "Syahputra", "Siregar", "Nasution",
    "Simanjuntak", "Hutapea", "Gunawan", "Halim", "Tanuwijaya", "Mulyani", "Suryani", "Oktaviani",
    "Pangestu", "Ardiansyah", "Mahardika", "Yuliana", "Hartono", "Sudrajat", "Iskandar", "Arifin",
    "Rosyidi", "Zulkifli", "Basuki", "Cahyono", "Darmawan", "Effendi", "Febrianto",
]
KOTA = ["Surabaya", "Sidoarjo", "Gresik", "Malang", "Mojokerto", "Lamongan", "Pasuruan", "Jombang"]
PENDIDIKAN = ["SMA/SMK", "D3", "S1", "S2"]

JABATAN = {
    "HR": ["Staff HR", "Senior Staff HR", "Supervisor HR", "Manager HR"],
    "Finance": ["Staff Akuntansi", "Senior Staff Keuangan", "Supervisor Keuangan", "Manager Finance"],
    "Operasional": ["Staff Operasional", "Senior Staff Operasional", "Supervisor Operasional", "Manager Operasional"],
    "Marketing": ["Staff Marketing", "Senior Staff Marketing", "Supervisor Marketing", "Manager Marketing"],
    "IT": ["Staff IT Support", "Senior Programmer", "Supervisor IT", "Manager IT"],
    "Produksi": ["Operator Produksi", "Senior Operator", "Supervisor Produksi", "Manager Produksi"],
}
LEVEL = ["Staff", "Senior Staff", "Supervisor", "Manager"]
# rentang gaji pokok per level (rupiah), acuan terendah UMK Surabaya 2024 = Rp4.725.479
GAJI_LEVEL = {
    "Staff": (4_750_000, 6_500_000),
    "Senior Staff": (6_500_000, 9_500_000),
    "Supervisor": (9_500_000, 14_500_000),
    "Manager": (15_000_000, 28_000_000),
}
BOBOT_DEPT = [10, 12, 22, 14, 14, 28]  # proporsi karyawan per departemen (%)


def buat_karyawan(n, seed=2024, tahun_acuan=2024, dengan_keluar=False):
    """Kembalikan list dict karyawan sintetis yang konsisten antar proyek."""
    rng = random.Random(seed)
    data = []
    urut = {}
    nama_terpakai = set()
    for i in range(n):
        jk = rng.choice(["L", "P"])
        while True:
            depan = rng.choice(NAMA_DEPAN_L if jk == "L" else NAMA_DEPAN_P)
            nama = f"{depan} {rng.choice(NAMA_BELAKANG)}"
            if nama not in nama_terpakai:
                nama_terpakai.add(nama)
                break
        dept = rng.choices(DEPARTEMEN, weights=BOBOT_DEPT)[0]
        lvl_idx = rng.choices([0, 1, 2, 3], weights=[55, 25, 14, 6])[0]
        # satu manager per departemen cukup
        if lvl_idx == 3 and any(d["Departemen"] == dept and d["Level"] == "Manager" for d in data):
            lvl_idx = 2
        level = LEVEL[lvl_idx]
        jabatan = JABATAN[dept][lvl_idx]
        tahun_masuk = rng.randint(2012, tahun_acuan)
        if lvl_idx >= 2:
            tahun_masuk = min(tahun_masuk, tahun_acuan - 3)
        tgl_masuk = dt.date(tahun_masuk, rng.randint(1, 12), rng.randint(1, 28))
        urut[tahun_masuk] = urut.get(tahun_masuk, 0) + 1
        nik = f"{tahun_masuk}{urut[tahun_masuk]:04d}"
        usia_saat_masuk = rng.randint(21, 24) + lvl_idx * rng.randint(3, 6)
        tgl_lahir = dt.date(tahun_masuk - usia_saat_masuk, rng.randint(1, 12), rng.randint(1, 28))
        if tgl_masuk.year == tahun_acuan:
            status = rng.choice(["Kontrak", "Kontrak", "Magang"])
        elif tahun_acuan - tgl_masuk.year <= 2 and lvl_idx == 0:
            status = rng.choice(["Kontrak", "Tetap"])
        else:
            status = "Tetap"
        lo, hi = GAJI_LEVEL[level]
        if status == "Magang":
            gaji = 4_750_000
        else:
            gaji = int(round(rng.uniform(lo, hi) / 50_000)) * 50_000
        pend = rng.choices(PENDIDIKAN, weights=[25, 25, 45, 5])[0] if lvl_idx < 3 else "S1"
        nama_email = nama.lower().replace(" ", ".")
        k = {
            "NIK": nik,
            "Nama": nama,
            "JK": jk,
            "TglLahir": tgl_lahir,
            "Departemen": dept,
            "Jabatan": jabatan,
            "Level": level,
            "Status": status,
            "TglMasuk": tgl_masuk,
            "Gaji": gaji,
            "Email": f"{nama_email}@contohperusahaan.co.id",
            "HP": f"08{rng.randint(11, 99)}{rng.randint(1000, 9999)}{rng.randint(1000, 9999)}",
            "Kota": rng.choice(KOTA),
            "Pendidikan": pend,
            "StatusAktif": "Aktif",
            "TglKeluar": None,
            "AlasanKeluar": None,
        }
        data.append(k)

    if dengan_keluar:
        # sebagian karyawan keluar selama 2023-2024 (turnover)
        alasan = ["Resign", "Habis Kontrak", "Pensiun", "Pindah Kota", "PHK"]
        kandidat = [k for k in data if k["TglMasuk"] < dt.date(2024, 6, 1)]
        rng.shuffle(kandidat)
        for k in kandidat[: int(len(data) * 0.18)]:
            tgl = dt.date(rng.choice([2023, 2023, 2024, 2024]), rng.randint(1, 12), rng.randint(1, 28))
            if tgl <= k["TglMasuk"]:
                tgl = k["TglMasuk"] + dt.timedelta(days=rng.randint(60, 400))
            if tgl > dt.date(2024, 12, 31):
                continue
            k["StatusAktif"] = "Keluar"
            k["TglKeluar"] = tgl
            k["AlasanKeluar"] = rng.choice(alasan)
    data.sort(key=lambda d: d["NIK"])
    return data
