"""Audit CSV dan pembanding statistik independen untuk rekonsiliasi saya."""
from pathlib import Path
import json
import numpy as np
import pandas as pd
from scipy import stats

BASE = Path(__file__).resolve().parents[1]
ITEMS = [f"q{i:02}" for i in range(1, 13)]
DIMENSIONS = {"Produk": ITEMS[:3], "Layanan": ITEMS[3:6], "Harga": ITEMS[6:9], "Pengiriman": ITEMS[9:]}


def counts(raw):
    score = raw[ITEMS + ["kepuasan_keseluruhan", "nps"]]
    lik = score.iloc[:, :13].apply(lambda c: pd.to_numeric(c.str.strip(), errors="coerce"))
    age = pd.to_numeric(raw.usia, errors="coerce")
    return {
        "sel_kosong_skor": int(score.eq("").sum().sum()),
        "teks_na": int(score.eq("N/A").sum().sum()),
        "teks_strip": int(score.eq("-").sum().sum()),
        "angka_berspasi": int(score.map(lambda x: x != x.strip()).sum().sum()),
        **{f"luar_rentang_{v}": int(lik.eq(v).sum().sum()) for v in [0, 6, 7]},
        "usia_outlier": int((~age.between(17, 80)).sum()),
        "tanggal_format_dmy": int(raw.tanggal_survei.str.fullmatch(r"\d{2}/\d{2}/\d{4}").sum()),
        "kanal_kapitalisasi": int((~raw.kanal.isin(["Toko", "Online", "Call Center"])).sum()),
    }


def clean(raw):
    d = raw.drop_duplicates("respondent_id", keep="first").copy().reset_index(drop=True)
    for c in ITEMS + ["kepuasan_keseluruhan", "nps", "usia", "lama_pelanggan_bulan"]:
        v = pd.to_numeric(d[c].str.strip(), errors="coerce")
        lo, hi = (0, 10) if c == "nps" else (17, 80) if c == "usia" else (1, 120) if c == "lama_pelanggan_bulan" else (1, 5)
        d[c] = v.where(v.between(lo, hi) & v.eq(v.round()))
    d["tanggal_survei"] = d.tanggal_survei.map(lambda x: pd.to_datetime(x, format="%Y-%m-%d" if "-" in x else "%d/%m/%Y").strftime("%Y-%m-%d"))
    d["kanal"] = d.kanal.str.strip().str.lower().map({"toko": "Toko", "online": "Online", "call center": "Call Center"})
    for c in ["kota", "segmen", "jenis_kelamin"]: d[c] = d[c].str.strip()
    for name, cols in DIMENSIONS.items():
        d[f"dim_{name.lower()}"] = d[cols].mean(axis=1).where(d[cols].notna().all(axis=1))
    return d


def audit(raw, notes):
    assert len(raw) == 420 and raw.respondent_id.nunique() == 414
    assert raw.respondent_id.str.fullmatch(r"R-\d{4}").all()
    unique = raw.drop_duplicates("respondent_id")
    # Duplicates must be exact copies, not conflicting survey submissions.
    assert raw.drop_duplicates().shape[0] == unique.shape[0]
    before, after = counts(unique), counts(raw)
    for k, v in before.items(): assert notes[k] == v, (k, notes[k], v)
    assert notes["total_baris_raw"] == len(raw)
    ids = sorted(raw.loc[raw.respondent_id.duplicated(), "respondent_id"].tolist())
    assert notes["duplikat_id"] == ids and notes["baris_duplikat"] == len(ids) == 6
    if "hitungan_csv_akhir" in notes: assert notes["hitungan_csv_akhir"] == after
    return {"sebelum_duplikasi": before, "csv_akhir": after, "duplikat_id": ids, "total_baris_raw": len(raw), "n_bersih": len(unique)}


def statistics(d):
    dims = {}
    for name, cols in DIMENSIONS.items():
        complete = d[cols].dropna()
        x = complete.mean(axis=1)
        n, mean, sd = len(x), float(x.mean()), float(x.std(ddof=1))
        margin = float(stats.t.ppf(.975, n-1) * sd / np.sqrt(n))
        alpha = float(1.5 * (1 - complete.var(ddof=1).sum() / complete.sum(axis=1).var(ddof=1)))
        dims[name] = dict(n=n, mean=mean, sd=sd, margin=margin, ci_low=mean-margin, ci_high=mean+margin, alpha=alpha)
    a = d.loc[d.kanal.eq("Toko"), "kepuasan_keseluruhan"].dropna()
    b = d.loc[d.kanal.eq("Online"), "kepuasan_keseluruhan"].dropna()
    test = stats.ttest_ind(a, b, equal_var=False)
    nps = d.nps.dropna()
    prom, det = int(nps.ge(9).sum()), int(nps.le(6).sum())
    cols = [f"dim_{n.lower()}" for n in DIMENSIONS] + ["kepuasan_keseluruhan"]
    cc = d[cols].dropna()
    X = np.column_stack([np.ones(len(cc)), cc.iloc[:, :4]])
    y = cc.iloc[:, 4].to_numpy()
    beta = np.linalg.lstsq(X, y, rcond=None)[0]
    residual = y-X@beta
    dfres = len(y)-5
    mse = float(residual@residual/dfres)
    se = np.sqrt(np.diag(mse*np.linalg.inv(X.T@X)))
    r2 = float(1 - residual@residual / ((y-y.mean())@(y-y.mean())))
    kk = d.kepuasan_keseluruhan.dropna()
    kk_m = float(stats.t.ppf(.975,len(kk)-1)*kk.std(ddof=1)/np.sqrt(len(kk)))
    return {
        "n_bersih": len(d), "dimensions": dims,
        "overall": dict(n=len(kk), mean=float(kk.mean()), sd=float(kk.std(ddof=1)), margin=kk_m, ci_low=float(kk.mean()-kk_m), ci_high=float(kk.mean()+kk_m)),
        "nps": dict(n=len(nps), promotor=prom, detraktor=det, pasif=len(nps)-prom-det, value=(prom-det)/len(nps)*100),
        "welch": dict(n_toko=len(a), n_online=len(b), mean_toko=float(a.mean()), mean_online=float(b.mean()), diff=float(a.mean()-b.mean()), t=float(test.statistic), df=float(test.df), p=float(test.pvalue), cohen_d=float((a.mean()-b.mean())/np.sqrt(((len(a)-1)*a.var(ddof=1)+(len(b)-1)*b.var(ddof=1))/(len(a)+len(b)-2)))),
        "regression": dict(n=len(cc), beta=beta.tolist(), se=se.tolist(), r2=r2, sey=np.sqrt(mse), dfres=dfres, F=r2/(1-r2)*dfres/4),
        "correlation": {name:float(stats.pearsonr(cc[f"dim_{name.lower()}"],cc.kepuasan_keseluruhan).statistic) for name in DIMENSIONS},
    }


def inputs():
    raw = pd.read_csv(BASE/"data/raw/survei_raw.csv", dtype=str, keep_default_na=False)
    notes = json.loads((BASE/"data/processed/catatan_masalah_disengaja.json").read_text())
    report = audit(raw, notes)
    d = clean(raw)
    ref = pd.read_csv(BASE/"data/processed/survei_bersih_referensi.csv", keep_default_na=False)
    for c in d:
        if pd.api.types.is_numeric_dtype(d[c]): ref[c] = pd.to_numeric(ref[c], errors="coerce")
    pd.testing.assert_frame_equal(d, ref, check_dtype=False, check_exact=False, rtol=0, atol=1e-12)
    return raw, d, report, statistics(d)
