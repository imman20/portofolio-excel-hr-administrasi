"""Saya membangkitkan simulasi persediaan dengan seed tetap dan buku stok harian."""
from pathlib import Path
import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
SEED = 202403


def main():
    rng = np.random.default_rng(SEED)
    raw, processed = ROOT/'data/raw', ROOT/'data/processed'
    raw.mkdir(parents=True, exist_ok=True)
    processed.mkdir(parents=True, exist_ok=True)
    dates = pd.date_range('2024-01-01', '2024-12-31')
    suppliers = pd.DataFrame([{'supplier_id':f'SUP{i:02}', 'nama_supplier':f'Pemasok Simulasi {i:02}', 'lead_time_nominal':7+i} for i in range(1,9)])
    products, orders, movements, daily = [], [], [], []
    for s in range(60):
        sku = f'SKU{s+1:03}'
        supplier = suppliers.iloc[s%8]
        scale = float(rng.uniform(.5, 2.2))
        cost = int(rng.integers(8,201))*1000
        initial = int(rng.integers(60,160)*scale)
        products.append(dict(sku=sku, nama_sku=f'Barang Simulasi {s+1:03}', supplier_id=supplier.supplier_id, kategori=['ATK','Kemasan','Peralatan'][s%3], hpp_satuan=cost, stok_awal=initial, tanggal_stok_awal='2024-01-01'))
        supply = [0.72,0.95,1.1,1.35,2.0][s%5]
        receipts = {}
        for k in range(8):
            arrival = pd.Timestamp('2024-01-18')+pd.Timedelta(days=k*47+int(rng.integers(-7,8)))
            lead = int(rng.integers(5,21))
            placed = arrival-pd.Timedelta(days=lead)
            if placed < dates[0]: placed = dates[0]
            qty = int(round(366*.42*16*scale*supply/8))
            po_id = f'PO{len(orders)+1:04}'
            orders.append(dict(po_id=po_id, sku=sku, supplier_id=supplier.supplier_id, tanggal_pesan=placed.strftime('%Y-%m-%d'), tanggal_terima=arrival.strftime('%Y-%m-%d'), qty_pesan=qty, qty_terima=qty, status='SELESAI'))
            receipts[arrival] = (qty, po_id)
        if s%3==0:
            orders.append(dict(po_id=f'PO{len(orders)+1:04}', sku=sku, supplier_id=supplier.supplier_id, tanggal_pesan='2024-12-27', tanggal_terima='', qty_pesan=int(80*scale), qty_terima=0, status='TERBUKA'))
        stock = initial
        adjust_days = set(rng.choice(np.arange(366), 2, replace=False).tolist())
        for di, date in enumerate(dates):
            opening = stock
            incoming = outgoing = adjustment = 0
            def add(kind, qty, po=''):
                movements.append(dict(movement_id=f'MOV{len(movements)+1:05}', tanggal=date.strftime('%Y-%m-%d'), sku=sku, jenis=kind, qty=qty, po_id=po))
            if date in receipts:
                incoming, po = receipts[date]
                stock += incoming
                add('MASUK', incoming, po)
            if di in adjust_days:
                adjustment = int(rng.integers(-6,7)) or 1
                adjustment = max(-stock, adjustment)
                if adjustment:
                    stock += adjustment
                    add('PENYESUAIAN', adjustment)
            if rng.random()<.42:
                requested = max(1,int(round(rng.normal(16,5)*scale)))
                outgoing = min(stock,requested)
                if outgoing:
                    stock -= outgoing
                    add('KELUAR',outgoing)
            daily.append(dict(tanggal=date.strftime('%Y-%m-%d'),sku=sku,stok_awal_hari=opening,masuk=incoming,keluar=outgoing,penyesuaian=adjustment,stok_akhir=stock,stockout=int(stock==0)))
    for name, records in [('master_sku',products),('purchase_order',orders),('pergerakan',movements)]:
        frame = pd.DataFrame(records)
        if name=='pergerakan': frame=frame.sort_values(['tanggal','movement_id'],kind='stable')
        frame.to_csv(raw/f'{name}.csv',index=False)
    suppliers.to_csv(raw/'supplier.csv',index=False)
    pd.DataFrame(daily).sort_values(['tanggal','sku']).to_csv(processed/'stok_harian.csv',index=False)
    print({'seed':SEED,'sku':len(products),'supplier':len(suppliers),'pergerakan':len(movements),'po':len(orders),'stok_harian':len(daily)})


if __name__=='__main__':
    main()
