let
    // Saya mengganti hanya lokasi folder ini sebelum menjalankan query.
    FolderData = "C:\portofolio\proyek\01_sales_profitability\data\raw",
    Baca = (Nama as text, Tipe as list) as table =>
        Table.TransformColumnTypes(
            Table.PromoteHeaders(Csv.Document(File.Contents(FolderData & "/" & Nama),
                [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]), [PromoteAllScalars=true]), Tipe, "en-US"),
    Unik = (Tabel as table, Kunci as list, Nama as text) as table =>
        if Table.RowCount(Tabel) <> Table.RowCount(Table.Distinct(Tabel, Kunci))
        then error ("Kunci duplikat: " & Nama) else Table.Buffer(Tabel),
    Transaksi = Unik(Baca("transaksi.csv", {{"order_id", type text}, {"tanggal", type date},
        {"customer_id", type text}, {"product_id", type text}, {"region_id", type text},
        {"salesperson_id", type text}, {"qty", Int64.Type}, {"harga_satuan", Currency.Type},
        {"diskon_pct", type number}}), {"order_id"}, "transaksi"),
    Retur = Unik(Baca("retur.csv", {{"retur_id", type text}, {"order_id", type text},
        {"tanggal_retur", type date}, {"qty_retur", Int64.Type}, {"alasan", type text}}), {"retur_id"}, "retur"),
    Produk = Unik(Baca("master_produk.csv", {{"product_id", type text}, {"nama_produk", type text},
        {"kategori", type text}, {"sub_kategori", type text}, {"harga_list", Currency.Type},
        {"hpp_satuan", Currency.Type}}), {"product_id"}, "produk"),
    Pelanggan = Unik(Baca("master_pelanggan.csv", {{"customer_id", type text}, {"nama_pelanggan", type text},
        {"segmen", type text}, {"kota", type text}, {"region_id", type text}, {"tanggal_bergabung", type date}}),
        {"customer_id"}, "pelanggan"),
    Wilayah = Unik(Baca("master_wilayah.csv", {{"region_id", type text}, {"nama_region", type text},
        {"pulau", type text}}), {"region_id"}, "wilayah"),
    Sales = Unik(Baca("master_sales.csv", {{"salesperson_id", type text}, {"nama", type text},
        {"region_id", type text}}), {"salesperson_id"}, "sales"),
    Target = Unik(Baca("target.csv", {{"bulan", type text}, {"region_id", type text},
        {"target_penjualan_bersih", Currency.Type}}), {"bulan", "region_id"}, "target"),
    TargetValid = Table.RowCount(Target)=12*Table.RowCount(Wilayah) and
        List.AllTrue(List.Transform(Table.ToRecords(Target), each
            List.Contains(Wilayah[region_id], _[region_id]) and
            List.Contains(List.Transform({1..12}, (m) => "2024-" & Text.PadStart(Text.From(m),2,"0")), _[bulan]) and
            _[target_penjualan_bersih]>0)),
    ReturJoin = Table.NestedJoin(Retur, {"order_id"}, Transaksi, {"order_id"}, "order", JoinKind.LeftOuter),
    ReturCheck = Table.AddColumn(ReturJoin, "valid", each
        if Table.RowCount([order])<>1 then false else
        [qty_retur]>0 and [tanggal_retur]>=[order]{0}[tanggal] and Date.Year([tanggal_retur])=2024, type logical),
    ReturValid = if List.AllTrue(ReturCheck[valid]) then Retur else error "Relasi/tanggal retur tidak valid",
    ReturAgregat = Table.Group(ReturValid, {"order_id"}, {{"qty_retur", each List.Sum([qty_retur]), Int64.Type}}),
    GabungRetur = Table.NestedJoin(Transaksi, {"order_id"}, ReturAgregat, {"order_id"}, "ret", JoinKind.LeftOuter),
    BukaRetur = Table.ExpandTableColumn(GabungRetur, "ret", {"qty_retur"}),
    ReturNol = Table.ReplaceValue(BukaRetur, null, 0, Replacer.ReplaceValue, {"qty_retur"}),
    GabungProduk = Table.NestedJoin(ReturNol, {"product_id"}, Produk, {"product_id"}, "prod", JoinKind.LeftOuter),
    BukaProduk = Table.ExpandTableColumn(GabungProduk, "prod", {"hpp_satuan", "kategori"}),
    GabungPelanggan = Table.NestedJoin(BukaProduk, {"customer_id"}, Pelanggan, {"customer_id"}, "cust", JoinKind.LeftOuter),
    BukaPelanggan = Table.ExpandTableColumn(GabungPelanggan, "cust", {"segmen", "region_id", "tanggal_bergabung"},
        {"segmen", "region_pelanggan", "tanggal_bergabung"}),
    GabungSales = Table.NestedJoin(BukaPelanggan, {"salesperson_id"}, Sales, {"salesperson_id"}, "sales", JoinKind.LeftOuter),
    BukaSales = Table.ExpandTableColumn(GabungSales, "sales", {"region_id"}, {"region_sales"}),
    Kontrol = Table.AddColumn(BukaSales, "valid", each
        [hpp_satuan]<>null and [hpp_satuan]>0 and [segmen]<>null and
        [region_id]=[region_pelanggan] and [region_id]=[region_sales] and
        List.Contains(Wilayah[region_id], [region_id]) and [tanggal_bergabung]<>null and
        [tanggal]>=[tanggal_bergabung] and Date.Year([tanggal])=2024 and
        [qty]>0 and [harga_satuan]>0 and [diskon_pct]>=0 and [diskon_pct]<=0.25 and
        [qty_retur]>=[qty]*0 and [qty_retur]<=[qty], type logical),
    Validasi = if TargetValid and List.AllTrue(Kontrol[valid]) and
        Table.RowCount(Kontrol)=Table.RowCount(Transaksi)
        then Table.RemoveColumns(Kontrol, {"valid", "region_pelanggan", "region_sales", "tanggal_bergabung"})
        else error "Target, nilai, atau relasi transaksi tidak valid",
    Bruto = Table.AddColumn(Validasi, "omzet_bruto", each [qty]*[harga_satuan], Currency.Type),
    Diskon = Table.AddColumn(Bruto, "nilai_diskon", each [omzet_bruto]*[diskon_pct], Currency.Type),
    Neto = Table.AddColumn(Diskon, "omzet_neto", each [omzet_bruto]-[nilai_diskon], Currency.Type),
    NilaiRetur = Table.AddColumn(Neto, "nilai_retur", each [qty_retur]*[harga_satuan]*(1-[diskon_pct]), Currency.Type),
    Bersih = Table.AddColumn(NilaiRetur, "penjualan_bersih", each [omzet_neto]-[nilai_retur], Currency.Type),
    Hpp = Table.AddColumn(Bersih, "hpp", each ([qty]-[qty_retur])*[hpp_satuan], Currency.Type),
    Laba = Table.AddColumn(Hpp, "laba_kotor", each [penjualan_bersih]-[hpp], Currency.Type),
    Margin = Table.AddColumn(Laba, "margin_kotor", each if [penjualan_bersih]=0 then 0 else [laba_kotor]/[penjualan_bersih], type number),
    Periode = Table.AddColumn(Margin, "periode", each Date.ToText([tanggal], "yyyy-MM", "en-US"), type text),
    Kunci = Table.AddColumn(Periode, "kunci_unik", each [order_id] & "|" & [product_id], type text),
    HitungKunci = Table.Group(Kunci, {"kunci_unik"}, {{"jumlah_kunci", each Table.RowCount(_), Int64.Type}}),
    JoinKunci = Table.NestedJoin(Kunci, {"kunci_unik"}, HitungKunci, {"kunci_unik"}, "hitungan", JoinKind.LeftOuter),
    BukaKunci = Table.ExpandTableColumn(JoinKunci, "hitungan", {"jumlah_kunci"}),
    Duplikat = Table.AddColumn(BukaKunci, "duplikat", each if [jumlah_kunci]>1 then 1 else 0, Int64.Type),
    Hasil = Table.SelectColumns(Duplikat, {"order_id", "tanggal", "customer_id", "product_id", "region_id", "salesperson_id",
        "qty", "harga_satuan", "diskon_pct", "omzet_bruto", "nilai_diskon", "omzet_neto", "qty_retur", "nilai_retur",
        "penjualan_bersih", "hpp_satuan", "hpp", "laba_kotor", "margin_kotor", "periode", "kategori", "segmen", "kunci_unik", "duplikat"})
in
    Hasil
