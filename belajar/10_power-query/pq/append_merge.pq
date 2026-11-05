let
    Files = Folder.Files(pFolder),
    Selected = Table.SelectRows(Files, each Text.StartsWith([Name], "pesanan_") and [Extension] = ".csv"),
    Tables = Table.AddColumn(Selected, "Tabel", each
        let fileName = [Name],
            csv = Csv.Document([Content], [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
            headers = Table.PromoteHeaders(csv, [PromoteAllScalars=true])
        in Table.AddColumn(headers, "sumber_file", each fileName, type text)),
    Appended = Table.Combine(Tables[Tabel]),
    Typed = Table.TransformColumnTypes(Appended, {{"order_id", type text}, {"kode", type text}, {"qty", Int64.Type}, {"sumber_file", type text}}),
    Dates = Table.TransformColumns(Typed, {{"tanggal", each Date.FromText(_, [Format="yyyy-MM-dd", Culture="en-US"]), type date}}),
    MasterCsv = Csv.Document(File.Contents(pFolder & "/master_produk.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    MasterHeaders = Table.PromoteHeaders(MasterCsv, [PromoteAllScalars=true]),
    Master = Table.TransformColumnTypes(MasterHeaders, {{"kode", type text}, {"kategori", type text}, {"harga", Int64.Type}}),
    KeyCounts = Table.Group(Master, {"kode"}, {{"n", each Table.RowCount(_), Int64.Type}}),
    CheckedMaster = if Table.RowCount(Table.SelectRows(KeyCounts, each [n] <> 1)) > 0
                    then error "Kunci master tidak unik; saya memeriksa duplikat sebelum Merge." else Master,
    Joined = Table.NestedJoin(Dates, {"kode"}, CheckedMaster, {"kode"}, "master", JoinKind.LeftOuter),
    Expanded = Table.ExpandTableColumn(Joined, "master", {"kategori", "harga"}, {"kategori", "harga"}),
    CheckedRows = if Table.RowCount(Expanded) <> Table.RowCount(Dates)
                  then error "Jumlah baris berubah setelah Merge." else Expanded,
    CheckedKeys = if Table.RowCount(Table.SelectRows(CheckedRows, each [harga] = null)) > 0
                  then error "Ada kode yang tidak ditemukan pada master." else CheckedRows,
    Result = Table.AddColumn(CheckedKeys, "nilai", each [qty] * [harga], Int64.Type)
in
    Result
