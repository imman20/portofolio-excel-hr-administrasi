let
    Csv = Csv.Document(File.Contents(pFolder & "/target_wide.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Csv, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Headers, {{"wilayah", type text}, {"2024-01", Int64.Type}, {"2024-02", Int64.Type}}),
    Long = Table.UnpivotOtherColumns(Typed, {"wilayah"}, "periode", "target"),
    Dated = Table.AddColumn(Long, "awal_bulan", each Date.FromText([periode] & "-01", [Format="yyyy-MM-dd", Culture="en-US"]), type date)
in
    Dated
