function main(workbook: ExcelScript.Workbook): { baris: number; status: string } {
  const table = workbook.getTable("tblPermintaan");
  if (!table) throw new Error("Tabel tblPermintaan tidak ditemukan.");
  const count = table.getRowCount();
  if (count === 0) return { baris: 0, status: "KOSONG" };
  const headers = table.getHeaderRowRange().getValues()[0].map(value => String(value));
  const required = ["request_id", "qty", "harga_satuan", "status", "nilai"];
  required.forEach(name => {
    if (headers.indexOf(name) < 0) throw new Error("Kolom tidak ditemukan: " + name);
  });
  const rows = table.getRangeBetweenHeaderAndTotal().getValues();
  const ids = new Set<string>();
  rows.forEach((row, index) => {
    const id = String(row[headers.indexOf("request_id")]).trim().toUpperCase();
    const qty = row[headers.indexOf("qty")];
    const price = row[headers.indexOf("harga_satuan")];
    const status = String(row[headers.indexOf("status")]);
    if (!id || ids.has(id)) throw new Error("ID kosong atau duplikat pada baris " + (index + 1));
    ids.add(id);
    if (typeof qty !== "number" || typeof price !== "number" || qty <= 0 || price < 0) {
      throw new Error("Qty atau harga tidak valid pada baris " + (index + 1));
    }
    if (["BARU", "DIPROSES", "SELESAI"].indexOf(status) < 0) {
      throw new Error("Status tidak dikenal pada baris " + (index + 1));
    }
  });
  table.getColumnByName("nilai").getRangeBetweenHeaderAndTotal().setFormula("=[@qty]*[@harga_satuan]");
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  return { baris: count, status: "LULUS" };
}
