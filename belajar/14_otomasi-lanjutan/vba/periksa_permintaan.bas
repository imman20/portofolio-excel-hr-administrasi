Attribute VB_Name = "PeriksaPermintaan"
Option Explicit

Public Sub PeriksaDanHitung()
    Dim tabel As ListObject
    Dim data As Variant
    Dim seen As Object
    Dim i As Long, id As String, status As String
    Dim kolId As Long, kolQty As Long, kolHarga As Long, kolStatus As Long
    On Error GoTo Gagal
    Set tabel = ThisWorkbook.Worksheets("Data").ListObjects("tblPermintaan")
    If tabel.DataBodyRange Is Nothing Then
        MsgBox "Tidak ada catatan untuk diperiksa.", vbInformation
        Exit Sub
    End If
    kolId = tabel.ListColumns("request_id").Index
    kolQty = tabel.ListColumns("qty").Index
    kolHarga = tabel.ListColumns("harga_satuan").Index
    kolStatus = tabel.ListColumns("status").Index
    data = tabel.DataBodyRange.Value2
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    For i = 1 To UBound(data, 1)
        id = Trim$(CStr(data(i, kolId)))
        If Len(id) = 0 Then Err.Raise vbObjectError + 100, , "ID kosong pada baris " & i
        If seen.Exists(id) Then Err.Raise vbObjectError + 101, , "ID duplikat: " & id
        seen.Add id, True
        If Not IsNumeric(data(i, kolQty)) Or Not IsNumeric(data(i, kolHarga)) Then _
            Err.Raise vbObjectError + 102, , "Qty atau harga bukan angka pada baris " & i
        If CDbl(data(i, kolQty)) <= 0 Or CDbl(data(i, kolHarga)) < 0 Then _
            Err.Raise vbObjectError + 103, , "Qty atau harga di luar batas pada baris " & i
        status = CStr(data(i, kolStatus))
        If status <> "BARU" And status <> "DIPROSES" And status <> "SELESAI" Then _
            Err.Raise vbObjectError + 104, , "Status tidak dikenal pada baris " & i
    Next i
    tabel.ListColumns("nilai").DataBodyRange.Formula = "=[@qty]*[@harga_satuan]"
    Application.CalculateFull
    MsgBox "Pemeriksaan selesai untuk " & UBound(data, 1) & " catatan. Saya tidak mengubah status atau menghapus sumber.", vbInformation
    Exit Sub
Gagal:
    MsgBox "Pemeriksaan dihentikan: " & Err.Description, vbExclamation
End Sub
