Attribute VB_Name = "ValidasiData"
' Makro: memeriksa kualitas data pada sheet "Data Karyawan"
' - NIK ganda, email kosong, gaji di bawah UMK (sel B126)
' Sel bermasalah disorot merah dan ringkasan ditulis ke sheet "Hasil Validasi".
Option Explicit

Public Sub ValidasiDataKaryawan()
    Dim ws As Worksheet, wsHasil As Worksheet
    Dim barisAkhir As Long, r As Long, n As Long
    Dim umk As Double
    Dim dict As Object

    Set ws = ThisWorkbook.Worksheets("Data Karyawan")
    barisAkhir = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    umk = ws.Range("B126").Value
    Set dict = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    Set wsHasil = ThisWorkbook.Worksheets("Hasil Validasi")
    On Error GoTo 0
    If wsHasil Is Nothing Then
        Set wsHasil = ThisWorkbook.Worksheets.Add(After:=ws)
        wsHasil.Name = "Hasil Validasi"
    Else
        wsHasil.Cells.Clear
    End If
    wsHasil.Range("A1:D1").Value = Array("Baris", "NIK", "Jenis Masalah", "Nilai")
    wsHasil.Range("A1:D1").Font.Bold = True

    ws.Range("A5:H" & barisAkhir).Interior.ColorIndex = xlNone
    n = 1
    For r = 5 To barisAkhir
        If ws.Cells(r, "A").Value = "" Then GoTo Lanjut
        ' NIK ganda
        If dict.Exists(CStr(ws.Cells(r, "A").Value)) Then
            Call Catat(wsHasil, n, r, ws, "NIK ganda", ws.Cells(r, "A").Value): ws.Cells(r, "A").Interior.Color = RGB(255, 199, 206)
        Else
            dict.Add CStr(ws.Cells(r, "A").Value), r
        End If
        ' email kosong
        If Trim(ws.Cells(r, "H").Value) = "" Then
            Call Catat(wsHasil, n, r, ws, "Email kosong", ""): ws.Cells(r, "H").Interior.Color = RGB(255, 199, 206)
        End If
        ' gaji di bawah UMK
        If IsNumeric(ws.Cells(r, "G").Value) Then
            If ws.Cells(r, "G").Value < umk Then
                Call Catat(wsHasil, n, r, ws, "Gaji di bawah UMK", ws.Cells(r, "G").Value): ws.Cells(r, "G").Interior.Color = RGB(255, 199, 206)
            End If
        End If
Lanjut:
    Next r
    wsHasil.Columns("A:D").AutoFit
    MsgBox "Validasi selesai. Ditemukan " & (n - 1) & " masalah. Lihat sheet Hasil Validasi.", vbInformation
End Sub

Private Sub Catat(wsHasil As Worksheet, ByRef n As Long, r As Long, ws As Worksheet, jenis As String, nilai As Variant)
    n = n + 1
    wsHasil.Cells(n, 1).Value = r
    wsHasil.Cells(n, 2).Value = ws.Cells(r, "A").Value
    wsHasil.Cells(n, 3).Value = jenis
    wsHasil.Cells(n, 4).Value = nilai
End Sub
