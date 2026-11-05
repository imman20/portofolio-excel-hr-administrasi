Attribute VB_Name = "RekapPerDepartemen"
' Makro: menyalin baris karyawan tiap departemen dari sheet "Data Karyawan"
' ke sheet baru bernama departemennya (dibuat ulang bila sudah ada), lengkap dengan header dan subtotal gaji.
Option Explicit

Public Sub RekapPerDepartemen()
    Dim ws As Worksheet, wsBaru As Worksheet
    Dim barisAkhir As Long, r As Long, tujuan As Long
    Dim dept As Variant, daftar As Object

    Set ws = ThisWorkbook.Worksheets("Data Karyawan")
    barisAkhir = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Set daftar = CreateObject("Scripting.Dictionary")
    For r = 5 To barisAkhir
        If ws.Cells(r, "C").Value <> "" Then daftar(CStr(ws.Cells(r, "C").Value)) = 1
    Next r

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    For Each dept In daftar.Keys
        On Error Resume Next
        ThisWorkbook.Worksheets(CStr(dept)).Delete
        On Error GoTo 0
        Set wsBaru = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsBaru.Name = CStr(dept)
        ws.Range("A4:H4").Copy wsBaru.Range("A1")
        tujuan = 2
        For r = 5 To barisAkhir
            If ws.Cells(r, "C").Value = dept Then
                ws.Range("A" & r & ":H" & r).Copy wsBaru.Range("A" & tujuan)
                tujuan = tujuan + 1
            End If
        Next r
        wsBaru.Cells(tujuan, 1).Value = "Total gaji"
        wsBaru.Cells(tujuan, 7).Formula = "=SUM(G2:G" & (tujuan - 1) & ")"
        wsBaru.Cells(tujuan, 2).Formula = "=COUNTA(A2:A" & (tujuan - 1) & ")&"" orang"""
        wsBaru.Rows(tujuan).Font.Bold = True
        wsBaru.Columns("A:H").AutoFit
    Next dept
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Rekap selesai: " & daftar.Count & " sheet departemen dibuat.", vbInformation
End Sub
