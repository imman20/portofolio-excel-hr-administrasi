Attribute VB_Name = "SlipGajiPDF"
' Makro: membuat slip gaji PDF massal dari workbook proyek 04 (04_penggajian_dan_slip_gaji.xlsx).
' Cara pakai: impor modul ini ke workbook proyek 04, simpan sebagai .xlsm, lalu jalankan BuatSlipGajiPDF.
' Makro mengganti nama di sel C6 sheet "Slip Gaji" satu per satu dan mengekspor area B2:F32 ke PDF.
Option Explicit

Public Sub BuatSlipGajiPDF()
    Dim wsSlip As Worksheet, wsMaster As Worksheet
    Dim barisAkhir As Long, r As Long, jumlah As Long
    Dim folder As String, namaFile As String

    Set wsSlip = ThisWorkbook.Worksheets("Slip Gaji")
    Set wsMaster = ThisWorkbook.Worksheets("Master Karyawan")
    barisAkhir = wsMaster.Cells(wsMaster.Rows.Count, "A").End(xlUp).Row

    folder = ThisWorkbook.Path & Application.PathSeparator & "slip_pdf"
    If Dir(folder, vbDirectory) = "" Then MkDir folder

    Application.ScreenUpdating = False
    For r = 5 To barisAkhir
        If wsMaster.Cells(r, "B").Value <> "" Then
            wsSlip.Range("C6").Value = wsMaster.Cells(r, "B").Value
            Application.Calculate
            namaFile = folder & Application.PathSeparator & "Slip_" & wsMaster.Cells(r, "A").Value & "_" & _
                       Replace(wsMaster.Cells(r, "B").Value, " ", "_") & ".pdf"
            wsSlip.Range("B2:F32").ExportAsFixedFormat Type:=xlTypePDF, Filename:=namaFile, _
                Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=True, OpenAfterPublish:=False
            jumlah = jumlah + 1
        End If
    Next r
    Application.ScreenUpdating = True
    MsgBox jumlah & " slip gaji PDF tersimpan di folder: " & folder, vbInformation
End Sub
