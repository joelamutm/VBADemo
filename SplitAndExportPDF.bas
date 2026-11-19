Attribute VB_Name = "SplitAndExportPDF"
'===========================================================
' Project   : ICT1002 Lab7
' Module    : vba1_breakDown() and vba2_exportPDF
' Author    : Joe Lam
' Created   : 24-Nov-2025
' Purpose   : For reviewing the functionality of VBA through two vba examples
' Version   : 2
' Change Log:
'   1.0 - Initial version with two procedures
'   2.0 - Improved code efficiency in procedure vba1_breakDown()
'===========================================================

Option Explicit

'-----------------------------------------------------------
' Procedure : vba1_breakDown()
' Author    : Joe Lam
' Created   : 24-Nov-2025
' Purpose   : Copy each record from the list of 100 students and paste each record individually to a newly created blank worksheet.
'-----------------------------------------------------------

Sub vba1_breakDown()
    Dim ncol As Long
    Dim I As Long
    Dim startTime As Double, endTime As Double
    Dim wsResult As Worksheet, wsNew As Worksheet
    
    startTime = Timer
    
    ' Reference the source sheet
    Set wsResult = Worksheets("Result")
    
    ' Count non-empty cells in column A
    ncol = WorksheetFunction.CountA(wsResult.Columns(1))
    
    ' Speed optimizations
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Loop through rows
    For I = 1 To ncol - 1
        ' Add new sheet
        Set wsNew = Sheets.Add(After:=Sheets(Sheets.Count))
        wsNew.Name = CStr(I)
        
        ' Copy header row (A1:H1)
        wsNew.Range("A1:H1").Value = wsResult.Range("A1:H1").Value
        wsResult.Range("A1:H1").Copy
        wsNew.Range("A1").PasteSpecial Paste:=xlPasteColumnWidths
        
        ' Copy data row (A2:H2, A3:H3, etc.)
        wsNew.Range("A2:H2").Value = wsResult.Range("A" & I + 1 & ":H" & I + 1).Value
        wsResult.Range("A" & I + 1 & ":H" & I + 1).Copy
        wsNew.Range("A2").PasteSpecial Paste:=xlPasteColumnWidths
    Next I
    
    ' Restore settings
    Application.CutCopyMode = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    wsResult.Activate
    wsResult.Range("A1").Select
    
    endTime = Timer
    
    MsgBox "All worksheets have been successfully created." & vbCrLf & _
           "Time taken: " & Format(endTime - startTime, "0.00") & " seconds"
End Sub


'-----------------------------------------------------------
' Procedure : vba2_exportPDF()
' Author    : Joe Lam
' Created   : 24-Nov-2025
' Purpose   : Exports each worksheet in the workbook as a separate PDF file into a folder on Desktop.
'-----------------------------------------------------------

Sub vba2_exportPDF()
    Dim FolderPath As String
    Dim I As Integer
    Dim startTime As Double, endTime As Double

    startTime = Timer

    FolderPath = Environ("USERPROFILE") & "\Desktop\PDFOutput\"
    MkDir FolderPath

    For I = 1 To Worksheets.Count
        Worksheets(I).ExportAsFixedFormat Type:=xlTypePDF, FileName:=FolderPath & "\" & _
        Worksheets(I).Name, openafterpublish:=False
    Next

    endTime = Timer

    MsgBox "All PDF's have been successfully exported." & vbCrLf & _
           "Time taken: " & Format(endTime - startTime, "0.00") & " seconds"
End Sub

