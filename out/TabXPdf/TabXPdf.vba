'
' ---------------------------------------------------------------------------
' Script Name: Tab Export to PDF
' Author: Brandon Johnson
' Date: 2026-01-11
' Version: 1.0.0
' ---------------------------------------------------------------------------
' Description:
' Exports each worksheet in the active workbook to a PDF file after applying
' basic formatting to columns H and I. Output files are saved alongside the
' workbook.
'
' Features:
' - Deletes column I when present.
' - Sets column widths and auto-fits columns H:I and row heights.
' - Sanitizes worksheet names for safe PDF filenames.
' - Exports every sheet as a separate PDF to the workbook folder.
'
' Notes:
' - Workbook must be saved before running to determine the output folder.
' - Column formatting is applied to each worksheet before export.
'
Sub PrepareAndExportSheetsToPDF()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim pdfPath As String
    Dim safeName As String
    
    ' Reference the workbook that’s currently active
    Set wb = ActiveWorkbook
    
    ' Use that workbook’s folder for output
    pdfPath = wb.Path & "\"
    
    ' If the workbook hasn’t been saved yet
    If pdfPath = "\" Then
        MsgBox "Please save your workbook first before exporting.", vbExclamation
        Exit Sub
    End If
    
    For Each ws In wb.Worksheets
        ws.Activate
        
        ' Delete column I (9th column)
        On Error Resume Next
        ws.Columns("I").Delete
        On Error GoTo 0
        
        ' Set width for columns H and I
        ws.Columns("H").ColumnWidth = 25
        ws.Columns("I").ColumnWidth = 25
        
        ' Auto-fit for content and row heights
        ws.Columns("H:I").AutoFit
        ws.Rows.AutoFit
        
        ' Clean up file name
        safeName = ws.Name
        safeName = Replace(safeName, "/", "_")
        safeName = Replace(safeName, "\", "_")
        safeName = Replace(safeName, ":", "_")
        safeName = Replace(safeName, "?", "_")
        safeName = Replace(safeName, "*", "_")
        safeName = Replace(safeName, "[", "_")
        safeName = Replace(safeName, "]", "_")
        
        ' Export as PDF to the same folder as the workbook
        ws.ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=pdfPath & safeName & ".pdf", _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
    Next ws
    
    MsgBox "All sheets exported as PDFs to: " & pdfPath, vbInformation
End Sub
