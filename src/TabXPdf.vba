Sub PrepareAndExportSheetsToPDF()
    Dim ws As Worksheet
    Dim pdfPath As String
    Dim safeName As String
    
    pdfPath = ThisWorkbook.Path & "\"
    
    For Each ws In ThisWorkbook.Worksheets
        ' Activate the sheet
        ws.Activate
        
        ' Delete column I (9th column)
        On Error Resume Next ' Ignore if column I doesn’t exist
        ws.Columns("I").Delete
        On Error GoTo 0
        
        ' Set width for columns H and I
        ws.Columns("H").ColumnWidth = 25
        ws.Columns("I").ColumnWidth = 25
        
        ' Prepare a filename-safe version of the sheet name
        safeName = ws.Name
        safeName = Replace(safeName, "/", "_")
        safeName = Replace(safeName, "\", "_")
        safeName = Replace(safeName, ":", "_")
        safeName = Replace(safeName, "?", "_")
        safeName = Replace(safeName, "*", "_")
        safeName = Replace(safeName, "[", "_")
        safeName = Replace(safeName, "]", "_")
        
        ' Export as PDF
        ws.ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=pdfPath & safeName & ".pdf", _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
    Next ws
    
    MsgBox "All sheets exported as PDFs to: " & pdfPath, vbInformation
End Sub
