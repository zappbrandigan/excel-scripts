Option Explicit

Public Sub Run_SBR_ExportBulkFiles()
    Call SBR_ExportBulkFiles
End Sub

Public Sub SBR_ExportBulkFiles()
    Const KEY_SEP As String = "||~||"

    Dim wb As Workbook
    Dim wsMech As Worksheet
    Dim used As Range
    Dim header As Range
    Dim lastRow As Long
    Dim lastCol As Long
    Dim artistCol As Long
    Dim upcCol As Long
    Dim titleCol As Long
    Dim groupFirstRow As Object
    Dim groupKeys As Collection
    Dim outputDir As String
    Dim prevCalc As XlCalculation
    Dim prevScreenUpdating As Boolean
    Dim prevEvents As Boolean
    Dim processed As Long
    Dim totalGroups As Long
    Dim r As Long
    Dim artist As String
    Dim title As String
    Dim groupKey As String
    Dim failures As Collection
    Dim i As Long
    Dim sepPos As Long
    Dim firstRow As Long
    Dim msg As String

    Set wb = ActiveWorkbook
    If wb.Path = "" Then
        MsgBox "Please save the workbook before running.", vbExclamation
        Exit Sub
    End If
    If InStr(1, wb.Path, "://", vbTextCompare) > 0 Then
        MsgBox "Workbook path is not local. Please open a local file before running.", vbExclamation
        Exit Sub
    End If

    Set wsMech = SBR_GetWorksheetByNameCaseInsensitive(wb, "MECH")
    If wsMech Is Nothing Then
        MsgBox "No worksheet named ""MECH"" was found.", vbExclamation
        Exit Sub
    End If

    Set used = wsMech.UsedRange
    If used Is Nothing Then
        MsgBox "No data found on the MECH worksheet.", vbExclamation
        Exit Sub
    End If

    lastRow = used.Row + used.Rows.Count - 1
    lastCol = used.Column + used.Columns.Count - 1
    If lastRow < 2 Then
        MsgBox "Need at least one data row beneath the header row.", vbExclamation
        Exit Sub
    End If

    Set header = wsMech.Range(wsMech.Cells(1, 1), wsMech.Cells(1, lastCol))
    artistCol = SBR_FindHeaderColumn(header, Array("Artist"))
    upcCol = SBR_FindHeaderColumn(header, Array("Product #", "Product#", "Product Number", "UPC"))
    titleCol = SBR_FindHeaderColumn(header, Array("Product Title", "Title"))

    If artistCol = 0 And lastCol >= 5 Then artistCol = 5
    If upcCol = 0 And lastCol >= 4 Then upcCol = 4
    If titleCol = 0 And lastCol >= 6 Then titleCol = 6

    If artistCol = 0 Then
        MsgBox "Couldn't find an ""Artist"" column.", vbExclamation
        Exit Sub
    End If
    If upcCol = 0 Then
        MsgBox "Couldn't find a ""Product #"" column.", vbExclamation
        Exit Sub
    End If
    If titleCol = 0 Then
        MsgBox "Couldn't find a ""Product Title"" column.", vbExclamation
        Exit Sub
    End If

    prevCalc = Application.Calculation
    prevScreenUpdating = Application.ScreenUpdating
    prevEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = True

    On Error GoTo CleanUp

    Set groupFirstRow = CreateObject("Scripting.Dictionary")
    Set groupKeys = New Collection

    For r = 2 To lastRow
        artist = SBR_NormalizeArtistName(Trim$(CStr(wsMech.Cells(r, artistCol).Value)))
        If Len(artist) = 0 Then artist = "Unknown Artist"

        title = Trim$(CStr(wsMech.Cells(r, titleCol).Value))
        If Len(title) = 0 Then GoTo ContinueCollect

        groupKey = artist & KEY_SEP & title
        If Not groupFirstRow.Exists(groupKey) Then
            groupFirstRow.Add groupKey, r
            groupKeys.Add groupKey
        End If
ContinueCollect:
    Next r

    totalGroups = groupKeys.Count
    If totalGroups = 0 Then
        MsgBox "No Product Titles found to export. Check the ""Product Title"" column and data.", vbExclamation
        GoTo CleanUp
    End If

    outputDir = SBR_CreateOutputDirectory(wb)
    SBR_ShowProgress totalGroups

    Set failures = New Collection

    For i = 1 To groupKeys.Count
        groupKey = CStr(groupKeys(i))
        sepPos = InStr(1, groupKey, KEY_SEP, vbBinaryCompare)
        If sepPos = 0 Then GoTo ContinueExport

        artist = Left$(groupKey, sepPos - 1)
        title = Mid$(groupKey, sepPos + Len(KEY_SEP))
        firstRow = CLng(groupFirstRow(groupKey))

        processed = processed + 1
        SBR_UpdateProgress processed, totalGroups, artist, title

        On Error GoTo ExportFail
        SBR_ExportGroupWorkbook wsMech, artist, title, firstRow, lastRow, lastCol, outputDir, artistCol, titleCol, upcCol
        On Error GoTo CleanUp
ContinueExport:
    Next i

    If failures.Count > 0 Then
        msg = "Some exports failed:" & vbCrLf
        For i = 1 To failures.Count
            msg = msg & "- " & failures(i) & vbCrLf
        Next i
        MsgBox msg, vbExclamation
    Else
        MsgBox "Exported " & processed & " file(s) to:" & vbCrLf & outputDir, vbInformation
    End If

CleanUp:
    On Error Resume Next
    Application.StatusBar = False
    On Error GoTo 0

    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEvents

    If Err.Number <> 0 Then
        MsgBox "Error: " & Err.Description, vbExclamation
    End If
    Exit Sub

ExportFail:
    failures.Add artist & " / " & title & " -> " & Err.Description
    Err.Clear
    Resume ContinueExport
End Sub

Private Sub SBR_ExportGroupWorkbook( _
    ByVal wsMech As Worksheet, _
    ByVal artistName As String, _
    ByVal productTitle As String, _
    ByVal firstRow As Long, _
    ByVal lastRow As Long, _
    ByVal lastCol As Long, _
    ByVal outputDir As String, _
    ByVal artistCol As Long, _
    ByVal titleCol As Long, _
    ByVal upcCol As Long)

    Dim bulkFeedDate As Variant
    Dim datePart As String
    Dim fileName As String
    Dim filePath As String
    Dim wbOut As Workbook
    Dim wsOut As Worksheet
    Dim r As Long
    Dim targetRow As Long
    Dim upc As String
    Dim lastUpc As String
    Dim srcRow As Range

    bulkFeedDate = wsMech.Cells(firstRow, 1).Value
    datePart = SBR_FormatBulkFeedDate(bulkFeedDate)

    fileName = SBR_BuildSafeFileName(artistName, productTitle, datePart, outputDir)
    filePath = outputDir & Application.PathSeparator & fileName

    Set wbOut = Workbooks.Add(xlWBATWorksheet)
    Set wsOut = wbOut.Worksheets(1)
    wsOut.Name = Left$(SBR_SanitizeSheetName(artistName), 31)

    SBR_SafeCopyPasteAll wsMech.Range(wsMech.Cells(1, 1), wsMech.Cells(1, lastCol)), wsOut.Range("A1")

    targetRow = 2
    lastUpc = vbNullString

    For r = 2 To lastRow
        If SBR_NormalizeArtistName(Trim$(CStr(wsMech.Cells(r, artistCol).Value))) = artistName Then
            If Trim$(CStr(wsMech.Cells(r, titleCol).Value)) = productTitle Then
                upc = Trim$(CStr(wsMech.Cells(r, upcCol).Value))
                If Len(upc) > 0 Then
                    If targetRow > 2 And upc <> lastUpc Then
                        With wsOut.Range(wsOut.Cells(targetRow, 1), wsOut.Cells(targetRow, lastCol))
                            .ClearContents
                            .Interior.Color = RGB(192, 192, 192)
                        End With
                        targetRow = targetRow + 1
                    End If

                    Set srcRow = wsMech.Range(wsMech.Cells(r, 1), wsMech.Cells(r, lastCol))
                    SBR_SafeCopyPasteAll srcRow, wsOut.Cells(targetRow, 1)
                    targetRow = targetRow + 1
                    lastUpc = upc
                End If
            End If
        End If
    Next r

    On Error GoTo SaveFail
    Application.DisplayAlerts = False
    wbOut.SaveAs Filename:=filePath, FileFormat:=xlOpenXMLWorkbook
    wbOut.Close SaveChanges:=False
    Application.DisplayAlerts = True

    If Len(Dir(filePath)) = 0 Then
        Err.Raise vbObjectError + 513, "SBR_ExportGroupWorkbook", "File not found after save: " & filePath
    End If
    Exit Sub

SaveFail:
    Application.DisplayAlerts = True
    On Error Resume Next
    wbOut.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise Err.Number, "SBR_ExportGroupWorkbook", Err.Description
End Sub
