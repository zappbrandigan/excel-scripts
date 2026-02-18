Option Explicit

Public Sub Run_SBR_ExportBulkFiles()
    Call SBR_ExportBulkFiles
End Sub

Public Sub SBR_ExportBulkFiles()
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
    Dim groupRows As Object
    Dim groupArtists As Object
    Dim groupTitles As Object
    Dim groupKeys As Collection
    Dim outputDir As String
    Dim prevCalc As XlCalculation
    Dim prevScreenUpdating As Boolean
    Dim prevEvents As Boolean
    Dim processed As Long
    Dim totalGroups As Long
    Dim r As Long
    Dim artist As String
    Dim rawTitle As String
    Dim canonicalTitle As String
    Dim groupKey As String
    Dim failures As Collection
    Dim i As Long
    Dim firstRow As Long
    Dim exportArtist As String
    Dim msg As String
    Dim rowBucket As Collection
    Dim artistBucket As Object

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
    Set groupRows = CreateObject("Scripting.Dictionary")
    Set groupArtists = CreateObject("Scripting.Dictionary")
    Set groupTitles = CreateObject("Scripting.Dictionary")
    Set groupKeys = New Collection

    For r = 2 To lastRow
        rawTitle = Trim$(CStr(wsMech.Cells(r, titleCol).Value))
        If Len(rawTitle) = 0 Then GoTo ContinueCollect

        canonicalTitle = SBR_CanonicalProductTitle(rawTitle)
        If Len(canonicalTitle) = 0 Then canonicalTitle = rawTitle
        groupKey = LCase$(canonicalTitle)

        If Not groupFirstRow.Exists(groupKey) Then
            groupFirstRow.Add groupKey, r
            groupTitles.Add groupKey, canonicalTitle

            Set rowBucket = New Collection
            rowBucket.Add r
            groupRows.Add groupKey, rowBucket

            Set artistBucket = CreateObject("Scripting.Dictionary")
            groupArtists.Add groupKey, artistBucket
            groupKeys.Add groupKey
        Else
            Set rowBucket = groupRows(groupKey)
            rowBucket.Add r
        End If

        artist = SBR_NormalizeArtistName(Trim$(CStr(wsMech.Cells(r, artistCol).Value)))
        If Len(artist) = 0 Then artist = "Unknown Artist"
        Set artistBucket = groupArtists(groupKey)
        If Not artistBucket.Exists(artist) Then artistBucket.Add artist, True
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
        firstRow = CLng(groupFirstRow(groupKey))
        canonicalTitle = CStr(groupTitles(groupKey))
        Set rowBucket = groupRows(groupKey)
        Set artistBucket = groupArtists(groupKey)
        exportArtist = SBR_ResolveExportArtistName(artistBucket)

        processed = processed + 1
        SBR_UpdateProgress processed, totalGroups, exportArtist, canonicalTitle

        On Error GoTo ExportFail
        SBR_ExportGroupWorkbook wsMech, exportArtist, canonicalTitle, rowBucket, firstRow, lastCol, outputDir, upcCol
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
    failures.Add exportArtist & " / " & canonicalTitle & " -> " & Err.Description
    Err.Clear
    Resume ContinueExport
End Sub

Private Function SBR_ResolveExportArtistName(ByVal artistBucket As Object) As String
    Dim key As Variant

    If artistBucket Is Nothing Then
        SBR_ResolveExportArtistName = "Unknown Artist"
        Exit Function
    End If

    If artistBucket.Count = 0 Then
        SBR_ResolveExportArtistName = "Unknown Artist"
    ElseIf artistBucket.Count = 1 Then
        For Each key In artistBucket.Keys
            SBR_ResolveExportArtistName = CStr(key)
            Exit Function
        Next key
    Else
        SBR_ResolveExportArtistName = "Various"
    End If
End Function

Private Sub SBR_ExportGroupWorkbook( _
    ByVal wsMech As Worksheet, _
    ByVal artistName As String, _
    ByVal productTitle As String, _
    ByVal rowBucket As Collection, _
    ByVal firstRow As Long, _
    ByVal lastCol As Long, _
    ByVal outputDir As String, _
    ByVal upcCol As Long)

    Dim bulkFeedDate As Variant
    Dim datePart As String
    Dim fileName As String
    Dim filePath As String
    Dim wbOut As Workbook
    Dim wsOut As Worksheet
    Dim targetRow As Long
    Dim upc As String
    Dim nextUpc As String
    Dim lastUpc As String
    Dim srcBlock As Range
    Dim idx As Long
    Dim rowNum As Long
    Dim runStartRow As Long
    Dim runEndRow As Long
    Dim nextRowNum As Long
    Dim upcOrder As Collection
    Dim upcRows As Object
    Dim upcKey As Variant
    Dim upcBucket As Collection

    bulkFeedDate = wsMech.Cells(firstRow, 1).Value
    datePart = SBR_FormatBulkFeedDate(bulkFeedDate)

    fileName = SBR_BuildSafeFileName(artistName, productTitle, datePart, outputDir)
    filePath = outputDir & Application.PathSeparator & fileName

    Set wbOut = Workbooks.Add(xlWBATWorksheet)
    Set wsOut = wbOut.Worksheets(1)
    wsOut.Name = Left$(SBR_SanitizeSheetName(artistName), 31)

    SBR_SafePasteFormatsThenValues wsMech.Range(wsMech.Cells(1, 1), wsMech.Cells(1, lastCol)), wsOut.Range("A1")

    targetRow = 2
    lastUpc = vbNullString

    Set upcOrder = New Collection
    Set upcRows = CreateObject("Scripting.Dictionary")
    For idx = 1 To rowBucket.Count
        rowNum = CLng(rowBucket(idx))
        upc = Trim$(CStr(wsMech.Cells(rowNum, upcCol).Value))
        If Len(upc) > 0 Then
            If Not upcRows.Exists(upc) Then
                Set upcBucket = New Collection
                upcRows.Add upc, upcBucket
                upcOrder.Add upc
            End If
            Set upcBucket = upcRows(upc)
            upcBucket.Add rowNum
        End If
    Next idx

    For Each upcKey In upcOrder
        upc = CStr(upcKey)
        Set upcBucket = upcRows(upc)

        If targetRow > 2 And upc <> lastUpc Then
            With wsOut.Range(wsOut.Cells(targetRow, 1), wsOut.Cells(targetRow, lastCol))
                .ClearContents
                .Interior.Color = RGB(192, 192, 192)
            End With
            targetRow = targetRow + 1
        End If

        idx = 1
        Do While idx <= upcBucket.Count
            rowNum = CLng(upcBucket(idx))
            runStartRow = rowNum
            runEndRow = rowNum

            Do While idx < upcBucket.Count
                nextRowNum = CLng(upcBucket(idx + 1))
                nextUpc = Trim$(CStr(wsMech.Cells(nextRowNum, upcCol).Value))

                If nextRowNum <> runEndRow + 1 Then Exit Do
                If nextUpc <> upc Then Exit Do

                runEndRow = nextRowNum
                idx = idx + 1
            Loop

            Set srcBlock = wsMech.Range(wsMech.Cells(runStartRow, 1), wsMech.Cells(runEndRow, lastCol))
            SBR_SafePasteFormatsThenValues srcBlock, wsOut.Cells(targetRow, 1)
            targetRow = targetRow + (runEndRow - runStartRow + 1)
            idx = idx + 1
        Loop

        lastUpc = upc
    Next upcKey

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
