'
' ---------------------------------------------------------------------------
' Script Name: Sony Bulk Report
' Author: Brandon Johnson
' Date: 2026-02-05
' Version: 1.0.0
' ---------------------------------------------------------------------------
' Description:
' Automates separation of data in an Excel sheet to multiple sheets and separate workbooks.
'
' Features:
' - Creates a backup of the original file before processing (*-backup.xlsx)
' - Identifies a sheet with name "MECH"
' - Creates a separate sheet for each artist name from the "Artist" column (tentatively column E)
'   - Each sheet is named the same as the artist (invalid sheet name characters are escaped)
'   - Each sheet contains all column data from every row that matches the artist for that sheet
'   - Original formatting and styling are preserved
'   - Creates separator row (a solid grey fill) to visually group the data by UPC: column title is "Product #" (tentatively column D)
'
' Notes:
' - Workbook must be saved before running (backup uses SaveCopyAs).
' - Optional MaxArtists parameter limits number of artist sheets created for testing.
'
Option Explicit

Public Sub Run_SonyBulkReport()
    Call SonyBulkReport(100)
End Sub

Public Sub SonyBulkReport(Optional ByVal MaxArtists As Long = 0)
    Dim wb As Workbook
    Dim wsMech As Worksheet
    Dim used As Range
    Dim header As Range
    Dim lastRow As Long, lastCol As Long
    Dim artistCol As Long, upcCol As Long
    Dim artists As Collection
    Dim dict As Object
    Dim r As Long
    Dim artist As String
    Dim artistKey As String
    Dim created As Long
    Dim createdSheets As Collection
    Dim outputDir As String
    Dim prevCalc As XlCalculation
    Dim prevScreenUpdating As Boolean
    Dim prevEvents As Boolean
    
    Set wb = ActiveWorkbook
    If wb.Path = "" Then
        MsgBox "Please save the workbook before running (backup requires a file path).", vbExclamation
        Exit Sub
    End If
    If InStr(1, wb.Path, "://", vbTextCompare) > 0 Then
        MsgBox "Workbook path is not local. Please open a local file before running.", vbExclamation
        Exit Sub
    End If
    
    Set wsMech = GetWorksheetByNameCaseInsensitive(wb, "MECH")
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
    artistCol = FindHeaderColumn(header, Array("Artist"))
    upcCol = FindHeaderColumn(header, Array("Product #", "Product#", "Product Number", "UPC"))
    
    ' Fallback to tentative columns if headers aren't found
    If artistCol = 0 And lastCol >= 5 Then artistCol = 5 ' column E
    If upcCol = 0 And lastCol >= 4 Then upcCol = 4 ' column D
    
    If artistCol = 0 Then
        MsgBox "Couldn't find an ""Artist"" column.", vbExclamation
        Exit Sub
    End If
    If upcCol = 0 Then
        MsgBox "Couldn't find a ""Product #"" column.", vbExclamation
        Exit Sub
    End If
    
    ' Performance settings
    prevCalc = Application.Calculation
    prevScreenUpdating = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = True
    
    On Error GoTo CleanUp
    
    ' Create file backup
    CreateBackupFile wb
    
    ' Collect artists in order of appearance
    Set artists = New Collection
    Set dict = CreateObject("Scripting.Dictionary")
    
    For r = 2 To lastRow
        artist = Trim$(CStr(wsMech.Cells(r, artistCol).Value))
        artistKey = NormalizeArtistName(artist)
        If Len(artistKey) = 0 Then artistKey = "Unknown Artist"
        If Not dict.Exists(artistKey) Then
            dict.Add artistKey, True
            artists.Add artistKey
        End If
    Next r
    
    Set createdSheets = New Collection
    
    ' Create sheets per artist
    created = 0
    Dim i As Long
    For i = 1 To artists.Count
        If MaxArtists > 0 And created >= MaxArtists Then Exit For
        artist = CStr(artists(i))
        Dim sheetName As String
        sheetName = CreateArtistSheet(wb, wsMech, artist, artistCol, upcCol, lastRow, lastCol)
        createdSheets.Add sheetName
        created = created + 1
    Next i
    
    ' Create output directory
    outputDir = CreateOutputDirectory(wb)
    
    ' Create workbooks per artist sheet / UPC
    ExportArtistSheetsToWorkbooks wb, createdSheets, outputDir, upcCol
    
    MsgBox "Created " & created & " artist sheet(s) and exported workbooks to:" & vbCrLf & outputDir, vbInformation
    
CleanUp:
    ' Restore application settings
    On Error Resume Next
    Application.StatusBar = False
    On Error GoTo 0
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEvents
    If Err.Number <> 0 Then
        MsgBox "Error: " & Err.Description, vbExclamation
    End If
End Sub

Private Function CreateArtistSheet( _
    ByVal wb As Workbook, _
    ByVal wsMech As Worksheet, _
    ByVal artist As String, _
    ByVal artistCol As Long, _
    ByVal upcCol As Long, _
    ByVal lastRow As Long, _
    ByVal lastCol As Long)
    
    Dim ws As Worksheet
    Dim nameBase As String
    Dim safeName As String
    Dim r As Long, targetRow As Long
    Dim upc As String, lastUpc As String
    
    nameBase = SanitizeSheetName(artist)
    safeName = UniqueSheetName(wb, nameBase)
    
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = safeName
    
    ' Copy header row with formatting
    wsMech.Range(wsMech.Cells(1, 1), wsMech.Cells(1, lastCol)).Copy _
        Destination:=ws.Cells(1, 1)
    
    targetRow = 2
    lastUpc = vbNullString
    
    For r = 2 To lastRow
        If NormalizeArtistName(Trim$(CStr(wsMech.Cells(r, artistCol).Value))) = artist Then
            upc = Trim$(CStr(wsMech.Cells(r, upcCol).Value))
            If targetRow > 2 And upc <> lastUpc Then
                With ws.Range(ws.Cells(targetRow, 1), ws.Cells(targetRow, lastCol))
                    .ClearContents
                    .Interior.Color = RGB(192, 192, 192)
                End With
                targetRow = targetRow + 1
            End If
            
            wsMech.Range(wsMech.Cells(r, 1), wsMech.Cells(r, lastCol)).Copy _
                Destination:=ws.Cells(targetRow, 1)
            targetRow = targetRow + 1
            lastUpc = upc
        End If
    Next r
    
    ws.Columns.AutoFit
    ws.Rows.AutoFit
    
    CreateArtistSheet = ws.Name
End Function

Private Sub ExportArtistSheetsToWorkbooks( _
    ByVal wb As Workbook, _
    ByVal createdSheets As Collection, _
    ByVal outputDir As String, _
    ByVal upcCol As Long)
    
    Dim i As Long
    Dim ws As Worksheet
    Dim upcKeys As Collection
    Dim upcFirstRow As Collection
    Dim r As Long
    Dim lastRow As Long, lastCol As Long
    Dim upc As String
    Dim key As String
    Dim totalUpcs As Long
    Dim processedUpcs As Long
    
    Dim failures As Collection
    Set failures = New Collection
    
    ' First pass: count total UPCs for progress
    totalUpcs = 0
    For i = 1 To createdSheets.Count
        Set ws = wb.Worksheets(CStr(createdSheets(i)))
        If ws Is Nothing Then GoTo NextCountSheet
        Set upcKeys = GetUniqueUpcs(ws, upcCol)
        totalUpcs = totalUpcs + upcKeys.Count
NextCountSheet:
        Set ws = Nothing
    Next i
    
    If totalUpcs = 0 Then
        MsgBox "No UPCs found to export. Check the ""Product #"" column and data.", vbExclamation
        Exit Sub
    End If
    
    ShowProgress totalUpcs
    
    For i = 1 To createdSheets.Count
        Set ws = wb.Worksheets(CStr(createdSheets(i)))
        If ws Is Nothing Then GoTo NextSheet
        
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If lastRow < 2 Then GoTo NextSheet
        
        Set upcKeys = New Collection
        Set upcFirstRow = New Collection
        
        For r = 2 To lastRow
            upc = Trim$(CStr(ws.Cells(r, upcCol).Value))
            If Len(upc) = 0 Then GoTo ContinueRow
            If Not CollectionKeyExists(upcKeys, upc) Then
                upcKeys.Add upc, upc
                upcFirstRow.Add r, upc
            End If
ContinueRow:
        Next r
        
        Dim u As Long
        For u = 1 To upcKeys.Count
            key = CStr(upcKeys(u))
            processedUpcs = processedUpcs + 1
            UpdateProgress processedUpcs, totalUpcs, ws.Name, key
            On Error GoTo ExportFail
            ExportUPCWorkbook ws, key, CLng(upcFirstRow(key)), lastCol, outputDir, upcCol
            On Error GoTo 0
        Next u
        
NextSheet:
        Set ws = Nothing
    Next i
    
    If failures.Count > 0 Then
        Dim msg As String
        msg = "Some exports failed:" & vbCrLf
        For i = 1 To failures.Count
            msg = msg & "- " & failures(i) & vbCrLf
        Next i
        MsgBox msg, vbExclamation
    End If
    Exit Sub
    
ExportFail:
    failures.Add ws.Name & " / UPC " & key & " -> " & Err.Description
    Err.Clear
    Resume Next
End Sub

Private Sub ExportUPCWorkbook( _
    ByVal ws As Worksheet, _
    ByVal upc As String, _
    ByVal firstRow As Long, _
    ByVal lastCol As Long, _
    ByVal outputDir As String, _
    ByVal upcCol As Long)
    
    Dim artistName As String
    Dim productTitle As String
    Dim bulkFeedDate As Variant
    Dim datePart As String
    Dim fileName As String
    Dim filePath As String
    Dim wbOut As Workbook
    Dim wsOut As Worksheet
    Dim r As Long
    Dim lastRow As Long
    Dim rng As Range
    Dim visible As Range
    
    artistName = ws.Name
    productTitle = Trim$(CStr(ws.Cells(firstRow, 6).Value)) ' Column F
    bulkFeedDate = ws.Cells(firstRow, 1).Value ' Column A
    datePart = FormatBulkFeedDate(bulkFeedDate)
    
    fileName = BuildSafeFileName(artistName, productTitle, datePart, outputDir)
    filePath = outputDir & Application.PathSeparator & fileName
    
    Set wbOut = Workbooks.Add(xlWBATWorksheet)
    Set wsOut = wbOut.Worksheets(1)
    wsOut.Name = Left$(SanitizeSheetName(artistName), 31)
    
    ' Filter and copy visible rows (preserves full styling)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    rng.AutoFilter Field:=upcCol, Criteria1:=upc
    
    On Error Resume Next
    Set visible = rng.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If Not visible Is Nothing Then
        SafeCopyPasteAll visible, wsOut.Range("A1")
    End If
    
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    Application.CutCopyMode = False
    
    On Error GoTo SaveFail
    Application.DisplayAlerts = False
    wbOut.SaveAs Filename:=filePath, FileFormat:=xlOpenXMLWorkbook
    wbOut.Close SaveChanges:=False
    Application.DisplayAlerts = True
    
    If Len(Dir(filePath)) = 0 Then
        Err.Raise vbObjectError + 513, "ExportUPCWorkbook", "File not found after save: " & filePath
    End If
    Exit Sub
    
SaveFail:
    Application.DisplayAlerts = True
    On Error Resume Next
    wbOut.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise Err.Number, "ExportUPCWorkbook", Err.Description
End Sub

Private Function CreateOutputDirectory(ByVal wb As Workbook) As String
    Dim folderPath As String
    Dim baseName As String
    Dim dotPos As Long
    Dim outDir As String
    
    folderPath = wb.Path
    dotPos = InStrRev(wb.Name, ".")
    If dotPos > 0 Then
        baseName = Left$(wb.Name, dotPos - 1)
    Else
        baseName = wb.Name
    End If
    
    outDir = folderPath & Application.PathSeparator & baseName & "_BLR"
    If Dir(outDir, vbDirectory) = "" Then
        MkDir outDir
    End If
    CreateOutputDirectory = outDir
End Function

Private Function SanitizeFileName(ByVal raw As String) As String
    Dim s As String
    s = Trim$(raw)
    If Len(s) = 0 Then s = "Unknown"
    s = Replace(s, "/", "_")
    s = Replace(s, "\", "_")
    s = Replace(s, ":", "_")
    s = Replace(s, "?", "_")
    s = Replace(s, "*", "_")
    s = Replace(s, "[", "_")
    s = Replace(s, "]", "_")
    s = Replace(s, """", "_")
    s = Replace(s, "<", "_")
    s = Replace(s, ">", "_")
    s = Replace(s, "|", "_")
    SanitizeFileName = s
End Function

Private Function FormatBulkFeedDate(ByVal v As Variant) As String
    If IsDate(v) Then
        FormatBulkFeedDate = "(" & Format$(CDate(v), "mm.dd.yyyy") & ")"
    Else
        FormatBulkFeedDate = "(" & Format$(Date, "mm.dd.yyyy") & ")"
    End If
End Function

Private Function BuildSafeFileName( _
    ByVal artistName As String, _
    ByVal productTitle As String, _
    ByVal datePart As String, _
    ByVal outputDir As String) As String
    
    Const MaxFullPath As Long = 240 ' conservative Windows-safe max
    Dim artistSafe As String
    Dim titleSafe As String
    Dim suffix As String
    Dim sep As String
    Dim maxNameLen As Long
    
    artistSafe = SanitizeFileName(artistName)
    titleSafe = SanitizeFileName(productTitle)
    suffix = " - Bulk License " & datePart & ".xlsx"
    sep = " - "
    
    maxNameLen = MaxFullPath - (Len(outputDir) + 1)
    If maxNameLen < 20 Then maxNameLen = 20
    
    Dim base As String
    base = artistSafe & sep & titleSafe & suffix
    If Len(base) <= maxNameLen Then
        BuildSafeFileName = base
        Exit Function
    End If
    
    ' First trim product title, then artist name if needed
    Dim available As Long
    available = maxNameLen - (Len(artistSafe) + Len(sep) + Len(suffix))
    If available < 1 Then available = 1
    titleSafe = Left$(titleSafe, available)
    base = artistSafe & sep & titleSafe & suffix
    If Len(base) <= maxNameLen Then
        BuildSafeFileName = base
        Exit Function
    End If
    
    available = maxNameLen - (Len(sep) + Len(titleSafe) + Len(suffix))
    If available < 1 Then available = 1
    artistSafe = Left$(artistSafe, available)
    
    BuildSafeFileName = artistSafe & sep & titleSafe & suffix
End Function

Private Sub SafeCopyPasteAll(ByVal src As Range, ByVal dest As Range)
    Dim i As Long
    Dim lastErr As Long
    Dim lastDesc As String
    
    For i = 1 To 10
        On Error Resume Next
        src.Copy
        If Err.Number <> 0 Then
            lastErr = Err.Number
            lastDesc = Err.Description
            Err.Clear
            GoTo RetryCopy
        End If
        dest.PasteSpecial xlPasteAll
        If Err.Number <> 0 Then
            lastErr = Err.Number
            lastDesc = Err.Description
            Err.Clear
            GoTo RetryCopy
        End If
        dest.PasteSpecial xlPasteColumnWidths
        If Err.Number <> 0 Then
            lastErr = Err.Number
            lastDesc = Err.Description
            Err.Clear
            GoTo RetryCopy
        End If
        On Error GoTo 0
        Exit Sub
RetryCopy:
        On Error GoTo 0
        Application.CutCopyMode = False
        DoEvents
        Application.Wait Now + TimeValue("0:00:01")
    Next i
    
    If lastErr <> 0 Then
        Err.Raise lastErr, "SafeCopyPasteAll", lastDesc
    End If
End Sub

Private Function CollectionKeyExists(ByVal col As Collection, ByVal key As String) As Boolean
    On Error Resume Next
    Dim tmp As Variant
    tmp = col(key)
    CollectionKeyExists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function GetUniqueUpcs(ByVal ws As Worksheet, ByVal upcCol As Long) As Collection
    Dim col As New Collection
    Dim lastRow As Long
    Dim r As Long
    Dim upc As String
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        upc = Trim$(CStr(ws.Cells(r, upcCol).Value))
        If Len(upc) = 0 Then GoTo ContinueRow
        If Not CollectionKeyExists(col, upc) Then col.Add upc, upc
ContinueRow:
    Next r
    Set GetUniqueUpcs = col
End Function

Private Sub ShowProgress(ByVal totalUpcs As Long)
    Application.StatusBar = "Preparing exports... (0 of " & totalUpcs & " UPCs)"
    DoEvents
End Sub

Private Sub UpdateProgress( _
    ByVal currentUpc As Long, _
    ByVal totalUpcs As Long, _
    ByVal artistName As String, _
    ByVal upc As String)
    
    Application.StatusBar = "Exporting " & currentUpc & " of " & totalUpcs & _
        " UPCs | Artist: " & artistName & " | UPC: " & upc
    DoEvents
End Sub

Private Function NormalizeArtistName(ByVal raw As String) As String
    Dim s As String
    Dim l As String
    Dim markers As Variant
    Dim i As Long
    Dim pos As Long
    Dim bestPos As Long
    
    s = Trim$(raw)
    If Len(s) = 0 Then
        NormalizeArtistName = ""
        Exit Function
    End If
    
    l = LCase$(s)
    markers = Array(" feat.", " feat", " featuring", " feature", " ft.", " ft")
    bestPos = 0
    
    For i = LBound(markers) To UBound(markers)
        pos = InStr(1, l, markers(i), vbTextCompare)
        If pos > 0 Then
            If bestPos = 0 Or pos < bestPos Then bestPos = pos
        End If
    Next i
    
    ' Also handle parenthetical forms like "Artist (feat. X)"
    pos = InStr(1, l, "(feat", vbTextCompare)
    If pos > 0 Then
        If bestPos = 0 Or pos < bestPos Then bestPos = pos
    End If
    pos = InStr(1, l, "(featuring", vbTextCompare)
    If pos > 0 Then
        If bestPos = 0 Or pos < bestPos Then bestPos = pos
    End If
    pos = InStr(1, l, "(feature", vbTextCompare)
    If pos > 0 Then
        If bestPos = 0 Or pos < bestPos Then bestPos = pos
    End If
    pos = InStr(1, l, "(ft", vbTextCompare)
    If pos > 0 Then
        If bestPos = 0 Or pos < bestPos Then bestPos = pos
    End If
    
    If bestPos > 0 Then
        s = Left$(s, bestPos - 1)
    End If
    
    ' If multiple artists are listed, keep the first one (split on comma or ampersand)
    Dim cutPos As Long
    cutPos = InStr(1, s, ",", vbTextCompare)
    If cutPos > 0 Then s = Left$(s, cutPos - 1)
    cutPos = InStr(1, s, "&", vbTextCompare)
    If cutPos > 0 Then s = Left$(s, cutPos - 1)
    
    NormalizeArtistName = Trim$(s)
End Function

Private Function FindHeaderColumn(ByVal headerRow As Range, ByVal names As Variant) As Long
    Dim c As Range
    Dim i As Long
    For Each c In headerRow.Cells
        For i = LBound(names) To UBound(names)
            If StrComp(Trim$(CStr(c.Value)), CStr(names(i)), vbTextCompare) = 0 Then
                FindHeaderColumn = c.Column
                Exit Function
            End If
        Next i
    Next c
    FindHeaderColumn = 0
End Function

Private Function GetWorksheetByNameCaseInsensitive(ByVal wb As Workbook, ByVal name As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, name, vbTextCompare) = 0 Then
            Set GetWorksheetByNameCaseInsensitive = ws
            Exit Function
        End If
    Next ws
    Set GetWorksheetByNameCaseInsensitive = Nothing
End Function

Private Function SanitizeSheetName(ByVal raw As String) As String
    Dim s As String
    s = Trim$(raw)
    If Len(s) = 0 Then s = "Unknown Artist"
    s = Replace(s, "/", "_")
    s = Replace(s, "\", "_")
    s = Replace(s, ":", "_")
    s = Replace(s, "?", "_")
    s = Replace(s, "*", "_")
    s = Replace(s, "[", "_")
    s = Replace(s, "]", "_")
    s = Replace(s, Chr$(39), "_") ' apostrophe
    If Len(s) > 31 Then s = Left$(s, 31)
    SanitizeSheetName = s
End Function

Private Function UniqueSheetName(ByVal wb As Workbook, ByVal baseName As String) As String
    Dim nameTry As String
    Dim i As Long
    nameTry = baseName
    If Not WorksheetExists(wb, nameTry) Then
        UniqueSheetName = nameTry
        Exit Function
    End If
    For i = 2 To 999
        nameTry = Left$(baseName, 31 - (Len(CStr(i)) + 3)) & " (" & i & ")"
        If Not WorksheetExists(wb, nameTry) Then
            UniqueSheetName = nameTry
            Exit Function
        End If
    Next i
    UniqueSheetName = Left$(baseName, 28) & "..."
End Function

Private Function WorksheetExists(ByVal wb As Workbook, ByVal name As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(name)
    WorksheetExists = Not ws Is Nothing
    Set ws = Nothing
    On Error GoTo 0
End Function

Private Sub CreateBackupFile(ByVal wb As Workbook)
    Dim folderPath As String
    Dim fileName As String
    Dim baseName As String
    Dim backupPath As String
    Dim dotPos As Long
    
    folderPath = wb.Path
    fileName = wb.Name
    dotPos = InStrRev(fileName, ".")
    If dotPos > 0 Then
        baseName = Left$(fileName, dotPos - 1)
    Else
        baseName = fileName
    End If
    
    backupPath = folderPath & "\" & baseName & "-backup.xlsx"
    wb.SaveCopyAs backupPath
End Sub
