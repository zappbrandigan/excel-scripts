Option Explicit

Public Function SBR_FindHeaderColumn(ByVal headerRow As Range, ByVal names As Variant) As Long
    Dim c As Range
    Dim i As Long
    For Each c In headerRow.Cells
        For i = LBound(names) To UBound(names)
            If StrComp(Trim$(CStr(c.Value)), CStr(names(i)), vbTextCompare) = 0 Then
                SBR_FindHeaderColumn = c.Column
                Exit Function
            End If
        Next i
    Next c
    SBR_FindHeaderColumn = 0
End Function

Public Function SBR_GetWorksheetByNameCaseInsensitive(ByVal wb As Workbook, ByVal name As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, name, vbTextCompare) = 0 Then
            Set SBR_GetWorksheetByNameCaseInsensitive = ws
            Exit Function
        End If
    Next ws
    Set SBR_GetWorksheetByNameCaseInsensitive = Nothing
End Function

Public Function SBR_NormalizeArtistName(ByVal raw As String) As String
    Dim s As String
    Dim l As String
    Dim markers As Variant
    Dim i As Long
    Dim pos As Long
    Dim bestPos As Long

    s = Trim$(raw)
    If Len(s) = 0 Then
        SBR_NormalizeArtistName = ""
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

    Dim cutPos As Long
    cutPos = InStr(1, s, ",", vbTextCompare)
    If cutPos > 0 Then s = Left$(s, cutPos - 1)
    cutPos = InStr(1, s, "&", vbTextCompare)
    If cutPos > 0 Then s = Left$(s, cutPos - 1)

    SBR_NormalizeArtistName = Trim$(s)
End Function

Public Function SBR_SanitizeSheetName(ByVal raw As String) As String
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
    s = Replace(s, Chr$(39), "_")
    If Len(s) > 31 Then s = Left$(s, 31)
    SBR_SanitizeSheetName = s
End Function

Public Function SBR_UniqueSheetName(ByVal wb As Workbook, ByVal baseName As String) As String
    Dim nameTry As String
    Dim i As Long

    nameTry = baseName
    If Not SBR_WorksheetExists(wb, nameTry) Then
        SBR_UniqueSheetName = nameTry
        Exit Function
    End If

    For i = 2 To 999
        nameTry = Left$(baseName, 31 - (Len(CStr(i)) + 3)) & " (" & i & ")"
        If Not SBR_WorksheetExists(wb, nameTry) Then
            SBR_UniqueSheetName = nameTry
            Exit Function
        End If
    Next i

    SBR_UniqueSheetName = Left$(baseName, 28) & "..."
End Function

Public Function SBR_WorksheetExists(ByVal wb As Workbook, ByVal name As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(name)
    SBR_WorksheetExists = Not ws Is Nothing
    Set ws = Nothing
    On Error GoTo 0
End Function

Public Function SBR_CreateOutputDirectory(ByVal wb As Workbook) As String
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

    SBR_CreateOutputDirectory = outDir
End Function

Public Function SBR_SanitizeFileName(ByVal raw As String) As String
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
    s = Replace(s, Chr$(34), "_")
    s = Replace(s, "<", "_")
    s = Replace(s, ">", "_")
    s = Replace(s, "|", "_")
    SBR_SanitizeFileName = s
End Function

Public Function SBR_FormatBulkFeedDate(ByVal v As Variant) As String
    If IsDate(v) Then
        SBR_FormatBulkFeedDate = "(" & Format$(CDate(v), "mm.dd.yyyy") & ")"
    Else
        SBR_FormatBulkFeedDate = "(" & Format$(Date, "mm.dd.yyyy") & ")"
    End If
End Function

Public Function SBR_BuildSafeFileName( _
    ByVal artistName As String, _
    ByVal productTitle As String, _
    ByVal datePart As String, _
    ByVal outputDir As String) As String

    Const MaxFullPath As Long = 240
    Dim artistSafe As String
    Dim titleSafe As String
    Dim suffix As String
    Dim sep As String
    Dim maxNameLen As Long
    Dim base As String
    Dim available As Long

    artistSafe = SBR_SanitizeFileName(artistName)
    titleSafe = SBR_SanitizeFileName(productTitle)
    suffix = " - Bulk License " & datePart & ".xlsx"
    sep = " - "

    maxNameLen = MaxFullPath - (Len(outputDir) + 1)
    If maxNameLen < 20 Then maxNameLen = 20

    base = artistSafe & sep & titleSafe & suffix
    If Len(base) <= maxNameLen Then
        SBR_BuildSafeFileName = base
        Exit Function
    End If

    available = maxNameLen - (Len(artistSafe) + Len(sep) + Len(suffix))
    If available < 1 Then available = 1
    titleSafe = Left$(titleSafe, available)
    base = artistSafe & sep & titleSafe & suffix
    If Len(base) <= maxNameLen Then
        SBR_BuildSafeFileName = base
        Exit Function
    End If

    available = maxNameLen - (Len(sep) + Len(titleSafe) + Len(suffix))
    If available < 1 Then available = 1
    artistSafe = Left$(artistSafe, available)

    SBR_BuildSafeFileName = artistSafe & sep & titleSafe & suffix
End Function

Public Sub SBR_SafeCopyPasteAll(ByVal src As Range, ByVal dest As Range)
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
        Err.Raise lastErr, "SBR_SafeCopyPasteAll", lastDesc
    End If
End Sub

Public Function SBR_CollectionKeyExists(ByVal col As Collection, ByVal key As String) As Boolean
    On Error Resume Next
    Dim tmp As Variant
    tmp = col(key)
    SBR_CollectionKeyExists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Public Function SBR_GetUniqueKeys(ByVal ws As Worksheet, ByVal keyCol As Long) As Collection
    Dim col As New Collection
    Dim lastRow As Long
    Dim r As Long
    Dim key As String

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        key = Trim$(CStr(ws.Cells(r, keyCol).Value))
        If Len(key) = 0 Then GoTo ContinueRow
        If Not SBR_CollectionKeyExists(col, key) Then col.Add key, key
ContinueRow:
    Next r

    Set SBR_GetUniqueKeys = col
End Function

Public Sub SBR_ShowProgress(ByVal totalTitles As Long)
    Application.StatusBar = "Preparing exports... (0 of " & totalTitles & " Titles)"
    DoEvents
End Sub

Public Sub SBR_UpdateProgress( _
    ByVal currentTitle As Long, _
    ByVal totalTitles As Long, _
    ByVal artistName As String, _
    ByVal title As String)

    Application.StatusBar = "Exporting " & currentTitle & " of " & totalTitles & _
        " Titles | Artist: " & artistName & " | Title: " & title
    DoEvents
End Sub
