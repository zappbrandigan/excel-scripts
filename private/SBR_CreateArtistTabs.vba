Option Explicit

Public Sub Run_SBR_CreateArtistTabs()
    Call SBR_CreateArtistTabs(1000)
End Sub

Public Sub SBR_CreateArtistTabs(Optional ByVal MaxArtists As Long = 0)
    Dim wb As Workbook
    Dim wsMech As Worksheet
    Dim used As Range
    Dim header As Range
    Dim lastRow As Long
    Dim lastCol As Long
    Dim artistCol As Long
    Dim upcCol As Long
    Dim artists As Collection
    Dim dict As Object
    Dim r As Long
    Dim artist As String
    Dim artistKey As String
    Dim created As Long
    Dim i As Long
    Dim prevCalc As XlCalculation
    Dim prevScreenUpdating As Boolean
    Dim prevEvents As Boolean

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

    If artistCol = 0 And lastCol >= 5 Then artistCol = 5
    If upcCol = 0 And lastCol >= 4 Then upcCol = 4

    If artistCol = 0 Then
        MsgBox "Couldn't find an ""Artist"" column.", vbExclamation
        Exit Sub
    End If
    If upcCol = 0 Then
        MsgBox "Couldn't find a ""Product #"" column.", vbExclamation
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

    Set artists = New Collection
    Set dict = CreateObject("Scripting.Dictionary")

    For r = 2 To lastRow
        artist = Trim$(CStr(wsMech.Cells(r, artistCol).Value))
        artistKey = SBR_NormalizeArtistName(artist)
        If Len(artistKey) = 0 Then artistKey = "Unknown Artist"

        If Not dict.Exists(artistKey) Then
            dict.Add artistKey, True
            artists.Add artistKey
        End If
    Next r

    created = 0
    For i = 1 To artists.Count
        If MaxArtists > 0 And created >= MaxArtists Then Exit For

        artist = CStr(artists(i))
        Application.StatusBar = "Creating sheet " & (created + 1) & " of " & _
            IIf(MaxArtists > 0 And MaxArtists < artists.Count, MaxArtists, artists.Count) & _
            " | Artist: " & artist
        DoEvents

        Call SBR_CreateSingleArtistSheet(wb, wsMech, artist, artistCol, upcCol, lastRow, lastCol)
        created = created + 1
    Next i

    MsgBox "Created " & created & " artist sheet(s).", vbInformation

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
End Sub

Private Sub SBR_CreateSingleArtistSheet( _
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
    Dim r As Long
    Dim targetRow As Long
    Dim upc As String
    Dim lastUpc As String

    nameBase = SBR_SanitizeSheetName(artist)
    safeName = SBR_UniqueSheetName(wb, nameBase)

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = safeName

    wsMech.Range(wsMech.Cells(1, 1), wsMech.Cells(1, lastCol)).Copy Destination:=ws.Cells(1, 1)

    targetRow = 2
    lastUpc = vbNullString

    For r = 2 To lastRow
        If SBR_NormalizeArtistName(Trim$(CStr(wsMech.Cells(r, artistCol).Value))) = artist Then
            upc = Trim$(CStr(wsMech.Cells(r, upcCol).Value))
            If targetRow > 2 And upc <> lastUpc Then
                With ws.Range(ws.Cells(targetRow, 1), ws.Cells(targetRow, lastCol))
                    .ClearContents
                    .Interior.Color = RGB(192, 192, 192)
                End With
                targetRow = targetRow + 1
            End If

            wsMech.Range(wsMech.Cells(r, 1), wsMech.Cells(r, lastCol)).Copy Destination:=ws.Cells(targetRow, 1)
            targetRow = targetRow + 1
            lastUpc = upc
        End If
    Next r

    ws.Columns.AutoFit
    ws.Rows.AutoFit
End Sub
