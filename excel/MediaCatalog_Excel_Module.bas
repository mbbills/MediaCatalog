Attribute VB_Name = "MediaCatalogExcel"
Option Explicit

' MediaCatalog Excel 2016 module v0.4.1
'
' Standard columns:
'   A  UPC
'   B  Blu-ray.com URL
'   C  Blu-ray.com Title
'   D  IMDb URL
'   E  IMDb ID
'   F  IMDb Title
'   G  Year
'   H  Runtime
'   I  Title Type
'   J  Season
'   K  Status / Error
'   L  Studio
'   M  Blu-ray Year
'   N  Blu-ray Runtime
'   O  Content Rating
'   P  Physical Release Date
'   Q  Disc Format
'   R  Video Codec
'   S  Resolution
'   T  Aspect Ratio
'   U  Disc Count / Capacities
'
' IMDb ID is the single authoritative correction field.

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
#End If

Private Const MENU_TAG As String = "MediaCatalogExcelMenu"
Private Const MENU_CAPTION As String = "Media Catalog"
Private Const HEADER_ROW As Long = 1


' ============================================================================
' Menu lifecycle
' ============================================================================

Public Sub InstallMediaCatalogMenu()
    Dim menuBar As Object
    Dim menuControl As Object

    RemoveMediaCatalogMenu

    Set menuBar = Application.CommandBars("Worksheet Menu Bar")
    Set menuControl = menuBar.Controls.Add(10, , , , True)

    menuControl.Caption = MENU_CAPTION
    menuControl.Tag = MENU_TAG

    AddMenuButton menuControl, "Resolve Selected Rows", "ResolveSelectedRows"
    AddMenuButton menuControl, "Resolve Selected UPCs with BRdC", "ResolveSelectedUPCsWithBRdC"
    AddMenuButton menuControl, "Resolve Selected UPCs with UPCdb", "ResolveSelectedUPCsWithUPCdb"
    AddMenuButton menuControl, "Resolve Selected UPCs with BarcodeLookup.com", "ResolveSelectedUPCsWithBarcodedCom"
    AddMenuButton menuControl, "Open UPC on BarcodeLookup.com (No API)", "ResolveSelectedUPCsWithBarcodedComNoAPI"
    AddMenuButton menuControl, "Enrich Selected Blu-ray Details", "EnrichSelectedBluRayDetails"
    AddMenuButton menuControl, "Lookup IMDb for Selected Rows", "LookupIMDbForCurrentRow"
    AddMenuButton menuControl, "Remove UPC-E Rows in Selection", "RemoveSelectedUPCERows"
    AddMenuButton menuControl, "Open Selected UPC on Blu-ray.com", "OpenSelectedUPCOnBluRay"
    AddMenuButton menuControl, "Check Configuration", "CheckMediaCatalogConfiguration"
End Sub


Public Sub RemoveMediaCatalogMenu()
    Dim menuBar As Object
    Dim control As Object

    On Error Resume Next
    Set menuBar = Application.CommandBars("Worksheet Menu Bar")

    For Each control In menuBar.Controls
        If control.Tag = MENU_TAG Then control.Delete
    Next control

    On Error GoTo 0
End Sub


Private Sub AddMenuButton(ByVal parentControl As Object, ByVal caption As String, ByVal macroName As String)
    Dim button As Object

    Set button = parentControl.Controls.Add(1, , , , True)
    button.Caption = caption
    button.OnAction = "'" & ThisWorkbook.Name & "'!" & macroName
End Sub


' ============================================================================
' Configuration and file helpers
' ============================================================================

Private Function ProjectPath() As String
    ProjectPath = ThisWorkbook.Path
End Function


Private Function JoinPath(ByVal parentPath As String, ByVal childName As String) As String
    If Right$(parentPath, 1) = "\" Then
        JoinPath = parentPath & childName
    Else
        JoinPath = parentPath & "\" & childName
    End If
End Function


Private Function FileExists(ByVal filePath As String) As Boolean
    FileExists = (Len(Dir$(filePath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) > 0)
End Function


Private Function ReadUtf8Text(ByVal filePath As String) As String
    Dim stream As Object

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    ReadUtf8Text = stream.ReadText
    stream.Close

    If Len(ReadUtf8Text) > 0 Then
        If AscW(Left$(ReadUtf8Text, 1)) = &HFEFF Then
            ReadUtf8Text = Mid$(ReadUtf8Text, 2)
        End If
    End If
End Function


Private Sub WriteUtf8Text(ByVal filePath As String, ByVal contents As String)
    Dim stream As Object

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText contents
    stream.SaveToFile filePath, 2
    stream.Close
End Sub


Private Function ReadIniValue( _
    ByVal iniPath As String, _
    ByVal sectionName As String, _
    ByVal keyName As String, _
    ByVal defaultValue As String _
) As String
    Dim lines As Variant
    Dim line As Variant
    Dim currentSection As String
    Dim trimmed As String
    Dim equalsPosition As Long
    Dim thisKey As String

    ReadIniValue = defaultValue
    If Not FileExists(iniPath) Then Exit Function

    lines = Split(Replace(ReadUtf8Text(iniPath), vbCrLf, vbLf), vbLf)

    For Each line In lines
        trimmed = Trim$(CStr(line))

        If Len(trimmed) > 0 Then
            If Left$(trimmed, 1) <> "#" And Left$(trimmed, 1) <> ";" Then
                If Left$(trimmed, 1) = "[" And Right$(trimmed, 1) = "]" Then
                    currentSection = LCase$(Trim$(Mid$(trimmed, 2, Len(trimmed) - 2)))
                ElseIf currentSection = LCase$(sectionName) Then
                    equalsPosition = InStr(1, trimmed, "=", vbBinaryCompare)
                    If equalsPosition > 0 Then
                        thisKey = LCase$(Trim$(Left$(trimmed, equalsPosition - 1)))
                        If thisKey = LCase$(keyName) Then
                            ReadIniValue = Trim$(Mid$(trimmed, equalsPosition + 1))
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next line
End Function


Private Function QuoteArgument(ByVal value As String) As String
    QuoteArgument = Chr$(34) & Replace(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function


Private Function TemporaryPath(ByVal extension As String) As String
    Dim token As String

    Randomize
    token = Format$(Now, "yyyymmdd_hhnnss") & "_" & Format$(CLng(Rnd() * 1000000), "000000")
    TemporaryPath = JoinPath(Environ$("TEMP"), "MediaCatalog_" & token & extension)
End Function


Private Function RunCommandAndWait( _
    ByVal commandLine As String, _
    ByVal timeoutSeconds As Long, _
    ByRef standardOutput As String, _
    ByRef standardError As String _
) As Long
    Dim shell As Object
    Dim process As Object
    Dim started As Date

    Set shell = CreateObject("WScript.Shell")
    Set process = shell.Exec(commandLine)
    started = Now

    Do While process.Status = 0
        DoEvents
        Sleep 100

        If DateDiff("s", started, Now) >= timeoutSeconds Then
            process.Terminate
            standardError = "Timed out after " & CStr(timeoutSeconds) & " seconds."
            RunCommandAndWait = -1
            Exit Function
        End If
    Loop

    standardOutput = process.StdOut.ReadAll
    standardError = process.StdErr.ReadAll
    RunCommandAndWait = process.ExitCode
End Function


Private Sub DeleteTemporaryFile(ByVal filePath As String)
    On Error Resume Next
    If Len(filePath) > 0 And FileExists(filePath) Then Kill filePath
    On Error GoTo 0
End Sub


Private Function GetPythonCommand(ByRef errorText As String) As String
    Dim rootPath As String
    Dim settingsPath As String

    rootPath = ProjectPath()

    If Len(rootPath) = 0 Then
        errorText = "Save the add-in before running MediaCatalog."
        Exit Function
    End If

    settingsPath = JoinPath(rootPath, "settings.ini")

    If Not FileExists(settingsPath) Then
        errorText = "settings.ini was not found:" & vbCrLf & settingsPath
        Exit Function
    End If

    GetPythonCommand = ReadIniValue(settingsPath, "runtime", "python", "")

    If Len(GetPythonCommand) = 0 Then
        errorText = "No Python executable is configured in settings.ini."
    End If
End Function


Private Function GetWindowlessPythonCommand( _
    ByVal pythonCommand As String, _
    ByRef errorText As String _
) As String
    Dim lowerCommand As String
    Dim candidate As String

    pythonCommand = Trim$(pythonCommand)
    lowerCommand = LCase$(pythonCommand)

    ' Bare command names are resolved through PATH by WScript.Shell.
    Select Case lowerCommand
        Case "python", "python.exe"
            GetWindowlessPythonCommand = IIf(lowerCommand = "python", "pythonw", "pythonw.exe")
            Exit Function
        Case "py", "py.exe"
            GetWindowlessPythonCommand = IIf(lowerCommand = "py", "pyw", "pyw.exe")
            Exit Function
        Case "pythonw", "pythonw.exe", "pyw", "pyw.exe"
            GetWindowlessPythonCommand = pythonCommand
            Exit Function
    End Select

    If Right$(lowerCommand, 10) = "python.exe" Then
        candidate = Left$(pythonCommand, Len(pythonCommand) - 10) & "pythonw.exe"
    ElseIf Right$(lowerCommand, 6) = "py.exe" Then
        candidate = Left$(pythonCommand, Len(pythonCommand) - 6) & "pyw.exe"
    ElseIf Right$(lowerCommand, 11) = "pythonw.exe" Or _
           Right$(lowerCommand, 7) = "pyw.exe" Then
        candidate = pythonCommand
    Else
        errorText = "The configured Python command cannot be converted to " & _
                    "a windowless command. Configure the full path to python.exe " & _
                    "in settings.ini."
        Exit Function
    End If

    If Not FileExists(Replace(candidate, "/", "\")) Then
        errorText = "The windowless Python executable was not found:" & vbCrLf & candidate
        Exit Function
    End If

    GetWindowlessPythonCommand = candidate
End Function


Private Sub SetCellHyperlink( _
    ByVal targetCell As Range, _
    ByVal address As String, _
    ByVal displayText As String _
)
    Dim targetSheet As Worksheet

    Set targetSheet = targetCell.Parent

    On Error Resume Next
    targetCell.Hyperlinks.Delete
    On Error GoTo 0

    targetCell.Value = displayText

    If Len(Trim$(address)) > 0 Then
        targetSheet.Hyperlinks.Add _
            Anchor:=targetCell, _
            Address:=address, _
            TextToDisplay:=displayText
    End If
End Sub


Private Function GetCellHyperlinkAddress(ByVal targetCell As Range) As String
    On Error Resume Next
    If targetCell.Hyperlinks.Count > 0 Then
        GetCellHyperlinkAddress = targetCell.Hyperlinks(1).Address
    End If
    On Error GoTo 0
End Function


Private Function AllDetailCellsPopulated( _
    ByVal sheet As Worksheet, _
    ByVal rowNumber As Long, _
    ByVal detailColumns As Variant _
) As Boolean
    Dim columnValue As Variant

    For Each columnValue In detailColumns
        If Len(Trim$(CStr(sheet.Cells(rowNumber, CLng(columnValue)).Value2))) = 0 Then
            Exit Function
        End If
    Next columnValue

    AllDetailCellsPopulated = True
End Function


Private Sub WriteTextIfBlank(ByVal targetCell As Range, ByVal value As String)
    If Len(Trim$(CStr(targetCell.Value2))) = 0 And Len(Trim$(value)) > 0 Then
        targetCell.Value = value
    End If
End Sub


Private Sub WriteNumberIfBlank(ByVal targetCell As Range, ByVal value As String)
    If Len(Trim$(CStr(targetCell.Value2))) = 0 And IsNumeric(value) Then
        targetCell.Value = CDbl(value)
    End If
End Sub


Private Sub WriteIsoDateIfBlank(ByVal targetCell As Range, ByVal value As String)
    If Len(Trim$(CStr(targetCell.Value2))) > 0 Or Len(value) = 0 Then Exit Sub

    If Len(value) = 10 And Mid$(value, 5, 1) = "-" And Mid$(value, 8, 1) = "-" Then
        targetCell.Value = DateSerial( _
            CInt(Left$(value, 4)), _
            CInt(Mid$(value, 6, 2)), _
            CInt(Right$(value, 2)) _
        )
        targetCell.NumberFormat = "yyyy-mm-dd"
    Else
        targetCell.Value = value
    End If
End Sub


' ============================================================================
' Selection and worksheet helpers
' ============================================================================

Private Function ActiveCatalogSheet(ByRef errorText As String) As Worksheet
    If ActiveWorkbook Is Nothing Then
        errorText = "Open a catalog workbook first."
        Exit Function
    End If

    If Not TypeOf ActiveSheet Is Worksheet Then
        errorText = "Activate a worksheet first."
        Exit Function
    End If

    Set ActiveCatalogSheet = ActiveSheet
End Function


Private Function SelectedRows(ByRef errorText As String) As Variant
    Dim area As Range
    Dim worksheetRow As Range
    Dim rows As Object
    Dim keys As Variant
    Dim i As Long
    Dim j As Long
    Dim swapValue As Variant

    If TypeName(Selection) <> "Range" Then
        errorText = "Select one or more data rows first."
        SelectedRows = Array()
        Exit Function
    End If

    Set rows = CreateObject("Scripting.Dictionary")

    For Each area In Selection.Areas
        For Each worksheetRow In area.Rows
            If worksheetRow.Row > HEADER_ROW Then rows(CStr(worksheetRow.Row)) = worksheetRow.Row
        Next worksheetRow
    Next area

    If rows.Count = 0 Then
        errorText = "Select one or more data rows below the header."
        SelectedRows = Array()
        Exit Function
    End If

    keys = rows.Items

    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If CLng(keys(j)) < CLng(keys(i)) Then
                swapValue = keys(i)
                keys(i) = keys(j)
                keys(j) = swapValue
            End If
        Next j
    Next i

    SelectedRows = keys
End Function


Private Function NormalizeHeader(ByVal value As String) As String
    Dim i As Long
    Dim character As String
    Dim result As String

    value = LCase$(Trim$(value))

    For i = 1 To Len(value)
        character = Mid$(value, i, 1)
        If character Like "[a-z0-9]" Then result = result & character
    Next i

    NormalizeHeader = result
End Function


Private Function FindHeaderColumn(ByVal sheet As Worksheet, ByVal candidates As Variant) As Long
    Dim lastColumn As Long
    Dim columnNumber As Long
    Dim candidate As Variant
    Dim normalized As String

    lastColumn = sheet.Cells(HEADER_ROW, sheet.Columns.Count).End(xlToLeft).Column

    For columnNumber = 1 To lastColumn
        normalized = NormalizeHeader(CStr(sheet.Cells(HEADER_ROW, columnNumber).Value2))

        For Each candidate In candidates
            If normalized = NormalizeHeader(CStr(candidate)) Then
                FindHeaderColumn = columnNumber
                Exit Function
            End If
        Next candidate
    Next columnNumber
End Function


Private Function NormalizeBarcode(ByVal cell As Range) As String
    Dim rawValue As String
    Dim i As Long
    Dim character As String
    Dim digits As String

    rawValue = Trim$(CStr(cell.Value2))

    If IsNumeric(cell.Value2) And InStr(1, rawValue, "E", vbTextCompare) > 0 Then
        rawValue = Format$(cell.Value2, "0")
    End If

    For i = 1 To Len(rawValue)
        character = Mid$(rawValue, i, 1)
        If character Like "[0-9]" Then digits = digits & character
    Next i

    If Len(digits) = 11 Then digits = "0" & digits
    NormalizeBarcode = digits
End Function


Private Function IsUPCERow(ByVal sheet As Worksheet, ByVal rowNumber As Long, ByVal typeColumn As Long) As Boolean
    If typeColumn <= 0 Then Exit Function
    IsUPCERow = (NormalizeHeader(CStr(sheet.Cells(rowNumber, typeColumn).Value2)) = "upce")
End Function


' ============================================================================
' Integrated selected-row resolver
' ============================================================================

Private Function TsvField(ByVal value As String) As String
    value = Replace(value, vbTab, " ")
    value = Replace(value, vbCr, " ")
    value = Replace(value, vbLf, " ")
    TsvField = Trim$(value)
End Function


Private Function CellTextOrHyperlink(ByVal targetCell As Range) As String
    CellTextOrHyperlink = Trim$(CStr(targetCell.Value2))
    If Len(CellTextOrHyperlink) = 0 Then
        CellTextOrHyperlink = GetCellHyperlinkAddress(targetCell)
    End If
End Function


Public Sub ResolveSelectedRows()
    Dim sheet As Worksheet
    Dim rows As Variant
    Dim rowValue As Variant
    Dim errorText As String
    Dim rootPath As String
    Dim pythonCommand As String
    Dim scriptPath As String
    Dim inputPath As String
    Dim outputPath As String
    Dim inputText As String
    Dim commandLine As String
    Dim standardOutput As String
    Dim standardError As String
    Dim exitCode As Long
    Dim upcColumn As Long
    Dim blurayUrlColumn As Long
    Dim releaseTitleColumn As Long
    Dim imdbUrlColumn As Long
    Dim imdbIdColumn As Long
    Dim titleColumn As Long
    Dim yearColumn As Long
    Dim runtimeColumn As Long
    Dim titleTypeColumn As Long
    Dim seasonColumn As Long
    Dim statusColumn As Long
    Dim studioColumn As Long
    Dim blurayYearColumn As Long
    Dim blurayRuntimeColumn As Long
    Dim ratingColumn As Long
    Dim releaseDateColumn As Long
    Dim discFormatColumn As Long
    Dim codecColumn As Long
    Dim resolutionColumn As Long
    Dim aspectColumn As Long
    Dim discCountColumn As Long
    Dim typeColumn As Long
    Dim code As String
    Dim blurayUrl As String
    Dim imdbUrl As String
    Dim releaseTitle As String
    Dim imdbId As String
    Dim canonicalTitle As String
    Dim season As String
    Dim requestedCount As Long
    Dim skippedBlank As Long
    Dim skippedUPCE As Long
    Dim lines As Variant
    Dim fields As Variant
    Dim lineIndex As Long
    Dim resultRow As Long
    Dim statusText As String
    Dim resolved As Long
    Dim partial As Long
    Dim review As Long
    Dim cancelled As Long
    Dim skipped As Long
    Dim failed As Long

    On Error GoTo FatalError

    Set sheet = ActiveCatalogSheet(errorText)
    If sheet Is Nothing Then GoTo ShowError

    rows = SelectedRows(errorText)
    If Len(errorText) > 0 Then GoTo ShowError

    upcColumn = FindHeaderColumn(sheet, Array("UPC", "UPC Code", "Barcode"))
    blurayUrlColumn = FindHeaderColumn(sheet, Array("Blu-ray.com URL", "Blu-ray URL", "Release URL"))
    releaseTitleColumn = FindHeaderColumn(sheet, Array("Blu-ray.com Title", "Release Title", "DVD Title"))
    imdbUrlColumn = FindHeaderColumn(sheet, Array("IMDb URL"))
    imdbIdColumn = FindHeaderColumn(sheet, Array("IMDb ID"))
    titleColumn = FindHeaderColumn(sheet, Array("IMDb Title", "Title"))
    yearColumn = FindHeaderColumn(sheet, Array("Year", "IMDb Year"))
    runtimeColumn = FindHeaderColumn(sheet, Array("Runtime", "IMDb Runtime"))
    titleTypeColumn = FindHeaderColumn(sheet, Array("Title Type", "IMDb Title Type"))
    seasonColumn = FindHeaderColumn(sheet, Array("Season", "IMDb Season"))
    statusColumn = FindHeaderColumn(sheet, Array("Status / Error", "Status", "Error"))
    studioColumn = FindHeaderColumn(sheet, Array("Studio"))
    blurayYearColumn = FindHeaderColumn(sheet, Array("Blu-ray Year"))
    blurayRuntimeColumn = FindHeaderColumn(sheet, Array("Blu-ray Runtime"))
    ratingColumn = FindHeaderColumn(sheet, Array("Content Rating", "Rating"))
    releaseDateColumn = FindHeaderColumn(sheet, Array("Physical Release Date", "Release Date"))
    discFormatColumn = FindHeaderColumn(sheet, Array("Disc Format", "Disk Format"))
    codecColumn = FindHeaderColumn(sheet, Array("Video Codec", "Codec"))
    resolutionColumn = FindHeaderColumn(sheet, Array("Resolution"))
    aspectColumn = FindHeaderColumn(sheet, Array("Aspect Ratio"))
    discCountColumn = FindHeaderColumn(sheet, Array("Disc Count / Capacities", "Disc Count", "Disk Count"))
    typeColumn = FindHeaderColumn(sheet, Array("format", "Barcode Type", "Symbology", "Type"))

    If upcColumn = 0 Or blurayUrlColumn = 0 Or releaseTitleColumn = 0 Or _
       imdbUrlColumn = 0 Or imdbIdColumn = 0 Or titleColumn = 0 Or _
       yearColumn = 0 Or runtimeColumn = 0 Or titleTypeColumn = 0 Or _
       seasonColumn = 0 Or statusColumn = 0 Or studioColumn = 0 Or _
       blurayYearColumn = 0 Or blurayRuntimeColumn = 0 Or ratingColumn = 0 Or _
       releaseDateColumn = 0 Or discFormatColumn = 0 Or codecColumn = 0 Or _
       resolutionColumn = 0 Or aspectColumn = 0 Or discCountColumn = 0 Then
        errorText = "The integrated resolver requires the standard MediaCatalog columns A through U."
        GoTo ShowError
    End If

    pythonCommand = GetPythonCommand(errorText)
    If Len(errorText) > 0 Then GoTo ShowError
    pythonCommand = GetWindowlessPythonCommand(pythonCommand, errorText)
    If Len(errorText) > 0 Then GoTo ShowError

    rootPath = ProjectPath()
    scriptPath = JoinPath(JoinPath(rootPath, "scripts"), "resolve_rows.py")
    If Not FileExists(scriptPath) Then
        errorText = "Integrated resolver was not found:" & vbCrLf & scriptPath
        GoTo ShowError
    End If

    inputText = "row" & vbTab & "upc" & vbTab & "bluray_url" & vbTab & _
                "release_title" & vbTab & "imdb_url" & vbTab & "imdb_id" & vbTab & _
                "title" & vbTab & "season" & vbCrLf

    For Each rowValue In rows
        If IsUPCERow(sheet, CLng(rowValue), typeColumn) Then
            skippedUPCE = skippedUPCE + 1
            sheet.Cells(CLng(rowValue), statusColumn).Value = "UPC_E skipped"
        Else
            code = NormalizeBarcode(sheet.Cells(CLng(rowValue), upcColumn))
            blurayUrl = CellTextOrHyperlink(sheet.Cells(CLng(rowValue), blurayUrlColumn))
            releaseTitle = Trim$(CStr(sheet.Cells(CLng(rowValue), releaseTitleColumn).Value2))
            imdbUrl = CellTextOrHyperlink(sheet.Cells(CLng(rowValue), imdbUrlColumn))
            imdbId = Trim$(CStr(sheet.Cells(CLng(rowValue), imdbIdColumn).Value2))
            canonicalTitle = Trim$(CStr(sheet.Cells(CLng(rowValue), titleColumn).Value2))
            season = Trim$(CStr(sheet.Cells(CLng(rowValue), seasonColumn).Value2))

            If Len(code & blurayUrl & releaseTitle & imdbUrl & imdbId & canonicalTitle) = 0 Then
                skippedBlank = skippedBlank + 1
            Else
                inputText = inputText & CStr(rowValue) & vbTab & _
                            TsvField(code) & vbTab & TsvField(blurayUrl) & vbTab & _
                            TsvField(releaseTitle) & vbTab & TsvField(imdbUrl) & vbTab & _
                            TsvField(imdbId) & vbTab & TsvField(canonicalTitle) & vbTab & _
                            TsvField(season) & vbCrLf
                requestedCount = requestedCount + 1
            End If
        End If
    Next rowValue

    If requestedCount = 0 Then
        MsgBox "No selected rows contained resolver input." & vbCrLf & vbCrLf & _
               "Blank: " & CStr(skippedBlank) & vbCrLf & _
               "UPC_E skipped: " & CStr(skippedUPCE), vbInformation, "MediaCatalog"
        Exit Sub
    End If

    inputPath = TemporaryPath("_integrated_input.tsv")
    outputPath = TemporaryPath("_integrated_output.tsv")
    WriteUtf8Text inputPath, inputText

    commandLine = QuoteArgument(pythonCommand) & " -E " & _
                  QuoteArgument(scriptPath) & " " & _
                  QuoteArgument(inputPath) & " " & QuoteArgument(outputPath)

    exitCode = RunCommandAndWait(commandLine, 43200, standardOutput, standardError)
    If exitCode <> 0 Or Not FileExists(outputPath) Then
        errorText = "Integrated resolver failed."
        If Len(Trim$(standardError)) > 0 Then
            errorText = errorText & vbCrLf & vbCrLf & Trim$(standardError)
        End If
        GoTo CleanupAndShowError
    End If

    lines = Split(Replace(ReadUtf8Text(outputPath), vbCrLf, vbLf), vbLf)
    For lineIndex = 1 To UBound(lines)
        If Len(Trim$(CStr(lines(lineIndex)))) > 0 Then
            fields = Split(CStr(lines(lineIndex)), vbTab)
            If UBound(fields) >= 24 Then
                resultRow = CLng(fields(0))

                If Len(fields(4)) > 0 Then
                    SetCellHyperlink sheet.Cells(resultRow, blurayUrlColumn), fields(4), fields(4)
                End If
                If Len(fields(5)) > 0 Then
                    SetCellHyperlink sheet.Cells(resultRow, releaseTitleColumn), fields(4), fields(5)
                End If

                If Len(fields(7)) > 0 Then
                    SetCellHyperlink sheet.Cells(resultRow, imdbUrlColumn), fields(6), fields(6)
                    sheet.Cells(resultRow, imdbIdColumn).Value = fields(7)
                End If
                If Len(fields(8)) > 0 Then
                    sheet.Cells(resultRow, titleColumn).Value = fields(8)
                    sheet.Cells(resultRow, yearColumn).Value = fields(9)
                    sheet.Cells(resultRow, runtimeColumn).Value = fields(10)
                    sheet.Cells(resultRow, titleTypeColumn).Value = fields(11)
                    sheet.Cells(resultRow, seasonColumn).Value = fields(12)
                End If

                WriteTextIfBlank sheet.Cells(resultRow, studioColumn), fields(13)
                WriteNumberIfBlank sheet.Cells(resultRow, blurayYearColumn), fields(14)
                WriteNumberIfBlank sheet.Cells(resultRow, blurayRuntimeColumn), fields(15)
                WriteTextIfBlank sheet.Cells(resultRow, ratingColumn), fields(16)
                WriteIsoDateIfBlank sheet.Cells(resultRow, releaseDateColumn), fields(17)
                WriteTextIfBlank sheet.Cells(resultRow, discFormatColumn), fields(18)
                WriteTextIfBlank sheet.Cells(resultRow, codecColumn), fields(19)
                WriteTextIfBlank sheet.Cells(resultRow, resolutionColumn), fields(20)
                WriteTextIfBlank sheet.Cells(resultRow, aspectColumn), fields(21)
                WriteTextIfBlank sheet.Cells(resultRow, discCountColumn), fields(22)

                statusText = fields(1)
                If Len(fields(2)) > 0 Then statusText = statusText & ": " & fields(2)
                sheet.Cells(resultRow, statusColumn).Value = statusText

                If Left$(fields(1), 3) = "OK " Or fields(1) = "OK" Then
                    resolved = resolved + 1
                ElseIf Left$(fields(1), 7) = "PARTIAL" Then
                    partial = partial + 1
                ElseIf Left$(fields(1), 12) = "NEEDS REVIEW" Then
                    review = review + 1
                ElseIf fields(1) = "CANCELLED" Then
                    cancelled = cancelled + 1
                ElseIf Left$(fields(1), 7) = "SKIPPED" Then
                    skipped = skipped + 1
                Else
                    failed = failed + 1
                End If
            Else
                failed = failed + 1
            End If
        End If
    Next lineIndex

    DeleteTemporaryFile inputPath
    DeleteTemporaryFile outputPath

    MsgBox "Integrated resolution finished." & vbCrLf & vbCrLf & _
           "Complete: " & CStr(resolved) & vbCrLf & _
           "Partial: " & CStr(partial) & vbCrLf & _
           "Needs review: " & CStr(review) & vbCrLf & _
           "Cancelled: " & CStr(cancelled) & vbCrLf & _
           "Skipped: " & CStr(skipped + skippedBlank + skippedUPCE) & vbCrLf & _
           "Errors: " & CStr(failed), _
           IIf(review + cancelled + failed > 0, vbExclamation, vbInformation), _
           "MediaCatalog"
    Exit Sub

CleanupAndShowError:
    DeleteTemporaryFile inputPath
    DeleteTemporaryFile outputPath

ShowError:
    MsgBox errorText, vbExclamation, "MediaCatalog"
    Exit Sub

FatalError:
    errorText = Err.Description
    Resume CleanupAndShowError
End Sub


' ============================================================================
' UPC/EAN release-title lookup
' ============================================================================

Public Sub ResolveSelectedUPCs()
    ResolveSelectedUPCsWithBRdC
End Sub


Public Sub ResolveSelectedUPCsWithBRdC()
    ResolveSelectedUPCsByProvider _
        "Blu-ray.com", "BRdC", "bluray_lookup_excel.py", True, True
End Sub


Public Sub ResolveSelectedUPCsWithUPCdb()
    ResolveSelectedUPCsByProvider _
        "UPCItemDB", "UPCdb", "upcitemdb_lookup.py", False, True
End Sub


Public Sub ResolveSelectedUPCsWithBarcodedCom()
    ResolveSelectedUPCsByProvider _
        "BarcodeLookup.com", "BarcodeLookup", _
        "barcodelookup_lookup.py", False, True
End Sub


Public Sub ResolveSelectedUPCsWithBarcodedComNoAPI()
    Dim sheet As Worksheet
    Dim rows As Variant
    Dim errorText As String
    Dim upcColumn As Long
    Dim rowNumber As Long
    Dim code As String

    Set sheet = ActiveCatalogSheet(errorText)
    If sheet Is Nothing Then GoTo ShowError

    rows = SelectedRows(errorText)
    If Len(errorText) > 0 Then GoTo ShowError

    If UBound(rows) <> LBound(rows) Then
        errorText = "Select exactly one data row for a manual Barcode Lookup search."
        GoTo ShowError
    End If

    upcColumn = FindHeaderColumn(sheet, Array("UPC", "UPC Code", "Barcode", "Barcode Value", "text"))
    If upcColumn = 0 Then
        errorText = "No UPC column was found in row 1."
        GoTo ShowError
    End If

    rowNumber = CLng(rows(LBound(rows)))
    code = NormalizeBarcode(sheet.Cells(rowNumber, upcColumn))
    If Len(code) <> 12 And Len(code) <> 13 Then
        errorText = "The selected row does not contain a 12-digit UPC or 13-digit EAN."
        GoTo ShowError
    End If

    sheet.Cells(rowNumber, upcColumn).Copy
    MsgBox "UPC " & code & " was copied to the clipboard." & vbCrLf & vbCrLf & _
           "BarcodeLookup.com will open next. Paste the UPC into its search box." & vbCrLf & _
           "This manual command does not read the website or modify catalog data.", _
           vbInformation, "MediaCatalog"
    ThisWorkbook.FollowHyperlink "https://www.barcodelookup.com/"
    Exit Sub

ShowError:
    MsgBox errorText, vbExclamation, "MediaCatalog"
End Sub


Private Function TitleNeedsResolver( _
    ByVal currentTitle As String, _
    ByVal retryMarkers As Boolean _
) As Boolean
    Dim normalized As String

    normalized = UCase$(Trim$(currentTitle))
    If Len(normalized) = 0 Then
        TitleNeedsResolver = True
    ElseIf retryMarkers Then
        TitleNeedsResolver = _
            (normalized = "[NOT FOUND]" Or normalized = "[AMBIGUOUS]")
    End If
End Function


Private Sub ResolveSelectedUPCsByProvider( _
    ByVal providerName As String, _
    ByVal providerCode As String, _
    ByVal scriptFileName As String, _
    ByVal addHyperlink As Boolean, _
    ByVal retryMarkers As Boolean _
)
    Dim sheet As Worksheet
    Dim rows As Variant
    Dim rowValue As Variant
    Dim errorText As String
    Dim rootPath As String
    Dim pythonCommand As String
    Dim scriptPath As String
    Dim inputPath As String
    Dim outputPath As String
    Dim inputText As String
    Dim commandLine As String
    Dim standardOutput As String
    Dim standardError As String
    Dim exitCode As Long
    Dim upcColumn As Long
    Dim titleColumn As Long
    Dim urlColumn As Long
    Dim statusColumn As Long
    Dim typeColumn As Long
    Dim code As String
    Dim releaseUrl As String
    Dim requestedCount As Long
    Dim skippedBlank As Long
    Dim skippedExisting As Long
    Dim skippedUPCE As Long
    Dim lines As Variant
    Dim fields As Variant
    Dim lineIndex As Long
    Dim resultRow As Long
    Dim matched As Long
    Dim notFound As Long
    Dim ambiguous As Long
    Dim failed As Long
    Dim cancelled As Long

    On Error GoTo FatalError

    Set sheet = ActiveCatalogSheet(errorText)
    If sheet Is Nothing Then GoTo ShowError

    rows = SelectedRows(errorText)
    If Len(errorText) > 0 Then GoTo ShowError

    upcColumn = FindHeaderColumn(sheet, Array("UPC", "UPC Code", "Barcode", "Barcode Value", "text"))
    titleColumn = FindHeaderColumn(sheet, Array("Blu-ray.com Title", "Release Title", "DVD Title", "UPCItemDB Name"))
    urlColumn = FindHeaderColumn(sheet, Array("Blu-ray.com URL", "Blu-ray URL", "Release URL"))
    statusColumn = FindHeaderColumn(sheet, Array("Status / Error", "Status", "Error"))
    typeColumn = FindHeaderColumn(sheet, Array("format", "Barcode Type", "Symbology", "Type"))

    If upcColumn = 0 Then
        errorText = "No UPC column was found in row 1."
        GoTo ShowError
    End If

    If titleColumn = 0 Then
        errorText = "No Blu-ray.com Title column was found in row 1."
        GoTo ShowError
    End If

    If addHyperlink And urlColumn = 0 Then
        errorText = "No Blu-ray.com URL column was found in row 1."
        GoTo ShowError
    End If

    pythonCommand = GetPythonCommand(errorText)
    If Len(errorText) > 0 Then GoTo ShowError

    rootPath = ProjectPath()
    scriptPath = JoinPath(JoinPath(rootPath, "scripts"), scriptFileName)

    If Not FileExists(scriptPath) Then
        errorText = providerName & " helper was not found:" & vbCrLf & scriptPath
        GoTo ShowError
    End If

    inputText = "row" & vbTab & "upc" & vbTab & "url" & vbCrLf

    For Each rowValue In rows
        If IsUPCERow(sheet, CLng(rowValue), typeColumn) Then
            skippedUPCE = skippedUPCE + 1
            If statusColumn > 0 Then sheet.Cells(CLng(rowValue), statusColumn).Value = "UPC_E skipped"
        Else
            code = NormalizeBarcode(sheet.Cells(CLng(rowValue), upcColumn))
            releaseUrl = ""
            If addHyperlink Then
                releaseUrl = Trim$(CStr(sheet.Cells(CLng(rowValue), urlColumn).Value2))
                If Len(releaseUrl) = 0 Then
                    releaseUrl = GetCellHyperlinkAddress(sheet.Cells(CLng(rowValue), urlColumn))
                End If
            End If

            If Len(code) = 0 And Len(releaseUrl) = 0 Then
                skippedBlank = skippedBlank + 1
            ElseIf Not TitleNeedsResolver( _
                CStr(sheet.Cells(CLng(rowValue), titleColumn).Value2), _
                retryMarkers _
            ) Then
                skippedExisting = skippedExisting + 1
            Else
                inputText = inputText & CStr(rowValue) & vbTab & code & vbTab & releaseUrl & vbCrLf
                requestedCount = requestedCount + 1
            End If
        End If
    Next rowValue

    If requestedCount = 0 Then
        MsgBox "No selected UPCs required lookup." & vbCrLf & vbCrLf & _
               "Blank: " & CStr(skippedBlank) & vbCrLf & _
               "Already populated: " & CStr(skippedExisting) & vbCrLf & _
               "UPC_E skipped: " & CStr(skippedUPCE), _
               vbInformation, "MediaCatalog"
        Exit Sub
    End If

    inputPath = TemporaryPath("_upc_input.tsv")
    outputPath = TemporaryPath("_upc_output.tsv")
    WriteUtf8Text inputPath, inputText

    commandLine = QuoteArgument(pythonCommand) & " -E " & _
                  QuoteArgument(scriptPath) & " " & _
                  QuoteArgument(inputPath) & " " & _
                  QuoteArgument(outputPath)

    exitCode = RunCommandAndWait(commandLine, 43200, standardOutput, standardError)

    If exitCode <> 0 Or Not FileExists(outputPath) Then
        errorText = providerName & " lookup failed."
        If Len(Trim$(standardError)) > 0 Then errorText = errorText & vbCrLf & vbCrLf & Trim$(standardError)
        GoTo CleanupAndShowError
    End If

    lines = Split(Replace(ReadUtf8Text(outputPath), vbCrLf, vbLf), vbLf)

    For lineIndex = 1 To UBound(lines)
        If Len(Trim$(CStr(lines(lineIndex)))) > 0 Then
            fields = Split(CStr(lines(lineIndex)), vbTab)

            If UBound(fields) >= 6 Then
                resultRow = CLng(fields(0))

                Select Case fields(2)
                    Case "OK"
                        If addHyperlink Then
                            SetCellHyperlink _
                                sheet.Cells(resultRow, titleColumn), _
                                fields(6), _
                                fields(3)
                            SetCellHyperlink _
                                sheet.Cells(resultRow, urlColumn), _
                                fields(6), _
                                fields(6)
                        Else
                            SetCellHyperlink _
                                sheet.Cells(resultRow, titleColumn), _
                                "", _
                                fields(3)
                        End If
                        If statusColumn > 0 Then
                            If fields(4) = "No UPC/EAN, URL OK" Then
                                sheet.Cells(resultRow, statusColumn).Value = fields(4)
                            Else
                                sheet.Cells(resultRow, statusColumn).Value = _
                                    "UPC OK - " & providerCode
                            End If
                        End If
                        matched = matched + 1

                    Case "NOT_FOUND"
                        sheet.Cells(resultRow, titleColumn).Value = "[NOT FOUND]"
                        If statusColumn > 0 Then sheet.Cells(resultRow, statusColumn).Value = providerCode & ": " & fields(5)
                        notFound = notFound + 1

                    Case "AMBIGUOUS"
                        sheet.Cells(resultRow, titleColumn).Value = "[AMBIGUOUS]"
                        If statusColumn > 0 Then sheet.Cells(resultRow, statusColumn).Value = providerCode & ": " & fields(5)
                        ambiguous = ambiguous + 1

                    Case "CANCELLED"
                        If statusColumn > 0 Then sheet.Cells(resultRow, statusColumn).Value = providerCode & ": Cancelled by user"
                        cancelled = cancelled + 1

                    Case Else
                        If statusColumn > 0 Then sheet.Cells(resultRow, statusColumn).Value = providerCode & ": " & fields(5)
                        failed = failed + 1
                End Select
            Else
                failed = failed + 1
            End If
        End If
    Next lineIndex

    DeleteTemporaryFile inputPath
    DeleteTemporaryFile outputPath

    MsgBox providerName & " lookup finished." & vbCrLf & vbCrLf & _
           "Matched: " & CStr(matched) & vbCrLf & _
           "Not found: " & CStr(notFound) & vbCrLf & _
           "Ambiguous: " & CStr(ambiguous) & vbCrLf & _
           "Cancelled: " & CStr(cancelled) & vbCrLf & _
           "Errors: " & CStr(failed) & vbCrLf & _
           "Already populated: " & CStr(skippedExisting) & vbCrLf & _
           "Blank: " & CStr(skippedBlank) & vbCrLf & _
           "UPC_E skipped: " & CStr(skippedUPCE), _
           IIf(failed + notFound + ambiguous + cancelled > 0, vbExclamation, vbInformation), _
           "MediaCatalog"
    Exit Sub

CleanupAndShowError:
    DeleteTemporaryFile inputPath
    DeleteTemporaryFile outputPath

ShowError:
    MsgBox errorText, vbExclamation, "MediaCatalog"
    Exit Sub

FatalError:
    errorText = Err.Description
    Resume CleanupAndShowError
End Sub


' ============================================================================
' Blu-ray.com release-detail enrichment
' ============================================================================

Public Sub EnrichSelectedBluRayDetails()
    Dim sheet As Worksheet
    Dim rows As Variant
    Dim rowValue As Variant
    Dim errorText As String
    Dim rootPath As String
    Dim pythonCommand As String
    Dim scriptPath As String
    Dim inputPath As String
    Dim outputPath As String
    Dim inputText As String
    Dim commandLine As String
    Dim standardOutput As String
    Dim standardError As String
    Dim exitCode As Long
    Dim upcColumn As Long
    Dim titleColumn As Long
    Dim urlColumn As Long
    Dim statusColumn As Long
    Dim studioColumn As Long
    Dim blurayYearColumn As Long
    Dim blurayRuntimeColumn As Long
    Dim ratingColumn As Long
    Dim releaseDateColumn As Long
    Dim discFormatColumn As Long
    Dim videoCodecColumn As Long
    Dim resolutionColumn As Long
    Dim aspectRatioColumn As Long
    Dim discCountColumn As Long
    Dim detailColumns As Variant
    Dim code As String
    Dim releaseUrl As String
    Dim requestedCount As Long
    Dim skippedBlank As Long
    Dim skippedExisting As Long
    Dim lines As Variant
    Dim fields As Variant
    Dim lineIndex As Long
    Dim resultRow As Long
    Dim enriched As Long
    Dim failed As Long
    Dim cancelled As Long

    On Error GoTo FatalError

    Set sheet = ActiveCatalogSheet(errorText)
    If sheet Is Nothing Then GoTo ShowError

    rows = SelectedRows(errorText)
    If Len(errorText) > 0 Then GoTo ShowError

    upcColumn = FindHeaderColumn(sheet, Array("UPC", "UPC Code", "Barcode", "Barcode Value", "text"))
    titleColumn = FindHeaderColumn(sheet, Array("Blu-ray.com Title", "Release Title", "DVD Title", "UPCItemDB Name"))
    urlColumn = FindHeaderColumn(sheet, Array("Blu-ray.com URL", "Blu-ray URL", "Release URL"))
    statusColumn = FindHeaderColumn(sheet, Array("Status / Error", "Status", "Error"))
    studioColumn = FindHeaderColumn(sheet, Array("Studio", "Blu-ray Studio"))
    blurayYearColumn = FindHeaderColumn(sheet, Array("Blu-ray Year", "Release Year"))
    blurayRuntimeColumn = FindHeaderColumn(sheet, Array("Blu-ray Runtime", "Release Runtime"))
    ratingColumn = FindHeaderColumn(sheet, Array("Content Rating", "Blu-ray Rating"))
    releaseDateColumn = FindHeaderColumn(sheet, Array("Physical Release Date", "Release Date"))
    discFormatColumn = FindHeaderColumn(sheet, Array("Disc Format", "Disk Format"))
    videoCodecColumn = FindHeaderColumn(sheet, Array("Video Codec", "Codec"))
    resolutionColumn = FindHeaderColumn(sheet, Array("Resolution", "Video Resolution"))
    aspectRatioColumn = FindHeaderColumn(sheet, Array("Aspect Ratio"))
    discCountColumn = FindHeaderColumn(sheet, Array("Disc Count / Capacities", "Disc Count and Capacities", "Disk Count and Capacities"))

    If upcColumn = 0 Or urlColumn = 0 Then
        errorText = "The UPC and Blu-ray.com URL columns are required."
        GoTo ShowError
    End If

    If studioColumn = 0 Or blurayYearColumn = 0 Or blurayRuntimeColumn = 0 Or _
       ratingColumn = 0 Or releaseDateColumn = 0 Or discFormatColumn = 0 Or _
       videoCodecColumn = 0 Or resolutionColumn = 0 Or aspectRatioColumn = 0 Or _
       discCountColumn = 0 Then
        errorText = "One or more Blu-ray detail columns are missing. Paste the v0.4.1 headers into row 1 or use the supplied template."
        GoTo ShowError
    End If

    detailColumns = Array( _
        studioColumn, blurayYearColumn, blurayRuntimeColumn, ratingColumn, _
        releaseDateColumn, discFormatColumn, videoCodecColumn, resolutionColumn, _
        aspectRatioColumn, discCountColumn _
    )

    pythonCommand = GetPythonCommand(errorText)
    If Len(errorText) > 0 Then GoTo ShowError
    pythonCommand = GetWindowlessPythonCommand(pythonCommand, errorText)
    If Len(errorText) > 0 Then GoTo ShowError

    rootPath = ProjectPath()
    scriptPath = JoinPath(JoinPath(rootPath, "scripts"), "bluray_details.py")
    If Not FileExists(scriptPath) Then
        errorText = "Blu-ray detail helper was not found:" & vbCrLf & scriptPath
        GoTo ShowError
    End If

    inputText = "row" & vbTab & "upc" & vbTab & "url" & vbCrLf

    For Each rowValue In rows
        If AllDetailCellsPopulated(sheet, CLng(rowValue), detailColumns) Then
            skippedExisting = skippedExisting + 1
        Else
            code = NormalizeBarcode(sheet.Cells(CLng(rowValue), upcColumn))
            releaseUrl = Trim$(CStr(sheet.Cells(CLng(rowValue), urlColumn).Value2))
            If Len(releaseUrl) = 0 Then
                releaseUrl = GetCellHyperlinkAddress(sheet.Cells(CLng(rowValue), urlColumn))
            End If

            If Len(code) = 0 And Len(releaseUrl) = 0 Then
                skippedBlank = skippedBlank + 1
            Else
                inputText = inputText & CStr(rowValue) & vbTab & code & vbTab & releaseUrl & vbCrLf
                requestedCount = requestedCount + 1
            End If
        End If
    Next rowValue

    If requestedCount = 0 Then
        MsgBox "No selected rows required Blu-ray detail enrichment." & vbCrLf & vbCrLf & _
               "Already populated: " & CStr(skippedExisting) & vbCrLf & _
               "No UPC or release link: " & CStr(skippedBlank), _
               vbInformation, "MediaCatalog"
        Exit Sub
    End If

    inputPath = TemporaryPath("_bluray_details_input.tsv")
    outputPath = TemporaryPath("_bluray_details_output.tsv")
    WriteUtf8Text inputPath, inputText

    commandLine = QuoteArgument(pythonCommand) & " -E " & _
                  QuoteArgument(scriptPath) & " " & _
                  QuoteArgument(inputPath) & " " & _
                  QuoteArgument(outputPath)

    exitCode = RunCommandAndWait(commandLine, 43200, standardOutput, standardError)

    If exitCode <> 0 Or Not FileExists(outputPath) Then
        errorText = "Blu-ray detail enrichment failed."
        If Len(Trim$(standardError)) > 0 Then errorText = errorText & vbCrLf & vbCrLf & Trim$(standardError)
        GoTo CleanupAndShowError
    End If

    lines = Split(Replace(ReadUtf8Text(outputPath), vbCrLf, vbLf), vbLf)

    For lineIndex = 1 To UBound(lines)
        If Len(Trim$(CStr(lines(lineIndex)))) > 0 Then
            fields = Split(CStr(lines(lineIndex)), vbTab)

            If UBound(fields) >= 15 Then
                resultRow = CLng(fields(0))

                If fields(2) = "OK" Then
                    WriteTextIfBlank sheet.Cells(resultRow, studioColumn), fields(6)
                    WriteTextIfBlank sheet.Cells(resultRow, blurayYearColumn), fields(7)
                    WriteNumberIfBlank sheet.Cells(resultRow, blurayRuntimeColumn), fields(8)
                    WriteTextIfBlank sheet.Cells(resultRow, ratingColumn), fields(9)
                    WriteIsoDateIfBlank sheet.Cells(resultRow, releaseDateColumn), fields(10)
                    WriteTextIfBlank sheet.Cells(resultRow, discFormatColumn), fields(11)
                    WriteTextIfBlank sheet.Cells(resultRow, videoCodecColumn), fields(12)
                    WriteTextIfBlank sheet.Cells(resultRow, resolutionColumn), fields(13)
                    WriteTextIfBlank sheet.Cells(resultRow, aspectRatioColumn), fields(14)
                    WriteTextIfBlank sheet.Cells(resultRow, discCountColumn), fields(15)
                    If Len(fields(5)) > 0 Then
                        SetCellHyperlink sheet.Cells(resultRow, urlColumn), fields(5), fields(5)
                    End If
                    If statusColumn > 0 Then
                        If fields(3) = "No UPC/EAN, URL OK" Then
                            sheet.Cells(resultRow, statusColumn).Value = fields(3)
                        ElseIf Len(Trim$(CStr(sheet.Cells(resultRow, statusColumn).Value2))) = 0 Then
                            sheet.Cells(resultRow, statusColumn).Value = "Blu-ray details OK"
                        End If
                    End If
                    enriched = enriched + 1
                ElseIf fields(2) = "CANCELLED" Then
                    If statusColumn > 0 Then sheet.Cells(resultRow, statusColumn).Value = "Blu-ray details cancelled"
                    cancelled = cancelled + 1
                Else
                    If statusColumn > 0 Then sheet.Cells(resultRow, statusColumn).Value = fields(4)
                    failed = failed + 1
                End If
            Else
                failed = failed + 1
            End If
        End If
    Next lineIndex

    DeleteTemporaryFile inputPath
    DeleteTemporaryFile outputPath

    MsgBox "Blu-ray detail enrichment finished." & vbCrLf & vbCrLf & _
           "Enriched: " & CStr(enriched) & vbCrLf & _
           "Cancelled: " & CStr(cancelled) & vbCrLf & _
           "Errors: " & CStr(failed) & vbCrLf & _
           "Already populated: " & CStr(skippedExisting) & vbCrLf & _
           "No UPC or release link: " & CStr(skippedBlank), _
           IIf(failed + cancelled > 0, vbExclamation, vbInformation), _
           "MediaCatalog"
    Exit Sub

CleanupAndShowError:
    DeleteTemporaryFile inputPath
    DeleteTemporaryFile outputPath

ShowError:
    MsgBox errorText, vbExclamation, "MediaCatalog"
    Exit Sub

FatalError:
    errorText = Err.Description
    Resume CleanupAndShowError
End Sub


' ============================================================================
' Local IMDb lookup
' ============================================================================

Public Sub LookupIMDbForCurrentRow()
    Dim sheet As Worksheet
    Dim rows As Variant
    Dim rowValue As Variant
    Dim errorText As String
    Dim rootPath As String
    Dim pythonCommand As String
    Dim scriptPath As String
    Dim rawName As String
    Dim imdbIdHint As String
    Dim inputPath As String
    Dim outputPath As String
    Dim commandLine As String
    Dim standardOutput As String
    Dim standardError As String
    Dim exitCode As Long
    Dim fields As Variant
    Dim response As String
    Dim resolved As Long
    Dim blankCount As Long
    Dim failed As Long
    Dim titleColumn As Long
    Dim statusColumn As Long
    Dim imdbUrlColumn As Long
    Dim imdbIdColumn As Long
    Dim imdbTitleColumn As Long
    Dim imdbYearColumn As Long
    Dim imdbRuntimeColumn As Long
    Dim imdbTitleTypeColumn As Long
    Dim imdbSeasonColumn As Long

    On Error GoTo FatalError

    Set sheet = ActiveCatalogSheet(errorText)
    If sheet Is Nothing Then GoTo ShowError

    rows = SelectedRows(errorText)
    If Len(errorText) > 0 Then GoTo ShowError

    titleColumn = FindHeaderColumn(sheet, Array("Blu-ray.com Title", "Release Title", "DVD Title", "UPCItemDB Name"))
    statusColumn = FindHeaderColumn(sheet, Array("Status / Error", "Status", "Error"))
    imdbUrlColumn = FindHeaderColumn(sheet, Array("IMDb URL"))
    imdbIdColumn = FindHeaderColumn(sheet, Array("IMDb ID"))
    imdbTitleColumn = FindHeaderColumn(sheet, Array("IMDb Title", "Title"))
    imdbYearColumn = FindHeaderColumn(sheet, Array("Year", "IMDb Year"))
    imdbRuntimeColumn = FindHeaderColumn(sheet, Array("Runtime", "IMDb Runtime"))
    imdbTitleTypeColumn = FindHeaderColumn(sheet, Array("Title Type", "IMDb Title Type"))
    imdbSeasonColumn = FindHeaderColumn(sheet, Array("Season", "IMDb Season"))

    If titleColumn = 0 Then
        errorText = "No release-title column was found in row 1."
        GoTo ShowError
    End If

    If imdbUrlColumn = 0 Or imdbIdColumn = 0 Or imdbTitleColumn = 0 Or _
       imdbYearColumn = 0 Or imdbRuntimeColumn = 0 Or imdbTitleTypeColumn = 0 Or _
       imdbSeasonColumn = 0 Or statusColumn = 0 Then
        errorText = "One or more IMDb output columns are missing from row 1."
        GoTo ShowError
    End If

    pythonCommand = GetPythonCommand(errorText)
    If Len(errorText) > 0 Then GoTo ShowError

    pythonCommand = GetWindowlessPythonCommand(pythonCommand, errorText)
    If Len(errorText) > 0 Then GoTo ShowError

    rootPath = ProjectPath()
    scriptPath = JoinPath(JoinPath(rootPath, "scripts"), "imdb_lookup_calc.py")

    If Not FileExists(scriptPath) Then
        errorText = "IMDb helper was not found:" & vbCrLf & scriptPath
        GoTo ShowError
    End If

    For Each rowValue In rows
        rawName = Trim$(CStr(sheet.Cells(CLng(rowValue), titleColumn).Value2))
        imdbIdHint = Trim$(CStr(sheet.Cells(CLng(rowValue), imdbIdColumn).Value2))

        If Len(rawName) = 0 Or rawName = "[NOT FOUND]" Or rawName = "[AMBIGUOUS]" Then
            blankCount = blankCount + 1
        Else
            inputPath = TemporaryPath("_imdb_input.txt")
            outputPath = TemporaryPath("_imdb_output.tsv")
            WriteUtf8Text inputPath, rawName & vbCrLf & imdbIdHint & vbCrLf

            commandLine = QuoteArgument(pythonCommand) & " -E " & _
                          QuoteArgument(scriptPath) & " " & _
                          QuoteArgument(inputPath) & " " & _
                          QuoteArgument(outputPath)

            standardOutput = ""
            standardError = ""
            exitCode = RunCommandAndWait(commandLine, 60, standardOutput, standardError)

            If exitCode = 0 And FileExists(outputPath) Then
                response = Replace(Replace(ReadUtf8Text(outputPath), vbCr, ""), vbLf, "")
                fields = Split(response, vbTab)

                If UBound(fields) >= 9 Then
                    If fields(0) = "1" Then
                        sheet.Cells(CLng(rowValue), imdbUrlColumn).Value = fields(2)
                        sheet.Cells(CLng(rowValue), imdbIdColumn).Value = fields(1)
                        sheet.Cells(CLng(rowValue), imdbTitleColumn).Value = fields(3)
                        sheet.Cells(CLng(rowValue), imdbYearColumn).Value = fields(4)
                        sheet.Cells(CLng(rowValue), imdbRuntimeColumn).Value = fields(5)
                        sheet.Cells(CLng(rowValue), imdbTitleTypeColumn).Value = fields(6)
                        sheet.Cells(CLng(rowValue), imdbSeasonColumn).Value = fields(7)

                        If fields(9) = "imdb_id" Then
                            sheet.Cells(CLng(rowValue), statusColumn).Value = "OK - IMDb ID"
                        Else
                            sheet.Cells(CLng(rowValue), statusColumn).Value = "OK"
                        End If

                        resolved = resolved + 1
                    Else
                        If statusColumn > 0 Then sheet.Cells(CLng(rowValue), statusColumn).Value = fields(8)
                        failed = failed + 1
                    End If
                Else
                    If statusColumn > 0 Then sheet.Cells(CLng(rowValue), statusColumn).Value = "Incomplete IMDb response"
                    failed = failed + 1
                End If
            Else
                If statusColumn > 0 Then sheet.Cells(CLng(rowValue), statusColumn).Value = Trim$(standardError)
                failed = failed + 1
            End If

            DeleteTemporaryFile inputPath
            DeleteTemporaryFile outputPath
        End If
    Next rowValue

    MsgBox "IMDb lookup finished." & vbCrLf & vbCrLf & _
           "Resolved/refreshed: " & CStr(resolved) & vbCrLf & _
           "Blank/unresolved release title: " & CStr(blankCount) & vbCrLf & _
           "Needs review: " & CStr(failed), _
           IIf(failed > 0, vbExclamation, vbInformation), _
           "MediaCatalog"
    Exit Sub

ShowError:
    MsgBox errorText, vbExclamation, "MediaCatalog"
    Exit Sub

FatalError:
    DeleteTemporaryFile inputPath
    DeleteTemporaryFile outputPath
    errorText = Err.Description
    Resume ShowError
End Sub


' ============================================================================
' UPC-E cleanup and manual lookup
' ============================================================================

Public Sub RemoveSelectedUPCERows()
    Dim sheet As Worksheet
    Dim rows As Variant
    Dim errorText As String
    Dim typeColumn As Long
    Dim index As Long
    Dim removed As Long

    Set sheet = ActiveCatalogSheet(errorText)
    If sheet Is Nothing Then GoTo ShowError

    rows = SelectedRows(errorText)
    If Len(errorText) > 0 Then GoTo ShowError

    typeColumn = FindHeaderColumn(sheet, Array("format", "Barcode Type", "Symbology", "Type"))

    If typeColumn = 0 Then
        errorText = "No barcode-type column was found. UPC-E rows can only be removed when the scanner export includes its symbology."
        GoTo ShowError
    End If

    Application.ScreenUpdating = False

    For index = UBound(rows) To LBound(rows) Step -1
        If IsUPCERow(sheet, CLng(rows(index)), typeColumn) Then
            sheet.Rows(CLng(rows(index))).Delete
            removed = removed + 1
        End If
    Next index

    Application.ScreenUpdating = True
    MsgBox CStr(removed) & " UPC-E row(s) removed.", vbInformation, "MediaCatalog"
    Exit Sub

ShowError:
    Application.ScreenUpdating = True
    MsgBox errorText, vbExclamation, "MediaCatalog"
End Sub


Public Sub OpenSelectedUPCOnBluRay()
    Dim sheet As Worksheet
    Dim errorText As String
    Dim rows As Variant
    Dim upcColumn As Long
    Dim code As String
    Dim fieldName As String
    Dim url As String

    Set sheet = ActiveCatalogSheet(errorText)
    If sheet Is Nothing Then GoTo ShowError

    rows = SelectedRows(errorText)
    If Len(errorText) > 0 Then GoTo ShowError

    upcColumn = FindHeaderColumn(sheet, Array("UPC", "UPC Code", "Barcode", "Barcode Value", "text"))
    If upcColumn = 0 Then
        errorText = "No UPC column was found in row 1."
        GoTo ShowError
    End If

    code = NormalizeBarcode(sheet.Cells(CLng(rows(LBound(rows))), upcColumn))
    If Len(code) <> 12 And Len(code) <> 13 Then
        errorText = "The first selected row does not contain a 12-digit UPC or 13-digit EAN."
        GoTo ShowError
    End If

    fieldName = IIf(Len(code) = 13, "ean", "upc")
    url = "https://www.blu-ray.com/dvd/search.php?" & fieldName & "=" & code & "&action=search"
    ThisWorkbook.FollowHyperlink url
    Exit Sub

ShowError:
    MsgBox errorText, vbExclamation, "MediaCatalog"
End Sub


' ============================================================================
' Diagnostics
' ============================================================================

Public Sub CheckMediaCatalogConfiguration()
    Dim rootPath As String
    Dim settingsPath As String
    Dim pythonCommand As String
    Dim windowlessPythonCommand As String
    Dim errorText As String
    Dim blurayScript As String
    Dim upcitemdbScript As String
    Dim barcodelookupScript As String
    Dim blurayDetailsScript As String
    Dim imdbScript As String
    Dim integratedScript As String
    Dim databaseSetting As String
    Dim databasePath As String
    Dim barcodeApiKey As String
    Dim barcodePaidSubscription As Boolean
    Dim report As String

    rootPath = ProjectPath()
    settingsPath = JoinPath(rootPath, "settings.ini")
    pythonCommand = GetPythonCommand(errorText)

    If Len(errorText) > 0 Then
        MsgBox errorText, vbExclamation, "MediaCatalog"
        Exit Sub
    End If

    windowlessPythonCommand = GetWindowlessPythonCommand(pythonCommand, errorText)

    If Len(errorText) > 0 Then
        MsgBox errorText, vbExclamation, "MediaCatalog"
        Exit Sub
    End If

    blurayScript = JoinPath(JoinPath(rootPath, "scripts"), "bluray_lookup_excel.py")
    upcitemdbScript = JoinPath(JoinPath(rootPath, "scripts"), "upcitemdb_lookup.py")
    barcodelookupScript = JoinPath(JoinPath(rootPath, "scripts"), "barcodelookup_lookup.py")
    blurayDetailsScript = JoinPath(JoinPath(rootPath, "scripts"), "bluray_details.py")
    imdbScript = JoinPath(JoinPath(rootPath, "scripts"), "imdb_lookup_calc.py")
    integratedScript = JoinPath(JoinPath(rootPath, "scripts"), "resolve_rows.py")
    databaseSetting = ReadIniValue(settingsPath, "paths", "imdb_database", "data/imdb.sqlite")
    barcodeApiKey = ReadIniValue(settingsPath, "barcodelookup", "api_key", "")
    barcodePaidSubscription = _
        (LCase$(Trim$(ReadIniValue( _
            settingsPath, "barcodelookup", "paid_subscription", "false" _
        ))) = "true")

    If InStr(databaseSetting, ":") > 0 Or Left$(databaseSetting, 2) = "\\" Then
        databasePath = Replace(databaseSetting, "/", "\")
    Else
        databasePath = JoinPath(rootPath, Replace(databaseSetting, "/", "\"))
    End If

    report = "MediaCatalog configuration" & vbCrLf & vbCrLf & _
             "Project: " & rootPath & vbCrLf & _
             "Python: " & pythonCommand & vbCrLf & _
             "IMDb Python (windowless): " & windowlessPythonCommand & vbCrLf & _
             "Blu-ray helper: " & IIf(FileExists(blurayScript), "OK", "MISSING") & vbCrLf & _
             "UPCItemDB helper: " & IIf(FileExists(upcitemdbScript), "OK", "MISSING") & vbCrLf & _
             "Barcode Lookup helper: " & IIf(FileExists(barcodelookupScript), "OK", "MISSING") & vbCrLf & _
             "Barcode Lookup API key: " & IIf(Len(Trim$(barcodeApiKey)) > 0, "OK", "MISSING") & vbCrLf & _
             "Barcode Lookup paid subscription: " & IIf(barcodePaidSubscription, "CONFIRMED", "NOT CONFIRMED") & vbCrLf & _
             "Blu-ray details helper: " & IIf(FileExists(blurayDetailsScript), "OK", "MISSING") & vbCrLf & _
             "IMDb helper: " & IIf(FileExists(imdbScript), "OK", "MISSING") & vbCrLf & _
             "Integrated resolver: " & IIf(FileExists(integratedScript), "OK", "MISSING") & vbCrLf & _
             "IMDb database: " & IIf(FileExists(databasePath), "OK", "MISSING") & vbCrLf & _
             databasePath

    MsgBox report, IIf(FileExists(blurayScript) And FileExists(upcitemdbScript) And FileExists(barcodelookupScript) And Len(Trim$(barcodeApiKey)) > 0 And barcodePaidSubscription And FileExists(blurayDetailsScript) And FileExists(imdbScript) And FileExists(integratedScript) And FileExists(databasePath), vbInformation, vbExclamation), "MediaCatalog"
End Sub
