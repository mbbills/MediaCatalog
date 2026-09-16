Attribute VB_Name = "MediaCatalogAccess"
Option Explicit

' MediaCatalog Access module -- Phase 2: integrated resolver (current
' record only).
'
' This ports Excel's ResolveSelectedRows (excel/MediaCatalog_Excel_Module.bas)
' to Access's object model: one bound form record instead of a selected
' spreadsheet range. The Shell/poll pattern (WScript.Shell.Exec, poll
' process.Status with Sleep 100 in a loop, 43200-second timeout) and the
' settings.ini/Python-path convention are copied from Excel's
' RunCommandAndWait/GetPythonCommand/GetWindowlessPythonCommand verbatim,
' since Access shares Excel's VBA dialect -- not reinvented here.
'
' Access field names differ from the canonical spreadsheet headers for two
' fields (Access field names cannot contain a period): "Blu-ray.com URL"
' is the Access field "Blu-ray com URL", and "Blu-ray.com Title" is
' "Blu-ray com Title". See the mapping table in access/README.md. Every
' other Access field name is identical to its spreadsheet header.
'
' Only "Resolve Selected Rows" is implemented here (as "Resolve Current
' Record" -- Access has no spreadsheet-style multi-row selection, so this
' phase resolves the one record currently open on the form). The
' BRdC-only and single-row-IMDb-lookup commands are explicitly deferred
' to later phases.

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
#End If


' ============================================================================
' Configuration and file helpers (ported from excel/MediaCatalog_Excel_Module.bas)
' ============================================================================

Private Function ProjectPath() As String
    ' MediaCatalog.accdb lives in access\ (see build_access_database.vbs),
    ' one level below the project root where settings.ini and scripts\
    ' actually live -- unlike Excel/Calc, whose templates sit at the
    ' project root directly. Strip the last path segment to get there.
    Dim accessDirPath As String
    Dim lastSlash As Long

    accessDirPath = CurrentProject.Path
    lastSlash = InStrRev(accessDirPath, "\")
    If lastSlash > 0 Then
        ProjectPath = Left$(accessDirPath, lastSlash - 1)
    Else
        ProjectPath = accessDirPath
    End If
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


Private Function TsvField(ByVal value As String) As String
    value = Replace(value, vbTab, " ")
    value = Replace(value, vbCr, " ")
    value = Replace(value, vbLf, " ")
    TsvField = Trim$(value)
End Function


' Only fills a currently-blank text field -- matches Excel's
' WriteTextIfBlank: manual/prior values are never overwritten by
' secondary enrichment fields (Studio, Blu-ray Year, etc.).
Private Sub WriteTextIfBlank(ByRef frm As Form, ByVal fieldName As String, ByVal value As String)
    If Len(Trim$(Nz(frm.Controls(fieldName).Value, ""))) = 0 And Len(Trim$(value)) > 0 Then
        frm.Controls(fieldName).Value = value
    End If
End Sub


' Only fills a currently-blank numeric field, and only when the resolver's
' value actually parses as a number -- matches Excel's WriteNumberIfBlank's
' IsNumeric guard (garbage in the TSV must not raise a type-mismatch error
' writing to an Integer field).
Private Sub WriteNumberIfBlank(ByRef frm As Form, ByVal fieldName As String, ByVal value As String)
    If IsNull(frm.Controls(fieldName).Value) And IsNumeric(value) Then
        frm.Controls(fieldName).Value = CLng(value)
    End If
End Sub


' Only fills a currently-blank date field, matching Excel's
' WriteIsoDateIfBlank. resolve_rows.py's release_date is either empty or a
' strict YYYY-MM-DD string; anything else is left alone rather than risk
' an invalid-date error.
Private Sub WriteIsoDateIfBlank(ByRef frm As Form, ByVal fieldName As String, ByVal value As String)
    If Not IsNull(frm.Controls(fieldName).Value) Then Exit Sub
    If Len(value) <> 10 Then Exit Sub
    If Mid$(value, 5, 1) <> "-" Or Mid$(value, 8, 1) <> "-" Then Exit Sub

    On Error Resume Next
    frm.Controls(fieldName).Value = DateSerial( _
        CInt(Left$(value, 4)), _
        CInt(Mid$(value, 6, 2)), _
        CInt(Right$(value, 2)) _
    )
    On Error GoTo 0
End Sub


' Always overwrites -- matches Excel's unconditional writes for the
' primary identity fields (Blu-ray release/IMDb work identity) and the
' Status / Error field, which the resolver is expected to refresh on
' every run.
Private Sub WriteNumberValue(ByRef frm As Form, ByVal fieldName As String, ByVal value As String)
    If IsNumeric(value) Then
        frm.Controls(fieldName).Value = CLng(value)
    Else
        frm.Controls(fieldName).Value = Null
    End If
End Sub


Private Sub WriteTextValue(ByRef frm As Form, ByVal fieldName As String, ByVal value As String)
    frm.Controls(fieldName).Value = value
End Sub


' ============================================================================
' Resolve Current Record (Phase 2: integrated resolver, ResolveSelectedRows
' equivalent)
' ============================================================================

Public Function ResolveCurrentRecord() As Variant
    Dim frm As Form
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
    Dim upc As String
    Dim blurayUrl As String
    Dim releaseTitle As String
    Dim imdbUrl As String
    Dim imdbId As String
    Dim canonicalTitle As String
    Dim season As String
    Dim lines As Variant
    Dim fields As Variant
    Dim statusText As String

    On Error GoTo FatalError

    Set frm = Screen.ActiveForm

    If frm.NewRecord Then
        MsgBox "Enter and save some data (a UPC, a Blu-ray.com URL, or an IMDb ID) before resolving.", _
               vbInformation, "MediaCatalog"
        Exit Function
    End If

    ' Commit any in-progress edit first so the field values read below (and
    ' the eventual re-save after the resolver writes back) reflect what is
    ' actually in the record, not a stale/uncommitted buffer.
    If frm.Dirty Then frm.Dirty = False

    ' Controls are referenced by name string (frm.Controls("Field Name"))
    ' rather than bang-bracket syntax (frm![Field Name]) throughout this
    ' function, to sidestep any doubt about how the VBA parser treats the
    ' hyphens/slashes/spaces in field names like "Blu-ray com URL" or
    ' "Status / Error" -- this project has already been burned twice by
    ' automation-syntax assumptions that seemed safe but weren't (see
    ' access/README.md's "Verified vs. not verified" history).
    upc = Trim$(Nz(frm.Controls("UPC").Value, ""))
    blurayUrl = Trim$(Nz(frm.Controls("Blu-ray com URL").Value, ""))
    releaseTitle = Trim$(Nz(frm.Controls("Blu-ray com Title").Value, ""))
    imdbUrl = Trim$(Nz(frm.Controls("IMDb URL").Value, ""))
    imdbId = Trim$(Nz(frm.Controls("IMDb ID").Value, ""))
    canonicalTitle = Trim$(Nz(frm.Controls("IMDb Title").Value, ""))
    If IsNull(frm.Controls("Season").Value) Then
        season = ""
    Else
        season = CStr(frm.Controls("Season").Value)
    End If

    If Len(upc & blurayUrl & releaseTitle & imdbUrl & imdbId & canonicalTitle) = 0 Then
        MsgBox "This record has no UPC, Blu-ray.com URL, Blu-ray.com Title, IMDb URL, IMDb ID, or IMDb Title to resolve from.", _
               vbInformation, "MediaCatalog"
        Exit Function
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

    ' Same input header/field order resolve_rows.py expects from Excel/Calc.
    inputText = "row" & vbTab & "upc" & vbTab & "bluray_url" & vbTab & _
                "release_title" & vbTab & "imdb_url" & vbTab & "imdb_id" & vbTab & _
                "title" & vbTab & "season" & vbCrLf
    inputText = inputText & "1" & vbTab & TsvField(upc) & vbTab & _
                TsvField(blurayUrl) & vbTab & TsvField(releaseTitle) & vbTab & _
                TsvField(imdbUrl) & vbTab & TsvField(imdbId) & vbTab & _
                TsvField(canonicalTitle) & vbTab & TsvField(season) & vbCrLf

    inputPath = TemporaryPath("_access_integrated_input.tsv")
    outputPath = TemporaryPath("_access_integrated_output.tsv")
    WriteUtf8Text inputPath, inputText

    commandLine = QuoteArgument(pythonCommand) & " -E " & _
                  QuoteArgument(scriptPath) & " " & _
                  QuoteArgument(inputPath) & " " & QuoteArgument(outputPath)

    exitCode = RunCommandAndWait(commandLine, 43200, standardOutput, standardError)
    If Not FileExists(outputPath) Then
        errorText = "Integrated resolver failed."
        If Len(Trim$(standardError)) > 0 Then
            errorText = errorText & vbCrLf & vbCrLf & Trim$(standardError)
        End If
        GoTo CleanupAndShowError
    End If

    ' As with Excel/Calc: the resolver writes its (one) row's result as
    ' soon as it is resolved, so a non-zero exit code does not by itself
    ' mean the output file is empty or unusable.
    If exitCode <> 0 And Len(Trim$(standardError)) > 0 Then
        errorText = "The integrated resolver did not finish cleanly:" & vbCrLf & Trim$(standardError)
    End If

    lines = Split(Replace(ReadUtf8Text(outputPath), vbCrLf, vbLf), vbLf)

    Dim resultLine As String
    Dim lineIndex As Long
    resultLine = ""
    For lineIndex = 1 To UBound(lines)
        If Len(Trim$(CStr(lines(lineIndex)))) > 0 Then
            resultLine = CStr(lines(lineIndex))
            Exit For
        End If
    Next lineIndex

    If Len(resultLine) = 0 Then
        If Len(errorText) = 0 Then
            errorText = "The integrated resolver produced no result for this record."
        End If
        GoTo CleanupAndShowError
    End If

    fields = Split(resultLine, vbTab)
    If UBound(fields) < 24 Then
        If Len(errorText) = 0 Then
            errorText = "The integrated resolver returned an incomplete response."
        End If
        GoTo CleanupAndShowError
    End If

    ' Field indices match resolve_rows.py's OUTPUT_FIELDS exactly:
    ' 0 row, 1 status, 2 error, 3 upc, 4 bluray_url, 5 release_title,
    ' 6 imdb_url, 7 imdb_id, 8 title, 9 year, 10 runtime, 11 title_type,
    ' 12 season, 13 studio, 14 bluray_year, 15 bluray_runtime, 16 rating,
    ' 17 release_date, 18 disc_format, 19 video_codec, 20 resolution,
    ' 21 aspect_ratio, 22 disc_count_capacities, 23 source, 24 warning.

    If Len(fields(4)) > 0 Then WriteTextValue frm, "Blu-ray com URL", CStr(fields(4))
    If Len(fields(5)) > 0 Then WriteTextValue frm, "Blu-ray com Title", CStr(fields(5))

    If Len(fields(7)) > 0 Then
        WriteTextValue frm, "IMDb URL", CStr(fields(6))
        WriteTextValue frm, "IMDb ID", CStr(fields(7))
    End If
    If Len(fields(8)) > 0 Then
        WriteTextValue frm, "IMDb Title", CStr(fields(8))
        WriteNumberValue frm, "Year", CStr(fields(9))
        WriteNumberValue frm, "Runtime", CStr(fields(10))
        WriteTextValue frm, "Title Type", CStr(fields(11))
        WriteNumberValue frm, "Season", CStr(fields(12))
    End If

    WriteTextIfBlank frm, "Studio", CStr(fields(13))
    WriteNumberIfBlank frm, "Blu-ray Year", CStr(fields(14))
    WriteNumberIfBlank frm, "Blu-ray Runtime", CStr(fields(15))
    WriteTextIfBlank frm, "Content Rating", CStr(fields(16))
    WriteIsoDateIfBlank frm, "Physical Release Date", CStr(fields(17))
    WriteTextIfBlank frm, "Disc Format", CStr(fields(18))
    WriteTextIfBlank frm, "Video Codec", CStr(fields(19))
    WriteTextIfBlank frm, "Resolution", CStr(fields(20))
    WriteTextIfBlank frm, "Aspect Ratio", CStr(fields(21))
    WriteTextIfBlank frm, "Disc Count / Capacities", CStr(fields(22))

    ' Same status vocabulary resolve_rows.py already emits (e.g.
    ' "OK - Blu-ray + IMDb", "PARTIAL - Blu-ray only", "NEEDS REVIEW",
    ' "SKIPPED - no resolver input"), not reinvented here.
    statusText = CStr(fields(1))
    If Len(fields(2)) > 0 Then statusText = statusText & ": " & CStr(fields(2))
    WriteTextValue frm, "Status / Error", statusText

    If frm.Dirty Then frm.Dirty = False

    DeleteTemporaryFile inputPath
    DeleteTemporaryFile outputPath

    If Len(errorText) > 0 Then
        MsgBox errorText & vbCrLf & vbCrLf & "Status: " & statusText, vbExclamation, "MediaCatalog"
    Else
        MsgBox "Resolved: " & statusText, vbInformation, "MediaCatalog"
    End If

    Exit Function

CleanupAndShowError:
    ' ReadUtf8Text/WriteUtf8Text each open and Close their own ADODB.Stream
    ' within a single call, so no stream handle is ever left open across
    ' statements the way this project's LibreOffice Calc macros once left
    ' a ScriptForge TextStream open across a failed run -- there is no
    ' handle to close here before deleting the temp files.
    DeleteTemporaryFile inputPath
    DeleteTemporaryFile outputPath

ShowError:
    MsgBox errorText, vbExclamation, "MediaCatalog"
    Exit Function

FatalError:
    errorText = Err.Description
    Resume CleanupAndShowError
End Function
