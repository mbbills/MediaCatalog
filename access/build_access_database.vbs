Option Explicit

' Builds access\MediaCatalog.accdb from source-controlled text:
'   - schema.sql defines the MediaCatalog table (executed verbatim as DDL).
'   - This script then adds one bound data-entry/browse form via the
'     classic CreateForm/CreateControl Access automation API.
'
' Run on a Windows machine with Access 2019 (or a compatible Access
' version) installed:
'
'   cscript access\build_access_database.vbs PROJECT_ROOT
'
' The .accdb itself is a proprietary binary format, so -- exactly like
' MediaCatalog_template.xlsm, which is built by build_excel_template.vbs
' rather than hand-edited -- it is not meaningfully hand-authorable or
' diffable. This script, and schema.sql, are the real source of truth;
' rerun this script any time the schema needs to change.
'
' NOTE: this script has not been run against a real copy of Access -- the
' development environment that produced it has no Windows, no Access, and
' no way to open or validate an .accdb file. Treat it as a first draft:
' run it, and if Access reports a COM/automation error, report the exact
' error message and line so it can be fixed (the same process
' build_excel_template.vbs went through during Windows 7 testing).

Dim fso, projectRoot, accessDir, schemaFile, outputDatabase
Dim textStream, schemaSql
Dim access

Const acTextBox = 109
Const acDetail = 0
Const acForm = 2
Const acSaveYes = 1

If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript build_access_database.vbs PROJECT_ROOT"
    WScript.Quit 2
End If

Set fso = CreateObject("Scripting.FileSystemObject")
projectRoot = fso.GetAbsolutePathName(WScript.Arguments(0))
accessDir = fso.BuildPath(projectRoot, "access")
schemaFile = fso.BuildPath(accessDir, "schema.sql")
outputDatabase = fso.BuildPath(accessDir, "MediaCatalog.accdb")

If Not fso.FileExists(schemaFile) Then Fail "Schema file not found: " & schemaFile

Set textStream = fso.OpenTextFile(schemaFile, 1, False, 0)
schemaSql = textStream.ReadAll
textStream.Close

' Strip SQL comment lines (Access's DDL executor does not accept "--"
' comments the way its interactive Query Designer's SQL view does).
schemaSql = StripSqlComments(schemaSql)

If fso.FileExists(outputDatabase) Then fso.DeleteFile outputDatabase, True

On Error Resume Next
Set access = CreateObject("Access.Application")
If Err.Number <> 0 Then Fail "Access could not be started: " & Err.Description
On Error GoTo 0

access.Visible = False

On Error Resume Next
access.NewCurrentDatabase outputDatabase
If Err.Number <> 0 Then
    Dim createError
    createError = Err.Description
    access.Quit
    Fail "Could not create a new .accdb: " & createError
End If
On Error GoTo 0

On Error Resume Next
access.CurrentDb.Execute schemaSql
If Err.Number <> 0 Then
    Dim ddlError
    ddlError = Err.Description
    access.CloseCurrentDatabase
    access.Quit
    Fail "Table creation failed (schema.sql): " & ddlError
End If
On Error GoTo 0

BuildBrowseForm access

access.CloseCurrentDatabase
access.Quit
Set access = Nothing

WScript.Echo "Created " & outputDatabase
WScript.Quit 0


Sub BuildBrowseForm(app)
    ' One plain bound form: a vertical stack of label+textbox pairs, one
    ' per field, enough to add a UPC and view/edit a row. CreateControl's
    ' ColumnName argument automatically attaches a matching label, so no
    ' separate label controls are created here.
    Dim fields, i, topPos, rowHeight, textLeft, textWidth
    Dim frm, ctl, finalName

    fields = Array( _
        "Inventory Number", "UPC", "Blu-ray.com URL", "Blu-ray.com Title", _
        "IMDb URL", "IMDb ID", "IMDb Title", "Year", "Runtime", _
        "Title Type", "Season", "Status / Error", "Studio", _
        "Blu-ray Year", "Blu-ray Runtime", "Content Rating", _
        "Physical Release Date", "Disc Format", "Video Codec", _
        "Resolution", "Aspect Ratio", "Disc Count / Capacities" _
    )

    rowHeight = 350   ' twips (~0.24in) per row
    textLeft = 2100   ' twips
    textWidth = 3200  ' twips
    topPos = 150

    On Error Resume Next
    Set frm = app.CreateForm()
    If Err.Number <> 0 Then Fail "CreateForm failed: " & Err.Description
    On Error GoTo 0

    frm.RecordSource = "MediaCatalog"

    For i = LBound(fields) To UBound(fields)
        On Error Resume Next
        ' VBScript's late-bound IDispatch calls into a COM object (unlike
        ' its own intrinsic functions, e.g. MsgBox) do not support skipping
        ' an argument with a bare "," ",": that compiles fine in VBA but is
        ' a syntax error here. Pass Empty explicitly for the unused
        ' ParentName slot, and set position/size as properties afterward
        ' instead of passing them positionally, to avoid the same risk.
        Set ctl = app.CreateControl(frm.Name, acTextBox, acDetail, Empty, fields(i))
        If Err.Number <> 0 Then
            Fail "CreateControl failed for field '" & fields(i) & "': " & Err.Description
        End If
        On Error GoTo 0
        ctl.Left = textLeft
        ctl.Top = topPos
        ctl.Width = textWidth
        ctl.Height = 250
        ' CreateControl auto-creates the attached label to the left of the
        ' textbox at its own default offset/width; left untouched here.
        topPos = topPos + rowHeight
    Next i

    frm.Caption = "MediaCatalog"
    finalName = frm.Name

    On Error Resume Next
    app.DoCmd.Close acForm, finalName, acSaveYes
    If Err.Number <> 0 Then Fail "Saving the form failed: " & Err.Description
    On Error GoTo 0

    On Error Resume Next
    app.DoCmd.Rename "frmMediaCatalog", acForm, finalName
    If Err.Number <> 0 Then Fail "Renaming the form failed: " & Err.Description
    On Error GoTo 0
End Sub


Function StripSqlComments(sql)
    Dim lines, out, i, line, trimmed
    lines = Split(sql, Chr(10))
    out = ""
    For i = LBound(lines) To UBound(lines)
        line = lines(i)
        trimmed = Trim(line)
        If Left(trimmed, 2) <> "--" Then
            out = out & line & Chr(10)
        End If
    Next i
    StripSqlComments = out
End Function


Sub Fail(message)
    WScript.Echo "ERROR: " & message
    WScript.Quit 1
End Sub
