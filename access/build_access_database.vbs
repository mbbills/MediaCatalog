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
Const acLabel = 100
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


' One plain bound form: a vertical stack of label+textbox pairs, one per
' field, enough to add a UPC and view/edit a row. CreateControl is called
' with only its required leading arguments (FormName, ControlType,
' Section); every position/size/binding property is set explicitly
' afterward, and the label for each field is created as its own separate
' control rather than relying on ColumnName auto-attaching one. This
' avoids passing any optional/skipped argument into a late-bound COM call
' at all -- an earlier version passed ParentName/ColumnName/Left/Top/
' Width/Height positionally (skipping ParentName, then passing Empty for
' it), and real Access 2019 rejected both forms at script-compile time
' with "Expected end of statement", which neither this project's own
' heuristic checks nor any available offline VBScript tool could
' reproduce -- so the fix here is to stop relying on that argument list
' shape entirely rather than guess at another variant of it.
Sub BuildBrowseForm(app)
    Dim fields, i, topPos, rowHeight, labelLeft, labelWidth, textLeft, textWidth
    Dim frm, ctl, lbl, finalName

    fields = Array( _
        "Inventory Number", "UPC", "Blu-ray.com URL", "Blu-ray.com Title", _
        "IMDb URL", "IMDb ID", "IMDb Title", "Year", "Runtime", _
        "Title Type", "Season", "Status / Error", "Studio", _
        "Blu-ray Year", "Blu-ray Runtime", "Content Rating", _
        "Physical Release Date", "Disc Format", "Video Codec", _
        "Resolution", "Aspect Ratio", "Disc Count / Capacities" _
    )

    rowHeight = 350    ' twips (~0.24in) per row
    labelLeft = 100    ' twips
    labelWidth = 1900  ' twips (~1.3in)
    textLeft = 2100    ' twips
    textWidth = 3200   ' twips
    topPos = 150

    On Error Resume Next
    Set frm = app.CreateForm()
    If Err.Number <> 0 Then Fail "CreateForm failed: " & Err.Description
    On Error GoTo 0

    frm.RecordSource = "MediaCatalog"

    For i = LBound(fields) To UBound(fields)
        On Error Resume Next
        Set ctl = app.CreateControl(frm.Name, acTextBox, acDetail)
        If Err.Number <> 0 Then
            Fail "CreateControl (textbox) failed for field '" & fields(i) & "': " & Err.Description
        End If
        On Error GoTo 0
        ctl.ControlSource = fields(i)
        ctl.Left = textLeft
        ctl.Top = topPos
        ctl.Width = textWidth
        ctl.Height = 250

        On Error Resume Next
        Set lbl = app.CreateControl(frm.Name, acLabel, acDetail)
        If Err.Number <> 0 Then
            Fail "CreateControl (label) failed for field '" & fields(i) & "': " & Err.Description
        End If
        On Error GoTo 0
        lbl.Caption = fields(i)
        lbl.Left = labelLeft
        lbl.Top = topPos
        lbl.Width = labelWidth
        lbl.Height = 250

        topPos = topPos + rowHeight
    Next

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
    Next
    StripSqlComments = out
End Function


Sub Fail(message)
    WScript.Echo "ERROR: " & message
    WScript.Quit 1
End Sub
