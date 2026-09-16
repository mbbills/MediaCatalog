Option Explicit

' Builds access\MediaCatalog.accdb from source-controlled text:
'   - schema.sql defines the MediaCatalog table (executed verbatim as DDL).
'   - MediaCatalog_Access_Module.bas is imported as a standard VBA module
'     (Phase 2: ResolveCurrentRecord; Phase 3: ResolveSelectedRecords --
'     see that file's own header comments).
'   - This script then builds three forms via the classic
'     CreateForm/CreateControl Access automation API:
'       frmMediaCatalog          -- Phase 1/2: single-record browse/entry
'                                    form, with the "Resolve Current
'                                    Record" button.
'       frmMediaCatalogDatasheet -- Phase 3: the same table in Datasheet
'                                    view, for multi-row selection. No
'                                    controls to create -- Datasheet view
'                                    auto-generates its grid from
'                                    RecordSource.
'       frmMediaCatalogTools     -- Phase 3: a small form with the
'                                    "Resolve Selected Records" button.
'                                    Datasheet view has no header/footer
'                                    section to put a button on, so the
'                                    button lives here instead and reads
'                                    the selection from
'                                    frmMediaCatalogDatasheet by name.
'
' Importing a VBA module via automation requires the same "Trust access to
' the VBA project object model" setting as creating the form's controls
' (see access/README.md) -- both go through Access's VBE object model.
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

Dim fso, projectRoot, accessDir, schemaFile, moduleFile, outputDatabase
Dim textStream, schemaSql
Dim access

Const acTextBox = 109
Const acLabel = 100
Const acCommandButton = 104
Const acDetail = 0
Const acForm = 2
Const acSaveYes = 1
' Form.DefaultView's own enumeration (unrelated to, and coincidentally
' sharing the value 2 with, the acForm AcObjectType constant above used
' by DoCmd.Close/Rename): 0 Single Form, 1 Continuous Forms, 2 Datasheet.
Const acFormViewDatasheet = 2

If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript build_access_database.vbs PROJECT_ROOT"
    WScript.Quit 2
End If

Set fso = CreateObject("Scripting.FileSystemObject")
projectRoot = fso.GetAbsolutePathName(WScript.Arguments(0))
accessDir = fso.BuildPath(projectRoot, "access")
schemaFile = fso.BuildPath(accessDir, "schema.sql")
moduleFile = fso.BuildPath(accessDir, "MediaCatalog_Access_Module.bas")
outputDatabase = fso.BuildPath(accessDir, "MediaCatalog.accdb")

If Not fso.FileExists(schemaFile) Then Fail "Schema file not found: " & schemaFile
If Not fso.FileExists(moduleFile) Then Fail "VBA module not found: " & moduleFile

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

On Error Resume Next
Dim component
Set component = access.VBE.ActiveVBProject.VBComponents.Import(moduleFile)
If Err.Number <> 0 Then
    Dim importError
    importError = Err.Description
    access.CloseCurrentDatabase
    access.Quit
    Fail "VBA module import failed. Programmatic VBA access may be blocked: " & importError
End If
On Error GoTo 0

BuildBrowseForm access
BuildDatasheetForm access
BuildToolsForm access

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
    Dim accessFieldNames, displayLabels, i, topPos, rowHeight
    Dim labelLeft, labelWidth, textLeft, textWidth
    Dim frm, ctl, lbl, finalName

    ' Must exactly match schema.sql's field names (see the Access-name <->
    ' spreadsheet-header mapping table in access/README.md), NOT the
    ' spreadsheet header text: ControlSource below binds a control to a
    ' field by its real Access name, and "Blu-ray.com URL"/"Blu-ray.com
    ' Title" are invalid Access field names (Access field names cannot
    ' contain a period). displayLabels carries the original, prettier
    ' spreadsheet header text for what the user actually sees on the form.
    accessFieldNames = Array( _
        "Inventory Number", "UPC", "Blu-ray com URL", "Blu-ray com Title", _
        "IMDb URL", "IMDb ID", "IMDb Title", "Year", "Runtime", _
        "Title Type", "Season", "Status / Error", "Studio", _
        "Blu-ray Year", "Blu-ray Runtime", "Content Rating", _
        "Physical Release Date", "Disc Format", "Video Codec", _
        "Resolution", "Aspect Ratio", "Disc Count / Capacities" _
    )
    displayLabels = Array( _
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

    For i = LBound(accessFieldNames) To UBound(accessFieldNames)
        On Error Resume Next
        Set ctl = app.CreateControl(frm.Name, acTextBox, acDetail)
        If Err.Number <> 0 Then
            Fail "CreateControl (textbox) failed for field '" & accessFieldNames(i) & "': " & Err.Description
        End If
        On Error GoTo 0
        ctl.ControlSource = accessFieldNames(i)
        ctl.Left = textLeft
        ctl.Top = topPos
        ctl.Width = textWidth
        ctl.Height = 250

        On Error Resume Next
        Set lbl = app.CreateControl(frm.Name, acLabel, acDetail)
        If Err.Number <> 0 Then
            Fail "CreateControl (label) failed for field '" & accessFieldNames(i) & "': " & Err.Description
        End If
        On Error GoTo 0
        lbl.Caption = displayLabels(i)
        lbl.Left = labelLeft
        lbl.Top = topPos
        lbl.Width = labelWidth
        lbl.Height = 250

        topPos = topPos + rowHeight
    Next

    ' Phase 2: "Resolve Current Record" runs the integrated resolver
    ' (MediaCatalog_Access_Module.bas, ResolveCurrentRecord) against this
    ' one open record -- Access has no spreadsheet-style row selection, so
    ' this is the equivalent of Excel/Calc's "Resolve Selected Rows" for a
    ' single record. The BRdC-only and single-row-IMDb-lookup commands are
    ' deferred to a later phase, same as Phase 1's form-only scope.
    Dim btn
    On Error Resume Next
    Set btn = app.CreateControl(frm.Name, acCommandButton, acDetail)
    If Err.Number <> 0 Then Fail "CreateControl (button) failed: " & Err.Description
    On Error GoTo 0
    btn.Caption = "Resolve Current Record"
    btn.OnClick = "=ResolveCurrentRecord()"
    btn.Left = textLeft
    btn.Top = topPos + 100
    btn.Width = textWidth
    btn.Height = 350

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


' Phase 3: the MediaCatalog table in Datasheet view, for multi-row
' selection (ResolveSelectedRecords reads this form's SelTop/SelHeight).
' No CreateControl calls needed -- Datasheet view generates its own grid
' of columns directly from RecordSource, unlike the Detail-section forms
' above.
Sub BuildDatasheetForm(app)
    Dim frm, finalName

    On Error Resume Next
    Set frm = app.CreateForm()
    If Err.Number <> 0 Then Fail "CreateForm (datasheet) failed: " & Err.Description
    On Error GoTo 0

    frm.RecordSource = "MediaCatalog"
    frm.DefaultView = acFormViewDatasheet
    frm.Caption = "MediaCatalog (Datasheet)"
    finalName = frm.Name

    On Error Resume Next
    app.DoCmd.Close acForm, finalName, acSaveYes
    If Err.Number <> 0 Then Fail "Saving the datasheet form failed: " & Err.Description
    On Error GoTo 0

    On Error Resume Next
    app.DoCmd.Rename "frmMediaCatalogDatasheet", acForm, finalName
    If Err.Number <> 0 Then Fail "Renaming the datasheet form failed: " & Err.Description
    On Error GoTo 0
End Sub


' Phase 3: a small form holding only the "Resolve Selected Records"
' button. Datasheet view (frmMediaCatalogDatasheet, above) has no
' header/footer section to put a button on, so this separate form is the
' button's home; ResolveSelectedRecords looks up the datasheet form by
' name rather than via Screen.ActiveForm, which at the moment this
' button's OnClick runs is this tools form itself, not the datasheet the
' user selected rows in.
Sub BuildToolsForm(app)
    Dim frm, btn, finalName

    On Error Resume Next
    Set frm = app.CreateForm()
    If Err.Number <> 0 Then Fail "CreateForm (tools) failed: " & Err.Description
    On Error GoTo 0

    On Error Resume Next
    Set btn = app.CreateControl(frm.Name, acCommandButton, acDetail)
    If Err.Number <> 0 Then Fail "CreateControl (Resolve Selected Records button) failed: " & Err.Description
    On Error GoTo 0
    btn.Caption = "Resolve Selected Records"
    btn.OnClick = "=ResolveSelectedRecords()"
    btn.Left = 100
    btn.Top = 100
    btn.Width = 3200
    btn.Height = 350

    frm.Caption = "MediaCatalog Tools"
    finalName = frm.Name

    On Error Resume Next
    app.DoCmd.Close acForm, finalName, acSaveYes
    If Err.Number <> 0 Then Fail "Saving the tools form failed: " & Err.Description
    On Error GoTo 0

    On Error Resume Next
    app.DoCmd.Rename "frmMediaCatalogTools", acForm, finalName
    If Err.Number <> 0 Then Fail "Renaming the tools form failed: " & Err.Description
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
