Option Explicit

Dim fso, projectRoot, sourceWorkbook, outputWorkbook
Dim moduleFile, workbookCodeFile, excel, workbook, component, codeModule
Dim textStream, workbookCode

If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript build_excel_template.vbs PROJECT_ROOT"
    WScript.Quit 2
End If

Set fso = CreateObject("Scripting.FileSystemObject")
projectRoot = fso.GetAbsolutePathName(WScript.Arguments(0))
sourceWorkbook = fso.BuildPath(fso.BuildPath(projectRoot, "excel"), "MediaCatalog_template.xlsx")
moduleFile = fso.BuildPath(fso.BuildPath(projectRoot, "excel"), "MediaCatalog_Excel_Module.bas")
workbookCodeFile = fso.BuildPath(fso.BuildPath(projectRoot, "excel"), "ThisWorkbook_Code.txt")
outputWorkbook = fso.BuildPath(projectRoot, "MediaCatalog_template.xlsm")

If Not fso.FileExists(sourceWorkbook) Then Fail "Source workbook not found: " & sourceWorkbook
If Not fso.FileExists(moduleFile) Then Fail "VBA module not found: " & moduleFile
If Not fso.FileExists(workbookCodeFile) Then Fail "ThisWorkbook code not found: " & workbookCodeFile

Set textStream = fso.OpenTextFile(workbookCodeFile, 1, False, 0)
workbookCode = textStream.ReadAll
textStream.Close

On Error Resume Next
Set excel = CreateObject("Excel.Application")
If Err.Number <> 0 Then Fail "Excel 2016 could not be started: " & Err.Description
On Error GoTo 0

excel.Visible = False
excel.DisplayAlerts = False
excel.EnableEvents = False

' Excel's COM server can return from CreateObject before its object model
' (including the Workbooks collection) has finished initializing. Opening
' a workbook during that window fails with the generic "Unable to get the
' Open property of the Workbooks class" -- a timing race, not a problem
' with the file. Give it a moment, and retry a couple of times before
' giving up.
WScript.Sleep 1000

Dim openAttempt, openError, openErrorNumber
Dim maxOpenAttempts
maxOpenAttempts = 3
On Error Resume Next
For openAttempt = 1 To maxOpenAttempts
    Err.Clear
    Set workbook = excel.Workbooks.Open(sourceWorkbook)
    openErrorNumber = Err.Number
    openError = Err.Description
    If openErrorNumber = 0 Then Exit For
    If openAttempt < maxOpenAttempts Then WScript.Sleep 1500
Next

If openErrorNumber <> 0 Then
    On Error GoTo 0
    excel.Quit
    Fail "The source workbook could not be opened after " & maxOpenAttempts & _
        " attempt(s) (error " & openErrorNumber & "): " & openError
End If

Set component = workbook.VBProject.VBComponents.Import(moduleFile)
If Err.Number <> 0 Then
    Dim importError
    importError = Err.Description
    workbook.Close False
    excel.Quit
    Fail "VBA import failed. Programmatic VBA access may be blocked: " & importError
End If
On Error GoTo 0

Set codeModule = workbook.VBProject.VBComponents("ThisWorkbook").CodeModule
codeModule.AddFromString workbookCode

If fso.FileExists(outputWorkbook) Then fso.DeleteFile outputWorkbook, True
workbook.SaveAs outputWorkbook, 52

Set codeModule = Nothing
Set component = Nothing
workbook.Close False
Set workbook = Nothing
excel.Quit
Set excel = Nothing

WScript.Quit 0


Sub Fail(message)
    WScript.Echo "ERROR: " & message
    WScript.Quit 1
End Sub
