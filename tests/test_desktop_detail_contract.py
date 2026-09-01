#!/usr/bin/env python3
"""Offline contract checks for Excel/Calc release-detail integration."""

from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
excel = (ROOT / "excel" / "MediaCatalog_Excel_Module.bas").read_text(
    encoding="utf-8-sig"
)
calc = (ROOT / "calc" / "MediaCatalog_Calc_Module.txt").read_text(
    encoding="utf-8-sig"
)

assert calc.splitlines()[2] == "' MediaCatalog LibreOffice Calc module v0.3.0"

headers = (
    "UPC\tBlu-ray.com URL\tRelease Title\tIMDb URL\tIMDb ID\tTitle\t"
    "Year\tRuntime\tTitle Type\tSeason\tStatus / Error\tStudio\tBlu-ray Year\t"
    "Blu-ray Runtime\tContent Rating\tPhysical Release Date\tDisc Format\t"
    "Video Codec\tResolution\tAspect Ratio\tDisc Count / Capacities"
)

assert (ROOT / "excel" / "HEADERS.txt").read_text(encoding="utf-8").strip() == headers
assert (ROOT / "calc" / "HEADERS.txt").read_text(encoding="utf-8").strip() == headers

assert '"Enrich Selected Blu-ray Details", "EnrichSelectedBluRayDetails"' in excel
assert '"Resolve Selected UPCs with BRdC", "ResolveSelectedUPCsWithBRdC"' in excel
assert '"Resolve Selected UPCs with UPCdb", "ResolveSelectedUPCsWithUPCdb"' in excel
assert '"Resolve Selected UPCs with BarcodeLookup.com", "ResolveSelectedUPCsWithBarcodedCom"' in excel
assert '"Open UPC on BarcodeLookup.com (No API)", "ResolveSelectedUPCsWithBarcodedComNoAPI"' in excel
assert "Public Sub ResolveSelectedUPCsWithBRdC()" in excel
assert "Public Sub ResolveSelectedUPCsWithUPCdb()" in excel
assert "Public Sub ResolveSelectedUPCsWithBarcodedCom()" in excel
assert "Public Sub ResolveSelectedUPCsWithBarcodedComNoAPI()" in excel
assert 'ThisWorkbook.FollowHyperlink "https://www.barcodelookup.com/"' in excel
assert '"upcitemdb_lookup.py"' in excel
assert '"UPCItemDB", "UPCdb", "upcitemdb_lookup.py", False, True' in excel
assert '"barcodelookup_lookup.py", False, True' in excel
assert '(normalized = "[NOT FOUND]" Or normalized = "[AMBIGUOUS]")' in excel
assert "Public Sub EnrichSelectedBluRayDetails()" in excel
assert '"bluray_details.py"' in excel
assert "GetWindowlessPythonCommand" in excel
assert "WriteIsoDateIfBlank" in excel
assert "WriteTextIfBlank sheet.Cells(resultRow, blurayYearColumn), fields(7)" in excel
assert 'FindHeaderColumn(sheet, Array("Blu-ray.com URL", "Blu-ray URL", "Release URL"))' in excel
assert 'fields(4) = "No UPC/EAN, URL OK"' in excel
assert 'Case "CANCELLED"' in excel
assert '"Cancelled: " & CStr(cancelled)' in excel
assert "RunCommandAndWait(commandLine, 43200" in excel
assert "sheet.Cells(CLng(rowValue), imdbIdColumn)" in excel

assert "Sub ResolveSelectedUPCs()" in calc
assert "Sub ResolveSelectedUPCsWithBRdC()" in calc
assert "Sub ResolveSelectedUPCsWithUPCdb()" in calc
assert "Sub ResolveSelectedUPCsWithBarcodedCom()" in calc
assert "Sub ResolveSelectedUPCsWithBarcodedComNoAPI()" in calc
assert 'shellExecute.execute("https://www.barcodelookup.com/", "", 0)' in calc
assert "Sub ResolveSelectedUPCsLegacy()" in calc
assert "Sub EnrichSelectedBluRayDetails()" in calc
assert '"bluray_lookup_excel.py"' in calc
assert '"upcitemdb_lookup.py"' in calc
assert '"UPCItemDB", "UPCdb", "upcitemdb_lookup.py", False, True' in calc
assert '"barcodelookup_lookup.py", False, True' in calc
assert '(normalized = "[NOT FOUND]" Or normalized = "[AMBIGUOUS]")' in calc
assert '"bluray_details.py"' in calc
assert "SetCalcCellHyperlink" in calc
assert "SetCalcIsoDateIfBlank" in calc
assert "SetCalcTextIfBlank oSheet.getCellByPosition(blurayYearColumn, rowNum), fields(7)" in calc
assert 'FindCalcHeaderColumn(oSheet, Array("Blu-ray.com URL", "Blu-ray URL", "Release URL"))' in calc
assert 'fields(4) = "No UPC/EAN, URL OK"' in calc
assert 'Case "CANCELLED"' in calc
assert '"Cancelled: " & CStr(cancelled)' in calc
assert "For tries = 1 To 432000" in calc
assert "oSheet.getCellByPosition(imdbIdColumn, rowNum)" in calc
assert "getCellByPosition(1, rowNum)" not in calc

assert (ROOT / "scripts" / "bluray_details.py").is_file()
assert (ROOT / "scripts" / "upcitemdb_lookup.py").is_file()
assert (ROOT / "scripts" / "barcodelookup_lookup.py").is_file()
assert (ROOT / "scripts" / "job_progress.py").is_file()
assert "# MediaCatalog portable v0.3.0" in (ROOT / "README.md").read_text(
    encoding="utf-8"
)

print("PASS: Excel/Calc Blu-ray detail integration contract tests")
