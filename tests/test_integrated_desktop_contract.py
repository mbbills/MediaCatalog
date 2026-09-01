from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent


def main():
    excel = (
        PROJECT_ROOT / "excel" / "MediaCatalog_Excel_Module.bas"
    ).read_text(encoding="utf-8")
    calc = (
        PROJECT_ROOT / "calc" / "MediaCatalog_Calc_Module.txt"
    ).read_text(encoding="utf-8")

    assert "' MediaCatalog Excel 2016 module v0.4.0" in excel
    assert "Public Sub ResolveSelectedRows()" in excel
    assert '"Resolve Selected Rows", "ResolveSelectedRows"' in excel
    assert '"resolve_rows.py"' in excel
    assert "' MediaCatalog LibreOffice Calc module v0.4.0" in calc
    assert "Sub ResolveSelectedRows()" in calc
    assert '"resolve_rows.py"' in calc
    assert "Sub RemoveSelectedUPCERows()" in calc
    assert "Sub OpenSelectedUPCOnBluRay()" in calc
    assert "Sub CheckMediaCatalogConfiguration()" in calc

    print("PASS: Excel and Calc integrated resolver contracts")


if __name__ == "__main__":
    main()
