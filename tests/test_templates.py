import zipfile
from pathlib import Path
from xml.etree import ElementTree


PROJECT_ROOT = Path(__file__).resolve().parent.parent


def main():
    ods = PROJECT_ROOT / "MediaCatalog_template.ods"
    assert ods.exists(), "Calc ODS template is missing"

    with zipfile.ZipFile(ods) as package:
        names = set(package.namelist())
        required = {
            "mimetype",
            "content.xml",
            "styles.xml",
            "META-INF/manifest.xml",
            "Basic/Standard/MediaCatalog.xml",
            "Configurations2/menubar/menubar.xml",
        }
        assert required.issubset(names)
        assert package.infolist()[0].filename == "mimetype"
        assert package.infolist()[0].compress_type == zipfile.ZIP_STORED
        assert (
            package.read("mimetype")
            == b"application/vnd.oasis.opendocument.spreadsheet"
        )

        content = package.read("content.xml")
        ElementTree.fromstring(content)
        ElementTree.fromstring(package.read("META-INF/manifest.xml"))
        macro = package.read("Basic/Standard/MediaCatalog.xml").decode("utf-8")
        menu = package.read(
            "Configurations2/menubar/menubar.xml"
        ).decode("utf-8")
        assert b">Blu-ray.com Title<" in content
        assert b">IMDb Title<" in content
        assert b">Release Title<" not in content
        assert b">Title<" not in content
        assert "MediaCatalog LibreOffice Calc module v0.4.1" in macro
        embedded_source = macro.split("<![CDATA[", 1)[1].rsplit("]]>", 1)[0]
        standalone_source = (
            PROJECT_ROOT / "calc" / "MediaCatalog_Calc_Module.txt"
        ).read_text(encoding="utf-8")
        assert embedded_source == standalone_source
        assert "Sub ResolveSelectedRows()" in macro
        assert 'Array("Blu-ray.com Title", "Release Title"' in macro
        assert 'Array("IMDb Title", "Title")' in macro
        assert "Media Catalog" in menu
        assert "ResolveSelectedRows?language=Basic" in menu
        assert "CheckMediaCatalogConfiguration?language=Basic" in menu

    xlsx = PROJECT_ROOT / "excel" / "MediaCatalog_template.xlsx"
    assert xlsx.exists()
    with zipfile.ZipFile(xlsx) as package:
        # The checked-in XLSX is intentionally a data-only source. install.cmd
        # uses Excel itself to create the macro-enabled root XLSM.
        assert "xl/vbaProject.bin" not in package.namelist()
        assert not any(
            name.startswith("xl/tables/") for name in package.namelist()
        )
        workbook_xml = package.read("xl/workbook.xml")
        assert b"absPath" not in workbook_xml
        assert b"fileRecoveryPr" not in workbook_xml
        worksheet_xml = b"\n".join(
            package.read(name)
            for name in package.namelist()
            if name.startswith("xl/") and name.endswith(".xml")
        )
        assert b"Blu-ray.com Title" in worksheet_xml
        assert b"IMDb Title" in worksheet_xml
        assert b"Release Title" not in worksheet_xml

    builder_path = PROJECT_ROOT / "scripts" / "build_excel_template.vbs"
    powershell_builder_path = (
        PROJECT_ROOT / "scripts" / "build_excel_template.ps1"
    )

    assert builder_path.exists()
    assert powershell_builder_path.exists()

    builder_source = builder_path.read_text(encoding="utf-8")

    assert "codeModule.AddFromString workbookCode" in builder_source
    assert "codeModule.DeleteLines" not in builder_source
    assert (
        "OpenTextFile(workbookCodeFile, 1, False, 0)"
        in builder_source
    )
    assert "excel.EnableEvents = False" in builder_source
    assert 'WScript.Echo "Created " & outputWorkbook' not in builder_source

    print("PASS: ODS macro/menu package and Excel template builder contract")


if __name__ == "__main__":
    main()
