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

        ElementTree.fromstring(package.read("content.xml"))
        ElementTree.fromstring(package.read("META-INF/manifest.xml"))
        macro = package.read("Basic/Standard/MediaCatalog.xml").decode("utf-8")
        menu = package.read(
            "Configurations2/menubar/menubar.xml"
        ).decode("utf-8")
        assert "Sub ResolveSelectedRows()" in macro
        assert "Media Catalog" in menu
        assert "ResolveSelectedRows?language=Basic" in menu
        assert "CheckMediaCatalogConfiguration?language=Basic" in menu

    xlsx = PROJECT_ROOT / "excel" / "MediaCatalog_template.xlsx"
    assert xlsx.exists()
    with zipfile.ZipFile(xlsx) as package:
        # The checked-in XLSX is intentionally a data-only source. install.cmd
        # uses Excel itself to create the macro-enabled root XLSM.
        assert "xl/vbaProject.bin" not in package.namelist()

    assert (PROJECT_ROOT / "scripts" / "build_excel_template.vbs").exists()
    assert (PROJECT_ROOT / "scripts" / "build_excel_template.ps1").exists()
    print("PASS: ODS macro/menu package and Excel template builder contract")


if __name__ == "__main__":
    main()
