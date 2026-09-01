# MediaCatalog portable v0.4.0

MediaCatalog is a local physical-media catalog workflow for Excel 2016 and
LibreOffice Calc. One integrated command can begin with a UPC/EAN, an exact
Blu-ray.com release URL, an IMDb URL or ID, a release title, or a canonical
title and fill every field it can resolve confidently.

The workflow keeps physical-product identity separate from canonical content
identity:

- Blu-ray.com supplies the particular DVD, Blu-ray, or UHD release and its disc
  specifications.
- IMDb supplies the canonical movie or television identity used for Jellyfin
  naming and other library work.
- An IMDb ID entered in the sheet remains the authoritative correction when
  Blu-ray.com is missing, ambiguous, or wrong.

## Standard columns

| Column | Field |
|---|---|
| A | UPC |
| B | Blu-ray.com URL |
| C | Release Title |
| D | IMDb URL |
| E | IMDb ID |
| F | Title |
| G | Year |
| H | Runtime |
| I | Title Type |
| J | Season |
| K | Status / Error |
| L | Studio |
| M | Blu-ray Year |
| N | Blu-ray Runtime |
| O | Content Rating |
| P | Physical Release Date |
| Q | Disc Format |
| R | Video Codec |
| S | Resolution |
| T | Aspect Ratio |
| U | Disc Count / Capacities |

The macros locate columns by their row-1 labels, but A through U is the
supported standard layout.

## Quick installation on Windows

Run:

```bat
install.cmd
```

The installer:

1. finds an installed Python 3.8 or newer;
2. if necessary, downloads and installs Python 3.8.10, the final official
   Python release with a Windows 7 installer;
3. creates the private, Git-ignored `settings.ini`;
4. checks for `title.basics.tsv.gz`, `title.episode.tsv.gz`, and
   `title.ratings.tsv.gz`;
5. downloads missing or invalid datasets from IMDb;
6. builds or refreshes `data/imdb.sqlite`; and
7. offers to use Excel 2016 to build `MediaCatalog_template.xlsm` with the
   VBA module and **Media Catalog** menu embedded.

Dataset downloads and the database build are safe to rerun. A failed or
cancelled build does not replace a working database.

The Excel template builder temporarily enables programmatic VBA-project access
for the current user, starts Excel while it is hidden, creates the XLSM, and
restores the previous security setting. Close Excel before allowing that step.

At the end, the installer offers to open this README.

## Templates and menus

### LibreOffice Calc

`MediaCatalog_template.ods` is ready to use. It contains:

- the standard A-through-U catalog sheet;
- the current MediaCatalog Basic module;
- a document-specific **Media Catalog** menu; and
- menu assignments for the integrated resolver and the retained individual
  tools.

Open the ODS from the project root and enable its document macros. Use **Save
As** for your working catalog, but keep the workbook in the MediaCatalog project
folder so its macros can find `settings.ini`, `scripts`, and `data`.

LibreOffice officially supports saving document-scoped menu customizations in
templates. The ODS stores its menu and macros in the document rather than
modifying every Calc installation.

### Excel 2016

`excel\MediaCatalog_template.xlsx` is intentionally a data-only source
workbook. The XLSX format cannot contain VBA, so it does **not** contain the
MediaCatalog menu.

Running `install.cmd` creates this usable macro-enabled template in the
project root:

```text
MediaCatalog_template.xlsm
```

When it opens with macros enabled, its `Workbook_Open` handler installs the
**Media Catalog** menu on Excel's **Add-ins** tab. Closing the workbook removes
that menu.

If automated template creation is blocked, use the manual procedure:

1. Open `excel\MediaCatalog_template.xlsx` in Excel.
2. Press **Alt+F11**.
3. Choose **File → Import File** and import
   `excel\MediaCatalog_Excel_Module.bas`.
4. Double-click **ThisWorkbook** and paste
   `excel\ThisWorkbook_Code.txt`.
5. Save the workbook in the project root as an Excel Macro-Enabled Workbook
   named `MediaCatalog_template.xlsm`.
6. Reopen it and enable macros.

The `Attribute VB_Name` line belongs to the imported BAS file. Do not paste
that line directly into the VBA code editor.

## Integrated resolution

Select one or more data rows and run:

```text
Media Catalog → Resolve Selected Rows
```

The resolver processes each row according to the strongest available input:

| Starting information | Result |
|---|---|
| UPC only | UPC → Blu-ray.com release → disc details and structured IMDb link → local IMDb metadata |
| Blu-ray.com URL only | Exact release page → disc details and structured IMDb link → local IMDb metadata |
| IMDb ID or URL only | Exact local IMDb lookup |
| Title only | Local IMDb title matcher; physical-release fields remain unresolved |
| UPC plus IMDb ID | UPC supplies physical-release data; the entered IMDb ID supplies canonical content data |
| Blu-ray.com URL plus UPC | The supplied release URL is authoritative |

Identifier precedence is:

1. entered IMDb ID;
2. entered IMDb URL containing an ID;
3. Blu-ray.com's structured `imdb_icon` title link; and
4. local title/year matching.

The parser deliberately ignores IMDb links appearing in reviews, forum text, or
technical discussions. Only Blu-ray.com's structured title link is eligible.

If the entered IMDb ID differs from Blu-ray.com's link, MediaCatalog retains the
entered ID and records a warning such as:

```text
OK - Blu-ray + IMDb; Manual IMDb override; Blu-ray.com links tt1234567
```

IMDb-derived Title, Year, Runtime, Title Type, and Season are refreshed from
the winning IMDb ID. Existing physical-detail cells are filled only when blank,
preserving manually reviewed Studio, release date, format, codec, and related
values.

A title or IMDb ID identifies content, not a particular disc release.
MediaCatalog therefore does not guess a physical edition when only IMDb or
title information is supplied.

## Retained individual tools

The integrated command is the normal workflow. These commands remain available
for targeted retries and comparison:

- **Resolve Selected UPCs with BRdC**
- **Resolve Selected UPCs with UPCdb**
- **Resolve Selected UPCs with BarcodeLookup.com**
- **Open UPC on BarcodeLookup.com (No API)**
- **Enrich Selected Blu-ray Details**
- **Lookup IMDb for Selected Rows**
- **Remove UPC-E Rows in Selection**
- **Open Selected UPC on Blu-ray.com**
- **Check Configuration**

The automated Barcode Lookup resolver requires a paid subscription and API key.
Its free website remains available through the manual command. UPCItemDB's free
trial endpoint is throttled according to its published request limits.

## Progress, throttling, and cancellation

Blu-ray.com rows display a shared progress window with current item, progress,
elapsed time, estimated time remaining, and a **Cancel** button. Completed rows
are returned even when the remainder is cancelled.

The configured Blu-ray.com delay is applied between selected rows that actually
used the network, not before the first row. Title-only and IMDb-ID-only rows use
the local SQLite database and do not initialize a Blu-ray.com session.

Do not click or edit the workbook while a desktop macro is writing its results.
Use the progress window's Cancel button instead.

## IMDb datasets and database

Official resources:

- [IMDb non-commercial dataset documentation](https://developer.imdb.com/non-commercial-datasets/)
- [IMDb daily dataset directory](https://datasets.imdbws.com/)
- [title.basics.tsv.gz](https://datasets.imdbws.com/title.basics.tsv.gz)
- [title.episode.tsv.gz](https://datasets.imdbws.com/title.episode.tsv.gz)
- [title.ratings.tsv.gz](https://datasets.imdbws.com/title.ratings.tsv.gz)

The three compressed files belong in `data/source`. To rebuild manually:

```bat
python scripts\build_imdb_database.py --force
```

The builder writes a temporary `.building` database and replaces the live
database only after a complete successful import. Databases from v0.2.4 and
later remain compatible with v0.4.0.

## Configuration

Copy `settings.example.ini` to `settings.ini` when installing manually.
The local file is ignored by Git because it can contain machine-specific paths
and API credentials.

The project uses only Python's standard library. No `pip install` command is
required. Python's `-E` option prevents inherited `PYTHON*` environment
variables from contaminating the helper processes.

## Macro security

Do not globally enable every macro.

For Excel, add the MediaCatalog project directory as a trusted location:

```text
File → Options → Trust Center → Trust Center Settings → Trusted Locations
```

For Calc, use **Tools → Options → LibreOffice → Security → Macro Security** and
trust the project location or approve the document macros according to your
security policy.

## Tests

Copy `settings.example.ini` to `settings.ini`, then run the self-contained
tests individually from the project root. GitHub Actions runs the same set on
Python 3.8.

The database-backed matcher regression test additionally requires
`data/imdb.sqlite`:

```bat
python tests\test_imdb_matcher.py
```

## Project layout

```text
MediaCatalog/
├── install.cmd
├── MediaCatalog_template.ods
├── settings.example.ini
├── settings.ini                       # local and Git-ignored
├── README.md
├── CHANGELOG.md
├── excel/
│   ├── MediaCatalog_Excel_Module.bas
│   ├── ThisWorkbook_Code.txt
│   ├── MediaCatalog_template.xlsx     # data-only XLSM source
│   └── HEADERS.txt
├── calc/
│   ├── MediaCatalog_Calc_Module.txt
│   └── HEADERS.txt
├── scripts/
│   ├── resolve_rows.py
│   ├── install_media_catalog.py
│   ├── build_excel_template.ps1
│   ├── build_excel_template.vbs
│   ├── bluray_lookup_excel.py
│   ├── bluray_details.py
│   ├── imdb_lookup.py
│   ├── imdb_matcher.py
│   ├── build_imdb_database.py
│   └── retained provider/progress helpers
├── data/
│   └── source/
├── tests/
└── logs/
```

## Backups and portability

For a complete portable backup, preserve:

- the MediaCatalog project directory;
- the working catalog workbook;
- `data/imdb.sqlite`; and
- `CHANGELOG.md`.

The SQLite database and catalog files can be copied between supported Windows
and LibreOffice installations. The downloaded IMDb datasets are useful for
future rebuilds but can be downloaded again.
