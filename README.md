# MediaCatalog portable v0.3.0

MediaCatalog is a local physical-media catalog workflow for Excel 2016 and
LibreOffice Calc. It retains the local IMDb database and authoritative IMDb-ID
correction behavior developed in v0.1.2.

The desktop workflow is:

1. Scan or paste a disc UPC into Excel.
2. Resolve the exact 12-digit UPC or 13-digit EAN through Blu-ray.com (BRdC),
   with UPCItemDB (UPCdb) and BarcodeLookup.com available as separate fallback
   commands.
3. Preserve the matching Blu-ray.com release URL beside the UPC and write the
   returned release title. BRdC titles and URLs are clickable; UPCdb and
   Barcode Lookup titles remain plain text.
4. On a separate selected-row request, collect the chosen physical-release
   details from that individual release page.
5. Clean the release title only in memory for matching.
6. Match it against the local IMDb SQLite database.
7. Write canonical IMDb metadata back into the selected rows.

The standard workbook columns are:

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

## What is and is not included

This package includes the complete Python source, Excel VBA source, current Calc
module, settings, tests, documentation, and changelog.

It intentionally does not include:

- IMDb's downloadable datasets;
- the generated `data/imdb.sqlite` database;
- a compiled `.xlam` add-in; or
- a personal catalog workbook.

An IMDb database built by MediaCatalog v0.2.4 or later on another computer is
portable. Databases from v0.2.3 or earlier must be rebuilt once because they do
not contain the ratings table used by the current matcher.

## Requirements

- Windows 7
- Excel 2016
- LibreOffice Calc when using the Calc front end
- Python 3.8.10
- Internet access for Blu-ray.com, UPCItemDB, and Barcode Lookup searches
- A paid Barcode Lookup API subscription and key to use its automated resolver
- `data/imdb.sqlite`, either copied from the earlier project or rebuilt locally

The Python code uses only the standard library. No `pip install` step is needed.
The standard Windows Python installer includes Tkinter, which MediaCatalog uses
for its shared progress and cancellation window.

## Project layout

```text
MediaCatalog/
├── settings.example.ini
├── settings.ini                 # local configuration; ignored by Git
├── README.md
├── CHANGELOG.md
├── excel/
│   ├── MediaCatalog_Excel_Module.bas
│   ├── ThisWorkbook_Code.txt
│   ├── MediaCatalog_template.xlsx
│   └── HEADERS.txt
├── calc/
│   ├── MediaCatalog_Calc_Module.txt
│   └── HEADERS.txt
├── scripts/
│   ├── bluray_details.py
│   ├── bluray_lookup_excel.py
│   ├── upcitemdb_lookup.py
│   ├── barcodelookup_lookup.py
│   ├── build_imdb_database.py
│   ├── config.py
│   ├── job_progress.py
│   ├── imdb_matcher.py
│   ├── imdb_lookup.py
│   └── imdb_lookup_calc.py
├── data/
│   └── source/
├── tests/
│   ├── test_bluray_details.py
│   ├── test_bluray_lookup.py
│   ├── test_upcitemdb_lookup.py
│   ├── test_barcodelookup_lookup.py
│   ├── test_desktop_detail_contract.py
│   ├── test_job_progress.py
│   ├── test_imdb_ratings.py
│   └── test_imdb_matcher.py
└── logs/
```

## 1. Install Python 3.8.10

Python 3.8.10 is the last official Python installer supporting Windows 7.
During installation, enable the option to add Python to `PATH`.

Confirm the executable in Command Prompt:

```bat
python --version
where python
```

Copy `settings.example.ini` to `settings.ini` before running MediaCatalog.
The local `settings.ini` is ignored by Git so API credentials and
machine-specific paths are not committed accidentally.

Edit the `[runtime]` section of `settings.ini` so `python` contains the full
path to `python.exe`. Forward slashes avoid INI escaping concerns:

```ini
[runtime]
python = C:/Python38/python.exe
```

Do not place command-line arguments in this setting. The Excel module adds
Python's `-E` option itself. For IMDb lookups, MediaCatalog automatically uses
the companion `pythonw.exe` installed beside `python.exe`; this runs the local
lookup without opening a console window.

## 2. Supply the local IMDb database

Official IMDb resources:

- [IMDb non-commercial dataset documentation](https://developer.imdb.com/non-commercial-datasets/)
- [IMDb daily dataset download directory](https://datasets.imdbws.com/)
- [Download `title.basics.tsv.gz`](https://datasets.imdbws.com/title.basics.tsv.gz)
- [Download `title.episode.tsv.gz`](https://datasets.imdbws.com/title.episode.tsv.gz)
- [Download `title.ratings.tsv.gz`](https://datasets.imdbws.com/title.ratings.tsv.gz)

The files are refreshed daily. MediaCatalog v0.2.4 and later require the three
downloads listed above; the other files in IMDb's directory are not required
by this version.

If another computer already has an IMDb database built by v0.2.4 or later:

```text
data/imdb.sqlite
```

copy that file into the same location in this project. SQLite database files
are portable between Windows 10 and Windows 7. A database built by v0.2.3 or
earlier must be rebuilt using the three current datasets.

To build or upgrade the database, place the three official IMDb daily datasets
here:

```text
data/source/title.basics.tsv.gz
data/source/title.episode.tsv.gz
data/source/title.ratings.tsv.gz
```

For a new database, run from the project root:

```bat
python scripts\build_imdb_database.py
```

To intentionally replace an existing database:

```bat
python scripts\build_imdb_database.py --force
```

The builder writes a temporary `.building` database and replaces the live
database only after a complete successful import. The ratings dataset supplies
`averageRating` and `numVotes`. MediaCatalog stores both values but uses vote
count—not average rating—as a matching-confidence signal.

## 3. Run the Python tests

After `data/imdb.sqlite` exists:

```bat
python tests\test_bluray_lookup.py
python tests\test_upcitemdb_lookup.py
python tests\test_barcodelookup_lookup.py
python tests\test_bluray_details.py
python tests\test_job_progress.py
python tests\test_desktop_detail_contract.py
python tests\test_imdb_ratings.py
python tests\test_imdb_matcher.py
```

The Blu-ray parser and ratings-ranking tests are offline. The IMDb matcher test
uses the rebuilt local database and should end with:

```text
PASS: 13 matcher regression cases
```

## 4. Create the Excel add-in

The package supplies reviewable VBA source instead of an opaque precompiled
macro binary. Create the add-in once in Excel 2016:

1. Open a blank Excel workbook.
2. Press `Alt+F11` to open the Visual Basic Editor.
3. Choose **File → Import File** and import
   `excel\MediaCatalog_Excel_Module.bas`.
4. In the Project pane, double-click **ThisWorkbook**.
5. Paste the contents of `excel\ThisWorkbook_Code.txt` into that object.
6. Choose **File → Save As**.
7. Save as **Excel Add-In (`*.xlam`)** with the name
   `MediaCatalog.xlam` in the project root beside `settings.ini`.
8. In Excel, open **File → Options → Add-ins**.
9. At the bottom, choose **Excel Add-ins**, click **Go**, then **Browse**.
10. Select `MediaCatalog.xlam` and enable it.

When upgrading an existing add-in to v0.3.0, open its VBA project, right-click
the old `MediaCatalogExcel` module and choose **Remove MediaCatalogExcel**. The
old module does not need to be exported because it remains in the earlier ZIP.
Then import the new `excel\MediaCatalog_Excel_Module.bas`, choose
**Debug → Compile VBAProject**, save the add-in, and reopen Excel. The v0.2.4
IMDb database remains compatible and does not need another rebuild.

Catalogs already using the v0.2.11 A-through-U headers need no column changes
for v0.3.0.

When the add-in loads, Excel's **Add-ins** tab contains a **Media Catalog**
menu. Its commands are:

- **Resolve Selected UPCs with BRdC**
- **Resolve Selected UPCs with UPCdb**
- **Resolve Selected UPCs with BarcodeLookup.com**
- **Open UPC on BarcodeLookup.com (No API)**
- **Enrich Selected Blu-ray Details**
- **Lookup IMDb for Selected Rows**
- **Remove UPC-E Rows in Selection**
- **Open Selected UPC on Blu-ray.com**
- **Check Configuration**

The legacy CommandBar menu was chosen because it works reliably in Excel 2016
without Ribbon XML. A custom Ribbon can be added later without changing the
lookup code.

## 4A. Update the LibreOffice Calc module and menu

The Calc module now uses the same Blu-ray.com, UPCItemDB, and Barcode Lookup
Python helpers as Excel. Replace the previous MediaCatalog module with the contents of
`calc\MediaCatalog_Calc_Module.txt`. If you created a custom **Media Catalog**
menu through **Tools → Customize**, add or update these macro assignments:

- **Resolve Selected UPCs with BRdC** → `ResolveSelectedUPCsWithBRdC`
- **Resolve Selected UPCs with UPCdb** → `ResolveSelectedUPCsWithUPCdb`
- **Resolve Selected UPCs with BarcodeLookup.com** → `ResolveSelectedUPCsWithBarcodedCom`
- **Open UPC on BarcodeLookup.com (No API)** → `ResolveSelectedUPCsWithBarcodedComNoAPI`
- **Enrich Selected Blu-ray Details** → `EnrichSelectedBluRayDetails`
- **Lookup IMDb for Selected Rows** → `LookupIMDbForCurrentRow`

For an existing Excel or Calc catalog from v0.2.10, insert one new blank column
at **B**, then paste the complete tab-separated line from the appropriate
`HEADERS.txt` file into row 1. The insert moves the existing Release Title and
all later data one column to the right without overwriting it.

The active macros locate fields by their normalized row-1 labels, not by fixed
column offsets. You can move columns if the labels remain unique and unchanged;
the A-through-U layout above is the supported standard and the supplied
template follows it.

## 5. Excel macro security

Do not globally enable every macro. Add the MediaCatalog project directory as
a trusted location:

```text
File
→ Options
→ Trust Center
→ Trust Center Settings
→ Trusted Locations
```

Then close and reopen Excel. Application-control or antivirus software can also
block Excel from starting Python; check its logs if the menu runs but Python
never starts.

## 6. Prepare a catalog workbook

Start with `excel\MediaCatalog_template.xlsx`, or create a normal `.xlsx` or
`.xlsm` workbook and paste the tab-separated header line from
`excel\HEADERS.txt` into row 1.

Format the UPC column as **Text** before scanning or pasting barcodes. This is
essential for leading zeroes. The macro can repair the common case where one
leading zero was lost, but text storage remains authoritative.

The add-in operates on the active workbook; the workbook itself does not need
to contain macros.

## 7. UPC/EAN title lookup

Select one cell, one row, or a contiguous group of rows and choose:

```text
Add-ins → Media Catalog → Resolve Selected UPCs with BRdC
```

The helper:

1. searches Blu-ray.com's dedicated DVD UPC/EAN field;
2. searches the Blu-ray/UHD database only when the DVD search has no result;
3. collapses duplicate edition records returning the same title;
4. writes the matching individual release-page URL into column B;
5. writes the exact returned title, including a displayed year, into column C
   and hyperlinks it to that URL;
6. if the UPC/EAN is absent from Blu-ray.com's index or returns conflicting
   titles, uses a valid URL already present in column B as the explicit backup
   and records
   `No UPC/EAN, URL OK` in column K;
7. writes `[AMBIGUOUS]` for conflicting UPC/EAN titles only when no usable URL
   fallback is supplied; and
8. writes `[NOT FOUND]` only when the exact-code searches return no title and
   no usable URL fallback is supplied.

It does not use the site's general title search and does not guess from partial
barcodes. Existing normal titles in column C are preserved. To refresh a row,
clear its column-C title and run **Resolve
Selected UPCs with BRdC** again for the selected rows.

The request delay and network timeout are configured in `settings.ini`:

```ini
[bluray]
request_delay_seconds = 0.75
timeout_seconds = 45
```

For rows marked `[NOT FOUND]` or `[AMBIGUOUS]` by BRdC, keep those rows
selected and choose:

```text
Add-ins → Media Catalog → Resolve Selected UPCs with UPCdb
```

The UPCdb command retries blank cells and those two unresolved markers, but it
does not overwrite a normal existing title. Successful UPCItemDB titles are
plain text because they are not Blu-ray.com release-page links. Column K records
the successful source as `UPC OK - BRdC` or `UPC OK - UPCdb`.

The free UPCItemDB endpoint currently permits 100 combined requests per day and
6 requests per minute. MediaCatalog therefore defaults to a 10.5-second delay
between distinct codes and caches duplicate codes in the selected batch. The
settings are:

```ini
[upcitemdb]
endpoint = https://api.upcitemdb.com/prod/trial/lookup?upc=
request_delay_seconds = 10.5
timeout_seconds = 45
user_key =
key_type = 3scale
```

Paid DEV/PRO accounts use `https://api.upcitemdb.com/prod/v1/lookup` and place
their key in `user_key`. See the official
[UPCItemDB getting-started documentation](https://www.upcitemdb.com/wp/docs/main/development/getting-started/),
[response schema](https://www.upcitemdb.com/wp/docs/main/development/responses/),
and [rate-limit documentation](https://www.upcitemdb.com/wp/docs/main/development/api-rate-limits/).

Barcode Lookup is available as another automated fallback for paid API
subscribers:

```text
Add-ins → Media Catalog → Resolve Selected UPCs with BarcodeLookup.com
```

Its requested macro name is `ResolveSelectedUPCsWithBarcodedCom`. Like UPCdb,
it retries blank cells, `[NOT FOUND]`, and `[AMBIGUOUS]`, preserves normal
existing titles, and writes successful titles as plain text. Column K records
`UPC OK - BarcodeLookup` on success.

Barcode Lookup's supported automated interface requires a paid API subscription
and key. After subscribing through its
[official API page](https://www.barcodelookup.com/api), edit:

```ini
[barcodelookup]
endpoint = https://api.barcodelookup.com/v3/products
api_key = paste_your_key_here
paid_subscription = true
request_delay_seconds = 0.75
timeout_seconds = 45
```

The helper refuses to run until both the key is present and
`paid_subscription = true` explicitly confirms the subscription. Barcode
Lookup documents a maximum of 100 API requests per minute; the default delay
remains below that ceiling and duplicate codes in one batch are cached. See its
[API documentation](https://www.barcodelookup.com/api-documentation).

Barcode Lookup's terms prohibit automated access to its public website and to
a free API test account. MediaCatalog therefore does not scrape the website.
For a compliant no-API manual lookup, select exactly one data row and choose:

```text
Add-ins → Media Catalog → Open UPC on BarcodeLookup.com (No API)
```

The underlying macro name is `ResolveSelectedUPCsWithBarcodedComNoAPI`. It
copies the selected row's UPC to the clipboard and opens BarcodeLookup.com's
home page. Paste the UPC into the site's search box and review the result
yourself. The command does not read the page, fill the Release Title, or modify
any catalog cell. See Barcode Lookup's
[terms and conditions](https://www.barcodelookup.com/terms-and-conditions),
especially sections 3 and 5.

`ResolveSelectedUPCs` remains as a compatibility alias for the BRdC resolver.
The menus use explicit names for the three automated resolvers and the manual
Barcode Lookup helper.

### Batch progress and cancellation

The three automated UPC resolvers and **Enrich Selected Blu-ray Details** open
one shared progress window. It reports completed and total rows, the current
UPC/EAN, result counts, elapsed time, and an estimated time remaining. The
estimate uses the average time of completed rows and can change as provider
response times vary.

Choose **Cancel** to stop cooperatively. A request already in progress may take
up to the provider's configured network timeout to return. MediaCatalog then
imports every completed result, leaves unfinished title/detail fields
resumable, records cancelled rows in the status field, and includes a
**Cancelled** count in the final summary. Running the same command again skips
completed rows and resumes work that is still incomplete.

Do not click in the spreadsheet while the progress window is open. The old
30-minute batch ceiling has been extended to a twelve-hour safety limit so
large collections can finish or be stopped with **Cancel**.

Progress-window behavior is configured in `settings.ini`:

```ini
[progress]
show_window = true
always_on_top = true
```

If Tkinter is unavailable, the batch still runs without the progress window.
You can verify Tkinter separately with `python -m tkinter`.

## 8. Blu-ray.com release-detail enrichment

After **Resolve Selected UPCs with BRdC** has supplied the title and URL,
select the rows to enrich and choose:

```text
Media Catalog → Enrich Selected Blu-ray Details
```

This is deliberately a separate request. It opens one individual release page
per selected release and fills blank cells in columns L through U:

- studio;
- Blu-ray.com's production year or year range and runtime;
- content rating;
- physical-media release date;
- disc format;
- video codec, resolution, and aspect ratio; and
- disc count and capacities.

The physical release date is stored as a real spreadsheet date and displayed
as `yyyy-mm-dd`. The production year and runtime in columns M and N are
Blu-ray.com values and remain separate from the IMDb values in columns G and H.
The Blu-ray Year field is text so television ranges such as `2012-2013` are
preserved rather than reduced to one year.

Release summary fields are identified by their contents rather than fixed
positions. Optional labels such as `Season 7` and `1 Movie, 2 Cuts` therefore
do not displace runtime, rating, or release date. Both `Disc` and `Discs`
sections are recognized. When a page omits `Aspect ratio` but supplies
`Original aspect ratio`, the original ratio is used as the fallback.
The Disc Format field records `Blu-ray 3D` whenever that format appears in the
release's Disc section. Combo sets can therefore show values such as
`Blu-ray 3D + Blu-ray Disc` or include 4K/DVD formats as applicable.

Existing detail cells are preserved. A row is skipped completely when all ten
detail fields are already populated; clear selected detail cells before
rerunning if you intentionally want them refreshed. Both Excel and Calc first
try the exact UPC/EAN. If Blu-ray.com has no indexed UPC/EAN result or returns
conflicting titles, they use the validated individual release URL in column B and write
`No UPC/EAN, URL OK` to the status field. A successful UPC/EAN result remains
authoritative even when column B already contains a different URL.

Some releases do not list every field. A successful request can therefore
leave an unavailable field blank. Prices, country, audio, subtitles, HDR,
packaging, popularity, and user ratings are intentionally outside this release.

## 9. UPC-E handling

UPC-E entries are removed only when a scanner export includes a barcode-type
column with a value that normalizes exactly to `UPC_E`. Supported header names
include `format`, `Barcode Type`, `Symbology`, and `Type`.

Select the imported rows and run:

```text
Add-ins → Media Catalog → Remove UPC-E Rows in Selection
```

Rows are deleted from bottom to top so remaining row numbers are not shifted
before they are inspected. The software never expands UPC-E to UPC-A and never
guesses symbology from barcode length alone.

The standard twenty-one-column catalog does not need a barcode-type column when UPCs
are scanned directly and the scanner is configured to supply UPC-A/EAN.

## 10. IMDb lookup and correction

After column C contains a release title, select the desired rows and choose:

```text
Add-ins → Media Catalog → Lookup IMDb for Selected Rows
```

On success the add-in fills columns D through K.

The release title in column C is preserved as an audit field. During
IMDb matching, MediaCatalog extracts a parenthesized year as a ranking hint and
removes product wording only from its in-memory search title. A trailing `4K`
format marker is ignored. A trailing `3D` is first searched exactly because it
can be part of a real title such as `Piranha 3D`; if that exact search misses,
MediaCatalog retries once without `3D`. If an exact title such as
`National Treasure 2: Book of Secrets` has no IMDb match, the matcher makes one
narrow fallback attempt as `National Treasure: Book of Secrets`. The fallback
is never tried when the original exact title already matches.

Season releases are detected before the IMDb search. Numeric and worded forms
such as `Season 2`, `Season Two`, `Second Season`, and `The Complete Seventh
Season` are recognized through season twenty. For a title such as
`Sherlock: Season Two`, MediaCatalog searches the base title `Sherlock`, limits
candidates to IMDb television series or miniseries that actually contain
season 2 in `title.episode`, and writes `2` into the Season column.

The displayed year is a useful hint but not an absolute truth. IMDb and release
catalogs can differ by one year because one may use an early screening while
another uses the wider theatrical release. MediaCatalog therefore treats a
zero- or one-year difference as the same confidence band. Within that band it
prefers the appropriate IMDb title type and then the candidate with the larger
IMDb vote count. A missing runtime and the exact year distance are later
tie-breakers. This resolves cases such as Blu-ray.com's `300 (2007)` selecting
IMDb's widely recognized `300 (2006)` instead of an obscure same-title entry.

Column E (`IMDb ID`) remains the single authoritative correction field. To
correct a wrong automatic match:

1. replace column E with the correct ID, such as `tt0114709`, or paste a full
   IMDb title URL;
2. select the row; and
3. run the IMDb lookup again.

The exact ID bypasses heuristic matching and refreshes the IMDb-derived fields.
Invalid or unknown IDs are reported rather than silently rematched.

The IMDb bridge is launched with `pythonw.exe`, so a console window should not
flash or remain visible while Excel processes the selected rows. The normal
`python.exe` setting remains authoritative; MediaCatalog derives and validates
the matching windowless executable automatically.

## Why Python is launched with `-E`

The Excel bridge retains the `-E` safeguard introduced during Calc development.
It tells Python to ignore inherited `PYTHON*` environment variables and prevents
another application's Python environment from contaminating the external
interpreter.

## Portability and backups

Paths under `[paths]` are relative to the project root unless made absolute.
Keep `MediaCatalog.xlam`, `settings.ini`, `scripts`, and `data` together.

For a complete portable backup, preserve:

- the project directory;
- the catalog workbook;
- `data/imdb.sqlite`; and
- the changelog.

The Calc module in v0.3.0 provides current BRdC, UPCdb, and Barcode Lookup front
ends. Its old
`ResolveSelectedUPCsLegacy` procedure is retained only as source-level
reference and is not assigned to the menu.
