# Changelog

## v0.4.0

- Added **Resolve Selected Rows**, a single Excel/Calc workflow that can begin
  with a UPC/EAN, exact Blu-ray.com URL, IMDb URL/ID, release title, or
  canonical title and populate every confidently resolvable field.
- Added structured IMDb-link extraction from individual Blu-ray.com release
  pages. Only the page's `imdb_icon` title link is accepted; IMDb links in
  reviews and unrelated prose are ignored.
- Established identifier precedence: entered IMDb ID, entered IMDb URL,
  Blu-ray.com's structured IMDb link, then local title/year matching.
- Preserved entered IMDb IDs when they conflict with Blu-ray.com and recorded
  an explicit manual-override warning containing Blu-ray.com's linked ID.
- Made a supplied Blu-ray.com release URL authoritative for physical-release
  data, while existing physical-detail cells remain protected from overwrite.
- Reused one downloaded Blu-ray.com release page for its title, disc details,
  and IMDb identity instead of issuing separate enrichment requests.
- Added offline integrated-resolver regression tests for UPC-first, URL-first,
  IMDb-ID override, title-only, and invalid-manual-ID paths.
- Added `install.cmd`, which finds Python 3.8+, installs Python 3.8.10 when
  necessary on Windows, downloads and validates the three required IMDb daily
  datasets, builds `imdb.sqlite`, and offers to create the Excel template.
- Added a comment-preserving Python installation helper and safe dataset/database
  refresh behavior.
- Added an Excel 2016 template builder. Excel itself imports the reviewed BAS
  source and ThisWorkbook events into a root-level macro-enabled XLSM; the
  builder temporarily enables VBA-project access and restores its prior value.
- Added `MediaCatalog_template.ods` with the standard A-through-U layout,
  embedded v0.4.0 Basic module, and document-specific **Media Catalog** menu.
- Added Calc parity commands for UPC-E removal, opening the selected UPC on
  Blu-ray.com, and checking configuration.
- Added GitHub Actions coverage on Python 3.8 plus package-contract tests for
  the ODS macro/menu and Excel template-building inputs.
- The workbook columns and IMDb SQLite schema remain unchanged; an existing
  v0.2.4-or-newer database does not require rebuilding.

## v0.3.0

- Added a shared native progress window for BRdC, UPCItemDB, Barcode Lookup,
  and Blu-ray.com release-detail batches. It displays the current item,
  completed/total rows, result counts, elapsed time, and estimated remaining
  time.
- Added a **Cancel** button with cooperative cancellation. The current network
  request is allowed to finish, completed results are preserved and imported,
  unfinished rows are marked cancelled without changing their title/detail
  fields, and rerunning the command resumes incomplete work.
- Made configured provider delays interruptible so cancellation does not have
  to wait through UPCItemDB's full rate-limit pause.
- Added explicit `CANCELLED` result handling and summary counts to both Excel
  and Calc.
- Extended desktop network-batch waiting from thirty minutes to a twelve-hour
  safety limit; users can stop long jobs through the progress window.
- Converted fatal provider/configuration failures into per-row error responses
  when possible, avoiding an uninformative missing-response timeout.
- Added `[progress]` settings for showing the window and keeping it on top,
  plus offline regression coverage for progress state and partial-result
  preservation.
- Kept the twenty-one-column workbook layout and IMDb schema unchanged; no
  spreadsheet migration or IMDb database rebuild is required.

## v0.2.13

- Changed the BRdC title and details URL fallback to run after either `[NOT FOUND]` or
  `[AMBIGUOUS]`. Entering a definitive individual Blu-ray.com release URL and
  rerunning **Resolve Selected UPCs with BRdC** now replaces either marker with
  the URL page's title and hyperlink.
- Added regression coverage using UPC `027616857804` for the ambiguous-result
  override path.
- Kept the v0.2.11 twenty-one-column workbook layout and IMDb schema unchanged;
  no database rebuild is required.

## v0.2.12

- Added a trailing-`3D` IMDb fallback: the exact title is tried first to protect
  genuine titles such as `Piranha 3D`, then a no-match retry removes the likely
  Blu-ray.com format marker. The catalog's Release Title remains unchanged.
- Added `Blu-ray 3D` detection to release-detail parsing. Combo sets retain all
  detected physical formats, such as `Blu-ray 3D + Blu-ray Disc`.
- Expanded season detection to recognize numeric, cardinal, and ordinal forms
  through season twenty, including `Season Two`, `Season One`, `Second Season`,
  and `The Complete Seventh Season`.
- Made worded season releases search IMDb using the base series title before
  the season phrase, restrict results to television series/miniseries that
  contain the requested season, and write the parsed number to the Season
  column.
- Added offline regression coverage for the trailing-3D fallback, Blu-ray 3D disc
  details, and `Sherlock: Season Two` series matching.
- Kept the v0.2.11 twenty-one-column workbook layout and IMDb schema unchanged;
  no database rebuild is required.

## v0.2.11

- Added **Blu-ray.com URL** immediately after UPC as standard column B, moving
  Release Title and all later fields one column to the right.
- Made the active Excel and Calc workflows resolve columns from normalized
  row-1 labels instead of fixed offsets, including all IMDb outputs.
- Updated the BRdC resolver to store both the matching release URL and linked
  release title on a successful UPC/EAN lookup.
- Added an explicit URL backup: after a Blu-ray.com UPC/EAN miss (or when no
  usable UPC/EAN is present), a validated individual Blu-ray.com release URL
  can supply the release title and details. The status is recorded exactly as
  `No UPC/EAN, URL OK`.
- Kept UPC/EAN lookup authoritative when both a valid UPC/EAN and a URL are
  supplied; the URL is used only after an exact-code miss.
- Updated the Excel template, Calc/Excel headers, documentation, and offline
  regression tests for the twenty-one-column layout.
- No IMDb database rebuild is needed when upgrading from v0.2.10.

## v0.2.10

- Added the release version to line 3 of the Calc module header and covered it
  with a regression check for future releases.
- Reworked Blu-ray.com summary parsing to identify production year, runtime,
  content rating, and physical release date by value instead of fixed position.
- Preserved television production-year ranges such as `2012-2013` in the
  Blu-ray Year column by writing that field as text in Excel and Calc.
- Ignored optional summary labels such as `Season 7` and
  `1 Movie, 2 Cuts` without shifting later fields.
- Recognized both singular `Disc` and plural `Discs` sections and inferred
  `Single disc` when the singular section provides a disc format but no count.
- Used `Original aspect ratio` as a fallback when the release page does not
  supply a separate `Aspect ratio` value.
- Added regression coverage for The Incredibles (`786936244250`), Curious
  George: The Complete Seventh Season (`025192213984`), and A Bug's Life
  (`786936217896`).
- Kept the workbook layout and v0.2.4 IMDb schema unchanged; no database rebuild
  is needed when upgrading from v0.2.9.

## v0.2.9

- Corrected the Barcode Lookup guidance: its terms prohibit automated use of
  the public website and of a free API test account.
- Added `paid_subscription = false` as a safety gate. The automated
  `ResolveSelectedUPCsWithBarcodedCom` helper now refuses to run until a paid
  subscription is explicitly confirmed and an API key is configured.
- Added `ResolveSelectedUPCsWithBarcodedComNoAPI` to Excel and Calc as a manual,
  non-scraping helper for one selected row.
- The manual helper copies the UPC to the clipboard, opens BarcodeLookup.com,
  and leaves all catalog cells untouched for user review.
- Added the manual command to Excel's menu and documented its Calc menu
  assignment and terms-based limitations.
- Kept the workbook layout and v0.2.4 IMDb schema unchanged; no database rebuild
  is needed when upgrading from v0.2.8.

## v0.2.8

- Added `ResolveSelectedUPCsWithBarcodedCom` to Excel and Calc as a third
  selected-row UPC/EAN resolver using BarcodeLookup.com's documented v3 API.
- Added the corresponding Excel menu command and documented Calc menu
  assignment.
- Preserved the established fallback rules: Barcode Lookup retries blank,
  `[NOT FOUND]`, and `[AMBIGUOUS]` title cells without overwriting a normal
  title.
- Kept Barcode Lookup titles as plain text so they cannot be mistaken for
  Blu-ray.com release-detail links.
- Added API-key, endpoint, delay, and timeout settings with a clear no-key
  configuration error and no attempted request.
- Added duplicate-code caching, safe HTTP error handling that does not expose
  the API key, and an offline parser/TSV contract test.
- Kept the workbook layout and v0.2.4 IMDb schema unchanged; no database rebuild
  is needed when upgrading from v0.2.7.

## v0.2.7

- Renamed the explicit Blu-ray.com resolver macro to
  `ResolveSelectedUPCsWithBRdC` in Excel and Calc, while retaining
  `ResolveSelectedUPCs` as a compatibility alias.
- Added `ResolveSelectedUPCsWithUPCdb` to both desktop front ends, backed by a
  shared Python 3.8-compatible UPCItemDB helper.
- Added separate Excel menu entries and documented Calc menu assignments for
  the two resolver sources.
- Made column B source-neutral as **Release Title** while continuing to
  recognize existing **Blu-ray.com Title**, **DVD Title**, and **UPCItemDB
  Name** headers.
- Made UPCdb a practical fallback: it retries blank cells plus `[NOT FOUND]`
  and `[AMBIGUOUS]` markers, but preserves normal existing titles.
- Kept BRdC titles hyperlinked to their matching release pages and wrote UPCdb
  titles as plain text so they cannot be mistaken for Blu-ray.com detail URLs.
- Added free- and paid-plan UPCItemDB configuration, duplicate-code caching,
  conservative free-plan throttling, row-level API errors, and offline tests.
- Kept the v0.2.4 IMDb database schema unchanged; no database rebuild is needed
  when upgrading from v0.2.6.

## v0.2.6

- Added a separate **Enrich Selected Blu-ray Details** command so release-page
  requests are made only for explicitly selected rows.
- Added ten optional physical-release fields: studio, Blu-ray year, Blu-ray
  runtime, content rating, physical release date, disc format, video codec,
  resolution, aspect ratio, and disc count/capacities.
- Added a shared Python release-page parser with validation, duplicate-request
  caching, ISO release-date output, and offline DVD/Blu-ray regression tests.
- Preserved existing detail-cell values. Rows with all ten detail fields already
  populated are skipped; partially populated rows fill only blank cells.
- Extended the Excel 2016 menu, VBA bridge, twenty-column template, and header
  file for the new separate enrichment request.
- Ported Calc's active UPC resolver from UPCItemDB to the shared Blu-ray.com
  helper, including release-title hyperlinks, and added the matching Calc
  enrichment macro and twenty-column header file.
- Retained the old Calc UPCItemDB procedure under the non-menu name
  `ResolveSelectedUPCsLegacy` as source reference only.
- Kept the v0.2.4 IMDb database schema unchanged; no database rebuild is needed
  when upgrading from v0.2.5.

## v0.2.5

- Added the individual Blu-ray.com release-page URL to successful UPC/EAN
  lookup results and made the returned title in column B a clickable hyperlink.
- Changed Excel's local IMDb bridge to use the companion `pythonw.exe`,
  suppressing the visible console window while retaining the existing polling,
  timeout, and response-file handling.
- Added validation and a clear configuration error when the windowless Python
  executable cannot be found.
- Kept the v0.2.4 IMDb database schema unchanged; no database rebuild is needed
  when upgrading from v0.2.4.

## v0.2.4

- Added `title.ratings.tsv.gz` as a required third IMDb dataset, with official
  direct-download links and placement instructions in the README.
- Added ratings import support to `build_imdb_database.py`; the generated
  SQLite database stores both `average_rating` and `num_votes`.
- Added schema-version validation with a clear `--force` rebuild instruction
  when an older `imdb.sqlite` is used.
- Changed year handling so zero- and one-year differences share the same
  confidence band.
- Added IMDb vote count as the popularity tie-breaker within a year-confidence
  band and title type. Average rating is retained for diagnostics but does not
  choose a match.
- Added offline ratings-import/ranking tests and a regression case ensuring
  `300 (2007)` resolves to IMDb `tt0416449`, the 2006 feature film.

## v0.2.3

- Preserved Blu-ray.com's full result title and parenthesized year in column B;
  the year continues to rank otherwise identical IMDb title matches.
- Added removal of a trailing `4K` product marker from the in-memory IMDb
  search title without modifying column B.
- Added one narrowly scoped fallback for titles formatted as
  `Title 2: Subtitle`: if the original exact-title search fails, retry as
  `Title: Subtitle`.
- Added `National Treasure 2: Book of Secrets 4K (2007)` to the IMDb matcher
  regression cases.
- Added official IMDb documentation, dataset-directory, and direct dataset
  download links to the README.

## v0.2.2 — withdrawn

- This revision removed the displayed year from the Blu-ray.com title. It was
  withdrawn because that year is useful evidence for IMDb disambiguation.

## v0.2.1

- Fixed `tests\test_imdb_matcher.py` when launched using the documented
  `python tests\test_imdb_matcher.py` command. The test now adds the shared
  `scripts` directory to its import path before importing `config` and
  `imdb_matcher`.
- Added an explicit missing-database check so the test does not create an empty
  `data\imdb.sqlite` when the real database has not yet been copied or built.

## v0.2.0

- Added an Excel 2016 front end for Windows 7.
- Added a persistent `Media Catalog` menu under Excel's Add-ins tab.
- Preserved selected-row/batch UPC and IMDb processing.
- Replaced UPCItemDB lookup in the Excel workflow with Blu-ray.com's exact
  12-digit UPC and 13-digit EAN database searches.
- Search order is DVD first, followed by Blu-ray/UHD when the DVD catalog has
  no exact-code result.
- Added an explicit selected-row command for removing rows whose symbology
  normalizes exactly to `UPC_E`; UPC-E is never expanded or guessed.
- Preserved Blu-ray.com's returned release title in column B as an audit field.
- Retained the authoritative IMDb-ID correction workflow from v0.1.2.
- Targeted the Python helper code to Python 3.8.10, the last official Python
  installer supporting Windows 7.
- Added a configuration check and a manual Blu-ray.com search command.
- Added a ready-to-use `.xlsx` catalog template with UPC and IMDb ID columns
  formatted as text.
- Retained the LibreOffice Calc v0.1.2 module as a legacy reference.

## v0.1.2

- Added an authoritative IMDb-ID correction workflow.
- Existing IMDb IDs are no longer skipped during IMDb lookup.
- If column D is blank, MediaCatalog performs normal automatic matching.
- If column D contains an IMDb ID, that ID is resolved exactly and all
  IMDb-derived fields are refreshed from the local database.
- A full IMDb title URL pasted into column D is accepted and normalized to
  its `tt...` ID.
- Invalid or unknown IMDb IDs are reported in the Status / Error column
  instead of falling back to a heuristic match.
- Batch IMDb lookup continues to work across selected rows.

## v0.1.1

- Added selected-range/batch IMDb enrichment.
- Added one summary dialog per batch.

## v0.1.0

- Clean portable baseline.
- UPCItemDB lookup.
- Local IMDb SQLite matching.
- LibreOffice Calc bridge.
