# Changelog

## Unreleased

- **Renamed "Blu-ray.com URL"/"Blu-ray.com Title" to "Blu-ray URL"/"Blu-ray
  Title" everywhere** (dropping ".com" entirely, not just the period), for
  consistency with "Blu-ray Year"/"Blu-ray Runtime" (which never had
  ".com") and so the Access field names -- previously "Blu-ray com
  URL"/"Blu-ray com Title", a workaround for Access's no-period rule --
  can become character-for-character identical to the canonical
  spreadsheet header text, with no separate name mapping needed anywhere
  in the project. Updated in sync:
  - `excel/MediaCatalog_template.xlsx` and `MediaCatalog_template.ods`
    header rows (B1/C1 -> C1/D1 shifted from the earlier Inventory Number
    insertion, text only changed here, no columns move), `excel/
    HEADERS.txt`, `calc/HEADERS.txt`.
  - `excel/MediaCatalog_Excel_Module.bas` and `calc/
    MediaCatalog_Calc_Module.txt`: every `FindHeaderColumn`/
    `FindCalcHeaderColumn` alias array now leads with the new canonical
    name, with `"Blu-ray.com URL"`/`"Blu-ray.com Title"` kept as a legacy
    alias (same pattern as the existing `"Release Title"`/`"DVD Title"`
    aliases) so spreadsheets built before this change keep working
    unmodified. Guard/error messages, diagnostic labels, and the
    top-of-file column comment updated to match.
  - `access/schema.sql`: `[Blu-ray com URL]`/`[Blu-ray com Title]` ->
    `[Blu-ray URL]`/`[Blu-ray Title]`.
  - `access/build_access_database.vbs`: `accessFieldNames`/
    `displayLabels` both updated (and are now literally identical arrays
    for every field, not just these two).
  - `access/MediaCatalog_Access_Module.bas`: every `Controls(...)`/
    `Fields(...)` reference in both `ResolveCurrentRecord` (Phase 2) and
    `ResolveSelectedRecords` (Phase 3) updated -- these would otherwise
    have broken silently (a field-name lookup miss, not a compile error)
    after the schema rename.
  - `access/README.md`'s field-name mapping table rewritten as a
    historical note: it no longer applies, since every Access field name
    now matches its spreadsheet header exactly.
  - Root `README.md`'s column table and the two scenario-table rows that
    named the column.
  - Regenerated the checked-in `.xlsx`/`.ods` templates' header cells
    directly (raw XML/zip surgery, not a rebuild) and re-embedded the
    updated Calc module source into the `.ods`'s Basic macro storage,
    verified byte-identical to the standalone `.txt` file.
  - Updated every test asserting the old exact strings:
    `tests/test_desktop_detail_contract.py`,
    `tests/test_integrated_desktop_contract.py`,
    `tests/test_templates.py`, `tests/test_access_schema_names.py`.
  - Scanned `schema.sql`, `HEADERS.txt`, and both live template header
    rows for any other Access-illegal character (period, `!`, `` ` ``,
    `[`, `]`) -- found none; this was the only occurrence.
  - `resolve_rows.py`'s "Manual Blu-ray.com URL" source-attribution label
    and every "Blu-ray.com" mention describing the actual website (menu
    items, comments, status text) were deliberately left unchanged --
    only the spreadsheet-header/Access-field-name usage was in scope.

- Fixed two more gaps found in `"16 Blocks (DVD  2006  Widescreen) NEW"`:
  - `extract_year()` only recognized a year that was the sole content of a
    bracket group (e.g. `"(2006)"`); a year embedded alongside other words
    in the same group (`"(DVD 2006 Widescreen)"`) was missed entirely.
    Broadened to find a year token anywhere inside a bracket group, while
    still requiring it be inside *some* bracket group -- a bare year
    outside brackets is too easily part of a real title ("300", "1984",
    "2001: A Space Odyssey") to search for freely.
  - A trailing, ALL-CAPS condition/listing tag ("NEW") is seller
    metadata, not title text, and wasn't stripped. Added as a
    case-**sensitive** strip specifically so it can never touch a
    legitimately title-cased title that happens to end the same way --
    confirmed `"Something New (2006)"` is untouched while `"16 Blocks
    (DVD 2006 Widescreen) NEW"` correctly loses the trailing tag.
    Deliberately conservative: only "NEW" is added for now, since that's
    the only condition tag seen in a real title so far; more (USED,
    SEALED, MINT, etc.) can be added the same way if/when they show up in
    a real failing title.
  - Extended `tests/test_title_cleanup.py` (18 cases total), including an
    explicit case-sensitivity regression guard.

- Added a "- Part N" fallback candidate to `title_match_candidates()` for
  titles like `"Dragons: Riders of Berk - Part 1"`, where "Part N" marks
  how a season was split across discs/volumes rather than being part of
  IMDb's actual title (`"Dragons: Riders of Berk"`). This is deliberately
  a *fallback* candidate only, tried after the full title, not a
  `clean_release_name()` strip -- "Part N" is sometimes genuinely part of
  the real IMDb title (`"Harry Potter and the Deathly Hallows: Part 1"`
  and `"Part 2"` are both real, separate, exact IMDb titles), so
  unconditionally removing it would have broken those. Verified both
  directions: the Dragons case now gets `"...Riders of Berk"` as its
  second candidate, while Harry Potter's exact `"...Part 1"` title stays
  the untouched first candidate. Also confirmed as correct with no
  changes needed: `"Father of the Bride 2 (1995)"` (repeat sanity check)
  and `"Madonna: The Video Collection 93:99 (1993-1999)"` -- the
  colon-embedded "93:99" is real product-name text, not clutter, and
  doesn't trigger the "Title 2: Subtitle" candidate rewrite (which
  requires a literal digit "2" before the colon). Extended
  `tests/test_title_cleanup.py` with a candidate-list check for both
  directions of the Part-N fallback.

- Added two more missing packaging phrases to `imdb_matcher`'s
  packaging-word list, found in a third batch of failing titles: "Full
  Screen"/"Wide Screen" as space-separated two-word forms (only the
  single-word "fullscreen"/"widescreen" was previously recognized), and
  "Snap case" as a physical case-type descriptor (same category as the
  earlier DigiBook/DigiPack/SteelBook additions). Fixes `"Jarhead DVD
  (Full Screen) (2005)"`, `"National Lampoon's Vacation DVD (Snap case)
  (1983)"`, and `"National Lampoon's European Vacation DVD (Snap case)
  (1985)"`. Extended `tests/test_title_cleanup.py` (17 cases total).

- Fixed two more real-title cleanup gaps found in a second batch of
  failing Blu-ray.com titles:
  - `detect_season()` required a season ordinal to be immediately
    followed by "Season" (e.g. "Eighth Season"), so "The Eighth and
    Final Season" (a common final-season phrasing) wasn't recognized as
    a season release at all. Now tolerates an optional "and Final"/"and
    Last" between the ordinal and "Season".
  - "Director's Cut" and "Warner Archive Collection" -- both real,
    recurring Blu-ray.com packaging/label phrases -- weren't in the
    packaging-word list, so parentheticals containing them survived
    cleanup intact (e.g. `"Dark City Blu-ray (Director's Cut) (1998)"`
    stayed as `"Dark City Blu-ray (Director's Cut)"` instead of "Dark
    City"). This also indirectly fixed a related case: a bare trailing
    format word (the earlier "Blu-ray"/"DVD" fix) only got stripped when
    it was already at the very end of the string, so it was missed
    whenever a still-unrecognized parenthetical came after it. Once the
    parenthetical is now correctly recognized as clutter and removed,
    the format word ends up trailing and gets stripped as before -- no
    separate architecture change was needed once the packaging-word gap
    itself was closed. Added "extended cut", "theatrical cut", "unrated
    cut", and "final cut" alongside "director's cut" since they're the
    same category of clutter and equally unambiguous.
  - Extended `tests/test_title_cleanup.py` with both new cases (14
    total).
  - Confirmed NOT fixable by cleanup, and deliberately left alone --
    these are title *alias* mismatches or source-data typos, not
    clutter, and no safe general regex removes them without risking
    false positives on real titles:
    - `"Highlander 2: Renegade Version (1991)"` -- "Renegade Version" is
      real branding for a specific home-video cut of the film, but
      IMDb's title is "Highlander II: The Quickening"; there's no
      generic "cut name" text to strip that gets you there.
    - `"Clue: The Movie (1985)"` -- IMDb's title is just "Clue"; but
      other real titles legitimately keep an exact ": The Movie" suffix
      on IMDb (e.g. "Digimon: The Movie"), so blindly stripping it would
      trade one failure mode for another.
    - `"Kung Fu Panda Holiday Special (2010)"` -- same shape of problem;
      IMDb's title is "Kung Fu Panda Holiday" without "Special", but
      "Special" is too common a word to safely strip in general.
    - `"Enterprise - The Complete Second Season (2002-2003)"` -- season
      detection itself now works correctly (series "Enterprise", season
      2), but IMDb's actual primary title is "Star Trek: Enterprise";
      Blu-ray.com's shortened franchise name won't exact-match it. Same
      class of problem as the two above.
    - `"Gentlemen Prefer Blonds"` -- this is a typo in the source title
      itself (the real film and IMDb title is "Gentlemen Prefer
      Blondes"); not a cleanup defect at all, and exact-title matching
      can't be expected to absorb a misspelling.
    - `"Father of the Bride 2 (1995)"` cleaned correctly with no changes
      needed -- included as a verified sanity check, not a bug report.
    All five of the alias/typo cases above share a fix shape that's
    different from everything fixed so far: they'd need either an
    explicit alias table (known alternate/marketing title -> canonical
    IMDb title) or a fuzzy-matching fallback, not another packaging-word
    or regex tweak. Worth scoping as its own follow-up rather than
    folding into more cleanup-regex patches.

- Fixed `imdb_matcher.clean_release_name()` leaving product/packaging
  clutter in the search title, causing exact-title IMDb matching to miss
  releases that should have resolved cleanly. Found via a batch of real
  failing titles walked through with the user. Specific fixes:
  - A bare format word left dangling outside any parentheses (e.g. the
    "Blu-ray" in `"Cars 2 Blu-ray (Blu-ray + DVD) (2011)"`, which only
    had the parenthesized "(Blu-ray + DVD)" stripped, leaving "Cars 2
    Blu-ray") is now also stripped. This was the single highest-impact
    defect, hit by 8 of the 10 reported titles (Cars 2; X-Men Origins:
    Wolverine; The Angry Birds Movie; Toy Story; Toy Story 2; Schindler's
    List; The Sum of All Fears; 16 Blocks; Beverly Hills Cop).
  - Packaging-word matching inside a parenthesized group used substring
    containment (`word in inner`), so "disc" (a packaging word) matched
    inside unrelated text like "Disclosure" and could wipe out a
    legitimate parenthetical subtitle. Now matched with word/phrase
    boundaries.
  - A parenthesized box-set year range, e.g. "(2008-2013)", was not
    recognized as a year by either the "drop this group, it's just a
    year" check or by `extract_year()`; both now accept a range and use
    its first year.
  - "DigiBook" and "DigiPack" (real Blu-ray.com packaging terms) were
    missing from the packaging-word list and so survived cleanup intact,
    e.g. `"Schindler's List DVD (DigiBook) (1993)"`.
  - "(The) Complete Series" box-set suffixes (e.g. "Breaking Bad: The
    Complete Series") were not recognized by `detect_season()`, which
    only matches "Season N" phrasing, and so were left in the search
    title verbatim. `clean_release_name()` now strips a trailing
    "Complete Series" suffix (and everything after it, packaging clutter
    included), which resolves to the plain series title -- not a full
    season-aware fix, but sufficient for the exact-title match to find
    the series' own IMDb entry.
  - Confirmed NOT a bug and deliberately left alone: a disc covering more
    than one title at once (e.g. "Toy Story and Toy Story 2 DVD (Disc 3:
    Supplemental Features) (1995)") has no single correct IMDb target: it
    still won't exact-match anything after cleanup, which is correct --
    there's no title fix that could make this resolve to one IMDb ID.
  - Also confirmed but explicitly out of scope for this pass: a title
    collision ambiguity for short, reused titles (e.g. "Defiance", which
    IMDb lists more than once across unrelated works) could plausibly
    cause the year/vote ranking in `find_matches()` to pick the wrong
    entry. This needs verification against the real `imdb.sqlite`, which
    this sandbox does not have; flagged for follow-up rather than
    guessed at blind.
  - Added `tests/test_title_cleanup.py`, a database-free regression test
    covering all 10 reported titles plus the "Disclosure" false-positive
    and the multi-title-disc no-match case. The existing
    `tests/test_imdb_matcher.py` regression cases were re-verified by
    hand (string-level, no database available in this environment) to
    still clean identically to before this change.

- Fixed `scripts/build_excel_template.ps1` failing with "Unable to get the
  Open property of the Workbooks class" when opening
  `excel/MediaCatalog_template.xlsx`. This is the standard symptom of
  Protected View blocking an invisible (`Application.Visible = False`)
  Excel automation session: Excel can't show the "this file was
  downloaded from the Internet" banner, so `Workbooks.Open` fails outright
  instead of prompting. Files can pick up that Mark-of-the-Web flag when
  the repository is obtained as a downloaded/extracted ZIP rather than a
  plain `git clone`. The script now runs `Unblock-File` on the Excel
  source files it touches before invoking the VBScript builder. Confirmed
  live that `Unblock-File` itself is unavailable on the user's Windows 7
  install (PowerShell 2.0 predates that cmdlet), and that
  `$ErrorActionPreference = "Stop"` turned the resulting
  `CommandNotFoundException` into a hard failure despite
  `-ErrorAction SilentlyContinue` (it can't apply to a command that was
  never resolved). Now detects whether `Unblock-File` exists and falls
  back to removing the `Zone.Identifier` alternate data stream directly
  when it doesn't, wrapped in try/catch so a missing stream or a
  filesystem without ADS support isn't fatal either. Not yet
  re-verified against real Excel -- please re-run `install.cmd` and
  confirm.
- Confirmed live that MOTW/Protected View was not the (or not the only)
  cause of the Excel template build failure above: unblocking the source
  files made no difference and the identical error persisted. Audited
  `excel/MediaCatalog_template.xlsx`'s internal XML (shared strings,
  styles/dxf counts, content types, relationships, row spans) for
  corruption from the earlier Inventory Number column-insertion surgery
  -- all internally consistent, so that's very likely not the cause
  either. `scripts/build_excel_template.vbs` opened the workbook
  immediately after `CreateObject("Excel.Application")` with no pause for
  Excel's COM server to finish initializing, which is the most commonly
  reported real-world cause of this exact generic error when the file and
  trust settings are otherwise fine. Added a one-second settle delay plus
  a 3-attempt retry (1.5s apart) around `Workbooks.Open`, and the error
  message now includes `Err.Number` for further diagnosis if it still
  fails. Confirmed live: this fixed the template build.
- Rebuilt `excel/MediaCatalog_template.xlsx`'s Inventory Number column
  insertion using `openpyxl` instead of hand-edited raw XML. The
  hand-edited version was internally consistent enough to pass zip/XML
  well-formedness checks, but comparing it cell-by-cell against a proper
  rebuild turned up real defects the manual surgery had introduced: the
  old header cell for UPC kept the leftmost-column border style instead
  of switching to the interior-column style used by every other header,
  and (caught while building this replacement, not in the shipped file)
  a naive column-width shift loses widths on any column whose original
  `<col>` entry spanned more than one column (e.g. Status/Error and
  Studio shared one `<col min="11" max="12">` entry; a per-letter copy
  only carries the width to the first of the two). The regenerated file
  was verified against the original: all 22 headers, conditional
  formatting range (now correctly `C2:C1000`), and every column width
  match the original 21-column layout shifted by one, and it passes the
  full test suite (`tests/test_templates.py`,
  `tests/test_desktop_detail_contract.py`,
  `tests/test_integrated_desktop_contract.py`). Not yet verified against
  real Excel.
- Added `install.cmd`/`install_media_catalog.py --skip-database`, to
  reconfigure `settings.ini` or rebuild a template without touching
  `data/imdb.sqlite` at all. The already-existing "don't redownload/rebuild
  if unchanged" logic was verified correct and given regression test
  coverage (it previously had none).
- Access front end phase 3: a "Resolve Selected Records" button (on a new
  `frmMediaCatalogTools` form) that batch-resolves a multi-row selection
  made in a new `frmMediaCatalogDatasheet` Datasheet-view form -- the
  closest Access equivalent to Excel/Calc's spreadsheet row selection.
  Records are correlated to the resolver's response by their table ID
  (AutoNumber primary key) rather than a spreadsheet row position. Phase
  2 confirmed working live in the meantime; see the "Phase 3" section of
  `access/README.md`. Not yet verified against real Access.
- Access front end phase 2: a "Resolve Current Record" button running the
  same integrated resolver (`scripts/resolve_rows.py`) Excel/Calc's
  "Resolve Selected Rows" uses, against the current record only (Access
  has no spreadsheet-style row selection). See
  `access/MediaCatalog_Access_Module.bas` and the "Phase 2" section of
  `access/README.md`. Not yet verified against real Access.
- Started the Access 2019 front end (phase 1): a flat `MediaCatalog` table
  schema (`access/schema.sql`) and a basic bound data-entry/browse form,
  built via `access/build_access_database.vbs` since `.accdb` is a binary
  format. No resolver integration yet. See `access/README.md`.
- Added an `Inventory Number` column (new column A in both templates) as the
  user's manual physical-copy/location identifier, e.g. `BR37`. `Spine Tag`
  is an older name for the same concept; existing catalogs that already use
  that header are unaffected, since no macro reads this column by name.
- Fixed the Windows installer path passed to the Excel template builder when
  the project folder contains spaces or the batch-file directory ends in a
  quoted backslash.
- Repaired the checked-in Excel source template, removed its invalid table and
  repair metadata, and hardened the builder for Excel 2016 COM cleanup.
- Made older partial `settings.ini` files inherit newly shipped defaults, so a
  preserved Python path does not hide later IMDb dataset settings.
- Corrected the integrated resolver's progress counts for complete, partial,
  and needs-review results.
- Made the legacy Excel `ResolveSelectedUPCs` entry point run the integrated
  workflow, matching Calc, and remove stale Media Catalog menus by caption.
- Made an Excel-template build failure return installer failure instead of
  continuing to a misleading `Installation complete` message.

## v0.4.1

- Renamed the standard column C header from **Release Title** to
  **Blu-ray.com Title** so the physical-release title's source is explicit.
- Renamed the standard column F header from **Title** to **IMDb Title** so it is
  clearly distinguished from the Blu-ray.com product title.
- Updated the Calc ODS template, Excel XLSX source template, embedded Calc
  module, standalone Excel and Calc modules, header files, installer identity,
  tests, and documentation to use the new canonical names.
- Retained **Release Title** and **Title** as legacy header aliases so existing
  v0.4.0 and earlier catalogs continue to work without renaming their columns.

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
