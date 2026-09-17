# MediaCatalog Access front end

This is the Microsoft Access 2019 front end for MediaCatalog. It is new,
separate work alongside the Excel and LibreOffice Calc front ends -- see
the project handoff document, section 16.1, for the long-term intent
(shared Python resolver/IMDb backend, Access providing forms/queries/reports).

## What's implemented so far

**Phase 1** (schema and a basic form):

- **One flat table**, `MediaCatalog`, using the canonical column list from
  the handoff (section 7.2), with `Inventory Number` in place of the
  earlier `Spine Tag` name. No parent/child/box-set hierarchy yet -- that is
  explicitly deferred, per the handoff (section 7.3).
- **One basic data-entry/browse form**, `frmMediaCatalog`, bound to that
  table: a plain vertical stack of labeled text boxes, one per field, in
  the same order as the table.

See `schema.sql` for the exact field list and types, and the "Field types"
section below for the reasoning.

**Phase 2** (confirmed working live): a **"Resolve Current Record" button**
on `frmMediaCatalog` that runs the same integrated resolver Excel/Calc's
"Resolve Selected Rows" uses, against the one record currently open on the
form. See "Phase 2: Resolve Current Record" below for what it does, its
dependencies, and its limitations.

**Phase 3** (this round): a **"Resolve Selected Records" button**, on a
separate small `frmMediaCatalogTools` form, that resolves a multi-row
selection made in a new `frmMediaCatalogDatasheet` Datasheet-view form --
the closest Access equivalent to Excel/Calc's spreadsheet row selection.
See "Phase 3: Resolve Selected Records" below.

## What's explicitly NOT implemented yet

- The BRdC-only (Blu-ray.com-only) and single-row IMDb-lookup commands --
  deferred to later phases, one at a time.
- No import/export script (CSV, or from the Excel/Calc catalogs).
- No reports.
- No parent/child (box-set) schema.
- No friendly-title natural sort (handoff section 15.2) -- noted for a
  later phase.

## Files here

| File | Purpose |
|---|---|
| `schema.sql` | The canonical `CREATE TABLE` statement (Jet/ACE SQL). Hand-maintained source of truth for the schema. |
| `MediaCatalog_Access_Module.bas` | The integrated-resolver VBA module (`ResolveCurrentRecord`, Phase 2; `ResolveSelectedRecords`, Phase 3; and their helpers), ported from Excel's `ResolveSelectedRows`. Imported into the `.accdb`'s VBA project by the build script. |
| `build_access_database.vbs` | Automation script that creates a fresh `MediaCatalog.accdb`, runs `schema.sql` against it, imports `MediaCatalog_Access_Module.bas`, and builds `frmMediaCatalog`, `frmMediaCatalogDatasheet`, and `frmMediaCatalogTools`. |
| `MediaCatalog.accdb` | **Not included in this repository.** See below. |

## Why there's no `.accdb` file checked in yet

`.accdb` is a proprietary binary database format (the Access Database
Engine/ACE format), not something the development environment used to
prepare this phase can create, open, or validate: it has no Windows, no
Microsoft Access, no Access-compatible ODBC/OLEDB driver, and no tool
capable of authoring a genuine `.accdb`. Committing a hand-fabricated
binary claiming to be a working Access database would be worse than not
providing one.

This mirrors how this project already handles `MediaCatalog_template.xlsm`:
that file is also not hand-edited in source control. Instead,
`scripts/build_excel_template.vbs` and `build_excel_template.ps1` build it
from text sources (`excel/MediaCatalog_Excel_Module.bas`,
`excel/ThisWorkbook_Code.txt`) using Excel's own COM automation, on a
Windows machine with Excel installed. `access/build_access_database.vbs`
follows the same pattern for Access: run it once on a machine with Access
installed to produce `access/MediaCatalog.accdb`, and rerun it whenever
`schema.sql` changes.

### Building it

```
cscript access\build_access_database.vbs .
```

(run from the project root, or pass whatever path points at it). Requires
Access 2019 or a compatible version, installed and licensed on that
machine, with VBA/macro project access enabled (see below).

**This script has not been run against a real copy of Access.** It was
written by reasoning about the well-established, decades-stable
`CreateForm`/`CreateControl`/`DoCmd` Access automation API, but there was
no way to execute or verify it in the environment that produced it. Treat
the first run as a test: if Access reports a COM/automation error, report
the exact error message so the script can be corrected -- the same process
`build_excel_template.vbs` went through during this project's Windows 7
testing (see `CHANGELOG.md`).

## Access-specific setup the user needs

Access 2019 applies the same trust model as Excel/Word to files opened
from an untrusted location:

- **Macro/VBA security**: a database opened from a location Access doesn't
  trust shows a yellow "SECURITY WARNING -- some active content has been
  disabled" bar. This build script itself only needs COM automation
  (`Access.Application`), not the target database's own VBA project, so it
  is unaffected by this -- but once later phases add resolver macros to
  `MediaCatalog.accdb` itself, the user will need to either enable content
  each time or add a **Trusted Location** (Access Options -> Trust Center ->
  Trust Center Settings -> Trusted Locations) pointing at the MediaCatalog
  project folder.
- **"Enable Content"**: even with the file trusted, Access may still show a
  message bar the first time; click "Enable Content" if so.
- Running `build_access_database.vbs` itself may trigger a one-time
  "programmatic access to Visual Basic Project is not trusted" style error
  if Access's VBA object model access is locked down (mirrors the same
  issue documented for the Excel builder in this project's `README.md`).
  If so, enable **Trust access to the VBA project object model** under
  Access Options -> Trust Center -> Trust Center Settings -> Macro Settings,
  run the build, then it can be turned back off.

## Field types

- **Identifier-like fields are `TEXT`, never numeric**: `Inventory Number`,
  `UPC`, `Blu-ray URL`, `IMDb URL`, `IMDb ID`. This project has already
  been bitten twice by numeric-string bugs in Excel and LibreOffice Calc
  (leading zeros silently dropped from UPCs; a movie literally titled
  "1917" turning into a number) -- see `CHANGELOG.md`. Access doesn't have
  the same silent-auto-coercion-on-assignment behavior those two apps do,
  so this isn't strictly required for correctness here, but keeping these
  fields as text keeps the schema honest about what they are (opaque
  codes, not quantities) and matches every other MediaCatalog front end.
- **Year, Runtime, Blu-ray Year, Blu-ray Runtime, and Season are `INTEGER`**,
  a deliberate departure from "everything as text": these are genuine
  numeric quantities used for sorting/filtering, and Access doesn't have
  the bug class above, so storing them as real numbers is the more useful
  and more correct relational modeling choice, not a repeat of the earlier
  mistake. Flagging this explicitly in case a different call is wanted.
- **Physical Release Date** is a real `DATETIME` field, as requested.
- **Status / Error** is `MEMO` (Long Text), since it can exceed 255
  characters once warnings are appended (see `resolve_rows.py`'s
  `warning`/`status` concatenation).
- Every other field is `TEXT(255)` (Short Text).
- `ID` is a surrogate `COUNTER` (AutoNumber) primary key. Nothing in this
  phase enforces uniqueness on `Inventory Number` or `UPC` -- a box set may
  have several owned copies, or no UPC at all (see the handoff's box-set
  notes) -- so neither is a unique/candidate key here.

## Access field names vs. spreadsheet header names

**No longer applicable -- history kept below for context.** Every Access
field name is now identical to its canonical spreadsheet header text; no
mapping table is needed anywhere in this project.

Table creation originally failed on the first real run: `Blu-ray.com URL`
is not a valid Access field name (Access field names cannot contain a
period). Checked every field name in `schema.sql` against Access's actual
naming restrictions -- no period, `!`, `` ` ``, `[`, or `]`, and no
leading space -- rather than assuming this was the only one; `Blu-ray.com
Title` had the same problem and was the only other one that did. The
original fix renamed just the two Access field names (to `Blu-ray com
URL`/`Blu-ray com Title`, period dropped and replaced with a space) while
keeping the spreadsheet headers as `Blu-ray.com URL`/`Blu-ray.com Title`,
which meant `build_access_database.vbs` needed two parallel arrays
(`accessFieldNames` for `ControlSource` binding, `displayLabels` for the
prettier header text with the period) and Phase 2/3's resolver code had to
know about the mismatch.

That mismatch is now gone: the canonical spreadsheet headers were renamed
from `Blu-ray.com URL`/`Blu-ray.com Title` to `Blu-ray URL`/`Blu-ray
Title` (dropping ".com" entirely, rather than just the period) for
consistency with `Blu-ray Year`/`Blu-ray Runtime`, which never had ".com"
in the first place -- and that rename incidentally also makes them valid
Access field names verbatim. `schema.sql`, `build_access_database.vbs`,
and every field reference in `MediaCatalog_Access_Module.bas` now use the
exact same text as the spreadsheet header, for every field, with no
exceptions.

**Safeguard remains in place**: `tests/test_access_schema_names.py` parses
`schema.sql`'s field list and asserts none of them contain a period, `!`,
`` ` ``, `[`, `]`, or a leading space, so a future schema edit that
reintroduces an Access-illegal character is caught by the test suite
instead of by another live Access run. It does not (and cannot, without a
real Access install) validate anything beyond Access's field-naming
character restrictions -- reserved-word collisions and the 64-character
name-length limit are not checked.

## What the form looks like

`frmMediaCatalog` is a single, plain Detail-section form (no header/footer
sections used), one label+text-box pair per field, stacked vertically in
the same order as the table: Inventory Number, UPC, Blu-ray URL,
Blu-ray Title, IMDb URL, IMDb ID, IMDb Title, Year, Runtime, Title
Type, Season, Status / Error, Studio, Blu-ray Year, Blu-ray Runtime,
Content Rating, Physical Release Date, Disc Format, Video Codec,
Resolution, Aspect Ratio, Disc Count / Capacities (the labels show the
same text as the underlying Access field name -- see "Access field names
vs. spreadsheet header names" above). Below the last field is a **"Resolve
Current Record"** button (Phase 2). Each field's label is its own
explicitly-created and positioned control (Access's auto-attached-label
behavior turned out not to be usable here -- see "Verified vs. not
verified"). The form is taller than a typical window; Access shows its
normal vertical scrollbar for that automatically. There is no other
navigation/toolbar customization -- Access's default record-navigation
buttons at the bottom of the form window are what a user gets by default
when opening a bound form in Form view.

## Phase 2: Resolve Current Record

### What it does

Clicking **"Resolve Current Record"** runs the exact same integrated
resolver Excel/Calc's "Resolve Selected Rows" uses
(`scripts/resolve_rows.py`, same TSV-in/TSV-out contract, same 25-field
`OUTPUT_FIELDS` header, same status vocabulary), against the one record
currently open on the form:

1. Saves the current record if it has unsaved edits (`Me.Dirty = False`),
   so the fields read below reflect what's actually in the record.
2. Reads `UPC`, `Blu-ray URL`, `Blu-ray Title`, `IMDb URL`,
   `IMDb ID`, `IMDb Title`, and `Season` from the current record. If all
   of those (except Season) are blank, it says so and stops -- there's
   nothing to resolve from.
3. Writes those fields to a temp TSV file (in the user's `%TEMP%` folder,
   like Excel/Calc) matching `resolve_rows.py`'s expected input header
   exactly (`row`, `upc`, `bluray_url`, `release_title`, `imdb_url`,
   `imdb_id`, `title`, `season`) -- one data row, `row` always `1` since
   there is exactly one record.
4. Launches the configured Python interpreter against `resolve_rows.py`
   and waits for it to finish, then reads back its one-row TSV response.
5. Writes the returned fields into the current record and re-saves it:
   - Blu-ray release identity (`Blu-ray URL`, `Blu-ray Title`)
     and IMDb work identity (`IMDb URL`, `IMDb ID`, `IMDb Title`, `Year`,
     `Runtime`, `Title Type`, `Season`) are always overwritten when the
     resolver returns a value, matching Excel/Calc.
   - Everything else the resolver enriches (`Studio`, `Blu-ray Year`,
     `Blu-ray Runtime`, `Content Rating`, `Physical Release Date`,
     `Disc Format`, `Video Codec`, `Resolution`, `Aspect Ratio`,
     `Disc Count / Capacities`) only fills in if that field is currently
     blank -- manual corrections are never silently overwritten, matching
     Excel/Calc's existing behavior (not a new provenance policy; just
     matching what "Resolve Selected Rows" already does elsewhere).
   - `Year`, `Runtime`, `Season`, `Blu-ray Year`, and `Blu-ray Runtime`
     are parsed with `IsNumeric`/`CLng` before being written to their
     Integer fields (never a raw string dumped into a numeric field), and
     `Physical Release Date` is parsed as a real date only when the
     resolver's value is a strict `YYYY-MM-DD` string; anything else is
     left alone.
   - `Status / Error` gets the resolver's own status text verbatim (e.g.
     `OK - Blu-ray + IMDb`, `PARTIAL - Blu-ray only`, `NEEDS REVIEW`,
     `SKIPPED - no resolver input`), with the error field appended after a
     colon when present -- the same status vocabulary `resolve_rows.py`
     already emits, not reinvented here.
6. Shows a summary message box: the resolver's status on success, or the
   failure text on error. Temp files are deleted in both cases -- and
   since `ReadUtf8Text`/`WriteUtf8Text` each open and close their own
   `ADODB.Stream` within a single call (unlike this project's earlier
   LibreOffice Calc `ScriptForge.TextStream` handles, which were once left
   open across multiple statements), there's no lingering file handle to
   close before that delete.

### Why current-record-only

Access has no spreadsheet-style "select several rows" gesture. Resolving
just the current record is the simplest correct behavior, and matches
what the user was already doing manually in Phase 1 (open a record, fill
in a UPC by hand). Multi-record/batch resolution -- e.g. resolving every
record in a filtered datasheet view -- is a reasonable later phase, not
attempted here.

### Dependencies

Same as Excel/Calc, and deliberately reusing their existing configuration
rather than inventing an Access-only one:

- `settings.ini` at the **project root** (one level above `access\`) --
  specifically its `[runtime] python = ...` key, read with the exact same
  `ReadIniValue` parser Excel uses. `ProjectPath()` in
  `MediaCatalog_Access_Module.bas` walks up one directory from
  `CurrentProject.Path` to find it, since `MediaCatalog.accdb` lives in
  `access\` rather than at the project root the way the Excel/Calc
  templates do -- **this assumes `MediaCatalog.accdb` stays directly
  inside the `access\` folder**; moving it elsewhere will break this path
  resolution.
- `scripts\resolve_rows.py`, invoked the same way Excel does: the
  configured Python command converted to its windowless (`pythonw`/`pyw`)
  form, run via `WScript.Shell.Exec` and polled with `Sleep 100` in a loop
  (`Do While process.Status = 0 ... Loop`) up to a 43200-second (12-hour)
  timeout -- copied from Excel's `RunCommandAndWait`, not reimplemented
  from scratch. (Calc's LibreOffice Basic equivalent, `Session.RunApplication`
  plus a `FileExists` poll loop, is Basic/UNO-specific and not directly
  portable to VBA; Excel's WScript.Shell-based pattern is the correct port
  for Access, which shares Excel's VBA dialect.)
- Access's VBA project trust setting, for the button's `OnClick` to reach
  `ResolveCurrentRecord` at all -- covered under "Access-specific setup the
  user needs" above.

### Known limitations

- **Current record only** -- for resolving several records at once, see
  "Phase 3: Resolve Selected Records" below.
- No progress/cancel window (Excel/Calc's integrated resolver shows one
  via `job_progress.py`'s `run_with_progress`); a single record resolves
  quickly enough that this wasn't judged necessary for this phase, but a
  long-running Blu-ray.com lookup will make Access appear to "hang" (it
  isn't -- `DoEvents` runs during the poll loop, so the form stays
  responsive to Windows messages, but there's no visual feedback that
  work is in progress). **Confirmed live**: a real run took longer than
  expected with no progress indicator and looked stalled; it wasn't --
  give it time before assuming something is wrong.
- No cancellation.
- Whatever else "Resolve Selected Rows" doesn't do either (BRdC-only
  fallback, box-set/hierarchy awareness, etc.) -- this is a straight port
  of that command's behavior, not an enhancement of it.

## Phase 3: Resolve Selected Records

### What it does

Access has no spreadsheet-style "select several rows" gesture on a normal
bound form -- but it does on a form in **Datasheet view**, which behaves
like an Excel/Calc grid (click a row selector, Shift/Ctrl-click or drag to
extend the selection). That view can't host a command button, though: in
Datasheet view, Access renders only the grid -- a form's Header/Footer/
Detail sections (where buttons normally live) don't appear at all. So this
phase adds two new forms instead of one:

- **`frmMediaCatalogDatasheet`** -- the `MediaCatalog` table in Datasheet
  view. Select the records to resolve here.
- **`frmMediaCatalogTools`** -- a small form with just the **"Resolve
  Selected Records"** button. Keep it open alongside the datasheet (e.g.
  side by side, or one in front of the other); clicking its button reads
  the selection from `frmMediaCatalogDatasheet` *by name*, not from
  whichever window currently has focus -- clicking a button always makes
  its own form the active one, so "the active form" would just be the
  tools form itself, never the datasheet.

Clicking the button:

1. Requires `frmMediaCatalogDatasheet` to be open (shows a message and
   stops if it isn't), brings it to the foreground so its selection state
   is current, and reads which rows are selected (`SelTop`/`SelHeight`).
2. Builds **one batch TSV** covering every selected record (not one
   resolver call per record) -- same input contract as Phase 2, except
   the `row` value sent to `resolve_rows.py` is each record's real `ID`
   (the table's AutoNumber primary key), not a position. This is a
   deliberate difference from Excel/Calc's spreadsheet row numbers: a
   selected *record* has no fixed position to return to the way a
   spreadsheet row does (sorting/filtering the datasheet between the read
   and write phases would silently corrupt a position-based approach),
   but every record has a stable ID regardless.
3. Runs `resolve_rows.py` once against the whole batch (same Shell/poll
   pattern as Phase 2), then for each line of its response, looks up the
   matching record **by ID** via a fresh `DAO.Recordset` (`FindFirst`/
   `Edit`/`Update`) and writes its fields back -- same overwrite-always
   vs. overwrite-if-blank policy, same numeric/date parsing, same status
   vocabulary as Phase 2 (see "Phase 2: Resolve Current Record" above;
   not repeated or changed here).
4. Requeries `frmMediaCatalogDatasheet` so the results are visible, then
   shows a summary message box with per-status counts (Complete, Partial,
   Needs review, Cancelled, Skipped, Errors) -- the same style of summary
   Excel/Calc's integrated resolver shows, not Phase 2's single-status
   message (there's more than one result to report now).

### Why a separate form pair instead of one combined UI

Considered and rejected:

- **Making `frmMediaCatalog` itself switchable to Datasheet view**: Access
  does support this (View menu / `Ctrl+.`), but there'd be nowhere for a
  "Resolve Selected Records" button to live while in that view for the
  same header/footer-section reason above, so a second form/button was
  needed regardless.
- **A keyboard shortcut (AutoKeys macro) instead of a second form**: would
  avoid needing `frmMediaCatalogTools` at all, but creating a Macro object
  via automation uses a different, less-proven API
  (`Application.SaveAsText`/`LoadFromText`) than anything already working
  in this project (`CreateForm`/`CreateControl`), and this project has
  already needed two live-tested rounds to get `CreateForm`/`CreateControl`
  right. The two-form approach reuses only already-proven techniques.

### Dependencies

Same as Phase 2 (`settings.ini`, `scripts\resolve_rows.py`, the Shell/poll
pattern, the VBA project trust setting) -- nothing new.

### Known limitations

- Selection must be **contiguous** (`SelTop`/`SelHeight` describe one
  range of rows, the same model Access's own Datasheet view enforces --
  Ctrl-clicking to select a *non-contiguous* set of rows is not something
  Access's own row selectors support in the first place, so this isn't a
  restriction beyond what the UI already allows).
- Both forms must be open (`frmMediaCatalogDatasheet` with something
  selected, `frmMediaCatalogTools` to click the button) -- there's no
  single combined window for this workflow.
- Same "no progress/cancel window" limitation as Phase 2, more noticeable
  here since a batch of records takes proportionally longer.
- Whatever else "Resolve Selected Rows" doesn't do either -- see Phase 2's
  own "Known limitations" for the full list; this phase doesn't add or
  remove any of it, only the selection/batching mechanism around it.

## Verified vs. not verified

**First real run** found `CreateControl` called with a skipped `ParentName`
argument using VBA's `acDetail, , fields(i)` comma-gap idiom, which real
Access rejected at compile time with `Expected end of statement`. That was
fixed by passing `Empty` explicitly instead of skipping the argument.

**Second real run failed the same way, in the same place**, even after
that fix. This means the comma-gap diagnosis, while a real defect, either
wasn't the actual (or wasn't the only) cause. Rather than guess at a
third variant of the same argument-list shape, `CreateControl` is now
called with only its three required leading arguments (`FormName`,
`ControlType`, `Section`) everywhere in this script; every other property
(`ControlSource`, `Left`, `Top`, `Width`, `Height`, and the label's
`Caption`) is set explicitly afterward instead of passed positionally,
and each field's label is now created as its own separate `acLabel`
control rather than relying on `CreateControl`'s `ColumnName` argument to
auto-attach one. No call anywhere in the script now passes more than
three positional arguments or skips any argument.

Verified in this environment (still no Access available):
- `schema.sql` is valid, well-formed Jet/ACE DDL by inspection.
- Tried two offline VBScript tools against the file, and against a
  minimal known-bad snippet reproducing the original comma-gap defect, to
  calibrate how much to trust a clean result from either:
  - `vbspretty` (npm) parsed and reformatted the known-bad snippet
    without complaint -- too lenient to trust for this defect class.
  - `@devscholar/vbs-engine-js` (npm), a real VBScript-to-JS interpreter,
    also accepted the known-bad snippet via both its `addCode` and
    `executeStatement` APIs (only failing later, at simulated runtime,
    with an unrelated "Invalid procedure call" -- not the compile-time
    error real Access produced). With `WScript`/`Access.Application`
    stubbed via its `addObject` API, it did execute the *entire* current
    script end-to-end (with all `On Error Resume Next` guards stripped
    out, to surface anything they might hide) and created all 44
    expected controls (22 fields x textbox + label) with no error. Useful
    as a general structural check, but neither tool models whatever
    Access-specific compile-time restriction real `cscript.exe` is
    actually enforcing, so a clean result from either is not proof this
    is fixed.
  - A purpose-built regex-based heuristic checker (comma-gap arguments,
    `As Type`/`ByVal`/`ByRef`/`:=` VBA-only constructs, trailing
    whitespace after `_` continuations, unbalanced parens/quotes per
    logical statement, `Sub`/`Function`/`For`-`Next` counts) plus a full
    manual line-by-line re-read both came back clean on the version that
    still failed -- which is exactly why the fix this round changes
    strategy (eliminate the whole risky argument-list shape) rather than
    hunting for one more token-level mistake the same tools would likely
    miss again.
- No real Windows VBScript interpreter (`cscript.exe`) was found or
  obtainable in this environment; `wine` alone doesn't provide one (it's
  not a redistributable component of Windows).

Not verified (no Access, no Windows, no ODBC driver available here):
- That `build_access_database.vbs` now runs to completion against a real
  Access installation.
- That the resulting form looks/behaves as described, including whether
  each label ends up positioned/sized sensibly next to its textbox now
  that positioning is fully manual rather than Access's own auto-label
  default.

Please rerun the build script and report back **the full error text
verbatim, including the line number**, if it fails again -- and if it
succeeds, what the resulting form actually looks like.

### Phase 2 (Resolve Current Record): what's verified vs. not

Everything above (Phase 1: schema, form, field names) is confirmed
working live by the user. Phase 2 is new and **entirely unverified
against real Access** -- same constraint as every round before it (no
Windows, no Access, no ODBC driver in this environment).

Verified in this environment:
- `MediaCatalog_Access_Module.bas` reviewed by hand for VBA correctness
  (block balance -- 13 `Function`/`End Function` and 6 `Sub`/`End Sub`
  pairs, manually counted since the project's own VBScript-oriented
  heuristic checker doesn't apply to a real `.bas` VBA module and its
  `ByVal`/`ByRef`/`As Type` false-positives were discarded). Every field
  reference uses `frm.Controls("Field Name")` (plain string-indexed
  collection access) rather than bang-bracket syntax
  (`frm![Field Name]`), specifically to avoid any doubt about how VBA
  parses the hyphens/slashes/spaces in names like `Blu-ray URL` or
  `Status / Error` -- given this project's history, bang-bracket syntax
  was judged not worth the risk even though it's probably fine.
- The Shell/poll pattern, `settings.ini` parsing, and windowless-Python
  conversion are copied verbatim from Excel's already-working
  `RunCommandAndWait`/`GetPythonCommand`/`GetWindowlessPythonCommand`
  (excel/MediaCatalog_Excel_Module.bas) -- not reimplemented, so they
  inherit that code's track record, though the surrounding Access-specific
  code (control lookup, field writes) is new and unproven.
- `build_access_database.vbs`'s new module-import and button-creation
  code follows the exact same "only required leading arguments, set
  everything else as a property afterward" discipline established after
  the two live CreateControl failures above -- checked with the same
  heuristic checker (clean) and the same JS-based VBScript engine stub
  (executes the whole updated build script end-to-end, including the
  module import and button creation, with `On Error Resume Next` guards
  stripped out; creates all 45 expected controls). As before, neither
  tool models Access's own compile-time restrictions, so this is a
  structural check, not proof.
- The project's Python test suite (`tests/`) still passes unchanged --
  Phase 2 touched no Python code.

**Phase 2 is now confirmed working live**: the build succeeded, the
button reached `ResolveCurrentRecord`, `VBComponents.Import` worked,
`Screen.ActiveForm`/`frm.Controls("...")`/`Me.Dirty` all behaved as
expected against a real bound form, and the resolver ran end-to-end and
wrote correct data back (the one live surprise was cosmetic -- see
"Known limitations" above, not a defect). Every "not verified" item from
that first Phase 2 release is now verified.

### How to test Phase 2

1. Rebuild the database: `cscript access\build_access_database.vbs .`
   (from the project root). This should now also report importing
   `MediaCatalog_Access_Module.bas` and creating the "Resolve Current
   Record" button, alongside everything Phase 1 already did.
2. Open `access\MediaCatalog.accdb`, open `frmMediaCatalog`.
3. Make sure `settings.ini` (project root) has a working `[runtime]
   python = ...` entry -- the same one Excel/Calc already use.
4. On a fresh record, enter **UPC `031398108450`** (used earlier in this
   project's own Calc testing -- resolves to "The Spirit" on Blu-ray.com)
   in the `UPC` field, then click **"Resolve Current Record"**.
5. **What success looks like**: after a short pause (a live Blu-ray.com
   lookup, so likely several seconds, with no progress indicator -- see
   "Known limitations" above), a message box reports something like
   `Resolved: OK - Blu-ray + IMDb`, and the record's `Blu-ray URL`,
   `Blu-ray Title`, `IMDb URL`, `IMDb ID`, `IMDb Title`, `Year`,
   `Runtime`, `Title Type`, `Studio`, and other enrichment fields are now
   filled in, with `Status / Error` showing the same status text as the
   message box.
6. **What failure looks like**, and what it means:
   - A message box about `settings.ini` or Python not being found ->
     configuration issue, same fix as for Excel/Calc.
   - A message box naming a COM/automation error (e.g. about
     `Screen.ActiveForm`, `Controls`, or a specific field name) -> a real
     bug in `MediaCatalog_Access_Module.bas`; please paste the exact
     message and which step you were on (build, or clicking the button).
   - The button does nothing at all when clicked -> most likely the VBA
     project trust setting, or the module import silently didn't happen;
     check the Visual Basic Editor (Alt+F11) for whether
     `MediaCatalogAccess` appears as a module, and whether
     `ResolveCurrentRecord` exists in it.
7. Worth also trying a record with only an `IMDb ID` filled in (no UPC),
   to exercise the IMDb-only path, and an empty record (should show the
   "nothing to resolve from" message and do nothing else).

### Phase 3 (Resolve Selected Records): what's verified vs. not

Entirely unverified against real Access -- same constraint as Phase 2 was
before its own live testing.

Verified in this environment:
- `MediaCatalog_Access_Module.bas`'s block structure (12
  `Sub`/`End Sub` and 12 `Function`/`End Function` pairs, counted
  precisely with a small script rather than eyeballed) and a full manual
  re-read of `ResolveSelectedRecords` and its `DAO.Recordset` helpers.
- All field reads/writes go through `rsSource.Fields("...")` /
  `rsTable.Fields("...")` (plain DAO field access by name), never
  `Controls("...")` or bang-bracket syntax -- sidestepping both of the
  control-lookup and syntax concerns raised earlier in this document.
- `build_access_database.vbs`'s two new form-building Subs
  (`BuildDatasheetForm`, `BuildToolsForm`) follow the same
  "only-required-arguments" `CreateControl`/`CreateForm` discipline as
  every other control this script creates. Checked with the same
  heuristic checker (clean: zero comma-gap arguments anywhere in either
  file, block counts balanced) and the same JS-based VBScript engine stub
  (executes the whole updated build script end-to-end and creates exactly
  the expected 3 forms and 46 controls).
- The project's Python test suite still passes unchanged.

Not verified at all, and inherently *can't* be checked by the JS-engine
stub the way the build script's own structure can (the stub has no model
of Access's own intrinsic `Forms`/`Screen`/`CurrentDb`/`DAO` object model
-- only external `CreateObject` calls can be stubbed):
- That `frmMediaCatalogDatasheet` actually opens in Datasheet view and
  lets the user select multiple rows the way described.
- That `Forms("frmMediaCatalogDatasheet")`, `.SetFocus`,
  `.SelTop`/`.SelHeight`, and `.RecordsetClone` behave as expected against
  a real Datasheet-view form and a real multi-row selection.
- That `CurrentDb.OpenRecordset(...)`, `.FindFirst "ID = " & recordId`,
  `.Edit`, and `.Update` correctly locate and update the right records.
- That the batch resolver call and per-record write-back actually produce
  correct results across more than one record.

### How to test Phase 3

1. Rebuild: `cscript access\build_access_database.vbs .` -- should now
   also report creating `frmMediaCatalogDatasheet` and
   `frmMediaCatalogTools`, alongside everything Phases 1-2 already did.
2. In `MediaCatalog.accdb`, add at least 2-3 records via `frmMediaCatalog`
   (Phase 1/2's form) if there aren't already some -- e.g. **UPC
   `031398108450`** (from the Phase 2 test) plus one or two more,
   left unresolved (blank `Status / Error`).
3. Open `frmMediaCatalogDatasheet`. Confirm it opens directly in
   Datasheet view (a grid, not the single-record layout).
4. Select 2 or more rows using their row selectors (click the leftmost
   gray bar of one row, then Shift-click another to extend the
   selection).
5. Without closing the datasheet, also open `frmMediaCatalogTools` (it
   can overlap/tile with the datasheet window) and click **"Resolve
   Selected Records"**.
6. **What success looks like**: after a pause (proportionally longer than
   Phase 2's single-record wait, and again with no progress indicator --
   see "Known limitations" above), a message box reports something like
   `Resolved 3 record(s).` followed by the same per-status counts
   Excel/Calc's integrated resolver shows (Complete/Partial/Needs
   review/Cancelled/Skipped/Errors), and switching back to
   `frmMediaCatalogDatasheet` (it requeries automatically) shows the
   selected rows' fields filled in and their `Status / Error` columns
   updated.
7. **What failure looks like**, and what it means:
   - "Open frmMediaCatalogDatasheet..." message -> the tools form's
     button was clicked before the datasheet form was opened, or it was
     closed; open it, select rows, try again.
   - "Select one or more rows..." message -> nothing was selected (or the
     selection was lost) when the button was clicked.
   - A message box naming a COM/automation error (e.g. about
     `SelTop`/`SelHeight`, `RecordsetClone`, `OpenRecordset`, or
     `FindFirst`) -> a real bug in `ResolveSelectedRecords`; please paste
     the exact message and how many rows were selected.
   - Some records update but not others in the same batch -> please note
     which ones (and their `ID` values, visible as a column in the
     datasheet) so the correlation-by-ID logic can be checked against
     what actually happened.
8. Worth also trying: selecting just one row (should behave like Phase
   2's single-record case, just through the batch code path), and
   selecting a row that's already fully resolved alongside an unresolved
   one (confirms the manual-value-preserving fields aren't clobbered).
