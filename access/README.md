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

**Phase 2** (this round): a **"Resolve Current Record" button** on
`frmMediaCatalog` that runs the same integrated resolver Excel/Calc call
"Resolve Selected Rows", against the one record currently open on the
form. See "Phase 2: Resolve Current Record" below for what it does, its
dependencies, and its limitations.

## What's explicitly NOT implemented yet

- The BRdC-only (Blu-ray.com-only) and single-row IMDb-lookup commands --
  deferred to later phases, one at a time, after Phase 2 is confirmed
  working live.
- Multi-record / batch resolution -- Phase 2 is current-record-only; see
  "Phase 2: Resolve Current Record" for why.
- No import/export script (CSV, or from the Excel/Calc catalogs).
- No reports.
- No parent/child (box-set) schema.
- No friendly-title natural sort (handoff section 15.2) -- noted for a
  later phase.

## Files here

| File | Purpose |
|---|---|
| `schema.sql` | The canonical `CREATE TABLE` statement (Jet/ACE SQL). Hand-maintained source of truth for the schema. |
| `MediaCatalog_Access_Module.bas` | Phase 2: the integrated-resolver VBA module (`ResolveCurrentRecord` and its helpers), ported from Excel's `ResolveSelectedRows`. Imported into the `.accdb`'s VBA project by the build script. |
| `build_access_database.vbs` | Automation script that creates a fresh `MediaCatalog.accdb`, runs `schema.sql` against it, imports `MediaCatalog_Access_Module.bas`, and builds `frmMediaCatalog` (including the Phase 2 button). |
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
  `UPC`, `Blu-ray com URL`, `IMDb URL`, `IMDb ID`. This project has already
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

Table creation failed on the first real run: `Blu-ray.com URL` is not a
valid Access field name (Access field names cannot contain a period).
Checked every field name in `schema.sql` against Access's actual naming
restrictions -- no period, `!`, `` ` ``, `[`, or `]`, and no leading space
-- rather than assuming this was the only one; `Blu-ray.com Title` has the
same problem and was the only other one that did. Every other field name
in this schema was already Access-safe and identical to its spreadsheet
header text.

Both were renamed by replacing the period with a space, matching the
convention already used for every other multi-word field name in this
schema (`Content Rating`, `Physical Release Date`, etc.), so the Access
field name differs from the canonical spreadsheet header for exactly
these two fields:

| Access field name | Canonical spreadsheet header (Excel/Calc) |
|---|---|
| `Blu-ray com URL` | `Blu-ray.com URL` |
| `Blu-ray com Title` | `Blu-ray.com Title` |
| *(every other field)* | identical in both |

**This affects more than just `schema.sql`.** `build_access_database.vbs`
independently hardcodes its own copy of the field list to build the
browse form (it does not parse `schema.sql`), and had the same "Blu-ray.
com URL"/"Blu-ray.com Title" strings baked into it -- used both to bind
each textbox's `ControlSource` (which must match the real Access field
name) and to caption each label (which can display anything, including
the period). Fixed by splitting that single list into two parallel
arrays: `accessFieldNames` (must track `schema.sql` exactly) for
`ControlSource`, and `displayLabels` (the original, prettier spreadsheet
header text, period included) for each label's `Caption`. Phase 2, which
wires up the resolver against this table, needs to use `accessFieldNames`
-- i.e. the table's real column names -- not the spreadsheet header text,
for exactly these two fields.

**Safeguard added**: `tests/test_access_schema_names.py` parses
`schema.sql`'s field list and asserts none of them contain a period, `!`,
`` ` ``, `[`, `]`, or a leading space, so a third occurrence of this
mistake (in a future schema edit) is caught by the existing test suite
instead of by another live Access run. It does not (and cannot, without a
real Access install) validate anything beyond Access's field-naming
character restrictions -- reserved-word collisions and the 64-character
name-length limit are not checked.

## What the form looks like

`frmMediaCatalog` is a single, plain Detail-section form (no header/footer
sections used), one label+text-box pair per field, stacked vertically in
the same order as the table: Inventory Number, UPC, Blu-ray.com URL,
Blu-ray.com Title, IMDb URL, IMDb ID, IMDb Title, Year, Runtime, Title
Type, Season, Status / Error, Studio, Blu-ray Year, Blu-ray Runtime,
Content Rating, Physical Release Date, Disc Format, Video Codec,
Resolution, Aspect Ratio, Disc Count / Capacities (the labels show the
original spreadsheet header text, period included, even where the
underlying Access field name differs -- see "Access field names vs.
spreadsheet header names" above). Below the last field is a **"Resolve
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
2. Reads `UPC`, `Blu-ray com URL`, `Blu-ray com Title`, `IMDb URL`,
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
   - Blu-ray release identity (`Blu-ray com URL`, `Blu-ray com Title`)
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

- **Current record only** -- see above.
- No progress/cancel window (Excel/Calc's integrated resolver shows one
  via `job_progress.py`'s `run_with_progress`); a single record resolves
  quickly enough that this wasn't judged necessary for this phase, but a
  long-running Blu-ray.com lookup will make Access appear to "hang" (it
  isn't -- `DoEvents` runs during the poll loop, so the form stays
  responsive to Windows messages, but there's no visual feedback that
  work is in progress).
- No cancellation.
- Whatever else "Resolve Selected Rows" doesn't do either (BRdC-only
  fallback, box-set/hierarchy awareness, etc.) -- this is a straight port
  of that command's behavior, not an enhancement of it.

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
  parses the hyphens/slashes/spaces in names like `Blu-ray com URL` or
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

Not verified at all:
- That the VBA module actually imports successfully via
  `VBE.ActiveVBProject.VBComponents.Import` (a new automation call this
  phase adds -- the button-creation calls were already proven live-safe
  last phase, but this one hasn't been exercised).
- That the button's `OnClick = "=ResolveCurrentRecord()"` expression
  binding actually reaches the imported module's function.
- That `Screen.ActiveForm`, `frm.Controls("...")`, `Me.Dirty`, and the
  rest of the Access object model calls inside `ResolveCurrentRecord`
  behave as expected against a real bound form and a real Integer/Date
  field.
- That the resolver actually runs end-to-end and writes correct data back.

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
   `Resolved: OK - Blu-ray + IMDb`, and the record's `Blu-ray com URL`,
   `Blu-ray com Title`, `IMDb URL`, `IMDb ID`, `IMDb Title`, `Year`,
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
