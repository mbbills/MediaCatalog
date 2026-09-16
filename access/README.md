# MediaCatalog Access front end (phase 1: schema and a basic form)

This is the initial Microsoft Access 2019 front end for MediaCatalog. It is
new, separate work alongside the Excel and LibreOffice Calc front ends --
see the project handoff document, section 16.1, for the long-term intent
(shared Python resolver/IMDb backend, Access providing forms/queries/reports).

## What's implemented in this phase

- **One flat table**, `MediaCatalog`, using the canonical column list from
  the handoff (section 7.2), with `Inventory Number` in place of the
  earlier `Spine Tag` name. No parent/child/box-set hierarchy yet -- that is
  explicitly deferred, per the handoff (section 7.3).
- **One basic data-entry/browse form**, `frmMediaCatalog`, bound to that
  table: a plain vertical stack of labeled text boxes, one per field, in
  the same order as the table. Enough to type in a UPC and view or edit a
  row. Nothing beyond that -- no validation, no lookups, no buttons.

See `schema.sql` for the exact field list and types, and the "Field types"
section below for the reasoning.

## What's explicitly NOT in this phase

- No VBA resolver macros, no `Shell()`/COM call into the Python backend.
- No import/export script (CSV, or from the Excel/Calc catalogs).
- No reports.
- No parent/child (box-set) schema.
- No friendly-title natural sort (handoff section 15.2) -- noted for a
  later phase.

## Files here

| File | Purpose |
|---|---|
| `schema.sql` | The canonical `CREATE TABLE` statement (Jet/ACE SQL). Hand-maintained source of truth for the schema. |
| `build_access_database.vbs` | Automation script that creates a fresh `MediaCatalog.accdb`, runs `schema.sql` against it, and builds `frmMediaCatalog`. |
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
  `UPC`, `Blu-ray.com URL`, `IMDb URL`, `IMDb ID`. This project has already
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

## What the form looks like

`frmMediaCatalog` is a single, plain Detail-section form (no header/footer
sections used), one label+text-box pair per field, stacked vertically in
the same order as the table: Inventory Number, UPC, Blu-ray.com URL,
Blu-ray.com Title, IMDb URL, IMDb ID, IMDb Title, Year, Runtime, Title
Type, Season, Status / Error, Studio, Blu-ray Year, Blu-ray Runtime,
Content Rating, Physical Release Date, Disc Format, Video Codec,
Resolution, Aspect Ratio, Disc Count / Capacities. Labels use Access's
default attached-label behavior (auto-created and placed by
`CreateControl` when a bound field name is supplied), left at Access's
default size/offset rather than manually positioned. The form is taller
than a typical window; Access shows its normal vertical scrollbar for that
automatically. There is no navigation/toolbar customization -- Access's
default record-navigation buttons at the bottom of the form window are
what a user gets by default when opening a bound form in Form view.

## Verified vs. not verified

**First real run found a genuine bug**: `CreateControl` was called with a
skipped `ParentName` argument using VBA's `acDetail, , fields(i)`
comma-gap idiom. That compiles fine inside a real VBA project, but
VBScript's late-bound `IDispatch` calls into an external COM object (as
opposed to VBScript's own intrinsic functions like `MsgBox`, which get
special compiler support for the same idiom) do not support it, and it
failed with `Expected end of statement` -- reported on the `Next i` two
lines later, not on the actual bad line, since that's where the parser's
recovery gave up. Fixed by passing `Empty` explicitly for that argument
and moving `Left`/`Top`/`Width`/`Height` to property assignments after
`CreateControl` returns, instead of passing them positionally, removing
the only other place a similar gap could have been introduced.

Verified in this environment (still no Access available):
- `schema.sql` is valid, well-formed Jet/ACE DDL by inspection.
- `build_access_database.vbs` was checked with a purpose-built heuristic
  static analyzer (regex-based: comma-gap arguments in late-bound calls,
  `As Type`/`ByVal`/`ByRef`/`:=` VBA-only constructs, trailing whitespace
  after `_` continuations, unbalanced parens and quotes per logical
  statement, `Sub`/`End Sub` and `For`/`Next` counts) plus a full manual
  line-by-line re-read after the fix above. No real VBScript interpreter
  was available to actually execute it (no Windows, and `wine` alone
  doesn't provide a redistributable `cscript.exe`).

Not verified (no Access, no Windows, no ODBC driver available here):
- That `build_access_database.vbs` now runs to completion against a real
  Access installation -- this fixes the specific reported error, but a
  static check is not the same as a real run.
- That the resulting form looks/behaves as described.
- Exact default label width/offset Access chooses for the auto-attached
  labels (cosmetic only).

Please rerun the build script and report back what actually happens
(including any further errors) before the next phase (resolver
integration) begins.
