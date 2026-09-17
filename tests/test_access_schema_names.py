#!/usr/bin/env python3
"""Guard against Access-illegal field names in access/schema.sql.

Access field/object names cannot contain a period, "!", "`", "[", or "]",
and cannot start with a leading space. schema.sql already hit this once
("Blu-ray.com URL"/"Blu-ray.com Title", copied verbatim from the
spreadsheet header list, both contain a period) -- this is a cheap static
check so the next occurrence is caught here instead of by another live
Access run.

This does not (and cannot, without a real Access install) check anything
beyond these character/leading-space rules: reserved-word collisions and
the 64-character name-length limit are out of scope.
"""

import re
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent
SCHEMA_PATH = PROJECT_ROOT / "access" / "schema.sql"

BANNED_CHARS = {".", "!", "`", "[", "]"}

# One field definition per line inside CREATE TABLE: either a
# bracket-quoted name ("[Some Field]") or a bare identifier ("UPC"),
# followed by its type. Skip blank lines and "--" comment lines.
FIELD_LINE = re.compile(
    r"""^\s*
    (?:\[(?P<bracketed>[^\]]*)\]|(?P<bare>[A-Za-z_][A-Za-z0-9_ ]*?))
    \s+
    (?:COUNTER\b|TEXT\(\d+\)|INTEGER\b|DATETIME\b|MEMO\b)
    """,
    re.VERBOSE,
)


def extract_field_names(schema_sql):
    names = []
    for line in schema_sql.splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("--"):
            continue
        match = FIELD_LINE.match(stripped)
        if match:
            name = match.group("bracketed")
            if name is None:
                name = match.group("bare")
            names.append(name)
    return names


def find_naming_violations(name):
    violations = []
    for ch in BANNED_CHARS:
        if ch in name:
            violations.append("contains %r" % ch)
    if name.startswith(" "):
        violations.append("starts with a leading space")
    return violations


def main():
    schema_sql = SCHEMA_PATH.read_text(encoding="utf-8")
    names = extract_field_names(schema_sql)

    # Sanity check that the parser actually found the real field list, not
    # an empty/malformed result that would make this test vacuously pass.
    assert len(names) >= 20, (
        "Expected at least 20 fields in schema.sql, found %d -- "
        "the FIELD_LINE parser may need updating" % len(names)
    )
    assert "ID" in names
    assert "Inventory Number" in names
    assert "Blu-ray URL" in names
    assert "Blu-ray Title" in names

    failures = []
    for name in names:
        violations = find_naming_violations(name)
        if violations:
            failures.append("%r: %s" % (name, "; ".join(violations)))

    if failures:
        print("Access-illegal field name(s) found in schema.sql:")
        for failure in failures:
            print("  - " + failure)
        sys.exit(1)

    print("PASS: schema.sql field names are Access-safe (%d fields checked)" % len(names))


if __name__ == "__main__":
    main()
