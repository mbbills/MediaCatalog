#!/usr/bin/env python3
"""Offline regression tests for the UPCItemDB batch helper."""

import csv
import sys
import tempfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

from upcitemdb_lookup import (
    build_lookup_url,
    process_rows,
    read_input,
    resolve_code,
    write_output,
)


class FakeClient:
    def __init__(self, items):
        self.items = items

    def lookup(self, code):
        return self.items


class CancelAfterOne:
    def __init__(self):
        self.is_cancelled = False

    def cancelled(self):
        return self.is_cancelled

    def start_item(self, current, phase=""):
        pass

    def finish_item(self, status):
        self.is_cancelled = True

    def wait(self, seconds, phase=""):
        return False


assert build_lookup_url(
    "https://api.upcitemdb.com/prod/trial/lookup?lang=en&upc=old",
    "012345678905",
) == "https://api.upcitemdb.com/prod/trial/lookup?lang=en&upc=012345678905"

ok = resolve_code(FakeClient([{"title": "Example Movie [Blu-ray]"}]), "012345678905")
assert ok["status"] == "OK"
assert ok["title"] == "Example Movie [Blu-ray]"
assert ok["url"] == ""

duplicate = resolve_code(
    FakeClient([{"title": "Same Title"}, {"title": "Same Title"}]),
    "012345678905",
)
assert duplicate["status"] == "OK"
assert duplicate["title"] == "Same Title"

ambiguous = resolve_code(
    FakeClient([{"title": "First"}, {"title": "Second"}]),
    "012345678905",
)
assert ambiguous["status"] == "AMBIGUOUS"

missing = resolve_code(FakeClient([]), "012345678905")
assert missing["status"] == "NOT_FOUND"

invalid = resolve_code(FakeClient([]), "12345678")
assert invalid["status"] == "ERROR"

cancelled_rows = process_rows(
    [
        {"row": "2", "upc": "012345678905"},
        {"row": "3", "upc": "098765432109"},
    ],
    FakeClient([{"title": "Example Movie"}]),
    10.5,
    CancelAfterOne(),
)
assert [row["status"] for row in cancelled_rows] == ["OK", "CANCELLED"]
assert cancelled_rows[1]["error"] == "Cancelled by user"

with tempfile.TemporaryDirectory() as temp_dir:
    input_path = Path(temp_dir) / "input.tsv"
    output_path = Path(temp_dir) / "output.tsv"
    input_path.write_text("row\tupc\n2\t012345678905\n", encoding="utf-8")
    rows = read_input(input_path)
    assert rows == [{"row": "2", "upc": "012345678905"}]
    ok["row"] = "2"
    ok["upc"] = "012345678905"
    write_output(output_path, [ok])
    with output_path.open("r", encoding="utf-8", newline="") as source:
        parsed = list(csv.DictReader(source, dialect="excel-tab"))
    assert parsed[0]["status"] == "OK"
    assert parsed[0]["source"] == "UPCItemDB"

print("PASS: UPCItemDB lookup parser and TSV contract")
