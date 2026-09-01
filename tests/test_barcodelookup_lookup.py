#!/usr/bin/env python3
"""Offline regression tests for the Barcode Lookup batch helper."""

import csv
import subprocess
import sys
import tempfile
import urllib.parse
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

from barcodelookup_lookup import build_lookup_url, read_input, resolve_code, write_output


class FakeClient:
    def __init__(self, products):
        self.products = products

    def lookup(self, code):
        return self.products


url = build_lookup_url(
    "https://api.barcodelookup.com/v3/products?geo=us&barcode=old&key=old",
    "012345678905",
    "test-key",
)
parsed = urllib.parse.urlsplit(url)
query = dict(urllib.parse.parse_qsl(parsed.query))
assert parsed.netloc == "api.barcodelookup.com"
assert query == {"geo": "us", "barcode": "012345678905", "key": "test-key"}

ok = resolve_code(FakeClient([{"title": "Example Movie [DVD]"}]), "012345678905")
assert ok["status"] == "OK"
assert ok["title"] == "Example Movie [DVD]"
assert ok["source"] == "BarcodeLookup.com"
assert ok["url"] == ""

duplicate = resolve_code(
    FakeClient([{"title": "Same Title"}, {"title": "Same Title"}]),
    "012345678905",
)
assert duplicate["status"] == "OK"

ambiguous = resolve_code(
    FakeClient([{"title": "First"}, {"title": "Second"}]),
    "012345678905",
)
assert ambiguous["status"] == "AMBIGUOUS"

missing = resolve_code(FakeClient([]), "012345678905")
assert missing["status"] == "NOT_FOUND"

invalid = resolve_code(FakeClient([]), "12345678")
assert invalid["status"] == "ERROR"

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
        parsed_rows = list(csv.DictReader(source, dialect="excel-tab"))
    assert parsed_rows[0]["status"] == "OK"
    assert parsed_rows[0]["source"] == "BarcodeLookup.com"

    blocked_output = Path(temp_dir) / "blocked-automation.tsv"
    completed = subprocess.run(
        [sys.executable, str(SCRIPTS / "barcodelookup_lookup.py"), str(input_path), str(blocked_output)],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        check=False,
    )
    assert completed.returncode == 0
    assert blocked_output.exists()
    with blocked_output.open("r", encoding="utf-8", newline="") as source:
        blocked_rows = list(csv.DictReader(source, dialect="excel-tab"))
    assert blocked_rows[0]["status"] == "ERROR"
    assert "requires a paid API subscription" in blocked_rows[0]["error"]

print("PASS: Barcode Lookup parser and TSV contract")
