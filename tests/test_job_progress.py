#!/usr/bin/env python3
"""Offline checks for shared progress and cancellation state."""

import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

from job_progress import (
    ProgressReporter,
    format_duration,
    normalize_status,
    run_with_progress,
)


assert format_duration(0) == "0:00"
assert format_duration(65) == "1:05"
assert format_duration(3661) == "1:01:01"
assert normalize_status("OK - Blu-ray + IMDb") == "OK"
assert normalize_status("PARTIAL - Blu-ray only") == "PARTIAL"
assert normalize_status("NEEDS REVIEW; title mismatch") == "NEEDS_REVIEW"
assert normalize_status("SKIPPED - no resolver input") == "SKIPPED"

reporter = ProgressReporter(3)
reporter.start_item("012345678905")
reporter.finish_item("OK")
state = reporter.snapshot()
assert state["completed"] == 1
assert state["counts"]["OK"] == 1
assert state["current"] == "012345678905"

reporter.finish_item("PARTIAL - Blu-ray only")
reporter.finish_item("NEEDS REVIEW; title mismatch")
state = reporter.snapshot()
assert state["counts"]["PARTIAL"] == 1
assert state["counts"]["NEEDS_REVIEW"] == 1
assert state["counts"]["ERROR"] == 0

reporter.request_cancel()
assert reporter.cancelled()
assert reporter.wait(0.01) is False

result = run_with_progress(
    "Headless test",
    1,
    lambda progress: (progress.finish_item("OK"), "done")[1],
    enabled=False,
)
assert result == "done"

print("PASS: shared progress and cancellation state")
