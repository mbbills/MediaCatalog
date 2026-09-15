import csv
import sys
import tempfile
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent
SCRIPTS_DIR = PROJECT_ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

from job_progress import ProgressReporter
from resolve_rows import (
    OUTPUT_FIELDS,
    open_incremental_writer,
    process_rows,
    write_result_row,
)


def fake_resolve_row(row, client=None):
    # Row "2" is a deliberate mid-batch fault: a bug or an unexpected crash
    # partway through a large selection should not erase rows already
    # resolved before it.
    if row.get("row") == "2":
        raise RuntimeError("simulated crash while resolving row 2")
    return {
        "row": row.get("row"),
        "status": "OK - Blu-ray + IMDb",
        "error": "",
        "upc": row.get("upc", ""),
        "bluray_url": "",
        "release_title": "Title {}".format(row.get("row")),
        "imdb_url": "",
        "imdb_id": "",
        "title": "",
        "year": "",
        "runtime": "",
        "title_type": "",
        "season": "",
        "studio": "",
        "bluray_year": "",
        "bluray_runtime": "",
        "rating": "",
        "release_date": "",
        "disc_format": "",
        "video_codec": "",
        "resolution": "",
        "aspect_ratio": "",
        "disc_count_capacities": "",
        "source": "",
        "warning": "",
        "_network_used": False,
    }


def read_tsv_rows(path):
    with path.open("r", encoding="utf-8", newline="") as handle:
        reader = csv.reader(handle, dialect="excel-tab")
        header = next(reader)
        assert header == OUTPUT_FIELDS
        return list(reader)


def main():
    rows = [
        {"row": "1", "upc": "012345678905"},
        {"row": "2", "upc": "012345678906"},
        {"row": "3", "upc": "012345678907"},
    ]

    with tempfile.TemporaryDirectory() as tmp:
        output_path = Path(tmp) / "output.tsv"
        partial_path = output_path.with_name(output_path.name + ".partial")

        handle, writer = open_incremental_writer(partial_path)
        rows_written = 0

        def on_result(result):
            nonlocal rows_written
            write_result_row(handle, writer, result)
            rows_written += 1

        progress = ProgressReporter(len(rows))
        raised = None
        try:
            process_rows(
                rows,
                timeout=1.0,
                delay=0.0,
                progress=progress,
                client_factory=lambda timeout: object(),
                resolve_row_func=fake_resolve_row,
                on_result=on_result,
            )
        except RuntimeError as exc:
            raised = exc
        finally:
            handle.close()

        assert raised is not None, "expected the simulated crash to propagate"
        assert rows_written == 1, "only row 1 should have been written before the crash"

        # This mirrors what main() does in its own finally block: whatever
        # was written before the fault is a complete, valid file already,
        # so it is kept (never deleted) for the front end to import.
        import os

        os.replace(partial_path, output_path)
        assert output_path.exists()
        assert not partial_path.exists()

        data_rows = read_tsv_rows(output_path)
        assert len(data_rows) == 1
        assert data_rows[0][0] == "1"
        assert data_rows[0][1] == "OK - Blu-ray + IMDb"
        assert data_rows[0][5] == "Title 1"

    print("PASS: partial results survive a mid-batch crash")


if __name__ == "__main__":
    main()
