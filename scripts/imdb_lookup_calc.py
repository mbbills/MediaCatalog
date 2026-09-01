from pathlib import Path
import sys
import traceback


def safe_field(value):
    """Make a value safe for the single-line TSV response Calc reads."""
    if value is None:
        return ""

    return (
        str(value)
        .replace("\t", " ")
        .replace("\r", " ")
        .replace("\n", " ")
    )


def write_result(output_file, fields):
    output_file.write_text(
        "\t".join(safe_field(value) for value in fields),
        encoding="utf-8",
    )


def main():
    if len(sys.argv) != 3:
        raise SystemExit(
            "Usage: imdb_lookup_calc.py input.txt output.tsv"
        )

    input_file = Path(sys.argv[1])
    output_file = Path(sys.argv[2])

    try:
        # Import inside the exception handler so module/configuration
        # failures are returned to Calc instead of disappearing silently.
        from imdb_lookup import lookup

        # utf-8-sig accepts ordinary UTF-8 and strips a BOM when present.
        lines = input_file.read_text(
            encoding="utf-8-sig"
        ).splitlines()

        raw_name = lines[0].strip() if lines else ""
        imdb_id_hint = lines[1].strip() if len(lines) > 1 else ""

        result = lookup(
            raw_name,
            imdb_id_hint=imdb_id_hint,
        )

        fields = [
            "1" if result.get("success") else "0",
            result.get("imdb_id", ""),
            result.get("imdb_url", ""),
            result.get("title", ""),
            result.get("year", ""),
            result.get("runtime", ""),
            result.get("title_type", ""),
            result.get("season", ""),
            result.get("error", ""),
            result.get("source", ""),
        ]

        write_result(output_file, fields)

    except Exception:
        # Preserve the response shape even on an unexpected failure.
        write_result(
            output_file,
            [
                "0",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                traceback.format_exc(),
                "",
            ],
        )


if __name__ == "__main__":
    main()
