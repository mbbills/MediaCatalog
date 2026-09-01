#!/usr/bin/env python3
"""Integrated MediaCatalog row resolver.

A selected row can begin with any reliable identifier MediaCatalog already
supports: UPC/EAN, an exact Blu-ray.com release URL, an IMDb title URL/ID, a
physical-release title, or a canonical title.  Manual identifiers take
precedence over discovered identifiers; downstream metadata is refreshed only
from the winning identifier.
"""

import csv
import sys
import urllib.error
from pathlib import Path

from bluray_lookup_excel import (
    BlurayClient,
    fetch_release_page,
    normalize_code,
    normalize_release_url,
    resolve_code,
    safe_field,
)
from config import load_settings
from imdb_lookup import lookup as lookup_imdb
from imdb_lookup import normalize_imdb_id
from job_progress import run_with_progress


INPUT_FIELDS = [
    "row",
    "upc",
    "bluray_url",
    "release_title",
    "imdb_url",
    "imdb_id",
    "title",
    "season",
]

OUTPUT_FIELDS = [
    "row",
    "status",
    "error",
    "upc",
    "bluray_url",
    "release_title",
    "imdb_url",
    "imdb_id",
    "title",
    "year",
    "runtime",
    "title_type",
    "season",
    "studio",
    "bluray_year",
    "bluray_runtime",
    "rating",
    "release_date",
    "disc_format",
    "video_codec",
    "resolution",
    "aspect_ratio",
    "disc_count_capacities",
    "source",
    "warning",
]


def clean_input(value):
    return (value or "").strip()


def read_input(input_path):
    with input_path.open("r", encoding="utf-8-sig", newline="") as source:
        reader = csv.DictReader(source, dialect="excel-tab")
        if not reader.fieldnames or "row" not in reader.fieldnames:
            raise ValueError("Input TSV must contain a row header")
        return list(reader)


def base_result(row):
    return {
        "row": clean_input(row.get("row")),
        "status": "",
        "error": "",
        "upc": normalize_code(row.get("upc")),
        "bluray_url": clean_input(row.get("bluray_url")),
        "release_title": clean_input(row.get("release_title")),
        "imdb_url": clean_input(row.get("imdb_url")),
        "imdb_id": clean_input(row.get("imdb_id")),
        "title": clean_input(row.get("title")),
        "year": "",
        "runtime": "",
        "title_type": "",
        "season": clean_input(row.get("season")),
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
    }


def normalize_manual_imdb_identity(imdb_id, imdb_url):
    """Return the authoritative manually supplied IMDb ID and warnings."""
    imdb_id = clean_input(imdb_id)
    imdb_url = clean_input(imdb_url)
    warnings = []

    if imdb_id:
        normalized = normalize_imdb_id(imdb_id)
        if imdb_url:
            url_id = normalize_imdb_id(imdb_url)
            if url_id != normalized:
                warnings.append(
                    "IMDb ID overrides conflicting IMDb URL {}".format(url_id)
                )
        return normalized, warnings

    if imdb_url:
        return normalize_imdb_id(imdb_url), warnings

    return None, warnings


def resolve_bluray_stage(
    row,
    client,
    resolve_code_func=resolve_code,
    fetch_release_page_func=fetch_release_page,
):
    """Resolve and read one physical release, preferring a supplied URL."""
    code = normalize_code(row.get("upc"))
    supplied_url = clean_input(row.get("bluray_url"))

    if supplied_url:
        release_url = normalize_release_url(supplied_url)
        source = "Manual Blu-ray.com URL"
    elif code:
        if len(code) not in (12, 13):
            raise ValueError("UPC/EAN must contain exactly 12 or 13 digits")
        resolved = resolve_code_func(client, code)
        if resolved.get("status") != "OK":
            return {
                "success": False,
                "status": resolved.get("status", "ERROR"),
                "error": resolved.get("error", "Blu-ray.com lookup failed"),
                "source": resolved.get("source", "Blu-ray.com"),
                "network_used": True,
            }
        release_url = normalize_release_url(resolved["url"])
        source = resolved.get("source", "Blu-ray.com")
    else:
        return {
            "success": False,
            "status": "NO_INPUT",
            "error": "",
            "source": "",
            "network_used": False,
        }

    page = fetch_release_page_func(client, release_url)
    page.update(
        {
            "success": True,
            "status": "OK",
            "source": source,
            "network_used": True,
            "url": release_url,
        }
    )
    return page


def resolve_row(
    row,
    client=None,
    resolve_code_func=resolve_code,
    fetch_release_page_func=fetch_release_page,
    imdb_lookup_func=lookup_imdb,
):
    result = base_result(row)
    errors = []
    warnings = []
    sources = []
    bluray_ok = False
    imdb_ok = False
    network_used = False
    linked_imdb_id = None

    try:
        manual_imdb_id, identity_warnings = normalize_manual_imdb_identity(
            row.get("imdb_id"),
            row.get("imdb_url"),
        )
        warnings.extend(identity_warnings)
    except ValueError as exc:
        manual_imdb_id = None
        errors.append(str(exc))
        # An invalid manual identity is deliberate input. Do not silently
        # replace it with Blu-ray.com's link or a title heuristic.
        manual_identity_invalid = True
    else:
        manual_identity_invalid = False

    try:
        bluray = resolve_bluray_stage(
            row,
            client,
            resolve_code_func=resolve_code_func,
            fetch_release_page_func=fetch_release_page_func,
        )
        network_used = bluray.get("network_used", False)
        if bluray.get("success"):
            bluray_ok = True
            sources.append(bluray.get("source", "Blu-ray.com"))
            result["bluray_url"] = bluray.get("url", result["bluray_url"])
            result["release_title"] = bluray.get(
                "release_title", result["release_title"]
            )
            result["studio"] = bluray.get("studio", "")
            result["bluray_year"] = bluray.get("year", "")
            result["bluray_runtime"] = bluray.get("runtime", "")
            result["rating"] = bluray.get("rating", "")
            result["release_date"] = bluray.get("release_date", "")
            result["disc_format"] = bluray.get("disc_format", "")
            result["video_codec"] = bluray.get("video_codec", "")
            result["resolution"] = bluray.get("resolution", "")
            result["aspect_ratio"] = bluray.get("aspect_ratio", "")
            result["disc_count_capacities"] = bluray.get(
                "disc_count_capacities", ""
            )
            linked_imdb_id = bluray.get("imdb_id") or None
            if bluray.get("imdb_warning"):
                warnings.append(bluray["imdb_warning"])
        elif bluray.get("status") not in ("NO_INPUT",):
            errors.append(bluray.get("error") or bluray.get("status"))
    except (OSError, ValueError, urllib.error.URLError) as exc:
        network_used = bool(
            clean_input(row.get("bluray_url")) or normalize_code(row.get("upc"))
        )
        errors.append(str(exc))

    winning_imdb_id = None
    identity_source = ""

    if not manual_identity_invalid:
        if manual_imdb_id:
            winning_imdb_id = manual_imdb_id
            identity_source = "Manual IMDb ID"
            if linked_imdb_id and linked_imdb_id != manual_imdb_id:
                warnings.append(
                    "Manual IMDb override; Blu-ray.com links {}".format(
                        linked_imdb_id
                    )
                )
        elif linked_imdb_id:
            winning_imdb_id = linked_imdb_id
            identity_source = "Blu-ray.com IMDb link"

        if winning_imdb_id:
            result["imdb_id"] = winning_imdb_id
            result["imdb_url"] = "https://www.imdb.com/title/{}/".format(
                winning_imdb_id
            )

        raw_title = (
            result["release_title"]
            if result["release_title"]
            and result["release_title"] not in ("[NOT FOUND]", "[AMBIGUOUS]")
            else clean_input(row.get("title"))
        )

        if winning_imdb_id or raw_title:
            imdb = imdb_lookup_func(raw_title, imdb_id_hint=winning_imdb_id)
            if imdb.get("success"):
                imdb_ok = True
                sources.append(identity_source or "IMDb title match")
                result["imdb_id"] = imdb.get("imdb_id", "")
                result["imdb_url"] = imdb.get("imdb_url", "")
                result["title"] = imdb.get("title", "")
                result["year"] = imdb.get("year", "")
                result["runtime"] = imdb.get("runtime", "")
                result["title_type"] = imdb.get("title_type", "")
                result["season"] = imdb.get("season", "")
            else:
                errors.append(imdb.get("error", "IMDb lookup failed"))

    if bluray_ok and imdb_ok:
        result["status"] = "OK - Blu-ray + IMDb"
    elif bluray_ok:
        result["status"] = "PARTIAL - Blu-ray only"
    elif imdb_ok:
        result["status"] = "PARTIAL - IMDb only"
    elif not any(
        clean_input(row.get(name))
        for name in (
            "upc",
            "bluray_url",
            "release_title",
            "imdb_url",
            "imdb_id",
            "title",
        )
    ):
        result["status"] = "SKIPPED - no resolver input"
    else:
        result["status"] = "NEEDS REVIEW"

    result["error"] = "; ".join(dict.fromkeys(error for error in errors if error))
    result["warning"] = "; ".join(
        dict.fromkeys(warning for warning in warnings if warning)
    )
    if result["warning"]:
        result["status"] += "; " + result["warning"]
    result["source"] = " | ".join(dict.fromkeys(source for source in sources if source))
    result["_network_used"] = network_used
    return result


def cancelled_result(row):
    result = base_result(row)
    result["status"] = "CANCELLED"
    result["error"] = "Cancelled by user"
    result["_network_used"] = False
    return result


def process_rows(
    rows,
    timeout,
    delay,
    progress,
    client_factory=BlurayClient,
    resolve_row_func=resolve_row,
):
    results = []
    cache = {}
    client = None

    for index, row in enumerate(rows):
        if progress.cancelled():
            break

        label = (
            normalize_code(row.get("upc"))
            or clean_input(row.get("bluray_url"))
            or clean_input(row.get("imdb_id"))
            or clean_input(row.get("title"))
            or "row {}".format(row.get("row", ""))
        )
        progress.start_item(label, "Resolving physical release and IMDb data...")

        cache_key = tuple(clean_input(row.get(field)) for field in INPUT_FIELDS[1:])
        if cache_key in cache:
            resolved = dict(cache[cache_key])
            resolved["row"] = clean_input(row.get("row"))
            made_network_lookup = False
        else:
            needs_bluray = bool(
                normalize_code(row.get("upc"))
                or clean_input(row.get("bluray_url"))
            )
            if client is None and needs_bluray:
                client = client_factory(timeout)
            resolved = resolve_row_func(row, client=client)
            cache[cache_key] = dict(resolved)
            made_network_lookup = resolved.pop("_network_used", False)
            cache[cache_key].pop("_network_used", None)

        results.append(resolved)
        progress.finish_item(resolved.get("status"))

        if progress.cancelled():
            break
        if index + 1 < len(rows) and delay and made_network_lookup:
            if not progress.wait(
                delay,
                "Waiting {:.1f}s before the next Blu-ray.com row...".format(
                    delay
                ),
            ):
                break

    for row in rows[len(results):]:
        results.append(cancelled_result(row))

    return results


def write_output(output_path, results):
    with output_path.open("w", encoding="utf-8", newline="") as destination:
        writer = csv.writer(destination, dialect="excel-tab", lineterminator="\n")
        writer.writerow(OUTPUT_FIELDS)
        for result in results:
            writer.writerow([safe_field(result.get(field)) for field in OUTPUT_FIELDS])


def main():
    if len(sys.argv) != 3:
        raise SystemExit("Usage: resolve_rows.py input.tsv output.tsv")

    input_path = Path(sys.argv[1])
    output_path = Path(sys.argv[2])
    rows = read_input(input_path)

    try:
        settings = load_settings()
        delay = settings.getfloat(
            "bluray", "request_delay_seconds", fallback=0.75
        )
        timeout = settings.getfloat(
            "bluray", "timeout_seconds", fallback=45.0
        )
        show_progress = settings.getboolean(
            "progress", "show_window", fallback=True
        )
        always_on_top = settings.getboolean(
            "progress", "always_on_top", fallback=True
        )

        if delay < 0:
            raise ValueError("request_delay_seconds cannot be negative")
        if timeout <= 0:
            raise ValueError("timeout_seconds must be positive")

        results = run_with_progress(
            "MediaCatalog - Integrated resolver",
            len(rows),
            lambda progress: process_rows(rows, timeout, delay, progress),
            enabled=show_progress,
            always_on_top=always_on_top,
        )
        write_output(output_path, results)
    except Exception:
        output_path.unlink(missing_ok=True)
        raise


if __name__ == "__main__":
    try:
        main()
    except Exception:
        import traceback

        traceback.print_exc()
        raise SystemExit(1)
