#!/usr/bin/env python3
"""Batch Blu-ray.com release-detail enrichment bridge."""

import csv
import sys
import traceback
import urllib.error
from pathlib import Path

from bluray_lookup_excel import (
    BlurayClient,
    fetch_release_details,
    normalize_code,
    normalize_release_url,
    resolve_code,
    safe_field,
)
from config import load_settings
from job_progress import run_with_progress


OUTPUT_FIELDS = [
    "row",
    "upc",
    "status",
    "source",
    "error",
    "url",
    "studio",
    "year",
    "runtime",
    "rating",
    "release_date",
    "disc_format",
    "video_codec",
    "resolution",
    "aspect_ratio",
    "disc_count_capacities",
]


def read_input(input_path):
    with input_path.open("r", encoding="utf-8-sig", newline="") as source:
        reader = csv.DictReader(source, dialect="excel-tab")
        required = {"row", "upc", "url"}
        if not reader.fieldnames or not required.issubset(reader.fieldnames):
            raise ValueError("Input TSV must contain row, upc, and url headers")
        return list(reader)


def error_result(row, code, status, source, message, url=""):
    result = {field: "" for field in OUTPUT_FIELDS}
    result.update(
        {
            "row": row,
            "upc": code,
            "status": status,
            "source": source,
            "error": message,
            "url": url,
        }
    )
    return result


def resolve_release_url(client, code, supplied_url):
    if len(code) in (12, 13):
        resolved = resolve_code(client, code)
        if resolved["status"] == "OK":
            return normalize_release_url(resolved["url"]), resolved["source"]
        if resolved["status"] not in ("NOT_FOUND", "AMBIGUOUS") or not supplied_url:
            return resolved, ""
    elif not supplied_url:
        return resolve_code(client, code), ""

    return normalize_release_url(supplied_url), "No UPC/EAN, URL OK"


def enrich_row(client, row):
    code = normalize_code(row.get("upc"))
    supplied_url = (row.get("url") or "").strip()
    row_number = row.get("row", "")

    if not supplied_url and len(code) not in (12, 13):
        return error_result(
            row_number,
            code,
            "ERROR",
            "",
            "A release hyperlink or exact 12/13-digit UPC/EAN is required",
        )

    resolved_url, source = resolve_release_url(client, code, supplied_url)
    if isinstance(resolved_url, dict):
        return error_result(
            row_number,
            code,
            resolved_url["status"],
            resolved_url.get("source", ""),
            resolved_url.get("error", ""),
            resolved_url.get("url", ""),
        )

    details = fetch_release_details(client, resolved_url)
    result = {field: "" for field in OUTPUT_FIELDS}
    result.update(details)
    result.update(
        {
            "row": row_number,
            "upc": code,
            "status": "OK",
            "source": source,
            "error": "",
        }
    )
    return result


def write_output(output_path, results):
    with output_path.open("w", encoding="utf-8", newline="") as destination:
        writer = csv.writer(destination, dialect="excel-tab", lineterminator="\n")
        writer.writerow(OUTPUT_FIELDS)
        for result in results:
            writer.writerow([safe_field(result.get(field)) for field in OUTPUT_FIELDS])


def process_rows(
    rows,
    timeout,
    delay,
    progress,
    client_factory=BlurayClient,
):
    results = []
    cache = {}
    client = None

    for index, row in enumerate(rows):
        if progress.cancelled():
            break

        supplied_url = (row.get("url") or "").strip()
        code = normalize_code(row.get("upc"))
        cache_key = (code, supplied_url)
        made_network_lookup = False
        progress.start_item(
            code or supplied_url,
            "Reading Blu-ray.com release details...",
        )

        try:
            if cache_key in cache:
                resolved = dict(cache[cache_key])
                resolved["row"] = row.get("row", "")
            else:
                if client is None:
                    client = client_factory(timeout)
                made_network_lookup = True
                resolved = enrich_row(client, row)
                cache[cache_key] = dict(resolved)
        except (OSError, ValueError, urllib.error.URLError) as exc:
            resolved = error_result(
                row.get("row", ""),
                code,
                "ERROR",
                "",
                str(exc),
                supplied_url,
            )

        results.append(resolved)
        progress.finish_item(resolved.get("status"))

        if progress.cancelled():
            break
        if index + 1 < len(rows) and delay and made_network_lookup:
            if not progress.wait(
                delay,
                "Waiting {:.1f}s before the next release page...".format(delay),
            ):
                break

    for row in rows[len(results):]:
        results.append(
            error_result(
                row.get("row", ""),
                normalize_code(row.get("upc")),
                "CANCELLED",
                "",
                "Cancelled by user",
                (row.get("url") or "").strip(),
            )
        )

    return results


def main():
    if len(sys.argv) != 3:
        raise SystemExit("Usage: bluray_details.py input.tsv output.tsv")

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
            "MediaCatalog - Blu-ray.com detail enrichment",
            len(rows),
            lambda progress: process_rows(rows, timeout, delay, progress),
            enabled=show_progress,
            always_on_top=always_on_top,
        )
    except Exception as exc:
        results = [
            error_result(
                row.get("row", ""),
                normalize_code(row.get("upc")),
                "ERROR",
                "",
                str(exc),
                (row.get("url") or "").strip(),
            )
            for row in rows
        ]

    write_output(output_path, results)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
        raise SystemExit(2)
