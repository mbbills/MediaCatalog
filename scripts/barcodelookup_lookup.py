#!/usr/bin/env python3
"""Resolve exact UPC/EAN codes through the Barcode Lookup API."""

import csv
import json
import sys
import traceback
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

from config import load_settings
from job_progress import run_with_progress


DEFAULT_ENDPOINT = "https://api.barcodelookup.com/v3/products"
USER_AGENT = "MediaCatalog/0.3.0 (Windows 7; Python 3.8)"


def safe_field(value):
    """Keep a value on one TSV line."""
    if value is None:
        return ""
    return str(value).replace("\t", " ").replace("\r", " ").replace("\n", " ")


def normalize_code(value):
    """Return digits only while preserving leading zeroes."""
    return "".join(character for character in (value or "").strip() if character.isdigit())


def build_lookup_url(endpoint, code, api_key):
    """Build an exact-barcode API URL without logging the configured key."""
    parsed = urllib.parse.urlsplit(endpoint.strip())
    if parsed.scheme not in ("http", "https") or not parsed.netloc:
        raise ValueError("Barcode Lookup endpoint must be an HTTP or HTTPS URL")
    query = urllib.parse.parse_qsl(parsed.query, keep_blank_values=True)
    query = [
        (key, value)
        for key, value in query
        if key.lower() not in ("barcode", "key", "formatted")
    ]
    query.extend((("barcode", code), ("key", api_key)))
    return urllib.parse.urlunsplit(
        (parsed.scheme, parsed.netloc, parsed.path, urllib.parse.urlencode(query), "")
    )


class BarcodeLookupClient:
    """Standard-library client for Barcode Lookup's documented v3 API."""

    def __init__(self, endpoint, timeout, api_key):
        self.endpoint = endpoint
        self.timeout = timeout
        self.api_key = api_key.strip()
        self.opener = urllib.request.build_opener()

    def lookup(self, code):
        request = urllib.request.Request(
            build_lookup_url(self.endpoint, code, self.api_key),
            headers={
                "User-Agent": USER_AGENT,
                "Accept": "application/json",
                "Accept-Encoding": "identity",
            },
        )
        try:
            with self.opener.open(request, timeout=self.timeout) as response:
                payload = response.read().decode("utf-8", "replace")
        except urllib.error.HTTPError as exc:
            if exc.code == 403:
                raise ValueError("Barcode Lookup rejected the configured API key")
            if exc.code == 404:
                return []
            if exc.code == 429:
                raise ValueError("Barcode Lookup API call limit was exceeded")
            raise ValueError(
                "Barcode Lookup request failed with HTTP {}".format(exc.code)
            )

        try:
            data = json.loads(payload)
        except ValueError as exc:
            raise ValueError("Barcode Lookup returned invalid JSON") from exc

        products = data.get("products", [])
        if not isinstance(products, list):
            raise ValueError("Barcode Lookup response did not contain a products list")
        return products


def resolve_code(client, code):
    """Resolve one exact 12-digit UPC or 13-digit EAN."""
    if len(code) not in (12, 13):
        return {
            "status": "ERROR",
            "title": "",
            "source": "BarcodeLookup.com",
            "error": "UPC/EAN must contain exactly 12 or 13 digits",
            "url": "",
        }

    products = client.lookup(code)
    titles = []
    for product in products:
        if isinstance(product, dict):
            title = str(product.get("title") or "").strip()
            if title and title not in titles:
                titles.append(title)

    if len(titles) == 1:
        return {
            "status": "OK",
            "title": titles[0],
            "source": "BarcodeLookup.com",
            "error": "",
            "url": "",
        }
    if len(titles) > 1:
        return {
            "status": "AMBIGUOUS",
            "title": "[AMBIGUOUS]",
            "source": "BarcodeLookup.com",
            "error": "Multiple distinct titles were returned for this exact code",
            "url": "",
        }
    return {
        "status": "NOT_FOUND",
        "title": "[NOT FOUND]",
        "source": "BarcodeLookup.com",
        "error": "No exact UPC/EAN result was returned",
        "url": "",
    }


def read_input(input_path):
    with input_path.open("r", encoding="utf-8-sig", newline="") as source:
        reader = csv.DictReader(source, dialect="excel-tab")
        required = {"row", "upc"}
        if not reader.fieldnames or not required.issubset(reader.fieldnames):
            raise ValueError("Input TSV must contain row and upc headers")
        return list(reader)


def write_output(output_path, results):
    with output_path.open("w", encoding="utf-8", newline="") as destination:
        writer = csv.writer(destination, dialect="excel-tab", lineterminator="\n")
        writer.writerow(["row", "upc", "status", "title", "source", "error", "url"])
        for result in results:
            writer.writerow(
                [
                    safe_field(result.get("row")),
                    safe_field(result.get("upc")),
                    safe_field(result.get("status")),
                    safe_field(result.get("title")),
                    safe_field(result.get("source")),
                    safe_field(result.get("error")),
                    safe_field(result.get("url")),
                ]
            )


def result_for_row(row, status, error, source="BarcodeLookup.com"):
    return {
        "row": row.get("row", ""),
        "upc": normalize_code(row.get("upc")),
        "status": status,
        "title": "",
        "source": source,
        "error": error,
        "url": "",
    }


def process_rows(rows, client, delay, progress):
    results = []
    cache = {}

    for index, row in enumerate(rows):
        if progress.cancelled():
            break

        code = normalize_code(row.get("upc"))
        progress.start_item(code)
        made_network_lookup = False
        try:
            if code in cache:
                resolved = dict(cache[code])
            else:
                made_network_lookup = len(code) in (12, 13)
                resolved = resolve_code(client, code)
                cache[code] = dict(resolved)
        except (OSError, ValueError, urllib.error.URLError) as exc:
            resolved = {
                "status": "ERROR",
                "title": "",
                "source": "BarcodeLookup.com",
                "error": str(exc),
                "url": "",
            }

        resolved["row"] = row.get("row", "")
        resolved["upc"] = code
        results.append(resolved)
        progress.finish_item(resolved.get("status"))

        if progress.cancelled():
            break
        if made_network_lookup and index + 1 < len(rows) and delay:
            if not progress.wait(
                delay,
                "Waiting {:.1f}s for Barcode Lookup rate limit...".format(delay),
            ):
                break

    for row in rows[len(results):]:
        results.append(result_for_row(row, "CANCELLED", "Cancelled by user"))

    return results


def main():
    if len(sys.argv) != 3:
        raise SystemExit("Usage: barcodelookup_lookup.py input.tsv output.tsv")

    input_path = Path(sys.argv[1])
    output_path = Path(sys.argv[2])
    rows = read_input(input_path)
    try:
        settings = load_settings()
        endpoint = settings.get(
            "barcodelookup", "endpoint", fallback=DEFAULT_ENDPOINT
        )
        delay = settings.getfloat(
            "barcodelookup", "request_delay_seconds", fallback=0.75
        )
        timeout = settings.getfloat(
            "barcodelookup", "timeout_seconds", fallback=45.0
        )
        api_key = settings.get(
            "barcodelookup", "api_key", fallback=""
        ).strip()
        paid_subscription = settings.getboolean(
            "barcodelookup", "paid_subscription", fallback=False
        )
        show_progress = settings.getboolean(
            "progress", "show_window", fallback=True
        )
        always_on_top = settings.getboolean(
            "progress", "always_on_top", fallback=True
        )

        if not paid_subscription:
            raise ValueError(
                "Barcode Lookup automation requires a paid API subscription; "
                "set paid_subscription = true only after activating one"
            )
        if not api_key:
            raise ValueError(
                "Barcode Lookup API key is not configured in settings.ini"
            )
        if delay < 0:
            raise ValueError("request_delay_seconds cannot be negative")
        if timeout <= 0:
            raise ValueError("timeout_seconds must be positive")

        client = BarcodeLookupClient(endpoint, timeout, api_key)
        results = run_with_progress(
            "MediaCatalog - Barcode Lookup",
            len(rows),
            lambda progress: process_rows(rows, client, delay, progress),
            enabled=show_progress,
            always_on_top=always_on_top,
        )
    except Exception as exc:
        results = [result_for_row(row, "ERROR", str(exc)) for row in rows]

    write_output(output_path, results)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
        raise SystemExit(2)
