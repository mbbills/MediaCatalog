#!/usr/bin/env python3
"""Shared Blu-ray.com UPC/EAN and release-detail parser for desktop clients."""

import csv
import datetime
import html
import re
import sys
import traceback
import urllib.error
import urllib.parse
import urllib.request
from html.parser import HTMLParser
from http.cookiejar import CookieJar
from pathlib import Path

from config import load_settings
from job_progress import run_with_progress


HOME = "https://www.blu-ray.com/"
SEARCHES = {
    "DVD": ("https://www.blu-ray.com/dvd/search.php", "/dvd/"),
    "Blu-ray": ("https://www.blu-ray.com/movies/search.php", "/movies/"),
}
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:115.0) "
    "Gecko/20100101 Firefox/115.0"
)


def safe_field(value):
    """Keep a value on one TSV line."""
    if value is None:
        return ""
    return str(value).replace("\t", " ").replace("\r", " ").replace("\n", " ")


def normalize_code(value):
    """Return digits only while preserving leading zeroes."""
    value = (value or "").strip()
    return "".join(character for character in value if character.isdigit())


class ResultParser(HTMLParser):
    """Extract exact release titles and page URLs from result cards."""

    def __init__(self, path_fragment):
        HTMLParser.__init__(self)
        self.path_fragment = path_fragment
        self.titles = []
        self.matches = []

    def handle_starttag(self, tag, attrs):
        if tag != "a":
            return

        values = dict(attrs)
        classes = (values.get("class") or "").split()
        href = values.get("href") or ""
        title = values.get("title")

        if "hoverlink" in classes and self.path_fragment in href and title:
            page_url = urllib.parse.urljoin(HOME, href)
            parsed_url = urllib.parse.urlsplit(page_url)

            if (
                parsed_url.scheme in ("http", "https")
                and parsed_url.hostname in ("blu-ray.com", "www.blu-ray.com")
            ):
                clean_title = html.unescape(title).strip()
                self.titles.append(clean_title)
                self.matches.append(
                    {
                        "title": clean_title,
                        "url": page_url,
                    }
                )


class ReleaseDetailParser(HTMLParser):
    """Collect the summary line plus the Video and Disc(s) detail sections."""

    TARGET_SECTIONS = {"Video", "Disc", "Discs"}

    def __init__(self):
        HTMLParser.__init__(self)
        self.summary_parts = []
        self.sections = {}
        self.current_section = ""
        self.heading_parts = []
        self.heading_span_depth = 0
        self.summary_span_depth = 0
        self.single_disc_heading = False
        self.page_title_parts = []
        self.page_title_depth = 0
        self.imdb_ids = []

    def handle_starttag(self, tag, attrs):
        values = dict(attrs)

        if tag == "title":
            self.page_title_depth = 1

        if tag == "a" and values.get("id") == "imdb_icon":
            href = html.unescape(values.get("href") or "")
            match = re.search(
                r"https?://(?:www\.)?imdb\.com/title/(tt\d+)(?:/|[?#]|$)",
                href,
                flags=re.I,
            )
            if match:
                imdb_id = match.group(1).lower()
                if imdb_id not in self.imdb_ids:
                    self.imdb_ids.append(imdb_id)

        if tag == "span":
            classes = set((values.get("class") or "").split())

            if self.summary_span_depth:
                self.summary_span_depth += 1
            elif self.heading_span_depth:
                self.heading_span_depth += 1
            elif "subheading" in classes and "grey" in classes:
                self.summary_span_depth = 1
            elif "subheading" in classes:
                self.heading_span_depth = 1
                self.heading_parts = []

        if tag == "br":
            if self.summary_span_depth:
                self.summary_parts.append("\n")
            elif self.current_section in self.TARGET_SECTIONS:
                self.sections.setdefault(self.current_section, []).append("\n")

    def handle_endtag(self, tag):
        if tag == "title" and self.page_title_depth:
            self.page_title_depth = 0
            return

        if tag != "span":
            return

        if self.summary_span_depth:
            self.summary_span_depth -= 1
            return

        if self.heading_span_depth:
            self.heading_span_depth -= 1
            if self.heading_span_depth == 0:
                heading = normalize_text(" ".join(self.heading_parts))
                if heading == "Disc":
                    self.single_disc_heading = True
                    heading = "Discs"
                self.current_section = heading
                if heading in self.TARGET_SECTIONS:
                    self.sections.setdefault(heading, [])

    def handle_data(self, data):
        if self.page_title_depth:
            self.page_title_parts.append(data)
        elif self.summary_span_depth:
            self.summary_parts.append(data)
        elif self.heading_span_depth:
            self.heading_parts.append(data)
        elif self.current_section in self.TARGET_SECTIONS:
            self.sections.setdefault(self.current_section, []).append(data)


def normalize_text(value):
    """Collapse HTML whitespace without losing meaningful punctuation."""
    return re.sub(r"\s+", " ", html.unescape(value or "")).strip()


def section_lines(parts):
    """Turn collected text and BR markers into compact nonblank lines."""
    raw = html.unescape("".join(parts or []))
    return [normalize_text(line) for line in raw.split("\n") if normalize_text(line)]


def parse_release_date(value):
    """Return a recognized Blu-ray.com release date as ISO text."""
    value = normalize_text(value)
    for date_format in ("%b %d, %Y", "%B %d, %Y"):
        try:
            return datetime.datetime.strptime(value, date_format).date().isoformat()
        except ValueError:
            pass
    return ""


def parse_release_details(document):
    """Extract the selected stable catalog fields from one release page."""
    parser = ReleaseDetailParser()
    parser.feed(document)

    summary = normalize_text(" ".join(parser.summary_parts))
    summary_fields = [normalize_text(part) for part in summary.split("|")]
    summary_fields = [part for part in summary_fields if part]

    details = {
        "studio": "",
        "year": "",
        "runtime": "",
        "rating": "",
        "release_date": "",
        "disc_format": "",
        "video_codec": "",
        "resolution": "",
        "aspect_ratio": "",
        "disc_count_capacities": "",
    }

    if summary_fields:
        details["studio"] = summary_fields[0]

        for field in summary_fields[1:]:
            if not details["year"]:
                year_match = re.fullmatch(
                    r"(\d{4})(?:\s*[-\N{EN DASH}]\s*(\d{4}))?", field
                )
                if year_match:
                    details["year"] = year_match.group(1)
                    if year_match.group(2):
                        details["year"] += "-" + year_match.group(2)
                    continue

            if not details["runtime"]:
                runtime_match = re.search(r"\b(\d+)\s*min\b", field, re.I)
                if runtime_match:
                    details["runtime"] = runtime_match.group(1)
                    continue

            if not details["rating"]:
                if re.match(r"^Rated\s+", field, re.I):
                    details["rating"] = re.sub(
                        r"^Rated\s+", "", field, flags=re.I
                    ).strip()
                    continue
                if re.fullmatch(r"(?:Not rated|Unrated|NR)", field, re.I):
                    details["rating"] = field
                    continue

            if not details["release_date"]:
                release_date = parse_release_date(field)
                if release_date:
                    details["release_date"] = release_date

    original_aspect_ratio = ""
    for line in section_lines(parser.sections.get("Video")):
        if line.startswith("Codec:") and not details["video_codec"]:
            details["video_codec"] = normalize_text(line.split(":", 1)[1])
        elif line.startswith("Resolution:") and not details["resolution"]:
            details["resolution"] = normalize_text(line.split(":", 1)[1])
        elif line.startswith("Aspect ratio:") and not details["aspect_ratio"]:
            details["aspect_ratio"] = normalize_text(line.split(":", 1)[1])
        elif line.startswith("Original aspect ratio:") and not original_aspect_ratio:
            original_aspect_ratio = normalize_text(line.split(":", 1)[1])

    if not details["aspect_ratio"]:
        details["aspect_ratio"] = original_aspect_ratio

    disc_lines = section_lines(parser.sections.get("Discs"))
    formats = []
    for line in disc_lines:
        if line in (
            "4K Ultra HD",
            "Blu-ray 3D",
            "Blu-ray Disc",
            "DVD",
        ) and line not in formats:
            formats.append(line)
        if (
            not details["disc_count_capacities"]
            and re.search(r"\b(?:single|two|three|four|five|six|seven|eight|nine|ten|\d+)[ -]disc\b", line, re.I)
        ):
            details["disc_count_capacities"] = line

    details["disc_format"] = " + ".join(formats)
    if (
        not details["disc_count_capacities"]
        and parser.single_disc_heading
        and formats
    ):
        details["disc_count_capacities"] = "Single disc"
    return details


def parse_release_title(document):
    """Return a catalog-style title from one individual release page."""
    parser = ReleaseDetailParser()
    parser.feed(document)
    page_title = normalize_text(" ".join(parser.page_title_parts))

    title = re.sub(
        r"\s+(?:4K Ultra HD|4K Blu-ray|Blu-ray|DVD)\s+Release Date\b.*$",
        "",
        page_title,
        flags=re.I,
    ).strip()
    if title == page_title:
        title = re.sub(r"\s+-\s+Blu-ray\.com.*$", "", title, flags=re.I).strip()
        title = re.sub(
            r"\s+(?:4K Ultra HD|4K Blu-ray|Blu-ray|DVD)\s*$",
            "",
            title,
            flags=re.I,
        ).strip()

    if not title:
        raise ValueError("Release title was not found on the page")

    details = parse_release_details(document)
    year = details.get("year", "")
    if year and not re.search(r"\(\d{4}(?:-\d{4})?\)\s*$", title):
        title += " (" + year + ")"
    return title


def parse_release_imdb_identity(document):
    """Return only Blu-ray.com's structured IMDb title identity.

    Review prose and unrelated external links are intentionally ignored.  The
    page's rating-area anchor uses id="imdb_icon" for its canonical title link.
    """
    parser = ReleaseDetailParser()
    parser.feed(document)

    if len(parser.imdb_ids) == 1:
        imdb_id = parser.imdb_ids[0]
        return {
            "imdb_id": imdb_id,
            "imdb_url": "https://www.imdb.com/title/{}/".format(imdb_id),
            "imdb_warning": "",
        }

    if len(parser.imdb_ids) > 1:
        return {
            "imdb_id": "",
            "imdb_url": "",
            "imdb_warning": "Multiple structured IMDb links were found",
        }

    return {
        "imdb_id": "",
        "imdb_url": "",
        "imdb_warning": "",
    }


def parse_release_page(document):
    """Parse all supported fields from one already-downloaded release page."""
    page = parse_release_details(document)
    page["release_title"] = parse_release_title(document)
    page.update(parse_release_imdb_identity(document))
    return page


def normalize_release_url(value):
    """Accept only individual Blu-ray.com DVD or movie release pages."""
    parsed = urllib.parse.urlsplit((value or "").strip())
    if parsed.hostname not in ("blu-ray.com", "www.blu-ray.com"):
        raise ValueError("Release URL is not on Blu-ray.com")
    if not (parsed.path.startswith("/movies/") or parsed.path.startswith("/dvd/")):
        raise ValueError("URL is not an individual Blu-ray.com release page")
    return urllib.parse.urlunsplit(("https", "www.blu-ray.com", parsed.path, "", ""))


def fetch_release_details(client, release_url):
    """Fetch and parse one validated individual release page."""
    release_url = normalize_release_url(release_url)
    body = client.request(release_url, HOME)
    details = parse_release_details(body.decode("utf-8", "replace"))

    if not any(details.values()):
        raise ValueError("Release details were not found on the page")

    details["url"] = release_url
    return details


def fetch_release_page(client, release_url):
    """Fetch one page once and return title, details, and structured IMDb link."""
    release_url = normalize_release_url(release_url)
    body = client.request(release_url, HOME)
    page = parse_release_page(body.decode("utf-8", "replace"))

    if not page.get("release_title") and not any(page.values()):
        raise ValueError("Release data was not found on the page")

    page["url"] = release_url
    return page


def fetch_release_title(client, release_url):
    """Fetch a validated release page and return its catalog-style title."""
    release_url = normalize_release_url(release_url)
    body = client.request(release_url, HOME)
    title = parse_release_title(body.decode("utf-8", "replace"))
    return {
        "status": "OK",
        "title": title,
        "url": release_url,
        "source": "No UPC/EAN, URL OK",
        "error": "",
    }


class BlurayClient:
    """Maintain the cookies Blu-ray.com expects across database searches."""

    def __init__(self, timeout):
        self.timeout = timeout
        self.reset()

    def request(self, url, referer=None):
        headers = {
            "User-Agent": USER_AGENT,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.9",
        }
        if referer:
            headers["Referer"] = referer

        request = urllib.request.Request(url, headers=headers)
        with self.opener.open(request, timeout=self.timeout) as response:
            return response.read()

    def reset(self):
        self.opener = urllib.request.build_opener(
            urllib.request.HTTPCookieProcessor(CookieJar())
        )
        self.seeded = set()
        self.request(HOME)

    def search(self, code, catalog):
        base, path_fragment = SEARCHES[catalog]

        if catalog not in self.seeded:
            self.request(base, HOME)
            self.seeded.add(catalog)

        field = "ean" if len(code) == 13 else "upc"
        query = urllib.parse.urlencode({field: code, "action": "search"})
        url = base + "?" + query
        body = self.request(url, base)

        if body.strip() == b"error42":
            self.reset()
            self.request(base, HOME)
            self.seeded.add(catalog)
            body = self.request(url, base)

        if body.strip() == b"error42":
            raise ValueError("Blu-ray.com returned error42 after a fresh session")

        parser = ResultParser(path_fragment)
        parser.feed(body.decode("utf-8", "replace"))

        matches_by_title = {}
        for match in parser.matches:
            if match["title"] and match["title"] not in matches_by_title:
                matches_by_title[match["title"]] = match

        return list(matches_by_title.values())


def resolve_code(client, code):
    """Resolve only exact 12-digit UPC or 13-digit EAN searches."""
    if len(code) not in (12, 13):
        return {
            "status": "ERROR",
            "title": "",
            "url": "",
            "source": "",
            "error": "UPC/EAN must contain exactly 12 or 13 digits",
        }

    for catalog in ("DVD", "Blu-ray"):
        matches = client.search(code, catalog)

        if not matches:
            continue

        if len(matches) == 1:
            return {
                "status": "OK",
                "title": matches[0]["title"],
                "url": matches[0]["url"],
                "source": "Blu-ray.com " + catalog,
                "error": "",
            }

        return {
            "status": "AMBIGUOUS",
            "title": "[AMBIGUOUS]",
            "url": "",
            "source": "Blu-ray.com " + catalog,
            "error": "Multiple distinct titles were returned for this exact code",
        }

    return {
        "status": "NOT_FOUND",
        "title": "[NOT FOUND]",
        "url": "",
        "source": "Blu-ray.com",
        "error": "No exact UPC/EAN result in the DVD or Blu-ray catalog",
    }


def read_input(input_path):
    with input_path.open("r", encoding="utf-8-sig", newline="") as source:
        reader = csv.DictReader(source, dialect="excel-tab")
        required = {"row", "upc"}
        if not reader.fieldnames or not required.issubset(reader.fieldnames):
            raise ValueError("Input TSV must contain row and upc headers")
        return list(reader)


def resolve_code_with_url_fallback(client, code, supplied_url):
    """Try the exact UPC/EAN first, then a supplied URL if unresolved."""
    if len(code) in (12, 13):
        resolved = resolve_code(client, code)
        if resolved["status"] not in ("NOT_FOUND", "AMBIGUOUS") or not supplied_url:
            return resolved
    elif not supplied_url:
        return resolve_code(client, code)

    return fetch_release_title(client, supplied_url)


def write_output(output_path, results):
    with output_path.open("w", encoding="utf-8", newline="") as destination:
        writer = csv.writer(destination, dialect="excel-tab", lineterminator="\n")
        writer.writerow(
            ["row", "upc", "status", "title", "source", "error", "url"]
        )
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


def result_for_row(row, status, error, source="Blu-ray.com"):
    return {
        "row": row.get("row", ""),
        "upc": normalize_code(row.get("upc")),
        "status": status,
        "title": "",
        "url": "",
        "source": source,
        "error": error,
    }


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

        code = normalize_code(row.get("upc"))
        supplied_url = (row.get("url") or "").strip()
        cache_key = (code, supplied_url)
        made_network_lookup = False
        progress.start_item(code or supplied_url, "Searching Blu-ray.com...")

        try:
            if cache_key in cache:
                resolved = dict(cache[cache_key])
            else:
                if len(code) in (12, 13) or supplied_url:
                    if client is None:
                        client = client_factory(timeout)
                    made_network_lookup = True
                resolved = resolve_code_with_url_fallback(
                    client, code, supplied_url
                )
                cache[cache_key] = dict(resolved)
        except (OSError, ValueError, urllib.error.URLError) as exc:
            resolved = {
                "status": "ERROR",
                "title": "",
                "url": "",
                "source": "Blu-ray.com",
                "error": str(exc),
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
                "Waiting {:.1f}s before the next Blu-ray.com lookup...".format(
                    delay
                ),
            ):
                break

    for row in rows[len(results):]:
        results.append(result_for_row(row, "CANCELLED", "Cancelled by user"))

    return results


def main():
    if len(sys.argv) != 3:
        raise SystemExit("Usage: bluray_lookup_excel.py input.tsv output.tsv")

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
            "MediaCatalog - Blu-ray.com UPC lookup",
            len(rows),
            lambda progress: process_rows(rows, timeout, delay, progress),
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
