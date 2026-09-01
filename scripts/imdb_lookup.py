import json
import re
import sqlite3
import sys

from config import load_settings, resolve_path
from imdb_matcher import (
    clean_release_name,
    extract_year,
    detect_season,
    find_matches,
    title_match_candidates,
)


IMDB_ID_PATTERN = re.compile(r"(tt\d+)", re.IGNORECASE)
REQUIRED_SCHEMA_VERSION = "2"


def normalize_imdb_id(value):
    """
    Normalize user-supplied IMDb identification.

    The correction field is the IMDb ID column, but users may paste either:
        tt0114709
    or:
        https://www.imdb.com/title/tt0114709/

    Returns a normalized lowercase tconst or None when the value is blank.
    Raises ValueError when a nonblank value contains no IMDb title ID.
    """
    value = (value or "").strip()

    if not value:
        return None

    match = IMDB_ID_PATTERN.search(value)

    if not match:
        raise ValueError(
            "IMDb ID must contain an identifier such as tt0114709"
        )

    return match.group(1).lower()


def result_from_row(row, season_number=None, source="match"):
    """
    Convert one titles-table row into the spreadsheet bridge response format.
    """
    (
        tconst,
        title_type,
        title,
        year,
        runtime,
        average_rating,
        num_votes,
    ) = row

    return {
        "success": True,
        "imdb_id": tconst,
        "imdb_url": f"https://www.imdb.com/title/{tconst}/",
        "title": title,
        "year": year,
        "runtime": runtime,
        "title_type": title_type,
        "season": season_number,
        "average_rating": average_rating,
        "num_votes": num_votes,
        "source": source,
    }


def database_schema_error(conn):
    """Return a rebuild instruction when imdb.sqlite predates v0.2.4."""
    try:
        row = conn.execute(
            "SELECT value FROM metadata WHERE key = ?",
            ("schema_version",),
        ).fetchone()

        ratings_table = conn.execute(
            """
            SELECT 1
            FROM sqlite_master
            WHERE type = 'table' AND name = 'ratings'
            """
        ).fetchone()
    except sqlite3.OperationalError:
        row = None
        ratings_table = None

    if row is None or row[0] != REQUIRED_SCHEMA_VERSION or ratings_table is None:
        return (
            "IMDb database predates MediaCatalog v0.2.4. Download "
            "title.ratings.tsv.gz and rebuild with: "
            "python scripts\\build_imdb_database.py --force"
        )

    return ""


def lookup_by_imdb_id(conn, imdb_id, season_number=None):
    """
    Resolve an exact IMDb title ID.

    A user-supplied IMDb ID is authoritative. No title/year/runtime
    heuristic is used once an ID has been supplied.
    """
    row = conn.execute(
        """
        SELECT
            t.tconst,
            t.title_type,
            t.primary_title,
            t.start_year,
            t.runtime_minutes,
            r.average_rating,
            COALESCE(r.num_votes, 0)
        FROM titles AS t
        LEFT JOIN ratings AS r
          ON r.tconst = t.tconst
        WHERE t.tconst = ?
        """,
        (imdb_id,),
    ).fetchone()

    if row is None:
        return {
            "success": False,
            "imdb_id": imdb_id,
            "season": season_number,
            "error": f"IMDb ID not found in local database: {imdb_id}",
        }

    return result_from_row(
        row,
        season_number=season_number,
        source="imdb_id",
    )


def lookup(raw_name, imdb_id_hint=None):
    """
    Resolve one catalog row.

    Normal path:
        Blu-ray.com release title -> cleaned title -> ranked IMDb match

    Correction path:
        Existing/user-entered IMDb ID -> exact local IMDb lookup

    The IMDb ID is deliberately the single authoritative correction field.
    """
    settings = load_settings()

    database_file = resolve_path(
        settings["paths"]["imdb_database"]
    )

    if not database_file.exists():
        return {
            "success": False,
            "error": f"IMDb database not found: {database_file}",
        }

    cleaned = clean_release_name(raw_name)
    year_hint = extract_year(raw_name)

    series_title, season_number = detect_season(cleaned)

    try:
        normalized_id = normalize_imdb_id(imdb_id_hint)
    except ValueError as exc:
        return {
            "success": False,
            "season": season_number,
            "error": str(exc),
        }

    conn = sqlite3.connect(database_file)

    try:
        schema_error = database_schema_error(conn)
        if schema_error:
            return {
                "success": False,
                "season": season_number,
                "error": schema_error,
            }

        # Gold-standard correction path.
        if normalized_id is not None:
            return lookup_by_imdb_id(
                conn,
                normalized_id,
                season_number=season_number,
            )

        # Automatic matching path.
        if series_title is not None:
            search_title = series_title
            is_season = True
        else:
            search_title = cleaned
            is_season = False

        matches = []
        matched_search_title = search_title

        for candidate in title_match_candidates(
            search_title,
            is_season=is_season,
        ):
            matches = find_matches(
                conn,
                candidate,
                year_hint=year_hint,
                is_season=is_season,
                season_number=season_number,
            )

            if matches:
                matched_search_title = candidate
                break

    finally:
        conn.close()

    if not matches:
        return {
            "success": False,
            "search_title": search_title,
            "season": season_number,
            "error": "No exact IMDb match found",
        }

    result = result_from_row(
        matches[0],
        season_number=season_number,
        source="match",
    )
    result["search_title"] = matched_search_title

    if matched_search_title != search_title:
        without_3d = re.sub(
            r"\s+3D\s*$",
            "",
            search_title,
            flags=re.IGNORECASE,
        ).strip()
        if matched_search_title == without_3d:
            result["source"] = "match_3d_removed"
        else:
            result["source"] = "match_sequel_number_removed"

    return result


def main():
    if len(sys.argv) < 2:
        print(
            json.dumps(
                {
                    "success": False,
                    "error": "Release title was not supplied",
                }
            )
        )
        raise SystemExit(1)

    raw_name = sys.argv[1]
    imdb_id_hint = sys.argv[2] if len(sys.argv) >= 3 else None

    result = lookup(
        raw_name,
        imdb_id_hint=imdb_id_hint,
    )

    print(
        json.dumps(
            result,
            ensure_ascii=False,
        )
    )


if __name__ == "__main__":
    main()
