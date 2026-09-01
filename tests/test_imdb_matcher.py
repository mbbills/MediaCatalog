import sqlite3
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent
SCRIPTS_DIR = PROJECT_ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

from config import load_settings, resolve_path
from imdb_matcher import (
    clean_upcitemdb_name,
    detect_season,
    extract_year,
    find_matches,
    title_match_candidates,
)
from imdb_lookup import database_schema_error


CASES = [
    (
        "Buena Vista Toy Story 10th Anniversary Dvd",
        "tt0114709",
    ),
    (
        "The Incredibles (DVD) (2nd Collector s Edition) [2004]",
        "tt0317705",
    ),
    (
        "Universal Studios - Curious George [DIGITAL VIDEO DISC]",
        "tt0381971",
    ),
    (
        "Finding Nemo (Two-Disc Collector s Edition)",
        "tt0266543",
    ),
    (
        "Curious George Swings into Spring (DVD)",
        "tt2776304",
    ),
    (
        "Curious George 3 - Back to the Jungle [DVD]",
        "tt4622340",
    ),
    (
        "Curious George: The Complete Sixth Season [DVD]",
        "tt0449545",
    ),
    (
        "Curious George 2 - Follow That Monkey [DVD]",
        "tt1350484",
    ),
    (
        "Curious George: A Very Monkey Christmas (2009) [DVD]",
        "tt1570964",
    ),
    (
        "Curious George: Season 8 [DVD]",
        "tt0449545",
    ),
    (
        "Curious George: The Complete Ninth Season [DVD]",
        "tt0449545",
    ),
    (
        "National Treasure 2: Book of Secrets 4K (2007)",
        "tt0465234",
    ),
    (
        "Avatar 3D (2009)",
        "tt0499549",
    ),
    (
        "Sherlock: Season Two (2012)",
        "tt1475582",
    ),
    (
        "300 (2007)",
        "tt0416449",
    ),
]


def resolve_one(conn, raw_name):
    cleaned = clean_upcitemdb_name(raw_name)
    year_hint = extract_year(raw_name)
    series_title, season_number = detect_season(cleaned)

    if series_title is not None:
        search_title = series_title
        is_season = True
    else:
        search_title = cleaned
        is_season = False

    matches = []

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
            break

    return matches[0][0] if matches else None


def main():
    settings = load_settings()
    db_path = resolve_path(settings["paths"]["imdb_database"])

    if not db_path.exists():
        raise SystemExit(
            "IMDb database not found: " + str(db_path)
        )

    conn = sqlite3.connect(db_path)

    failures = []

    try:
        schema_error = database_schema_error(conn)
        if schema_error:
            raise SystemExit(schema_error)

        for raw_name, expected in CASES:
            actual = resolve_one(conn, raw_name)

            if actual != expected:
                failures.append(
                    (raw_name, expected, actual)
                )
    finally:
        conn.close()

    if failures:
        for raw_name, expected, actual in failures:
            print(f"FAIL: {raw_name}")
            print(f"  expected: {expected}")
            print(f"  actual:   {actual}")
        raise SystemExit(1)

    print(f"PASS: {len(CASES)} matcher regression cases")


if __name__ == "__main__":
    main()
