import csv
import gzip
import sqlite3
import sys
import tempfile
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent
SCRIPTS_DIR = PROJECT_ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

from build_imdb_database import import_ratings
from imdb_lookup import database_schema_error, result_from_row
from imdb_matcher import (
    clean_release_name,
    detect_season,
    find_matches,
    normalize_title,
    title_match_candidates,
    year_rank,
)


def create_tables(conn):
    conn.executescript(
        """
        CREATE TABLE metadata (
            key    TEXT PRIMARY KEY,
            value  TEXT NOT NULL
        );
        INSERT INTO metadata (key, value) VALUES ('schema_version', '2');

        CREATE TABLE titles (
            tconst            TEXT PRIMARY KEY,
            title_type        TEXT,
            primary_title     TEXT,
            original_title    TEXT,
            normalized_title  TEXT,
            start_year        INTEGER,
            end_year          INTEGER,
            runtime_minutes   INTEGER,
            genres            TEXT
        );

        CREATE TABLE ratings (
            tconst          TEXT PRIMARY KEY,
            average_rating  REAL,
            num_votes       INTEGER
        );

        CREATE TABLE episodes (
            tconst           TEXT PRIMARY KEY,
            parent_tconst    TEXT NOT NULL,
            season_number    INTEGER,
            episode_number   INTEGER
        );
        """
    )


def insert_title(conn, tconst, title, year, runtime):
    conn.execute(
        """
        INSERT INTO titles (
            tconst,
            title_type,
            primary_title,
            original_title,
            normalized_title,
            start_year,
            end_year,
            runtime_minutes,
            genres
        )
        VALUES (?, 'movie', ?, ?, ?, ?, NULL, ?, 'Action,Drama')
        """,
        (tconst, title, title, normalize_title(title), year, runtime),
    )


def insert_series(conn, tconst, title, year):
    conn.execute(
        """
        INSERT INTO titles (
            tconst,
            title_type,
            primary_title,
            original_title,
            normalized_title,
            start_year,
            end_year,
            runtime_minutes,
            genres
        )
        VALUES (?, 'tvSeries', ?, ?, ?, ?, NULL, NULL, 'Crime,Drama')
        """,
        (tconst, title, title, normalize_title(title), year),
    )


def write_test_ratings(path):
    with gzip.open(path, "wt", encoding="utf-8", newline="") as destination:
        writer = csv.writer(destination, delimiter="\t", lineterminator="\n")
        writer.writerow(["tconst", "averageRating", "numVotes"])
        writer.writerow(["tt0416449", "7.6", "900000"])
        writer.writerow(["tt33397756", "6.0", "5"])
        writer.writerow(["tt0000300", "8.5", "2000000"])


def main():
    assert year_rank(2006, 2007) == (0, 1)
    assert year_rank(2007, 2007) == (0, 0)
    assert year_rank(2004, 2007) == (1, 3)
    assert clean_release_name("Avatar 3D (2009)") == "Avatar 3D"
    assert clean_release_name("Avatar (2009) 3D") == "Avatar 3D"
    assert title_match_candidates("Avatar 3D") == ["Avatar 3D", "Avatar"]
    assert title_match_candidates("Piranha 3D") == ["Piranha 3D", "Piranha"]

    season_cases = {
        "Sherlock: Season Two": ("Sherlock", 2),
        "Sherlock: Season One": ("Sherlock", 1),
        "Sherlock: Second Season": ("Sherlock", 2),
        "Sherlock: The Complete Seventh Season": ("Sherlock", 7),
        "Sherlock - Complete Season 8": ("Sherlock", 8),
    }
    for release_title, expected in season_cases.items():
        assert detect_season(release_title) == expected

    conn = sqlite3.connect(":memory:")
    create_tables(conn)

    insert_title(conn, "tt0416449", "300", 2006, 117)
    insert_title(conn, "tt33397756", "300", 2007, None)
    insert_title(conn, "tt0000300", "300", 1980, 120)
    insert_title(conn, "tt0499549", "Avatar", 2009, 162)
    insert_title(conn, "tt0464154", "Piranha 3D", 2010, 88)
    insert_title(conn, "tt0078087", "Piranha", 1978, 94)
    insert_series(conn, "tt1475582", "Sherlock", 2010)
    conn.execute(
        """
        INSERT INTO episodes (
            tconst, parent_tconst, season_number, episode_number
        )
        VALUES ('tt1942612', 'tt1475582', 2, 1)
        """
    )

    with tempfile.TemporaryDirectory() as temporary_dir:
        ratings_file = Path(temporary_dir) / "title.ratings.tsv.gz"
        write_test_ratings(ratings_file)
        imported = import_ratings(conn.cursor(), conn, ratings_file)

    assert imported == 3
    assert database_schema_error(conn) == ""

    matches = find_matches(conn, "300", year_hint=2007)
    assert matches[0][0] == "tt0416449"
    assert matches[0][3] == 2006
    assert matches[0][6] == 900000

    result = result_from_row(matches[0])
    assert result["imdb_id"] == "tt0416449"
    assert result["year"] == 2006
    assert result["average_rating"] == 7.6
    assert result["num_votes"] == 900000

    matches = []
    for candidate in title_match_candidates("Avatar 3D"):
        matches = find_matches(conn, candidate, year_hint=2009)
        if matches:
            break
    assert matches[0][0] == "tt0499549"

    matches = []
    for candidate in title_match_candidates("Piranha 3D"):
        matches = find_matches(conn, candidate, year_hint=2010)
        if matches:
            break
    assert matches[0][0] == "tt0464154"

    series_title, season_number = detect_season("Sherlock: Season Two")
    matches = find_matches(
        conn,
        series_title,
        is_season=True,
        season_number=season_number,
    )
    assert matches[0][0] == "tt1475582"
    assert matches[0][1] == "tvSeries"
    result = result_from_row(matches[0], season_number=season_number)
    assert result["title"] == "Sherlock"
    assert result["season"] == 2

    conn.close()
    print("PASS: IMDb ratings import and vote-aware ranking")


if __name__ == "__main__":
    main()
