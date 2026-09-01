import argparse
import csv
import gzip
import os
import sqlite3
from pathlib import Path

from config import load_settings, resolve_path
from imdb_matcher import normalize_title


def parse_int(value):
    if value == r"\N":
        return None

    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def parse_float(value):
    if value == r"\N":
        return None

    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def import_titles(cur, conn, source_file):
    insert_sql = """
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
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """

    batch = []
    count = 0
    batch_size = 10000

    with gzip.open(source_file, "rt", encoding="utf-8", newline="") as infile:
        reader = csv.DictReader(infile, delimiter="\t")

        for row in reader:
            primary_title = row["primaryTitle"]

            batch.append(
                (
                    row["tconst"],
                    row["titleType"],
                    primary_title,
                    row["originalTitle"],
                    normalize_title(primary_title),
                    parse_int(row["startYear"]),
                    parse_int(row["endYear"]),
                    parse_int(row["runtimeMinutes"]),
                    None if row["genres"] == r"\N" else row["genres"],
                )
            )

            if len(batch) >= batch_size:
                cur.executemany(insert_sql, batch)
                conn.commit()
                count += len(batch)
                batch.clear()
                print(f"{count:,} titles imported...")

        if batch:
            cur.executemany(insert_sql, batch)
            conn.commit()
            count += len(batch)

    return count


def import_episodes(cur, conn, source_file):
    insert_sql = """
        INSERT INTO episodes (
            tconst,
            parent_tconst,
            season_number,
            episode_number
        )
        VALUES (?, ?, ?, ?)
    """

    batch = []
    count = 0
    batch_size = 10000

    with gzip.open(source_file, "rt", encoding="utf-8", newline="") as infile:
        reader = csv.DictReader(infile, delimiter="\t")

        for row in reader:
            batch.append(
                (
                    row["tconst"],
                    row["parentTconst"],
                    parse_int(row["seasonNumber"]),
                    parse_int(row["episodeNumber"]),
                )
            )

            if len(batch) >= batch_size:
                cur.executemany(insert_sql, batch)
                conn.commit()
                count += len(batch)
                batch.clear()
                print(f"{count:,} episode relationships imported...")

        if batch:
            cur.executemany(insert_sql, batch)
            conn.commit()
            count += len(batch)

    return count


def import_ratings(cur, conn, source_file):
    insert_sql = """
        INSERT INTO ratings (
            tconst,
            average_rating,
            num_votes
        )
        VALUES (?, ?, ?)
    """

    batch = []
    count = 0
    batch_size = 10000

    with gzip.open(source_file, "rt", encoding="utf-8", newline="") as infile:
        reader = csv.DictReader(infile, delimiter="\t")

        for row in reader:
            batch.append(
                (
                    row["tconst"],
                    parse_float(row["averageRating"]),
                    parse_int(row["numVotes"]),
                )
            )

            if len(batch) >= batch_size:
                cur.executemany(insert_sql, batch)
                conn.commit()
                count += len(batch)
                batch.clear()
                print(f"{count:,} title ratings imported...")

        if batch:
            cur.executemany(insert_sql, batch)
            conn.commit()
            count += len(batch)

    return count


def main():
    parser = argparse.ArgumentParser(
        description="Build the local MediaCatalog IMDb SQLite database."
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Replace an existing IMDb database.",
    )
    args = parser.parse_args()

    settings = load_settings()

    database_file = resolve_path(
        settings["paths"]["imdb_database"]
    )
    source_dir = resolve_path(
        settings["paths"]["imdb_source_dir"]
    )

    basics_file = source_dir / settings["imdb"]["title_basics"]
    episode_file = source_dir / settings["imdb"]["title_episode"]
    ratings_file = source_dir / settings["imdb"]["title_ratings"]

    for source_file in (basics_file, episode_file, ratings_file):
        if not source_file.exists():
            raise FileNotFoundError(
                f"IMDb source file not found: {source_file}"
            )

    if database_file.exists() and not args.force:
        raise SystemExit(
            f"Database already exists: {database_file}\n"
            "Run again with --force only if you intend to rebuild it."
        )

    database_file.parent.mkdir(parents=True, exist_ok=True)

    building_file = database_file.with_suffix(
        database_file.suffix + ".building"
    )

    if building_file.exists():
        building_file.unlink()

    print(f"Building: {database_file}")
    print(f"Titles:   {basics_file}")
    print(f"Episodes: {episode_file}")
    print(f"Ratings:  {ratings_file}")
    print()

    conn = sqlite3.connect(building_file)

    try:
        cur = conn.cursor()

        # Bulk imports are significantly faster with these temporary settings.
        cur.execute("PRAGMA journal_mode = OFF")
        cur.execute("PRAGMA synchronous = OFF")
        cur.execute("PRAGMA temp_store = MEMORY")

        cur.execute(
            """
            CREATE TABLE metadata (
                key    TEXT PRIMARY KEY,
                value  TEXT NOT NULL
            )
            """
        )
        cur.execute(
            "INSERT INTO metadata (key, value) VALUES (?, ?)",
            ("schema_version", "2"),
        )

        cur.execute(
            """
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
            )
            """
        )

        title_count = import_titles(
            cur,
            conn,
            basics_file,
        )

        print()
        print("Creating title indexes...")

        cur.execute(
            """
            CREATE INDEX idx_titles_normalized
            ON titles(normalized_title)
            """
        )
        cur.execute(
            """
            CREATE INDEX idx_titles_year
            ON titles(start_year)
            """
        )
        cur.execute(
            """
            CREATE INDEX idx_titles_type
            ON titles(title_type)
            """
        )
        conn.commit()

        print()
        cur.execute(
            """
            CREATE TABLE ratings (
                tconst          TEXT PRIMARY KEY,
                average_rating  REAL,
                num_votes       INTEGER
            )
            """
        )

        ratings_count = import_ratings(
            cur,
            conn,
            ratings_file,
        )

        print()
        cur.execute(
            """
            CREATE TABLE episodes (
                tconst           TEXT PRIMARY KEY,
                parent_tconst    TEXT NOT NULL,
                season_number    INTEGER,
                episode_number   INTEGER
            )
            """
        )

        episode_count = import_episodes(
            cur,
            conn,
            episode_file,
        )

        print()
        print("Creating episode indexes...")

        cur.execute(
            """
            CREATE INDEX idx_episodes_parent_season
            ON episodes(parent_tconst, season_number)
            """
        )
        cur.execute(
            """
            CREATE INDEX idx_episodes_parent
            ON episodes(parent_tconst)
            """
        )
        conn.commit()

    finally:
        conn.close()

    # Build into a temporary file and replace the live database only after
    # the entire import succeeds.
    if database_file.exists():
        database_file.unlink()

    os.replace(building_file, database_file)

    print()
    print(f"Finished. Imported {title_count:,} IMDb titles.")
    print(
        f"Finished. Imported {episode_count:,} "
        "IMDb episode relationships."
    )
    print(f"Finished. Imported {ratings_count:,} IMDb title ratings.")
    print(f"Database: {database_file}")


if __name__ == "__main__":
    main()
