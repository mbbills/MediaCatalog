#!/usr/bin/env python3
"""First-run installer for MediaCatalog's local configuration and IMDb data."""

import argparse
import gzip
import os
import re
import shutil
import subprocess
import sys
import urllib.request
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent
SETTINGS_EXAMPLE = PROJECT_ROOT / "settings.example.ini"
SETTINGS_FILE = PROJECT_ROOT / "settings.ini"
SOURCE_DIR = PROJECT_ROOT / "data" / "source"
DATABASE_FILE = PROJECT_ROOT / "data" / "imdb.sqlite"
DATASET_BASE = "https://datasets.imdbws.com/"
DATASETS = {
    "title.basics.tsv.gz": b"tconst\ttitleType\tprimaryTitle",
    "title.episode.tsv.gz": b"tconst\tparentTconst\tseasonNumber",
    "title.ratings.tsv.gz": b"tconst\taverageRating\tnumVotes",
}
USER_AGENT = "MediaCatalog installer/0.4.1"


def update_runtime_python(settings_path, executable):
    text = settings_path.read_text(encoding="utf-8")
    replacement = str(Path(executable).resolve()).replace("\\", "/")

    section_pattern = re.compile(
        r"(^\[runtime\][\s\S]*?^python\s*=\s*).*$",
        flags=re.MULTILINE,
    )
    if not section_pattern.search(text):
        raise ValueError("[runtime] python setting was not found")

    text = section_pattern.sub(
        lambda match: match.group(1) + replacement,
        text,
        count=1,
    )
    settings_path.write_text(text, encoding="utf-8")


def ensure_settings():
    if not SETTINGS_FILE.exists():
        if not SETTINGS_EXAMPLE.exists():
            raise FileNotFoundError(
                "settings.example.ini was not found: {}".format(
                    SETTINGS_EXAMPLE
                )
            )
        shutil.copy2(SETTINGS_EXAMPLE, SETTINGS_FILE)
        print("Created settings.ini from settings.example.ini")
    else:
        print("Keeping existing settings.ini")

    update_runtime_python(SETTINGS_FILE, sys.executable)
    print("Configured Python: {}".format(sys.executable))


def validate_dataset(path, expected_header):
    try:
        if path.stat().st_size < 100:
            return False
        with gzip.open(path, "rb") as source:
            header = source.readline(4096).rstrip(b"\r\n")
        return header.startswith(expected_header)
    except (OSError, EOFError):
        return False


def download_dataset(filename, expected_header):
    SOURCE_DIR.mkdir(parents=True, exist_ok=True)
    destination = SOURCE_DIR / filename

    if destination.exists() and validate_dataset(destination, expected_header):
        print("Dataset OK: {}".format(filename))
        return destination

    if destination.exists():
        print("Replacing incomplete or invalid dataset: {}".format(filename))

    temporary = destination.with_suffix(destination.suffix + ".part")
    url = DATASET_BASE + filename
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})

    print("Downloading {}".format(url))
    downloaded = 0
    with urllib.request.urlopen(request, timeout=120) as response:
        total_text = response.headers.get("Content-Length")
        total = int(total_text) if total_text and total_text.isdigit() else 0

        with temporary.open("wb") as output:
            while True:
                block = response.read(1024 * 1024)
                if not block:
                    break
                output.write(block)
                downloaded += len(block)
                if total:
                    percent = downloaded * 100.0 / total
                    print(
                        "\r  {:6.1f}% ({:.1f}/{:.1f} MiB)".format(
                            percent,
                            downloaded / 1048576.0,
                            total / 1048576.0,
                        ),
                        end="",
                        flush=True,
                    )
                else:
                    print(
                        "\r  {:.1f} MiB".format(downloaded / 1048576.0),
                        end="",
                        flush=True,
                    )
    print()

    if not validate_dataset(temporary, expected_header):
        temporary.unlink(missing_ok=True)
        raise ValueError("Downloaded dataset failed validation: {}".format(filename))

    os.replace(str(temporary), str(destination))
    print("Installed dataset: {}".format(filename))
    return destination


def check_free_space():
    free = shutil.disk_usage(PROJECT_ROOT).free
    recommended = 3 * 1024 * 1024 * 1024
    if free < recommended:
        print(
            "WARNING: only {:.1f} GiB is free. Building the IMDb database may "
            "require about 3 GiB.".format(free / 1073741824.0)
        )


def database_needs_build(dataset_paths, force):
    if force or not DATABASE_FILE.exists():
        return True
    return any(path.stat().st_mtime > DATABASE_FILE.stat().st_mtime for path in dataset_paths)


def build_database(force=False):
    dataset_paths = [
        download_dataset(filename, expected_header)
        for filename, expected_header in DATASETS.items()
    ]

    if not database_needs_build(dataset_paths, force):
        print("IMDb database is current: {}".format(DATABASE_FILE))
        return

    build_script = PROJECT_ROOT / "scripts" / "build_imdb_database.py"
    print("Building the local IMDb database. This can take several minutes...")
    command = [sys.executable, "-E", str(build_script), "--force"]
    subprocess.run(command, cwd=str(PROJECT_ROOT), check=True)

    if not DATABASE_FILE.exists():
        raise RuntimeError("IMDb database builder completed without creating imdb.sqlite")


def main():
    parser = argparse.ArgumentParser(
        description="Configure MediaCatalog and install its IMDb data."
    )
    parser.add_argument(
        "--force-database",
        action="store_true",
        help="Rebuild imdb.sqlite even when it appears current.",
    )
    args = parser.parse_args()

    if sys.version_info < (3, 8):
        raise SystemExit("MediaCatalog requires Python 3.8 or newer")

    print("MediaCatalog integrated installer")
    print("Project folder: {}".format(PROJECT_ROOT))
    check_free_space()
    ensure_settings()
    build_database(force=args.force_database)

    print()
    print("MediaCatalog data installation completed successfully.")
    print("Next: read README.md, open a template, enter catalog data,")
    print("select the rows, and run Media Catalog > Resolve Selected Rows.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nInstallation cancelled. A future run will retry incomplete work.")
        raise SystemExit(130)
    except Exception as exc:
        print("\nINSTALLATION FAILED: {}".format(exc), file=sys.stderr)
        raise SystemExit(1)
