import gzip
import sys
import tempfile
import time
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent
SCRIPTS_DIR = PROJECT_ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

import install_media_catalog as installer
from config import load_settings_files
from install_media_catalog import update_runtime_python, validate_dataset


def main():
    with tempfile.TemporaryDirectory() as temp_dir:
        install_cmd = (PROJECT_ROOT / "install.cmd").read_text(encoding="utf-8")
        assert '-ProjectRoot "%~dp0."' in install_cmd
        builder_failure = install_cmd.split(
            'powershell.exe -NoProfile -ExecutionPolicy Bypass -File', 1
        )[1].split(":finish", 1)[0]
        assert "goto failed" in builder_failure
        temp = Path(temp_dir)
        settings = temp / "settings.ini"
        settings.write_text(
            "[runtime]\n# keep this comment\npython = python\n\n[paths]\n"
            "imdb_database = data/imdb.sqlite\n",
            encoding="utf-8",
        )
        update_runtime_python(settings, sys.executable)
        updated = settings.read_text(encoding="utf-8")
        assert "# keep this comment" in updated
        assert "imdb_database = data/imdb.sqlite" in updated
        assert str(Path(sys.executable).resolve()).replace("\\", "/") in updated

        defaults = temp / "settings.example.ini"
        defaults.write_text(
            "[runtime]\npython = python\n\n[imdb]\n"
            "title_basics = title.basics.tsv.gz\n"
            "title_episode = title.episode.tsv.gz\n"
            "title_ratings = title.ratings.tsv.gz\n",
            encoding="utf-8",
        )
        legacy_settings = temp / "legacy.ini"
        legacy_settings.write_text(
            "[runtime]\npython = C:/Python38/python.exe\n\n[imdb]\n"
            "title_basics = old.basics.tsv.gz\n",
            encoding="utf-8",
        )
        merged = load_settings_files(defaults, legacy_settings)
        assert merged["runtime"]["python"] == "C:/Python38/python.exe"
        assert merged["imdb"]["title_basics"] == "old.basics.tsv.gz"
        assert merged["imdb"]["title_ratings"] == "title.ratings.tsv.gz"

        dataset = temp / "title.ratings.tsv.gz"
        with gzip.open(dataset, "wb") as output:
            output.write(b"tconst\taverageRating\tnumVotes\n")
            for number in range(40):
                output.write(
                    "tt{:07d}\t5.7\t{}\n".format(number, number + 10).encode("ascii")
                )
        assert dataset.stat().st_size >= 100
        assert validate_dataset(
            dataset,
            b"tconst\taverageRating\tnumVotes",
        )

        # install.cmd offers to skip the IMDb database step entirely, and
        # wires the choice through to install_media_catalog.py's
        # --skip-database flag.
        assert "Download and build IMDb database?" in install_cmd
        assert "--skip-database" in install_cmd

        check_skip_unnecessary_rebuild(temp)

    print("PASS: installer settings preservation and dataset validation")


def check_skip_unnecessary_rebuild(temp):
    """A second install run must not re-download or rebuild anything that
    is already present and valid -- this is what --skip-database exists
    to guarantee explicitly, but it should already be true implicitly."""
    original_root = installer.PROJECT_ROOT
    original_source_dir = installer.SOURCE_DIR
    original_database_file = installer.DATABASE_FILE
    try:
        installer.PROJECT_ROOT = temp
        installer.SOURCE_DIR = temp / "data" / "source"
        installer.DATABASE_FILE = temp / "data" / "imdb.sqlite"
        installer.SOURCE_DIR.mkdir(parents=True, exist_ok=True)
        installer.DATABASE_FILE.parent.mkdir(parents=True, exist_ok=True)

        dataset_paths = []
        for filename, header in installer.DATASETS.items():
            path = installer.SOURCE_DIR / filename
            with gzip.open(path, "wb", compresslevel=0) as output:
                output.write(header + b"\n")
                output.write(b"filler\tdata\trow\n" * 20)
            assert path.stat().st_size >= 100, (
                "test dataset must clear validate_dataset's 100-byte floor "
                "to exercise the real skip path, not the too-small-to-"
                "validate path"
            )
            dataset_paths.append(path)

        installer.DATABASE_FILE.write_bytes(b"fake database contents")
        time.sleep(0.05)

        assert installer.database_needs_build(dataset_paths, force=False) is False
        assert installer.database_needs_build(dataset_paths, force=True) is True

        def fail_if_called(*args, **kwargs):
            raise AssertionError(
                "download_dataset must not touch the network for an "
                "already-valid dataset"
            )

        original_urlopen = installer.urllib.request.urlopen
        installer.urllib.request.urlopen = fail_if_called
        try:
            for filename, header in installer.DATASETS.items():
                installer.download_dataset(filename, header)
        finally:
            installer.urllib.request.urlopen = original_urlopen
    finally:
        installer.PROJECT_ROOT = original_root
        installer.SOURCE_DIR = original_source_dir
        installer.DATABASE_FILE = original_database_file


if __name__ == "__main__":
    main()
