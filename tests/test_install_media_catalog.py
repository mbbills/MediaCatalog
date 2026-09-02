import gzip
import sys
import tempfile
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent
SCRIPTS_DIR = PROJECT_ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

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

    print("PASS: installer settings preservation and dataset validation")


if __name__ == "__main__":
    main()
