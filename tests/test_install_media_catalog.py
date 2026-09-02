import gzip
import sys
import tempfile
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent
SCRIPTS_DIR = PROJECT_ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

from install_media_catalog import update_runtime_python, validate_dataset


def main():
    with tempfile.TemporaryDirectory() as temp_dir:
        install_cmd = (PROJECT_ROOT / "install.cmd").read_text(encoding="utf-8")
        assert '-ProjectRoot "%~dp0."' in install_cmd
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
