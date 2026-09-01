import csv
import sys
import tempfile
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent
SCRIPTS_DIR = PROJECT_ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

from bluray_lookup_excel import (
    ResultParser,
    normalize_code,
    parse_release_title,
    process_rows,
    resolve_code,
    resolve_code_with_url_fallback,
    write_output,
)


class FakeClient:
    def __init__(self, results, release_html=None):
        self.results = results
        self.calls = []
        self.request_calls = []
        self.release_html = release_html or b"""
        <title>Example Movie Blu-ray Release Date May 1, 2010</title>
        <span class=\"subheading grey\">Example Studio | 2001 | 99 min</span>
        """

    def search(self, code, catalog):
        self.calls.append((code, catalog))
        return self.results.get(catalog, [])

    def request(self, url, referer=None):
        self.request_calls.append((url, referer))
        return self.release_html


class CancelAfterOne:
    def __init__(self):
        self.is_cancelled = False

    def cancelled(self):
        return self.is_cancelled

    def start_item(self, current, phase=""):
        pass

    def finish_item(self, status):
        self.is_cancelled = True

    def wait(self, seconds, phase=""):
        return False


def main():
    assert normalize_code("097368122840") == "097368122840"
    assert normalize_code("0-97368-12284-0") == "097368122840"

    dvd_html = """
    <a class="hoverlink cover" href="/dvd/Example-Movie-DVD/123/"
       title="Example Movie (2001)">Example</a>
    <a class="other" href="/dvd/Wrong/456/" title="Wrong">Wrong</a>
    """
    parser = ResultParser("/dvd/")
    parser.feed(dvd_html)
    assert parser.titles == ["Example Movie (2001)"]
    assert parser.matches == [
        {
            "title": "Example Movie (2001)",
            "url": "https://www.blu-ray.com/dvd/Example-Movie-DVD/123/",
        }
    ]

    duplicate_html = dvd_html + """
    <a class="hoverlink" href="/dvd/Another-Edition/999/"
       title="Example Movie (2001)">Duplicate title</a>
    """
    parser = ResultParser("/dvd/")
    parser.feed(duplicate_html)
    assert len(parser.matches) == 2

    dvd_match = {
        "title": "Example Movie (2001)",
        "url": "https://www.blu-ray.com/dvd/Example-Movie-DVD/123/",
    }
    bluray_match = {
        "title": "Example Movie (2001)",
        "url": "https://www.blu-ray.com/movies/Example-Movie-Blu-ray/789/",
    }

    client = FakeClient({"DVD": [dvd_match]})
    result = resolve_code(client, "097368122840")
    assert result["status"] == "OK"
    assert result["title"] == "Example Movie (2001)"
    assert result["url"] == dvd_match["url"]
    assert client.calls == [("097368122840", "DVD")]

    client = FakeClient({"DVD": [], "Blu-ray": [bluray_match]})
    result = resolve_code(client, "097368122840")
    assert result["source"] == "Blu-ray.com Blu-ray"
    assert result["url"] == bluray_match["url"]
    assert client.calls == [
        ("097368122840", "DVD"),
        ("097368122840", "Blu-ray"),
    ]

    client = FakeClient(
        {
            "DVD": [
                {"title": "One", "url": "https://www.blu-ray.com/dvd/One/1/"},
                {"title": "Two", "url": "https://www.blu-ray.com/dvd/Two/2/"},
            ]
        }
    )
    assert resolve_code(client, "097368122840")["status"] == "AMBIGUOUS"

    result = resolve_code(None, "12345678")
    assert result["status"] == "ERROR"

    release_html = """
    <title>National Treasure 2: Book of Secrets Blu-ray Release Date May 20, 2008</title>
    <span class="subheading grey">Disney | 2007 | 125 min</span>
    """
    assert parse_release_title(release_html) == (
        "National Treasure 2: Book of Secrets (2007)"
    )

    client = FakeClient({"DVD": [dvd_match]})
    result = resolve_code_with_url_fallback(
        client,
        "097368122840",
        "https://www.blu-ray.com/movies/Wrong-Release/999/",
    )
    assert result["url"] == dvd_match["url"]
    assert client.request_calls == []

    client = FakeClient({"DVD": [], "Blu-ray": []})
    supplied_url = "https://www.blu-ray.com/movies/Example-Movie-Blu-ray/123/"
    result = resolve_code_with_url_fallback(
        client, "097368122840", supplied_url
    )
    assert result["status"] == "OK"
    assert result["title"] == "Example Movie (2001)"
    assert result["url"] == supplied_url
    assert result["source"] == "No UPC/EAN, URL OK"

    client = FakeClient({})
    result = resolve_code_with_url_fallback(client, "", supplied_url)
    assert result["status"] == "OK"
    assert result["source"] == "No UPC/EAN, URL OK"

    # Real-world regression: this UPC returns conflicting Blu-ray.com titles.
    # A user-supplied individual release URL must override [AMBIGUOUS].
    ambiguous_upc = "027616857804"
    definitive_url = (
        "https://www.blu-ray.com/dvd/When-Harry-Met-Sally-DVD/123/"
    )
    client = FakeClient(
        {
            "DVD": [
                {
                    "title": "When Harry Met Sally... (1989)",
                    "url": "https://www.blu-ray.com/dvd/First-Result/1/",
                },
                {
                    "title": "When Harry Met Sally: Special Edition (1989)",
                    "url": "https://www.blu-ray.com/dvd/Second-Result/2/",
                },
            ]
        },
        release_html=b"""
        <title>When Harry Met Sally... DVD Release Date January 9, 2001</title>
        <span class=\"subheading grey\">MGM | 1989 | 96 min</span>
        """,
    )
    result = resolve_code_with_url_fallback(
        client, ambiguous_upc, definitive_url
    )
    assert result["status"] == "OK"
    assert result["title"] == "When Harry Met Sally... (1989)"
    assert result["url"] == definitive_url
    assert result["source"] == "No UPC/EAN, URL OK"
    assert client.calls == [(ambiguous_upc, "DVD")]
    assert client.request_calls == [(definitive_url, "https://www.blu-ray.com/")]

    client = FakeClient({"DVD": [dvd_match]})
    cancelled_rows = process_rows(
        [
            {"row": "2", "upc": "097368122840", "url": ""},
            {"row": "3", "upc": "012345678905", "url": ""},
        ],
        45.0,
        0,
        CancelAfterOne(),
        client_factory=lambda timeout: client,
    )
    assert [row["status"] for row in cancelled_rows] == ["OK", "CANCELLED"]

    with tempfile.TemporaryDirectory() as temporary_dir:
        output_path = Path(temporary_dir) / "results.tsv"
        write_output(
            output_path,
            [
                {
                    "row": 2,
                    "upc": "097368122840",
                    "status": "OK",
                    "title": dvd_match["title"],
                    "source": "Blu-ray.com DVD",
                    "error": "",
                    "url": dvd_match["url"],
                }
            ],
        )
        with output_path.open(encoding="utf-8", newline="") as source:
            row = next(csv.DictReader(source, dialect="excel-tab"))
        assert row["title"] == dvd_match["title"]
        assert row["url"] == dvd_match["url"]

    print("PASS: Blu-ray.com bridge parser tests")


if __name__ == "__main__":
    main()
