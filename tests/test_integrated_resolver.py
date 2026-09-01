import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent
SCRIPTS_DIR = PROJECT_ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

from bluray_lookup_excel import (
    parse_release_imdb_identity,
    parse_release_page,
)
from resolve_rows import normalize_manual_imdb_identity, resolve_row


SAMPLE_PAGE = """
<title>300 Blu-ray Release Date July 31, 2007</title>
<span class="subheading grey">
Warner Bros. | 2007 | 117 min | Rated R | Jul 31, 2007
</span>
<span class="subheading">Video</span>
Codec: VC-1<br>
Resolution: 1080p<br>
Aspect ratio: 2.40:1<br>
<span class="subheading">Discs</span>
Blu-ray Disc<br>
Single disc (1 BD-50)<br>
<a id="imdb_icon" href="https://www.imdb.com/title/tt0416449/">
  <img title="IMDb">
</a>
<a href="https://www.imdb.com/title/tt9999999/technical">review prose</a>
"""


def fake_page(client, url):
    page = parse_release_page(SAMPLE_PAGE)
    page["url"] = url
    return page


def fake_resolve(client, code):
    assert code == "012345678905"
    return {
        "status": "OK",
        "title": "300 (2007)",
        "url": "https://www.blu-ray.com/movies/300-Blu-ray/123/",
        "source": "Blu-ray.com Blu-ray",
        "error": "",
    }


def fake_imdb(raw_title, imdb_id_hint=None):
    assert imdb_id_hint in ("tt0416449", "tt7654321")
    return {
        "success": True,
        "imdb_id": imdb_id_hint,
        "imdb_url": "https://www.imdb.com/title/{}/".format(imdb_id_hint),
        "title": "300" if imdb_id_hint == "tt0416449" else "Manual Choice",
        "year": "2006",
        "runtime": "117",
        "title_type": "movie",
        "season": "",
        "source": "imdb_id",
    }


def fake_title_imdb(raw_title, imdb_id_hint=None):
    assert raw_title == "Sherlock"
    assert imdb_id_hint is None
    return {
        "success": True,
        "imdb_id": "tt1475582",
        "imdb_url": "https://www.imdb.com/title/tt1475582/",
        "title": "Sherlock",
        "year": "2010",
        "runtime": "88",
        "title_type": "tvSeries",
        "season": "",
        "source": "match",
    }


def main():
    identity = parse_release_imdb_identity(SAMPLE_PAGE)
    assert identity["imdb_id"] == "tt0416449"
    assert identity["imdb_url"].endswith("/tt0416449/")

    # A successful UPC lookup now performs all physical and IMDb enrichment.
    result = resolve_row(
        {"row": "2", "upc": "012345678905"},
        client=object(),
        resolve_code_func=fake_resolve,
        fetch_release_page_func=fake_page,
        imdb_lookup_func=fake_imdb,
    )
    assert result["status"] == "OK - Blu-ray + IMDb"
    assert result["release_title"] == "300 (2007)"
    assert result["imdb_id"] == "tt0416449"
    assert result["title"] == "300"
    assert result["year"] == "2006"
    assert result["bluray_year"] == "2007"
    assert result["studio"] == "Warner Bros."
    assert result["disc_format"] == "Blu-ray Disc"

    # A manually supplied IMDb ID wins over a conflicting Blu-ray.com link.
    result = resolve_row(
        {
            "row": "3",
            "bluray_url": "https://www.blu-ray.com/movies/300-Blu-ray/123/",
            "imdb_id": "tt7654321",
        },
        client=object(),
        fetch_release_page_func=fake_page,
        imdb_lookup_func=fake_imdb,
    )
    assert result["imdb_id"] == "tt7654321"
    assert result["title"] == "Manual Choice"
    assert "Manual IMDb override" in result["warning"]
    assert "tt0416449" in result["warning"]

    # IMDb ID also wins when the pre-existing IMDb URL disagrees.
    imdb_id, warnings = normalize_manual_imdb_identity(
        "tt7654321",
        "https://www.imdb.com/title/tt0416449/",
    )
    assert imdb_id == "tt7654321"
    assert warnings and "overrides conflicting IMDb URL" in warnings[0]

    # A title-only row stays offline and fills only canonical IMDb metadata.
    result = resolve_row(
        {"row": "4", "title": "Sherlock"},
        client=None,
        imdb_lookup_func=fake_title_imdb,
    )
    assert result["status"] == "PARTIAL - IMDb only"
    assert result["imdb_id"] == "tt1475582"
    assert not result["bluray_url"]

    # An invalid manual ID is reported and is never silently replaced.
    result = resolve_row(
        {"row": "5", "title": "Sherlock", "imdb_id": "not-an-id"},
        client=None,
        imdb_lookup_func=fake_title_imdb,
    )
    assert result["status"] == "NEEDS REVIEW"
    assert "must contain an identifier" in result["error"]
    assert result["imdb_id"] == "not-an-id"

    print("PASS: integrated resolver precedence and full-row enrichment")


if __name__ == "__main__":
    main()
