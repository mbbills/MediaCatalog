import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent
SCRIPTS_DIR = PROJECT_ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

from imdb_matcher import clean_release_name, detect_season, extract_year, normalize_title


# Real Blu-ray.com/UPC release titles that were failing to resolve because
# clean_release_name() left product/packaging clutter in the search title.
# Each case is (raw release title, expected cleaned title, expected year hint).
CLEAN_CASES = [
    (
        "Cars 2 Blu-ray (Blu-ray + DVD) (2011)",
        "Cars 2",
        2011,
    ),
    (
        "X-Men Origins: Wolverine Blu-ray (Blu-ray + Digital) (2009)",
        "X-Men Origins: Wolverine",
        2009,
    ),
    (
        "The Angry Birds Movie Blu-ray (Blu-ray + DVD) (2016)",
        "The Angry Birds Movie",
        2016,
    ),
    (
        "Toy Story DVD (Disc 1: Toy Story) (1995)",
        "Toy Story",
        1995,
    ),
    (
        "Toy Story 2 DVD (Disc 2: Toy Story 2) (1999)",
        "Toy Story 2",
        1999,
    ),
    (
        "Schindler's List DVD (DigiBook) (1993)",
        "Schindler's List",
        1993,
    ),
    (
        "The Sum of All Fears DVD (Special Collector's Edition) (2002)",
        "The Sum of All Fears",
        2002,
    ),
    (
        "16 Blocks DVD (Widescreen Edition) (2006)",
        "16 Blocks",
        2006,
    ),
    (
        "Beverly Hills Cop DVD (Special Collector's Edition | Widescreen Collection) (1984)",
        "Beverly Hills Cop",
        1984,
    ),
    (
        "Breaking Bad: The Complete Series Blu-ray (DigiPack) (2008-2013)",
        "Breaking Bad",
        2008,
    ),
    (
        "Dark City Blu-ray (Director's Cut) (1998)",
        "Dark City",
        1998,
    ),
    (
        "Young Catherine DVD (Warner Archive Collection) (1991)",
        "Young Catherine",
        1991,
    ),
    (
        "Jarhead DVD (Full Screen) (2005)",
        "Jarhead",
        2005,
    ),
    (
        "National Lampoon's Vacation DVD (Snap case) (1983)",
        "National Lampoon's Vacation",
        1983,
    ),
    (
        "National Lampoon's European Vacation DVD (Snap case) (1985)",
        "National Lampoon's European Vacation",
        1985,
    ),
]

# (raw release title, expected series title, expected season number)
SEASON_CASES = [
    (
        "Enterprise - The Complete Second Season (2002-2003)",
        "Enterprise",
        2,
    ),
    (
        "Psych: The Eighth and Final Season (2014)",
        "Psych",
        8,
    ),
]


def check_cleanup_cases():
    for raw_name, expected_title, expected_year in CLEAN_CASES:
        cleaned = clean_release_name(raw_name)
        series_title, _season_number = detect_season(cleaned)
        actual_title = series_title if series_title is not None else cleaned

        assert normalize_title(actual_title) == normalize_title(expected_title), (
            f"{raw_name!r} cleaned to {actual_title!r}, "
            f"expected {expected_title!r}"
        )

        actual_year = extract_year(raw_name)
        assert actual_year == expected_year, (
            f"{raw_name!r} extracted year {actual_year!r}, "
            f"expected {expected_year!r}"
        )


def check_packaging_word_uses_whole_word_matching():
    """
    "disc" (a packaging word) must not match inside an unrelated word like
    "Disclosure" -- packaging-word matching inside a parenthesized group
    must respect word boundaries, not treat the words as substrings.
    """
    cleaned = clean_release_name("Some Movie (Disclosure Cut) (2001)")
    assert cleaned == "Some Movie (Disclosure Cut)", cleaned


def check_season_cases():
    for raw_name, expected_series, expected_season in SEASON_CASES:
        cleaned = clean_release_name(raw_name)
        series_title, season_number = detect_season(cleaned)

        assert series_title is not None, (
            f"{raw_name!r} was not detected as a season release"
        )
        assert normalize_title(series_title) == normalize_title(expected_series), (
            f"{raw_name!r} series title {series_title!r}, "
            f"expected {expected_series!r}"
        )
        assert season_number == expected_season, (
            f"{raw_name!r} season {season_number!r}, "
            f"expected {expected_season!r}"
        )


def check_multi_title_bonus_disc_is_left_unmatched():
    """
    A disc that covers more than one film has no single correct IMDb
    target. Cleanup should not invent one; it's expected that this search
    title will not exact-match anything in imdb.sqlite.
    """
    cleaned = clean_release_name(
        "Toy Story and Toy Story 2 DVD (Disc 3: Supplemental Features) (1995)"
    )
    assert cleaned == "Toy Story and Toy Story 2", cleaned


def main():
    check_cleanup_cases()
    check_season_cases()
    check_packaging_word_uses_whole_word_matching()
    check_multi_title_bonus_disc_is_left_unmatched()
    total = len(CLEAN_CASES) + len(SEASON_CASES)
    print(f"PASS: {total} title-cleanup regression cases")


if __name__ == "__main__":
    main()
