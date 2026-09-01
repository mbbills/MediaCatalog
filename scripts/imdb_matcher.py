import re
import unicodedata


ORDINALS = {
    "first": 1,
    "second": 2,
    "third": 3,
    "fourth": 4,
    "fifth": 5,
    "sixth": 6,
    "seventh": 7,
    "eighth": 8,
    "ninth": 9,
    "tenth": 10,
    "eleventh": 11,
    "twelfth": 12,
    "thirteenth": 13,
    "fourteenth": 14,
    "fifteenth": 15,
    "sixteenth": 16,
    "seventeenth": 17,
    "eighteenth": 18,
    "nineteenth": 19,
    "twentieth": 20,
}

CARDINALS = {
    "one": 1,
    "two": 2,
    "three": 3,
    "four": 4,
    "five": 5,
    "six": 6,
    "seven": 7,
    "eight": 8,
    "nine": 9,
    "ten": 10,
    "eleven": 11,
    "twelve": 12,
    "thirteen": 13,
    "fourteen": 14,
    "fifteen": 15,
    "sixteen": 16,
    "seventeen": 17,
    "eighteen": 18,
    "nineteen": 19,
    "twenty": 20,
}

SEASON_NUMBERS = dict(CARDINALS)
SEASON_NUMBERS.update(ORDINALS)
SEASON_WORD_PATTERN = "|".join(
    sorted(SEASON_NUMBERS, key=len, reverse=True)
)


def normalize_title(text):
    """
    Convert a title to the normalized form stored in imdb.sqlite.

    This is used only for matching. IMDb's original title is preserved
    and returned to the catalog unchanged.
    """
    text = unicodedata.normalize("NFKD", text)
    text = text.casefold()
    text = re.sub(r"[^a-z0-9]+", " ", text)

    return " ".join(text.split())


def extract_year(text):
    """
    Extract a year hint when a release title includes one in parentheses or
    brackets, such as "(2009)" or "[2004]".
    """
    match = re.search(
        r"[\[(](19\d{2}|20\d{2})[\])]",
        text,
    )

    if match:
        return int(match.group(1))

    return None


def clean_release_name(text):
    """
    Create a temporary IMDb search title from a physical-media release name.

    The raw release title stored in the spreadsheet is never modified.
    Only obvious physical-media/distributor clutter is removed here.
    """
    cleaned = text.strip()

    # Distributor prefixes observed in physical-media UPC data.
    cleaned = re.sub(
        r"^Universal Studios\s*-\s*",
        "",
        cleaned,
        flags=re.IGNORECASE,
    )

    cleaned = re.sub(
        r"^Buena Vista\s+",
        "",
        cleaned,
        flags=re.IGNORECASE,
    )

    packaging_words = (
        "dvd",
        "disc",
        "disk",
        "edition",
        "collector",
        "anniversary",
        "digital video",
        "blu-ray",
        "bluray",
        "widescreen",
        "fullscreen",
    )

    def remove_packaging_group(match):
        inner = (match.group(1) or match.group(2)).strip()

        # A year can be useful as a matching hint but is not part of
        # the search title itself.
        if re.fullmatch(r"(19|20)\d{2}", inner):
            return ""

        if any(word in inner.casefold() for word in packaging_words):
            return ""

        return match.group(0)

    cleaned = re.sub(
        r"\(([^()]*)\)|\[([^\[\]]*)\]",
        remove_packaging_group,
        cleaned,
    )

    # Example:
    #   Buena Vista Toy Story 10th Anniversary Dvd
    # becomes:
    #   Toy Story
    cleaned = re.sub(
        r"\s+\d+(?:st|nd|rd|th)\s+Anniversary\b.*$",
        "",
        cleaned,
        flags=re.IGNORECASE,
    )

    # Blu-ray.com places this UHD marker immediately before the parenthesized
    # year in some result titles. After the year group has been removed above,
    # 4K is a trailing product marker rather than title text.
    cleaned = re.sub(
        r"\s+4K\s*$",
        "",
        cleaned,
        flags=re.IGNORECASE,
    )

    cleaned = re.sub(r"\s+", " ", cleaned)

    return cleaned.strip(" -")


def title_match_candidates(search_title, is_season=False):
    """
    Return narrowly scoped exact-title candidates in lookup order.

    Physical-media listings sometimes insert the sequel number before a
    subtitle even though IMDb omits it, for example:

        National Treasure 2: Book of Secrets
        National Treasure: Book of Secrets

    The original title always remains the first candidate. The alternate is
    tried only when the original exact-title search returns no result.
    """
    candidates = [search_title]

    if not is_season:
        # 3D is often a Blu-ray.com format marker, but it is also part of
        # legitimate IMDb titles such as Piranha 3D. Preserve the exact title
        # as the first attempt and remove the marker only as a no-match fallback.
        without_3d = re.sub(
            r"\s+3D\s*$",
            "",
            search_title,
            flags=re.IGNORECASE,
        ).strip()
        if (
            without_3d
            and normalize_title(without_3d) != normalize_title(search_title)
        ):
            candidates.append(without_3d)

        for candidate in list(candidates):
            match = re.fullmatch(
                r"(.+?)\s+2\s*:\s*(.+)",
                candidate,
            )

            if match:
                alternate = (
                    match.group(1).strip() + ": " + match.group(2).strip()
                )
                if all(
                    normalize_title(alternate) != normalize_title(existing)
                    for existing in candidates
                ):
                    candidates.append(alternate)

    return candidates


def clean_upcitemdb_name(text):
    """Backward-compatible name retained for the Calc v0.1.x test suite."""
    return clean_release_name(text)


def detect_season(text):
    """
    Detect TV-season releases.

    Returns:
        (series_title, season_number)

    or:
        (None, None)
    """
    patterns = (
        # Sherlock: Season Two; Curious George: Complete Season 8
        r"(?:\s*[:-]\s*|\s+)"
        r"(?:The\s+Complete\s+|Complete\s+)?"
        r"Season\s+(?P<number>\d+|" + SEASON_WORD_PATTERN + r")\b.*$",
        # Seventh Season; The Complete Ninth Season; The Second Season
        r"(?:\s*[:-]\s*|\s+)"
        r"(?:The\s+Complete\s+|Complete\s+|The\s+)?"
        r"(?P<number>" + SEASON_WORD_PATTERN + r")\s+Season\b.*$",
    )

    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if not match:
            continue

        number_text = match.group("number").casefold()
        if number_text.isdigit():
            season_number = int(number_text)
        else:
            season_number = SEASON_NUMBERS[number_text]

        series_title = text[:match.start()].strip(" :-")
        if series_title:
            return series_title, season_number

    return None, None


def type_rank(title_type, is_season):
    """
    Rank IMDb title types according to what normally makes sense for
    physical-media releases.
    """
    if is_season:
        ranking = {
            "tvSeries": 0,
            "tvMiniSeries": 1,
            "tvEpisode": 20,
            "movie": 30,
        }
    else:
        ranking = {
            "movie": 0,
            "tvMovie": 1,
            "video": 2,
            "tvSpecial": 3,
            "tvMiniSeries": 10,
            "tvSeries": 11,
            "short": 20,
            "tvEpisode": 30,
        }

    return ranking.get(title_type, 50)


def year_rank(year, year_hint):
    """
    Treat a one-year discrepancy as equivalent for matching purposes.

    IMDb may use the earliest screening year while a physical-media catalog
    uses the wider theatrical-release year. Larger differences remain
    increasingly strong negative signals.
    """
    if year_hint is None:
        return 0, 0

    if year is None:
        return 4, 9999

    distance = abs(year - year_hint)

    if distance <= 1:
        band = 0
    elif distance <= 3:
        band = 1
    elif distance <= 10:
        band = 2
    else:
        band = 3

    return band, distance


def find_matches(
    conn,
    search_title,
    year_hint=None,
    is_season=False,
    season_number=None,
):
    """
    Find and rank exact normalized IMDb title matches.

    For season releases, a candidate series must actually contain the
    requested season in IMDb's episode relationship data. Within the same
    year-confidence band and title type, IMDb vote count is used as a
    popularity signal; average rating is retained for diagnostics but is not
    used to choose a result.
    """
    normalized = normalize_title(search_title)

    if is_season and season_number is not None:
        rows = conn.execute(
            """
            SELECT
                t.tconst,
                t.title_type,
                t.primary_title,
                t.start_year,
                t.runtime_minutes,
                r.average_rating,
                COALESCE(r.num_votes, 0)
            FROM titles AS t
            LEFT JOIN ratings AS r
              ON r.tconst = t.tconst
            WHERE t.normalized_title = ?
              AND t.title_type IN ('tvSeries', 'tvMiniSeries')
              AND EXISTS (
                  SELECT 1
                  FROM episodes AS e
                  WHERE e.parent_tconst = t.tconst
                    AND e.season_number = ?
              )
            """,
            (normalized, season_number),
        ).fetchall()
    else:
        rows = conn.execute(
            """
            SELECT
                t.tconst,
                t.title_type,
                t.primary_title,
                t.start_year,
                t.runtime_minutes,
                r.average_rating,
                COALESCE(r.num_votes, 0)
            FROM titles AS t
            LEFT JOIN ratings AS r
              ON r.tconst = t.tconst
            WHERE t.normalized_title = ?
            """,
            (normalized,),
        ).fetchall()

    def sort_key(row):
        tconst, title_type, _, year, runtime, _, num_votes = row
        year_band, year_distance = year_rank(year, year_hint)

        return (
            year_band,
            type_rank(title_type, is_season),
            -(num_votes or 0),
            0 if runtime is not None else 1,
            year_distance,
            year if year is not None else 9999,
            tconst,
        )

    return sorted(rows, key=sort_key)
