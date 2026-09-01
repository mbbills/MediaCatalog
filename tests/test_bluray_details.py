#!/usr/bin/env python3
"""Offline release-detail parser tests."""

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

from bluray_lookup_excel import normalize_release_url, parse_release_details
from bluray_details import enrich_row, process_rows


bluray_html = """
<span class="subheading grey">
<a>Warner Bros.</a> | <a>2008</a> |
<span id="runtime">152 min</span> | Rated PG-13 |
<a>Dec 19, 2017</a></span>
<span class="subheading">Video</span><br>
Codec: HEVC / H.265 (47.16 Mbps)<br>
Resolution: Native 4K (2160p)<br>
HDR: HDR10<br>
Aspect ratio: 2.40:1, 1.78:1<br>
Original aspect ratio: 2.39:1<br>
<span class="subheading">Audio</span><br>English: DTS-HD MA 5.1<br>
<span class="subheading">Discs</span><br>
4K Ultra HD<br>Blu-ray Disc<br>
Three-disc set (1 BD-100, 1 BD-25, 1 BD-50)<br>BD-Live<br>
<span class="subheading">Digital</span><br>Digital 4K<br>
"""

details = parse_release_details(bluray_html)
assert details == {
    "studio": "Warner Bros.",
    "year": "2008",
    "runtime": "152",
    "rating": "PG-13",
    "release_date": "2017-12-19",
    "disc_format": "4K Ultra HD + Blu-ray Disc",
    "video_codec": "HEVC / H.265 (47.16 Mbps)",
    "resolution": "Native 4K (2160p)",
    "aspect_ratio": "2.40:1, 1.78:1",
    "disc_count_capacities": "Three-disc set (1 BD-100, 1 BD-25, 1 BD-50)",
}

dvd_html = """
<span class="subheading grey">
<a>Disney / Buena Vista</a> | <a>2007</a> |
<span id="runtime">119 min</span> | Rated ACB: PG |
<a>May 20, 2008</a></span>
<span class="subheading">Video</span><br>
Codec: MPEG-2<br>Encoding format: 16:9<br>Resolution: 576i (PAL)<br>
Aspect ratio: 2.35:1<br>Original aspect ratio: 2.39:1<br>
<span class="subheading">Discs</span><br>
DVD<br>Single disc (1 DVD-9)<br>
<span class="subheading">Playback</span><br>Region 4<br>
"""

dvd = parse_release_details(dvd_html)
assert dvd["studio"] == "Disney / Buena Vista"
assert dvd["rating"] == "ACB: PG"
assert dvd["release_date"] == "2008-05-20"
assert dvd["disc_format"] == "DVD"
assert dvd["video_codec"] == "MPEG-2"
assert dvd["resolution"] == "576i (PAL)"
assert dvd["aspect_ratio"] == "2.35:1"
assert dvd["disc_count_capacities"] == "Single disc (1 DVD-9)"

incredibles_html = """
<span class="subheading grey">
Disney / Buena Vista | 2004 | 115 min | Rated PG | Mar 15, 2005
</span>
<span class="subheading">Video</span><br>
Codec: MPEG-2<br>Resolution: 480i (NTSC)<br>Aspect ratio: 2.39:1<br>
<span class="subheading">Discs</span><br>
DVD<br>Two-disc set (2 DVD-9)<br>
"""

incredibles = parse_release_details(incredibles_html)
assert incredibles["year"] == "2004"
assert incredibles["runtime"] == "115"
assert incredibles["rating"] == "PG"
assert incredibles["release_date"] == "2005-03-15"

curious_george_html = """
<span class="subheading grey">
Universal Studios | 2012-2013 | Season 7 | 142 min | Not rated | Jun 03, 2014
</span>
<span class="subheading">Video</span><br>
Codec: MPEG-2<br>Resolution: 480i (NTSC)<br>
Original aspect ratio: 1.78:1<br>
<span class="subheading">Disc</span><br>
DVD<br>
<span class="subheading">Playback</span><br>Region 1<br>
"""

curious_george = parse_release_details(curious_george_html)
assert curious_george == {
    "studio": "Universal Studios",
    "year": "2012-2013",
    "runtime": "142",
    "rating": "Not rated",
    "release_date": "2014-06-03",
    "disc_format": "DVD",
    "video_codec": "MPEG-2",
    "resolution": "480i (NTSC)",
    "aspect_ratio": "1.78:1",
    "disc_count_capacities": "Single disc",
}

bugs_life_html = """
<span class="subheading grey">
Disney / Buena Vista | 1998 | 1 Movie, 2 Cuts | 95 min | Rated G | May 27, 2003
</span>
<span class="subheading">Video</span><br>
Codec: MPEG-2<br>Resolution: 480i (NTSC)<br>
Aspect ratio: 1.33:1, 2.35:1<br>Original aspect ratio: 2.39:1<br>
<span class="subheading">Discs</span><br>
DVD<br>Two-disc set (1 DVD-5, 1 DVD-9)<br>
"""

bugs_life = parse_release_details(bugs_life_html)
assert bugs_life["year"] == "1998"
assert bugs_life["runtime"] == "95"
assert bugs_life["rating"] == "G"
assert bugs_life["release_date"] == "2003-05-27"
assert bugs_life["aspect_ratio"] == "1.33:1, 2.35:1"
assert bugs_life["disc_count_capacities"] == "Two-disc set (1 DVD-5, 1 DVD-9)"

three_d_html = """
<span class="subheading grey">
Disney / Buena Vista | 2009 | 96 min | Rated PG | Nov 08, 2011
</span>
<span class="subheading">Video</span><br>
Codec: MPEG-4 MVC<br>Resolution: 1080p<br>Aspect ratio: 1.78:1<br>
<span class="subheading">Discs</span><br>
Blu-ray 3D<br>Blu-ray Disc<br>
Five-disc set (3 BD-50, 2 DVDs)<br>DVD copy<br>
"""

three_d = parse_release_details(three_d_html)
assert three_d["disc_format"] == "Blu-ray 3D + Blu-ray Disc"
assert three_d["disc_count_capacities"] == (
    "Five-disc set (3 BD-50, 2 DVDs)"
)

assert normalize_release_url(
    "http://blu-ray.com/movies/Example-Blu-ray/123/?x=1#top"
) == "https://www.blu-ray.com/movies/Example-Blu-ray/123/"

try:
    normalize_release_url("https://example.com/movies/Example/123/")
except ValueError:
    pass
else:
    raise AssertionError("Non-Blu-ray.com URL was accepted")


class FakeDetailClient:
    def __init__(self, matches):
        self.matches = matches
        self.search_calls = []
        self.request_calls = []

    def search(self, code, catalog):
        self.search_calls.append((code, catalog))
        return self.matches.get(catalog, [])

    def request(self, url, referer=None):
        self.request_calls.append((url, referer))
        return bluray_html.encode("utf-8")


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


resolved_url = "https://www.blu-ray.com/movies/Resolved-Blu-ray/111/"
client = FakeDetailClient(
    {"DVD": [{"title": "Resolved (2008)", "url": resolved_url}]}
)
result = enrich_row(
    client,
    {
        "row": "2",
        "upc": "097368122840",
        "url": "https://www.blu-ray.com/movies/Wrong-Blu-ray/999/",
    },
)
assert result["status"] == "OK"
assert result["url"] == resolved_url
assert result["source"] == "Blu-ray.com DVD"

client = FakeDetailClient({"DVD": [], "Blu-ray": []})
fallback_url = "https://www.blu-ray.com/movies/Fallback-Blu-ray/222/"
result = enrich_row(
    client,
    {"row": "3", "upc": "097368122840", "url": fallback_url},
)
assert result["status"] == "OK"
assert result["url"] == fallback_url
assert result["source"] == "No UPC/EAN, URL OK"

client = FakeDetailClient(
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
    }
)
ambiguous_fallback_url = (
    "https://www.blu-ray.com/dvd/When-Harry-Met-Sally-DVD/123/"
)
result = enrich_row(
    client,
    {
        "row": "4",
        "upc": "027616857804",
        "url": ambiguous_fallback_url,
    },
)
assert result["status"] == "OK"
assert result["url"] == ambiguous_fallback_url
assert result["source"] == "No UPC/EAN, URL OK"

client = FakeDetailClient(
    {"DVD": [{"title": "Resolved (2008)", "url": resolved_url}]}
)
cancelled_rows = process_rows(
    [
        {"row": "5", "upc": "097368122840", "url": ""},
        {"row": "6", "upc": "012345678905", "url": ""},
    ],
    45.0,
    0,
    CancelAfterOne(),
    client_factory=lambda timeout: client,
)
assert [row["status"] for row in cancelled_rows] == ["OK", "CANCELLED"]

print("PASS: Blu-ray.com release-detail parser tests")
