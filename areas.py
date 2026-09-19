"""
areas.py — free, offline Area 1-4 (State/District-City/Locality/Street)
auto-fill for the export. Self-contained: doesn't touch or depend on
the scraping/scrolling code, and nothing else depends on its internals
except detect_areas() itself, which excel_export.py calls.
"""

import os
import re

# ============================================================
# Area auto-fill (free, offline — postcode lookup + keyword matching)
# ============================================================
#
# This is a SEPARATE, self-contained block from the scraping/scrolling
# code above and below it — it doesn't touch, call, or depend on any
# of that. It only produces suggested values for the Area 1-4 columns,
# which write_output() then applies with one rule: NEVER overwrite a
# cell that already has something in it (typed by hand OR auto-filled
# on an earlier run). Only a cell that's still genuinely blank gets a
# new auto-filled value. Tested against a real export and a batch of
# deliberately messy addresses before being built in here. Known
# limitation, accepted on purpose: this will leave plenty of cells
# blank for you to fill in by hand, especially Area 3/Area 4 on
# addresses with no commas — that's the safe trade-off on purpose,
# not a bug.
#
# POSTCODE_LOOKUP_FILE is a free, offline dataset (Malaysia postcodes
# -> state/city, from a public GitHub repo, saved locally as
# postcodes_my.json next to this script) — no internet call at
# runtime, no API key, no cost. If that file is ever missing, Area 1/
# Area 2 just won't auto-fill (everything else still works normally).

POSTCODE_LOOKUP_FILE = "postcodes_my.json"

# Locality-type words that can appear BEFORE the name they describe
# (the usual Malay convention, e.g. "Taman Wangsa Melawati").
LOCALITY_KEYWORDS = [
    "Taman", "Kampung", "Kg", "Bandar", "Presint", "Precint",
    "Seksyen", "Desa", "Pangsapuri", "Kondominium", "Residensi",
]
# Road-type words — same "keyword before the name" pattern.
STREET_KEYWORDS = ["Jalan", "Jln", "Lorong", "Persiaran", "Lebuh", "Lebuhraya"]

# A match this long is almost certainly two different things run
# together (e.g. a street name merged with the next town's name because
# there was no comma to stop at) — confirmed by testing. Past this
# length, the match is treated as unreliable and dropped rather than
# risking a wrong value sitting in the sheet looking legit.
_MAX_CONFIDENT_MATCH_CHARS = 45


def _load_postcode_lookup(path=POSTCODE_LOOKUP_FILE):
    """Loads the free Malaysia postcode -> (state, city) dataset. Returns
    {} (and prints a one-time warning) if the file isn't there — Area 1/
    Area 2 simply won't auto-fill in that case, nothing else breaks."""
    if not os.path.exists(path):
        print(f"  !! {path} not found next to the script — Area 1 (State) "
              f"and Area 2 (District / City) won't auto-fill this run. "
              f"Everything else still works normally.")
        return {}
    try:
        import json
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        lookup = {}
        for state in data.get("state", []):
            for city in state.get("city", []):
                for pc in city.get("postcode", []):
                    lookup[pc] = (state["name"], city["name"])
        return lookup
    except Exception as e:
        print(f"  !! Could not read {path} ({e}) — Area 1/Area 2 won't "
              f"auto-fill this run.")
        return {}


_POSTCODE_LOOKUP = _load_postcode_lookup()

_LOCALITY_PATTERN = re.compile(
    r"\b(" + "|".join(re.escape(k) for k in LOCALITY_KEYWORDS) + r")\.?\s+"
    r"[A-Za-z0-9'\-\s]*?(?=,|\d{5}|$)",
    re.IGNORECASE,
)
_STREET_PATTERN = re.compile(
    r"\b(" + "|".join(re.escape(k) for k in STREET_KEYWORDS) + r")\.?\s+"
    r"[A-Za-z0-9'\-/\s]+?(?=,|\d{5}|$)",
    re.IGNORECASE,
)


def detect_areas(address):
    """
    Takes one raw address string and returns whatever it can confidently
    work out, as {"area1": ..., "area2": ..., "area3": ..., "area4": ...}.
    Any level it's not confident about is left as "" — by design, not
    left for this function to guess at.

      area1 = State       )  both from the postcode found in the address,
      area2 = District/City) looked up in the free postcode dataset
      area3 = Locality    -  from a keyword match (Taman/Kampung/Presint/...)
      area4 = Street      -  from a keyword match (Jalan/Lorong/...)

    Area 3/Area 4 are only kept if the match is short enough to be
    trustworthy (see _MAX_CONFIDENT_MATCH_CHARS) — a long match usually
    means it accidentally swallowed extra, unrelated text because the
    address had no comma to stop it cleanly.
    """
    result = {"area1": "", "area2": "", "area3": "", "area4": ""}
    if not address:
        return result

    postcode_match = re.search(r"\b(\d{5})\b", address)
    if postcode_match:
        postcode = postcode_match.group(1)
        if postcode in _POSTCODE_LOOKUP:
            state, city = _POSTCODE_LOOKUP[postcode]
            result["area1"] = state
            result["area2"] = city

    locality_match = _LOCALITY_PATTERN.search(address)
    if locality_match:
        text = locality_match.group(0).strip().rstrip(",")
        if len(text) <= _MAX_CONFIDENT_MATCH_CHARS:
            result["area3"] = text

    street_match = _STREET_PATTERN.search(address)
    if street_match:
        text = street_match.group(0).strip().rstrip(",")
        if len(text) <= _MAX_CONFIDENT_MATCH_CHARS:
            result["area4"] = text

    return result

