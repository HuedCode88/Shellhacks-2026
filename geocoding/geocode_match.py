"""
Gridlock Challenge -- Part 2: match project names to real-world coordinates.

WHAT THIS DOES (three stages, run in order):

  1. LOAD      Reads the two project tables (DESC + GA ITS) from the Excel
               workbook produced in Part 1, and the two Overpass GeoJSON
               exports produced by geocoding.search_overpass (this module
               does NOT hit any network API itself).

  2. MATCH     Splits each project's raw title into candidate endpoint
               names (e.g. "Queensboro - Ft Johnson 115 kV" -> "Queensboro",
               "Ft Johnson"), then fuzzy-matches each candidate against the
               `name` tag of every feature in the relevant GeoJSON file.
               Two matched endpoints -> center point is their midpoint.
               One matched endpoint -> that point is the center.
               No confident match -> flagged NEEDS_NOMINATIM.

  3. FALLBACK  (optional, needs network) For anything flagged
               NEEDS_NOMINATIM, calls the public Nominatim API one request
               per second (per Nominatim's usage policy) with a proper
               User-Agent, and fills in whatever it finds.

OUTPUT: a workbook with two sheets -- "DESC Geocoded" and "GA ITS Geocoded" --
        each containing the original project rows plus: matched endpoint
        names, lat/lon per endpoint, a center lat/lon, a match confidence,
        and match method. (Earlier versions of this pipeline also wrote
        copies of the raw, ungeocoded sheets alongside these; that's been
        dropped since nothing downstream ever reads them.)

Importable API
---------------
  run(excel, desc_geojson, ga_geojson, out, nominatim_fallback=False,
      contact_email="") -> str
      Runs the whole load+match(+fallback) stage and writes `out`.
      This is what geocoding.search_overpass calls after it fetches the
      GeoJSON files, and what this module's own CLI (`main()`) calls too.

CLI
---
  python3 -m geocoding.geocode_match \
      --excel gridlock_project_tables_raw.xlsx \
      --desc-geojson dominion_energy_sc_all.geojson \
      --ga-geojson georgia_power_all.geojson \
      --out gridlock_project_tables_geocoded.xlsx

  # Add --nominatim-fallback to also call Nominatim for unmatched rows
  # (needs network; respects 1 req/sec; set --contact-email to something
  # real, Nominatim's usage policy requires an identifiable User-Agent):
  python3 -m geocoding.geocode_match ... --nominatim-fallback --contact-email you@example.com

DEPENDENCIES
------------
  pandas, openpyxl.
  rapidfuzz is used if installed (better matching for reordered/abbreviated
  names) but is optional -- falls back to Python's built-in difflib.
  requests is only needed if nominatim_fallback=True.
"""

import argparse
import json
import re
import sys
import time
from dataclasses import dataclass, field
from difflib import SequenceMatcher

import pandas as pd

try:
    from rapidfuzz import fuzz as _rf_fuzz
    HAVE_RAPIDFUZZ = True
except ImportError:
    HAVE_RAPIDFUZZ = False


# --------------------------------------------------------------------------
# Config
# --------------------------------------------------------------------------

# Confidence thresholds on a 0-100 similarity scale.
HIGH_CONFIDENCE = 85
LOW_CONFIDENCE = 75  # below this -> NEEDS_NOMINATIM
MIN_MATCH_MARGIN = 8  # reject a match when the runner-up is too close

# Words stripped out when extracting "endpoint" candidates from a raw
# project title, since they describe the work/voltage, not a place.
NOISE_WORDS = {
    "kv", "kv:", "sub", "sub:", "substation", "tap", "line", "lines",
    "rebuild", "rebld", "reconductor", "construct", "constructing",
    "installing", "upgrade", "replace", "replacing", "fold-in", "fold",
    "in", "new", "spdc", "acsr", "structures", "conductor", "and",
    "with", "the", "of", "section", "sect", "approx", "miles", "mile",
    "bus", "bank", "autobank", "transformer", "relay", "panel",
    "modernization", "improvements", "phase", "gtc:", "meag:", "sav:",
    "du:",
}

VOLTAGE_RE = re.compile(
    r"\b\d{2,3}(?:\.\d+)?\s*(?:[-/]\s*\d{2,3}(?:\.\d+)?\s*)?k?v\b",
    re.IGNORECASE,
)
NUMBER_RE = re.compile(r"\b\d+(?:\.\d+)?\b")
SPONSOR_PREFIX_RE = re.compile(r"^(?:sav|gtc|meag|du|spdc)\s*:\s*", re.IGNORECASE)
PAREN_RE = re.compile(r"\([^)]*\)")


# --------------------------------------------------------------------------
# Endpoint extraction
# --------------------------------------------------------------------------

def extract_endpoints(raw_title: str) -> list[str]:
    """Turn a raw project title into a short list of candidate place names.

    e.g. "Queensboro - Ft Johnson 115 kV & Queensboro-Bayfront 115kV
          (Queensboro-James Island Sect)"
      -> ["Queensboro", "Ft Johnson", "Queensboro", "Bayfront"]
    """
    title = PAREN_RE.sub(" ", raw_title)          # drop parenthetical notes
    title = SPONSOR_PREFIX_RE.sub("", title)
    title = VOLTAGE_RE.sub(" ", title)             # drop "115 kV" etc.
    title = title.replace("&", " - ")
    # split on dash variants (-, –, —) and colons
    parts = re.split(r"[-–—:]", title)

    endpoints = []
    for part in parts:
        words = [w.strip(",.#") for w in part.split()]
        words = [w for w in words if not NUMBER_RE.fullmatch(w)]
        words = [w for w in words if w.lower() not in NOISE_WORDS and w]
        cleaned = " ".join(words).strip()
        if len(cleaned) >= 3:
            endpoints.append(cleaned)

    # de-dup while preserving order
    seen = set()
    unique = []
    for e in endpoints:
        key = e.lower()
        if key not in seen:
            seen.add(key)
            unique.append(e)
    return unique[:4]  # a project name rarely has more than 2-3 real endpoints


# --------------------------------------------------------------------------
# Fuzzy matching
# --------------------------------------------------------------------------

def similarity(a: str, b: str) -> float:
    """0-100 similarity score. Uses rapidfuzz's token_sort_ratio if
    available (handles word-order and partial differences better),
    otherwise falls back to difflib's SequenceMatcher ratio."""
    a, b = a.lower().strip(), b.lower().strip()
    if not a or not b:
        return 0.0
    if HAVE_RAPIDFUZZ:
        return max(
            _rf_fuzz.token_sort_ratio(a, b),
            _rf_fuzz.partial_ratio(a, b),
        )
    ratio = SequenceMatcher(None, a, b).ratio() * 100
    # crude partial-match boost: if one string is fully contained in the
    # other (common for "Okatie" inside "Okatie 230-115kV Substation")
    if a in b or b in a:
        ratio = max(ratio, 80.0)
    return ratio


@dataclass
class OSMFeature:
    name: str
    lat: float
    lon: float
    tags: dict = field(default_factory=dict)


def _coordinate_pairs(value):
    """Yield longitude/latitude pairs from any GeoJSON coordinate nesting."""
    if (
        isinstance(value, (list, tuple))
        and len(value) >= 2
        and all(isinstance(item, (int, float)) for item in value[:2])
    ):
        yield float(value[0]), float(value[1])
        return
    if isinstance(value, (list, tuple)):
        for child in value:
            yield from _coordinate_pairs(child)


def load_geojson_features(path: str) -> list[OSMFeature]:
    with open(path) as f:
        gj = json.load(f)
    features = []
    for feat in gj.get("features", []):
        props = feat.get("properties", {}) or {}
        name = props.get("name")
        if not name:
            continue
        geometry = feat.get("geometry") or {}
        coordinates = list(_coordinate_pairs(geometry.get("coordinates")))
        for child_geometry in geometry.get("geometries", []):
            coordinates.extend(_coordinate_pairs(child_geometry.get("coordinates")))
        if not coordinates:
            continue
        lon = sum(pair[0] for pair in coordinates) / len(coordinates)
        lat = sum(pair[1] for pair in coordinates) / len(coordinates)
        features.append(OSMFeature(name=name, lat=lat, lon=lon, tags=props))
    return features


def best_match(endpoint: str, features: list[OSMFeature]):
    """Return the best unambiguous match and its score."""
    ranked = sorted(
        ((similarity(endpoint, feat.name), feat) for feat in features),
        key=lambda item: item[0],
        reverse=True,
    )
    if not ranked:
        return None, 0.0
    best_score, best_feat = ranked[0]
    second_score = ranked[1][0] if len(ranked) > 1 else 0.0
    if best_score < LOW_CONFIDENCE or best_score - second_score < MIN_MATCH_MARGIN:
        return None, best_score
    return best_feat, best_score


# --------------------------------------------------------------------------
# Nominatim fallback (network required, rate-limited to Nominatim's policy)
# --------------------------------------------------------------------------

def geocode_via_nominatim(query: str, contact_email: str, state_hint: str):
    """One request to the public Nominatim API. Sleeps 1s AFTER returning
    so callers that loop over many queries stay within the 1 req/sec usage
    policy (https://operations.osmfoundation.org/policies/nominatim/).
    Requires `requests` and network access -- will raise ImportError/
    ConnectionError if unavailable, which callers should catch."""
    import requests  # imported lazily so the rest of the module has no hard dep

    headers = {"User-Agent": f"GridlockChallenge-Geocoder/1.0 ({contact_email})"}
    params = {
        "q": f"{query}, {state_hint}, USA",
        "format": "json",
        "limit": 1,
    }
    resp = requests.get(
        "https://nominatim.openstreetmap.org/search",
        params=params,
        headers=headers,
        timeout=15,
    )
    resp.raise_for_status()
    results = resp.json()
    time.sleep(1.0)  # respect 1 req/sec regardless of how fast the call was
    if not results:
        return None
    r = results[0]
    return {"lat": float(r["lat"]), "lon": float(r["lon"]), "display_name": r["display_name"]}


# --------------------------------------------------------------------------
# Per-project geocoding
# --------------------------------------------------------------------------

def geocode_project(raw_title: str, features: list[OSMFeature]):
    """Returns a dict describing the match outcome for one project."""
    endpoints = extract_endpoints(raw_title)
    matches = []
    for ep in endpoints:
        feat, score = best_match(ep, features)
        if feat and score >= LOW_CONFIDENCE:
            matches.append((ep, feat, score))

    # de-dup matches that landed on the exact same OSM feature
    dedup = []
    seen_names = set()
    for ep, feat, score in matches:
        if feat.name not in seen_names:
            seen_names.add(feat.name)
            dedup.append((ep, feat, score))
    matches = dedup[:2]  # a project's center point uses at most 2 endpoints

    if not matches:
        return {
            "endpoints_tried": endpoints,
            "matched_endpoint_1": None, "osm_name_1": None,
            "lat_1": None, "lon_1": None, "score_1": None,
            "matched_endpoint_2": None, "osm_name_2": None,
            "lat_2": None, "lon_2": None, "score_2": None,
            "center_lat": None, "center_lon": None,
            "confidence": "UNMATCHED", "match_method": "NEEDS_NOMINATIM",
        }

    row = {
        "endpoints_tried": endpoints,
        "matched_endpoint_1": matches[0][0], "osm_name_1": matches[0][1].name,
        "lat_1": matches[0][1].lat, "lon_1": matches[0][1].lon, "score_1": round(matches[0][2], 1),
        "matched_endpoint_2": None, "osm_name_2": None,
        "lat_2": None, "lon_2": None, "score_2": None,
    }

    if len(matches) == 2:
        row.update({
            "matched_endpoint_2": matches[1][0], "osm_name_2": matches[1][1].name,
            "lat_2": matches[1][1].lat, "lon_2": matches[1][1].lon, "score_2": round(matches[1][2], 1),
        })
        row["center_lat"] = (matches[0][1].lat + matches[1][1].lat) / 2
        row["center_lon"] = (matches[0][1].lon + matches[1][1].lon) / 2
    else:
        row["center_lat"] = matches[0][1].lat
        row["center_lon"] = matches[0][1].lon

    min_score = min(m[2] for m in matches)
    row["confidence"] = "HIGH" if min_score >= HIGH_CONFIDENCE else "MEDIUM"
    row["match_method"] = "OVERPASS"
    return row


# --------------------------------------------------------------------------
# Sheet-level and workbook-level drivers
# --------------------------------------------------------------------------

def process_sheet(df: pd.DataFrame, title_col: str, features: list[OSMFeature],
                   nominatim_fallback: bool, contact_email: str, state_hint: str) -> pd.DataFrame:
    results = []
    for _, row in df.iterrows():
        raw_title = str(row[title_col])
        result = geocode_project(raw_title, features)

        if result["confidence"] == "UNMATCHED" and nominatim_fallback:
            queries = result["endpoints_tried"] or [raw_title]
            nominatim_hits = []
            for query in queries[:2]:
                try:
                    hit = geocode_via_nominatim(query, contact_email, state_hint)
                except Exception as e:  # network unavailable, rate-limited, etc.
                    print(f"  [nominatim] failed for '{query[:60]}': {e}", file=sys.stderr)
                    hit = None
                if hit:
                    nominatim_hits.append((query, hit))

            if nominatim_hits:
                result.update({
                    "matched_endpoint_1": nominatim_hits[0][0],
                    "osm_name_1": nominatim_hits[0][1]["display_name"],
                    "lat_1": nominatim_hits[0][1]["lat"],
                    "lon_1": nominatim_hits[0][1]["lon"],
                    "center_lat": nominatim_hits[0][1]["lat"],
                    "center_lon": nominatim_hits[0][1]["lon"],
                    "confidence": "LOW", "match_method": "NOMINATIM",
                })
                if len(nominatim_hits) == 2:
                    result.update({
                        "matched_endpoint_2": nominatim_hits[1][0],
                        "osm_name_2": nominatim_hits[1][1]["display_name"],
                        "lat_2": nominatim_hits[1][1]["lat"],
                        "lon_2": nominatim_hits[1][1]["lon"],
                        "center_lat": (nominatim_hits[0][1]["lat"] + nominatim_hits[1][1]["lat"]) / 2,
                        "center_lon": (nominatim_hits[0][1]["lon"] + nominatim_hits[1][1]["lon"]) / 2,
                    })

        results.append(result)

    result_df = pd.DataFrame(results)
    return pd.concat([df.reset_index(drop=True), result_df.reset_index(drop=True)], axis=1)


def run(excel: str, desc_geojson: str, ga_geojson: str, out: str,
        nominatim_fallback: bool = False, contact_email: str = "") -> str:
    """Runs the load+match(+fallback) stage end-to-end and writes `out`.

    This is the single entry point other modules should call -- both
    geocoding.search_overpass (after it fetches the GeoJSON files) and
    this module's own CLI use it, so the logic only lives in one place.
    """
    if nominatim_fallback and not contact_email:
        raise ValueError(
            "nominatim_fallback=True requires contact_email "
            "(Nominatim's usage policy requires an identifiable User-Agent)"
        )

    print(f"Fuzzy matcher: {'rapidfuzz' if HAVE_RAPIDFUZZ else 'difflib (rapidfuzz not installed -- pip install rapidfuzz for better matches)'}")

    desc_features = load_geojson_features(desc_geojson)
    ga_features = load_geojson_features(ga_geojson)
    print(f"Loaded {len(desc_features)} DESC OSM features, {len(ga_features)} GA/ITS OSM features")

    desc_df = pd.read_excel(excel, sheet_name="DESC Projects")
    ga_df = pd.read_excel(excel, sheet_name="GA ITS Projects")

    print("Matching DESC projects...")
    desc_out = process_sheet(
        desc_df, "Project Name / Endpoints (raw title)", desc_features,
        nominatim_fallback, contact_email, "South Carolina",
    )
    print("Matching GA ITS projects...")
    ga_out = process_sheet(
        ga_df, "Project Name / Endpoints (raw title)", ga_features,
        nominatim_fallback, contact_email, "Georgia",
    )

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        # Only the geocoded sheets are ever read downstream (database.py
        # only picks up rows with lat/lon, which the raw sheets never
        # have), so that's all this writes -- no raw-table duplicates.
        desc_out.to_excel(writer, sheet_name="DESC Geocoded", index=False)
        ga_out.to_excel(writer, sheet_name="GA ITS Geocoded", index=False)

    for name, sheet_out in [("DESC", desc_out), ("GA ITS", ga_out)]:
        counts = sheet_out["confidence"].value_counts().to_dict()
        print(f"{name}: {counts}")

    print(f"\nSaved: {out}")
    return out


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--excel", required=True, help="Part 1 output workbook")
    ap.add_argument("--desc-geojson", required=True, help="Overpass export for Dominion Energy SC")
    ap.add_argument("--ga-geojson", required=True, help="Overpass export for Georgia Power/ITS")
    ap.add_argument("--out", default="gridlock_project_tables_geocoded.xlsx")
    ap.add_argument("--nominatim-fallback", action="store_true",
                     help="Call the public Nominatim API for anything Overpass couldn't match (needs network)")
    ap.add_argument("--contact-email", default="", help="Required by Nominatim's usage policy if using --nominatim-fallback")
    args = ap.parse_args()

    if args.nominatim_fallback and not args.contact_email:
        ap.error("--nominatim-fallback requires --contact-email (Nominatim's usage policy requires an identifiable User-Agent)")

    run(
        excel=args.excel,
        desc_geojson=args.desc_geojson,
        ga_geojson=args.ga_geojson,
        out=args.out,
        nominatim_fallback=args.nominatim_fallback,
        contact_email=args.contact_email,
    )


if __name__ == "__main__":
    main()