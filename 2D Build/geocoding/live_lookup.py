"""
Live geocoding for uploaded documents.

Unlike the batch pipeline (geocode_match.run / run_pipeline.py), this
module doesn't need a pre-built raw workbook. It answers one question on
demand: "given this project's name, where is it?" -- by fuzzy-matching the
name against real OpenStreetMap power infrastructure (substations, plants,
transmission lines, towers) fetched for a configurable region, regardless
of which utility operates it.

Used by services.normalize_uploaded_projects() as the first geocoding
attempt for an uploaded project with no explicit coordinates. Falls back
to services.geocode_location() (Nominatim, address-text based) for
anything this doesn't recognize by name -- the two are complementary:
this module is good at "Queensboro - Ft Johnson 115 kV" (an infrastructure
name with no street address), Nominatim is good at "123 Main St, Anytown".

The fetched OSM data is cached to a GeoJSON file on disk so the (fairly
slow, multi-minute) Overpass query only runs once per cache lifetime, not
once per upload. Call warm_cache() at app startup to pay that cost before
any user is waiting on it; if you don't, the first geocode_by_name() call
after a cold start or an expired cache will pay it inline.
"""

import json
import time
from pathlib import Path
from threading import Lock

from . import geocode_match
from .search_overpass import fetch_named_power_with_retry, get_coords, HEADERS

CACHE_LOCK = Lock()
_CACHED_FEATURES = None


def _cache_path():
    return Path(__file__).resolve().parent / "overpass_feature_cache.geojson"


def _fetch_features(bboxes, contact_email="", timeout=180):
    """Runs one Overpass query per bbox and merges the named power
    features found into a single flat list."""
    if contact_email:
        HEADERS["User-Agent"] = f"GridLockChallenge-DataResearch/1.0 ({contact_email})"

    features = []
    for bbox in bboxes:
        data = fetch_named_power_with_retry(bbox, timeout=timeout)
        for el in data.get("elements", []):
            tags = el.get("tags", {}) or {}
            name = tags.get("name")
            if not name:
                continue
            lat, lon = get_coords(el)
            if lat is None:
                continue
            features.append(geocode_match.OSMFeature(name=name, lat=lat, lon=lon, tags=tags))

    return features


def _save_cache(features):
    geojson = {
        "type": "FeatureCollection",
        "features": [
            {
                "type": "Feature",
                "properties": feat.tags,
                "geometry": {"type": "Point", "coordinates": [feat.lon, feat.lat]},
            }
            for feat in features
        ],
    }
    with open(_cache_path(), "w") as f:
        json.dump(geojson, f)


def warm_cache(bboxes, contact_email="", max_age_days=7, force_refresh=False):
    """Ensures the on-disk feature cache exists and is fresh, fetching from
    Overpass if needed. Safe to call at app startup (recommended, so users
    never pay this cost on their first upload) or lazily on first use.
    Returns the loaded feature list."""
    global _CACHED_FEATURES

    with CACHE_LOCK:
        cache_file = _cache_path()
        is_stale = (
            force_refresh
            or not cache_file.exists()
            or (time.time() - cache_file.stat().st_mtime) > max_age_days * 86400
        )

        if is_stale:
            print("Fetching OSM power infrastructure for live geocoding cache "
                  "(this can take a minute or two per region)...")
            features = _fetch_features(bboxes, contact_email=contact_email)
            if features:
                _save_cache(features)
                print(f"Cached {len(features)} named power features for live geocoding.")
            else:
                print("  No features fetched -- leaving any existing cache in place.")
                if cache_file.exists():
                    features = geocode_match.load_geojson_features(str(cache_file))
            _CACHED_FEATURES = features
        elif _CACHED_FEATURES is None:
            _CACHED_FEATURES = geocode_match.load_geojson_features(str(cache_file))

        return _CACHED_FEATURES


def geocode_by_name(name, bboxes, contact_email=""):
    """Fuzzy-matches `name` against the cached OSM power infrastructure.

    Returns (lat, lon, matched_osm_name, confidence) on a confident match,
    or None if nothing in the cache resembles this project closely enough
    to trust. Loads/fetches the cache on first call if warm_cache() hasn't
    already been called.
    """
    features = _CACHED_FEATURES if _CACHED_FEATURES is not None else warm_cache(
        bboxes, contact_email=contact_email
    )
    if not features:
        return None

    endpoints = geocode_match.extract_endpoints(name)
    matches = []
    for endpoint in endpoints:
        feat, score = geocode_match.best_match(endpoint, features)
        if feat and score >= geocode_match.LOW_CONFIDENCE:
            matches.append((feat, score))

    if not matches:
        return None

    # de-dup matches that landed on the same OSM feature, same as the
    # batch pipeline's geocode_project()
    seen = set()
    dedup = []
    for feat, score in matches:
        if feat.name not in seen:
            seen.add(feat.name)
            dedup.append((feat, score))
    dedup = dedup[:2]

    if len(dedup) == 2:
        lat = (dedup[0][0].lat + dedup[1][0].lat) / 2
        lon = (dedup[0][0].lon + dedup[1][0].lon) / 2
        matched_name = f"{dedup[0][0].name} / {dedup[1][0].name}"
    else:
        lat, lon = dedup[0][0].lat, dedup[0][0].lon
        matched_name = dedup[0][0].name

    min_score = min(score for _, score in dedup)
    confidence = "HIGH" if min_score >= geocode_match.HIGH_CONFIDENCE else "MEDIUM"
    return lat, lon, matched_name, confidence