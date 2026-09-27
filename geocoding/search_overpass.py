"""
Stage 1 of the pipeline: pull utility power infrastructure out of
OpenStreetMap by OPERATOR (not name) across an entire region, save it as
GeoJSON, then hand off to geocoding.geocode_match to match it against the
project workbook from Part 1.

Run standalone:
  python3 -m geocoding.search_overpass --excel raw.xlsx --out geocoded.xlsx

Or import run_pipeline() / search_overpass_by_operator() directly.
"""

import argparse
import json
from pathlib import Path

import requests

from . import geocode_match

HEADERS = {
    "User-Agent": "GridLockChallenge-DataResearch/1.0 (xxxx@yyyy.com)"
}


def _run_overpass_query(query, timeout=180):
    """Posts a raw Overpass QL query to the primary endpoint, then the
    mirror, returning parsed JSON. Shared by the operator-filtered query
    below and the name-only query used for live, any-utility lookups."""
    endpoints = [
        "https://overpass-api.de/api/interpreter",
        "https://overpass.kumi.systems/api/interpreter",
    ]

    for endpoint in endpoints:
        try:
            print(f"Trying {endpoint} ...")
            response = requests.post(
                endpoint,
                data={"data": query},
                headers=HEADERS,
                timeout=timeout + 15,
            )
            response.raise_for_status()
            print("  Success.")
            return response.json()
        except Exception as e:
            print(f"  Failed: {e}")
            continue

    raise RuntimeError("All Overpass endpoints failed or timed out.")


def search_overpass_by_operator(operator_pattern, bbox, timeout=180):
    """
    Search OpenStreetMap power infrastructure by OPERATOR (not name) across
    an entire region. Returns every tagged node/way/relation whose operator
    field matches the pattern -- useful for pulling ALL of a utility's
    infrastructure at once, rather than searching name-by-name.
    """
    query = (
        f'[out:json][timeout:{timeout}][bbox:{bbox}];'
        f'nwr["power"]["operator"~"{operator_pattern}",i];'
        f'out center;'
    )
    return _run_overpass_query(query, timeout)


def search_all_named_power(bbox, timeout=180):
    """
    Search OpenStreetMap for named power-infrastructure feature in a
    region, regardless of operator. Used for live geocoding of uploaded
    documents where the utility isn't known ahead of time -- unlike
    search_overpass_by_operator, this doesn't require knowing which
    company's infrastructure to look for.

    Restricted to feature types that plausibly carry a project-relevant
    name (substations, plants, generators, transformers, switches) rather
    than every power=* tag -- an unrestricted scan of a whole state
    includes tens of thousands of poles/lines and reliably times out on
    public Overpass instances.
    """
    query = (
        f'[out:json][timeout:{timeout}][bbox:{bbox}];'
        f'nwr["power"~"^(substation|plant|generator|transformer|switch|converter|compensator)$"]["name"];'
        f'out center;'
    )
    return _run_overpass_query(query, timeout)


def _split_bbox(bbox):
    """Splits a "south,west,north,east" bbox in half along its longer
    dimension, returning the two halves."""
    south, west, north, east = (float(x) for x in bbox.split(","))
    if (north - south) >= (east - west):
        mid = (north + south) / 2
        return (f"{south},{west},{mid},{east}", f"{mid},{west},{north},{east}")
    mid = (east + west) / 2
    return (f"{south},{west},{north},{mid}", f"{south},{mid},{north},{east}")


def fetch_named_power_with_retry(bbox, timeout=180, max_split_depth=2):
    """Runs search_all_named_power(bbox), and if that region times out,
    automatically halves it and retries each half independently (up to
    max_split_depth levels deep) rather than giving up on the whole
    region. Returns the combined elements from whichever sub-regions
    ultimately succeeded; a sub-region that still fails at max depth is
    skipped (logged, not raised) so one stubborn corner doesn't block
    everything else."""
    try:
        return search_all_named_power(bbox, timeout=timeout)
    except RuntimeError as error:
        if max_split_depth <= 0:
            print(f"  [overpass] giving up on {bbox} (still failing after splitting): {error}")
            return {"elements": []}

        print(f"  [overpass] {bbox} timed out -- splitting in half and retrying each side...")
        half_a, half_b = _split_bbox(bbox)
        data_a = fetch_named_power_with_retry(half_a, timeout=timeout, max_split_depth=max_split_depth - 1)
        data_b = fetch_named_power_with_retry(half_b, timeout=timeout, max_split_depth=max_split_depth - 1)
        return {"elements": data_a.get("elements", []) + data_b.get("elements", [])}


def get_coords(el):
    if "lat" in el and "lon" in el:
        return el["lat"], el["lon"]
    elif "center" in el:
        return el["center"]["lat"], el["center"]["lon"]
    return None, None


def summarize_results(data):
    elements = data.get("elements", [])
    named = [el for el in elements if el.get("tags", {}).get("name")]

    print(f"\nTotal features returned: {len(elements)}")
    print(f"Named features: {len(named)}\n")

    if named:
        print("=== Named features ===")
        for el in named:
            tags = el["tags"]
            lat, lon = get_coords(el)
            print(f"{tags.get('name')} ({tags.get('power')}) "
                  f"[operator={tags.get('operator')}] -> {lat}, {lon}")


def save_as_geojson(data, filepath):
    """Convert Overpass JSON to a GeoJSON FeatureCollection and save it."""
    features = []
    for el in data.get("elements", []):
        lat, lon = get_coords(el)
        if lat is None:
            continue
        features.append({
            "type": "Feature",
            "properties": el.get("tags", {}),
            "geometry": {"type": "Point", "coordinates": [lon, lat]},
        })

    geojson = {"type": "FeatureCollection", "features": features}
    with open(filepath, "w") as f:
        json.dump(geojson, f, indent=2)
    print(f"\nSaved {len(features)} features to {filepath}")


def fetch_all(desc_geojson_path, ga_geojson_path):
    """Runs both Overpass queries (Dominion Energy SC, Georgia Power/ITS)
    and writes each result to its own GeoJSON path."""
    desc_result = search_overpass_by_operator(
        "Dominion Energy|South Carolina Electric",
        bbox="32.0,-83.5,35.3,-78.5",
    )
    summarize_results(desc_result)
    save_as_geojson(desc_result, desc_geojson_path)

    ga_result = search_overpass_by_operator(
        "Georgia Power|Georgia Transmission|Savannah Electric",
        bbox="30.3,-85.7,35.1,-80.7",
    )
    summarize_results(ga_result)
    save_as_geojson(ga_result, ga_geojson_path)


def run_pipeline(excel, out, desc_geojson, ga_geojson, contact_email="",
                  skip_query=False, nominatim_fallback=False):
    """Fetch stage + match stage in one call. Returns the output path."""
    if contact_email:
        HEADERS["User-Agent"] = f"GridLockChallenge-DataResearch/1.0 ({contact_email})"

    if not skip_query:
        fetch_all(desc_geojson, ga_geojson)

    return geocode_match.run(
        excel=excel,
        desc_geojson=desc_geojson,
        ga_geojson=ga_geojson,
        out=out,
        nominatim_fallback=nominatim_fallback,
        contact_email=contact_email,
    )


def main():
    script_dir = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(description="Fetch utility infrastructure and geocode the project workbook.")
    parser.add_argument("--excel", default=str(script_dir / "gridlock_project_tables_raw.xlsx"))
    parser.add_argument("--out", default=str(script_dir / "gridlock_project_tables_geocoded.xlsx"))
    parser.add_argument("--desc-geojson", default=str(script_dir / "dominion_energy_sc_all.geojson"))
    parser.add_argument("--ga-geojson", default=str(script_dir / "georgia_power_all.geojson"))
    parser.add_argument("--contact-email", default="", help="Contact included in the Overpass User-Agent")
    parser.add_argument("--skip-query", action="store_true", help="Use existing GeoJSON files instead of calling Overpass")
    parser.add_argument("--nominatim-fallback", action="store_true",
                         help="Also call Nominatim for anything Overpass couldn't match (needs network)")
    args = parser.parse_args()

    run_pipeline(
        excel=args.excel,
        out=args.out,
        desc_geojson=args.desc_geojson,
        ga_geojson=args.ga_geojson,
        contact_email=args.contact_email,
        skip_query=args.skip_query,
        nominatim_fallback=args.nominatim_fallback,
    )


if __name__ == "__main__":
    main()