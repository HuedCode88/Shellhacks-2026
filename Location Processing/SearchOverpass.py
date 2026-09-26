import argparse
import json
from pathlib import Path

import requests

HEADERS = {
    "User-Agent": "GridLockChallenge-DataResearch/1.0 (xxxx@yyyy.com)"
}

def search_overpass_by_operator(operator_pattern, bbox, timeout=180):
    """
    Search OpenStreetMap power infrastructure by OPERATOR (not name) across
    an entire region. Returns every tagged node/way/relation whose operator
    field matches the pattern -- useful for pulling ALL of a utility's
    infrastructure at once, rather than searching name-by-name.
    """
    endpoints = [
        "https://overpass-api.de/api/interpreter",
        "https://overpass.kumi.systems/api/interpreter",
    ]

    query = (
        f'[out:json][timeout:{timeout}][bbox:{bbox}];'
        f'nwr["power"]["operator"~"{operator_pattern}",i];'
        f'out center;'
    )

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
            print(f"  Success.")
            return response.json()
        except Exception as e:
            print(f"  Failed: {e}")
            continue

    raise RuntimeError("All Overpass endpoints failed or timed out.")


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


if __name__ == "__main__":
    script_dir = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(description="Fetch utility infrastructure and geocode the project workbook.")
    parser.add_argument("--excel", default=str(script_dir / "gridlock_project_tables_raw.xlsx"))
    parser.add_argument("--out", default=str(script_dir / "gridlock_project_tables_geocoded.xlsx"))
    parser.add_argument("--desc-geojson", default=str(script_dir / "dominion_energy_sc_all.geojson"))
    parser.add_argument("--ga-geojson", default=str(script_dir / "georgia_power_all.geojson"))
    parser.add_argument("--contact-email", default="", help="Contact included in the Overpass User-Agent")
    parser.add_argument("--skip-query", action="store_true", help="Use existing GeoJSON files instead of calling Overpass")
    args = parser.parse_args()

    if args.contact_email:
        HEADERS["User-Agent"] = f"GridLockChallenge-DataResearch/1.0 ({args.contact_email})"

    if not args.skip_query:
        desc_result = search_overpass_by_operator(
            "Dominion Energy|South Carolina Electric",
            bbox="32.0,-83.5,35.3,-78.5",
        )
        summarize_results(desc_result)
        save_as_geojson(desc_result, args.desc_geojson)

        ga_result = search_overpass_by_operator(
            "Georgia Power|Georgia Transmission|Savannah Electric",
            bbox="30.3,-85.7,35.1,-80.7",
        )
        summarize_results(ga_result)
        save_as_geojson(ga_result, args.ga_geojson)

    from GeocodeMatch import main as match_workbook

    import sys
    sys.argv = [
        "GeocodeMatch.py",
        "--excel", str(Path(args.excel)),
        "--desc-geojson", str(Path(args.desc_geojson)),
        "--ga-geojson", str(Path(args.ga_geojson)),
        "--out", str(Path(args.out)),
    ]
    match_workbook()