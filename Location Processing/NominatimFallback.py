import argparse
import ast
from pathlib import Path

import pandas as pd

from GeocodeMatch import extract_endpoints, geocode_via_nominatim


TARGET_SHEETS = {
    "DESC Geocoded": "South Carolina",
    "GA ITS Geocoded": "Georgia",
}


def parse_endpoints(value, raw_title):
    if isinstance(value, list):
        return [str(item) for item in value if str(item).strip()]
    if isinstance(value, str) and value.strip():
        try:
            parsed = ast.literal_eval(value)
            if isinstance(parsed, list):
                return [str(item) for item in parsed if str(item).strip()]
        except (SyntaxError, ValueError):
            pass
    return extract_endpoints(str(raw_title))


def update_sheet(df, state_hint, contact_email):
    updated = df.copy()
    fallback_rows = updated[
        updated["confidence"].eq("UNMATCHED")
        | updated["match_method"].eq("NEEDS_NOMINATIM")
    ]
    print(f"{state_hint}: {len(fallback_rows)} rows need Nominatim")

    for index in fallback_rows.index:
        raw_title = updated.at[index, "Project Name / Endpoints (raw title)"]
        endpoints = parse_endpoints(updated.at[index, "endpoints_tried"], raw_title)
        hits = []

        for endpoint in endpoints[:2]:
            try:
                hit = geocode_via_nominatim(endpoint, contact_email, state_hint)
            except Exception as error:
                print(f"  Failed '{endpoint[:60]}': {error}")
                hit = None
            if hit:
                hits.append((endpoint, hit))

        if not hits:
            continue

        first_endpoint, first_hit = hits[0]
        updated.at[index, "matched_endpoint_1"] = first_endpoint
        updated.at[index, "osm_name_1"] = first_hit["display_name"]
        updated.at[index, "lat_1"] = first_hit["lat"]
        updated.at[index, "lon_1"] = first_hit["lon"]
        updated.at[index, "score_1"] = None
        updated.at[index, "center_lat"] = first_hit["lat"]
        updated.at[index, "center_lon"] = first_hit["lon"]
        updated.at[index, "confidence"] = "LOW"
        updated.at[index, "match_method"] = "NOMINATIM"

        if len(hits) == 2:
            second_endpoint, second_hit = hits[1]
            updated.at[index, "matched_endpoint_2"] = second_endpoint
            updated.at[index, "osm_name_2"] = second_hit["display_name"]
            updated.at[index, "lat_2"] = second_hit["lat"]
            updated.at[index, "lon_2"] = second_hit["lon"]
            updated.at[index, "score_2"] = None
            updated.at[index, "center_lat"] = (first_hit["lat"] + second_hit["lat"]) / 2
            updated.at[index, "center_lon"] = (first_hit["lon"] + second_hit["lon"]) / 2

    return updated


def main():
    script_dir = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(
        description="Populate unmatched rows in an existing geocoded workbook with Nominatim."
    )
    parser.add_argument(
        "--excel",
        default=str(script_dir / "gridlock_project_tables_geocoded.xlsx"),
        help="Existing geocoded workbook",
    )
    parser.add_argument(
        "--out",
        default=str(script_dir / "gridlock_project_tables_geocoded_nominatim.xlsx"),
        help="New workbook to create",
    )
    parser.add_argument("--contact-email", required=True)
    args = parser.parse_args()

    workbook = pd.ExcelFile(args.excel)
    sheets = {
        sheet_name: pd.read_excel(args.excel, sheet_name=sheet_name)
        for sheet_name in workbook.sheet_names
    }

    for sheet_name, state_hint in TARGET_SHEETS.items():
        if sheet_name in sheets:
            sheets[sheet_name] = update_sheet(sheets[sheet_name], state_hint, args.contact_email)

    with pd.ExcelWriter(args.out, engine="openpyxl") as writer:
        for sheet_name, sheet in sheets.items():
            sheet.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"Saved: {args.out}")


if __name__ == "__main__":
    main()
