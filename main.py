"""
Maps every geocoded project in the workbook onto a real,
interactive OpenStreetMap-based map (via Folium/Leaflet) and saves
it as a standalone HTML file you can open in any browser.

Unlike the in-chat map tool, this has no marker limit, combines
both utilities on one map, and can be re-run any time the workbook
is updated -- just run:

    pip install openpyxl folium
    python map_projects.py

It looks for the workbook in the same folder as this script by
default; change EXCEL_FILE below if yours lives elsewhere.
"""

import os
import openpyxl
import folium
from folium.plugins import MarkerCluster

EXCEL_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "gridlock_project_tables_geocoded_nominatim (1).xlsx"
)

OUTPUT_HTML = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "project_map.html"
)

# Which column identifies a project's category on each sheet, and
# what color to draw that category's markers in. Folium's built-in
# Icon colors are a fixed palette (not arbitrary hex), so we map
# categories onto the closest ones.
SHEET_CATEGORY_FIELD = {
    "DESC Geocoded": "Status",
    "GA ITS Geocoded": "Sponsor (GPC/GTC/MEAG/DU/SAV)",
}

CATEGORY_COLORS = {
    ("DESC Geocoded", "In Progress"): "green",
    ("DESC Geocoded", "Planned"): "orange",

    ("GA ITS Geocoded", "GPC"): "blue",
    ("GA ITS Geocoded", "GTC"): "purple",
    ("GA ITS Geocoded", "SAV"): "pink",
    ("GA ITS Geocoded", "MEAG"): "cadetblue",
    ("GA ITS Geocoded", "DU"): "darkred",
}

DEFAULT_COLOR = "gray"

ID_FIELD_CANDIDATES = ["Project ID", "TEAMS Number"]
NAME_FIELD = "Project Name / Endpoints (raw title)"

# Extra fields to show in each marker's popup, per sheet, in order.
# Only fields that exist and have a value are shown.
POPUP_FIELDS = {
    "DESC Geocoded": ["Status", "Planned In-Service Date", "Description", "confidence"],
    "GA ITS Geocoded": ["Sponsor (GPC/GTC/MEAG/DU/SAV)", "Need / In-Service Date", "Zone", "confidence"],
}


def safe_float(value):
    """Coerces a coordinate cell to a float, tolerating stray commas/whitespace."""

    if value is None:
        return None

    if isinstance(value, (int, float)):
        return float(value)

    if isinstance(value, str):

        text = value.strip()

        try:
            return float(text)
        except ValueError:
            pass

        parts = [p.strip() for p in text.split(",") if p.strip()]

        for part in reversed(parts):
            try:
                return float(part)
            except ValueError:
                continue

    return None


def _row_dict(header, row):
    return {header[i]: row[i] for i in range(len(header))}


def load_projects(path):
    """
    Reads every sheet in the workbook and returns one dict per row
    that has a valid center_lat/center_lon -- rows the geocoder
    marked UNMATCHED are skipped since there's nowhere to plot them.
    """

    wb = openpyxl.load_workbook(path, data_only=True)

    projects = []

    for sheet_name in wb.sheetnames:

        ws = wb[sheet_name]

        rows = list(ws.iter_rows(values_only=True))

        if not rows:
            continue

        header = rows[0]

        category_field = SHEET_CATEGORY_FIELD.get(sheet_name)
        popup_fields = POPUP_FIELDS.get(sheet_name, [])

        for row in rows[1:]:

            record = _row_dict(header, row)

            raw_lat = record.get("center_lat")
            raw_lon = record.get("center_lon")

            lat = safe_float(raw_lat)
            lon = safe_float(raw_lon)

            if lat is None or lon is None:

                if raw_lat is not None or raw_lon is not None:
                    print(
                        f"Skipping row with unparseable coordinates "
                        f"in '{sheet_name}': lat={raw_lat!r} lon={raw_lon!r}"
                    )

                continue

            project_id = None

            for field in ID_FIELD_CANDIDATES:

                if field in record and record[field] is not None:
                    project_id = record[field]
                    break

            category = record.get(category_field) if category_field else None

            popup_lines = []

            for field in popup_fields:

                value = record.get(field)

                if value is not None and str(value).strip():
                    popup_lines.append(f"<b>{field}:</b> {value}")

            projects.append({
                "sheet": sheet_name,
                "id": project_id,
                "name": record.get(NAME_FIELD),
                "category": category,
                "lat": lat,
                "lon": lon,
                "popup_lines": popup_lines,
                "color": CATEGORY_COLORS.get((sheet_name, category), DEFAULT_COLOR),
            })

    return projects


def build_map(projects, output_path):

    if not projects:
        raise ValueError("No geocoded projects found -- nothing to map.")

    center_lat = sum(p["lat"] for p in projects) / len(projects)
    center_lon = sum(p["lon"] for p in projects) / len(projects)

    fmap = folium.Map(
        location=[center_lat, center_lon],
        zoom_start=7,
        tiles="OpenStreetMap",
    )

    # One toggleable layer per sheet, so you can hide/show DESC vs
    # GA ITS independently from the layer control in the corner.
    layers = {}

    for p in projects:

        sheet = p["sheet"]

        if sheet not in layers:
            layers[sheet] = folium.FeatureGroup(name=sheet)
            layers[sheet].add_to(fmap)

        title_bits = [str(p["name"])]

        if p["id"] is not None:
            title_bits.append(f"[{p['id']}]")

        popup_html = f"<b>{' '.join(title_bits)}</b><br>"
        popup_html += "<br>".join(p["popup_lines"])

        folium.Marker(
            location=[p["lat"], p["lon"]],
            tooltip=" ".join(title_bits),
            popup=folium.Popup(popup_html, max_width=350),
            icon=folium.Icon(color=p["color"]),
        ).add_to(layers[sheet])

    folium.LayerControl(collapsed=False).add_to(fmap)

    fmap.save(output_path)


def main():

    projects = load_projects(EXCEL_FILE)

    print(f"Loaded {len(projects)} geocoded projects.")

    build_map(projects, OUTPUT_HTML)

    print(f"Map saved to: {OUTPUT_HTML}")
    print("Open that file in a browser to view it.")


if __name__ == "__main__":
    main()