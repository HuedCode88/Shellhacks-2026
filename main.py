"""
Maps every geocoded project in the workbook onto an interactive
OpenStreetMap-based webpage (via Folium/Leaflet).

Unlike the in-chat map tool, this has no marker limit, combines
both utilities on one map, and can be re-run any time the workbook
is updated -- just run:

    pip install openpyxl folium
    python main.py

Open the local URL printed by the script in a browser. The map reads
the workbook again whenever the page is requested, so refreshing it
shows workbook updates. The workbook must be in this folder.
"""

import os
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from threading import Lock
from urllib.parse import urlsplit

import openpyxl
import folium
from folium.plugins import MarkerCluster

EXCEL_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "gridlock_project_tables_geocoded_nominatim (1).xlsx"
)

MAP_RENDER_LOCK = Lock()

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


def build_map(projects):

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

    return fmap.get_root().render()


class ProjectMapHandler(BaseHTTPRequestHandler):

    def do_GET(self):

        if urlsplit(self.path).path not in ("/", "/project_map.html"):
            self.send_error(404)
            return

        try:
            with MAP_RENDER_LOCK:
                projects = load_projects(EXCEL_FILE)
                page = build_map(projects).encode("utf-8")
        except (OSError, ValueError) as error:
            self.send_error(500, str(error))
            return

        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(page)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(page)


def serve_map():

    server = None

    for port in range(8000, 8011):
        try:
            server = ThreadingHTTPServer(("127.0.0.1", port), ProjectMapHandler)
            break
        except OSError:
            continue

    if server is None:
        raise RuntimeError("No available local port between 8000 and 8010.")

    print(f"Interactive map available at http://127.0.0.1:{server.server_port}/")
    print("Refresh the page to load the latest workbook data. Press Ctrl+C to stop.")

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nMap server stopped.")
    finally:
        server.server_close()


def main():

    projects = load_projects(EXCEL_FILE)

    print(f"Loaded {len(projects)} geocoded projects.")
    serve_map()


if __name__ == "__main__":
    main()