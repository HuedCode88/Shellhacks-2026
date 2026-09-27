import os

import openpyxl
import folium

from build_overlap_table import (
    load_desc_projects,
    load_gpc_projects,
    build_overlap_table,
)


# ============================================================
# FILES
# ============================================================

EXCEL_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "gridlock_project_tables_geocoded_with_descriptions(1).xlsx"
)

OUTPUT_HTML = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "project_map.html"
)


# ============================================================
# PROJECT MARKER SETTINGS
# ============================================================

SHEET_CATEGORY_FIELD = {
    "DESC Geocoded": "Status",
    "GA ITS Geocoded": "Sponsor (GPC/GTC/MEAG/DU/SAV)",
}


CATEGORY_COLORS = {
    ("DESC Geocoded", "In Progress"): "#28a745",
    ("DESC Geocoded", "Planned"): "#fd7e14",

    ("GA ITS Geocoded", "GPC"): "#007bff",
    ("GA ITS Geocoded", "GTC"): "#6f42c1",
    ("GA ITS Geocoded", "SAV"): "#e83e8c",
    ("GA ITS Geocoded", "MEAG"): "#17a2b8",
    ("GA ITS Geocoded", "DU"): "#8b0000",
}



DEFAULT_COLOR = "gray"


ID_FIELD_CANDIDATES = [
    "Project ID",
    "TEAMS Number",
]


NAME_FIELD = "Project Name / Endpoints (raw title)"


POPUP_FIELDS = {

    "DESC Geocoded": [
        "Status",
        "Planned In-Service Date",
        "Description",
        "confidence",
    ],

    "GA ITS Geocoded": [
        "Sponsor (GPC/GTC/MEAG/DU/SAV)",
        "Need / In-Service Date",
        "Zone",
        "Description",
        "confidence",
    ],
}


# ============================================================
# OVERLAP COLORS
# ============================================================

OVERLAP_COLORS = {

    "Touching / Crossing -- must coordinate (outage timing, crossing structures)":
        "#00FFFF",

    "Under 1.6 km -- can share the right-of-way (access roads, permits)":
        "#00FFFF",

    "Under 8 km -- can share site logistics (laydown yards, deliveries)":
        "#00FFFF",

    "Under 40 km -- can share crews & equipment":
        "#00FFFF",
}



# ============================================================
# HELPERS
# ============================================================

def safe_float(value):

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

        parts = [
            p.strip()
            for p in text.split(",")
            if p.strip()
        ]

        for part in reversed(parts):

            try:
                return float(part)

            except ValueError:
                continue

    return None


def _row_dict(header, row):

    return {
        header[i]: row[i]
        for i in range(len(header))
    }


# ============================================================
# LOAD PROJECTS FOR MAP
# ============================================================

def load_projects(path):

    wb = openpyxl.load_workbook(
        path,
        data_only=True
    )

    projects = []

    for sheet_name in wb.sheetnames:

        ws = wb[sheet_name]

        rows = list(
            ws.iter_rows(
                values_only=True
            )
        )

        if not rows:
            continue

        header = rows[0]

        category_field = (
            SHEET_CATEGORY_FIELD.get(
                sheet_name
            )
        )

        popup_fields = (
            POPUP_FIELDS.get(
                sheet_name,
                []
            )
        )

        for row in rows[1:]:

            record = _row_dict(
                header,
                row
            )

            lat = safe_float(
                record.get("center_lat")
            )

            lon = safe_float(
                record.get("center_lon")
            )

            if lat is None or lon is None:
                continue

            # ------------------------------------------------
            # Project ID
            # ------------------------------------------------

            project_id = None

            for field in ID_FIELD_CANDIDATES:

                if (
                    field in record
                    and record[field] is not None
                ):

                    project_id = record[field]

                    break

            # ------------------------------------------------
            # Category
            # ------------------------------------------------

            category = None

            if category_field:

                category = record.get(
                    category_field
                )

            # ------------------------------------------------
            # Popup
            # ------------------------------------------------

            popup_lines = []

            for field in popup_fields:

                value = record.get(
                    field
                )

                if (
                    value is not None
                    and str(value).strip()
                ):

                    popup_lines.append(
                        f"<b>{field}:</b> {value}"
                    )

            # ------------------------------------------------
            # Store project
            # ------------------------------------------------

            projects.append({

                "sheet": sheet_name,

                "id": project_id,

                "name": record.get(
                    NAME_FIELD
                ),

                "category": category,

                "lat": lat,

                "lon": lon,

                "popup_lines": popup_lines,

                "color": CATEGORY_COLORS.get(
                    (
                        sheet_name,
                        category
                    ),
                    DEFAULT_COLOR
                ),

            })

    return projects


# ============================================================
# BUILD LOOKUP FOR PROJECTS
# ============================================================

def build_project_lookup(projects):

    lookup = {}

    for project in projects:

        project_id = project["id"]

        if project_id is None:
            continue

        key = (
            project["sheet"],
            str(project_id).strip()
        )

        lookup[key] = project

    return lookup


# ============================================================
# ADD OVERLAP LINES
# ============================================================

def add_overlap_lines(
    fmap,
    overlaps,
    project_lookup
):

    overlap_layer = folium.FeatureGroup(
        name="Project Overlaps",
        show=True
    )

    overlap_count = 0

    for overlap in overlaps:

        # ----------------------------------------------------
        # Find DESC project
        # ----------------------------------------------------

        desc_key = (
            "DESC Geocoded",
            str(
                overlap["desc_id"]
            ).strip()
        )

        desc_project = project_lookup.get(
            desc_key
        )

        # ----------------------------------------------------
        # Find GPC project
        # ----------------------------------------------------

        gpc_key = (
            "GA ITS Geocoded",
            str(
                overlap["gpc_id"]
            ).strip()
        )

        gpc_project = project_lookup.get(
            gpc_key
        )

        # If either project can't be found on
        # the map, don't draw the line.
        if (
            desc_project is None
            or gpc_project is None
        ):
            print(
                "Could not map overlap:",
                overlap["desc_id"],
                overlap["gpc_id"]
            )

            continue

        # ----------------------------------------------------
        # Coordinates
        # ----------------------------------------------------

        start = [
            desc_project["lat"],
            desc_project["lon"]
        ]

        end = [
            gpc_project["lat"],
            gpc_project["lon"]
        ]

        # ----------------------------------------------------
        # Color
        # ----------------------------------------------------

        color = OVERLAP_COLORS.get(
            overlap["geographic_tier"],
            "red"
        )

        # ----------------------------------------------------
        # Timeline
        # ----------------------------------------------------

        timeline = overlap[
            "timeline_overlap"
        ]

        if timeline is None:

            timeline_text = "Unknown"

        elif timeline:

            timeline_text = "Yes"

        else:

            timeline_text = "No"

        # ----------------------------------------------------
        # Popup
        # ----------------------------------------------------

        desc_name = (
            desc_project["name"]
            or overlap["desc_id"]
        )

        gpc_name = (
            gpc_project["name"]
            or overlap["gpc_id"]
        )

        popup_html = f"""
        <div style="font-family: Arial;">

            <h4 style="margin-bottom: 8px;">
                Project Overlap
            </h4>

            <b>DESC:</b><br>
            {desc_name}<br>
            ID: {overlap["desc_id"]}

            <br><br>

            <b>GPC:</b><br>
            {gpc_name}<br>
            ID: {overlap["gpc_id"]}

            <br><br>

            <b>Distance:</b>
            {overlap["distance_km"]} km
            ({overlap["distance_mi"]} mi)

            <br><br>

            <b>Geographic tier:</b><br>
            {overlap["geographic_tier"]}

            <br><br>

            <b>Timeline overlap:</b>
            {timeline_text}

            <br>

            <b>In-service date gap:</b>
            {overlap["day_gap"] if overlap["day_gap"] is not None else "Unknown"}
            days

        </div>
        """

        # ----------------------------------------------------
        # Draw line
        # ----------------------------------------------------

        folium.PolyLine(

            locations=[
                start,
                end
            ],

            color="#111111",

            weight=5,

            opacity=0.9,

        ).add_to(overlap_layer)


        folium.PolyLine(

            locations=[
                start,
                end
            ],

            color="#00FFFF",

            weight=3,

            opacity=1.0,

            tooltip=(
                f"{desc_name} ↔ "
                f"{gpc_name} | "
                f"{overlap['distance_mi']} mi"
            ),

            popup=folium.Popup(
                popup_html,
                max_width=400
            ),

        ).add_to(overlap_layer)


        overlap_count += 1

    overlap_layer.add_to(
        fmap
    )

    print(
        f"Added {overlap_count} overlap lines to map."
    )

# ============================================================
# BUILD POPUP
# ============================================================
def add_overlap_ranking_panel(
    fmap,
    overlaps,
    project_lookup
):
    """
    Adds a collapsible overlap-ranking button to the map.

    Clicking an overlap entry:
      - zooms the map to both projects
      - highlights the two projects
      - draws attention to the selected pair
    """

    if not overlaps:
        return

    from branca.element import Element

    # --------------------------------------------------------
    # Build table rows
    # --------------------------------------------------------

    rows_html = ""

    for rank, overlap in enumerate(
        overlaps,
        start=1
    ):

        desc_id = overlap["desc_id"]
        gpc_id = overlap["gpc_id"]

        desc_name = (
            overlap["desc_name"]
            or "Unknown DESC Project"
        )

        gpc_name = (
            overlap["gpc_name"]
            or "Unknown GPC Project"
        )

        distance_km = overlap[
            "distance_km"
        ]

        distance_mi = overlap[
            "distance_mi"
        ]

        day_gap = overlap[
            "day_gap"
        ]

        if day_gap is None:
            time_text = "Unknown"
        else:
            time_text = f"{day_gap} days"

        tier = overlap[
            "geographic_tier"
        ]

        # ----------------------------------------------------
        # Color by geographic tier
        # ----------------------------------------------------

        if tier.startswith(
            "Touching"
        ):

            color = "#dc3545"

        elif tier.startswith(
            "Under 1.6"
        ):

            color = "#fd7e14"

        elif tier.startswith(
            "Under 8"
        ):

            color = "#6f42c1"

        else:

            color = "#0d6efd"

        # ----------------------------------------------------
        # Find the projects
        # ----------------------------------------------------

        desc_key = (
            "DESC Geocoded",
            str(desc_id).strip()
        )

        gpc_key = (
            "GA ITS Geocoded",
            str(gpc_id).strip()
        )

        desc_project = project_lookup.get(
            desc_key
        )

        gpc_project = project_lookup.get(
            gpc_key
        )

        if (
            desc_project is None
            or gpc_project is None
        ):
            continue

        # ----------------------------------------------------
        # JavaScript coordinates
        # ----------------------------------------------------

        desc_lat = desc_project["lat"]
        desc_lon = desc_project["lon"]

        gpc_lat = gpc_project["lat"]
        gpc_lon = gpc_project["lon"]

        # ----------------------------------------------------
        # Row
        # ----------------------------------------------------

        rows_html += f"""

        <div
            class="overlap-row"
            onclick="
                focusOverlap(
                    {desc_lat},
                    {desc_lon},
                    {gpc_lat},
                    {gpc_lon},
                    {rank}
                );
            "
            style="
                border-bottom: 1px solid #ddd;
                padding: 10px 8px;
                cursor: pointer;
                transition: background 0.15s;
            "
            onmouseover="
                this.style.background='#f0f0f0';
            "
            onmouseout="
                this.style.background='white';
            "
        >

            <div style="
                display: flex;
                align-items: center;
                margin-bottom: 5px;
            ">

                <span style="
                    background: {color};
                    color: white;

                    border-radius: 50%;

                    width: 25px;
                    height: 25px;

                    display: inline-flex;

                    align-items: center;
                    justify-content: center;

                    font-weight: bold;

                    margin-right: 8px;
                ">
                    {rank}
                </span>

                <strong>
                    {desc_id} ↔ {gpc_id}
                </strong>

            </div>

            <div style="
                font-size: 12px;
                color: #444;
                margin-left: 33px;
            ">

                <div>
                    <b>DESC:</b>
                    {desc_name}
                </div>

                <div>
                    <b>GPC:</b>
                    {gpc_name}
                </div>

                <div style="
                    margin-top: 5px;
                    color: #222;
                ">

                    <b>Distance:</b>
                    {distance_km:.2f} km
                    ({distance_mi:.2f} mi)

                    &nbsp; | &nbsp;

                    <b>Time:</b>
                    {time_text}

                </div>

                <div style="
                    margin-top: 3px;
                    color: {color};
                    font-weight: bold;
                ">

                    {tier}

                </div>

            </div>

        </div>

        """

    # --------------------------------------------------------
    # Complete panel
    # --------------------------------------------------------

    panel_html = f"""

    <!-- OVERLAP BUTTON -->

    <div id="overlap-container"
        style="
            position: fixed;

            bottom: 20px;
            right: 20px;

            z-index: 9999;

            font-family: Arial, sans-serif;
        "
    >

        <!-- COLLAPSED BUTTON -->

        <button
            id="overlap-toggle"
            onclick="toggleOverlapPanel()"
            style="
                background: #222;
                color: white;

                border: none;
                border-radius: 6px;

                padding: 11px 16px;

                font-size: 14px;
                font-weight: bold;

                cursor: pointer;

                box-shadow:
                    0 3px 10px
                    rgba(0,0,0,0.35);
            "
        >

            ⚠ Project Overlaps
            ({len(overlaps)})

        </button>


        <!-- PANEL -->

        <div
            id="overlap-panel"
            style="
                display: none;

                width: 400px;

                max-height: 80vh;

                margin-top: 8px;

                background: white;

                border: 2px solid #333;

                border-radius: 8px;

                box-shadow:
                    0 3px 15px
                    rgba(0,0,0,0.35);

                overflow: hidden;
            "
        >

            <!-- HEADER -->

            <div style="
                background: #222;
                color: white;

                padding: 12px;

                font-size: 16px;
                font-weight: bold;
            ">

                Project Overlap Ranking

                <span style="
                    float: right;

                    font-size: 12px;

                    font-weight: normal;

                    opacity: 0.8;
                ">

                    {len(overlaps)} overlaps

                </span>

            </div>


            <!-- DESCRIPTION -->

            <div style="
                padding: 8px 12px;

                background: #f4f4f4;

                border-bottom:
                    1px solid #ccc;

                font-size: 12px;

                color: #555;
            ">

                Ranked by geographic distance.
                Click an entry to focus the map.

            </div>


            <!-- TABLE -->

            <div style="
                max-height:
                    calc(80vh - 110px);

                overflow-y: auto;
            ">

                {rows_html}

            </div>

        </div>

    </div>


    <!-- JAVASCRIPT -->

    <script>

        function toggleOverlapPanel() {{

            var panel =
                document.getElementById(
                    "overlap-panel"
                );

            var button =
                document.getElementById(
                    "overlap-toggle"
                );

            if (
                panel.style.display === "none"
                || panel.style.display === ""
            ) {{

                panel.style.display = "block";

                button.innerHTML =
                    "✕ Close Overlaps";

            }} else {{

                panel.style.display = "none";

                button.innerHTML =
                    "⚠ Project Overlaps ({len(overlaps)})";

            }}

        }}


        function focusOverlap(
            lat1,
            lon1,
            lat2,
            lon2,
            rank
        ) {{

            /*
             * Find the Leaflet map generated by Folium.
             */

            var mapObject = null;

            for (
                var key in window
            ) {{

                if (
                    key.startsWith("map_")
                    &&
                    window[key]
                    &&
                    typeof window[key].fitBounds
                        === "function"
                ) {{

                    mapObject =
                        window[key];

                    break;

                }}

            }}


            if (!mapObject) {{

                console.error(
                    "Could not find Leaflet map."
                );

                return;

            }}


            /*
             * Create bounds around the
             * two project locations.
             */

            var bounds = [
                [lat1, lon1],
                [lat2, lon2]
            ];


            /*
             * Zoom map to both projects.
             */

            mapObject.fitBounds(
                bounds,
                {{
                    padding: [
                        100,
                        100
                    ],

                    maxZoom: 12
                }}
            );


            /*
             * Add a temporary highlighted
             * connection between the two.
             */

            if (
                window.activeOverlapLine
            ) {{

                mapObject.removeLayer(
                    window.activeOverlapLine
                );

            }}


            window.activeOverlapLine =
                L.polyline(
                    bounds,
                    {{

                        color: "#ff0000",

                        weight: 8,

                        opacity: 0.9,

                        dashArray: "10, 8"

                    }}
                ).addTo(
                    mapObject
                );


            /*
             * Remove highlight after 5 seconds.
             */

            setTimeout(
                function() {{

                    if (
                        window.activeOverlapLine
                    ) {{

                        mapObject.removeLayer(
                            window.activeOverlapLine
                        );

                        window.activeOverlapLine =
                            null;

                    }}

                }},
                5000
            );

        }}

    </script>

    """

    fmap.get_root().html.add_child(
        Element(panel_html)
    )

def add_map_legend(fmap):
    """
    Adds a collapsible legend to the left side of the map.

    Closed:
        The legend is hidden off-screen and only the > tab is visible.

    Open:
        The legend slides out and the arrow changes to <.
    """

    from branca.element import Element

    legend_html = """
    <style>
        #map-legend-container {
            position: fixed;
            bottom: 20px;
            left: 0;
            z-index: 9999;
            font-family: Arial, sans-serif;
        }

        #map-legend-panel {
            position: relative;

            width: 220px;

            background: white;

            border: 2px solid #333;
            border-left: none;

            border-radius: 0 8px 8px 0;

            padding: 14px 16px;

            box-shadow:
                0 3px 12px rgba(0,0,0,0.3);

            transition:
                transform 0.3s ease;
        }

        #map-legend-container.legend-hidden
        #map-legend-panel {
            transform: translateX(-220px);
        }

        #map-legend-toggle {
            position: absolute;

            right: -38px;
            top: 50%;

            transform: translateY(-50%);

            width: 38px;
            height: 70px;

            border: 2px solid #333;
            border-left: none;

            border-radius: 0 8px 8px 0;

            background: white;

            cursor: pointer;

            font-size: 22px;
            font-weight: bold;

            box-shadow:
                3px 0 8px rgba(0,0,0,0.2);
        }

        #map-legend-toggle:hover {
            background: #f0f0f0;
        }

        .legend-title {
            font-size: 15px;
            font-weight: bold;
            margin-bottom: 10px;
        }

        .legend-section {
            font-weight: bold;
            margin-top: 8px;
            margin-bottom: 5px;
        }

        .legend-item {
            display: flex;
            align-items: center;
            margin-bottom: 6px;
        }

        .legend-color {
            width: 14px;
            height: 14px;
            border-radius: 50%;
            display: inline-block;
            margin-right: 8px;
            flex-shrink: 0;
        }
    </style>

    <div
        id="map-legend-container"
        class="legend-hidden"
    >

        <div id="map-legend-panel">

            <button
                id="map-legend-toggle"
                onclick="toggleMapLegend()"
                title="Show legend"
            >❯</button>

            <div class="legend-title">
                Project Legend
            </div>

            <div class="legend-section">
                DESC
            </div>

            <div class="legend-item">
                <span
                    class="legend-color"
                    style="background: green;"
                ></span>
                In Progress
            </div>

            <div class="legend-item">
                <span
                    class="legend-color"
                    style="background: orange;"
                ></span>
                Planned
            </div>

            <div class="legend-section">
                GA ITS
            </div>

            <div class="legend-item">
                <span
                    class="legend-color"
                    style="background: blue;"
                ></span>
                GPC
            </div>

            <div class="legend-item">
                <span
                    class="legend-color"
                    style="background: purple;"
                ></span>
                GTC
            </div>

            <div class="legend-item">
                <span
                    class="legend-color"
                    style="background: pink;"
                ></span>
                SAV
            </div>

            <div class="legend-item">
                <span
                    class="legend-color"
                    style="background: cadetblue;"
                ></span>
                MEAG
            </div>

            <div class="legend-item">
                <span
                    class="legend-color"
                    style="background: darkred;"
                ></span>
                DU
            </div>

            <div style="
                border-top: 1px solid #ccc;
                margin-top: 10px;
                padding-top: 10px;
            ">

                <div class="legend-section">
                    Connections
                </div>

                <div class="legend-item">

                    <span style="
                        width: 24px;
                        height: 3px;
                        background: #00FFFF;
                        display: inline-block;
                        margin-right: 8px;
                        flex-shrink: 0;
                    "></span>

                    Project Overlap

                </div>

            </div>

        </div>

    </div>

    <script>
        function toggleMapLegend() {

            var container =
                document.getElementById(
                    "map-legend-container"
                );

            var button =
                document.getElementById(
                    "map-legend-toggle"
                );

            if (
                container.classList.contains(
                    "legend-hidden"
                )
            ) {

                container.classList.remove(
                    "legend-hidden"
                );

                button.innerHTML = "❮";

                button.title =
                    "Hide legend";

            } else {

                container.classList.add(
                    "legend-hidden"
                );

                button.innerHTML = "❯";

                button.title =
                    "Show legend";
            }
        }
    </script>
    """

    # IMPORTANT:
    # Jinja interprets {{ }} as template expressions.
    # Escape them before giving the HTML to Branca.
    legend_html = legend_html.replace(
        "{{",
        "{% raw %}{{{% endraw %}"
    ).replace(
        "}}",
        "{% raw %}}}{% endraw %}"
    )

    fmap.get_root().html.add_child(
        Element(legend_html)
    )


def add_project_search(fmap, projects):
    """
    Adds a project search bar to the top-right of the map.

    Searches:
        - Project ID
        - Project name
        - Sheet / utility
        - Category

    Clicking a result zooms the map to that project.
    """

    from branca.element import Element
    import json
    import html

    search_projects = []

    for project in projects:

        search_projects.append({
            "id": str(project["id"])
            if project["id"] is not None
            else "",

            "name": str(project["name"])
            if project["name"] is not None
            else "",

            "sheet": str(project["sheet"])
            if project["sheet"] is not None
            else "",

            "category": str(project["category"])
            if project["category"] is not None
            else "",

            "lat": project["lat"],
            "lon": project["lon"],
        })

    projects_json = json.dumps(
        search_projects
    )

    search_html = f"""
    <style>

        #project-search-container {{
            position: fixed;

            top: 20px;
            left: 50%;

            transform: translateX(-50%);

            z-index: 9999;

            font-family: Arial, sans-serif;

            width: 330px;
        }}


        #project-search-box {{
            display: flex;

            background: white;

            border-radius: 7px;

            box-shadow:
                0 3px 12px
                rgba(0,0,0,0.3);

            overflow: hidden;
        }}


        #project-search-input {{
            flex: 1;

            border: none;

            padding: 11px 12px;

            font-size: 14px;

            outline: none;

            min-width: 0;
        }}


        #project-search-button {{
            border: none;

            background: #222;

            color: white;

            padding: 0 14px;

            cursor: pointer;

            font-size: 16px;
        }}


        #project-search-button:hover {{
            background: #444;
        }}


        #project-search-results {{
            display: none;

            margin-top: 6px;

            background: white;

            border-radius: 7px;

            box-shadow:
                0 3px 12px
                rgba(0,0,0,0.3);

            max-height: 400px;

            overflow-y: auto;
        }}


        .project-search-result {{
            padding: 10px 12px;

            border-bottom:
                1px solid #ddd;

            cursor: pointer;

            font-size: 13px;
        }}


        .project-search-result:hover {{
            background: #f0f0f0;
        }}


        .project-search-id {{
            font-weight: bold;

            font-size: 14px;

            margin-bottom: 3px;
        }}


        .project-search-name {{
            color: #444;

            line-height: 1.3;
        }}


        .project-search-meta {{
            color: #777;

            font-size: 11px;

            margin-top: 4px;
        }}


        .project-search-no-results {{
            padding: 12px;

            color: #666;

            font-size: 13px;
        }}

    </style>


    <div id="project-search-container">

        <div id="project-search-box">

            <input
                id="project-search-input"
                type="text"
                placeholder="Search projects..."
                autocomplete="off"
            >

            <button
                id="project-search-button"
                onclick="searchProjects()"
                title="Search"
            >
                
            </button>

        </div>


        <div id="project-search-results"></div>

    </div>


    <script>

        var projectSearchData =
            {projects_json};


        function searchProjects() {{

            var input =
                document.getElementById(
                    "project-search-input"
                );

            var results =
                document.getElementById(
                    "project-search-results"
                );


            var query =
                input.value
                    .trim()
                    .toLowerCase();


            if (!query) {{

                results.style.display = "none";

                results.innerHTML = "";

                return;

            }}


            /*
             * Search ID, name, utility,
             * and category.
             */

            var matches =
                projectSearchData.filter(
                    function(project) {{

                        var searchable =
                            (
                                project.id +
                                " " +
                                project.name +
                                " " +
                                project.sheet +
                                " " +
                                project.category
                            ).toLowerCase();

                        return searchable.includes(
                            query
                        );

                    }}
                );


            /*
             * Limit displayed results.
             */

            matches =
                matches.slice(0, 15);


            if (matches.length === 0) {{

                results.innerHTML =
                    '<div class="project-search-no-results">' +
                    'No matching projects found.' +
                    '</div>';

                results.style.display =
                    "block";

                return;

            }}


            /*
             * Build result list.
             */

            var html = "";


            matches.forEach(
                function(project) {{

                    var safeId =
                        escapeSearchHtml(
                            project.id
                        );

                    var safeName =
                        escapeSearchHtml(
                            project.name
                        );

                    var safeSheet =
                        escapeSearchHtml(
                            project.sheet
                        );

                    var safeCategory =
                        escapeSearchHtml(
                            project.category
                        );


                    html +=
                        '<div ' +
                        'class="project-search-result" ' +
                        'onclick="focusProject(' +
                        project.lat + ',' +
                        project.lon +
                        ')">' +

                            '<div class="project-search-id">' +
                                safeId +
                            '</div>' +

                            '<div class="project-search-name">' +
                                safeName +
                            '</div>' +

                            '<div class="project-search-meta">' +
                                safeSheet +
                                (
                                    safeCategory
                                    ? " • " + safeCategory
                                    : ""
                                ) +
                            '</div>' +

                        '</div>';

                }}
            );


            results.innerHTML =
                html;

            results.style.display =
                "block";

        }}


        function focusProject(
            lat,
            lon
        ) {{

            /*
             * Find Folium's Leaflet map.
             */

            var mapObject = null;


            for (
                var key in window
            ) {{

                if (
                    key.startsWith("map_") &&
                    window[key] &&
                    typeof window[key].setView
                        === "function"
                ) {{

                    mapObject =
                        window[key];

                    break;

                }}

            }}


            if (!mapObject) {{

                console.error(
                    "Could not find Leaflet map."
                );

                return;

            }}


            /*
             * Zoom to project.
             */

            mapObject.setView(
                [lat, lon],
                13
            );


            /*
             * Close results.
             */

            document.getElementById(
                "project-search-results"
            ).style.display =
                "none";


            /*
             * Clear search box.
             */

            document.getElementById(
                "project-search-input"
            ).value = "";

        }}


        function escapeSearchHtml(
            value
        ) {{

            return String(value)
                .replace(/&/g, "&amp;")
                .replace(/</g, "&lt;")
                .replace(/>/g, "&gt;")
                .replace(/"/g, "&quot;")
                .replace(/'/g, "&#039;");

        }}


        /*
         * Pressing Enter searches.
         */

        document
            .getElementById(
                "project-search-input"
            )
            .addEventListener(
                "keydown",
                function(event) {{

                    if (
                        event.key === "Enter"
                    ) {{

                        searchProjects();

                    }}

                }}
            );


        /*
         * Search while typing.
         *
         * This makes it feel more like
         * an autocomplete box.
         */

        document
            .getElementById(
                "project-search-input"
            )
            .addEventListener(
                "input",
                function() {{

                    searchProjects();

                }}
            );

    </script>
    """

    fmap.get_root().html.add_child(
        Element(search_html)
    )



# ============================================================
# BUILD MAP
# ============================================================

def build_map(
    projects,
    overlaps,
    output_path
):

    if not projects:

        raise ValueError(
            "No geocoded projects found."
        )

    # --------------------------------------------------------
    # Center map
    # --------------------------------------------------------

    center_lat = (
        sum(
            p["lat"]
            for p in projects
        )
        / len(projects)
    )

    center_lon = (
        sum(
            p["lon"]
            for p in projects
        )
        / len(projects)
    )

    fmap = folium.Map(

        location=[
            center_lat,
            center_lon
        ],

        zoom_start=7,

        tiles="OpenStreetMap",
    )

    add_project_search(
        fmap,
        projects
    )


    # --------------------------------------------------------
    # Project layers
    # --------------------------------------------------------

    layers = {}

    for project in projects:

        sheet = project["sheet"]

        if sheet not in layers:

            layers[sheet] = (
                folium.FeatureGroup(
                    name=sheet
                )
            )

            layers[sheet].add_to(
                fmap
            )

        title_bits = [
            str(
                project["name"]
            )
        ]

        if project["id"] is not None:

            title_bits.append(
                f"[{project['id']}]"
            )

        popup_html = (
            f"<b>{' '.join(title_bits)}</b>"
            "<br>"
        )

        popup_html += (
            "<br>".join(
                project["popup_lines"]
            )
        )

        folium.CircleMarker(

            location=[
                project["lat"],
                project["lon"]
            ],

            radius=7,

            color="#333333",

            weight=1,

            fill=True,

            fill_color=project["color"],

            fill_opacity=0.9,

            tooltip=" ".join(
                title_bits
            ),

            popup=folium.Popup(
                popup_html,
                max_width=350
            ),

        ).add_to(
            layers[sheet]
        )


    # --------------------------------------------------------
    # OVERLAPS
    # --------------------------------------------------------

    project_lookup = (
        build_project_lookup(
            projects
        )
    )

    add_overlap_lines(
        fmap,
        overlaps,
        project_lookup
    )

    add_overlap_ranking_panel(
        fmap,
        overlaps,
        project_lookup
    )


    # --------------------------------------------------------
    # Layer controls
    # --------------------------------------------------------

    folium.LayerControl(
        collapsed=False
    ).add_to(
        fmap
    )

    add_map_legend(fmap)


    # --------------------------------------------------------
    # Save
    # --------------------------------------------------------

    fmap.save(
        output_path
    )


# ============================================================
# MAIN
# ============================================================

def main():

    print(
        "Loading project map data..."
    )

    projects = load_projects(
        EXCEL_FILE
    )

    print(
        f"Loaded {len(projects)} "
        f"geocoded projects."
    )

    # --------------------------------------------------------
    # Load the same projects used by
    # build_overlap_table.py
    # --------------------------------------------------------

    print(
        "Calculating project overlaps..."
    )

    wb = openpyxl.load_workbook(
        EXCEL_FILE,
        data_only=True
    )

    desc_projects = (
        load_desc_projects(wb)
    )

    gpc_projects = (
        load_gpc_projects(wb)
    )

    print(
        f"DESC projects: "
        f"{len(desc_projects)}"
    )

    print(
        f"GPC projects: "
        f"{len(gpc_projects)}"
    )

    # --------------------------------------------------------
    # Calculate overlaps
    # --------------------------------------------------------

    overlaps = build_overlap_table(
        desc_projects,
        gpc_projects
    )

    print(
        f"Found {len(overlaps)} "
        f"overlapping project pairs."
    )

    # --------------------------------------------------------
    # Build map
    # --------------------------------------------------------

    build_map(
        projects,
        overlaps,
        OUTPUT_HTML
    )

    print()
    print(
        f"Map saved to:"
    )
    print(
        OUTPUT_HTML
    )


# ============================================================
# START
# ============================================================

if __name__ == "__main__":
    main()
