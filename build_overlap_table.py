"""
Stage 5 -- Geographic overlap (closest points, not centers).

Two projects overlap if the CLOSEST POINTS between them are under
40 km (25 mi) -- not their centers. A project is either:

  - a single point  (only one endpoint was geocoded), or
  - a line segment  (both endpoints were geocoded)

and we need the true minimum distance between whatever geometry
each side has: point-to-point, point-to-segment, or
segment-to-segment (including detecting an actual crossing, which
counts as distance 0 regardless of endpoint positions).

Distance is computed by converting lat/lon to a local flat
kilometer grid for each pair (fine at this scale -- straight-line/
haversine-equivalent accuracy is all the brief asks for) and then
running standard 2D segment-distance geometry on that grid.

Overlaps are ranked into four coordination tiers by how close they
are, and timeline overlap (in-service dates falling in the same
build window) is attached as a secondary signal alongside the
primary geographic one.
"""

import math
import os
import openpyxl
from dateutil import parser as dateparser

SOURCE_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "gridlock_project_tables_geocoded_nominatim(1).xlsx"
)

OUTPUT_XLSX = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "overlap_table.xlsx"
)

# --- Distance thresholds, in km ---------------------------------
KM_PER_MILE = 1.60934

TOUCH_EPS_KM = 0.05          # essentially touching/crossing
ROW_SHARE_KM = 1.6           # ~1 mi -- can share the right-of-way itself
SITE_SHARE_KM = 8.0          # ~5 mi -- can share laydown yards/deliveries
OVERLAP_THRESHOLD_KM = 40.0  # ~25 mi -- can share crews/equipment; beyond this, ignore

# How close two in-service dates need to be (in days) to call it
# the "same build window." The brief doesn't pin an exact number --
# this is a reasonable default (roughly a construction season) and
# is meant to be tuned; it only affects the secondary flag, not
# whether a row appears in the table at all.
TIMELINE_WINDOW_DAYS = 365

# Coordination tiers, closest first. Each is (max_km, label).
# max_km is exclusive except the first (touching), which uses <=.
TIERS = [
    (TOUCH_EPS_KM, "Touching / Crossing -- must coordinate (outage timing, crossing structures)"),
    (ROW_SHARE_KM, "Under 1.6 km -- can share the right-of-way (access roads, permits)"),
    (SITE_SHARE_KM, "Under 8 km -- can share site logistics (laydown yards, deliveries)"),
    (OVERLAP_THRESHOLD_KM, "Under 40 km -- can share crews & equipment"),
]


# ------------------------------------------------------------
# Load raw rows
# ------------------------------------------------------------

def _row_dicts(ws):
    rows = list(ws.iter_rows(values_only=True))
    header = rows[0]
    return [dict(zip(header, row)) for row in rows[1:]]


def safe_float(value):
    """Coerces a coordinate cell to float, tolerating stray commas/whitespace."""

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


def parse_date(value):
    """Best-effort date parse. Returns None if empty/unparseable."""

    if value is None:
        return None

    if isinstance(value, str) and not value.strip():
        return None

    try:
        return dateparser.parse(str(value)).date()
    except (ValueError, OverflowError):
        return None


def load_desc_projects(wb):
    """
    DESC (Utility A) projects. Each project keeps BOTH endpoints
    when available -- endpoint_2 may be None, meaning the project
    is a single point rather than a line segment.
    """

    projects = []

    for rec in _row_dicts(wb["DESC Geocoded"]):

        lat1 = safe_float(rec.get("lat_1"))
        lon1 = safe_float(rec.get("lon_1"))

        if lat1 is None or lon1 is None:
            continue

        lat2 = safe_float(rec.get("lat_2"))
        lon2 = safe_float(rec.get("lon_2"))

        has_endpoint_2 = lat2 is not None and lon2 is not None

        center_lat = safe_float(rec.get("center_lat"))
        center_lon = safe_float(rec.get("center_lon"))

        if center_lat is None or center_lon is None:
            continue

        projects.append({
            "utility": "DESC",
            "id": rec.get("Project ID"),
            "name": rec.get("Project Name / Endpoints (raw title)"),

            "p1": (center_lat, center_lon),
            "p2": (center_lat, center_lon),

            "has_endpoint_2": False,

            "in_service_date": parse_date(
                rec.get("Planned In-Service Date")
            ),
        })

    return projects


def load_gpc_projects(wb):
    """GPC-sponsored (Utility B) projects from the GA ITS sheet, same shape as above."""

    projects = []

    for rec in _row_dicts(wb["GA ITS Geocoded"]):

        if rec.get("Sponsor (GPC/GTC/MEAG/DU/SAV)") != "GPC":
            continue

        lat1 = safe_float(rec.get("lat_1"))
        lon1 = safe_float(rec.get("lon_1"))

        if lat1 is None or lon1 is None:
            continue

        lat2 = safe_float(rec.get("lat_2"))
        lon2 = safe_float(rec.get("lon_2"))

        has_endpoint_2 = lat2 is not None and lon2 is not None

        center_lat = safe_float(rec.get("center_lat"))
        center_lon = safe_float(rec.get("center_lon"))

        if center_lat is None or center_lon is None:
            continue

        projects.append({
            "utility": "GPC",
            "id": rec.get("TEAMS Number"),
            "name": rec.get("Project Name / Endpoints (raw title)"),

            "p1": (center_lat, center_lon),
            "p2": (center_lat, center_lon),

            "has_endpoint_2": False,

            "in_service_date": parse_date(
                rec.get("Need / In-Service Date")
            ),
        })

    return projects


# ------------------------------------------------------------
# Local flat-km projection (accurate enough under ~100 km)
# ------------------------------------------------------------

KM_PER_DEG_LAT = 110.574


def haversine_km(lat1, lon1, lat2, lon2):
    """
    Great-circle distance between two latitude/longitude points.
    Returns kilometers.
    """

    R = 6371.0088

    lat1 = math.radians(lat1)
    lat2 = math.radians(lat2)

    dlat = lat2 - lat1
    dlon = math.radians(lon2 - lon1)

    a = (
        math.sin(dlat / 2) ** 2
        + math.cos(lat1)
        * math.cos(lat2)
        * math.sin(dlon / 2) ** 2
    )

    c = 2 * math.atan2(
        math.sqrt(a),
        math.sqrt(1 - a)
    )

    return R * c

# ------------------------------------------------------------
# 2D segment geometry
# ------------------------------------------------------------

def _dist(p, q):
    return math.hypot(p[0] - q[0], p[1] - q[1])


def _is_point(a, b, eps=1e-9):
    """True if a segment's two endpoints are (numerically) the same point."""
    return abs(a[0] - b[0]) < eps and abs(a[1] - b[1]) < eps


def _closest_point_on_segment(p, a, b):
    """Closest point to p on segment ab (clamped to the segment)."""

    ax, ay = a
    bx, by = b
    px, py = p

    abx, aby = bx - ax, by - ay
    len_sq = abx * abx + aby * aby

    if len_sq < 1e-12:
        return a

    t = ((px - ax) * abx + (py - ay) * aby) / len_sq
    t = max(0.0, min(1.0, t))

    return (ax + t * abx, ay + t * aby)


def _orientation(p, q, r, eps=1e-9):
    val = (q[1] - p[1]) * (r[0] - q[0]) - (q[0] - p[0]) * (r[1] - q[1])
    if abs(val) < eps:
        return 0
    return 1 if val > 0 else 2


def _on_segment(p, q, r, eps=1e-9):
    return (
        min(p[0], r[0]) - eps <= q[0] <= max(p[0], r[0]) + eps
        and min(p[1], r[1]) - eps <= q[1] <= max(p[1], r[1]) + eps
    )


def _segments_intersect(p1, p2, p3, p4):
    """Standard orientation-based segment intersection test (incl. collinear touches)."""

    o1 = _orientation(p1, p2, p3)
    o2 = _orientation(p1, p2, p4)
    o3 = _orientation(p3, p4, p1)
    o4 = _orientation(p3, p4, p2)

    if o1 != o2 and o3 != o4:
        return True

    if o1 == 0 and _on_segment(p1, p3, p2):
        return True
    if o2 == 0 and _on_segment(p1, p4, p2):
        return True
    if o3 == 0 and _on_segment(p3, p1, p4):
        return True
    if o4 == 0 and _on_segment(p3, p2, p4):
        return True

    return False


def closest_distance_km(a1_ll, a2_ll, b1_ll, b2_ll):

    """
    Calculates the closest distance between two project geometries.

    Projects can be:
        - a single point
        - a line segment

    For short distances, the coordinates are projected onto a local
    kilometer grid and standard 2D geometry is used.
    """

    # Average latitude of the four points
    ref_lat = (
        a1_ll[0]
        + a2_ll[0]
        + b1_ll[0]
        + b2_ll[0]
    ) / 4.0

    # Correct longitude scale at this latitude
    km_per_deg_lon = (
        111.320
        * math.cos(
            math.radians(ref_lat)
        )
    )

    def project(point):

        lat, lon = point

        return (
            lon * km_per_deg_lon,
            lat * KM_PER_DEG_LAT
        )

    a1 = project(a1_ll)
    a2 = project(a2_ll)

    b1 = project(b1_ll)
    b2 = project(b2_ll)

    a_is_point = _is_point(
        a1,
        a2
    )

    b_is_point = _is_point(
        b1,
        b2
    )

    # --------------------------------------------------------
    # Two line segments
    # --------------------------------------------------------

    if not a_is_point and not b_is_point:

        if _segments_intersect(
            a1,
            a2,
            b1,
            b2
        ):
            return 0.0

    # --------------------------------------------------------
    # Calculate candidate distances
    # --------------------------------------------------------

    candidates = [

        _dist(
            a1,
            _closest_point_on_segment(
                a1,
                b1,
                b2
            )
        ),

        _dist(
            b1,
            _closest_point_on_segment(
                b1,
                a1,
                a2
            )
        ),
    ]

    if not a_is_point:

        candidates.append(
            _dist(
                a2,
                _closest_point_on_segment(
                    a2,
                    b1,
                    b2
                )
            )
        )

    if not b_is_point:

        candidates.append(
            _dist(
                b2,
                _closest_point_on_segment(
                    b2,
                    a1,
                    a2
                )
            )
        )

    return min(candidates)



def geographic_tier(distance_km):
    """Returns the coordination-tier label for a distance, or None if beyond 40 km."""

    for max_km, label in TIERS:
        if distance_km <= max_km:
            return label

    return None


# ------------------------------------------------------------
# Stage 5: build the overlap table
# ------------------------------------------------------------

def build_overlap_table(desc_projects, gpc_projects):

    overlaps = []

    for d in desc_projects:

        for g in gpc_projects:

            distance = closest_distance_km(d["p1"], d["p2"], g["p1"], g["p2"])

            tier = geographic_tier(distance)

            if tier is None:
                continue

            if d["in_service_date"] and g["in_service_date"]:
                day_gap = abs((d["in_service_date"] - g["in_service_date"]).days)
                timeline_overlap = day_gap <= TIMELINE_WINDOW_DAYS
            else:
                day_gap = None
                timeline_overlap = None  # unknown, not "no"

            overlaps.append({
                "desc_id": d["id"],
                "desc_name": d["name"],
                "desc_in_service": d["in_service_date"],
                "gpc_id": g["id"],
                "gpc_name": g["name"],
                "gpc_in_service": g["in_service_date"],
                "distance_km": round(distance, 3),
                "distance_mi": round(distance / KM_PER_MILE, 2),
                "geographic_tier": tier,
                "day_gap": day_gap,
                "timeline_overlap": timeline_overlap,
            })

    # Primary signal is geographic closeness; sort by distance first.
    overlaps.sort(key=lambda r: r["distance_km"])

    return overlaps


# ------------------------------------------------------------
# Output
# ------------------------------------------------------------

def save_overlap_table(overlaps, path):

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Overlap Table"

    headers = [
        "DESC ID", "DESC Name", "DESC In-Service Date",
        "GPC ID", "GPC Name", "GPC In-Service Date",
        "Distance (km)", "Distance (mi)", "Geographic Tier",
        "Day Gap", "Timeline Overlap (<= {} days)".format(TIMELINE_WINDOW_DAYS),
    ]
    ws.append(headers)

    for row in overlaps:
        ws.append([
            row["desc_id"], row["desc_name"], str(row["desc_in_service"]) if row["desc_in_service"] else "",
            row["gpc_id"], row["gpc_name"], str(row["gpc_in_service"]) if row["gpc_in_service"] else "",
            row["distance_km"], row["distance_mi"], row["geographic_tier"],
            row["day_gap"] if row["day_gap"] is not None else "",
            "" if row["timeline_overlap"] is None else ("Yes" if row["timeline_overlap"] else "No"),
        ])

    for col_cells in ws.columns:
        max_len = max(len(str(c.value)) if c.value is not None else 0 for c in col_cells)
        ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 60)

    wb.save(path)


def main():

    wb = openpyxl.load_workbook(SOURCE_FILE, data_only=True)

    desc_projects = load_desc_projects(wb)
    gpc_projects = load_gpc_projects(wb)

    both_endpoints = sum(1 for p in desc_projects + gpc_projects if p["has_endpoint_2"])

    print(f"DESC projects: {len(desc_projects)}")
    print(f"GPC projects:  {len(gpc_projects)}")
    print(f"Projects with a real two-endpoint line segment: {both_endpoints}")
    print(f"Pairs checked: {len(desc_projects) * len(gpc_projects)}")

    overlaps = build_overlap_table(desc_projects, gpc_projects)

    print(f"Overlaps found (<= {OVERLAP_THRESHOLD_KM} km): {len(overlaps)}")
    print()

    header = (
        f"{'DESC ID':<12} {'GPC ID':<10} {'Dist (km)':>10} {'Timeline':>9}  Tier"
    )
    print(header)
    print("-" * 100)

    for row in overlaps:
        timeline = (
            "n/a" if row["timeline_overlap"] is None
            else ("YES" if row["timeline_overlap"] else "no")
        )
        print(
            f"{str(row['desc_id']):<12} {str(row['gpc_id']):<10} "
            f"{row['distance_km']:>10} {timeline:>9}  {row['geographic_tier']}"
        )

    save_overlap_table(overlaps, OUTPUT_XLSX)
    print()
    print(f"Saved to: {OUTPUT_XLSX}")

    return overlaps


if __name__ == "__main__":
    main()
