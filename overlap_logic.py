from datetime import date, datetime
from math import asin, cos, radians, sin, sqrt


DISTANCE_TIERS = (
    (1.6, "Under 1.6 km -- can share the right-of-way (access roads, permits)", "orange"),
    (8.0, "Under 8 km -- can share site logistics (laydown yards, deliveries)", "purple"),
    (40.0, "Under 40 km -- can share crews & equipment", "blue"),
)

OVERLAP_COLORS = {
    "Touching / Crossing -- must coordinate (outage timing, crossing structures)": (0.95, 0.15, 0.12),
    "Under 1.6 km -- can share the right-of-way (access roads, permits)": (1.0, 0.55, 0.05),
    "Under 8 km -- can share site logistics (laydown yards, deliveries)": (0.72, 0.2, 0.9),
    "Under 40 km -- can share crews & equipment": (0.15, 0.45, 1.0),
}


def parse_project_date(value):
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    for format_string in ("%m/%d/%Y", "%Y-%m-%d", "%m/%d/%y"):
        try:
            return datetime.strptime(str(value).strip(), format_string).date()
        except ValueError:
            continue
    return None


def distance_km(first, second):
    earth_radius_km = 6371.0088
    lat1, lat2 = radians(first["lat"]), radians(second["lat"])
    delta_lat = lat2 - lat1
    delta_lon = radians(second["lon"] - first["lon"])
    haversine = sin(delta_lat / 2) ** 2 + cos(lat1) * cos(lat2) * sin(delta_lon / 2) ** 2
    return earth_radius_km * 2 * asin(sqrt(haversine))


def timeline_info(first, second):
    first_date = parse_project_date(first.get("details", {}).get(first["date_field"]))
    second_date = parse_project_date(second.get("details", {}).get(second["date_field"]))
    if first_date is None or second_date is None:
        return None, None
    day_gap = abs((first_date - second_date).days)
    return day_gap <= 365, day_gap


def classify_distance(distance):
    for threshold, tier, color_name in DISTANCE_TIERS:
        if distance < threshold:
            return tier, color_name
    if distance < 40.0:
        return DISTANCE_TIERS[-1][1], DISTANCE_TIERS[-1][2]
    return None, None


def build_project_overlaps(projects):
    """Build DESC-to-GA ITS relationships using the mockup's distance tiers."""
    desc_projects = [project for project in projects if project["sheet"] == "DESC Geocoded"]
    ga_projects = [project for project in projects if project["sheet"] == "GA ITS Geocoded"]
    overlaps = []

    for desc_project in desc_projects:
        for ga_project in ga_projects:
            distance = distance_km(desc_project, ga_project)
            tier, color_name = classify_distance(distance)
            if tier is None:
                continue

            timeline_overlap, day_gap = timeline_info(desc_project, ga_project)
            overlap = {
                "first": desc_project,
                "second": ga_project,
                "distance_km": round(distance, 2),
                "distance_mi": round(distance * 0.621371, 2),
                "geographic_tier": tier,
                "color": OVERLAP_COLORS[tier],
                "timeline_overlap": timeline_overlap,
                "day_gap": day_gap,
            }
            overlaps.append(overlap)
            desc_project.setdefault("overlaps", []).append(overlap)
            ga_project.setdefault("overlaps", []).append(overlap)

    return overlaps
