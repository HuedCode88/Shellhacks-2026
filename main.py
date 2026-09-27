"""Interactive 3D OpenStreetMap floor with geocoded project markers."""

import math
import os
import argparse
import threading
from concurrent.futures import ThreadPoolExecutor
from io import BytesIO

import gridfs
import openpyxl
import pygame
import requests
from dotenv import load_dotenv
from OpenGL.GL import *
from OpenGL.GLU import *
from PIL import Image
from pygame.locals import DOUBLEBUF, FULLSCREEN, OPENGL
from pymongo import MongoClient


from overlap_logic import build_project_overlaps
from project_ui import ProjectInfoPanel


# ============================================================
# CONFIG
# ============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

EXCEL_FILE = os.path.join(
    BASE_DIR,
    "gridlock_project_tables_geocoded.xlsx",
)

WINDOW_SIZE = (1200, 800)

TILE_ZOOM = 8

MAX_DETAIL_ZOOM = 17
MIN_DETAIL_ZOOM = 2

TILE_SIZE = 256
WORLD_UNITS_PER_TILE = 60.0
TILE_CACHE_DIR = os.path.join(BASE_DIR, "map_tiles")

MAX_TILE_GRID = 8

FOCUSED_TILE_SPAN = 6

MONGO_TILE_WORKERS = 12

MAX_MEMORY_TILES = 300

# When focusing on a project (search result, ranking click, etc.),
# always teleport straight to this zoom level rather than only
# zooming in if the current zoom is shallower.
TELEPORT_ZOOM = 10
INITIAL_MAP_FOCUS = (30.0, -82.0)  # slightly north toward the Georgia border

# Used only as a fallback when MongoDB doesn't have a tile yet.
OSM_USER_AGENT = "Shellhacks-2026-3D-map/1.0 (local visualization)"


# ============================================================
# MONGODB
# ============================================================

load_dotenv()

MONGO_DB_NAME = "map"
MONGO_BUCKET_NAME = "images"

MONGO_CLIENT = MongoClient(
    os.getenv("MONGODB_URI"),
    maxPoolSize=MONGO_TILE_WORKERS + 4,
)

MONGO_DB = MONGO_CLIENT[MONGO_DB_NAME]

TILE_FS = gridfs.GridFS(
    MONGO_DB,
    collection=MONGO_BUCKET_NAME,
)

# GridFS doesn't index metadata fields by default, so without this,
# every tile lookup below collection-scans images.files. create_index
# is a cheap no-op if the index already exists.
MONGO_DB[f"{MONGO_BUCKET_NAME}.files"].create_index(
    [("metadata.z", 1), ("metadata.x", 1), ("metadata.y", 1)]
)


# ============================================================
# IN-MEMORY TILE CACHE
# ============================================================

TILE_MEMORY_CACHE = {}

TILE_CACHE_LOCK = threading.Lock()


def get_cached_tile(cache_key):

    with TILE_CACHE_LOCK:
        return TILE_MEMORY_CACHE.get(cache_key)


def cache_tile(cache_key, image):

    with TILE_CACHE_LOCK:

        TILE_MEMORY_CACHE[cache_key] = image

        while (
            len(TILE_MEMORY_CACHE)
            > MAX_MEMORY_TILES
        ):

            oldest_key = next(
                iter(TILE_MEMORY_CACHE)
            )

            del TILE_MEMORY_CACHE[
                oldest_key
            ]


# ============================================================
# PROJECT DATA
# ============================================================

SHEET_CATEGORY_FIELD = {
    "DESC Geocoded": "Status",
    "GA ITS Geocoded": "Sponsor (GPC/GTC/MEAG/DU/SAV)",
}


CATEGORY_COLORS = {
    ("DESC Geocoded", "In Progress"): (0.1, 0.85, 0.3),
    ("DESC Geocoded", "Planned"): (1.0, 0.55, 0.05),
    ("GA ITS Geocoded", "GPC"): (0.15, 0.45, 1.0),
    ("GA ITS Geocoded", "GTC"): (0.75, 0.2, 1.0),
    ("GA ITS Geocoded", "SAV"): (1.0, 0.3, 0.65),
    ("GA ITS Geocoded", "MEAG"): (0.1, 0.75, 0.8),
    ("GA ITS Geocoded", "DU"): (0.9, 0.2, 0.15),
}


def safe_float(value):

    if value is None:
        return None

    try:
        return float(value)

    except (TypeError, ValueError):
        return None


def load_projects(path):

    workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)

    projects = []

    for sheet_name, category_field in SHEET_CATEGORY_FIELD.items():

        if sheet_name not in workbook.sheetnames:
            continue

        rows = workbook[sheet_name].iter_rows(values_only=True)

        header = next(rows, None)

        if not header:
            continue

        for row in rows:

            record = dict(zip(header, row))

            latitude = safe_float(record.get("center_lat"))
            longitude = safe_float(record.get("center_lon"))

            if latitude is None or longitude is None:
                continue

            category = record.get(category_field)

            project_id = record.get("Project ID") or record.get("TEAMS Number")

            detail_fields = (
                "Status",
                "Description",
                "Need / Area Note",
                "Planned In-Service Date",
                "Need / In-Service Date",
                "Sponsor (GPC/GTC/MEAG/DU/SAV)",
                "Zone",
                "Year Filed",
                "endpoints_tried",
                "matched_endpoint_1",
                "matched_endpoint_2",
                "osm_name_1",
                "osm_name_2",
                "score_1",
                "score_2",
            )

            details = {
                field: record[field]
                for field in detail_fields
                if record.get(field) is not None and str(record[field]).strip()
            }

            projects.append({
                "sheet": sheet_name,
                "id": project_id,
                "name": record.get("Project Name / Endpoints (raw title)") or "Unnamed project",
                "category": category or "Uncategorized",
                "lat": latitude,
                "lon": longitude,
                "color": CATEGORY_COLORS.get((sheet_name, category), (0.95, 0.9, 0.1)),
                "details": details,
                "date_field": (
                    "Planned In-Service Date"
                    if sheet_name == "DESC Geocoded"
                    else "Need / In-Service Date"
                ),
            })

    if not projects:
        raise ValueError("No geocoded projects with center_lat/center_lon were found.")

    build_project_overlaps(projects)

    return projects


# ============================================================
# TILE / GEOGRAPHY CONVERSION
# ============================================================

def lon_to_tile_x(longitude, zoom):
    return (longitude + 180.0) / 360.0 * (2 ** zoom)


def lat_to_tile_y(latitude, zoom):
    latitude = max(-85.05112878, min(85.05112878, latitude))
    radians = math.radians(latitude)
    return (1.0 - (math.asinh(math.tan(radians)) / math.pi)) / 2.0 * (2 ** zoom)


def tile_x_to_lon(tile_x, zoom):
    return tile_x / (2 ** zoom) * 360.0 - 180.0


def tile_y_to_lat(tile_y, zoom):
    value = math.pi * (1.0 - (2.0 * tile_y / (2 ** zoom)))
    return math.degrees(math.atan(math.sinh(value)))


def choose_tile_zoom(projects):

    min_lat = min(project["lat"] for project in projects)
    max_lat = max(project["lat"] for project in projects)
    min_lon = min(project["lon"] for project in projects)
    max_lon = max(project["lon"] for project in projects)

    for zoom in range(8, 1, -1):

        tile_width = (
            math.floor(lon_to_tile_x(max_lon, zoom))
            - math.floor(lon_to_tile_x(min_lon, zoom))
            + 3
        )

        tile_height = (
            math.floor(lat_to_tile_y(min_lat, zoom))
            - math.floor(lat_to_tile_y(max_lat, zoom))
            + 3
        )

        if tile_width <= MAX_TILE_GRID and tile_height <= MAX_TILE_GRID:
            return zoom

    return 2


# ============================================================
# MONGODB TILE LOADING (with OSM fallback + write-back)
# ============================================================

def save_tile_to_disk(zoom, tile_x, tile_y, image_bytes):
    """Mirror every tile fetched by the viewer into the local PNG cache."""
    tile_path = os.path.join(
        TILE_CACHE_DIR,
        str(zoom),
        str(tile_x),
        f"{tile_y}.png",
    )
    try:
        os.makedirs(os.path.dirname(tile_path), exist_ok=True)
        temporary_path = f"{tile_path}.tmp"
        with open(temporary_path, "wb") as tile_file:
            tile_file.write(image_bytes)
        os.replace(temporary_path, tile_path)
    except OSError as error:
        print(f"Could not save local tile z={zoom} x={tile_x} y={tile_y}: {error}")


def load_tile_from_disk(zoom, tile_x, tile_y):
    """Read a locally cached PNG tile, if one exists."""
    tile_path = os.path.join(
        TILE_CACHE_DIR,
        str(zoom),
        str(tile_x),
        f"{tile_y}.png",
    )
    try:
        with open(tile_path, "rb") as tile_file:
            image_bytes = tile_file.read()
        return image_bytes
    except OSError:
        return None

def download_tile_from_osm(zoom, tile_x, tile_y):
    """
    Fetches one tile from the public OSM tile server. Used only when
    MongoDB doesn't have this tile yet. Returns raw PNG bytes, or
    None if the download failed.
    """
    url = f"https://tile.openstreetmap.org/{zoom}/{tile_x}/{tile_y}.png"

    try:
        response = requests.get(
            url,
            headers={"User-Agent": OSM_USER_AGENT},
            timeout=15,
        )
        response.raise_for_status()
        return response.content
    except requests.RequestException as error:
        print(f"OSM fallback download failed z={zoom} x={tile_x} y={tile_y}: {error}")
        return None


def upload_tile_to_mongo(zoom, tile_x, tile_y, image_bytes):
    """
    Write-through cache: saves a freshly OSM-fetched tile into GridFS
    so this exact tile never needs OSM again -- everything ends up
    living in MongoDB over time, with no physical map_tiles folder.
    """
    try:
        existing = MONGO_DB[f"{MONGO_BUCKET_NAME}.files"].find_one(
            {"metadata.z": zoom, "metadata.x": tile_x, "metadata.y": tile_y},
            {"_id": 1},
        )
        if existing is not None:
            return
        TILE_FS.put(
            image_bytes,
            filename=f"{zoom}_{tile_x}_{tile_y}.png",
            contentType="image/png",
            metadata={
                "z": zoom,
                "x": tile_x,
                "y": tile_y,
                "source": "osm_fallback",
            },
        )
    except Exception as error:
        print(f"Could not save tile z={zoom} x={tile_x} y={tile_y} to MongoDB: {error}")


def fetch_tiles_with_fallback(zoom, min_x, max_x, min_y, max_y):
    """
    Ensures every tile in this z/x/y rectangle is in the in-memory
    cache. MongoDB is checked first with ONE batched query for the
    whole rectangle (fast, indexed). Any tile MongoDB doesn't have
    falls back to a direct OSM download, and that result is written
    back into MongoDB so the gap doesn't need filling again.
    """

    needed = [
        (tile_x, tile_y)
        for tile_x in range(min_x, max_x + 1)
        for tile_y in range(min_y, max_y + 1)
    ]

    uncached = [
        coordinates for coordinates in needed
        if get_cached_tile((zoom, *coordinates)) is None
    ]

    if not uncached:
        return

    try:
        metadata_cursor = TILE_FS.find({
            "metadata.z": zoom,
            "metadata.x": {"$gte": min_x, "$lte": max_x},
            "metadata.y": {"$gte": min_y, "$lte": max_y},
        })

        grid_out_by_coordinate = {
            (grid_out.metadata["x"], grid_out.metadata["y"]): grid_out
            for grid_out in metadata_cursor
        }
    except Exception as error:
        print(f"MongoDB tile lookup failed; using local/OSM fallback: {error}")
        grid_out_by_coordinate = {}

    def read_tile(coordinates):
        tile_x, tile_y = coordinates
        grid_out = grid_out_by_coordinate.get(coordinates)

        if grid_out is not None:
            try:
                image_bytes = grid_out.read()
                save_tile_to_disk(zoom, tile_x, tile_y, image_bytes)
                image = Image.open(BytesIO(image_bytes)).convert("RGB")
                cache_tile((zoom, *coordinates), image)
                return
            except Exception as error:
                print(
                    "MongoDB tile decode error "
                    f"z={zoom} x={tile_x} y={tile_y}: {error} -- falling back to OSM"
                )

        # MongoDB missed the tile. Use the local map_tiles cache before
        # contacting OpenStreetMap, preserving the original fast fallback.
        local_bytes = load_tile_from_disk(zoom, tile_x, tile_y)
        if local_bytes is not None:
            try:
                image = Image.open(BytesIO(local_bytes)).convert("RGB")
                cache_tile((zoom, *coordinates), image)
                upload_tile_to_mongo(zoom, tile_x, tile_y, local_bytes)
                return
            except OSError as error:
                print(
                    "Local tile decode error "
                    f"z={zoom} x={tile_x} y={tile_y}: {error} -- falling back to OSM"
                )

        # Not in MongoDB (or unreadable) -- fall back to OSM, then
        # write it back so it's covered by MongoDB from now on.
        image_bytes = download_tile_from_osm(zoom, tile_x, tile_y)

        if image_bytes is None:
            return

        try:
            image = Image.open(BytesIO(image_bytes)).convert("RGB")
        except OSError as error:
            print(f"Could not decode OSM tile z={zoom} x={tile_x} y={tile_y}: {error}")
            return

        cache_tile((zoom, *coordinates), image)
        save_tile_to_disk(zoom, tile_x, tile_y, image_bytes)
        upload_tile_to_mongo(zoom, tile_x, tile_y, image_bytes)

    with ThreadPoolExecutor(max_workers=MONGO_TILE_WORKERS) as executor:
        list(executor.map(read_tile, uncached))


def calculate_tile_bounds(projects, zoom, focus=None):

    tile_count = 2 ** zoom

    if focus is None:

        min_lat = min(project["lat"] for project in projects)
        max_lat = max(project["lat"] for project in projects)
        min_lon = min(project["lon"] for project in projects)
        max_lon = max(project["lon"] for project in projects)

        min_x = max(0, math.floor(lon_to_tile_x(min_lon, zoom)) - 1)
        max_x = min(tile_count - 1, math.floor(lon_to_tile_x(max_lon, zoom)) + 1)
        min_y = max(0, math.floor(lat_to_tile_y(max_lat, zoom)) - 1)
        max_y = min(tile_count - 1, math.floor(lat_to_tile_y(min_lat, zoom)) + 1)

    else:

        focus_lat, focus_lon = focus

        focus_x = math.floor(lon_to_tile_x(focus_lon, zoom))
        focus_y = math.floor(lat_to_tile_y(focus_lat, zoom))

        half_span = FOCUSED_TILE_SPAN // 2

        min_x = max(0, focus_x - half_span)
        max_x = min(tile_count - 1, min_x + FOCUSED_TILE_SPAN - 1)
        min_y = max(0, focus_y - half_span)
        max_y = min(tile_count - 1, min_y + FOCUSED_TILE_SPAN - 1)

    return min_x, max_x, min_y, max_y


def build_texture_image(projects, zoom, focus=None):

    min_x, max_x, min_y, max_y = calculate_tile_bounds(projects, zoom, focus)

    width_tiles = max_x - min_x + 1
    height_tiles = max_y - min_y + 1

    texture_image = Image.new(
        "RGB",
        (width_tiles * TILE_SIZE, height_tiles * TILE_SIZE),
        (45, 52, 58),
    )

    coordinates = [
        (tile_x, tile_y)
        for tile_x in range(min_x, max_x + 1)
        for tile_y in range(min_y, max_y + 1)
    ]

    fetch_tiles_with_fallback(zoom, min_x, max_x, min_y, max_y)

    for tile_x, tile_y in coordinates:

        tile = get_cached_tile((zoom, tile_x, tile_y))

        if tile is None:
            continue

        texture_image.paste(
            tile,
            ((tile_x - min_x) * TILE_SIZE, (tile_y - min_y) * TILE_SIZE),
        )

    center_x = (min_x + max_x + 1) / 2.0
    center_y = (min_y + max_y + 1) / 2.0

    return texture_image, min_x, min_y, center_x, center_y


# ============================================================
# ASYNCHRONOUS MAP LOADER
# ============================================================

class AsyncMapLoader:

    def __init__(self, projects):

        self.projects = projects
        self.executor = ThreadPoolExecutor(max_workers=1)
        self.lock = threading.Lock()
        self.future = None
        self.future_generation = 0
        self.future_zoom = None
        self.request_generation = 0
        self.pending_request = None

    def request(self, zoom, focus=None):

        with self.lock:

            self.request_generation += 1
            generation = self.request_generation

            # Always retain only the newest request.
            self.pending_request = (generation, zoom, focus)

            if self.future is None:
                self._start_pending_locked()

    def _start_pending_locked(self):

        if self.future is not None:
            return

        if self.pending_request is None:
            return

        generation, zoom, focus = self.pending_request
        self.pending_request = None

        self.future_generation = generation
        self.future_zoom = zoom

        self.future = self.executor.submit(
            build_texture_image,
            self.projects,
            zoom,
            focus,
        )

    def poll(self):

        with self.lock:

            if self.future is None:
                return None

            if not self.future.done():
                return None

            future = self.future
            generation = self.future_generation
            zoom = self.future_zoom

            self.future = None
            self.future_zoom = None

            try:
                result = future.result()
            except Exception as error:
                print(f"Background MongoDB map load failed: {error}")
                self._start_pending_locked()
                return None

            # Ignore stale results.
            if generation != self.request_generation:
                self._start_pending_locked()
                return None

            self._start_pending_locked()

            return zoom, result

    def close(self):
        self.executor.shutdown(wait=False, cancel_futures=True)


# ============================================================
# OPENGL TEXTURE
# ============================================================

def make_texture(image):

    texture_id = glGenTextures(1)
    glBindTexture(GL_TEXTURE_2D, texture_id)
    glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR)
    glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR)
    glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_WRAP_S, GL_CLAMP_TO_EDGE)
    glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_WRAP_T, GL_CLAMP_TO_EDGE)
    glTexImage2D(
        GL_TEXTURE_2D, 0, GL_RGB, image.width, image.height, 0,
        GL_RGB, GL_UNSIGNED_BYTE, image.tobytes(),
    )
    glBindTexture(GL_TEXTURE_2D, 0)

    return texture_id


# ============================================================
# OPENGL SETUP
# ============================================================

def setup_opengl(width, height):

    glClearColor(0.035, 0.05, 0.07, 1.0)
    glViewport(0, 0, width, height)
    glMatrixMode(GL_PROJECTION)
    glLoadIdentity()
    gluPerspective(55.0, width / height, 0.1, 1000.0)
    glMatrixMode(GL_MODELVIEW)
    glEnable(GL_DEPTH_TEST)
    glEnable(GL_BLEND)
    glBlendFunc(GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA)


# ============================================================
# MAP FLOOR
# ============================================================

def draw_floor(texture_id, image, min_x, min_y, center_x, center_y):

    world_width = image.width / TILE_SIZE * WORLD_UNITS_PER_TILE
    world_depth = image.height / TILE_SIZE * WORLD_UNITS_PER_TILE
    floor_x = (min_x - center_x) * WORLD_UNITS_PER_TILE
    floor_z = (min_y - center_y) * WORLD_UNITS_PER_TILE

    glEnable(GL_TEXTURE_2D)
    glBindTexture(GL_TEXTURE_2D, texture_id)
    glColor3f(1.0, 1.0, 1.0)
    glBegin(GL_QUADS)
    glTexCoord2f(0.0, 0.0)
    glVertex3f(floor_x, 0.0, floor_z)
    glTexCoord2f(1.0, 0.0)
    glVertex3f(floor_x + world_width, 0.0, floor_z)
    glTexCoord2f(1.0, 1.0)
    glVertex3f(floor_x + world_width, 0.0, floor_z + world_depth)
    glTexCoord2f(0.0, 1.0)
    glVertex3f(floor_x, 0.0, floor_z + world_depth)
    glEnd()
    glDisable(GL_TEXTURE_2D)


# ============================================================
# PROJECT POSITION
# ============================================================

def project_position(project, center_x, center_y, zoom=None):

    if zoom is None:
        zoom = TILE_ZOOM

    return (
        (lon_to_tile_x(project["lon"], zoom) - center_x) * WORLD_UNITS_PER_TILE,
        (lat_to_tile_y(project["lat"], zoom) - center_y) * WORLD_UNITS_PER_TILE,
    )


def project_is_visible(project, min_x, max_x, min_y, max_y, zoom=None):

    if zoom is None:
        zoom = TILE_ZOOM

    project_x = lon_to_tile_x(project["lon"], zoom)
    project_y = lat_to_tile_y(project["lat"], zoom)

    return min_x <= project_x <= max_x + 1 and min_y <= project_y <= max_y + 1


# ============================================================
# MOUSE -> MAP
# ============================================================

def map_point_under_mouse(mouse_position, center_x, center_y, zoom):

    viewport = glGetIntegerv(GL_VIEWPORT)
    mouse_x, mouse_y = mouse_position
    mouse_y = viewport[3] - mouse_y

    modelview = glGetDoublev(GL_MODELVIEW_MATRIX)
    projection = glGetDoublev(GL_PROJECTION_MATRIX)

    near_point = gluUnProject(mouse_x, mouse_y, 0.0, modelview, projection, viewport)
    far_point = gluUnProject(mouse_x, mouse_y, 1.0, modelview, projection, viewport)

    ray_y = far_point[1] - near_point[1]

    if abs(ray_y) < 1e-9:
        return None

    fraction = -near_point[1] / ray_y

    map_x = near_point[0] + fraction * (far_point[0] - near_point[0])
    map_z = near_point[2] + fraction * (far_point[2] - near_point[2])

    tile_x = center_x + map_x / WORLD_UNITS_PER_TILE
    tile_y = center_y + map_z / WORLD_UNITS_PER_TILE

    return map_x, map_z, tile_y_to_lat(tile_y, zoom), tile_x_to_lon(tile_x, zoom)


# ============================================================
# PROJECT MARKER
# ============================================================

def draw_marker(project, center_x, center_y, quadric, zoom, selected=False):

    x, z = project_position(project, center_x, center_y, zoom)
    red, green, blue = project["color"]
    marker_radius = 1.8 if selected else 1.2

    glColor3f(red, green, blue)
    glPushMatrix()
    glTranslatef(x, marker_radius, z)
    gluSphere(quadric, marker_radius, 12, 8)
    glPopMatrix()
    glBegin(GL_LINES)
    glVertex3f(x, 0.0, z)
    glVertex3f(x, marker_radius, z)
    glEnd()


# ============================================================
# PROJECT PICKING
# ============================================================

def pick_project(projects, click_position, center_x, center_y, zoom, show_desc=True, show_ga_its=True):

    viewport = glGetIntegerv(GL_VIEWPORT)
    window_height = viewport[3]
    click_x, click_y = click_position
    click_y = window_height - click_y

    closest_project = None
    closest_distance = 18.0

    for project in projects:

        if project["sheet"] == "DESC Geocoded" and not show_desc:
            continue
        if project["sheet"] == "GA ITS Geocoded" and not show_ga_its:
            continue

        x, z = project_position(project, center_x, center_y, zoom)
        screen_x, screen_y, depth = gluProject(x, 1.2, z)
        distance = math.hypot(screen_x - click_x, screen_y - click_y)

        if 0.0 <= depth <= 1.0 and distance < closest_distance:
            closest_project = project
            closest_distance = distance

    return closest_project


# ============================================================
# PROJECT DESCRIPTION
# ============================================================

def describe_project(project):
    return (
        f"{project['name']} | {project['sheet']} | "
        f"ID: {project['id'] or 'N/A'} | "
        f"Category: {project['category']} | "
        f"Lat: {project['lat']:.5f}, Lon: {project['lon']:.5f}"
    )


# ============================================================
# AXES
# ============================================================

def draw_axes(length):

    glLineWidth(2.0)
    glBegin(GL_LINES)
    glColor3f(1.0, 0.1, 0.1)
    glVertex3f(0, 0.02, 0)
    glVertex3f(length, 0.02, 0)
    glColor3f(0.1, 1.0, 0.1)
    glVertex3f(0, 0.02, 0)
    glVertex3f(0, length, 0)
    glColor3f(0.1, 0.4, 1.0)
    glVertex3f(0, 0.02, 0)
    glVertex3f(0, 0.02, length)
    glEnd()


# ============================================================
# OVERLAP LINES
# ============================================================

def draw_overlap_lines(overlaps, center_x, center_y, zoom, selected_project=None):

    glPushAttrib(GL_ENABLE_BIT | GL_COLOR_BUFFER_BIT | GL_LINE_BIT | GL_DEPTH_BUFFER_BIT)
    glDisable(GL_DEPTH_TEST)
    glDisable(GL_TEXTURE_2D)
    glEnable(GL_BLEND)
    glBlendFunc(GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA)

    rendered_lines = 0

    for overlap in overlaps:

        first_x, first_z = project_position(overlap["first"], center_x, center_y, zoom)
        second_x, second_z = project_position(overlap["second"], center_x, center_y, zoom)

        is_selected = selected_project in (overlap["first"], overlap["second"])

        red, green, blue = overlap["color"]
        line_width = 7.0 if is_selected else 4.0
        alpha = 0.95 if is_selected else 0.75

        glLineWidth(line_width)
        glColor4f(red, green, blue, alpha)
        glBegin(GL_LINES)
        glVertex3f(first_x, 0.12, first_z)
        glVertex3f(second_x, 0.12, second_z)
        glEnd()

        rendered_lines += 1

    glPopAttrib()

    return rendered_lines


def draw_overlap_domes(overlaps, center_x, center_y, zoom, selected_project=None):
    """Render translucent coordination domes at every ranked overlap site."""
    dome_radii_km = {
        "Touching / Crossing -- must coordinate (outage timing, crossing structures)": 0.4,
        "Under 1.6 km -- can share the right-of-way (access roads, permits)": 1.6,
        "Under 8 km -- can share site logistics (laydown yards, deliveries)": 8.0,
        "Under 40 km -- can share crews & equipment": 40.0,
    }
    sites = {}
    for overlap in overlaps:
        radius_km = dome_radii_km.get(overlap["geographic_tier"], 0.4)
        for project in (overlap["first"], overlap["second"]):
            key = id(project)
            previous = sites.get(key)
            if previous is None or radius_km > previous[1]:
                sites[key] = (project, radius_km, overlap["color"])

    glPushAttrib(GL_ENABLE_BIT | GL_COLOR_BUFFER_BIT | GL_DEPTH_BUFFER_BIT)
    glEnable(GL_BLEND)
    glBlendFunc(GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA)
    glDepthMask(GL_FALSE)
    glDisable(GL_TEXTURE_2D)
    for project, radius_km, color in sites.values():
        x, z = project_position(project, center_x, center_y, zoom)
        latitude_scale = math.cos(math.radians(project["lat"]))
        kilometers_per_tile = (
            156543.03392 * latitude_scale / (2 ** zoom) * TILE_SIZE / 1000.0
        )
        radius = radius_km / kilometers_per_tile * WORLD_UNITS_PER_TILE
        glColor4f(color[0], color[1], color[2], 0.28)
        for ring in range(8):
            phi0 = (math.pi / 2.0) * ring / 8.0
            phi1 = (math.pi / 2.0) * (ring + 1) / 8.0
            glBegin(GL_TRIANGLE_STRIP)
            for segment in range(25):
                theta = 2.0 * math.pi * segment / 24.0
                glVertex3f(x + radius * math.cos(phi0) * math.cos(theta), 0.08 + radius * math.sin(phi0), z + radius * math.cos(phi0) * math.sin(theta))
                glVertex3f(x + radius * math.cos(phi1) * math.cos(theta), 0.08 + radius * math.sin(phi1), z + radius * math.cos(phi1) * math.sin(theta))
            glEnd()
    glDepthMask(GL_TRUE)
    glPopAttrib()


# ============================================================
# CACHE VISIBLE PROJECTS
# ============================================================

def calculate_visible_projects(projects, min_x, max_x, min_y, max_y, zoom):
    return [
        project
        for project in projects
        if project_is_visible(project, min_x, max_x, min_y, max_y, zoom)
    ]


def calculate_visible_overlaps(overlaps, min_x, max_x, min_y, max_y, zoom):

    visible = []

    for overlap in overlaps:

        if not project_is_visible(overlap["first"], min_x, max_x, min_y, max_y, zoom):
            continue

        if project_is_visible(overlap["second"], min_x, max_x, min_y, max_y, zoom):
            visible.append(overlap)

    return visible


# ============================================================
# MAIN
# ============================================================

def main(preload=False):

    global TILE_ZOOM

    projects = load_projects(EXCEL_FILE)

    if preload:

        # Sanity check only. Loads one map at a time.
        preload_zoom = choose_tile_zoom(projects)
        print(f"Preloading zoom {preload_zoom}...")
        build_texture_image(projects, preload_zoom)
        print("Preload complete.")
        return

    overlaps = [
        overlap
        for project in projects
        for overlap in project.get("overlaps", [])
        if overlap["first"] is project
    ]

    print(f"Loaded {len(projects)} geocoded projects.")
    print(f"Found {len(overlaps)} geographic project overlaps.")

    # --------------------------------------------------------
    # Initial map
    # --------------------------------------------------------

    map_zoom = choose_tile_zoom(projects)
    TILE_ZOOM = map_zoom

    print("Loading initial map floor tiles (MongoDB first, OSM fallback)...")

    map_image, min_x, min_y, center_x, center_y = build_texture_image(
        projects,
        map_zoom,
        focus=INITIAL_MAP_FOCUS,
    )

    max_x = min_x + map_image.width // TILE_SIZE - 1
    max_y = min_y + map_image.height // TILE_SIZE - 1

    # --------------------------------------------------------
    # Cached visibility
    # --------------------------------------------------------

    visible_projects = calculate_visible_projects(projects, min_x, max_x, min_y, max_y, map_zoom)
    visible_overlaps = calculate_visible_overlaps(overlaps, min_x, max_x, min_y, max_y, map_zoom)

    # --------------------------------------------------------
    # Pygame / OpenGL
    # --------------------------------------------------------

    pygame.init()
    pygame.display.set_mode((0, 0), DOUBLEBUF | FULLSCREEN | OPENGL)
    window_size = pygame.display.get_window_size()
    pygame.display.set_caption("3D Gridlock Project Map | © OpenStreetMap contributors")
    setup_opengl(*window_size)
    texture_id = make_texture(map_image)
    quadric = gluNewQuadric()
    info_panel = ProjectInfoPanel(window_size)
    info_panel.set_projects(projects)
    info_panel.set_overlaps(overlaps)

    # --------------------------------------------------------
    # Async loader
    # --------------------------------------------------------

    map_loader = AsyncMapLoader(projects)

    # --------------------------------------------------------
    # Camera
    # --------------------------------------------------------

    camera_distance = max(240.0, max(map_image.width, map_image.height) / TILE_SIZE * 20.0)
    camera_rotation_x = 90.0
    camera_rotation_y = 0.0
    camera_offset_x = 0.0
    camera_offset_z = 0.0
    pending_focus = None
    mouse_down = False
    mouse_down_position = (0, 0)
    last_mouse_pos = (0, 0)
    selected_project = None
    pending_click = None
    pending_double_click = False
    last_click_time = 0
    last_click_position = (0, 0)
    show_desc = True
    show_ga_its = True
    show_overlaps = True
    show_domes = True
    clock = pygame.time.Clock()
    running = True

    def refresh_visible_cache():
        nonlocal visible_projects, visible_overlaps
        visible_projects = calculate_visible_projects(projects, min_x, max_x, min_y, max_y, map_zoom)
        visible_overlaps = calculate_visible_overlaps(overlaps, min_x, max_x, min_y, max_y, map_zoom)

    def focus_camera_on_project(project):
        nonlocal camera_distance, camera_rotation_x, camera_rotation_y
        nonlocal camera_offset_x, camera_offset_z, pending_focus

        target_zoom = min(MAX_DETAIL_ZOOM, TELEPORT_ZOOM)

        map_loader.request(target_zoom, focus=(project["lat"], project["lon"]))
        pending_focus = (None, None, project["lat"], project["lon"])

        camera_rotation_x = 90.0
        camera_rotation_y = 0.0
        camera_distance = 110.0

    while running:

        for event in pygame.event.get():

            if event.type == pygame.QUIT:
                running = False

            elif event.type == pygame.KEYDOWN:
                if event.key == pygame.K_ESCAPE:
                    if info_panel.details_expanded:
                        info_panel.collapse_details()
                    elif info_panel.search_active:
                        info_panel.close_search()
                    else:
                        running = False
                elif event.key == pygame.K_r:
                    info_panel.toggle_ranking()
                elif event.key == pygame.K_1:
                    show_desc = not show_desc
                    info_panel.set_layers(show_desc, show_ga_its, show_overlaps, show_domes)
                elif event.key == pygame.K_2:
                    show_ga_its = not show_ga_its
                    info_panel.set_layers(show_desc, show_ga_its, show_overlaps, show_domes)
                elif event.key == pygame.K_3:
                    show_overlaps = not show_overlaps
                    info_panel.set_layers(show_desc, show_ga_its, show_overlaps, show_domes)
                elif event.key == pygame.K_4:
                    show_domes = not show_domes
                    info_panel.set_layers(show_desc, show_ga_its, show_overlaps, show_domes)
                elif event.key == pygame.K_l:
                    info_panel.toggle_panel()
                elif info_panel.search_active and event.key == pygame.K_BACKSPACE:
                    info_panel.remove_search_character()
                elif info_panel.search_active and event.key == pygame.K_RETURN:
                    search_result = info_panel.choose_search_result()
                    if search_result is not None:
                        selected_project = search_result
                        focus_camera_on_project(selected_project)
                        info_panel.set_project(selected_project)
                        pygame.display.set_caption(describe_project(selected_project))

            elif event.type == pygame.TEXTINPUT and info_panel.search_active:
                info_panel.append_search_text(event.text)

            elif event.type == pygame.MOUSEBUTTONDOWN:

                if event.button == 1 and info_panel.details_close_hit(event.pos):
                    info_panel.collapse_details()
                    continue

                if event.button == 1 and info_panel.panel_button_hit(event.pos):
                    info_panel.toggle_panel()
                    continue

                if event.button == 1 and info_panel.layer_button_hit(event.pos):
                    info_panel.toggle_layer_panel()
                    continue

                if event.button == 1:
                    layer = info_panel.layer_at(event.pos)
                    if layer is not None:
                        if layer == "desc":
                            show_desc = not show_desc
                        elif layer == "ga_its":
                            show_ga_its = not show_ga_its
                        else:
                            if layer == "overlaps":
                                show_overlaps = not show_overlaps
                            else:
                                show_domes = not show_domes
                        info_panel.set_layers(show_desc, show_ga_its, show_overlaps, show_domes)
                        continue

                search_result, search_handled = info_panel.search_result_at(event.pos)
                if search_result is not None:
                    selected_project = search_result
                    focus_camera_on_project(selected_project)
                    info_panel.set_project(selected_project)
                    pygame.display.set_caption(describe_project(selected_project))
                    continue
                if search_handled:
                    continue

                if event.button == 1:
                    selected_overlap = info_panel.ranking_result_at(event.pos)
                    if selected_overlap is not None:
                        selected_project = selected_overlap["first"]
                        focus_camera_on_project(selected_project)
                        info_panel.set_project(selected_project)
                        pygame.display.set_caption(describe_project(selected_project))
                        continue

                if event.button in (4, 5) and info_panel.ranking_active and info_panel.ranking_wheel_hit(event.pos):
                    info_panel.scroll_ranking(-1 if event.button == 4 else 1)
                    continue

                if event.button == 1 and info_panel.ranking_active and info_panel.ranking_button_hit(event.pos):
                    info_panel.toggle_ranking_minimized()
                    continue

                if event.button == 1:
                    mouse_down = True
                    mouse_down_position = event.pos
                    last_mouse_pos = event.pos

                elif event.button == 4:
                    new_zoom = min(MAX_DETAIL_ZOOM, map_zoom + 1)
                    if new_zoom != map_zoom:
                        mouse_focus = map_point_under_mouse(event.pos, center_x, center_y, map_zoom)
                        if mouse_focus is None:
                            mouse_focus = (
                                0.0, 0.0,
                                tile_y_to_lat(center_y, map_zoom),
                                tile_x_to_lon(center_x, map_zoom),
                            )
                        old_map_x, old_map_z, focus_lat, focus_lon = mouse_focus
                        pending_focus = (old_map_x, old_map_z, focus_lat, focus_lon)
                        map_loader.request(new_zoom, focus=(focus_lat, focus_lon))

                elif event.button == 5:
                    new_zoom = max(MIN_DETAIL_ZOOM, map_zoom - 1)
                    if new_zoom != map_zoom:
                        mouse_focus = map_point_under_mouse(event.pos, center_x, center_y, map_zoom)
                        if mouse_focus is None:
                            mouse_focus = (
                                0.0, 0.0,
                                tile_y_to_lat(center_y, map_zoom),
                                tile_x_to_lon(center_x, map_zoom),
                            )
                        old_map_x, old_map_z, focus_lat, focus_lon = mouse_focus
                        pending_focus = (old_map_x, old_map_z, focus_lat, focus_lon)
                        map_loader.request(new_zoom, focus=(focus_lat, focus_lon))

            elif event.type == pygame.MOUSEBUTTONUP and event.button == 1:
                mouse_down = False
                if math.dist(event.pos, mouse_down_position) <= 6.0:
                    current_time = pygame.time.get_ticks()
                    pending_double_click = (
                        current_time - last_click_time <= 450
                        and math.dist(event.pos, last_click_position) <= 12.0
                    )
                    last_click_time = current_time
                    last_click_position = event.pos
                    pending_click = event.pos

            elif event.type == pygame.MOUSEMOTION:
                info_panel.update_pointer(event.pos)
                if mouse_down and not info_panel.search_active:
                    delta_x = event.pos[0] - last_mouse_pos[0]
                    delta_y = event.pos[1] - last_mouse_pos[1]
                    camera_rotation_y += delta_x * 0.5
                    camera_rotation_x = max(15.0, min(85.0, camera_rotation_x + delta_y * 0.5))
                    last_mouse_pos = event.pos

        # ====================================================
        # CHECK COMPLETED MAP LOAD
        # ====================================================

        loaded_map = map_loader.poll()

        if loaded_map is not None:

            loaded_zoom, result = loaded_map
            new_image, new_min_x, new_min_y, new_center_x, new_center_y = result

            new_texture_id = None
            try:
                new_texture_id = make_texture(new_image)
            except Exception as error:
                print(f"OpenGL texture creation failed: {error}")

            if new_texture_id is not None:

                old_texture = texture_id
                texture_id = new_texture_id
                if old_texture:
                    glDeleteTextures([old_texture])

                map_image = new_image
                min_x = new_min_x
                min_y = new_min_y
                center_x = new_center_x
                center_y = new_center_y
                map_zoom = loaded_zoom
                TILE_ZOOM = loaded_zoom

                max_x = min_x + map_image.width // TILE_SIZE - 1
                max_y = min_y + map_image.height // TILE_SIZE - 1

                refresh_visible_cache()

                if pending_focus is not None:

                    old_map_x, old_map_z, focus_lat, focus_lon = pending_focus

                    new_map_x = (lon_to_tile_x(focus_lon, map_zoom) - center_x) * WORLD_UNITS_PER_TILE
                    new_map_z = (lat_to_tile_y(focus_lat, map_zoom) - center_y) * WORLD_UNITS_PER_TILE

                    if old_map_x is not None:
                        camera_offset_x += new_map_x - old_map_x
                        camera_offset_z += new_map_z - old_map_z
                    else:
                        camera_offset_x = new_map_x
                        camera_offset_z = new_map_z

                    pending_focus = None

        # ====================================================
        # CAMERA MOVEMENT
        # ====================================================

        frame_seconds = clock.get_time() / 1000.0
        pressed_keys = pygame.key.get_pressed()
        movement_speed = 42.0 if (pressed_keys[pygame.K_LSHIFT] or pressed_keys[pygame.K_RSHIFT]) else 18.0
        movement = movement_speed * frame_seconds

        if not info_panel.search_active and pressed_keys[pygame.K_w]:
            camera_offset_z -= movement
        if not info_panel.search_active and pressed_keys[pygame.K_s]:
            camera_offset_z += movement
        if not info_panel.search_active and pressed_keys[pygame.K_a]:
            camera_offset_x -= movement
        if not info_panel.search_active and pressed_keys[pygame.K_d]:
            camera_offset_x += movement

        # ====================================================
        # OPENGL RENDER
        # ====================================================

        glClear(GL_COLOR_BUFFER_BIT | GL_DEPTH_BUFFER_BIT)
        glLoadIdentity()
        glTranslatef(0.0, -4.0, -camera_distance)
        glRotatef(camera_rotation_x, 1.0, 0.0, 0.0)
        glRotatef(camera_rotation_y, 0.0, 1.0, 0.0)
        glTranslatef(-camera_offset_x, 0.0, -camera_offset_z)

        if pending_click is not None:
            previously_selected_project = selected_project
            selected_project = pick_project(
                visible_projects, pending_click, center_x, center_y, map_zoom, show_desc, show_ga_its,
            )
            pending_click = None
            was_double_click = pending_double_click and selected_project is previously_selected_project
            pending_double_click = False
            info_panel.set_project(selected_project)
            if selected_project is None:
                pygame.display.set_caption("3D Gridlock Project Map | © OpenStreetMap contributors")
            else:
                if was_double_click:
                    info_panel.expand_details()
                project_info = describe_project(selected_project)
                print(f"Selected: {project_info}")
                pygame.display.set_caption(project_info)

        draw_floor(texture_id, map_image, min_x, min_y, center_x, center_y)

        if show_domes:
            draw_overlap_domes(visible_overlaps, center_x, center_y, map_zoom, selected_project)

        rendered_overlap_lines = (
            draw_overlap_lines(visible_overlaps, center_x, center_y, map_zoom, selected_project)
            if show_overlaps else 0
        )
        info_panel.set_overlap_status(len(overlaps), rendered_overlap_lines)

        draw_axes(12.0)

        for project in visible_projects:
            if project["sheet"] == "DESC Geocoded" and not show_desc:
                continue
            if project["sheet"] == "GA ITS Geocoded" and not show_ga_its:
                continue
            draw_marker(project, center_x, center_y, quadric, map_zoom, selected=project is selected_project)

        info_panel.draw()
        pygame.display.flip()
        clock.tick(60)

    # ========================================================
    # CLEANUP
    # ========================================================

    map_loader.close()
    gluDeleteQuadric(quadric)
    if texture_id:
        glDeleteTextures([texture_id])
    info_panel.close()
    pygame.quit()


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":

    parser = argparse.ArgumentParser(description="Run the 3D project map viewer.")
    parser.add_argument(
        "--preload-map",
        action="store_true",
        help="Fetch the initial map tiles from MongoDB (with OSM fallback), then exit.",
    )

    args = parser.parse_args()
    main(preload=args.preload_map)