"""Interactive 3D OpenStreetMap floor with geocoded project markers."""

import math
import os
import argparse
from concurrent.futures import ThreadPoolExecutor
from io import BytesIO

import openpyxl
import pygame
import requests
from OpenGL.GL import *
from OpenGL.GLU import *
from PIL import Image
from pygame.locals import DOUBLEBUF, FULLSCREEN, OPENGL
from overlap_logic import build_project_overlaps
from project_ui import ProjectInfoPanel


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, "gridlock_project_tables_geocoded_nominatim.xlsx")
WINDOW_SIZE = (1200, 800)
TILE_ZOOM = 8
MAX_DETAIL_ZOOM = 17
TILE_SIZE = 256
MAX_TILE_GRID = 8
FOCUSED_TILE_SPAN = 8
TELEPORT_ZOOM = 10
OSM_USER_AGENT = "Shellhacks-2026-3D-map/1.0 (local visualization)"
TILE_CACHE = {}
TILE_CACHE_DIR = os.path.join(BASE_DIR, "map_tiles")

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
                "confidence",
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
                "color": CATEGORY_COLORS.get(
                    (sheet_name, category),
                    (0.95, 0.9, 0.1),
                ),
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


def lon_to_tile_x(longitude, zoom):
    return (longitude + 180.0) / 360.0 * (2 ** zoom)


def lat_to_tile_y(latitude, zoom):
    latitude = max(-85.05112878, min(85.05112878, latitude))
    radians = math.radians(latitude)
    return (1.0 - math.asinh(math.tan(radians)) / math.pi) / 2.0 * (2 ** zoom)


def tile_x_to_lon(tile_x, zoom):
    return tile_x / (2 ** zoom) * 360.0 - 180.0


def tile_y_to_lat(tile_y, zoom):
    value = math.pi * (1.0 - 2.0 * tile_y / (2 ** zoom))
    return math.degrees(math.atan(math.sinh(value)))


def choose_tile_zoom(projects):
    """Choose the highest zoom that keeps every project in the tile mosaic."""
    min_lat = min(project["lat"] for project in projects)
    max_lat = max(project["lat"] for project in projects)
    min_lon = min(project["lon"] for project in projects)
    max_lon = max(project["lon"] for project in projects)

    for zoom in range(8, 1, -1):
        tile_width = math.floor(lon_to_tile_x(max_lon, zoom)) - math.floor(lon_to_tile_x(min_lon, zoom)) + 3
        tile_height = math.floor(lat_to_tile_y(min_lat, zoom)) - math.floor(lat_to_tile_y(max_lat, zoom)) + 3
        if tile_width <= MAX_TILE_GRID and tile_height <= MAX_TILE_GRID:
            return zoom

    return 2


def download_map_texture(projects, zoom=None, focus=None):
    """Download nearby OSM tiles and return a texture plus geographic bounds."""
    global TILE_ZOOM
    TILE_ZOOM = zoom if zoom is not None else choose_tile_zoom(projects)
    print(f"Using OpenStreetMap zoom level {TILE_ZOOM} for the current view.")

    min_lat = min(project["lat"] for project in projects)
    max_lat = max(project["lat"] for project in projects)
    min_lon = min(project["lon"] for project in projects)
    max_lon = max(project["lon"] for project in projects)

    if focus is None:
        center_lat = (min_lat + max_lat) / 2.0
        center_lon = (min_lon + max_lon) / 2.0
    else:
        center_lat, center_lon = focus
    center_x = lon_to_tile_x(center_lon, TILE_ZOOM)
    center_y = lat_to_tile_y(center_lat, TILE_ZOOM)
    tile_count = 2 ** TILE_ZOOM

    if zoom is None or TILE_ZOOM <= choose_tile_zoom(projects):
        min_x = max(0, math.floor(lon_to_tile_x(min_lon, TILE_ZOOM)) - 1)
        max_x = min(tile_count - 1, math.floor(lon_to_tile_x(max_lon, TILE_ZOOM)) + 1)
        min_y = max(0, math.floor(lat_to_tile_y(max_lat, TILE_ZOOM)) - 1)
        max_y = min(tile_count - 1, math.floor(lat_to_tile_y(min_lat, TILE_ZOOM)) + 1)
    else:
        tile_span = FOCUSED_TILE_SPAN
        focus_tile_x = math.floor(center_x)
        focus_tile_y = math.floor(center_y)
        min_x = max(0, focus_tile_x - tile_span // 2)
        max_x = min(tile_count - 1, min_x + tile_span - 1)
        min_y = max(0, focus_tile_y - tile_span // 2)
        max_y = min(tile_count - 1, min_y + tile_span - 1)

    center_x = (min_x + max_x + 1) / 2.0
    center_y = (min_y + max_y + 1) / 2.0

    texture_image = Image.new(
        "RGB",
        ((max_x - min_x + 1) * TILE_SIZE, (max_y - min_y + 1) * TILE_SIZE),
        (45, 52, 58),
    )
    def fetch_tile(tile_coordinates):
        tile_x, tile_y = tile_coordinates
        cache_key = (TILE_ZOOM, tile_x, tile_y)
        if cache_key in TILE_CACHE:
            return tile_coordinates, TILE_CACHE[cache_key]

        tile_path = os.path.join(
            TILE_CACHE_DIR,
            str(TILE_ZOOM),
            str(tile_x),
            f"{tile_y}.png",
        )
        try:
            with open(tile_path, "rb") as tile_file:
                tile = Image.open(tile_file).convert("RGB")
                TILE_CACHE[cache_key] = tile
                return tile_coordinates, tile
        except (FileNotFoundError, OSError):
            pass

        url = f"https://tile.openstreetmap.org/{TILE_ZOOM}/{tile_x}/{tile_y}.png"
        try:
            response = requests.get(
                url,
                headers={"User-Agent": OSM_USER_AGENT},
                timeout=15,
            )
            response.raise_for_status()
            os.makedirs(os.path.dirname(tile_path), exist_ok=True)
            with open(tile_path, "wb") as tile_file:
                tile_file.write(response.content)
            tile = Image.open(BytesIO(response.content)).convert("RGB")
            TILE_CACHE[cache_key] = tile
            return tile_coordinates, tile
        except requests.RequestException as error:
            print(f"Could not download map tile: {error}")
            return tile_coordinates, None

    tile_coordinates = [
        (tile_x, tile_y)
        for tile_x in range(min_x, max_x + 1)
        for tile_y in range(min_y, max_y + 1)
    ]
    with ThreadPoolExecutor(max_workers=8) as executor:
        for (tile_x, tile_y), tile in executor.map(fetch_tile, tile_coordinates):
            if tile is not None:
                texture_image.paste(
                    tile,
                    ((tile_x - min_x) * TILE_SIZE, (tile_y - min_y) * TILE_SIZE),
                )

    return texture_image, min_x, min_y, center_x, center_y


def preload_map(projects):
    """Persist the base view and detail zoom levels before opening the viewer."""
    base_zoom = choose_tile_zoom(projects)
    print(f"Preloading map tiles into {TILE_CACHE_DIR}...")
    download_map_texture(projects, zoom=base_zoom)
    min_lat = min(project["lat"] for project in projects)
    max_lat = max(project["lat"] for project in projects)
    min_lon = min(project["lon"] for project in projects)
    max_lon = max(project["lon"] for project in projects)
    focus = ((min_lat + max_lat) / 2.0, (min_lon + max_lon) / 2.0)
    for zoom in range(base_zoom + 1, MAX_DETAIL_ZOOM + 1):
        download_map_texture(projects, zoom=zoom, focus=focus)
    print("Map preload complete. Future launches will use the local tile cache.")


def make_texture(image):
    texture_id = glGenTextures(1)
    glBindTexture(GL_TEXTURE_2D, texture_id)
    glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR)
    glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR)
    glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_WRAP_S, GL_CLAMP_TO_EDGE)
    glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_WRAP_T, GL_CLAMP_TO_EDGE)
    glTexImage2D(
        GL_TEXTURE_2D,
        0,
        GL_RGB,
        image.width,
        image.height,
        0,
        GL_RGB,
        GL_UNSIGNED_BYTE,
        image.tobytes(),
    )
    return texture_id


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


def draw_floor(texture_id, image, min_x, min_y, center_x, center_y):
    world_width = image.width / TILE_SIZE * 40.0
    world_depth = image.height / TILE_SIZE * 40.0
    floor_x = (min_x - center_x) * 40.0
    floor_z = (min_y - center_y) * 40.0

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


def project_position(project, center_x, center_y):
    return (
        (lon_to_tile_x(project["lon"], TILE_ZOOM) - center_x) * 40.0,
        (lat_to_tile_y(project["lat"], TILE_ZOOM) - center_y) * 40.0,
    )


def project_is_visible(project, min_x, max_x, min_y, max_y):
    project_x = lon_to_tile_x(project["lon"], TILE_ZOOM)
    project_y = lat_to_tile_y(project["lat"], TILE_ZOOM)
    return min_x <= project_x <= max_x + 1 and min_y <= project_y <= max_y + 1


def map_point_under_mouse(mouse_position, center_x, center_y):
    """Return map-local coordinates and geographic coordinates under the cursor."""
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
    tile_x = center_x + map_x / 40.0
    tile_y = center_y + map_z / 40.0
    return map_x, map_z, tile_y_to_lat(tile_y, TILE_ZOOM), tile_x_to_lon(tile_x, TILE_ZOOM)


def draw_marker(project, center_x, center_y, quadric, selected=False):
    x, z = project_position(project, center_x, center_y)
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


def pick_project(projects, click_position, center_x, center_y, show_desc=True, show_ga_its=True):
    """Return the closest marker under a screen-space click."""
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
        x, z = project_position(project, center_x, center_y)
        screen_x, screen_y, depth = gluProject(x, 1.2, z)
        distance = math.hypot(screen_x - click_x, screen_y - click_y)
        if 0.0 <= depth <= 1.0 and distance < closest_distance:
            closest_project = project
            closest_distance = distance

    return closest_project


def describe_project(project):
    return (
        f"{project['name']} | {project['sheet']} | "
        f"ID: {project['id'] or 'N/A'} | "
        f"Category: {project['category']} | "
        f"Lat: {project['lat']:.5f}, Lon: {project['lon']:.5f}"
    )


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


def draw_overlap_lines(overlaps, center_x, center_y, min_x, max_x, min_y, max_y, selected_project=None):
    """Render the 2D mockup's cross-project relationships on the 3D floor."""
    glPushAttrib(GL_ENABLE_BIT | GL_COLOR_BUFFER_BIT | GL_LINE_BIT | GL_DEPTH_BUFFER_BIT)
    glDisable(GL_DEPTH_TEST)
    glDisable(GL_TEXTURE_2D)
    glEnable(GL_BLEND)
    glBlendFunc(GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA)
    glLineWidth(4.0)
    rendered_lines = 0
    for overlap in overlaps:
        if not (
            project_is_visible(overlap["first"], min_x, max_x, min_y, max_y)
            and project_is_visible(overlap["second"], min_x, max_x, min_y, max_y)
        ):
            continue
        first_x, first_z = project_position(overlap["first"], center_x, center_y)
        second_x, second_z = project_position(overlap["second"], center_x, center_y)
        is_selected = selected_project in (overlap["first"], overlap["second"])
        red, green, blue = (1.0, 0.95, 0.15) if is_selected else overlap["color"]
        glColor4f(red, green, blue, 0.95 if is_selected else 0.65)
        if is_selected:
            glLineWidth(7.0)
        glBegin(GL_LINES)
        glVertex3f(first_x, 0.12, first_z)
        glVertex3f(second_x, 0.12, second_z)
        glEnd()
        rendered_lines += 1
    glPopAttrib()
    return rendered_lines


def main(preload=False):
    projects = load_projects(EXCEL_FILE)
    if preload:
        preload_map(projects)
        return

    overlaps = [
        overlap
        for project in projects
        for overlap in project.get("overlaps", [])
        if overlap["first"] is project
    ]
    print(f"Loaded {len(projects)} geocoded projects.")
    print(f"Found {len(overlaps)} geographic project overlaps.")
    print(f"Overlap checker: {len(overlaps)} relationships processed; line rendering enabled.")
    print("Downloading OpenStreetMap floor tiles...")
    map_zoom = choose_tile_zoom(projects)
    map_image, min_x, min_y, center_x, center_y = download_map_texture(
        projects,
        zoom=map_zoom,
    )
    max_x = min_x + map_image.width // TILE_SIZE - 1
    max_y = min_y + map_image.height // TILE_SIZE - 1

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

    camera_distance = max(95.0, max(map_image.width, map_image.height) / TILE_SIZE * 20.0)
    camera_rotation_x = 52.0
    camera_rotation_y = 0.0
    camera_offset_x = 0.0
    camera_offset_z = 0.0
    mouse_down = False
    mouse_down_position = (0, 0)
    last_mouse_pos = (0, 0)
    selected_project = None
    pending_click = None
    show_desc = True
    show_ga_its = True
    show_overlaps = True
    clock = pygame.time.Clock()
    running = True

    def focus_camera_on_project(project):
        nonlocal map_image, min_x, min_y, max_x, max_y
        nonlocal center_x, center_y, texture_id, map_zoom
        nonlocal camera_distance, camera_rotation_x, camera_rotation_y
        nonlocal camera_offset_x, camera_offset_z

        target_zoom = min(MAX_DETAIL_ZOOM, TELEPORT_ZOOM)
        map_zoom = target_zoom
        glDeleteTextures([texture_id])
        map_image, min_x, min_y, center_x, center_y = download_map_texture(
            projects,
            zoom=map_zoom,
            focus=(project["lat"], project["lon"]),
        )
        max_x = min_x + map_image.width // TILE_SIZE - 1
        max_y = min_y + map_image.height // TILE_SIZE - 1
        texture_id = make_texture(map_image)

        camera_offset_x, camera_offset_z = project_position(
            project,
            center_x,
            center_y,
        )
        camera_rotation_x = 90.0
        camera_rotation_y = 0.0
        camera_distance = 110.0

    while running:
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                running = False
            elif event.type == pygame.KEYDOWN:
                if event.key == pygame.K_ESCAPE:
                    if info_panel.search_active:
                        info_panel.close_search()
                    else:
                        running = False
                elif event.key == pygame.K_r:
                    info_panel.toggle_ranking()
                elif event.key == pygame.K_1:
                    show_desc = not show_desc
                    info_panel.set_layers(show_desc, show_ga_its, show_overlaps)
                elif event.key == pygame.K_2:
                    show_ga_its = not show_ga_its
                    info_panel.set_layers(show_desc, show_ga_its, show_overlaps)
                elif event.key == pygame.K_3:
                    show_overlaps = not show_overlaps
                    info_panel.set_layers(show_desc, show_ga_its, show_overlaps)
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
                if event.button == 1 and info_panel.panel_button_hit(event.pos):
                    info_panel.toggle_panel()
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
                        mouse_focus = map_point_under_mouse(
                            event.pos,
                            center_x,
                            center_y,
                        )
                        if mouse_focus is None:
                            mouse_focus = (0.0, 0.0, tile_y_to_lat(center_y, map_zoom), tile_x_to_lon(center_x, map_zoom))
                        old_map_x, old_map_z, focus_lat, focus_lon = mouse_focus
                        map_zoom = new_zoom
                        glDeleteTextures([texture_id])
                        map_image, min_x, min_y, center_x, center_y = download_map_texture(
                            projects,
                            zoom=map_zoom,
                            focus=(focus_lat, focus_lon),
                        )
                        max_x = min_x + map_image.width // TILE_SIZE - 1
                        max_y = min_y + map_image.height // TILE_SIZE - 1
                        texture_id = make_texture(map_image)
                        camera_offset_x = (lon_to_tile_x(focus_lon, map_zoom) - center_x) * 40.0 - old_map_x
                        camera_offset_z = (lat_to_tile_y(focus_lat, map_zoom) - center_y) * 40.0 - old_map_z
                        camera_distance = max(95.0, max(map_image.width, map_image.height) / TILE_SIZE * 20.0)
                elif event.button == 5:
                    new_zoom = max(2, map_zoom - 1)
                    if new_zoom != map_zoom:
                        mouse_focus = map_point_under_mouse(
                            event.pos,
                            center_x,
                            center_y,
                        )
                        if mouse_focus is None:
                            mouse_focus = (0.0, 0.0, tile_y_to_lat(center_y, map_zoom), tile_x_to_lon(center_x, map_zoom))
                        old_map_x, old_map_z, focus_lat, focus_lon = mouse_focus
                        map_zoom = new_zoom
                        glDeleteTextures([texture_id])
                        map_image, min_x, min_y, center_x, center_y = download_map_texture(
                            projects,
                            zoom=map_zoom,
                            focus=(focus_lat, focus_lon),
                        )
                        max_x = min_x + map_image.width // TILE_SIZE - 1
                        max_y = min_y + map_image.height // TILE_SIZE - 1
                        texture_id = make_texture(map_image)
                        camera_offset_x = (lon_to_tile_x(focus_lon, map_zoom) - center_x) * 40.0 - old_map_x
                        camera_offset_z = (lat_to_tile_y(focus_lat, map_zoom) - center_y) * 40.0 - old_map_z
                        camera_distance = max(95.0, max(map_image.width, map_image.height) / TILE_SIZE * 20.0)
            elif event.type == pygame.MOUSEBUTTONUP and event.button == 1:
                mouse_down = False
                if math.dist(event.pos, mouse_down_position) <= 6.0:
                    pending_click = event.pos
            elif event.type == pygame.MOUSEMOTION:
                info_panel.update_pointer(event.pos)
                if mouse_down and not info_panel.search_active:
                    delta_x = event.pos[0] - last_mouse_pos[0]
                    delta_y = event.pos[1] - last_mouse_pos[1]
                    camera_rotation_y += delta_x * 0.5
                    camera_rotation_x = max(15.0, min(85.0, camera_rotation_x + delta_y * 0.5))
                    last_mouse_pos = event.pos

        frame_seconds = clock.get_time() / 1000.0
        pressed_keys = pygame.key.get_pressed()
        movement_speed = 42.0 if pressed_keys[pygame.K_LSHIFT] or pressed_keys[pygame.K_RSHIFT] else 18.0
        movement = movement_speed * frame_seconds
        if not info_panel.search_active and pressed_keys[pygame.K_w]:
            camera_offset_z -= movement
        if not info_panel.search_active and pressed_keys[pygame.K_s]:
            camera_offset_z += movement
        if not info_panel.search_active and pressed_keys[pygame.K_a]:
            camera_offset_x -= movement
        if not info_panel.search_active and pressed_keys[pygame.K_d]:
            camera_offset_x += movement

        glClear(GL_COLOR_BUFFER_BIT | GL_DEPTH_BUFFER_BIT)
        glLoadIdentity()
        glTranslatef(0.0, -4.0, -camera_distance)
        glRotatef(camera_rotation_x, 1.0, 0.0, 0.0)
        glRotatef(camera_rotation_y, 0.0, 1.0, 0.0)
        glTranslatef(-camera_offset_x, 0.0, -camera_offset_z)
        if pending_click is not None:
            selected_project = pick_project(
                projects,
                pending_click,
                center_x,
                center_y,
                show_desc,
                show_ga_its,
            )
            pending_click = None
            info_panel.set_project(selected_project)
            if selected_project is None:
                pygame.display.set_caption(
                    "3D Gridlock Project Map | © OpenStreetMap contributors"
                )
            else:
                project_info = describe_project(selected_project)
                print(f"Selected: {project_info}")
                pygame.display.set_caption(project_info)
        draw_floor(texture_id, map_image, min_x, min_y, center_x, center_y)
        rendered_overlap_lines = (
            draw_overlap_lines(
                overlaps,
                center_x,
                center_y,
                min_x,
                max_x,
                min_y,
                max_y,
                selected_project,
            )
            if show_overlaps
            else 0
        )
        info_panel.set_overlap_status(len(overlaps), rendered_overlap_lines)
        draw_axes(12.0)
        for project in projects:
            if not project_is_visible(project, min_x, max_x, min_y, max_y):
                continue
            if project["sheet"] == "DESC Geocoded" and not show_desc:
                continue
            if project["sheet"] == "GA ITS Geocoded" and not show_ga_its:
                continue
            draw_marker(
                project,
                center_x,
                center_y,
                quadric,
                selected=project is selected_project,
            )
        info_panel.draw()
        pygame.display.flip()
        clock.tick(60)

    gluDeleteQuadric(quadric)
    glDeleteTextures([texture_id])
    info_panel.close()
    pygame.quit()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Run the 3D project map viewer.")
    parser.add_argument(
        "--preload-map",
        action="store_true",
        help="Download and save the map tiles used by all zoom levels, then exit.",
    )
    main(preload=parser.parse_args().preload_map)
