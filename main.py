import pygame
from pygame.locals import *
from OpenGL.GL import *
from OpenGL.GLU import *
from datetime import datetime
import random
import math
import os

import openpyxl

# ============================================================
# PARSER
# ============================================================
#
# Reads the geocoded project workbook and returns a flat list of
# project dicts, one per row that actually has coordinates.
#
# Only rows where center_lat/center_lon are populated (i.e. the
# geocoder found a match) are kept -- rows marked UNMATCHED are
# skipped since we have nowhere to place them.

EXCEL_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "gridlock_project_tables_geocoded_nominatim.xlsx"
)

# Half-width (in world units) that the full lat/lon spread should
# be scaled to fit inside. Keeps the layout roughly centered
# around the origin like the old random site did.
TARGET_HALF_EXTENT = 22.0

# name -> category field to use per sheet, and a color per category.
# Anything not listed here falls back to a neutral gray.
SHEET_CATEGORY_FIELD = {
    "DESC Geocoded": "Status",
    "GA ITS Geocoded": "Sponsor (GPC/GTC/MEAG/DU/SAV)",
}

CATEGORY_COLORS = {
    ("DESC Geocoded", "In Progress"): (0.3, 0.8, 0.4),
    ("DESC Geocoded", "Planned"): (0.9, 0.7, 0.2),

    ("GA ITS Geocoded", "GPC"): (0.2, 0.6, 1.0),
    ("GA ITS Geocoded", "GTC"): (0.7, 0.4, 1.0),
    ("GA ITS Geocoded", "SAV"): (0.9, 0.3, 0.5),
    ("GA ITS Geocoded", "MEAG"): (0.4, 0.8, 0.8),
    ("GA ITS Geocoded", "DU"): (1.0, 1.0, 1.0),
}

DEFAULT_COLOR = (0.6, 0.6, 0.6)

# Fields that could hold a human-readable ID depending on the sheet.
ID_FIELD_CANDIDATES = ["Project ID", "TEAMS Number"]

NAME_FIELD = "Project Name / Endpoints (raw title)"


def _row_dict(header, row):
    return {header[i]: row[i] for i in range(len(header))}


def load_projects_from_excel(path):
    """
    Opens the workbook, walks every sheet, and returns a list of
    dicts describing each geocoded project:

        {
            "sheet": "DESC Geocoded",
            "id": "6807 B",
            "name": "Queensboro - Ft Johnson 115 kV ...",
            "category": "In Progress",
            "lat": 32.79,
            "lon": -79.93,
        }
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

        for row in rows[1:]:

            record = _row_dict(header, row)

            lat = record.get("center_lat")
            lon = record.get("center_lon")

            # Skip anything the geocoder couldn't match
            if lat is None or lon is None:
                continue

            project_id = None

            for field in ID_FIELD_CANDIDATES:

                if field in record and record[field] is not None:
                    project_id = record[field]
                    break

            category = record.get(category_field) if category_field else None

            projects.append({
                "sheet": sheet_name,
                "id": project_id,
                "name": record.get(NAME_FIELD),
                "category": category,
                "lat": float(lat),
                "lon": float(lon),
            })

    return projects


def latlon_to_world(projects):
    """
    Projects lat/lon onto a flat X/Z plane using an equirectangular
    projection centered on the data's own mean position, then scales
    it uniformly (same scale on both axes, so nothing gets stretched)
    to fit inside TARGET_HALF_EXTENT.

    Returns the same list of dicts with "x" and "z" added.
    """

    if not projects:
        return projects

    lats = [p["lat"] for p in projects]
    lons = [p["lon"] for p in projects]

    lat0 = (min(lats) + max(lats)) / 2
    lon0 = (min(lons) + max(lons)) / 2

    # Longitude degrees are "shorter" than latitude degrees away from
    # the equator -- this correction keeps the layout from looking
    # squished east-west.
    lon_correction = math.cos(math.radians(lat0))

    span_x = (max(lons) - min(lons)) * lon_correction
    span_z = (max(lats) - min(lats))

    span = max(span_x, span_z, 0.0001)

    scale = (TARGET_HALF_EXTENT * 2) / span

    for p in projects:

        p["x"] = (p["lon"] - lon0) * lon_correction * scale

        # Flip so increasing latitude (north) moves toward -Z,
        # matching a conventional "north is up/back" layout.
        p["z"] = -(p["lat"] - lat0) * scale

    return projects


# ============================================================
# SETTINGS
# ============================================================

# Fixed footprint for every project marker; height varies a bit
# for visual interest only (it carries no data meaning).
MARKER_WIDTH = 0.9
MARKER_DEPTH = 0.9
MIN_HEIGHT = 1.5
MAX_HEIGHT = 4.0


# --------------------------------------------------------
# SKYBOX COLORS
# --------------------------------------------------------
def get_sky_color():

    hour = datetime.now().hour + datetime.now().minute / 60

    # --------------------------------------------------------
    # NIGHT: 7 PM - 6 AM
    # --------------------------------------------------------
    if hour >= 19 or hour < 6:

        return (
            0.02,
            0.03,
            0.10
        )

    # --------------------------------------------------------
    # SUNRISE: 6 AM - 8 AM
    # --------------------------------------------------------
    elif hour < 8:

        return (
            0.35,
            0.20,
            0.35
        )

    # --------------------------------------------------------
    # DAY: 8 AM - 5 PM
    # --------------------------------------------------------
    elif hour < 17:

        return (
            0.15,
            0.35,
            0.60
        )

    # --------------------------------------------------------
    # SUNSET: 5 PM - 7 PM
    # --------------------------------------------------------
    else:

        return (
            0.45,
            0.18,
            0.10
        )

# ============================================================
# CAMERA
# ============================================================

camera_distance = 45.0
camera_rotation_x = 35.0
camera_rotation_y = 45.0

mouse_down = False
last_mouse_pos = (0, 0)


# ============================================================
# CREATE A BOX
# ============================================================

def create_box(name, x, y, z, width, height, depth, color):

    # Center of box
    x1 = x - width / 2
    x2 = x + width / 2

    y1 = y
    y2 = y + height

    z1 = z - depth / 2
    z2 = z + depth / 2

    points = [
        # Bottom
        (x1, y1, z1),  # 0
        (x2, y1, z1),  # 1
        (x2, y1, z2),  # 2
        (x1, y1, z2),  # 3

        # Top
        (x1, y2, z1),  # 4
        (x2, y2, z1),  # 5
        (x2, y2, z2),  # 6
        (x1, y2, z2),  # 7
    ]

    connections = [
        # Bottom
        (0, 1),
        (1, 2),
        (2, 3),
        (3, 0),

        # Top
        (4, 5),
        (5, 6),
        (6, 7),
        (7, 4),

        # Vertical
        (0, 4),
        (1, 5),
        (2, 6),
        (3, 7),
    ]

    return {
        "name": name,
        "points": points,
        "connections": connections,
        "color": color,

        # Store dimensions for intersection checking
        "min": (x1, y1, z1),
        "max": (x2, y2, z2),

        "intersects": False,
    }


# ============================================================
# GENERATE SITE FROM THE GEOCODED EXCEL WORKBOOK
# ============================================================

def generate_site_from_excel(path):

    projects = load_projects_from_excel(path)

    projects = latlon_to_world(projects)

    areas = []

    for p in projects:

        height = random.uniform(
            MIN_HEIGHT,
            MAX_HEIGHT
        )

        color = CATEGORY_COLORS.get(
            (p["sheet"], p["category"]),
            DEFAULT_COLOR
        )

        label_bits = [str(p["name"])]

        if p["id"] is not None:
            label_bits.append(f"[{p['id']}]")

        label_bits.append(f"({p['sheet']})")

        name = " ".join(label_bits)

        area = create_box(
            name,
            p["x"],
            0,
            p["z"],
            MARKER_WIDTH,
            height,
            MARKER_DEPTH,
            color
        )

        areas.append(area)

    return areas


# ============================================================
# CHECK IF TWO BOXES INTERSECT
# ============================================================

def boxes_intersect(a, b):

    a_min = a["min"]
    a_max = a["max"]

    b_min = b["min"]
    b_max = b["max"]

    # X overlap
    x_overlap = (
        a_min[0] <= b_max[0]
        and a_max[0] >= b_min[0]
    )

    # Y overlap
    y_overlap = (
        a_min[1] <= b_max[1]
        and a_max[1] >= b_min[1]
    )

    # Z overlap
    z_overlap = (
        a_min[2] <= b_max[2]
        and a_max[2] >= b_min[2]
    )

    return (
        x_overlap
        and y_overlap
        and z_overlap
    )


# ============================================================
# FIND ALL INTERSECTIONS
# ============================================================

def find_intersections(areas):

    # Reset
    for area in areas:
        area["intersects"] = False

    # Compare every box with every other box
    for i in range(len(areas)):

        for j in range(i + 1, len(areas)):

            a = areas[i]
            b = areas[j]

            if boxes_intersect(a, b):

                # Both boxes become orange
                a["intersects"] = True
                b["intersects"] = True


# ============================================================
# OPENGL SETUP
# ============================================================

def setup_opengl(width, height):

    glViewport(
        0,
        0,
        width,
        height
    )

    glMatrixMode(GL_PROJECTION)

    glLoadIdentity()

    gluPerspective(
        45,
        width / height,
        0.1,
        1000
    )

    glMatrixMode(GL_MODELVIEW)

    glEnable(GL_DEPTH_TEST)

    glEnable(GL_LINE_SMOOTH)


# ============================================================
# DRAW GRID
# ============================================================

def draw_grid(size=25):

    glColor3f(
        0.08,   # Red
        0.35,   # Green
        0.08    # Blue
    )

    glBegin(GL_QUADS)

    glVertex3f(
        -size,
        0,
        -size
    )

    glVertex3f(
        size,
        0,
        -size
    )

    glVertex3f(
        size,
        0,
        size
    )

    glVertex3f(
        -size,
        0,
        size
    )

    glEnd()



# ============================================================
# DRAW AXES
# ============================================================

def draw_axes(length=10):

    glLineWidth(3)

    glBegin(GL_LINES)

    # X = RED
    glColor3f(1, 0, 0)

    glVertex3f(
        0,
        0,
        0
    )

    glVertex3f(
        length,
        0,
        0
    )

    # Y = GREEN
    glColor3f(0, 1, 0)

    glVertex3f(
        0,
        0,
        0
    )

    glVertex3f(
        0,
        length,
        0
    )

    # Z = BLUE
    glColor3f(0, 0, 1)

    glVertex3f(
        0,
        0,
        0
    )

    glVertex3f(
        0,
        0,
        length
    )

    glEnd()


# ============================================================
# DRAW A BOX
# ============================================================

def draw_area(area, hovered=False):

    points = area["points"]
    connections = area["connections"]

    # --------------------------------------------------------
    # Determine color
    # --------------------------------------------------------

    if area["intersects"]:

        # ORANGE = overlapping another area
        color = (1.0, 0.3, 0.0)

    elif hovered:

        # YELLOW = mouse is over it
        color = (1.0, 1.0, 0.0)

    else:

        color = area["color"]

    # --------------------------------------------------------
    # Draw lines
    # --------------------------------------------------------

    if area["intersects"]:

        glLineWidth(2)

    elif hovered:

        glLineWidth(2)

    else:

        glLineWidth(2)

    glColor3f(*color)

    glBegin(GL_LINES)

    for a, b in connections:

        glVertex3f(
            *points[a]
        )

        glVertex3f(
            *points[b]
        )

    glEnd()

    # --------------------------------------------------------
    # Draw points
    # --------------------------------------------------------

    glPointSize(
        9 if hovered else 6
    )

    glBegin(GL_POINTS)

    for point in points:

        glColor3f(
            1.0,
            1.0,
            0.0
        )

        glVertex3f(
            *point
        )

    glEnd()


# ============================================================
# CAMERA
# ============================================================

def update_camera():

    glLoadIdentity()

    glTranslatef(
        0,
        0,
        -camera_distance
    )

    glRotatef(
        camera_rotation_x,
        1,
        0,
        0
    )

    glRotatef(
        camera_rotation_y,
        0,
        1,
        0
    )


# ============================================================
# RAY / BOX INTERSECTION
# ============================================================

def ray_box_intersection(
    ray_origin,
    ray_direction,
    box_min,
    box_max
):

    t_min = -float("inf")
    t_max = float("inf")

    for i in range(3):

        origin = ray_origin[i]
        direction = ray_direction[i]

        minimum = box_min[i]
        maximum = box_max[i]

        if abs(direction) < 0.000001:

            if (
                origin < minimum
                or origin > maximum
            ):

                return False

        else:

            t1 = (
                minimum - origin
            ) / direction

            t2 = (
                maximum - origin
            ) / direction

            if t1 > t2:
                t1, t2 = t2, t1

            t_min = max(
                t_min,
                t1
            )

            t_max = min(
                t_max,
                t2
            )

            if t_min > t_max:
                return False

    return t_max >= max(
        t_min,
        0
    )


# ============================================================
# FIND HOVERED AREA
# ============================================================

def find_hovered_area(
    mouse_x,
    mouse_y,
    width,
    height,
    areas
):

    modelview = glGetDoublev(
        GL_MODELVIEW_MATRIX
    )

    projection = glGetDoublev(
        GL_PROJECTION_MATRIX
    )

    viewport = glGetIntegerv(
        GL_VIEWPORT
    )

    # Convert Pygame Y to OpenGL Y
    opengl_y = height - mouse_y

    try:

        near_point = gluUnProject(
            mouse_x,
            opengl_y,
            0,
            modelview,
            projection,
            viewport
        )

        far_point = gluUnProject(
            mouse_x,
            opengl_y,
            1,
            modelview,
            projection,
            viewport
        )

    except Exception:

        return None

    ray_origin = near_point

    direction = (
        far_point[0] - near_point[0],
        far_point[1] - near_point[1],
        far_point[2] - near_point[2],
    )

    length = math.sqrt(
        direction[0] ** 2
        + direction[1] ** 2
        + direction[2] ** 2
    )

    if length == 0:
        return None

    ray_direction = (
        direction[0] / length,
        direction[1] / length,
        direction[2] / length
    )

    # Find first box hit
    for area in areas:

        if ray_box_intersection(
            ray_origin,
            ray_direction,
            area["min"],
            area["max"]
        ):

            return area

    return None


# ============================================================
# DRAW TEXT
# ============================================================

def draw_text(text, x, y):

    font = pygame.font.Font(
        None,
        28
    )

    surface = font.render(
        text,
        True,
        (255, 255, 255)
    )

    text_width = surface.get_width()
    text_height = surface.get_height()

    text_data = pygame.image.tostring(
        surface,
        "RGBA",
        True
    )

    texture_id = glGenTextures(1)

    glBindTexture(
        GL_TEXTURE_2D,
        texture_id
    )

    glTexParameteri(
        GL_TEXTURE_2D,
        GL_TEXTURE_MIN_FILTER,
        GL_LINEAR
    )

    glTexParameteri(
        GL_TEXTURE_2D,
        GL_TEXTURE_MAG_FILTER,
        GL_LINEAR
    )

    glTexImage2D(
        GL_TEXTURE_2D,
        0,
        GL_RGBA,
        text_width,
        text_height,
        0,
        GL_RGBA,
        GL_UNSIGNED_BYTE,
        text_data
    )

    # Switch to 2D
    glMatrixMode(GL_PROJECTION)

    glPushMatrix()

    glLoadIdentity()

    screen_width, screen_height = (
        pygame.display
        .get_surface()
        .get_size()
    )

    glOrtho(
        0,
        screen_width,
        screen_height,
        0,
        -1,
        1
    )

    glMatrixMode(GL_MODELVIEW)

    glPushMatrix()

    glLoadIdentity()

    glDisable(GL_DEPTH_TEST)

    glEnable(GL_BLEND)

    glBlendFunc(
        GL_SRC_ALPHA,
        GL_ONE_MINUS_SRC_ALPHA
    )

    glEnable(GL_TEXTURE_2D)

    glBindTexture(
        GL_TEXTURE_2D,
        texture_id
    )

    glColor4f(
        1,
        1,
        1,
        1
    )

    glBegin(GL_QUADS)

    glTexCoord2f(0, 1)
    glVertex2f(
        x,
        y
    )

    glTexCoord2f(1, 1)
    glVertex2f(
        x + text_width,
        y
    )

    glTexCoord2f(1, 0)
    glVertex2f(
        x + text_width,
        y + text_height
    )

    glTexCoord2f(0, 0)
    glVertex2f(
        x,
        y + text_height
    )

    glEnd()

    glDisable(GL_TEXTURE_2D)

    glDisable(GL_BLEND)

    glEnable(GL_DEPTH_TEST)

    glDeleteTextures(
        [texture_id]
    )

    glPopMatrix()

    glMatrixMode(GL_PROJECTION)

    glPopMatrix()

    glMatrixMode(GL_MODELVIEW)


# ============================================================
# MAIN
# ============================================================

def main():

    global camera_distance
    global camera_rotation_x
    global camera_rotation_y

    global mouse_down
    global last_mouse_pos

    pygame.init()

    width = 1200
    height = 800

    pygame.display.set_mode(
        (
            width,
            height
        ),
        DOUBLEBUF | OPENGL
    )

    pygame.display.set_caption(
        "3D Construction Site"
    )

    setup_opengl(
        width,
        height
    )

    # --------------------------------------------------------
    # Build the site from the geocoded Excel workbook
    # --------------------------------------------------------

    areas = generate_site_from_excel(
        EXCEL_FILE
    )

    # Find overlaps
    find_intersections(
        areas
    )

    # Count intersections
    intersection_count = sum(
        1
        for area in areas
        if area["intersects"]
    )

    print(
        f"Loaded {len(areas)} geocoded projects."
    )

    print(
        f"{intersection_count} projects "
        f"are involved in intersections."
    )

    clock = pygame.time.Clock()

    running = True

    hovered_area = None

    # ========================================================
    # MAIN LOOP
    # ========================================================

    while running:

        # ----------------------------------------------------
        # EVENTS
        # ----------------------------------------------------

        for event in pygame.event.get():

            if event.type == pygame.QUIT:

                running = False

            elif event.type == pygame.MOUSEBUTTONDOWN:

                if event.button == 1:

                    mouse_down = True

                    last_mouse_pos = (
                        event.pos
                    )

                elif event.button == 4:

                    camera_distance -= 1

                elif event.button == 5:

                    camera_distance += 1

            elif event.type == pygame.MOUSEBUTTONUP:

                if event.button == 1:

                    mouse_down = False

            elif event.type == pygame.MOUSEMOTION:

                if mouse_down:

                    x, y = event.pos

                    old_x, old_y = (
                        last_mouse_pos
                    )

                    dx = x - old_x
                    dy = y - old_y

                    camera_rotation_y += (
                        dx * 0.5
                    )

                    camera_rotation_x += (
                        dy * 0.5
                    )

                    last_mouse_pos = (
                        event.pos
                    )

        # ----------------------------------------------------
        # Camera limits
        # ----------------------------------------------------

        camera_distance = max(
            2,
            min(
                camera_distance,
                150
            )
        )

        # ----------------------------------------------------
        # Clear screen
        # ----------------------------------------------------

        sky_color = get_sky_color()

        glClearColor(
            sky_color[0],
            sky_color[1],
            sky_color[2],
            1.0
        )


        glClear(
            GL_COLOR_BUFFER_BIT
            | GL_DEPTH_BUFFER_BIT
        )

        # ----------------------------------------------------
        # Camera
        # ----------------------------------------------------

        update_camera()

        # ----------------------------------------------------
        # Environment
        # ----------------------------------------------------

        draw_grid()

        #draw_axes()

        # ----------------------------------------------------
        # Mouse hover
        # ----------------------------------------------------

        mouse_x, mouse_y = (
            pygame.mouse.get_pos()
        )

        hovered_area = find_hovered_area(
            mouse_x,
            mouse_y,
            width,
            height,
            areas
        )

        # ----------------------------------------------------
        # Draw boxes
        # ----------------------------------------------------

        for area in areas:

            draw_area(
                area,
                hovered=(
                    area == hovered_area
                )
            )

        # ----------------------------------------------------
        # Project name
        # ----------------------------------------------------

        if hovered_area:

            draw_text(
                hovered_area["name"],
                mouse_x + 15,
                mouse_y + 15
            )

        # ----------------------------------------------------
        # Display
        # ----------------------------------------------------

        pygame.display.flip()

        clock.tick(60)

    pygame.quit()


# ============================================================
# START 
# ============================================================

if __name__ == "__main__":
    main()