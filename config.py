import os
from dotenv import load_dotenv

APP_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(APP_DIR, ".env"))

EXCEL_FILE = os.path.join(
    APP_DIR, "gridlock_project_tables_geocoded.xlsx"
)
UPLOADS_FILE = os.path.join(APP_DIR, "uploaded_projects.json")
DATABASE_FILE = os.path.join(APP_DIR, "projects.sqlite3")

MAX_UPLOAD_BYTES = 15 * 1024 * 1024
MAX_DOCUMENT_CHARS = 120_000  # size of ONE chunk sent to Gemini, not a hard
# document-size ceiling anymore -- see MAX_CHUNKS_PER_UPLOAD below.
MAX_CHUNKS_PER_UPLOAD = 20  # bounds total Gemini calls (and cost) for one
# upload: applies to PDF page-chunks and, now, chunks of any other long
# document too, instead of silently truncating anything past the first
# MAX_DOCUMENT_CHARS characters.
MAX_ARCHIVE_UNCOMPRESSED_BYTES = 25 * 1024 * 1024
MAX_PROJECTS_PER_UPLOAD = 100
GEOCODES_PER_PROJECT_BUDGET = 3  # whole-location + up to 2 endpoint-level Nominatim calls
MAX_GEOCODES_HARD_CAP = 300  # absolute ceiling regardless of document size;
# at 1 request/sec this bounds worst-case geocoding time to ~5 minutes
PDF_PAGES_PER_REQUEST = 4
PDF_PAGE_OVERLAP = 1

GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-3.8-flash")

PROJECT_RESPONSE_SCHEMA = {
    "type": "object",
    "properties": {
        "projects": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "name": {"type": "string"},
                    "project_id": {"type": "string"},
                    "location": {"type": "string"},
                    "latitude": {"type": "string"},
                    "longitude": {"type": "string"},
                    "status": {"type": "string"},
                    "category": {"type": "string"},
                    "description": {"type": "string"},
                },
                "required": [
                    "name", "project_id", "location", "latitude", "longitude",
                    "status", "category", "description",
                ],
                "additionalProperties": False,
            },
        },
    },
    "required": ["projects"],
    "additionalProperties": False,
}

ALLOWED_EXTENSIONS = {".pdf", ".docx", ".xlsx", ".txt", ".md", ".csv"}

# Bounding boxes (south,west,north,east) searched for live Overpass-based
# geocoding of uploaded projects -- South Carolina, then Georgia. Add more
# entries here if you need to cover additional states.
OVERPASS_BBOXES = (
    "32.0,-83.5,35.3,-78.5",   # South Carolina
    "30.3,-85.7,35.1,-80.7",   # Georgia
)
OVERPASS_CACHE_MAX_AGE_DAYS = 7

ICON_COLORS = (
    "blue", "cadetblue", "darkblue", "darkgreen", "darkpurple", "darkred",
    "green", "lightblue", "lightgreen", "lightred", "orange", "pink",
    "purple", "red", "beige", "gray", "black",
)  # "white" deliberately excluded -- poor contrast against the map canvas

FIELD_ALIASES = {
    "latitude": ("centerlatitude", "centerlat", "latitude", "lat", "y"),
    "longitude": ("centerlongitude", "centerlon", "longitude", "lon", "lng", "x"),
    "name": ("projectname", "projecttitle", "name", "title", "project"),
    "identifier": ("projectid", "teamsnumber", "projectnumber", "id", "number"),
    "category": ("category", "sponsor", "status", "type"),
}