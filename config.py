import os
from dotenv import load_dotenv

APP_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(APP_DIR, ".env"))

EXCEL_FILE = os.path.join(
    APP_DIR, "gridlock_project_tables_geocoded_nominatim (1).xlsx"
)
UPLOADS_FILE = os.path.join(APP_DIR, "uploaded_projects.json")
DATABASE_FILE = os.path.join(APP_DIR, "projects.sqlite3")

MAX_UPLOAD_BYTES = 15 * 1024 * 1024
MAX_DOCUMENT_CHARS = 120_000
MAX_ARCHIVE_UNCOMPRESSED_BYTES = 25 * 1024 * 1024
MAX_PROJECTS_PER_UPLOAD = 100
MAX_GEOCODES_PER_UPLOAD = 10
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

ICON_COLORS = (
    "blue", "cadetblue", "darkgreen", "darkpurple", "green", "orange",
    "purple", "red", "darkred",
)

FIELD_ALIASES = {
    "latitude": ("centerlatitude", "centerlat", "latitude", "lat", "y"),
    "longitude": ("centerlongitude", "centerlon", "longitude", "lon", "lng", "x"),
    "name": ("projectname", "projecttitle", "name", "title", "project"),
    "identifier": ("projectid", "teamsnumber", "projectnumber", "id", "number"),
    "category": ("category", "sponsor", "status", "type"),
}