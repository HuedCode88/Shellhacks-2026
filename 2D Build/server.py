import json
import sys
import os
import zipfile
from email import policy
from email.parser import BytesParser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlsplit, urlparse, parse_qs
import subprocess

from config import APP_DIR, MAX_UPLOAD_BYTES
from database import (
    load_all_projects, database_backend, save_uploaded_projects,
    clear_uploaded_projects
)
from services import (
    extract_projects_with_gemini, normalize_uploaded_projects,
    build_map, _safe_gemini_error
)

class ProjectMapHandler(BaseHTTPRequestHandler):

    def _send_bytes(self, status, content_type, body):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _send_json(self, status, payload):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self._send_bytes(status, "application/json; charset=utf-8", body)

    def _send_file(self, path, content_type):
        try:
            with open(path, "rb") as source:
                body = source.read()
        except OSError:
            self._send_json(500, {"error": "The map interface could not be loaded."})
            return
        self._send_bytes(200, content_type, body)

    def _parse_document_upload(self, body):
        content_type = self.headers.get("Content-Type", "")
        if not content_type.lower().startswith("multipart/form-data;"):
            raise ValueError("Choose a document to upload.")

        message = BytesParser(policy=policy.default).parsebytes(
            f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("latin-1")
            + body
        )

        if not message.is_multipart():
            raise ValueError("The upload form was not formatted correctly.")

        uploads = []
        utility_name = ""
        for part in message.iter_parts():
            if part.get_content_disposition() != "form-data":
                continue
            field_name = part.get_param("name", header="content-disposition")

            if field_name == "utility_name":
                utility_name = (part.get_content() or "").strip()
                continue

            if field_name != "document":
                continue

            filename = part.get_filename()
            content = part.get_payload(decode=True)
            if filename and content is not None:
                uploads.append((filename, content))

        if len(uploads) != 1:
            raise ValueError("Upload one document at a time.")

        filename, content = uploads[0]
        filename = filename.replace("\\", "/").rsplit("/", 1)[-1]

        if len(content) > MAX_UPLOAD_BYTES:
            raise OverflowError("The document exceeds the 15 MB upload limit.")

        return filename, content, utility_name

    def do_GET(self):
        path = urlsplit(self.path).path

        if path == "/":
            self._send_file(
                os.path.join(APP_DIR, "templates", "index.html"),
                "text/html; charset=utf-8",
            )
            return

        static_files = {
            "/static/app.js": ("static/app.js", "text/javascript; charset=utf-8"),
            "/static/styles.css": ("static/styles.css", "text/css; charset=utf-8"),
        }
        if path in static_files:
            relative_path, content_type = static_files[path]
            self._send_file(os.path.join(APP_DIR, relative_path), content_type)
            return

        if path == "/api/status":
            try:
                projects = load_all_projects()
            except (OSError, ValueError):
                self._send_json(500, {"error": "Project data could not be read."})
                return

            uploaded_count = sum(project["source_type"] == "upload" for project in projects)
            self._send_json(200, {
                "uploaded_count": uploaded_count,
                "total_count": len(projects),
                "gemini_configured": bool(os.environ.get("GEMINI_API_KEY")),
                "database_backend": database_backend(),
            })
            return

        if path not in ("/map", "/project_map.html"):
            self._send_json(404, {"error": "Not found."})
            return

        try:
            projects = load_all_projects()
            page = build_map(projects).encode("utf-8")
        except (OSError, ValueError) as error:
            self._send_json(500, {"error": str(error)})
            return

        self._send_bytes(200, "text/html; charset=utf-8", page)

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path == '/run-main':
            project_dir = os.path.normpath(os.path.join(APP_DIR, "..", "3D Build"))
            subprocess.Popen(
                [sys.executable, 'main.py'],
                cwd=project_dir
            )
            self._send_json(200, {"status": "launched"})
            return

        if urlsplit(self.path).path != "/api/upload":
            self._send_json(404, {"error": "Not found."})
            return

        if not os.environ.get("GEMINI_API_KEY"):
            self._send_json(
                503,
                {"error": "Gemini is not configured. Set GEMINI_API_KEY and restart the app."},
            )
            return

        origin = self.headers.get("Origin")
        allowed_hosts = {
            f"127.0.0.1:{self.server.server_port}",
            f"localhost:{self.server.server_port}",
        }
        if origin and (
            not origin.startswith("http://")
            or origin.removeprefix("http://") not in allowed_hosts
        ):
            self._send_json(403, {"error": "Cross-origin uploads are not allowed."})
            return

        try:
            content_length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            self._send_json(400, {"error": "Invalid upload size."})
            return

        if content_length <= 0:
            self._send_json(400, {"error": "Choose a document to upload."})
            return

        if content_length > MAX_UPLOAD_BYTES + 128 * 1024:
            self._send_json(413, {"error": "The document exceeds the 15 MB upload limit."})
            return

        body = self.rfile.read(content_length)
        if len(body) != content_length:
            self._send_json(400, {"error": "The upload was incomplete."})
            return

        try:
            filename, content, utility_name = self._parse_document_upload(body)
            extracted = extract_projects_with_gemini(filename, content)
            projects, skipped, geocode_count = normalize_uploaded_projects(
                extracted, filename, utility_name
            )
        except OverflowError as error:
            self._send_json(413, {"error": str(error)})
            return
        except RuntimeError as error:
            self._send_json(503, {"error": str(error)})
            return
        except (ValueError, zipfile.BadZipFile, KeyError) as error:
            self._send_json(422, {"error": str(error)})
            return
        except Exception as error:
            safe_error = _safe_gemini_error(error)
            print(safe_error)
            self._send_json(
                502,
                {"error": safe_error},
            )
            return

        if not projects:
            self._send_json(422, {
                "error": "No mappable projects were found. Include project locations or coordinates.",
                "extracted": len(extracted),
                "skipped": skipped,
            })
            return

        try:
            added, total = save_uploaded_projects(projects)
        except (OSError, ValueError):
            self._send_json(500, {"error": "Extracted projects could not be saved locally."})
            return

        self._send_json(200, {
            "added": added,
            "uploaded_count": total,
            "extracted": len(extracted),
            "not_mappable": skipped,
            "already_imported": len(projects) - added,
            "skipped": skipped + len(projects) - added,
            "geocoded": geocode_count,
        })

    def do_DELETE(self):
        if urlsplit(self.path).path != "/api/uploaded-projects":
            self._send_json(404, {"error": "Not found."})
            return

        clear_uploaded_projects()
        self._send_json(200, {"uploaded_count": 0})

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
    print(f"Project database: {database_backend()}")
    print("Refresh the page to load the latest workbook data. Press Ctrl+C to stop.")

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nMap server stopped.")
    finally:
        server.server_close()