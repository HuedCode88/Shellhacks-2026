import pygame
from OpenGL.GL import *


class ProjectInfoPanel:
    """Draws selected workbook project data as an in-window OpenGL overlay."""

    PANEL_SIZE = (500, 330)

    def __init__(self, window_size):
        self.window_width, self.window_height = window_size
        self.surface = pygame.Surface(self.PANEL_SIZE, pygame.SRCALPHA)
        self.title_font = pygame.font.Font(None, 25)
        self.body_font = pygame.font.Font(None, 20)
        self.small_font = pygame.font.Font(None, 17)
        self.texture_id = glGenTextures(1)
        self.selected_project = None
        self.projects = []
        self.search_query = ""
        self.search_active = False
        self.overlap_count = 0
        self.overlap_lines_rendered = 0
        self.dirty = True

    def set_project(self, project):
        if project is not self.selected_project:
            self.selected_project = project
            self.dirty = True

    def set_projects(self, projects):
        self.projects = projects

    def begin_search(self):
        self.search_active = True
        self.search_query = ""
        self.dirty = True

    def close_search(self):
        self.search_active = False
        self.search_query = ""
        self.dirty = True

    def append_search_text(self, text):
        if self.search_active:
            self.search_query += text
            self.dirty = True

    def remove_search_character(self):
        if self.search_active and self.search_query:
            self.search_query = self.search_query[:-1]
            self.dirty = True

    def search_matches(self):
        query = self.search_query.strip().lower()
        if not query:
            return []
        return [
            project
            for project in self.projects
            if query in " ".join(
                str(project.get(field, ""))
                for field in ("id", "name", "sheet", "category")
            ).lower()
        ][:8]

    def choose_search_result(self):
        matches = self.search_matches()
        self.close_search()
        return matches[0] if matches else None

    def set_overlap_status(self, processed, rendered):
        if processed != self.overlap_count or rendered != self.overlap_lines_rendered:
            self.overlap_count = processed
            self.overlap_lines_rendered = rendered
            self.dirty = True

    def _fit_text(self, text, font, width):
        text = str(text)
        if font.size(text)[0] <= width:
            return text
        while text and font.size(text + "...")[0] > width:
            text = text[:-1]
        return text + "..."

    def _render_panel(self):
        self.surface.fill((13, 20, 29, 238))
        pygame.draw.rect(
            self.surface,
            (66, 190, 220, 255),
            self.surface.get_rect(),
            width=2,
        )

        project = self.selected_project
        if project is None:
            self.surface.blit(
                self.title_font.render("Select a job marker", True, (235, 247, 255)),
                (18, 42 if self.search_active else 18),
            )
            self.surface.blit(
                self.body_font.render("Click a colored 3D point to inspect its XLSX data.", True, (205, 220, 232)),
                (18, 80 if self.search_active else 56),
            )
            self.surface.blit(
                self.body_font.render("WASD move | Shift: faster | Drag: orbit | Wheel: detail zoom", True, (205, 220, 232)),
                (18, 108 if self.search_active else 84),
            )
            self.surface.blit(
                self.body_font.render(
                    f"Overlap checker: {self.overlap_count} processed | {self.overlap_lines_rendered} lines",
                    True,
                    (120, 240, 160) if self.overlap_count == self.overlap_lines_rendered else (255, 190, 80),
                ),
                (18, 136 if self.search_active else 112),
            )
            self.surface.blit(
                self.small_font.render(
                    "Legend: green/orange DESC | blue/purple/pink GA ITS | cyan overlap",
                    True,
                    (150, 190, 205),
                ),
                (18, 164 if self.search_active else 140),
            )
        else:
            margin = 18
            text_width = self.PANEL_SIZE[0] - margin * 2
            title = self._fit_text(project["name"], self.title_font, text_width)
            self.surface.blit(
                self.title_font.render(title, True, (235, 247, 255)),
                (margin, 42 if self.search_active else margin),
            )

        margin = 18
        text_width = self.PANEL_SIZE[0] - margin * 2
        lines = [] if project is None else [
            ("Source", project["sheet"]),
            ("ID", project["id"] or "N/A"),
            ("Category", project["category"]),
            ("Coordinates", f"{project['lat']:.5f}, {project['lon']:.5f}"),
        ]
        details = project.get("details", {}) if project is not None else {}
        for field in (
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
        ):
            if project is not None and field in details:
                lines.append((field, details[field]))

        overlaps = sorted(
            project.get("overlaps", []) if project is not None else [],
            key=lambda overlap: overlap["distance_km"],
        )
        if overlaps:
            lines.append(("Nearby coordination", f"{len(overlaps)} related job(s)"))
            for overlap in overlaps[:5]:
                other_project = (
                    overlap["second"]
                    if overlap["first"] is project
                    else overlap["first"]
                )
                timeline = (
                    "Unknown"
                    if overlap["timeline_overlap"] is None
                    else "Yes"
                    if overlap["timeline_overlap"]
                    else "No"
                )
                lines.append(
                    (
                        "Overlap",
                        f"{other_project['name']} | {overlap['distance_mi']} mi | "
                        f"Timeline: {timeline} | {overlap['geographic_tier']}",
                    )
                )

            lines.append(("Overlap checker", f"{self.overlap_count} processed / {self.overlap_lines_rendered} lines"))

        y = 78 if self.search_active else 52
        for label, value in lines:
            label_text = self._fit_text(f"{label}: {value}", self.body_font, text_width)
            self.surface.blit(
                self.body_font.render(label_text, True, (205, 220, 232)),
                (margin, y),
            )
            y += 25
            if y > self.PANEL_SIZE[1] - 26:
                break

        if self.search_active:
            search_text = self._fit_text(
                f"Search: {self.search_query or 'type project ID, name, sheet, or category'}",
                self.body_font,
                text_width,
            )
            pygame.draw.rect(self.surface, (34, 48, 62), (12, 10, self.PANEL_SIZE[0] - 24, 26))
            self.surface.blit(self.body_font.render(search_text, True, (245, 250, 255)), (18, 14))
            matches = self.search_matches()
            if matches:
                match_text = " | ".join(str(match["name"]) for match in matches[:3])
                self.surface.blit(
                    self.small_font.render(self._fit_text(match_text, self.small_font, text_width), True, (150, 220, 240)),
                    (18, 38),
                )
            elif self.search_query:
                self.surface.blit(self.small_font.render("No matching projects", True, (255, 170, 130)), (18, 38))

        hint = self.small_font.render("Ctrl+F search | Enter select | Esc close search", True, (132, 158, 174))
        self.surface.blit(hint, (margin, self.PANEL_SIZE[1] - 22))

    def _upload_texture(self):
        self._render_panel()
        pixels = pygame.image.tostring(self.surface, "RGBA", True)
        glBindTexture(GL_TEXTURE_2D, self.texture_id)
        glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR)
        glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR)
        glTexImage2D(
            GL_TEXTURE_2D,
            0,
            GL_RGBA,
            self.surface.get_width(),
            self.surface.get_height(),
            0,
            GL_RGBA,
            GL_UNSIGNED_BYTE,
            pixels,
        )
        self.dirty = False

    def draw(self):
        if self.dirty:
            self._upload_texture()

        panel_width, panel_height = self.PANEL_SIZE
        x = 24
        y = self.window_height - panel_height - 24

        glPushAttrib(GL_ENABLE_BIT | GL_COLOR_BUFFER_BIT | GL_DEPTH_BUFFER_BIT)
        glDisable(GL_DEPTH_TEST)
        glEnable(GL_BLEND)
        glBlendFunc(GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA)
        glEnable(GL_TEXTURE_2D)
        glBindTexture(GL_TEXTURE_2D, self.texture_id)

        glMatrixMode(GL_PROJECTION)
        glPushMatrix()
        glLoadIdentity()
        glOrtho(0, self.window_width, self.window_height, 0, -1, 1)
        glMatrixMode(GL_MODELVIEW)
        glPushMatrix()
        glLoadIdentity()

        glColor4f(1, 1, 1, 1)
        glBegin(GL_QUADS)
        glTexCoord2f(0, 1)
        glVertex2f(x, y)
        glTexCoord2f(1, 1)
        glVertex2f(x + panel_width, y)
        glTexCoord2f(1, 0)
        glVertex2f(x + panel_width, y + panel_height)
        glTexCoord2f(0, 0)
        glVertex2f(x, y + panel_height)
        glEnd()

        glPopMatrix()
        glMatrixMode(GL_PROJECTION)
        glPopMatrix()
        glMatrixMode(GL_MODELVIEW)
        glPopAttrib()

    def close(self):
        glDeleteTextures([self.texture_id])
