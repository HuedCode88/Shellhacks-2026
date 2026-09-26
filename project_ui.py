import pygame
from OpenGL.GL import *


class ProjectInfoPanel:
    """Draws selected workbook project data as an in-window OpenGL overlay."""

    PANEL_SIZE = (500, 330)
    SEARCH_SIZE = (520, 250)
    RANKING_TAB_SIZE = (48, 72)
    PANEL_TAB_SIZE = (48, 72)

    def __init__(self, window_size):
        self.window_width, self.window_height = window_size
        self.surface = pygame.Surface(self.PANEL_SIZE, pygame.SRCALPHA)
        self.ranking_surface = pygame.Surface(self.PANEL_SIZE, pygame.SRCALPHA)
        self.ranking_tab_surface = pygame.Surface(self.RANKING_TAB_SIZE, pygame.SRCALPHA)
        self.search_surface = pygame.Surface(self.SEARCH_SIZE, pygame.SRCALPHA)
        self.panel_tab_surface = pygame.Surface(self.PANEL_TAB_SIZE, pygame.SRCALPHA)
        self.title_font = pygame.font.Font(None, 25)
        self.body_font = pygame.font.Font(None, 20)
        self.small_font = pygame.font.Font(None, 17)
        self.texture_id = glGenTextures(1)
        self.ranking_texture_id = glGenTextures(1)
        self.ranking_tab_texture_id = glGenTextures(1)
        self.search_texture_id = glGenTextures(1)
        self.panel_tab_texture_id = glGenTextures(1)
        self.selected_project = None
        self.projects = []
        self.overlaps = []
        self.overlaps = []
        self.search_query = ""
        self.search_active = False
        self.search_hover_index = None
        self.ranking_active = True
        self.ranking_minimized = True
        self.ranking_offset = 0
        self.ranking_hover_index = None
        self.show_desc = True
        self.show_ga_its = True
        self.show_overlaps = True
        self.panel_visible = True
        self.overlap_count = 0
        self.overlap_lines_rendered = 0
        self.dirty = True

    def set_project(self, project):
        if project is not self.selected_project:
            self.selected_project = project
            self.dirty = True

    def set_projects(self, projects):
        self.projects = projects

    def set_overlaps(self, overlaps):
        self.overlaps = sorted(overlaps, key=lambda overlap: overlap["distance_km"])

    def toggle_ranking_minimized(self):
        self.ranking_minimized = not self.ranking_minimized
        self.dirty = True

    def ranking_wheel_hit(self, position):
        x, y = position
        panel_x = self.window_width - self.PANEL_SIZE[0]
        return panel_x <= x <= panel_x + self.PANEL_SIZE[0] and 24 <= y <= 24 + self.PANEL_SIZE[1]

    def ranking_result_at(self, position):
        if self.ranking_minimized or not self.ranking_active:
            return None
        x, y = position
        panel_x = self.window_width - self.PANEL_SIZE[0]
        if not (panel_x <= x <= self.window_width and 76 <= y <= 76 + 44 * 5):
            return None
        index = self.ranking_offset + (y - 76) // 44
        return self.overlaps[index] if 0 <= index < len(self.overlaps) else None

    def scroll_ranking(self, amount):
        max_offset = max(0, len(self.overlaps) - 5)
        self.ranking_offset = max(0, min(max_offset, self.ranking_offset + amount))
        self.dirty = True

    def ranking_button_hit(self, position):
        x = self.window_width - self.RANKING_TAB_SIZE[0] if self.ranking_minimized else self.window_width - 72
        y = self.window_height // 3 - self.RANKING_TAB_SIZE[1] // 2 if self.ranking_minimized else 24
        width = self.RANKING_TAB_SIZE[0] if self.ranking_minimized else self.PANEL_SIZE[0]
        height = self.RANKING_TAB_SIZE[1] if self.ranking_minimized else 48
        return x <= position[0] <= x + width and y <= position[1] <= y + height

    def search_result_at(self, position):
        x, y = position
        if not (24 <= x <= 24 + self.SEARCH_SIZE[0] and 24 <= y <= 24 + self.SEARCH_SIZE[1]):
            return None, False
        if y < 68:
            self.begin_search()
            return None, True
        matches = self.search_matches()
        index = (y - 72) // 28
        if 0 <= index < len(matches):
            self.close_search()
            return matches[index], True
        return None, True

    def set_overlaps(self, overlaps):
        self.overlaps = sorted(overlaps, key=lambda overlap: overlap["distance_km"])

    def toggle_ranking(self):
        self.ranking_active = not self.ranking_active
        if self.ranking_active:
            self.ranking_minimized = False
        self.dirty = True

    def set_layers(self, show_desc, show_ga_its, show_overlaps):
        self.show_desc = show_desc
        self.show_ga_its = show_ga_its
        self.show_overlaps = show_overlaps
        self.dirty = True

    def toggle_panel(self):
        self.panel_visible = not self.panel_visible
        self.dirty = True

    def panel_button_hit(self, position):
        panel_x = 0
        panel_y = self.window_height - self.PANEL_SIZE[1] - 24
        if self.panel_visible:
            return (
                panel_x + self.PANEL_SIZE[0] - 44 <= position[0] <= panel_x + self.PANEL_SIZE[0]
                and panel_y <= position[1] <= panel_y + 44
            )
        tab_y = self.window_height - 24 - self.PANEL_TAB_SIZE[1]
        return panel_x <= position[0] <= panel_x + self.PANEL_TAB_SIZE[0] and tab_y <= position[1] <= tab_y + self.PANEL_TAB_SIZE[1]

    def begin_search(self):
        self.search_active = True
        self.search_query = ""
        self.dirty = True

    def close_search(self):
        self.search_active = False
        self.search_query = ""
        self.search_hover_index = None
        self.dirty = True

    def update_pointer(self, position):
        x, y = position
        local_x = x - 24
        local_y = y - 24
        if 0 <= local_x <= self.SEARCH_SIZE[0] and 48 <= local_y < self.SEARCH_SIZE[1]:
            index = (local_y - 48) // 28
            matches = self.search_matches()
            new_index = index if index < len(matches) else None
        else:
            new_index = None
        ranking_index = None
        if self.ranking_active and not self.ranking_minimized:
            panel_x = self.window_width - self.PANEL_SIZE[0]
            if panel_x <= x <= self.window_width and 76 <= y < 76 + 44 * 5:
                candidate = self.ranking_offset + (y - 76) // 44
                if candidate < len(self.overlaps):
                    ranking_index = candidate
        if new_index != self.search_hover_index or ranking_index != self.ranking_hover_index:
            self.search_hover_index = new_index
            self.ranking_hover_index = ranking_index
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
        pygame.draw.rect(self.surface, (34, 48, 62, 255), (0, 0, self.PANEL_SIZE[0], 34))
        self.surface.blit(
            self.small_font.render("Project details / legend", True, (235, 247, 255)),
            (14, 9),
        )
        close_label = "<" if self.panel_visible else ">"
        self.surface.blit(self.title_font.render(close_label, True, (235, 247, 255)), (self.PANEL_SIZE[0] - 30, 5))

        project = self.selected_project
        if project is None:
            self.surface.blit(
                self.title_font.render("Select a job marker", True, (235, 247, 255)),
                (18, 58),
            )
            self.surface.blit(
                self.body_font.render("Click a colored 3D point to inspect its XLSX data.", True, (205, 220, 232)),
                (18, 92),
            )
            self.surface.blit(
                self.body_font.render("WASD move | Shift: faster | Drag: orbit | Wheel: detail zoom", True, (205, 220, 232)),
                (18, 120),
            )
            self.surface.blit(
                self.body_font.render(
                    f"Overlap checker: {self.overlap_count} processed | {self.overlap_lines_rendered} lines",
                    True,
                    (120, 240, 160) if self.overlap_count == self.overlap_lines_rendered else (255, 190, 80),
                ),
                (18, 148),
            )
            self.surface.blit(
                self.small_font.render(
                    "Legend: green/orange DESC | blue/purple/pink GA ITS | cyan overlap",
                    True,
                    (150, 190, 205),
                ),
                (18, 176),
            )
            pygame.draw.circle(self.surface, (40, 210, 110), (24, 208), 6)
            pygame.draw.circle(self.surface, (255, 170, 40), (24, 232), 6)
            pygame.draw.circle(self.surface, (35, 150, 245), (24, 256), 6)
            self.surface.blit(self.small_font.render("DESC: in progress / planned", True, (205, 220, 232)), (38, 201))
            self.surface.blit(self.small_font.render("GA ITS: sponsors / utilities", True, (205, 220, 232)), (38, 225))
            self.surface.blit(self.small_font.render("Cyan lines: geographic overlap", True, (205, 220, 232)), (38, 249))
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

        y = 58
        for label, value in lines:
            label_text = self._fit_text(f"{label}: {value}", self.body_font, text_width)
            self.surface.blit(
                self.body_font.render(label_text, True, (205, 220, 232)),
                (margin, y),
            )
            y += 25
            if y > self.PANEL_SIZE[1] - 26:
                break

        hint = self.small_font.render(
            "R ranking | 1 DESC | 2 GA ITS | 3 overlap lines",
            True,
            (132, 158, 174),
        )
        self.surface.blit(hint, (margin, self.PANEL_SIZE[1] - 22))

    def _render_ranking_panel(self):
        self.ranking_surface.fill((13, 20, 29, 245))
        pygame.draw.rect(self.ranking_surface, (66, 190, 220, 255), self.ranking_surface.get_rect(), width=2)
        if self.ranking_minimized:
            pygame.draw.rect(self.ranking_surface, (34, 72, 90, 255), (0, 0, 64, 72), border_radius=4)
            arrow_font = pygame.font.Font(None, 52)
            self.ranking_surface.blit(arrow_font.render("<", True, (235, 247, 255)), (19, 10))
            return
        margin = 18
        close_rect = pygame.Rect(self.PANEL_SIZE[0] - 48, 0, 48, 48)
        pygame.draw.rect(self.ranking_surface, (34, 72, 90, 255), close_rect, border_radius=4)
        close_arrow = pygame.font.Font(None, 40).render(">", True, (235, 247, 255))
        self.ranking_surface.blit(close_arrow, close_arrow.get_rect(center=close_rect.center))
        self.ranking_surface.blit(
            self.title_font.render("Project Overlap Ranking", True, (235, 247, 255)),
            (margin, margin),
        )
        self.ranking_surface.blit(
            self.small_font.render(
                f"Showing {self.ranking_offset + 1}-{min(self.ranking_offset + 5, len(self.overlaps))} "
                f"of {len(self.overlaps)} | Wheel scrolls | R closes",
                True,
                (150, 190, 205),
            ),
            (margin, 48),
        )
        y = 76
        visible_overlaps = self.overlaps[self.ranking_offset:self.ranking_offset + 5]
        for rank, overlap in enumerate(visible_overlaps, self.ranking_offset + 1):
            first = overlap["first"]
            second = overlap["second"]
            timeline = (
                "unknown"
                if overlap["timeline_overlap"] is None
                else "yes"
                if overlap["timeline_overlap"]
                else "no"
            )
            text = (
                f"{rank}. {first['id'] or first['name']} <-> "
                f"{second['id'] or second['name']}"
            )
            detail = (
                f"{overlap['distance_km']:.2f} km | "
                f"Timeline: {timeline} | {overlap['geographic_tier']}"
            )
            row_index = rank - 1
            if rank - 1 == self.ranking_hover_index:
                pygame.draw.rect(self.ranking_surface, (45, 100, 125, 245), (0, y - 4, self.PANEL_SIZE[0], 42), border_radius=3)
            self.ranking_surface.blit(
                self.body_font.render(self._fit_text(text, self.body_font, 460), True, (225, 235, 242)),
                (margin, y),
            )
            self.ranking_surface.blit(
                self.small_font.render(self._fit_text(detail, self.small_font, 460), True, (145, 190, 205)),
                (margin + 12, y + 20),
            )
            y += 44

        self.ranking_surface.blit(
            self.small_font.render(
                f"Layers: DESC {'ON' if self.show_desc else 'OFF'} | "
                f"GA ITS {'ON' if self.show_ga_its else 'OFF'} | "
                f"Lines {'ON' if self.show_overlaps else 'OFF'}",
                True,
                (120, 240, 160),
            ),
            (margin, self.PANEL_SIZE[1] - 22),
        )

    def _upload_texture(self):
        self._render_panel()
        self._render_ranking_panel()
        self.ranking_tab_surface.fill((13, 20, 29, 245))
        pygame.draw.rect(self.ranking_tab_surface, (66, 190, 220, 255), self.ranking_tab_surface.get_rect(), width=2, border_radius=4)
        pygame.draw.rect(self.ranking_tab_surface, (34, 72, 90, 255), self.ranking_tab_surface.get_rect(), border_radius=4)
        arrow_font = pygame.font.Font(None, 52)
        arrow = arrow_font.render("<", True, (235, 247, 255))
        self.ranking_tab_surface.blit(arrow, arrow.get_rect(center=self.ranking_tab_surface.get_rect().center))
        self.panel_tab_surface.fill((13, 20, 29, 245))
        pygame.draw.rect(self.panel_tab_surface, (66, 190, 220, 255), self.panel_tab_surface.get_rect(), width=2, border_radius=4)
        pygame.draw.rect(self.panel_tab_surface, (34, 72, 90, 255), self.panel_tab_surface.get_rect(), border_radius=4)
        panel_arrow = pygame.font.Font(None, 52).render(">", True, (235, 247, 255))
        self.panel_tab_surface.blit(panel_arrow, panel_arrow.get_rect(center=self.panel_tab_surface.get_rect().center))
        self.search_surface.fill((0, 0, 0, 0))
        pygame.draw.rect(self.search_surface, (13, 20, 29, 235), (0, 0, 500, 42), border_radius=6)
        pygame.draw.rect(self.search_surface, (66, 190, 220, 255), (0, 0, 500, 42), width=2, border_radius=6)
        caret = "|" if self.search_active else ""
        search_text = self._fit_text(
            f"{self.search_query}{caret}" if self.search_query or self.search_active else "Search projects...",
            self.body_font,
            470,
        )
        text_color = (245, 250, 255) if self.search_active or self.search_query else (150, 170, 180)
        self.search_surface.blit(self.body_font.render(search_text, True, text_color), (14, 11))
        if self.search_active and self.search_query:
            for index, project in enumerate(self.search_matches()[:6]):
                row_y = 48 + index * 28
                if index == self.search_hover_index:
                    pygame.draw.rect(self.search_surface, (45, 100, 125, 245), (0, row_y, 500, 27), border_radius=3)
                else:
                    pygame.draw.rect(self.search_surface, (13, 20, 29, 235), (0, row_y, 500, 27), border_radius=3)
                row = f"{project['id'] or 'N/A'} | {project['name']} | {project['sheet']}"
                self.search_surface.blit(self.small_font.render(self._fit_text(row, self.small_font, 470), True, (235, 245, 250)), (14, row_y + 5))

        self._upload_surface(self.surface, self.texture_id)
        self._upload_surface(self.ranking_surface, self.ranking_texture_id)
        self._upload_surface(self.ranking_tab_surface, self.ranking_tab_texture_id)
        self._upload_surface(self.search_surface, self.search_texture_id)
        self._upload_surface(self.panel_tab_surface, self.panel_tab_texture_id)
        self.dirty = False

    def _upload_surface(self, surface, texture_id):
        pixels = pygame.image.tostring(surface, "RGBA", True)
        glBindTexture(GL_TEXTURE_2D, texture_id)
        glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR)
        glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR)
        glTexImage2D(
            GL_TEXTURE_2D,
            0,
            GL_RGBA,
            surface.get_width(),
            surface.get_height(),
            0,
            GL_RGBA,
            GL_UNSIGNED_BYTE,
            pixels,
        )
    def draw(self):
        if self.dirty:
            self._upload_texture()

        panel_width, panel_height = self.PANEL_SIZE

        glPushAttrib(GL_ENABLE_BIT | GL_COLOR_BUFFER_BIT | GL_DEPTH_BUFFER_BIT)
        glDisable(GL_DEPTH_TEST)
        glEnable(GL_BLEND)
        glBlendFunc(GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA)
        glEnable(GL_TEXTURE_2D)

        glMatrixMode(GL_PROJECTION)
        glPushMatrix()
        glLoadIdentity()
        glOrtho(0, self.window_width, self.window_height, 0, -1, 1)
        glMatrixMode(GL_MODELVIEW)
        glPushMatrix()
        glLoadIdentity()

        def draw_texture(texture_id, width, height, x, y):
            glBindTexture(GL_TEXTURE_2D, texture_id)
            glColor4f(1, 1, 1, 1)
            glBegin(GL_QUADS)
            glTexCoord2f(0, 1); glVertex2f(x, y)
            glTexCoord2f(1, 1); glVertex2f(x + width, y)
            glTexCoord2f(1, 0); glVertex2f(x + width, y + height)
            glTexCoord2f(0, 0); glVertex2f(x, y + height)
            glEnd()

        project_panel_y = self.window_height - panel_height - 24
        if self.panel_visible:
            draw_texture(self.texture_id, panel_width, panel_height, 0, project_panel_y)
        else:
            panel_tab_y = self.window_height - 24 - self.PANEL_TAB_SIZE[1]
            draw_texture(self.panel_tab_texture_id, *self.PANEL_TAB_SIZE, 0, panel_tab_y)
        draw_texture(self.search_texture_id, self.SEARCH_SIZE[0], self.SEARCH_SIZE[1], 24, 24)
        if self.ranking_active:
            if self.ranking_minimized:
                draw_texture(self.ranking_tab_texture_id, *self.RANKING_TAB_SIZE, self.window_width - self.RANKING_TAB_SIZE[0], self.window_height // 3 - self.RANKING_TAB_SIZE[1] // 2)
            else:
                draw_texture(self.ranking_texture_id, panel_width, panel_height, self.window_width - panel_width, 24)

        glPopMatrix()
        glMatrixMode(GL_PROJECTION)
        glPopMatrix()
        glMatrixMode(GL_MODELVIEW)
        glPopAttrib()

    def close(self):
        glDeleteTextures([self.texture_id])
        glDeleteTextures([self.ranking_texture_id, self.ranking_tab_texture_id, self.search_texture_id])
        glDeleteTextures([self.panel_tab_texture_id])
