import pygame
from OpenGL.GL import *


class ProjectInfoPanel:
    """Draws selected workbook project data as an in-window OpenGL overlay."""

    PANEL_SIZE = (440, 272)

    def __init__(self, window_size):
        self.window_width, self.window_height = window_size
        self.surface = pygame.Surface(self.PANEL_SIZE, pygame.SRCALPHA)
        self.title_font = pygame.font.Font(None, 25)
        self.body_font = pygame.font.Font(None, 20)
        self.small_font = pygame.font.Font(None, 17)
        self.texture_id = glGenTextures(1)
        self.selected_project = None
        self.dirty = True

    def set_project(self, project):
        if project is not self.selected_project:
            self.selected_project = project
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
            return

        margin = 18
        text_width = self.PANEL_SIZE[0] - margin * 2
        title = self._fit_text(project["name"], self.title_font, text_width)
        self.surface.blit(
            self.title_font.render(title, True, (235, 247, 255)),
            (margin, margin),
        )

        lines = [
            ("Source", project["sheet"]),
            ("ID", project["id"] or "N/A"),
            ("Category", project["category"]),
            ("Coordinates", f"{project['lat']:.5f}, {project['lon']:.5f}"),
        ]
        details = project.get("details", {})
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
        ):
            if field in details:
                lines.append((field, details[field]))

        y = 52
        for label, value in lines:
            label_text = self._fit_text(f"{label}: {value}", self.body_font, text_width)
            self.surface.blit(
                self.body_font.render(label_text, True, (205, 220, 232)),
                (margin, y),
            )
            y += 25
            if y > self.PANEL_SIZE[1] - 26:
                break

        hint = self.small_font.render("Click another marker to inspect its job", True, (132, 158, 174))
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
        if self.selected_project is None:
            return
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
