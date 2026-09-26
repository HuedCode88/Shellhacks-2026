import pygame
from pygame.locals import *
from OpenGL.GL import *
from OpenGL.GLU import *


# --------------------------------------------------
# Your 3D coordinates
# --------------------------------------------------

POINTS = [
    (0, 0, 0),
    (1, 2, 3),
    (2, 1, 4),
    (-2, 1, 2),
    (-1, -2, 1),
    (3, 0, -2),
]


# --------------------------------------------------
# Camera settings
# --------------------------------------------------

camera_distance = 15.0
camera_rotation_x = 25.0
camera_rotation_y = 45.0

mouse_down = False
last_mouse_pos = (0, 0)


# --------------------------------------------------
# OpenGL setup
# --------------------------------------------------

def setup_opengl(width, height):
    glViewport(0, 0, width, height)

    glMatrixMode(GL_PROJECTION)
    glLoadIdentity()

    gluPerspective(
        45,
        width / height,
        0.1,
        1000.0
    )

    glMatrixMode(GL_MODELVIEW)

    glEnable(GL_DEPTH_TEST)

    # Make points larger
    glPointSize(8)

    # Smooth lines
    glEnable(GL_LINE_SMOOTH)


# --------------------------------------------------
# Draw X/Y/Z axes
# --------------------------------------------------

def draw_axes(length=10):
    glLineWidth(2)

    glBegin(GL_LINES)

    # X axis - RED
    glColor3f(1, 0, 0)
    glVertex3f(0, 0, 0)
    glVertex3f(length, 0, 0)

    # Y axis - GREEN
    glColor3f(0, 1, 0)
    glVertex3f(0, 0, 0)
    glVertex3f(0, length, 0)

    # Z axis - BLUE
    glColor3f(0, 0, 1)
    glVertex3f(0, 0, 0)
    glVertex3f(0, 0, length)

    glEnd()


# --------------------------------------------------
# Draw a grid on the X/Z plane
# --------------------------------------------------

def draw_grid(size=10, spacing=1):
    glColor3f(0.25, 0.25, 0.25)
    glLineWidth(1)

    glBegin(GL_LINES)

    for i in range(-size, size + 1, spacing):

        # Lines parallel to X
        glVertex3f(-size, 0, i)
        glVertex3f(size, 0, i)

        # Lines parallel to Z
        glVertex3f(i, 0, -size)
        glVertex3f(i, 0, size)

    glEnd()


# --------------------------------------------------
# Draw your points
# --------------------------------------------------

def draw_points(points):
    glPointSize(8)

    glBegin(GL_POINTS)

    for x, y, z in points:
        # Yellow points
        glColor3f(1, 1, 0)

        glVertex3f(x, y, z)

    glEnd()


# --------------------------------------------------
# Draw a small sphere at each point
# --------------------------------------------------

def draw_points_as_spheres(points):
    quadric = gluNewQuadric()

    for x, y, z in points:

        glPushMatrix()

        glTranslatef(x, y, z)

        glColor3f(1, 1, 0)

        gluSphere(
            quadric,
            0.1,   # radius
            12,     # slices
            12      # stacks
        )

        glPopMatrix()

    gluDeleteQuadric(quadric)


# --------------------------------------------------
# Handle camera
# --------------------------------------------------

def update_camera():
    glLoadIdentity()

    # Move camera backwards
    glTranslatef(
        0,
        0,
        -camera_distance
    )

    # Rotate camera
    glRotatef(camera_rotation_x, 1, 0, 0)
    glRotatef(camera_rotation_y, 0, 1, 0)


# --------------------------------------------------
# Main program
# --------------------------------------------------

def main():

    global camera_distance
    global camera_rotation_x
    global camera_rotation_y

    global mouse_down
    global last_mouse_pos

    pygame.init()

    width = 1000
    height = 700

    pygame.display.set_mode(
        (width, height),
        DOUBLEBUF | OPENGL
    )

    pygame.display.set_caption(
        "3D Coordinate Viewer"
    )

    setup_opengl(width, height)

    clock = pygame.time.Clock()

    running = True

    while running:

        # ------------------------------------------
        # Events
        # ------------------------------------------

        for event in pygame.event.get():

            if event.type == pygame.QUIT:
                running = False

            # Mouse button pressed
            elif event.type == pygame.MOUSEBUTTONDOWN:

                if event.button == 1:
                    mouse_down = True
                    last_mouse_pos = event.pos

                # Mouse wheel up
                elif event.button == 4:
                    camera_distance -= 1

                # Mouse wheel down
                elif event.button == 5:
                    camera_distance += 1

            # Mouse button released
            elif event.type == pygame.MOUSEBUTTONUP:

                if event.button == 1:
                    mouse_down = False

            # Mouse movement
            elif event.type == pygame.MOUSEMOTION:

                if mouse_down:

                    x, y = event.pos
                    old_x, old_y = last_mouse_pos

                    dx = x - old_x
                    dy = y - old_y

                    camera_rotation_y += dx * 0.5
                    camera_rotation_x += dy * 0.5

                    last_mouse_pos = event.pos

        # Prevent camera from going inside the points
        camera_distance = max(
            1.0,
            min(camera_distance, 100.0)
        )

        # ------------------------------------------
        # Clear screen
        # ------------------------------------------

        glClear(
            GL_COLOR_BUFFER_BIT |
            GL_DEPTH_BUFFER_BIT
        )

        # ------------------------------------------
        # Camera
        # ------------------------------------------

        update_camera()

        # ------------------------------------------
        # Draw scene
        # ------------------------------------------

        draw_grid()
        draw_axes()

        draw_points(POINTS)

        # Or use this instead:
        # draw_points_as_spheres(POINTS)

        # ------------------------------------------
        # Display
        # ------------------------------------------

        pygame.display.flip()

        clock.tick(60)

    pygame.quit()


if __name__ == "__main__":
    main()
