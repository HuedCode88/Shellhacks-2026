"""
Pulls map tiles out of MongoDB (GridFS) and uploads them as PyOpenGL
textures, keyed by the same z/x/y coordinates used when they were
uploaded.

Usage:
    pip install pymongo pillow PyOpenGL numpy

    from mongo_tile_loader import TileLoader

    loader = TileLoader(connection_string, db_name="map", bucket_name="images")

    texture_id = loader.get_texture(z=12, x=3, y=4)
    if texture_id is not None:
        glBindTexture(GL_TEXTURE_2D, texture_id)
        # ... draw your quad with this texture bound ...
"""

import io

import gridfs
import numpy as np
from OpenGL.GL import (
    GL_CLAMP_TO_EDGE,
    GL_LINEAR,
    GL_RGBA,
    GL_TEXTURE_2D,
    GL_TEXTURE_MIN_FILTER,
    GL_TEXTURE_MAG_FILTER,
    GL_TEXTURE_WRAP_S,
    GL_TEXTURE_WRAP_T,
    GL_UNSIGNED_BYTE,
    glBindTexture,
    glGenTextures,
    glTexImage2D,
    glTexParameteri,
)
from PIL import Image
from pymongo import MongoClient


class TileLoader:
    """
    Fetches tiles from GridFS on demand and caches the resulting
    OpenGL texture IDs so each tile is only downloaded/decoded once
    per run, no matter how many frames re-use it.
    """

    def __init__(self, connection_string, db_name="map", bucket_name="images"):
        self.client = MongoClient(connection_string)
        self.db = self.client[db_name]
        self.fs = gridfs.GridFS(self.db, collection=bucket_name)

        # (z, x, y) -> OpenGL texture id. None means "looked up,
        # doesn't exist" so we don't hit the DB again for a missing tile.
        self._texture_cache = {}

    def close(self):
        self.client.close()

    def get_tile_bytes(self, z, x, y):
        """Returns the raw image bytes for a tile, or None if it doesn't exist."""

        grid_out = self.fs.find_one({"metadata.z": z, "metadata.x": x, "metadata.y": y})

        if grid_out is None:
            return None

        return grid_out.read()

    def get_texture(self, z, x, y):
        """
        Returns an OpenGL texture id for the given tile, uploading it
        to the GPU the first time it's requested and reusing the
        cached id on every call after that. Returns None if no tile
        exists at that coordinate.
        """

        key = (z, x, y)

        if key in self._texture_cache:
            return self._texture_cache[key]

        image_bytes = self.get_tile_bytes(z, x, y)

        if image_bytes is None:
            self._texture_cache[key] = None
            return None

        texture_id = self._upload_texture(image_bytes)
        self._texture_cache[key] = texture_id
        return texture_id

    @staticmethod
    def _upload_texture(image_bytes):
        """Decodes image bytes and uploads them as a new OpenGL texture."""

        image = Image.open(io.BytesIO(image_bytes)).convert("RGBA")

        # OpenGL expects the first row to be the bottom of the image;
        # PIL decodes top-down, so flip vertically before uploading.
        image = image.transpose(Image.FLIP_TOP_BOTTOM)

        width, height = image.size
        pixels = np.array(image, dtype=np.uint8)

        texture_id = glGenTextures(1)
        glBindTexture(GL_TEXTURE_2D, texture_id)

        glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_WRAP_S, GL_CLAMP_TO_EDGE)
        glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_WRAP_T, GL_CLAMP_TO_EDGE)
        glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR)
        glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR)

        glTexImage2D(
            GL_TEXTURE_2D, 0, GL_RGBA, width, height, 0,
            GL_RGBA, GL_UNSIGNED_BYTE, pixels,
        )

        return texture_id