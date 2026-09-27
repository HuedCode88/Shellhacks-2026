"""
Bulk-uploads a map tile pyramid into MongoDB using GridFS.

Expects the classic tile-server folder layout:

    map_tiles/
        {z}/            <- zoom level (numbered folder)
            {x}/        <- column (numbered folder)
                {y}.png <- row (image file, named after a number)

Target: db "map", GridFS bucket "images"
  -> creates/uses collections: images.files, images.chunks
  -> each file's metadata stores {z, x, y} so tiles can be queried
     back out by coordinate later.

Usage:
    pip install pymongo
    python upload_images_to_mongo.py
"""

import mimetypes
import os
import sys
import threading
from concurrent.futures import ThreadPoolExecutor

from pymongo import MongoClient
from pymongo.errors import PyMongoError
import gridfs
from dotenv import load_dotenv


# ============================================================
# CONFIG -- edit these
# ============================================================
load_dotenv()
# Your MongoDB connection string. For Atlas, grab this from
# Cluster > Connect > "Drivers", or reuse the one in your
# MongoDB for VS Code connection.
CONNECTION_STRING = os.getenv("MONGODB_URI")

DB_NAME = "map"
BUCKET_NAME = "images"

# Top-level folder containing the {z}/{x}/{y}.ext tile pyramid.
IMAGE_FOLDER = "C:/Users/Speci/Downloads/map_tiles/map_tiles"

# Only files with these extensions will be uploaded.
ALLOWED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp"}

# If a tile with the same z/x/y already exists, skip it instead of
# uploading a duplicate. Set to False to always upload (creates a
# second copy in GridFS rather than overwriting).
SKIP_DUPLICATES = True

# How many uploads to run at once. Each upload is a blocking network
# call, so running several in parallel overlaps that latency instead
# of paying it one file at a time. 16 is a reasonable default; raise
# it if your connection has headroom, lower it if you see connection
# errors.
UPLOAD_WORKERS = 16


# ============================================================
# UPLOAD LOGIC
# ============================================================

def find_tile_files(folder):
    """
    Walks map_tiles/{z}/{x}/{y}.ext and returns a list of
    (full_path, z, x, y) tuples. z and x come from folder names,
    y comes from the image filename (without extension).

    Any non-numeric folder or file name is skipped with a warning
    rather than crashing the whole run.
    """

    if not os.path.isdir(folder):
        raise FileNotFoundError(f"Folder not found: {folder}")

    tiles = []

    for z_name in sorted(os.listdir(folder)):
        z_path = os.path.join(folder, z_name)

        if not os.path.isdir(z_path) or not z_name.isdigit():
            continue

        for x_name in sorted(os.listdir(z_path)):
            x_path = os.path.join(z_path, x_name)

            if not os.path.isdir(x_path) or not x_name.isdigit():
                continue

            for entry in sorted(os.listdir(x_path)):
                full_path = os.path.join(x_path, entry)

                if not os.path.isfile(full_path):
                    continue

                y_name, ext = os.path.splitext(entry)

                if ext.lower() not in ALLOWED_EXTENSIONS:
                    continue

                if not y_name.isdigit():
                    print(f"  warning: unexpected filename, skipping: {full_path}")
                    continue

                tiles.append((full_path, int(z_name), int(x_name), int(y_name)))

    return tiles


def upload_folder(fs, db, folder):
    tiles = find_tile_files(folder)

    if not tiles:
        print(f"No tile images found under {folder}")
        return

    print(f"Found {len(tiles)} tile(s) to upload.\n")

    files_collection = db[f"{BUCKET_NAME}.files"]

    # Without this, fs.exists() below would collection-scan
    # images.files on every single file -- and that scan gets slower
    # as the collection grows mid-run. create_index is a cheap no-op
    # if it already exists.
    files_collection.create_index([("metadata.z", 1), ("metadata.x", 1), ("metadata.y", 1)])

    existing_coordinates = set()
    if SKIP_DUPLICATES:
        print("Checking for already-uploaded tiles...")
        cursor = files_collection.find({}, {"metadata.z": 1, "metadata.x": 1, "metadata.y": 1})
        for doc in cursor:
            metadata = doc.get("metadata") or {}
            if "z" in metadata and "x" in metadata and "y" in metadata:
                existing_coordinates.add((metadata["z"], metadata["x"], metadata["y"]))
        print(f"  {len(existing_coordinates)} tile(s) already in the database.\n")

    counts_lock = threading.Lock()
    counts = {"uploaded": 0, "skipped": 0, "failed": 0}

    def upload_one(tile):
        path, z, x, y = tile

        if SKIP_DUPLICATES and (z, x, y) in existing_coordinates:
            with counts_lock:
                counts["skipped"] += 1
            return

        ext = os.path.splitext(path)[1]
        filename = f"{z}_{x}_{y}{ext}"

        content_type, _ = mimetypes.guess_type(path)
        content_type = content_type or "application/octet-stream"

        try:
            with open(path, "rb") as f:
                fs.put(
                    f,
                    filename=filename,
                    contentType=content_type,
                    metadata={
                        "z": z,
                        "x": x,
                        "y": y,
                        "originalPath": path,
                        "sizeBytes": os.path.getsize(path),
                    },
                )
            with counts_lock:
                counts["uploaded"] += 1
        except (PyMongoError, OSError) as exc:
            print(f"  FAILED: z={z} x={x} y={y}  ({exc})")
            with counts_lock:
                counts["failed"] += 1

        done = counts["uploaded"] + counts["skipped"] + counts["failed"]
        if done % 200 == 0:
            print(f"  ...{done}/{len(tiles)} processed")

    with ThreadPoolExecutor(max_workers=UPLOAD_WORKERS) as executor:
        list(executor.map(upload_one, tiles))

    print()
    print(f"Done. Uploaded: {counts['uploaded']}  Skipped: {counts['skipped']}  Failed: {counts['failed']}")


def main():
    if "your/map_tiles" in IMAGE_FOLDER or "<" in CONNECTION_STRING:
        print("Edit CONNECTION_STRING and IMAGE_FOLDER at the top of this script before running.")
        sys.exit(1)

    # maxPoolSize needs to be at least UPLOAD_WORKERS, or threads will
    # queue up waiting for a free connection instead of actually
    # running in parallel.
    client = MongoClient(CONNECTION_STRING, maxPoolSize=UPLOAD_WORKERS + 4)
    db = client[DB_NAME]
    fs = gridfs.GridFS(db, collection=BUCKET_NAME)

    try:
        upload_folder(fs, db, IMAGE_FOLDER)
    finally:
        client.close()


if __name__ == "__main__":
    main()