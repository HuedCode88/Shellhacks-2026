"""Preload OpenStreetMap tiles around every geocoded project.

Examples:
    python preload_point_tiles.py --dry-run
    python preload_point_tiles.py --radius 3 --min-zoom 5 --max-zoom 17
"""

import argparse
import multiprocessing
import os
import queue
import time
from concurrent.futures import ProcessPoolExecutor, ThreadPoolExecutor, as_completed

import requests

from main import (
    EXCEL_FILE,
    TILE_CACHE_DIR,
    TILE_ZOOM,
    lat_to_tile_y,
    load_projects,
    lon_to_tile_x,
)

OSM_USER_AGENT = "Shellhacks-2026-point-tile-preloader/1.0 (local visualization)"
TILE_URL = "https://tile.openstreetmap.org/{zoom}/{x}/{y}.png"


def build_tile_manifest(projects, min_zoom, max_zoom, radius):
    """Return unique (zoom, x, y) tiles surrounding every project."""
    manifest = set()
    for zoom in range(min_zoom, max_zoom + 1):
        tile_count = 2 ** zoom
        for project in projects:
            center_x = int(lon_to_tile_x(project["lon"], zoom))
            center_y = int(lat_to_tile_y(project["lat"], zoom))
            for tile_x in range(center_x - radius, center_x + radius + 1):
                for tile_y in range(center_y - radius, center_y + radius + 1):
                    if 0 <= tile_x < tile_count and 0 <= tile_y < tile_count:
                        manifest.add((zoom, tile_x, tile_y))
    return sorted(manifest)


def tile_path(tile):
    zoom, tile_x, tile_y = tile
    return os.path.join(TILE_CACHE_DIR, str(zoom), str(tile_x), f"{tile_y}.png")


def download_tile(tile):
    path = tile_path(tile)
    if os.path.isfile(path):
        return "cached", tile

    zoom, tile_x, tile_y = tile
    try:
        response = requests.get(
            TILE_URL.format(zoom=zoom, x=tile_x, y=tile_y),
            headers={"User-Agent": OSM_USER_AGENT},
            timeout=20,
        )
        response.raise_for_status()
        os.makedirs(os.path.dirname(path), exist_ok=True)
        temporary_path = f"{path}.tmp"
        with open(temporary_path, "wb") as tile_file:
            tile_file.write(response.content)
        os.replace(temporary_path, path)
        return "downloaded", tile
    except requests.RequestException as error:
        return f"failed: {error}", tile


def download_shard(tiles, workers, shard_id, progress_queue):
    """Download one manifest shard in a separate process."""
    completed = 0
    failures = 0
    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = [executor.submit(download_tile, tile) for tile in tiles]
        for future in as_completed(futures):
            status, _ = future.result()
            completed += 1
            if status.startswith("failed"):
                failures += 1
            progress_queue.put((shard_id, completed, status))
    return completed, failures


def run_download_with_ui(missing, processes, workers, cached_count):
    """Run the parallel downloader with a native Tk progress window."""
    import tkinter as tk
    from tkinter import ttk

    total = len(missing)
    manager = multiprocessing.Manager()
    progress_queue = manager.Queue()
    shards = [missing[index::processes] for index in range(processes)]
    shard_totals = [len(shard) for shard in shards]
    shard_progress = [0] * len(shards)
    shard_failures = [0] * len(shards)
    executor = ProcessPoolExecutor(max_workers=processes)
    futures = [
        executor.submit(download_shard, shard, workers, shard_id, progress_queue)
        for shard_id, shard in enumerate(shards)
        if shard
    ]
    start = time.perf_counter()

    root = tk.Tk()
    root.title("Gridlock Map Tile Preloader")
    root.geometry("620x300")
    root.resizable(False, False)
    ttk.Label(root, text="Downloading map tiles", font=("Segoe UI", 15, "bold")).pack(pady=(16, 4))
    ttk.Label(root, text=f"Cached: {cached_count:,} | Remaining: {total:,}").pack()
    progress = ttk.Progressbar(root, maximum=total, length=560, mode="determinate")
    progress.pack(pady=14)
    status = ttk.Label(root, text="Starting download processes...")
    status.pack()
    process_status = ttk.Label(root, text="")
    process_status.pack(pady=8)
    cancel_button = ttk.Button(root, text="Cancel")
    cancel_button.pack()
    state = {"completed": 0, "finished": False}

    def finish(message):
        if state["finished"]:
            return
        state["finished"] = True
        status.configure(text=message)
        cancel_button.configure(text="Close", command=close_window)
        executor.shutdown(wait=False, cancel_futures=True)

    def close_window():
        manager.shutdown()
        root.destroy()

    def cancel():
        finish("Cancelled. Completed tiles remain cached.")

    cancel_button.configure(command=cancel)
    root.protocol("WM_DELETE_WINDOW", close_window)

    def poll_progress():
        while True:
            try:
                shard_id, shard_count, result_status = progress_queue.get_nowait()
            except queue.Empty:
                break
            state["completed"] += 1
            shard_progress[shard_id] = shard_count
            if result_status.startswith("failed"):
                shard_failures[shard_id] += 1

        completed = state["completed"]
        elapsed = max(time.perf_counter() - start, 0.001)
        rate = completed / elapsed
        remaining = max(total - completed, 0)
        eta_minutes = remaining / rate / 60.0 if rate else 0.0
        progress.configure(value=completed)
        status.configure(
            text=f"{completed:,}/{total:,} tiles | {rate:.1f} tiles/sec | ETA {eta_minutes:.1f} min"
        )
        process_status.configure(
            text="  ".join(
                f"P{index + 1}: {count:,}/{shard_totals[index]:,}"
                for index, count in enumerate(shard_progress)
                if shard_totals[index]
            )
        )

        if completed >= total or all(future.done() for future in futures):
            failures = sum(shard_failures)
            finish(f"Finished. Failures: {failures}. Cached tiles are ready to use.")
            return
        root.after(200, poll_progress)

    root.after(200, poll_progress)
    root.mainloop()


def main():
    parser = argparse.ArgumentParser(
        description="Download OSM tiles around every geocoded project."
    )
    parser.add_argument("--min-zoom", type=int, default=5)
    parser.add_argument(
        "--max-zoom",
        type=int,
        default=17,
        help="Maximum OSM detail level; 17 is suitable for close street context.",
    )
    parser.add_argument(
        "--radius",
        type=int,
        default=3,
        help="Tiles in each direction; radius 3 means a 7x7 area around each point.",
    )
    parser.add_argument(
        "--processes",
        type=int,
        default=4,
        help="Independent downloader processes used for separate tile shards.",
    )
    parser.add_argument(
        "--workers",
        type=int,
        default=4,
        help="Threads per process; total default concurrency is 16 downloads.",
    )
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument(
        "--no-ui",
        action="store_true",
        help="Use terminal progress instead of opening the desktop progress window.",
    )
    args = parser.parse_args()

    if not 1 <= args.min_zoom <= args.max_zoom <= 19:
        parser.error("zoom levels must satisfy 1 <= min-zoom <= max-zoom <= 19")
    if args.radius < 0:
        parser.error("radius must be zero or greater")
    if not 1 <= args.processes <= 8:
        parser.error("processes must be between 1 and 8")
    if not 1 <= args.workers <= 8:
        parser.error("workers must be between 1 and 8")
    if args.processes * args.workers > 32:
        parser.error("processes * workers must not exceed 32 total downloads")

    projects = load_projects(EXCEL_FILE)
    manifest = build_tile_manifest(
        projects,
        args.min_zoom,
        args.max_zoom,
        args.radius,
    )
    missing = [tile for tile in manifest if not os.path.isfile(tile_path(tile))]
    print(f"Projects: {len(projects)}")
    print(f"Tile manifest: {len(manifest)} unique tiles")
    print(f"Already cached: {len(manifest) - len(missing)}")
    print(f"To download: {len(missing)}")
    print(f"Cache directory: {TILE_CACHE_DIR}")

    if args.dry_run or not missing:
        return

    if not args.no_ui:
        run_download_with_ui(
            missing,
            args.processes,
            args.workers,
            len(manifest) - len(missing),
        )
        return

    start = time.perf_counter()
    shards = [missing[index::args.processes] for index in range(args.processes)]
    shard_totals = [len(shard) for shard in shards]
    shard_progress = [0] * len(shards)
    shard_failures = [0] * len(shards)
    print(
        f"Starting {args.processes} downloader processes with "
        f"{args.workers} threads each."
    )
    with multiprocessing.Manager() as manager:
        progress_queue = manager.Queue()
        with ProcessPoolExecutor(max_workers=args.processes) as executor:
            futures = [
                executor.submit(
                    download_shard,
                    shard,
                    args.workers,
                    shard_id,
                    progress_queue,
                )
                for shard_id, shard in enumerate(shards)
                if shard
            ]
            completed = 0
            last_report = 0.0
            while completed < len(missing):
                try:
                    shard_id, shard_count, status = progress_queue.get(timeout=0.5)
                    shard_progress[shard_id] = shard_count
                    completed += 1
                    if status.startswith("failed"):
                        shard_failures[shard_id] += 1
                except queue.Empty:
                    if all(future.done() for future in futures):
                        break
                    continue

                now = time.perf_counter()
                if now - last_report >= 1.0 or completed == len(missing):
                    elapsed = max(now - start, 0.001)
                    rate = completed / elapsed
                    remaining = max(len(missing) - completed, 0)
                    eta_seconds = remaining / rate if rate else 0
                    eta_minutes = eta_seconds / 60.0
                    process_status = " ".join(
                        f"P{index + 1}:{count}/{shard_totals[index]}"
                        for index, count in enumerate(shard_progress)
                        if shard_totals[index]
                    )
                    print(
                        f"Progress {completed}/{len(missing)} "
                        f"({completed / len(missing) * 100:.1f}%) | "
                        f"{rate:.1f} tiles/sec | ETA {eta_minutes:.1f} min | "
                        f"{process_status}",
                        flush=True,
                    )
                    last_report = now

            results = [future.result() for future in futures]

    completed = sum(item[0] for item in results)
    failures = sum(item[1] for item in results)

    print(f"Finished in {time.perf_counter() - start:.1f} seconds.")
    print(f"Downloaded: {completed - failures}; failures: {failures}")


if __name__ == "__main__":
    main()
