from config import OVERPASS_BBOXES, OVERPASS_CACHE_MAX_AGE_DAYS
from database import initialize_database, load_all_projects
from geocoding.live_lookup import warm_cache
from server import serve_map

def main():
    initialize_database()
    projects = load_all_projects()
    print(f"Loaded {len(projects)} geocoded projects.")

    # Pre-fetch/refresh the Overpass feature cache now, so the first
    # document a user uploads doesn't have to wait on a multi-minute
    # Overpass query before it can be geocoded.
    warm_cache(OVERPASS_BBOXES, max_age_days=OVERPASS_CACHE_MAX_AGE_DAYS)

    serve_map()

if __name__ == "__main__":
    main()