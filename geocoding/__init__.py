"""
geocoding
=========

Three-stage pipeline that turns raw project titles (from the Part 1
workbook) into real-world coordinates:

  1. search_overpass   -- pull utility infrastructure from OpenStreetMap
                           (network, no auth needed)
  2. geocode_match      -- fuzzy-match project titles against that data
                           (offline)
  3. nominatim_fallback -- fill in anything unmatched via the public
                           Nominatim API (network, rate-limited)

Use `geocoding.run_pipeline.run_full_pipeline(...)` to run all three in
one call, or import each stage's `run(...)` function to use them
individually.
"""