"""
One command that runs the full geocoding pipeline:

  1. search_overpass    -- fetch DESC / GA Power infrastructure from OSM
  2. geocode_match       -- fuzzy-match project titles against it
  3. nominatim_fallback  -- (optional) fill in anything still unmatched

Produces one final workbook, defaulting to the exact filename
config.EXCEL_FILE already expects ("gridlock_project_tables_geocoded_nominatim
(1).xlsx"), containing "DESC Geocoded" and "GA ITS Geocoded" sheets with a
center_lat / center_lon per project. Those column names match the
"centerlat" / "centerlon" aliases in config.FIELD_ALIASES, so
database.load_workbook_projects() picks them up with no further wiring --
just run this and refresh the app.

CLI
---
  python3 -m geocoding.run_pipeline \
      --excel gridlock_project_tables_raw.xlsx \
      --out gridlock_project_tables_geocoded.xlsx \
      --contact-email you@example.com \
      --nominatim-fallback

  # Reuse GeoJSON already fetched in a previous run:
  python3 -m geocoding.run_pipeline --skip-query ...
"""

import argparse
from pathlib import Path

from . import search_overpass, nominatim_fallback


def run_full_pipeline(excel, out, desc_geojson=None, ga_geojson=None,
                       contact_email="", skip_query=False,
                       nominatim_pass=False):
    """Runs stage 1+2 (search_overpass.run_pipeline, which itself calls
    geocode_match.run), then optionally re-runs stage 3 as a second pass
    over anything still unmatched. Returns the final output path."""
    script_dir = Path(__file__).resolve().parent
    desc_geojson = desc_geojson or str(script_dir / "dominion_energy_sc_all.geojson")
    ga_geojson = ga_geojson or str(script_dir / "georgia_power_all.geojson")

    stage_2_out = search_overpass.run_pipeline(
        excel=excel,
        out=out,
        desc_geojson=desc_geojson,
        ga_geojson=ga_geojson,
        contact_email=contact_email,
        skip_query=skip_query,
        # Nominatim runs inline during stage 2 whenever a project has no
        # Overpass match at all; this covers most cases in one pass.
        nominatim_fallback=nominatim_pass,
    )

    if not nominatim_pass:
        return stage_2_out

    # Second pass: catch anything still left over (e.g. Overpass features
    # that loaded but a rate-limited/failed Nominatim call skipped first
    # time around).
    return nominatim_fallback.run(
        excel=stage_2_out,
        out=stage_2_out,
        contact_email=contact_email,
    )


def main():
    script_dir = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--excel", default=str(script_dir / "gridlock_project_tables_raw.xlsx"),
                         help="Part 1 output workbook (raw project tables)")
    parser.add_argument(
        "--out",
        default=str(script_dir.parent / "gridlock_project_tables_geocoded_nominatim (1).xlsx"),
        help="Final geocoded workbook. Defaults to the exact filename AND location "
             "config.EXCEL_FILE already points at (the app's root directory, one "
             "level up from this geocoding/ package), so the app picks it up with "
             "no further changes.",
    )
    parser.add_argument("--desc-geojson", default=None)
    parser.add_argument("--ga-geojson", default=None)
    parser.add_argument("--contact-email", default="",
                         help="Required if --nominatim-fallback is set (Nominatim usage policy)")
    parser.add_argument("--skip-query", action="store_true",
                         help="Reuse existing GeoJSON files instead of re-hitting Overpass")
    parser.add_argument("--nominatim-fallback", action="store_true",
                         help="Fill in anything Overpass couldn't match via the public Nominatim API")
    args = parser.parse_args()

    if args.nominatim_fallback and not args.contact_email:
        parser.error("--nominatim-fallback requires --contact-email")

    run_full_pipeline(
        excel=args.excel,
        out=args.out,
        desc_geojson=args.desc_geojson,
        ga_geojson=args.ga_geojson,
        contact_email=args.contact_email,
        skip_query=args.skip_query,
        nominatim_pass=args.nominatim_fallback,
    )


if __name__ == "__main__":
    main()