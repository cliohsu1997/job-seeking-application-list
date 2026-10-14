# -*- coding: utf-8 -*-
"""Prepare data/economics/{region}/ workspaces for every coastal province.

Creates region folders, empty seeds.json (if missing), candidate stubs, and
refreshes universities.json from provincial_catalog.json.

Usage:
  python init_province_pipeline.py
  python init_province_pipeline.py --region beijing --region fujian
  python init_province_pipeline.py --all
"""
from __future__ import annotations

import argparse
import json
from datetime import date

from china_universities_data import REGIONS
from download_china_universities import write_catalog
from paths import DATA_ECON, ensure_dirs

TODAY = date.today().isoformat()


def empty_seeds_payload(region: str) -> dict:
    return {
        "region": region,
        "note": "Add known-good announcement URLs; pipeline merges then validates",
        "count": 0,
        "seeds": [],
    }


def ensure_province_ready(region: str, *, refresh_catalog: bool = True) -> dict:
    """Make one province runnable by the coastal pipeline."""
    ensure_dirs()
    base = DATA_ECON / region
    base.mkdir(parents=True, exist_ok=True)
    (base / "raw").mkdir(parents=True, exist_ok=True)

    seeds_path = base / "seeds.json"
    created_seeds = False
    if not seeds_path.exists():
        seeds_path.write_text(
            json.dumps(empty_seeds_payload(region), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        created_seeds = True

    candidates_path = base / "discovered_jobs_candidates.json"
    created_candidates = False
    if not candidates_path.exists():
        candidates_path.write_text("[]\n", encoding="utf-8")
        created_candidates = True

    catalog_info = None
    if refresh_catalog:
        catalog_info = write_catalog([region])["regions"].get(region)

    return {
        "region": region,
        "seeds_created": created_seeds,
        "candidates_created": created_candidates,
        "catalog": catalog_info,
        "universities": (base / "universities.json").exists(),
        "path": str(base),
    }


def init_regions(regions: list[str] | None = None) -> dict:
    targets = list(regions or REGIONS)
    write_catalog(targets)
    summary = {"date": TODAY, "regions": {}}
    for region in targets:
        info = ensure_province_ready(region, refresh_catalog=False)
        info["universities"] = (DATA_ECON / region / "universities.json").exists()
        summary["regions"][region] = info
        print(
            "ready %s: universities=%s seeds=%s candidates=%s"
            % (
                region,
                "ok" if info["universities"] else "missing",
                "new" if info["seeds_created"] else "kept",
                "new" if info["candidates_created"] else "kept",
            )
        )
    out = DATA_ECON / ("province_init_summary_%s.json" % TODAY)
    out.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print("Summary ->", out)
    return summary


def main() -> None:
    ap = argparse.ArgumentParser(description="Init coastal province pipeline folders")
    ap.add_argument("--region", action="append", choices=list(REGIONS))
    ap.add_argument("--all", action="store_true", help="Init every coastal region")
    args = ap.parse_args()
    targets = list(REGIONS) if args.all or not args.region else args.region
    init_regions(targets)
    print("Done.")


if __name__ == "__main__":
    main()
