# -*- coding: utf-8 -*-
"""Province-agnostic coastal economics faculty/postdoc pipeline.

Steps:
  1. Ensure province workspace (seeds + universities.json)
  2. Write 985/211/double-first-class + provincial university catalog
  3. One-by-one university discover (home / econ units / HR)
  4. Merge seeds.json (+ Python fallback) into candidates
  5. Validate + polish (dedupe http/https, drop dean/stale) + PDF

Usage:
  python run_province_pipeline.py --list
  python run_province_pipeline.py --region guangdong
  python run_province_pipeline.py --region fujian --region zhejiang
  python run_province_pipeline.py --all
  python run_province_pipeline.py --region beijing --skip-discover
  python run_province_pipeline.py --init-only
"""
from __future__ import annotations

import argparse
import json
from datetime import date

from china_universities_data import REGIONS
from download_china_universities import write_catalog
from init_province_pipeline import ensure_province_ready, init_regions
from paths import DATA_ECON, ensure_dirs
from seed_jobs import merge_seeds

TODAY = date.today().isoformat()


def run_region(
    region: str,
    *,
    skip_discover: bool = False,
    skip_validate: bool = False,
    workers: int = 6,
) -> dict:
    ensure_dirs()
    print(f"\n########## PIPELINE {region.upper()} ##########")
    out: dict = {"region": region, "date": TODAY}

    print("=== 0/4 ensure province workspace ===")
    out["init"] = ensure_province_ready(region, refresh_catalog=False)

    print("=== 1/4 university catalog ===")
    cat = write_catalog([region])
    out["catalog"] = cat["regions"].get(region)

    if not skip_discover:
        print("=== 2/4 discover by university (one-by-one) ===")
        from discover_by_university import discover_region

        out["discover"] = discover_region(region, merge_existing=True)
    else:
        print("=== 2/4 discover SKIPPED ===")

    print("=== 3/4 merge seeds ===")
    out["seeds"] = merge_seeds(region)

    if not skip_validate:
        print("=== 4/4 validate + PDF ===")
        from validate_job_links import validate_region

        out["validation"] = validate_region(region, workers=workers)
    else:
        print("=== 4/4 validate SKIPPED ===")

    return out


def main() -> None:
    ap = argparse.ArgumentParser(
        description="Coastal province faculty/postdoc pipeline (13 regions)."
    )
    ap.add_argument(
        "--region",
        action="append",
        choices=list(REGIONS),
        help="Region folder key (repeatable). Choices: %s" % ", ".join(REGIONS),
    )
    ap.add_argument(
        "--all",
        action="store_true",
        help="Run all coastal regions: %s" % ", ".join(REGIONS),
    )
    ap.add_argument(
        "--list",
        action="store_true",
        help="Print available regions and exit.",
    )
    ap.add_argument(
        "--init-only",
        action="store_true",
        help="Only scaffold seeds/universities for selected regions (no scrape).",
    )
    ap.add_argument(
        "--skip-discover",
        action="store_true",
        help="Skip university loop; only seed + validate existing candidates.",
    )
    ap.add_argument(
        "--skip-validate",
        action="store_true",
        help="Stop after discover + seeds (no fetch scoring).",
    )
    ap.add_argument("--workers", type=int, default=6)
    args = ap.parse_args()

    if args.list:
        print("Available regions (%d):" % len(REGIONS))
        for r in REGIONS:
            print(" ", r)
        return

    targets = list(REGIONS) if args.all else (args.region or [])
    if not targets:
        if args.init_only:
            targets = list(REGIONS)
        else:
            raise SystemExit(
                "Pass --region <name> (repeatable), --all, --list, or --init-only.\n"
                "Regions: %s" % ", ".join(REGIONS)
            )

    if args.init_only:
        init_regions(targets)
        return

    summary = {"date": TODAY, "regions": {}}
    for region in targets:
        summary["regions"][region] = run_region(
            region,
            skip_discover=args.skip_discover,
            skip_validate=args.skip_validate,
            workers=args.workers,
        )

    out_path = DATA_ECON / ("province_pipeline_summary_%s.json" % TODAY)
    out_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    print("Summary -> %s" % out_path)
    print("Done.")


if __name__ == "__main__":
    main()
