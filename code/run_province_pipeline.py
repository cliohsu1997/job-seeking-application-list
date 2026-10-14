# -*- coding: utf-8 -*-
"""Province-agnostic coastal economics faculty/postdoc pipeline.

Steps:
  1. Write 985/211/double-first-class + provincial university catalog
  2. One-by-one university discover (home / econ units / HR)
  3. Merge seeds.json (+ Python fallback) into candidates
  4. Validate + polish (dedupe http/https, drop dean/stale) + PDF

Usage:
  python run_province_pipeline.py --region guangdong
  python run_province_pipeline.py --region fujian --region zhejiang
  python run_province_pipeline.py --all
  python run_province_pipeline.py --region guangdong --skip-discover
"""
from __future__ import annotations

import argparse
import json
from datetime import date

from china_universities_data import REGIONS
from download_china_universities import write_catalog
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
    ap = argparse.ArgumentParser(description="Coastal province faculty/postdoc pipeline")
    ap.add_argument(
        "--region",
        action="append",
        choices=list(REGIONS),
        help="Region folder key (repeatable).",
    )
    ap.add_argument("--all", action="store_true", help="Run all four coastal regions.")
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

    targets = list(REGIONS) if args.all else (args.region or [])
    if not targets:
        raise SystemExit("Pass --region guangdong (or --all)")

    summary = {"date": TODAY, "regions": {}}
    for region in targets:
        summary["regions"][region] = run_region(
            region,
            skip_discover=args.skip_discover,
            skip_validate=args.skip_validate,
            workers=args.workers,
        )

    out_path = DATA_ECON / f"province_pipeline_summary_{TODAY}.json"
    out_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    print(f"Summary -> {out_path}")
    print("Done.")


if __name__ == "__main__":
    main()
