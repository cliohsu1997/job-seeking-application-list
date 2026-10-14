# -*- coding: utf-8 -*-
"""Province-agnostic coastal economics faculty/postdoc pipeline.

Steps per region:
  1. Ensure province workspace (seeds + universities.json)
  2. Write university catalog from provincial_catalog.json
  3. One-by-one university discover (home / econ units / HR)
  4. Merge seeds.json (+ Python fallback) into candidates
  5. Validate + polish + PDF
  6. Screen freshness (2026+ OR 2025 with rolling/open-until-filled)
  7. One-by-one verify each posting is a real faculty/postdoc ad
  8. (after all regions) summarize into coastal PDF

Usage:
  python run_province_pipeline.py --list
  python run_province_pipeline.py --region guangdong
  python run_province_pipeline.py --all --continue-on-error
  python run_province_pipeline.py --region beijing --skip-discover
  python run_province_pipeline.py --init-only
  python run_province_pipeline.py --summarize-only
"""
from __future__ import annotations

import argparse
import json
import traceback
from datetime import date

from china_universities_data import REGIONS
from download_china_universities import write_catalog
from init_province_pipeline import ensure_province_ready, init_regions
from paths import DATA_ECON, ensure_dirs
from seed_jobs import merge_seeds
from summarize_coastal import summarize as summarize_coastal

TODAY = date.today().isoformat()


def run_region(
    region: str,
    *,
    skip_discover: bool = False,
    skip_validate: bool = False,
    skip_screen: bool = False,
    skip_verify: bool = False,
    workers: int = 6,
) -> dict:
    ensure_dirs()
    print("\n########## PIPELINE %s ##########" % region.upper())
    out: dict = {"region": region, "date": TODAY, "ok": True}

    print("=== 1/7 ensure province workspace ===")
    out["init"] = ensure_province_ready(region, refresh_catalog=False)

    print("=== 2/7 university catalog ===")
    cat = write_catalog([region])
    out["catalog"] = cat["regions"].get(region)

    if not skip_discover:
        print("=== 3/7 discover by university (one-by-one) ===")
        from discover_by_university import discover_region

        out["discover"] = discover_region(region, merge_existing=True)
    else:
        print("=== 3/7 discover SKIPPED ===")

    print("=== 4/7 merge seeds ===")
    out["seeds"] = merge_seeds(region)

    if not skip_validate:
        print("=== 5/7 validate + PDF ===")
        from validate_job_links import validate_region

        out["validation"] = validate_region(region, workers=workers)
    else:
        print("=== 5/7 validate SKIPPED ===")

    if not skip_screen:
        print("=== 6/7 screen freshness (2026+ / 2025 rolling) ===")
        from screen_job_freshness import screen_region

        out["freshness"] = screen_region(region, refetch=True)
    else:
        print("=== 6/7 screen SKIPPED ===")

    if not skip_verify:
        print("=== 7/7 verify each posting one-by-one ===")
        from verify_jobs_one_by_one import verify_region

        out["verify"] = verify_region(region)
    else:
        print("=== 7/7 verify SKIPPED ===")

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
    ap.add_argument("--list", action="store_true", help="Print available regions.")
    ap.add_argument(
        "--init-only",
        action="store_true",
        help="Only scaffold seeds/universities for selected regions.",
    )
    ap.add_argument(
        "--summarize-only",
        action="store_true",
        help="Only merge existing province outputs into coastal PDF.",
    )
    ap.add_argument("--skip-discover", action="store_true")
    ap.add_argument("--skip-validate", action="store_true")
    ap.add_argument("--skip-screen", action="store_true")
    ap.add_argument("--skip-verify", action="store_true")
    ap.add_argument(
        "--no-summarize",
        action="store_true",
        help="Do not build coastal PDF after region runs.",
    )
    ap.add_argument(
        "--continue-on-error",
        action="store_true",
        help="When running multiple regions, continue after a region failure.",
    )
    ap.add_argument("--workers", type=int, default=6)
    args = ap.parse_args()

    if args.list:
        print("Available regions (%d):" % len(REGIONS))
        for r in REGIONS:
            uni = DATA_ECON / r / "universities.json"
            seeds = DATA_ECON / r / "seeds.json"
            print(
                "  %-12s universities=%s seeds=%s"
                % (
                    r,
                    "yes" if uni.exists() else "no",
                    "yes" if seeds.exists() else "no",
                )
            )
        return

    if args.summarize_only:
        print(json.dumps(summarize_coastal("verified"), ensure_ascii=False, indent=2))
        return

    targets = list(REGIONS) if args.all else (args.region or [])
    if not targets:
        if args.init_only:
            targets = list(REGIONS)
        else:
            raise SystemExit(
                "Pass --region <name>, --all, --list, --init-only, or --summarize-only.\n"
                "Regions: %s" % ", ".join(REGIONS)
            )

    if args.init_only:
        init_regions(targets)
        return

    summary: dict = {"date": TODAY, "regions": {}, "failed": []}
    for region in targets:
        try:
            summary["regions"][region] = run_region(
                region,
                skip_discover=args.skip_discover,
                skip_validate=args.skip_validate,
                skip_screen=args.skip_screen,
                skip_verify=args.skip_verify,
                workers=args.workers,
            )
        except Exception as exc:
            summary["regions"][region] = {
                "region": region,
                "ok": False,
                "error": str(exc),
            }
            summary["failed"].append(region)
            print("!! FAILED %s: %s" % (region, exc))
            traceback.print_exc()
            if not args.continue_on_error:
                raise

    if not args.no_summarize:
        print("\n########## COASTAL SUMMARIZE ##########")
        try:
            summary["coastal"] = summarize_coastal("verified")
        except Exception as exc:
            summary["coastal_error"] = str(exc)
            print("!! coastal summarize failed:", exc)

    out_path = DATA_ECON / ("province_pipeline_summary_%s.json" % TODAY)
    out_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    print("Summary -> %s" % out_path)
    if summary["failed"]:
        print("Failed regions: %s" % ", ".join(summary["failed"]))
    print("Done.")


if __name__ == "__main__":
    main()
