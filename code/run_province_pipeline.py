# -*- coding: utf-8 -*-
"""Province-agnostic coastal economics faculty/postdoc pipeline.

Two-stage filtering:
  PRELIMINARY (build full list ~2x schools)
    1. Ensure province workspace
    2. University catalog
    3. Discover by university
    4. Merge seeds
    5. Validate -> discovered_jobs_full.json (+ validated PDF)
  ADVANCED (one-by-one on the full list)
    6. Re-fetch each posting: right content, drop deadline-passed,
       drop 2025-without-rolling, drop other stale -> verified PDF
  COASTAL
    7. When provinces are ready, summarize into coastal PDF

Usage:
  python run_province_pipeline.py --list
  python run_province_pipeline.py --region shanghai
  python run_province_pipeline.py --all --continue-on-error
  python run_province_pipeline.py --check-ready --require-advanced
  python run_province_pipeline.py --summarize-only
"""
from __future__ import annotations

import argparse
import json
import traceback
from datetime import date
from pathlib import Path

from china_universities_data import REGIONS, universities_for_region
from check_province_ready import check_all
from download_china_universities import write_catalog
from init_province_pipeline import ensure_province_ready, init_regions
from paths import DATA_ECON, ensure_dirs
from seed_jobs import merge_seeds
from summarize_coastal import summarize as summarize_coastal

TODAY = date.today().isoformat()


def _save_full_list(region: str) -> dict:
    """Copy preliminary validated output to discovered_jobs_full.json and report 2x coverage."""
    base = DATA_ECON / region
    src = base / "discovered_jobs_validated.json"
    if not src.exists():
        src = base / "discovered_jobs.json"
    jobs = json.loads(src.read_text(encoding="utf-8")) if src.exists() else []
    (base / "discovered_jobs_full.json").write_text(
        json.dumps(jobs, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    schools = len(universities_for_region(region))
    need = schools * 2
    ready = len(jobs) >= need
    print(
        "  full list: %d jobs | schools=%d need>=%d | preliminary %s"
        % (len(jobs), schools, need, "READY" if ready else "BELOW TARGET")
    )
    return {
        "full_list": len(jobs),
        "schools": schools,
        "need": need,
        "preliminary_ready": ready,
    }


def run_region(
    region: str,
    *,
    skip_discover: bool = False,
    skip_preliminary: bool = False,
    skip_advanced: bool = False,
    workers: int = 6,
) -> dict:
    ensure_dirs()
    print("\n########## PIPELINE %s ##########" % region.upper())
    out: dict = {"region": region, "date": TODAY, "ok": True}

    print("=== 1/6 ensure province workspace ===")
    out["init"] = ensure_province_ready(region, refresh_catalog=False)

    print("=== 2/6 university catalog ===")
    cat = write_catalog([region])
    out["catalog"] = cat["regions"].get(region)

    if not skip_discover:
        print("=== 3/6 discover by university ===")
        from discover_by_university import discover_region

        out["discover"] = discover_region(region, merge_existing=True)
    else:
        print("=== 3/6 discover SKIPPED ===")

    print("=== 4/6 merge seeds ===")
    out["seeds"] = merge_seeds(region)

    if not skip_preliminary:
        print("=== 5/6 PRELIMINARY filter -> full list (~2x schools) ===")
        from validate_job_links import validate_region

        out["preliminary"] = validate_region(region, workers=workers)
        out["full_list"] = _save_full_list(region)
    else:
        print("=== 5/6 PRELIMINARY SKIPPED ===")
        if (DATA_ECON / region / "discovered_jobs_full.json").exists() or (
            DATA_ECON / region / "discovered_jobs_validated.json"
        ).exists():
            out["full_list"] = _save_full_list(region)

    if not skip_advanced:
        print("=== 6/6 ADVANCED filter (content + deadline + 2025/rolling) ===")
        from advanced_filter_jobs import advanced_filter_region

        out["advanced"] = advanced_filter_region(region)
    else:
        print("=== 6/6 ADVANCED SKIPPED ===")

    return out


def main() -> None:
    ap = argparse.ArgumentParser(
        description="Coastal province pipeline: preliminary full list then advanced filter."
    )
    ap.add_argument("--region", action="append", choices=list(REGIONS))
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--list", action="store_true")
    ap.add_argument("--init-only", action="store_true")
    ap.add_argument(
        "--check-ready",
        action="store_true",
        help="Print 2x-school readiness and exit.",
    )
    ap.add_argument(
        "--require-advanced",
        action="store_true",
        help="With --check-ready / summarize: require verified>=2x.",
    )
    ap.add_argument(
        "--summarize-only",
        action="store_true",
        help="Merge coastal PDF (warns if provinces not ready).",
    )
    ap.add_argument("--skip-discover", action="store_true")
    ap.add_argument(
        "--skip-preliminary",
        action="store_true",
        help="Skip validate/full-list stage.",
    )
    ap.add_argument(
        "--skip-advanced",
        action="store_true",
        help="Skip one-by-one advanced filter.",
    )
    ap.add_argument(
        "--no-summarize",
        action="store_true",
        help="Do not auto-summarize coastal after runs.",
    )
    ap.add_argument(
        "--summarize-when-ready",
        action="store_true",
        help="Only build coastal PDF if all provinces meet the 2x target.",
    )
    ap.add_argument("--continue-on-error", action="store_true")
    ap.add_argument("--workers", type=int, default=6)
    args = ap.parse_args()

    if args.list:
        print("Available regions (%d):" % len(REGIONS))
        for r in REGIONS:
            print(" ", r)
        return

    if args.check_ready:
        check_all(require_advanced=args.require_advanced)
        return

    if args.summarize_only:
        ready = check_all(require_advanced=args.require_advanced)
        if args.summarize_when_ready and not ready["all_ready"]:
            raise SystemExit(
                "Not all provinces ready: %s" % ", ".join(ready["not_ready"])
            )
        print(json.dumps(summarize_coastal("verified"), ensure_ascii=False, indent=2))
        return

    targets = list(REGIONS) if args.all else (args.region or [])
    if not targets:
        if args.init_only:
            targets = list(REGIONS)
        else:
            raise SystemExit(
                "Pass --region, --all, --list, --init-only, --check-ready, or --summarize-only.\n"
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
                skip_preliminary=args.skip_preliminary,
                skip_advanced=args.skip_advanced,
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

    ready = check_all(require_advanced=True)
    summary["readiness"] = {
        "all_ready": ready["all_ready"],
        "not_ready": ready["not_ready"],
    }

    if not args.no_summarize:
        if args.summarize_when_ready and not ready["all_ready"]:
            print(
                "Coastal summarize deferred; not ready: %s"
                % ", ".join(ready["not_ready"])
            )
        else:
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
