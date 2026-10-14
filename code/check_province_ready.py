# -*- coding: utf-8 -*-
"""Report whether each coastal province reached the coverage target.

Preliminary target: full list (validated/full) >= 2 * school_count
Advanced target: verified kept >= 2 * school_count (optional stricter)

Usage:
  python check_province_ready.py
  python check_province_ready.py --require-advanced
"""
from __future__ import annotations

import argparse
import json
from datetime import date
from pathlib import Path

from china_universities_data import REGIONS, universities_for_region
from paths import DATA_ECON, ensure_dirs

TODAY = date.today().isoformat()


def _count(path: Path) -> int:
    if not path.exists():
        return 0
    data = json.loads(path.read_text(encoding="utf-8"))
    return len(data) if isinstance(data, list) else int(data.get("count") or 0)


def region_status(region: str) -> dict:
    base = DATA_ECON / region
    schools = len(universities_for_region(region))
    need = schools * 2
    full = _count(base / "discovered_jobs_full.json")
    if full == 0:
        full = _count(base / "discovered_jobs_validated.json")
    verified = _count(base / "discovered_jobs_verified.json")
    if verified == 0:
        verified = _count(base / "discovered_jobs_fresh.json")
    return {
        "region": region,
        "schools": schools,
        "need": need,
        "full_list": full,
        "verified": verified,
        "preliminary_ready": full >= need,
        "advanced_ready": verified >= need,
    }


def check_all(*, require_advanced: bool = False) -> dict:
    ensure_dirs()
    regions = [region_status(r) for r in REGIONS]
    ready_key = "advanced_ready" if require_advanced else "preliminary_ready"
    not_ready = [r["region"] for r in regions if not r[ready_key]]
    summary = {
        "date": TODAY,
        "require_advanced": require_advanced,
        "all_ready": len(not_ready) == 0,
        "not_ready": not_ready,
        "regions": {r["region"]: r for r in regions},
    }
    out = DATA_ECON / ("province_ready_%s.json" % TODAY)
    out.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print("Province readiness (%s):" % ("advanced" if require_advanced else "preliminary"))
    for r in regions:
        flag = "OK" if r[ready_key] else "GAP"
        print(
            "  %-12s schools=%2d need>=%2d full=%3d verified=%3d  %s"
            % (
                r["region"],
                r["schools"],
                r["need"],
                r["full_list"],
                r["verified"],
                flag,
            )
        )
    print("all_ready=%s missing=%s" % (summary["all_ready"], ",".join(not_ready) or "-"))
    print("Summary ->", out)
    return summary


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument(
        "--require-advanced",
        action="store_true",
        help="Require verified>=2x schools (not just full list).",
    )
    args = ap.parse_args()
    check_all(require_advanced=args.require_advanced)


if __name__ == "__main__":
    main()
