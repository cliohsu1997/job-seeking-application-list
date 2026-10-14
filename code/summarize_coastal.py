# -*- coding: utf-8 -*-
"""Merge per-province verified/validated jobs into one coastal listing + PDF.

Usage:
  python summarize_coastal.py
  python summarize_coastal.py --prefer verified
"""
from __future__ import annotations

import argparse
import json
from datetime import date
from pathlib import Path

from china_universities_data import REGIONS
from download_china_provinces import job_to_row, write_xlsx
from export_clean_pdf import export_one
from paths import CLEAN, CLEAN_ECON, DATA_ECON, ensure_dirs

TODAY = date.today().isoformat()


def load_region_jobs(region: str, prefer: str) -> list[dict]:
    base = DATA_ECON / region
    names = []
    if prefer == "verified":
        names = [
            "discovered_jobs_verified.json",
            "discovered_jobs_fresh.json",
            "discovered_jobs_validated.json",
        ]
    else:
        names = [
            "discovered_jobs_validated.json",
            "discovered_jobs_verified.json",
            "discovered_jobs_fresh.json",
        ]
    for name in names:
        path = base / name
        if path.exists():
            return json.loads(path.read_text(encoding="utf-8"))
    return []


def summarize(prefer: str = "verified") -> dict:
    ensure_dirs()
    all_jobs: list[dict] = []
    per_region = {}
    for region in REGIONS:
        jobs = load_region_jobs(region, prefer)
        # annotate region if missing
        for j in jobs:
            j.setdefault("region", region)
        per_region[region] = {
            "count": len(jobs),
            "academic": sum(1 for j in jobs if j.get("JobType") == "Academic"),
            "postdoc": sum(1 for j in jobs if j.get("JobType") == "Postdoc"),
        }
        all_jobs.extend(jobs)
        print("  %s: %d" % (region, len(jobs)))

    # URL dedupe
    by_url = {}
    for j in all_jobs:
        u = j.get("JobPostUrl") or j.get("ApplicationLink")
        if u:
            by_url[u] = j
    merged = list(by_url.values())

    rows = [job_to_row(j) for j in merged]
    CLEAN.mkdir(parents=True, exist_ok=True)
    CLEAN_ECON.mkdir(parents=True, exist_ok=True)
    xlsx = CLEAN_ECON / ("china_coastal_listings_validated_%s.xlsx" % TODAY)
    pdf = CLEAN / ("china_coastal_listings_validated_%s.pdf" % TODAY)
    write_xlsx(rows, xlsx, "China Coastal Validated")
    export_one(
        xlsx,
        pdf,
        "China Coastal - Faculty + Postdoc (validated)",
        use_cjk=True,
    )
    (DATA_ECON / ("china_coastal_jobs_%s.json" % TODAY)).write_text(
        json.dumps(merged, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    summary = {
        "date": TODAY,
        "total": len(merged),
        "regions": per_region,
        "pdf": str(pdf),
        "xlsx": str(xlsx),
    }
    (DATA_ECON / ("china_coastal_summary_%s.json" % TODAY)).write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print("Coastal total=%d -> %s" % (len(merged), pdf))
    return summary


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument(
        "--prefer",
        choices=("verified", "validated"),
        default="verified",
        help="Which per-region file family to prefer",
    )
    args = ap.parse_args()
    print(json.dumps(summarize(args.prefer), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
