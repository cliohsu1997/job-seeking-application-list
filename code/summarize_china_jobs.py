# -*- coding: utf-8 -*-
"""Merge coastal province verified jobs + Macau/Taiwan into one final clean/ PDF.

Top-level clean/ gets only:
  china_macau_taiwan_listings_YYYY-MM-DD.pdf

Intermediate xlsx stay under clean/economics/.

Usage:
  python summarize_china_jobs.py
  python summarize_china_jobs.py --coastal-only
  python summarize_china_jobs.py --prefer verified
"""
from __future__ import annotations

import argparse
import json
from datetime import date, datetime
from pathlib import Path

from china_universities_data import REGIONS
from download_china_provinces import job_to_row, write_xlsx
from enrich_job_metadata import enrich_jobs
from export_clean_pdf import export_one
from paths import CLEAN, CLEAN_ECON, CLEAN_ECON_PROVINCE, DATA_ECON, ensure_dirs

TODAY = date.today().isoformat()


def load_region_jobs(region: str, prefer: str) -> list[dict]:
    base = DATA_ECON / region
    names = (
        [
            "discovered_jobs_verified.json",
            "discovered_jobs_fresh.json",
            "discovered_jobs_validated.json",
        ]
        if prefer == "verified"
        else [
            "discovered_jobs_validated.json",
            "discovered_jobs_verified.json",
            "discovered_jobs_fresh.json",
        ]
    )
    for name in names:
        path = base / name
        if path.exists():
            jobs = json.loads(path.read_text(encoding="utf-8"))
            for j in jobs:
                j.setdefault("region", region)
            return jobs
    return []


def _latest_listings(folder: Path) -> list[dict]:
    paths = sorted(folder.glob("listings_*.json"), reverse=True)
    if not paths:
        return []
    return json.loads(paths[0].read_text(encoding="utf-8"))


def macau_taiwan_to_internal(rows: list[dict], region_key: str) -> list[dict]:
    """Normalize Application-format MT rows into coastal-style dicts."""
    out = []
    for r in rows:
        link = r.get("Application Link") or r.get("ApplicationLink") or ""
        job_post = r.get("Job Post") or r.get("JobPostTitle") or ""
        title = job_post
        if " | " in job_post:
            title = job_post.split(" | ", 1)[0].strip()
        deadline = r.get("Deadline")
        if isinstance(deadline, datetime):
            deadline = deadline.strftime("%Y-%m-%d")
        out.append(
            {
                "region": region_key,
                "Country": r.get("Country")
                or ("Macau" if region_key == "macau" else "Taiwan"),
                "Location": r.get("Location"),
                "University": r.get("University"),
                "School": r.get("School"),
                "Deadline": deadline,
                "JobPostTitle": title,
                "JobPostUrl": link,
                "RequiredMaterial": r.get("Required Material")
                or r.get("RequiredMaterial"),
                "SubmissionChannel": r.get("Submission Channel")
                or r.get("SubmissionChannel"),
                "ApplicationLink": link,
                "Email": r.get("Email"),
                "JobType": r.get("Job Type") or r.get("JobType") or "Faculty",
            }
        )
    return out


def load_macau_taiwan() -> tuple[list[dict], list[dict]]:
    macau = macau_taiwan_to_internal(
        _latest_listings(DATA_ECON / "macau"), "macau"
    )
    taiwan = macau_taiwan_to_internal(
        _latest_listings(DATA_ECON / "taiwan"), "taiwan"
    )
    return macau, taiwan


def _dedupe(jobs: list[dict]) -> list[dict]:
    by_url: dict[str, dict] = {}
    for j in jobs:
        u = j.get("JobPostUrl") or j.get("ApplicationLink") or ""
        if not u:
            u = "%s|%s" % (j.get("University"), j.get("JobPostTitle"))
        by_url[u] = j
    return list(by_url.values())


def _write_xlsx(jobs: list[dict], stem: str, title: str) -> Path:
    rows = [job_to_row(j) for j in jobs]
    xlsx = CLEAN_ECON / ("%s_%s.xlsx" % (stem, TODAY))
    write_xlsx(rows, xlsx, title[:31])
    return xlsx


def summarize(*, prefer: str = "verified", coastal_only: bool = False) -> dict:
    ensure_dirs()
    coastal: list[dict] = []
    per_region = {}
    for region in REGIONS:
        jobs = load_region_jobs(region, prefer)
        per_region[region] = {
            "count": len(jobs),
            "academic": sum(1 for j in jobs if j.get("JobType") == "Academic"),
            "postdoc": sum(1 for j in jobs if j.get("JobType") == "Postdoc"),
        }
        coastal.extend(jobs)
        print("  coastal %-12s %d" % (region, len(jobs)))

    coastal = _dedupe(coastal)
    # Fill Published Year + College Ranking from title/URL + catalog (no page fetch)
    coastal = enrich_jobs(coastal, fetch_page=False)
    macau, taiwan = ([], []) if coastal_only else load_macau_taiwan()
    print("  macau          %d" % len(macau))
    print("  taiwan         %d" % len(taiwan))

    all_jobs = _dedupe(coastal + macau + taiwan)
    all_jobs = enrich_jobs(all_jobs, fetch_page=False)
    CLEAN.mkdir(parents=True, exist_ok=True)
    CLEAN_ECON.mkdir(parents=True, exist_ok=True)

    coastal_xlsx = _write_xlsx(
        coastal,
        "china_coastal_listings_validated",
        "China Coastal - Faculty + Postdoc (validated)",
    )
    CLEAN_ECON_PROVINCE.mkdir(parents=True, exist_ok=True)
    coastal_pdf = CLEAN_ECON_PROVINCE / (
        "china_coastal_listings_validated_%s.pdf" % TODAY
    )
    export_one(
        coastal_xlsx,
        coastal_pdf,
        "China Coastal - Faculty + Postdoc (validated)",
        use_cjk=True,
    )
    (DATA_ECON / ("china_coastal_jobs_%s.json" % TODAY)).write_text(
        json.dumps(coastal, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    greater_xlsx = greater_pdf = None
    if not coastal_only:
        # Intermediate xlsx under clean/economics/; final provincial_* lives in clean/
        # via combine_asia_pacific.py (preferred). Keep a root provincial export here too.
        from download_china_provinces import write_xlsx as _wx
        from download_china_provinces import job_to_row as _jtr

        rows = [_jtr(j) for j in all_jobs]
        greater_xlsx = CLEAN / ("provincial_listings_%s.xlsx" % TODAY)
        greater_pdf = CLEAN / ("provincial_listings_%s.pdf" % TODAY)
        _wx(rows, greater_xlsx, "Provincial China+MT")
        export_one(
            greater_xlsx,
            greater_pdf,
            "Provincial China + Macau/Taiwan - Faculty + Postdoc",
            use_cjk=True,
        )
        # legacy name under economics/ for back-compat
        _write_xlsx(
            all_jobs,
            "china_macau_taiwan_listings",
            "China + Macau + Taiwan - Faculty + Postdoc",
        )
        (DATA_ECON / ("china_macau_taiwan_jobs_%s.json" % TODAY)).write_text(
            json.dumps(all_jobs, ensure_ascii=False, indent=2), encoding="utf-8"
        )

    summary = {
        "date": TODAY,
        "coastal_total": len(coastal),
        "macau": len(macau),
        "taiwan": len(taiwan),
        "greater_china_total": len(all_jobs),
        "regions": per_region,
        "coastal_xlsx": str(coastal_xlsx),
        "coastal_pdf": str(coastal_pdf),
        "greater_pdf": str(greater_pdf) if greater_pdf else None,
        "greater_xlsx": str(greater_xlsx) if greater_xlsx else None,
    }
    (DATA_ECON / ("china_jobs_summary_%s.json" % TODAY)).write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(
        "Coastal=%d Macau=%d Taiwan=%d Greater=%d"
        % (len(coastal), len(macau), len(taiwan), len(all_jobs))
    )
    if greater_pdf:
        print("Final clean PDF ->", greater_pdf)
    else:
        print("Coastal xlsx only (no root PDF) ->", coastal_xlsx)
    return summary


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--prefer", choices=("verified", "validated"), default="verified")
    ap.add_argument("--coastal-only", action="store_true")
    args = ap.parse_args()
    print(
        json.dumps(
            summarize(prefer=args.prefer, coastal_only=args.coastal_only),
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
