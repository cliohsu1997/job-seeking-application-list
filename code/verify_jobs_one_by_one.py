# -*- coding: utf-8 -*-
"""One-by-one verification that each kept URL is a real open faculty/postdoc post.

Reads discovered_jobs_fresh.json (or validated), re-fetches each page, and applies
the same score_announcement rules as validate_job_links.

Usage:
  python verify_jobs_one_by_one.py --region shanghai
  python verify_jobs_one_by_one.py --all
"""
from __future__ import annotations

import argparse
import json
from datetime import date
from pathlib import Path

from china_universities_data import REGIONS
from download_china_provinces import fetch, job_to_row, write_xlsx
from export_clean_pdf import export_one
from paths import CLEAN, CLEAN_ECON, DATA_ECON, ensure_dirs
from validate_job_links import score_announcement

TODAY = date.today().isoformat()


def verify_region(region: str) -> dict:
    ensure_dirs()
    base = DATA_ECON / region
    src = base / "discovered_jobs_fresh.json"
    if not src.exists():
        src = base / "discovered_jobs_validated.json"
    if not src.exists():
        raise FileNotFoundError("No jobs to verify for %s" % region)

    jobs = json.loads(src.read_text(encoding="utf-8"))
    print("=== VERIFY ONE-BY-ONE %s -- %d jobs ===" % (region.upper(), len(jobs)))
    verified, rejected = [], []
    for i, j in enumerate(jobs, 1):
        url = j.get("JobPostUrl") or j.get("ApplicationLink") or ""
        title = j.get("JobPostTitle") or ""
        unit = j.get("School") or ""
        status, html = fetch(url) if url else (None, "")
        result = score_announcement(
            html or "", url, title, unit_zh=unit, status=status
        )
        row = dict(j)
        row["_verify"] = result
        if result.get("keep"):
            verified.append(row)
            mark = "KEEP"
        else:
            rejected.append(row)
            mark = "DROP"
        print(
            "  [%d/%d] %s %s | %s"
            % (i, len(jobs), mark, result.get("label"), (title or "")[:50])
        )

    kept_clean = [{k: v for k, v in j.items() if not k.startswith("_")} for j in verified]
    audit = {
        "date": TODAY,
        "region": region,
        "input": len(jobs),
        "verified": len(kept_clean),
        "rejected": len(rejected),
        "rows": [
            {
                "title": j.get("JobPostTitle"),
                "url": j.get("JobPostUrl"),
                "keep": j["_verify"].get("keep"),
                "label": j["_verify"].get("label"),
                "reasons": j["_verify"].get("reasons"),
                "job_type": j.get("JobType"),
            }
            for j in verified + rejected
        ],
    }
    (base / "verify_audit.json").write_text(
        json.dumps(audit, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (base / "discovered_jobs_verified.json").write_text(
        json.dumps(kept_clean, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    # Final listing for the province
    (base / "discovered_jobs_validated.json").write_text(
        json.dumps(kept_clean, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (base / "discovered_jobs.json").write_text(
        json.dumps(kept_clean, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    rows = [job_to_row(j) for j in kept_clean]
    xlsx = CLEAN_ECON / ("%s_listings_verified_%s.xlsx" % (region, TODAY))
    pdf = CLEAN / ("%s_listings_verified_%s.pdf" % (region, TODAY))
    CLEAN_ECON.mkdir(parents=True, exist_ok=True)
    CLEAN.mkdir(parents=True, exist_ok=True)
    write_xlsx(rows, base / ("listings_verified_%s.xlsx" % TODAY), region.title())
    write_xlsx(rows, xlsx, region.title())
    export_one(
        xlsx,
        pdf,
        "%s Economics - Faculty + Postdoc (verified)" % region.title(),
        use_cjk=True,
    )
    # Also refresh the validated-named PDF used by the rest of the pipeline
    export_one(
        xlsx,
        CLEAN / ("%s_listings_validated_%s.pdf" % (region, TODAY)),
        "%s Economics - Faculty + Postdoc (validated)" % region.title(),
        use_cjk=True,
    )
    print("-> %s verified=%d/%d PDF -> %s" % (region, len(kept_clean), len(jobs), pdf))
    return {
        "input": len(jobs),
        "verified": len(kept_clean),
        "rejected": len(rejected),
        "pdf": str(pdf),
    }


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--region", action="append", choices=list(REGIONS))
    ap.add_argument("--all", action="store_true")
    args = ap.parse_args()
    targets = list(REGIONS) if args.all else (args.region or [])
    if not targets:
        raise SystemExit("Pass --region or --all")
    summary = {"date": TODAY, "regions": {}}
    for region in targets:
        summary["regions"][region] = verify_region(region)
    out = DATA_ECON / ("verify_summary_%s.json" % TODAY)
    out.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
