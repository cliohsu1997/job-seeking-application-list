# -*- coding: utf-8 -*-
"""Download every provincial + Macau/Taiwan job page; set Job Type + Deadline.

Writes:
  data/economics/{region}/discovered_jobs_verified.json
  data/economics/macau/listings_*.json
  data/economics/taiwan/listings_*.json

Usage:
  python enrich_provincial_from_pages.py
  python enrich_provincial_from_pages.py --region shanghai
  python enrich_provincial_from_pages.py --force
"""
from __future__ import annotations

import argparse
import json
import time
from datetime import date
from pathlib import Path

from china_universities_data import REGIONS
from deadline_extract import extract_deadline, has_deadline
from download_china_provinces import fetch
from freshness_rules import strip_html
from paths import DATA_ECON, ensure_dirs
from validate_job_links import classify_type

TODAY = date.today().isoformat()
SLEEP = 0.25


def _norm_job_type(raw: str) -> str:
    """Provincial finals use Faculty | Postdoc (Faculty == former Academic)."""
    if raw == "Postdoc":
        return "Postdoc"
    if raw in ("Faculty", "Academic"):
        return "Faculty"
    return raw or "Faculty"


def enrich_job(job: dict, *, force: bool = False) -> tuple[dict, dict]:
    flags = {"fetched": 0, "deadline": 0, "type": 0, "ok": 0}
    url = job.get("JobPostUrl") or job.get("ApplicationLink") or ""
    if not str(url).startswith("http"):
        # Application-format MT rows
        url = job.get("Application Link") or url
    if not str(url).startswith("http"):
        return job, flags

    status, html = fetch(str(url))
    time.sleep(SLEEP)
    flags["fetched"] = 1
    if not html or (status and status >= 400):
        return job, flags
    flags["ok"] = 1

    body = strip_html(html)
    title = (
        job.get("JobPostTitle")
        or job.get("Job Post")
        or job.get("_title")
        or ""
    )
    if " | " in str(title):
        title = str(title).split(" | ", 1)[0].strip()
    blob = " ".join([str(title), str(url), body[:12000]])

    out = dict(job)
    deadline = extract_deadline(blob)
    if deadline and (force or not has_deadline(job)):
        if out.get("Deadline") != deadline:
            out["Deadline"] = deadline
            flags["deadline"] = 1

    new_type = classify_type(str(title), body)
    if new_type in ("Academic", "Postdoc"):
        mapped = _norm_job_type(new_type)
        old = _norm_job_type(
            out.get("JobType") or out.get("Job Type") or ""
        )
        out["JobType"] = mapped
        if "Job Type" in out:
            out["Job Type"] = mapped
        if mapped != old:
            flags["type"] = 1
    else:
        # Ensure Faculty label even when classifier unchanged
        mapped = _norm_job_type(
            out.get("JobType") or out.get("Job Type") or "Faculty"
        )
        out["JobType"] = mapped
        if "Job Type" in out:
            out["Job Type"] = mapped
    return out, flags


def enrich_region(region: str, *, force: bool = False) -> dict:
    ensure_dirs()
    base = DATA_ECON / region
    src = base / "discovered_jobs_verified.json"
    if not src.exists():
        src = base / "discovered_jobs_validated.json"
    if not src.exists():
        raise FileNotFoundError("No verified/validated jobs for %s" % region)

    jobs = json.loads(src.read_text(encoding="utf-8"))
    print("=== ENRICH PROVINCIAL %s -- %d jobs ===" % (region.upper(), len(jobs)))
    out = []
    n_dl = n_type = n_ok = 0
    for i, job in enumerate(jobs, 1):
        fixed, flags = enrich_job(job, force=force)
        out.append(fixed)
        n_dl += flags["deadline"]
        n_type += flags["type"]
        n_ok += flags["ok"]
        if flags["deadline"] or flags["type"]:
            print(
                "  [%d/%d] dl=%s type=%s | %s"
                % (
                    i,
                    len(jobs),
                    fixed.get("Deadline") or "-",
                    fixed.get("JobType") or "-",
                    (fixed.get("JobPostTitle") or "")[:48],
                )
            )
        elif i % 25 == 0 or i == len(jobs):
            print("  [%d/%d] scanned (pages_ok=%d)..." % (i, len(jobs), n_ok))

    for name in (
        "discovered_jobs_verified.json",
        "discovered_jobs_validated.json",
    ):
        path = base / name
        if path.exists() or name == "discovered_jobs_verified.json":
            path.write_text(
                json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8"
            )

    with_dl = sum(1 for j in out if has_deadline(j))
    summary = {
        "region": region,
        "jobs": len(out),
        "pages_ok": n_ok,
        "deadline_filled": n_dl,
        "type_changed": n_type,
        "with_deadline": with_dl,
        "faculty": sum(
            1 for j in out if _norm_job_type(j.get("JobType") or "") == "Faculty"
        ),
        "postdoc": sum(1 for j in out if j.get("JobType") == "Postdoc"),
    }
    print(
        "  pages_ok=%d deadline_filled=%d type_changed=%d with_deadline=%d/%d"
        % (n_ok, n_dl, n_type, with_dl, len(out))
    )
    return summary


def enrich_macau_taiwan(*, force: bool = False) -> list[dict]:
    """Enrich Macau/Taiwan listings_*.json (Application-format rows)."""
    summaries = []
    for region in ("macau", "taiwan"):
        folder = DATA_ECON / region
        paths = sorted(folder.glob("listings_*.json"), reverse=True)
        if not paths:
            print("!! no listings_*.json for %s" % region)
            continue
        path = paths[0]
        rows = json.loads(path.read_text(encoding="utf-8"))
        print(
            "=== ENRICH %s -- %d listings (%s) ==="
            % (region.upper(), len(rows), path.name)
        )
        out = []
        n_dl = n_type = n_ok = 0
        for i, row in enumerate(rows, 1):
            # Normalize keys for enrich_job
            job = dict(row)
            job.setdefault(
                "JobPostUrl",
                row.get("Application Link") or row.get("ApplicationLink"),
            )
            job.setdefault("JobPostTitle", row.get("Job Post") or "")
            job.setdefault("JobType", row.get("Job Type") or "Faculty")
            fixed, flags = enrich_job(job, force=force)
            n_dl += flags["deadline"]
            n_type += flags["type"]
            n_ok += flags["ok"]
            # Write back Application-format
            row_out = dict(row)
            row_out["Deadline"] = fixed.get("Deadline")
            row_out["Job Type"] = _norm_job_type(
                fixed.get("JobType") or fixed.get("Job Type") or "Faculty"
            )
            out.append(row_out)
            if flags["deadline"] or flags["type"] or i == len(rows):
                print(
                    "  [%d/%d] dl=%s type=%s | %s"
                    % (
                        i,
                        len(rows),
                        row_out.get("Deadline") or "-",
                        row_out.get("Job Type"),
                        str(row_out.get("University") or "")[:40],
                    )
                )
        path.write_text(
            json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        # Also stamp today's copy so loaders that prefer latest stay current
        today_path = folder / ("listings_%s.json" % TODAY)
        if today_path != path:
            today_path.write_text(
                json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8"
            )
        summaries.append(
            {
                "region": region,
                "jobs": len(out),
                "pages_ok": n_ok,
                "deadline_filled": n_dl,
                "type_changed": n_type,
                "with_deadline": sum(1 for r in out if has_deadline(r)),
                "path": str(path),
            }
        )
    return summaries


def run(*, regions: list[str] | None = None, force: bool = False) -> dict:
    targets = regions or list(REGIONS)
    summaries = []
    for region in targets:
        try:
            summaries.append(enrich_region(region, force=force))
        except FileNotFoundError as exc:
            print("!! skip %s: %s" % (region, exc))
    mt = enrich_macau_taiwan(force=force)
    summaries.extend(mt)
    out = {
        "date": TODAY,
        "regions": summaries,
        "jobs": sum(s.get("jobs", 0) for s in summaries),
        "deadline_filled": sum(s.get("deadline_filled", 0) for s in summaries),
        "type_changed": sum(s.get("type_changed", 0) for s in summaries),
        "pages_ok": sum(s.get("pages_ok", 0) for s in summaries),
    }
    (DATA_ECON / ("provincial_page_enrich_summary_%s.json" % TODAY)).write_text(
        json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(
        "Provincial+MT enrich done. jobs=%d pages_ok=%d deadline_filled=%d type_changed=%d"
        % (out["jobs"], out["pages_ok"], out["deadline_filled"], out["type_changed"])
    )
    return out


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--region", action="append", default=None)
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()
    print(
        json.dumps(
            run(regions=args.region, force=args.force),
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
