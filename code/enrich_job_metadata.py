# -*- coding: utf-8 -*-
"""Scan jobs one-by-one for Published Year (2025/2026) and College Ranking (985/211).

Usage:
  python enrich_job_metadata.py --region shanghai
  python enrich_job_metadata.py --all
  python enrich_job_metadata.py --file data/economics/beijing/discovered_jobs_verified.json
"""
from __future__ import annotations

import argparse
import json
import re
import time
from datetime import date
from pathlib import Path

from china_universities_data import REGIONS, universities_for_region
from download_china_provinces import fetch
from freshness_rules import strip_html
from paths import DATA_ECON, ensure_dirs

TODAY = date.today().isoformat()
SLEEP = 0.2

YEAR_2026 = re.compile(r"2026|2026[-/]\d{1,2}|26[-/]\d{2}")
YEAR_2025 = re.compile(r"2025|2025[-/]\d{1,2}|25[-/]\d{2}")
YEAR_ANY = re.compile(r"(20(?:2[4-9]|3[0-9]))")


def _norm(s: str) -> str:
    return re.sub(r"[\s\-_,.\u00b7]+", "", (s or "").lower())


def build_ranking_index() -> dict[str, str]:
    """Map normalized university name -> ranking label (985 / 211 / 985/211)."""
    idx: dict[str, str] = {}
    for region in REGIONS:
        for u in universities_for_region(region):
            tags = [t for t in (u.get("tags") or []) if t in ("985", "211")]
            if not tags:
                label = ""
            elif "985" in tags and "211" in tags:
                label = "985/211"
            elif "985" in tags:
                label = "985"
            else:
                label = "211"
            for key in (u.get("name_en"), u.get("name_zh")):
                if key:
                    idx[_norm(key)] = label
            en = u.get("name_en") or ""
            if en:
                idx[_norm(en.replace(" University", ""))] = label
    return idx


def ranking_for_university(name: str, index: dict[str, str]) -> str:
    if not name:
        return ""
    n = _norm(name)
    if n in index:
        return index[n]
    for key, label in index.items():
        if not label:
            continue
        if key and (key in n or n in key) and len(key) >= 4:
            return label
    return ""


def published_year_from_text(*parts: str) -> str:
    blob = " ".join(p for p in parts if p)
    if not blob:
        return ""
    if YEAR_2026.search(blob):
        return "2026"
    if YEAR_2025.search(blob):
        return "2025"
    years = YEAR_ANY.findall(blob)
    if years:
        return max(years)
    return ""


def enrich_job(
    job: dict,
    *,
    ranking_index: dict[str, str],
    fetch_page: bool = True,
) -> dict:
    out = dict(job)
    title = out.get("JobPostTitle") or ""
    url = out.get("JobPostUrl") or out.get("ApplicationLink") or ""
    year = published_year_from_text(title, url, out.get("RequiredMaterial") or "")
    if fetch_page and url.startswith("http") and not year:
        status, html = fetch(url)
        time.sleep(SLEEP)
        body = strip_html(html or "")[:6000]
        year = published_year_from_text(title, url, body)
        if status and status >= 400 and not year:
            year = published_year_from_text(title, url)

    ranking = ranking_for_university(out.get("University") or "", ranking_index)
    out["PublishedYear"] = year
    out["CollegeRanking"] = ranking
    out["Published Year"] = year
    out["College Ranking"] = ranking
    return out


def enrich_jobs(
    jobs: list[dict],
    *,
    fetch_page: bool = True,
    ranking_index: dict[str, str] | None = None,
) -> list[dict]:
    idx = ranking_index if ranking_index is not None else build_ranking_index()
    return [
        enrich_job(j, ranking_index=idx, fetch_page=fetch_page) for j in jobs
    ]


def enrich_region(region: str, *, fetch_page: bool = True) -> dict:
    ensure_dirs()
    base = DATA_ECON / region
    src = base / "discovered_jobs_verified.json"
    if not src.exists():
        src = base / "discovered_jobs_validated.json"
    if not src.exists():
        raise FileNotFoundError("No verified/validated jobs for %s" % region)
    jobs = json.loads(src.read_text(encoding="utf-8"))
    print("=== ENRICH %s -- %d jobs ===" % (region.upper(), len(jobs)))
    idx = build_ranking_index()
    enriched = enrich_jobs(jobs, fetch_page=fetch_page, ranking_index=idx)
    for name in (
        "discovered_jobs_verified.json",
        "discovered_jobs_validated.json",
        "discovered_jobs_fresh.json",
        "discovered_jobs.json",
    ):
        path = base / name
        if path.exists() or name == "discovered_jobs_verified.json":
            path.write_text(
                json.dumps(enriched, ensure_ascii=False, indent=2), encoding="utf-8"
            )
    years: dict[str, int] = {}
    ranks: dict[str, int] = {}
    for j in enriched:
        y = j.get("PublishedYear") or ""
        years[y] = years.get(y, 0) + 1
        r = j.get("CollegeRanking") or ""
        ranks[r] = ranks.get(r, 0) + 1
    summary = {
        "region": region,
        "count": len(enriched),
        "published_year": years,
        "college_ranking": ranks,
    }
    (base / "enrich_metadata_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print("  years=%s ranks=%s" % (years, {k or "(none)": v for k, v in ranks.items()}))
    return summary


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--region", action="append", choices=list(REGIONS))
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--no-fetch", action="store_true")
    ap.add_argument("--file", type=Path)
    args = ap.parse_args()
    if args.file:
        jobs = json.loads(args.file.read_text(encoding="utf-8"))
        enriched = enrich_jobs(jobs, fetch_page=not args.no_fetch)
        args.file.write_text(
            json.dumps(enriched, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print("Wrote", args.file, len(enriched))
        return
    targets = list(REGIONS) if args.all else (args.region or [])
    if not targets:
        raise SystemExit("Pass --region / --all / --file")
    summary = {"date": TODAY, "regions": {}}
    for region in targets:
        summary["regions"][region] = enrich_region(
            region, fetch_page=not args.no_fetch
        )
    out = DATA_ECON / ("enrich_metadata_summary_%s.json" % TODAY)
    out.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
