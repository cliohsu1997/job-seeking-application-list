# -*- coding: utf-8 -*-
"""Re-fetch job pages to repair titles/schools corrupted to ASCII '?' .

Caused by Windows console / agent writes that replaced CJK with '?'.
Run:
  python repair_corrupted_text.py --region fujian
  python repair_corrupted_text.py --all
"""
from __future__ import annotations

import argparse
import json
import re
import time
from datetime import date

from china_universities_data import REGIONS
from download_china_provinces import fetch, job_to_row, write_xlsx
from paths import CLEAN_ECON, DATA_ECON, ensure_dirs

TODAY = date.today().isoformat()
SLEEP = 0.25

_TITLE_RE = re.compile(r"<title[^>]*>(.*?)</title>", re.I | re.S)
_H1_RE = re.compile(r"<h1[^>]*>(.*?)</h1>", re.I | re.S)
_TAG_RE = re.compile(r"<[^>]+>")
_WS_RE = re.compile(r"\s+")
# Site suffixes often appended to <title>
_TITLE_SUFFIX = re.compile(
    r"[-_|]\s*[\u4e00-\u9fffA-Za-z0-9\u00b7\s]{0,40}"
    r"(\u5927\u5b66|\u5b66\u9662|University|College)\s*$"
)

_SCHOOL_KEYS = (
    "\u7ecf\u6d4e\u5b66\u9662",
    "\u7ecf\u6d4e\u4e0e\u7ba1\u7406\u5b66\u9662",
    "\u5546\u5b66\u9662",
    "\u91d1\u878d\u5b66\u9662",
    "\u738b\u4e9a\u5357\u7ecf\u6d4e\u7814\u7a76\u9662",
    "WISE",
)

# Fallback when page HTML has no recoverable school string.
_UNI_SCHOOL = {
    "Xiamen University": "\u7ecf\u6d4e\u5b66\u9662/WISE",
    "Fuzhou University": "\u7ecf\u6d4e\u4e0e\u7ba1\u7406\u5b66\u9662",
    "FJNU": "\u7ecf\u6d4e\u5b66\u9662",
    "FJXU": "\u7ecf\u6d4e\u4e0e\u8d38\u6613\u5b66\u9662",
    "FAFU": "\u7ecf\u6d4e\u4e0e\u7ba1\u7406\u5b66\u9662",
    "Huaqiao University": "\u7ecf\u6d4e\u4e0e\u91d1\u878d\u5b66\u9662",
    "Jimei University": "\u8d22\u7ecf\u5b66\u9662",
    "Quanzhou Normal University": "\u7ecf\u6d4e\u4e0e\u7ba1\u7406\u5b66\u9662",
    "Shandong University": "\u7ecf\u6d4e\u5b66\u9662",
    "Ocean University of China": "\u7ecf\u6d4e\u5b66\u9662",
    "CUP East China": "\u7ecf\u6d4e\u7ba1\u7406\u5b66\u9662",
    "Qingdao University": "\u7ecf\u6d4e\u5b66\u9662",
    "SDUFE": "\u7ecf\u6d4e\u5b66\u9662",
    "Shenzhen University": "\u7ecf\u6d4e\u5b66\u9662",
    "HIT Shenzhen": "\u7ecf\u6d4e\u7ba1\u7406\u5b66\u9662",
    "PKU Shenzhen": "HSBC Business School",
    "SMBU": "\u5546\u5b66\u9662",
}

_REGION_LOC = {
    "fujian": "\u798f\u5efa",
    "shandong": "\u5c71\u4e1c",
    "shenzhen": "\u6df1\u5733",
}


def looks_corrupted(text: str | None) -> bool:
    if not text:
        return False
    s = str(text)
    q = s.count("?")
    if q < 3:
        return False
    cjk = sum(1 for c in s if "\u4e00" <= c <= "\u9fff")
    latin = sum(1 for c in s if c.isalpha() and c.isascii())
    return q >= 3 and cjk < 2 and latin < 8


def strip_html(html: str) -> str:
    return _WS_RE.sub(" ", _TAG_RE.sub(" ", html or "")).strip()


def page_title(html: str) -> str | None:
    if not html:
        return None
    for cre in (_H1_RE, _TITLE_RE):
        m = cre.search(html)
        if not m:
            continue
        t = strip_html(m.group(1))
        if not t or looks_corrupted(t):
            continue
        t = _TITLE_SUFFIX.sub("", t).strip(" -_|")
        if len(t) >= 4:
            return t[:180]
    return None


def repair_job(job: dict) -> tuple[dict, bool]:
    url = job.get("JobPostUrl") or job.get("ApplicationLink") or ""
    need = (
        looks_corrupted(job.get("JobPostTitle"))
        or looks_corrupted(job.get("School"))
        or looks_corrupted(job.get("RequiredMaterial"))
    )
    if not need or not url.startswith("http"):
        return job, False
    status, html = fetch(url)
    time.sleep(SLEEP)
    if not html or (status and status >= 400):
        return job, False
    changed = False
    title = page_title(html)
    if title and looks_corrupted(job.get("JobPostTitle")):
        job = dict(job)
        job["JobPostTitle"] = title
        changed = True
    school = job.get("School")
    if looks_corrupted(school):
        blob = (title or "") + " " + (html or "")[:3000]
        new_school = None
        for key in _SCHOOL_KEYS:
            if key in blob:
                new_school = key
                break
        if not new_school:
            uni = job.get("University") or ""
            for key, val in _UNI_SCHOOL.items():
                if key.lower() in uni.lower() or uni.lower() in key.lower():
                    new_school = val
                    break
        if new_school is not None:
            if not changed:
                job = dict(job)
            job["School"] = new_school
            changed = True
        elif looks_corrupted(school):
            if not changed:
                job = dict(job)
            job["School"] = ""
            changed = True
    if looks_corrupted(job.get("Location")):
        if not changed:
            job = dict(job)
        job["Location"] = _REGION_LOC.get(job.get("region") or "", "")
        changed = True
    if looks_corrupted(job.get("RequiredMaterial")):
        if not changed:
            job = dict(job)
        job["RequiredMaterial"] = "See announcement page"
        changed = True
    return job, changed


def repair_region(region: str) -> dict:
    base = DATA_ECON / region
    targets = [
        "discovered_jobs_verified.json",
        "discovered_jobs_validated.json",
        "discovered_jobs_fresh.json",
        "discovered_jobs_full.json",
        "discovered_jobs.json",
    ]
    summary = {"region": region, "files": {}, "repaired": 0}
    cache: dict[str, dict] = {}

    for name in targets:
        path = base / name
        if not path.exists():
            continue
        jobs = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(jobs, list):
            continue
        out = []
        n = 0
        for job in jobs:
            url = job.get("JobPostUrl") or job.get("ApplicationLink") or ""
            if url in cache:
                fixed = dict(job)
                fixed.update(cache[url])
                if fixed != job:
                    n += 1
                out.append(fixed)
                continue
            fixed, changed = repair_job(job)
            if changed:
                n += 1
                cache[url] = {
                    k: fixed[k]
                    for k in ("JobPostTitle", "School", "RequiredMaterial")
                    if k in fixed
                }
            out.append(fixed)
        path.write_text(
            json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        summary["files"][name] = {"jobs": len(out), "repaired": n}
        summary["repaired"] += n
        print("  %s/%s repaired=%d/%d" % (region, name, n, len(out)))

    seeds_path = base / "seeds.json"
    if seeds_path.exists():
        payload = json.loads(seeds_path.read_text(encoding="utf-8"))
        seeds = payload.get("seeds") if isinstance(payload, dict) else payload
        if isinstance(seeds, list):
            new_seeds = []
            n = 0
            for job in seeds:
                url = job.get("JobPostUrl") or job.get("ApplicationLink") or ""
                if url in cache:
                    fixed = dict(job)
                    fixed.update(cache[url])
                    if fixed != job:
                        n += 1
                    new_seeds.append(fixed)
                else:
                    fixed, changed = repair_job(job)
                    if changed:
                        n += 1
                        cache[url] = {
                            k: fixed[k]
                            for k in ("JobPostTitle", "School", "RequiredMaterial")
                            if k in fixed
                        }
                    new_seeds.append(fixed)
            if isinstance(payload, dict):
                payload["seeds"] = new_seeds
                payload["count"] = len(new_seeds)
                seeds_path.write_text(
                    json.dumps(payload, ensure_ascii=False, indent=2),
                    encoding="utf-8",
                )
            else:
                seeds_path.write_text(
                    json.dumps(new_seeds, ensure_ascii=False, indent=2),
                    encoding="utf-8",
                )
            print("  %s/seeds.json repaired=%d/%d" % (region, n, len(new_seeds)))

    verified = base / "discovered_jobs_verified.json"
    if verified.exists():
        jobs = json.loads(verified.read_text(encoding="utf-8"))
        rows = [job_to_row(j) for j in jobs]
        CLEAN_ECON.mkdir(parents=True, exist_ok=True)
        write_xlsx(
            rows,
            CLEAN_ECON / ("%s_listings_verified_%s.xlsx" % (region, TODAY)),
            region.title(),
        )
    return summary


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--region", action="append", choices=list(REGIONS))
    ap.add_argument("--all", action="store_true")
    args = ap.parse_args()
    targets = list(REGIONS) if args.all else (args.region or [])
    if not targets:
        for region in REGIONS:
            path = DATA_ECON / region / "discovered_jobs_verified.json"
            if not path.exists():
                continue
            jobs = json.loads(path.read_text(encoding="utf-8"))
            if any(looks_corrupted(j.get("JobPostTitle")) for j in jobs):
                targets.append(region)
    if not targets:
        raise SystemExit("No corrupted regions found; pass --region or --all")
    ensure_dirs()
    summary = {"date": TODAY, "regions": {}}
    for region in targets:
        print("=== repair", region, "===")
        summary["regions"][region] = repair_region(region)
    out = DATA_ECON / ("repair_text_summary_%s.json" % TODAY)
    out.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
