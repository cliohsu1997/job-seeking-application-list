# -*- coding: utf-8 -*-
"""One-by-one university search: home/HR/econ units -> career hubs -> job links.

Works for any coastal region (guangdong / fujian / zhejiang / shenzhen).

Usage:
  python discover_by_university.py --region guangdong
  python discover_by_university.py --region fujian --no-merge
"""
from __future__ import annotations

import argparse
import json
import re
import time
from datetime import date
from pathlib import Path
from urllib.parse import urlparse

from china_universities_data import REGIONS
from discover_dept_jobs import (
    classify_job,
    extract_links,
    is_career_link,
    is_job_announcement,
    same_site,
)
from download_china_provinces import fetch, job_to_row, slug, write_xlsx
from download_china_universities import load_region_universities, write_catalog
from paths import CLEAN_ECON, DATA_ECON, ensure_dirs

TODAY = date.today().isoformat()
SLEEP = 0.28

# CJK via unicode escapes (Windows source encoding safe)
ECON_UNIT_KW = re.compile(
    "("
    "\u7ecf\u6d4e\u5b66\u9662|\u7ecf\u6d4e\u4e0e|\u7ecf\u6d4e\u7ba1\u7406|\u7ecf\u7ba1\u5b66\u9662|"
    "\u5546\u5b66\u9662|\u91d1\u878d\u5b66\u9662|\u8d22\u7ecf\u5b66\u9662|"
    "\u5cad\u5357\u5b66\u9662|\u5e94\u7528\u7ecf\u6d4e|\u7406\u8bba\u7ecf\u6d4e|"
    "\u7edf\u8ba1\u5b66\u9662|\u8d38\u6613\u5b66\u9662|\u56fd\u9645\u91d1\u878d|"
    "School of Economics|College of Economics|Business School"
    ")",
    re.I,
)
ECON_REJECT = re.compile(
    "("
    "\u9a6c\u514b\u601d\u4e3b\u4e49|\u4f53\u80b2|\u827a\u672f|\u97f3\u4e50|"
    "\u9644\u5c5e|\u533b\u9662|\u5e7c\u513f\u56ed|\u4e2d\u5b66|\u5c0f\u5b66"
    ")",
    re.I,
)
STUDENT_PORTAL = re.compile(
    r"(career\.|jy\.|\u5c31\u4e1a|jobfair|\u62db\u8058\u4f1a)",
    re.I,
)
ECON_TITLE_KW = re.compile(
    "(\u7ecf\u6d4e|\u91d1\u878d|\u7ecf\u7ba1|\u5546\u5b66\u9662|\u5cad\u5357)",
    re.I,
)


def _host_root(url: str) -> str | None:
    try:
        return urlparse(url).netloc.lower().removeprefix("www.") or None
    except Exception:
        return None


def guess_econ_urls(home: str) -> list[str]:
    if not home:
        return []
    host = _host_root(home)
    if not host:
        return []
    parts = host.split(".")
    if len(parts) >= 3 and parts[0] in ("www", "www2", "www1"):
        base = ".".join(parts[1:])
    else:
        base = host
    return [
        f"https://jjxy.{base}/",
        f"https://em.{base}/",
        f"https://sem.{base}/",
        f"https://soe.{base}/",
        f"https://econ.{base}/",
        f"https://eco.{base}/",
        f"https://bs.{base}/",
        f"https://cec.{base}/",
        f"https://sef.{base}/",
        f"https://finance.{base}/",
    ]


def guess_hr_urls(home: str, career: str | None = None) -> list[str]:
    urls = []
    if career:
        urls.append(career)
    host = _host_root(home or career or "")
    if not host:
        return urls
    parts = host.split(".")
    if len(parts) >= 3 and parts[0] in ("www", "www2", "www1"):
        base = ".".join(parts[1:])
    else:
        base = host
    for u in (
        f"https://rsc.{base}/",
        f"https://hr.{base}/",
        f"https://www.{base}/rsc/",
        f"https://www.{base}/hr/",
    ):
        if u not in urls:
            urls.append(u)
    return urls


def find_econ_units_from_html(html: str, base_url: str) -> list[dict]:
    links = extract_links(html, base_url)
    found = []
    seen = set()
    for lk in links:
        text = lk["text"] or ""
        url = lk["url"]
        if not ECON_UNIT_KW.search(text) and not ECON_UNIT_KW.search(url):
            continue
        if ECON_REJECT.search(text):
            continue
        if not same_site(url, base_url) and "edu.cn" not in urlparse(url).netloc:
            continue
        key = url.split("#")[0].rstrip("/")
        if key in seen:
            continue
        seen.add(key)
        found.append(
            {
                "unit_zh": text[:80] or "\u7ecf\u6d4e\u76f8\u5173\u5355\u4f4d",
                "unit_en": "",
                "url": key,
            }
        )
    return found[:12]


def _job_rec(
    region, city, university_en, university_zh, school, title, url, job_type
) -> dict:
    return {
        "region": region,
        "Country": "China",
        "Location": city,
        "University": university_en or university_zh,
        "School": school,
        "Deadline": None,
        "JobPostTitle": (title or url)[:180],
        "JobPostUrl": url,
        "RequiredMaterial": "See announcement page",
        "SubmissionChannel": "Dept/HR",
        "ApplicationLink": url,
        "Email": None,
        "JobType": job_type if job_type != "Unknown" else "Academic",
        "_source": "discover_by_university",
    }


def mine_page_for_jobs(
    page_url: str,
    html: str,
    *,
    university_en: str,
    university_zh: str,
    city: str,
    unit_zh: str,
    region: str,
    seen_job_urls: set[str],
    via: str,
) -> list[dict]:
    jobs = []
    links = extract_links(html, page_url) if html else []
    title_guess = unit_zh or page_url
    if is_job_announcement(title_guess, page_url):
        jt = classify_job(title_guess, page_url)
        if page_url not in seen_job_urls and jt != "Unknown":
            seen_job_urls.add(page_url)
            jobs.append(
                _job_rec(
                    region,
                    city,
                    university_en,
                    university_zh,
                    unit_zh,
                    title_guess,
                    page_url,
                    jt,
                )
            )

    career_hubs = [lk for lk in links if is_career_link(lk["text"], lk["url"])]
    for hub in career_hubs[:8]:
        if STUDENT_PORTAL.search(hub["url"]):
            continue
        _h_status, h_html = fetch(hub["url"])
        time.sleep(SLEEP)
        hub_links = extract_links(h_html, hub["url"]) if h_html else []
        if is_job_announcement(hub["text"], hub["url"]):
            jt = classify_job(hub["text"], hub["url"])
            if hub["url"] not in seen_job_urls and jt != "Unknown":
                seen_job_urls.add(hub["url"])
                jobs.append(
                    _job_rec(
                        region,
                        city,
                        university_en,
                        university_zh,
                        unit_zh or hub["text"],
                        hub["text"],
                        hub["url"],
                        jt,
                    )
                )
        for lk in hub_links:
            if not is_job_announcement(lk["text"], lk["url"]):
                continue
            host = urlparse(lk["url"]).netloc.lower()
            if not (
                same_site(lk["url"], page_url)
                or same_site(lk["url"], hub["url"])
                or "gaoxiaojob" in host
                or "shuobojob" in host
            ):
                continue
            blob = f"{lk['text']} {unit_zh}"
            if via == "hr" and not ECON_UNIT_KW.search(blob) and not ECON_TITLE_KW.search(
                lk["text"]
            ):
                continue
            if lk["url"] in seen_job_urls:
                continue
            jt = classify_job(lk["text"], lk["url"])
            if jt == "Unknown":
                continue
            seen_job_urls.add(lk["url"])
            jobs.append(
                _job_rec(
                    region,
                    city,
                    university_en,
                    university_zh,
                    unit_zh or "\u4eba\u4e8b/\u5b66\u9662",
                    lk["text"],
                    lk["url"],
                    jt,
                )
            )
            if len(jobs) >= 30:
                return jobs

    for lk in links:
        if len(jobs) >= 30:
            break
        if not is_job_announcement(lk["text"], lk["url"]):
            continue
        if lk["url"] in seen_job_urls:
            continue
        if not same_site(lk["url"], page_url) and "gaoxiaojob" not in lk["url"]:
            continue
        if via == "hr" and not ECON_UNIT_KW.search(lk["text"]):
            continue
        jt = classify_job(lk["text"], lk["url"])
        if jt == "Unknown":
            continue
        seen_job_urls.add(lk["url"])
        jobs.append(
            _job_rec(
                region,
                city,
                university_en,
                university_zh,
                unit_zh or "\u5b66\u9662",
                lk["text"],
                lk["url"],
                jt,
            )
        )
    return jobs


def discover_university(
    uni: dict,
    region: str,
    raw: Path,
    seen_job_urls: set[str],
) -> dict:
    name_zh = uni["name_zh"]
    name_en = uni.get("name_en") or name_zh
    city = uni.get("city") or ""
    home = uni.get("home") or ""
    tags = uni.get("tags") or []
    print(f"\n[uni] {name_zh} ({', '.join(tags)})")
    if home:
        print(f"      {home}")

    units = list(uni.get("econ_units") or [])
    discovered_units = []

    if home and not units:
        _status, html = fetch(home)
        time.sleep(SLEEP)
        if html and len(html) > 80:
            (raw / f"{slug(name_zh)}_home.html").write_text(
                html, encoding="utf-8", errors="replace"
            )
            units = find_econ_units_from_html(html, home)
            discovered_units = list(units)
            for guess in guess_econ_urls(home):
                if any(u.get("url", "").rstrip("/") == guess.rstrip("/") for u in units):
                    continue
                g_status, g_html = fetch(guess)
                time.sleep(SLEEP)
                if g_status and g_status < 400 and g_html and len(g_html) > 200:
                    units.append(
                        {
                            "unit_zh": "\u7ecf\u6d4e\u76f8\u5173\uff08\u5b50\u57df\u63a2\u6d4b\uff09",
                            "unit_en": "",
                            "url": guess,
                        }
                    )
                    print(f"  [guess-ok] {guess}")

    jobs_all: list[dict] = []
    unit_reports = []
    for unit in units:
        url = unit.get("url")
        if not url:
            continue
        unit_zh = unit.get("unit_zh") or unit.get("unit") or "\u7ecf\u6d4e\u5b66\u9662"
        print(f"  [dept] {unit_zh} -> {url[:70]}")
        status, html = fetch(url)
        time.sleep(SLEEP)
        fname = f"{slug(name_zh)}_{slug(unit_zh)}.html"
        if html and len(html) > 80:
            (raw / fname).write_text(html, encoding="utf-8", errors="replace")
        if unit.get("career"):
            _c_status, c_html = fetch(unit["career"])
            time.sleep(SLEEP)
            if c_html and len(c_html) > 80:
                (raw / f"{slug(name_zh)}_{slug(unit_zh)}_career.html").write_text(
                    c_html, encoding="utf-8", errors="replace"
                )
                jobs_all.extend(
                    mine_page_for_jobs(
                        unit["career"],
                        c_html,
                        university_en=name_en,
                        university_zh=name_zh,
                        city=city,
                        unit_zh=unit_zh,
                        region=region,
                        seen_job_urls=seen_job_urls,
                        via="unit_career",
                    )
                )
        jobs_all.extend(
            mine_page_for_jobs(
                url,
                html or "",
                university_en=name_en,
                university_zh=name_zh,
                city=city,
                unit_zh=unit_zh,
                region=region,
                seen_job_urls=seen_job_urls,
                via="dept",
            )
        )
        unit_reports.append(
            {"unit_zh": unit_zh, "url": url, "status": status, "jobs": len(jobs_all)}
        )

    for hr_url in guess_hr_urls(home, uni.get("career"))[:3]:
        print(f"  [hr] {hr_url[:70]}")
        _status, html = fetch(hr_url)
        time.sleep(SLEEP)
        if not html or len(html) < 80:
            continue
        (raw / f"{slug(name_zh)}_hr_{slug(hr_url)}.html").write_text(
            html, encoding="utf-8", errors="replace"
        )
        before = len(jobs_all)
        jobs_all.extend(
            mine_page_for_jobs(
                hr_url,
                html,
                university_en=name_en,
                university_zh=name_zh,
                city=city,
                unit_zh="\u4eba\u4e8b\u5904/\u4eba\u624d\u62db\u8058",
                region=region,
                seen_job_urls=seen_job_urls,
                via="hr",
            )
        )
        print(f"       +{len(jobs_all) - before} jobs")

    print(f"  -> {name_zh}: units={len(units)} jobs={len(jobs_all)}")
    return {
        "name_zh": name_zh,
        "name_en": name_en,
        "tags": tags,
        "units_used": len(units),
        "units_discovered": len(discovered_units),
        "jobs": jobs_all,
        "unit_reports": unit_reports,
    }


def discover_region(region: str, merge_existing: bool = True) -> dict:
    ensure_dirs()
    write_catalog([region])
    unis = load_region_universities(region)
    base = DATA_ECON / region
    raw = base / "raw" / "job_mine_by_uni"
    raw.mkdir(parents=True, exist_ok=True)

    print(f"=== BY-UNIVERSITY {region.upper()} -- {len(unis)} schools ===")
    seen_job_urls: set[str] = set()
    discovered: list[dict] = []

    if merge_existing:
        for fname in ("discovered_jobs_candidates.json", "discovered_jobs.json"):
            path = base / fname
            if not path.exists():
                continue
            try:
                prev = json.loads(path.read_text(encoding="utf-8"))
            except Exception:
                continue
            for j in prev:
                url = j.get("JobPostUrl") or j.get("ApplicationLink")
                if url and url not in seen_job_urls:
                    seen_job_urls.add(url)
                    discovered.append(j)
            break

    uni_reports = []
    for uni in unis:
        report = discover_university(uni, region, raw, seen_job_urls)
        uni_reports.append({k: v for k, v in report.items() if k != "jobs"})
        for j in report["jobs"]:
            url = j.get("JobPostUrl")
            if url and not any(x.get("JobPostUrl") == url for x in discovered):
                discovered.append(j)

    clean_jobs = [{k: v for k, v in j.items() if not k.startswith("_")} for j in discovered]
    by_url = {j["JobPostUrl"]: j for j in clean_jobs if j.get("JobPostUrl")}
    merged = list(by_url.values())

    n_acad = sum(1 for j in merged if j.get("JobType") == "Academic")
    n_post = sum(1 for j in merged if j.get("JobType") == "Postdoc")

    payload = {
        "date": TODAY,
        "region": region,
        "universities_searched": len(unis),
        "exact_job_links": len(merged),
        "academic": n_acad,
        "postdoc": n_post,
        "universities": uni_reports,
    }
    (base / "job_links_by_university.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (base / "discovered_jobs_candidates.json").write_text(
        json.dumps(merged, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (base / "discovered_jobs.json").write_text(
        json.dumps(merged, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    rows = [job_to_row(j) for j in merged]
    write_xlsx(rows, base / f"listings_by_uni_{TODAY}.xlsx", region.title())
    CLEAN_ECON.mkdir(parents=True, exist_ok=True)
    write_xlsx(rows, CLEAN_ECON / f"{region}_listings_by_uni_{TODAY}.xlsx", region.title())

    summary = {
        "universities_searched": len(unis),
        "exact_job_links": len(merged),
        "academic": n_acad,
        "postdoc": n_post,
    }
    print(
        f"\n-> {region}: schools={len(unis)} jobs={len(merged)} "
        f"(Academic={n_acad}, Postdoc={n_post})"
    )
    return summary


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument(
        "--region",
        action="append",
        choices=list(REGIONS),
        required=True,
        help="Region to mine (repeatable).",
    )
    ap.add_argument(
        "--no-merge",
        action="store_true",
        help="Do not merge prior candidates; start fresh from university loop.",
    )
    args = ap.parse_args()
    all_summary = {"date": TODAY, "regions": {}}
    for region in args.region:
        all_summary["regions"][region] = discover_region(
            region, merge_existing=not args.no_merge
        )
    (DATA_ECON / f"by_university_summary_{TODAY}.json").write_text(
        json.dumps(all_summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print("Done.")


if __name__ == "__main__":
    main()
