"""
Pipeline for all China coastal regions (guangdong / fujian / zhejiang / shenzhen):

  1. Start from economics departments (UNIVERSITIES catalog)
  2. Scrape each department page
  3. Detect career-related links (招聘 / 博士后 / 人才引进 / ...)
  4. Follow career hubs → extract exact job announcement links
  5. Write job_links.json + discovered_jobs.json + Application-column xlsx

Usage:
  python discover_dept_jobs.py                  # all regions
  python discover_dept_jobs.py --region guangdong
"""
from __future__ import annotations

import argparse
import json
import re
import time
from datetime import date
from html.parser import HTMLParser
from pathlib import Path
from urllib.parse import urljoin, urlparse

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from columns import COLUMNS
from download_china_provinces import (
    REGIONS,
    UNIVERSITIES,
    fetch,
    job_to_row,
    parse_deadline,
    slug,
    write_xlsx,
)
from paths import CLEAN_ECON, DATA_ECON, ensure_dirs

TODAY = date.today().isoformat()
SLEEP = 0.3

# 简体中文-first keywords (faculty + postdoc)
CAREER_KW = re.compile(
    r"(人才招聘|教师招聘|人才引进|诚聘|博士后招聘|博士后|师资博士后|"
    r"人事处|人力资源|加入我们|招贤纳士|师资引进|"
    r"师资招聘|招聘信息|人才网|英才招聘)",
    re.I,
)
JOB_ANN_KW = re.compile(
    r"(教师招聘|公开招聘|招聘教师|研究型教师|人才引进|诚聘|招聘启事|招聘公告|"
    r"专任教师|教研岗|全球招聘|海内外|"
    r"博士后招聘|博士后研究人员|师资博士后|博士后招收|招聘博士后)",
    re.I,
)
POSTDOC_KW = re.compile(r"(博士后|师资博士后)", re.I)
FACULTY_KW = re.compile(r"(教师招聘|专任教师|教研岗|人才引进|诚聘|教授|副教授|讲师)", re.I)
REJECT_TITLE = re.compile(
    r"(拟聘|公示|录用名单|表彰|荣休|慰问|论坛|讲座|研讨会|支部|"
    r"工作例会|招聘会|就业|校园招聘|秋招|春招|毕业生|"
    r"附属.*(中学|小学|幼儿园)|实验学校|转发|管理办法|"
    r"行政岗|行政秘书|秘书岗|科研助理|市场招生|辅导员|"
    r"纪检监察|马克思主义|"
    r"^博士后管理$|^招聘信息$|^人才招聘$|^招聘公告$|^通知公告$|"
    r"^更多$|^详情$|^阅读全文$)",
    re.I,
)
STUDENT_PORTAL = re.compile(r"(career\.|jy\.|就业|jobfair|招聘会)", re.I)
SKIP_HREF = re.compile(
    r"(javascript:|#$|\.pdf$|\.doc|\.xls|\.zip|\.jpg|\.png|mailto:|tel:)",
    re.I,
)
DETAIL_URL = re.compile(
    r"(article/\d+|info/\d+/\d+|/c\d+a\d+|content|detail|view\.php|"
    r"announcement|page\.htm|/\d{4}/\d{2}/)",
    re.I,
)


class LinkParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.links: list[tuple[str, str]] = []
        self._href: str | None = None
        self._parts: list[str] = []

    def handle_starttag(self, tag, attrs):
        if tag != "a":
            return
        href = dict(attrs).get("href")
        if href:
            self._href = href.strip()
            self._parts = []

    def handle_data(self, data):
        if self._href is not None:
            self._parts.append(data)

    def handle_endtag(self, tag):
        if tag == "a" and self._href is not None:
            text = re.sub(r"\s+", " ", "".join(self._parts)).strip()
            self.links.append((self._href, text))
            self._href = None
            self._parts = []


def extract_links(html: str, base_url: str) -> list[dict]:
    p = LinkParser()
    try:
        p.feed(html or "")
    except Exception:
        return []
    out = []
    seen = set()
    for href, text in p.links:
        if not href or SKIP_HREF.search(href):
            continue
        abs_url = urljoin(base_url, href)
        if not abs_url.startswith("http"):
            continue
        key = (abs_url.split("#")[0], text[:120])
        if key in seen:
            continue
        seen.add(key)
        out.append({"url": abs_url.split("#")[0], "text": text[:200]})
    return out


def same_site(a: str, b: str) -> bool:
    try:
        ha = urlparse(a).netloc.lower().removeprefix("www.")
        hb = urlparse(b).netloc.lower().removeprefix("www.")
        return ha == hb or ha.endswith("." + hb) or hb.endswith("." + ha)
    except Exception:
        return False


def classify_job(title: str, url: str) -> str:
    blob = f"{title} {url}"
    if POSTDOC_KW.search(blob) and not FACULTY_KW.search(title):
        return "Postdoc"
    if POSTDOC_KW.search(blob) and FACULTY_KW.search(blob):
        # announcement covering both — keep as Academic if teacher primary, else Postdoc
        if re.search(r"博士后", title) and not re.search(r"教师|教授", title):
            return "Postdoc"
        if re.search(r"教师|教授|教研", title):
            return "Academic"
        return "Postdoc"
    if FACULTY_KW.search(blob) or JOB_ANN_KW.search(blob):
        return "Academic"
    return "Unknown"


def is_career_link(text: str, url: str) -> bool:
    blob = f"{text} {url}"
    if STUDENT_PORTAL.search(blob):
        return False
    host = url.split("/")[2] if "://" in url else ""
    if host.startswith("jobs.") or ".jobs." in host:
        return True
    return bool(CAREER_KW.search(blob))


def is_job_announcement(text: str, url: str) -> bool:
    text = (text or "").strip()
    if len(text) < 6:
        return False
    if REJECT_TITLE.search(text):
        return False
    if STUDENT_PORTAL.search(f"{text} {url}"):
        return False
    if not JOB_ANN_KW.search(text) and not (
        POSTDOC_KW.search(text) and re.search(r"招聘|招收|启事|简章", text)
    ):
        return False
    # Final rows must be detail-like URLs, not list hubs
    if re.search(r"(list\.htm|list\.html|/list/?$|/zp/?$)", url, re.I):
        return False
    return bool(DETAIL_URL.search(url))


def region_depts(region: str) -> list[dict]:
    rows = []
    for u in UNIVERSITIES[region]:
        for unit in u.get("econ_units") or []:
            url = unit.get("url")
            if not url:
                continue
            rows.append(
                {
                    "university_zh": u["name_zh"],
                    "university_en": u["name_en"],
                    "city": u.get("city"),
                    "unit_zh": unit.get("unit_zh") or unit.get("unit") or "",
                    "unit_en": unit.get("unit_en") or unit.get("unit") or "",
                    "dept_url": url,
                    "unit_career": unit.get("career"),
                }
            )
    return rows


def discover_region(region: str) -> dict:
    base = DATA_ECON / region
    raw = base / "raw" / "job_mine"
    for d in (base, raw):
        d.mkdir(parents=True, exist_ok=True)

    departments = []
    discovered = []
    seen_job_urls: set[str] = set()

    depts = region_depts(region)
    print(f"=== MINE {region.upper()} — {len(depts)} economics departments ===")

    for dept in depts:
        dept_url = dept["dept_url"]
        name = f"{dept['university_zh']}_{dept['unit_zh']}"
        print(f"\n[dept] {name}")
        print(f"       {dept_url}")
        status, html = fetch(dept_url)
        time.sleep(SLEEP)
        html_path = raw / f"{slug(name)}.html"
        if html and len(html) > 80:
            html_path.write_text(html, encoding="utf-8", errors="replace")

        links = extract_links(html, dept_url) if html else []
        career_hubs = []
        # Only career hubs linked from the department page (简体中文),
        # plus optional unit-level career URL (not university-wide HR).
        for lk in links:
            if is_career_link(lk["text"], lk["url"]):
                career_hubs.append(lk)
        if dept.get("unit_career"):
            career_hubs.insert(
                0,
                {"url": dept["unit_career"], "text": "学院招聘专栏"},
            )

        uniq_hubs = []
        seen_h = set()
        for h in career_hubs:
            if h["url"] in seen_h or STUDENT_PORTAL.search(h["url"]):
                continue
            seen_h.add(h["url"])
            uniq_hubs.append(h)

        jobs = []
        for hub in uniq_hubs[:10]:
            print(f"  [career] {hub['text'][:40]} -> {hub['url'][:70]}")
            h_status, h_html = fetch(hub["url"])
            time.sleep(SLEEP)
            hub_file = raw / f"{slug(name)}_hub_{slug(hub['text'] or 'career')}.html"
            if h_html and len(h_html) > 80:
                hub_file.write_text(h_html, encoding="utf-8", errors="replace")
            hub_links = extract_links(h_html, hub["url"]) if h_html else []

            # Hub itself may be a single announcement
            title_guess = hub["text"] or hub["url"]
            if is_job_announcement(title_guess, hub["url"]):
                jt = classify_job(title_guess, hub["url"])
                if hub["url"] not in seen_job_urls and jt != "Unknown":
                    seen_job_urls.add(hub["url"])
                    jobs.append(
                        {
                            "title": title_guess[:180],
                            "url": hub["url"],
                            "job_type": jt,
                            "via": "hub_is_announcement",
                        }
                    )

            for lk in hub_links:
                if not is_job_announcement(lk["text"], lk["url"]):
                    continue
                host = urlparse(lk["url"]).netloc.lower()
                if not (
                    same_site(lk["url"], dept_url)
                    or same_site(lk["url"], hub["url"])
                    or "gaoxiaojob" in host
                    or "shuobojob" in host
                    or "liuxuehr" in host
                ):
                    continue
                if lk["url"] in seen_job_urls:
                    continue
                jt = classify_job(lk["text"], lk["url"])
                if jt == "Unknown":
                    continue
                seen_job_urls.add(lk["url"])
                jobs.append(
                    {
                        "title": (lk["text"] or lk["url"])[:180],
                        "url": lk["url"],
                        "job_type": jt,
                        "via": hub["url"],
                    }
                )
                print(f"    [job:{jt}] {lk['text'][:60]}")
                if len(jobs) >= 40:
                    break
            if len(jobs) >= 40:
                break

        # Direct announcement links on the dept page
        for lk in links:
            if len(jobs) >= 40:
                break
            if not is_job_announcement(lk["text"], lk["url"]):
                continue
            if lk["url"] in seen_job_urls:
                continue
            if not same_site(lk["url"], dept_url) and "gaoxiaojob" not in lk["url"]:
                continue
            jt = classify_job(lk["text"], lk["url"])
            if jt == "Unknown":
                continue
            seen_job_urls.add(lk["url"])
            jobs.append(
                {
                    "title": (lk["text"] or lk["url"])[:180],
                    "url": lk["url"],
                    "job_type": jt,
                    "via": "dept_page",
                }
            )
            print(f"    [job:{jt}] {lk['text'][:60]}")

        rec = {
            "university_zh": dept["university_zh"],
            "university_en": dept["university_en"],
            "city": dept["city"],
            "unit_zh": dept["unit_zh"],
            "unit_en": dept["unit_en"],
            "dept_url": dept_url,
            "dept_status": status,
            "has_career": bool(uniq_hubs) or bool(jobs),
            "career_hubs": [{"text": h["text"], "url": h["url"]} for h in uniq_hubs],
            "jobs": jobs,
        }
        departments.append(rec)

        for j in jobs:
            discovered.append(
                {
                    "region": region,
                    "Country": "China",
                    "Location": dept["city"],
                    "University": dept["university_en"] or dept["university_zh"],
                    "School": dept["unit_zh"] or dept["unit_en"],
                    "Deadline": None,
                    "JobPostTitle": j["title"],
                    "JobPostUrl": j["url"],
                    "RequiredMaterial": "See announcement page",
                    "SubmissionChannel": "Dept/HR",
                    "ApplicationLink": j["url"],
                    "Email": None,
                    "JobType": j["job_type"] if j["job_type"] != "Unknown" else "Academic",
                }
            )

    n_career = sum(1 for d in departments if d["has_career"])
    n_jobs = sum(len(d["jobs"]) for d in departments)
    n_post = sum(1 for j in discovered if j["JobType"] == "Postdoc")
    n_acad = sum(1 for j in discovered if j["JobType"] == "Academic")

    payload = {
        "date": TODAY,
        "region": region,
        "departments_scraped": len(departments),
        "departments_with_career": n_career,
        "exact_job_links": n_jobs,
        "academic": n_acad,
        "postdoc": n_post,
        "departments": departments,
    }
    (base / "job_links.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (base / f"job_links_{TODAY}.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    # Candidates only — final listing written after validation when --validate is used
    (base / "discovered_jobs_candidates.json").write_text(
        json.dumps(discovered, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (base / "discovered_jobs.json").write_text(
        json.dumps(discovered, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    rows = [job_to_row(j) for j in discovered]
    write_xlsx(rows, base / f"listings_mined_{TODAY}.xlsx", region.title())
    write_xlsx(rows, CLEAN_ECON / f"{region}_listings_mined_{TODAY}.xlsx", region.title())

    summary = {
        "departments_scraped": len(departments),
        "departments_with_career": n_career,
        "exact_job_links": n_jobs,
        "academic": n_acad,
        "postdoc": n_post,
        "job_links": str(base / "job_links.json"),
        "discovered_jobs": str(base / "discovered_jobs.json"),
    }
    print(
        f"\n-> {region}: depts={len(departments)} with_career={n_career} "
        f"jobs={n_jobs} (Academic={n_acad}, Postdoc={n_post})"
    )
    return summary


def main(regions: list[str] | None = None, do_validate: bool = False) -> None:
    ensure_dirs()
    CLEAN_ECON.mkdir(parents=True, exist_ok=True)
    targets = regions or list(REGIONS)
    all_summary = {"date": TODAY, "regions": {}}
    for region in targets:
        if region not in UNIVERSITIES:
            raise SystemExit(f"Unknown region: {region}")
        all_summary["regions"][region] = discover_region(region)
        if do_validate:
            from validate_job_links import validate_region

            all_summary["regions"][region]["validation"] = validate_region(region)

    (DATA_ECON / f"china_coastal_job_mine_summary_{TODAY}.json").write_text(
        json.dumps(all_summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(all_summary, ensure_ascii=False, indent=2))
    print("Done.")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument(
        "--region",
        action="append",
        choices=list(REGIONS),
        help="Mine one region (repeatable). Default: all four.",
    )
    ap.add_argument(
        "--validate",
        action="store_true",
        help="After mining, keep only real faculty/postdoc announcement pages.",
    )
    args = ap.parse_args()
    main(args.region, do_validate=args.validate)
