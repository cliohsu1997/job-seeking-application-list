# -*- coding: utf-8 -*-
"""Validate candidate job URLs: keep only real open faculty / postdoc pages.

Usage:
  python validate_job_links.py --region guangdong
  python validate_job_links.py --all
"""
from __future__ import annotations

import argparse
import json
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date
from pathlib import Path

from urllib.parse import urlparse, urlunparse

from china_universities_data import REGIONS
from download_china_provinces import fetch, job_to_row, write_xlsx
from export_clean_pdf import export_one
from paths import CLEAN, CLEAN_ECON, DATA_ECON, ensure_dirs

TODAY = date.today().isoformat()

# Simplified Chinese patterns via unicode escapes (encoding-safe on Windows).
REAL_SIG = re.compile(
    "("
    "\u5e94\u8058|\u62a5\u540d\u65b9\u5f0f|\u62a5\u540d|\u5c97\u4f4d\u804c\u8d23|"
    "\u4efb\u804c\u8981\u6c42|\u62db\u8058\u6761\u4ef6|\u62db\u8058\u5c97\u4f4d|"
    "\u62db\u8058\u4eba\u6570|\u535a\u58eb\u5b66\u4f4d|\u53d1\u9001\u81f3|"
    "\u8054\u7cfb\u90ae\u7bb1|\u7535\u5b50\u90ae\u7bb1|\u90ae\u7bb1[:\uff1a]|"
    "\u5f85\u9047|\u85aa\u916c|\u5c97\u4f4d\u8981\u6c42|"
    "apply|deadline|qualification|requirements|curriculum vitae|\\bcv\\b"
    ")",
    re.I,
)
FACULTY_SIG = re.compile(
    "("
    "\u6559\u5e08\u62db\u8058|\u4e13\u4efb\u6559\u5e08|\u6559\u7814\u5c97|"
    "\u4eba\u624d\u5f15\u8fdb|\u8bda\u8058|\u6559\u6388|\u526f\u6559\u6388|\u8bb2\u5e08|"
    "faculty|tenure|assistant professor|associate professor"
    ")",
    re.I,
)
POSTDOC_SIG = re.compile(
    "(\u535a\u58eb\u540e|\u5e08\u8d44\u535a\u58eb\u540e|post[- ]?doc)",
    re.I,
)
ECON_SIG = re.compile(
    "("
    "\u7ecf\u6d4e|\u91d1\u878d|\u5e94\u7528\u7ecf\u6d4e\u5b66|\u7406\u8bba\u7ecf\u6d4e\u5b66|"
    "\u6570\u5b57\u7ecf\u6d4e|\u8d22\u653f|\u56fd\u9645\u8d38\u6613|"
    "\u5cad\u5357\u5b66\u9662|\u7ecf\u6d4e\u5b66\u9662|\u7ecf\u6d4e\u4e0e|\u5546\u5b66\u9662|\u7ecf\u7ba1|"
    "\u7ba1\u7406\u5b66\u9662|\u521b\u4e1a\u4e0e\u7ba1\u7406|\u5de5\u5546\u7ba1\u7406|"
    "\u7ec4\u7ec7\u7ba1\u7406|\u8425\u9500\u7cfb|\u4f1a\u8ba1\u7cfb|Antai|SEM\\b|ACEM"
    ")",
    re.I,
)
REJECT_TITLE = re.compile(
    "("
    "\u62df\u8058|\u516c\u793a|\u5f55\u7528\u540d\u5355|\u8868\u5f70|\u8363\u4f11|\u6170\u95ee|"
    "\u8bba\u575b|\u8bb2\u5ea7|\u7814\u8ba8\u4f1a|\u652f\u90e8|"
    "\u5de5\u4f5c\u4f8b\u4f1a|\u62db\u8058\u4f1a|\u5c31\u4e1a|\u6821\u56ed\u62db\u8058|"
    "\u79cb\u62db|\u6625\u62db|\u6bd5\u4e1a\u751f|"
    "\u9644\u5c5e.*(\u4e2d\u5b66|\u5c0f\u5b66|\u5e7c\u513f\u56ed)|\u5b9e\u9a8c\u5b66\u6821|"
    "\u8f6c\u53d1|\u7ba1\u7406\u529e\u6cd5|"
    "\u7eaa\u68c0\u76d1\u5bdf|\u9a6c\u514b\u601d\u4e3b\u4e49|\u8f85\u5bfc\u5458|"
    "\u884c\u653f\u5c97|\u884c\u653f\u79d8\u4e66|\u79d8\u4e66\u5c97|\u79d1\u7814\u52a9\u7406|\u5e02\u573a\u62db\u751f|"
    "^\u535a\u58eb\u540e\u7ba1\u7406$|^\u62db\u8058\u4fe1\u606f$|^\u4eba\u624d\u62db\u8058$|"
    "^\u62db\u8058\u516c\u544a$|^\u901a\u77e5\u516c\u544a$|"
    "^\u66f4\u591a$|^\u8be6\u60c5$|^\u9605\u8bfb\u5168\u6587$|"
    "\u62a5\u5230\u987b\u77e5|\u8fdb\u7ad9\u5b58\u6863"
    ")",
    re.I,
)
REJECT_BODY = re.compile(
    "("
    "\u62df\u8058\u4eba\u5458\u516c\u793a|\u516c\u5f00\u9009\u8058.*\u516c\u793a|"
    "\u5f55\u7528\u540d\u5355|\u6559\u5e08\u8282|\u62db\u8058\u4f1a|"
    "\u5c31\u4e1a\u6307\u5bfc|\u6821\u56ed\u5ba3\u8bb2"
    ")",
    re.I,
)
LIST_URL = re.compile(r"(list\.htm|list\.html|/list/?$|/zp/?$|Recruitment/?$)", re.I)
DETAIL_URL = re.compile(
    r"(article/\d+|info/\d+/\d+|/c\d+a\d+|content\.|/detail|"
    r"view\.php|announcement/detail|page\.htm|/\d{4}/\d{2}/|"
    r"/acem/hr/\d+|Ads/Ads\?|/Home/Detail/)",
    re.I,
)
STALE_YEAR = re.compile(r"(201[0-9]|202[0-3])")
FRESH_YEAR = re.compile(
    "(202[5-6]|\u5e38\u5e74\u62db\u8058|\u957f\u671f\u62db\u8058|"
    "\u672c\u62db\u8058\u4fe1\u606f\u957f\u671f\u6709\u6548|\u62db\u6ee1\u5373\u6b62)"
)
ROLLING_OK = re.compile(
    "(\u957f\u671f\u6709\u6548|\u5e38\u5e74\u62db\u8058|\u957f\u671f\u62db\u8058|"
    "\u5e38\u5e74\u63a5\u6536|\u6eda\u52a8\u62db\u8058|\u968f\u65f6\u7533\u8bf7)"
)
DEADLINE_RE = re.compile(
    "(?:"
    "\u622a\u6b62\u65e5\u671f|\u62a5\u540d\u622a\u6b62|\u622a\u6b62\u65f6\u95f4|"
    "\u8bf7\u4e8e|deadline"
    ")[:\uff1a\\s]*"
    "(20\\d{2})[\\-/.\\u5e74](\\d{1,2})[\\-/.\\u6708]?(\\d{1,2})?\u65e5?\u524d?",
    re.I,
)
POSTED_RE = re.compile(
    "(20\\d{2})[\\-/.\\u5e74](\\d{1,2})[\\-/.\\u6708](\\d{1,2})"
)
STUDENT_HOST = re.compile(
    r"(career\.|jy\.|/jobfair|\u5c31\u4e1a)",
    re.I,
)
UNI_HR_HOST = re.compile(r"(^rsc\.|^hr\.|^hrdam\.|/rsc/)", re.I)
# Dean / center-director posts are not faculty/postdoc openings.
DEAN_ONLY = re.compile(
    "("
    "\u4e8c\u7ea7\u5b66\u9662\u9662\u957f|\u5b66\u9662\u9662\u957f|"
    "\u5b9e\u9a8c\u6559\u5b66\u4e2d\u5fc3\u4e3b\u4efb|\u9662\u957f\u516c\u544a"
    ")",
    re.I,
)
FACULTY_OR_POSTDOC_TITLE = re.compile(
    "("
    "\u6559\u5e08|\u535a\u58eb\u540e|\u4eba\u624d\u5f15\u8fdb|\u4e13\u4efb|"
    "\u6559\u7814\u5c97|\u5e08\u8d44\u535a\u58eb\u540e"
    ")",
    re.I,
)
STALE_YEAR_IN_TITLE = re.compile("2024\u5e74")
FRESH_YEAR_IN_TITLE = re.compile("202[5-6]")
# Title cycles from 2010-2023 are hard-stale (body footers must not clear them).
OLD_CYCLE_TITLE = re.compile(r"(201[0-9]|202[0-3])\u5e74")


def strip_html(html: str) -> str:
    text = re.sub(r"(?is)<script[^>]*>.*?</script>", " ", html or "")
    text = re.sub(r"(?is)<style[^>]*>.*?</style>", " ", text)
    text = re.sub(r"<[^>]+>", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def classify_type(title: str, body: str) -> str:
    # Prefer title; body alone often mentions both faculty and postdoc tracks.
    # Postdoc-primary titles first (e.g. ???????).
    if re.search(
        "\u535a\u58eb\u540e\u5168\u7403\u62db\u8058|\u62db\u8058\u535a\u58eb\u540e|"
        "\u535a\u58eb\u540e\u62db\u8058|\u5e08\u8d44\u535a\u58eb\u540e|"
        "\u79d1\u7814\u535a\u58eb\u540e|\u535a\u58eb\u540e\u6d41\u52a8\u7ad9",
        title,
    ):
        return "Postdoc"
    # Faculty-first titles (even if they also mention postdoc tracks).
    if re.search(
        "\u6559\u5e08\u62db\u8058|\u62db\u8058\u6559\u5e08|\u4eba\u624d\u62db\u8058|"
        "\u4eba\u624d\u5f15\u8fdb|\u8bda\u8058\u6d77\u5185\u5916|\u5168\u7403\u62db\u8058|"
        "\u9ad8\u5c42\u6b21\u4eba\u624d|\u9762\u5411\u6d77\u5185\u5916\u62db\u8058\u4eba\u624d",
        title,
    ):
        return "Academic"
    if re.search("\u535a\u58eb\u540e", title):
        return "Postdoc"
    if FACULTY_SIG.search(title) or re.search(
        "\u62db\u8058\u542f\u4e8b|\u62db\u8058\u516c\u544a|\u8bda\u8058|"
        "\u7814\u7a76\u578b\u6559\u5e08",
        title,
    ):
        return "Academic"
    blob = body[:3000]
    if POSTDOC_SIG.search(blob) and not FACULTY_SIG.search(blob):
        return "Postdoc"
    if FACULTY_SIG.search(blob) or POSTDOC_SIG.search(blob):
        return "Academic" if FACULTY_SIG.search(title + blob[:800]) else "Postdoc"
    return "Unknown"


def score_announcement(
    html: str | None,
    url: str,
    title: str,
    unit_zh: str = "",
    status: int | None = None,
) -> dict:
    reasons: list[str] = []
    title = (title or "").strip()
    body = strip_html(html or "")[:12000]
    blob = f"{title} {body}"

    if status is None or (isinstance(status, int) and status >= 400 and status != 403):
        return {
            "keep": False,
            "label": "dead",
            "job_type": "Unknown",
            "score": 0,
            "reasons": ["fetch_fail status=%s" % status],
        }
    if len(html or "") < 500 or len(body) < 80:
        return {
            "keep": False,
            "label": "empty",
            "job_type": "Unknown",
            "score": 0,
            "reasons": ["body_too_small"],
        }
    if STUDENT_HOST.search(url) or STUDENT_HOST.search(title):
        return {
            "keep": False,
            "label": "student_career",
            "job_type": "Unknown",
            "score": 0,
            "reasons": ["student_employment_portal"],
        }
    if REJECT_TITLE.search(title):
        return {
            "keep": False,
            "label": "rejected_title",
            "job_type": "Unknown",
            "score": 0,
            "reasons": ["reject_title_pattern"],
        }
    if DEAN_ONLY.search(title) and not FACULTY_OR_POSTDOC_TITLE.search(title):
        return {
            "keep": False,
            "label": "dean_admin",
            "job_type": "Unknown",
            "score": 0,
            "reasons": ["dean_or_admin_post"],
        }
    # Hard reject year-coded cycles in the title (2010-2023), incl. postdocs.
    # Do not let incidental fresh years in page footers clear this.
    if OLD_CYCLE_TITLE.search(title) and not ROLLING_OK.search(title + " " + body[:800]):
        return {
            "keep": False,
            "label": "stale_year_title",
            "job_type": "Unknown",
            "score": 0,
            "reasons": ["title_year_pre_2024"],
        }
    # Faculty ads whose title is only a 2024 cycle (no 2025/2026) are stale.
    if (
        STALE_YEAR_IN_TITLE.search(title)
        and not FRESH_YEAR_IN_TITLE.search(title)
        and not POSTDOC_SIG.search(title)
    ):
        return {
            "keep": False,
            "label": "stale_year_title",
            "job_type": "Unknown",
            "score": 0,
            "reasons": ["title_year_2024_faculty"],
        }
    if REJECT_BODY.search(title) or (
        REJECT_BODY.search(body) and not REAL_SIG.search(body)
    ):
        return {
            "keep": False,
            "label": "rejected_body",
            "job_type": "Unknown",
            "score": 0,
            "reasons": ["not_open_announcement"],
        }
    if LIST_URL.search(url) and not DETAIL_URL.search(url):
        return {
            "keep": False,
            "label": "list_hub",
            "job_type": "Unknown",
            "score": 0,
            "reasons": ["url_is_list_hub"],
        }

    real_hits = len(REAL_SIG.findall(blob))
    job_type = classify_type(title, body)
    # Do NOT trust School attribution alone (HR pages get misc assigned to econ units).
    # Require econ / unit mention in title, or strong econ signal early in the body.
    title_econ = bool(ECON_SIG.search(title) or (unit_zh and unit_zh and unit_zh in title))
    body_econ = bool(ECON_SIG.search(body[:2500]))
    unit_in_body = bool(unit_zh and unit_zh in body[:2500])
    # Trust curated School labels from seeds/catalog (e.g. SEM / Antai / ???).
    school_econ = bool(unit_zh and ECON_SIG.search(unit_zh))
    econ_ok = title_econ or unit_in_body or school_econ or (
        body_econ and re.search(
            "\u7ecf\u6d4e\u5b66\u9662|\u7ecf\u6d4e\u4e0e|\u5cad\u5357|\u91d1\u878d\u5b66\u9662|"
            "\u5e94\u7528\u7ecf\u6d4e|\u7406\u8bba\u7ecf\u6d4e|\u6570\u5b57\u7ecf\u6d4e|"
            "\u7ba1\u7406\u5b66\u9662|\u5546\u5b66\u9662|\u7ecf\u7ba1",
            body[:2500],
        )
    )
    # University-wide postdoc/talent ads that never name econ still fail
    if re.search(
        "\u7eaa\u68c0|\u9a6c\u514b\u601d|\u8f85\u5bfc\u5458|\u9644\u5c5e|\u5b9e\u9a8c\u5b66\u6821",
        title,
    ):
        econ_ok = False

    # Campus HR hosts: keep if econ is an explicit hired field (title or early body)
    host = url.split("/")[2] if "://" in url else url
    if UNI_HR_HOST.search(host) or host.startswith("rsc.") or ".rsc." in host:
        title_or_lead = title + " " + body[:2000]
        if not (
            ECON_SIG.search(title)
            or (unit_zh and unit_zh in title)
            or re.search(
                "\u7ecf\u6d4e\u5b66|\u5e94\u7528\u7ecf\u6d4e|\u7406\u8bba\u7ecf\u6d4e|"
                "\u6570\u5b57\u7ecf\u6d4e|\u91d1\u878d\u5b66|\u7ecf\u8d38",
                title_or_lead,
            )
        ):
            econ_ok = False
            reasons.append("campus_hr_not_econ")
        else:
            econ_ok = True
            reasons.append("campus_hr_econ_field")

    rolling = bool(ROLLING_OK.search(blob))
    fresh = bool(FRESH_YEAR.search(blob) or FRESH_YEAR.search(url) or rolling)
    stale = bool(STALE_YEAR.search(title)) and not fresh

    # Hard deadline in the past => reject (e.g. ??2026?4?15??)
    today = date.today()
    for m in DEADLINE_RE.finditer(blob):
        try:
            y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3) or 1)
            if date(y, mo, d) < today and not rolling:
                reasons.append("deadline_passed_%s-%s-%s" % (y, mo, d))
                return {
                    "keep": False,
                    "label": "expired_deadline",
                    "job_type": job_type,
                    "score": 0,
                    "reasons": reasons,
                }
        except ValueError:
            pass

    # Old post date without rolling language
    posted_years = [int(y) for y, _mo, _d in POSTED_RE.findall(blob[:1500])]
    if posted_years and max(posted_years) <= 2023 and not rolling:
        reasons.append("posted_year_too_old")
        return {
            "keep": False,
            "label": "stale",
            "job_type": job_type,
            "score": 0,
            "reasons": reasons,
        }
    # Year-coded cycle ads like 2025?? when we are late in 2026
    if re.search("2025\u5e74\u5ea6", title) and today.year >= 2026 and today.month >= 7:
        open_yearround = bool(
            re.search(
                "\u5168\u5e74\u5f00\u653e|\u5e38\u5e74|\u957f\u671f\u6709\u6548|\u62db\u6ee1\u5373\u6b62",
                blob,
            )
        )
        if not rolling and not open_yearround:
            reasons.append("prior_year_cycle")
            return {
                "keep": False,
                "label": "stale_cycle",
                "job_type": job_type,
                "score": 0,
                "reasons": reasons,
            }

    score = real_hits * 2
    if job_type in ("Academic", "Postdoc"):
        score += 3
    if econ_ok:
        score += 3
    if DETAIL_URL.search(url):
        score += 2
    if fresh:
        score += 2
    if stale:
        score -= 4
        reasons.append("possibly_stale")

    if real_hits < 2:
        reasons.append("few_application_signals")
        return {
            "keep": False,
            "label": "weak_job_signal",
            "job_type": job_type,
            "score": score,
            "reasons": reasons,
        }
    if job_type == "Unknown":
        reasons.append("not_faculty_or_postdoc")
        return {
            "keep": False,
            "label": "not_target_role",
            "job_type": job_type,
            "score": score,
            "reasons": reasons,
        }
    if not econ_ok:
        reasons.append("not_econ_related")
        return {
            "keep": False,
            "label": "wrong_field",
            "job_type": job_type,
            "score": score,
            "reasons": reasons,
        }
    if stale and real_hits < 4:
        reasons.append("stale_closed")
        return {
            "keep": False,
            "label": "stale",
            "job_type": job_type,
            "score": score,
            "reasons": reasons,
        }

    label = "real_announcement" if real_hits >= 3 else "likely_real"
    return {
        "keep": True,
        "label": label,
        "job_type": job_type,
        "score": score,
        "reasons": reasons or ["ok"],
    }


def validate_job(job: dict) -> dict:
    url = job.get("JobPostUrl") or job.get("ApplicationLink") or ""
    title = job.get("JobPostTitle") or ""
    unit = job.get("School") or ""
    status, html = fetch(url, timeout=18)
    time.sleep(0.15)
    result = score_announcement(html, url, title, unit_zh=unit, status=status)
    out = dict(job)
    out["JobType"] = (
        result["job_type"]
        if result["job_type"] != "Unknown"
        else job.get("JobType") or "Academic"
    )
    out["_validate"] = {
        "keep": result["keep"],
        "label": result["label"],
        "score": result["score"],
        "reasons": result["reasons"],
        "http_status": status,
        "bytes": len(html or ""),
    }
    return out


def normalize_job_url(url: str) -> str:
    """Collapse http/https and trailing-slash variants for dedupe."""
    p = urlparse((url or "").strip())
    host = (p.netloc or "").lower()
    if host.startswith("www."):
        host = host[4:]
    path = p.path.rstrip("/") or "/"
    return urlunparse(("https", host, path, "", "", ""))


def polish_kept(jobs: list[dict]) -> list[dict]:
    """Dedupe http/https twins; prefer https and richer metadata."""
    by_key: dict[str, dict] = {}
    for j in jobs:
        key = normalize_job_url(j.get("JobPostUrl") or j.get("ApplicationLink") or "")
        if not key or key == "https:///":
            continue
        prev = by_key.get(key)
        if not prev:
            by_key[key] = j
            continue
        cur_https = (j.get("JobPostUrl") or "").startswith("https")
        prev_https = (prev.get("JobPostUrl") or "").startswith("https")
        cur_len = len(j.get("JobPostTitle") or "")
        prev_len = len(prev.get("JobPostTitle") or "")
        if (cur_https and not prev_https) or (cur_https == prev_https and cur_len > prev_len):
            by_key[key] = j
    out = list(by_key.values())
    out.sort(
        key=lambda j: (
            j.get("University") or "",
            j.get("JobPostTitle") or "",
            j.get("JobPostUrl") or "",
        )
    )
    return out


def validate_region(region: str, workers: int = 6) -> dict:
    base = DATA_ECON / region
    src = base / "discovered_jobs_candidates.json"
    if not src.exists():
        src = base / "discovered_jobs.json"
    if not src.exists():
        raise FileNotFoundError("Missing candidates for %s" % region)

    jobs = json.loads(src.read_text(encoding="utf-8"))
    # Prefer candidates file; if discovered_jobs already validated-only, use candidates
    cand = base / "discovered_jobs_candidates.json"
    if cand.exists():
        jobs = json.loads(cand.read_text(encoding="utf-8"))

    print("=== VALIDATE %s -- %d candidates ===" % (region.upper(), len(jobs)))
    validated: list[dict] = []

    if not jobs:
        print("  no candidates; writing empty validated outputs")
    else:
        with ThreadPoolExecutor(max_workers=workers) as ex:
            futs = {ex.submit(validate_job, j): i for i, j in enumerate(jobs)}
            done = 0
            for fut in as_completed(futs):
                validated.append(fut.result())
                done += 1
                if done % 10 == 0 or done == len(jobs):
                    print("  scored %d/%d" % (done, len(jobs)))

    order = {j.get("JobPostUrl"): i for i, j in enumerate(jobs)}
    validated.sort(key=lambda j: order.get(j.get("JobPostUrl"), 10**9))

    kept_raw = [j for j in validated if j.get("_validate", {}).get("keep")]
    rejected = [j for j in validated if not j.get("_validate", {}).get("keep")]

    kept_clean = [{k: v for k, v in j.items() if not k.startswith("_")} for j in kept_raw]
    before_polish = len(kept_clean)
    kept_clean = polish_kept(kept_clean)
    if len(kept_clean) != before_polish:
        print(
            "  polish: deduped %d -> %d kept rows"
            % (before_polish, len(kept_clean))
        )

    audit = {
        "date": TODAY,
        "region": region,
        "candidates": len(jobs),
        "kept_before_polish": before_polish,
        "kept": len(kept_clean),
        "rejected": len(rejected),
        "academic": sum(1 for j in kept_clean if j.get("JobType") == "Academic"),
        "postdoc": sum(1 for j in kept_clean if j.get("JobType") == "Postdoc"),
        "rows": [
            {
                "title": j.get("JobPostTitle"),
                "url": j.get("JobPostUrl"),
                "keep": j["_validate"]["keep"],
                "label": j["_validate"]["label"],
                "reasons": j["_validate"]["reasons"],
                "job_type": j.get("JobType"),
            }
            for j in validated
        ],
    }
    (base / "validation_audit.json").write_text(
        json.dumps(audit, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (base / "discovered_jobs_validated.json").write_text(
        json.dumps(kept_clean, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    # Final listing = validated real postings only
    (base / "discovered_jobs.json").write_text(
        json.dumps(kept_clean, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    rows = [job_to_row(j) for j in kept_clean]
    xlsx_region = CLEAN_ECON / ("%s_listings_validated_%s.xlsx" % (region, TODAY))
    write_xlsx(rows, base / ("listings_validated_%s.xlsx" % TODAY), region.title())
    CLEAN_ECON.mkdir(parents=True, exist_ok=True)
    CLEAN.mkdir(parents=True, exist_ok=True)
    write_xlsx(rows, xlsx_region, region.title())

    # Always refresh PDF next to other clean/*.pdf (Cursor preview may need reopen)
    pdf_path = CLEAN / ("%s_listings_validated_%s.pdf" % (region, TODAY))
    export_one(
        xlsx_region,
        pdf_path,
        "%s Economics - Faculty + Postdoc (validated)" % region.title(),
        use_cjk=True,
    )
    print("  PDF -> %s (%d rows)" % (pdf_path, len(rows)))

    summary = {
        "candidates": len(jobs),
        "kept": len(kept_clean),
        "rejected": len(rejected),
        "academic": audit["academic"],
        "postdoc": audit["postdoc"],
        "pdf": str(pdf_path),
        "xlsx": str(xlsx_region),
    }
    print(
        "-> %s: kept=%d/%d (Academic=%d, Postdoc=%d)"
        % (
            region,
            summary["kept"],
            summary["candidates"],
            summary["academic"],
            summary["postdoc"],
        )
    )
    return summary


def main(regions: list[str] | None = None, workers: int = 6) -> None:
    ensure_dirs()
    targets = regions or list(REGIONS)
    all_summary = {"date": TODAY, "regions": {}}
    all_kept = []
    for region in targets:
        all_summary["regions"][region] = validate_region(region, workers=workers)
        path = DATA_ECON / region / "discovered_jobs_validated.json"
        if path.exists():
            all_kept.extend(json.loads(path.read_text(encoding="utf-8")))

    if all_kept:
        rows = [job_to_row(j) for j in all_kept]
        xlsx_all = CLEAN_ECON / ("china_coastal_listings_validated_%s.xlsx" % TODAY)
        write_xlsx(rows, xlsx_all, "China Coastal Validated")
        pdf_all = CLEAN / ("china_coastal_listings_validated_%s.pdf" % TODAY)
        export_one(
            xlsx_all,
            pdf_all,
            "China Coastal - Faculty + Postdoc (validated)",
            use_cjk=True,
        )
        print("  PDF -> %s (%d rows)" % (pdf_all, len(rows)))
    (DATA_ECON / ("china_coastal_validation_summary_%s.json" % TODAY)).write_text(
        json.dumps(all_summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(all_summary, ensure_ascii=False, indent=2))
    print("Done.")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--region", action="append", choices=list(REGIONS))
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--workers", type=int, default=6)
    args = ap.parse_args()
    regions = list(REGIONS) if args.all else args.region
    if not regions:
        raise SystemExit("Pass --region guangdong or --all")
    main(regions, workers=args.workers)
