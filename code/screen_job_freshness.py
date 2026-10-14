# -*- coding: utf-8 -*-
"""Screen validated jobs for freshness.

Keep only if:
  - title/url/body clearly indicates 2026+ cycle, OR
  - 2025 cycle BUT rolling / open-until-filled / long-term language is present.

Usage:
  python screen_job_freshness.py --region shanghai
  python screen_job_freshness.py --all
"""
from __future__ import annotations

import argparse
import json
import re
from datetime import date
from pathlib import Path

from china_universities_data import REGIONS
from download_china_provinces import fetch
from paths import DATA_ECON, ensure_dirs

TODAY = date.today().isoformat()

YEAR_2026_PLUS = re.compile(r"202[6-9]|20[3-9]\d")
YEAR_2025 = re.compile(r"2025")
YEAR_OLD = re.compile(r"201[0-9]|202[0-4]")
ROLLING = re.compile(
    "("
    "\u957f\u671f\u6709\u6548|\u5e38\u5e74\u62db\u8058|\u5e38\u5e74\u63a5\u6536|"
    "\u62db\u6ee1\u5373\u6b62|\u6eda\u52a8\u62db\u8058|\u968f\u65f6\u7533\u8bf7|"
    "\u5168\u5e74\u5f00\u653e|\u957f\u671f\u62db\u8058|until\\s+filled|open\\s+until|"
    "rolling|year[- ]?round"
    ")",
    re.I,
)


def strip_html(html: str) -> str:
    text = re.sub(r"(?is)<script[^>]*>.*?</script>", " ", html or "")
    text = re.sub(r"(?is)<style[^>]*>.*?</style>", " ", text)
    text = re.sub(r"<[^>]+>", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def freshness_decision(title: str, url: str, body: str) -> dict:
    blob = " ".join([title or "", url or "", (body or "")[:4000]])
    has_2026 = bool(YEAR_2026_PLUS.search(blob))
    has_2025 = bool(YEAR_2025.search(blob))
    has_old = bool(YEAR_OLD.search(title or "") or YEAR_OLD.search(url or ""))
    rolling = bool(ROLLING.search(blob))

    if has_2026:
        return {
            "keep": True,
            "label": "fresh_2026_plus",
            "reasons": ["has_2026_plus"],
            "rolling": rolling,
        }
    if has_2025 and rolling:
        return {
            "keep": True,
            "label": "fresh_2025_rolling",
            "reasons": ["year_2025_with_rolling"],
            "rolling": True,
        }
    if has_2025 and not has_old:
        # 2025 cycle without explicit rolling: keep only if no hard closed signal
        if re.search(r"\u622a\u6b62|\u5df2\u7ed3\u675f|closed|deadline", blob, re.I):
            return {
                "keep": False,
                "label": "stale_2025_closed",
                "reasons": ["2025_with_deadline_no_rolling"],
                "rolling": False,
            }
        return {
            "keep": True,
            "label": "fresh_2025_open",
            "reasons": ["year_2025_no_closed_signal"],
            "rolling": False,
        }
    if rolling and not has_old:
        return {
            "keep": True,
            "label": "rolling_undated",
            "reasons": ["rolling_language"],
            "rolling": True,
        }
    return {
        "keep": False,
        "label": "stale_pre_2025",
        "reasons": ["no_2025_or_2026_or_rolling"],
        "rolling": rolling,
    }


def screen_region(region: str, *, refetch: bool = True) -> dict:
    ensure_dirs()
    base = DATA_ECON / region
    src = base / "discovered_jobs_validated.json"
    if not src.exists():
        src = base / "discovered_jobs.json"
    if not src.exists():
        raise FileNotFoundError("No jobs to screen for %s" % region)

    jobs = json.loads(src.read_text(encoding="utf-8"))
    print("=== SCREEN FRESHNESS %s -- %d jobs ===" % (region.upper(), len(jobs)))
    kept, rejected = [], []
    for i, j in enumerate(jobs, 1):
        url = j.get("JobPostUrl") or j.get("ApplicationLink") or ""
        title = j.get("JobPostTitle") or ""
        body = ""
        if refetch and url:
            _st, html = fetch(url)
            body = strip_html(html or "")
        decision = freshness_decision(title, url, body)
        row = dict(j)
        row["_freshness"] = decision
        if decision["keep"]:
            kept.append(row)
        else:
            rejected.append(row)
        if i % 10 == 0 or i == len(jobs):
            print("  screened %d/%d (kept=%d)" % (i, len(jobs), len(kept)))

    kept_clean = [{k: v for k, v in j.items() if not k.startswith("_")} for j in kept]
    audit = {
        "date": TODAY,
        "region": region,
        "input": len(jobs),
        "kept": len(kept_clean),
        "rejected": len(rejected),
        "rows": [
            {
                "title": j.get("JobPostTitle"),
                "url": j.get("JobPostUrl"),
                "keep": j["_freshness"]["keep"],
                "label": j["_freshness"]["label"],
                "reasons": j["_freshness"]["reasons"],
            }
            for j in kept + rejected
        ],
    }
    (base / "freshness_audit.json").write_text(
        json.dumps(audit, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (base / "discovered_jobs_fresh.json").write_text(
        json.dumps(kept_clean, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print("-> %s freshness kept=%d/%d" % (region, len(kept_clean), len(jobs)))
    return {"input": len(jobs), "kept": len(kept_clean), "rejected": len(rejected)}


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--region", action="append", choices=list(REGIONS))
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--no-refetch", action="store_true")
    args = ap.parse_args()
    targets = list(REGIONS) if args.all else (args.region or [])
    if not targets:
        raise SystemExit("Pass --region or --all")
    summary = {"date": TODAY, "regions": {}}
    for region in targets:
        summary["regions"][region] = screen_region(
            region, refetch=not args.no_refetch
        )
    out = DATA_ECON / ("freshness_summary_%s.json" % TODAY)
    out.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
