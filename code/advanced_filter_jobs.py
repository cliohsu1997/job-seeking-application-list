# -*- coding: utf-8 -*-
"""Advanced one-by-one filter on the province full list.

Reads discovered_jobs_full.json (or validated), re-fetches each posting, and drops if:
  - not a real faculty/postdoc econ-related announcement (content check)
  - hard deadline already passed (and not rolling)
  - published / titled as 2025 without rolling / open-until-filled language
  - other stale signals (pre-2025 cycles)

Usage:
  python advanced_filter_jobs.py --region shanghai
  python advanced_filter_jobs.py --all
"""
from __future__ import annotations

import argparse
import json
import re
from datetime import date

from china_universities_data import REGIONS
from download_china_provinces import fetch, job_to_row, write_xlsx
from export_clean_pdf import export_one
from paths import CLEAN, CLEAN_ECON, DATA_ECON, ensure_dirs
from screen_job_freshness import ROLLING, YEAR_2025, YEAR_2026_PLUS, strip_html
from validate_job_links import DEADLINE_RE, ROLLING_OK, score_announcement

TODAY = date.today().isoformat()


def deadline_passed(body: str, title: str, url: str) -> tuple[bool, str | None]:
    blob = " ".join([title or "", url or "", (body or "")[:5000]])
    rolling = bool(ROLLING.search(blob) or ROLLING_OK.search(blob))
    today = date.today()
    for m in DEADLINE_RE.finditer(blob):
        try:
            y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3) or 1)
            if date(y, mo, d) < today and not rolling:
                return True, "deadline_passed_%s-%s-%s" % (y, mo, d)
        except ValueError:
            pass
    return False, None


def year_2025_without_rolling(title: str, url: str, body: str) -> bool:
    blob = " ".join([title or "", url or "", (body or "")[:4000]])
    if YEAR_2026_PLUS.search(blob):
        return False
    if not YEAR_2025.search(blob):
        return False
    return not bool(ROLLING.search(blob) or ROLLING_OK.search(blob))


def advanced_filter_region(region: str) -> dict:
    ensure_dirs()
    base = DATA_ECON / region
    src = base / "discovered_jobs_full.json"
    if not src.exists():
        src = base / "discovered_jobs_validated.json"
    if not src.exists():
        raise FileNotFoundError("No full list for %s (run preliminary validate first)" % region)

    jobs = json.loads(src.read_text(encoding="utf-8"))
    print("=== ADVANCED FILTER %s -- %d full-list jobs ===" % (region.upper(), len(jobs)))
    kept, rejected = [], []

    for i, j in enumerate(jobs, 1):
        url = j.get("JobPostUrl") or j.get("ApplicationLink") or ""
        title = j.get("JobPostTitle") or ""
        unit = j.get("School") or ""
        status, html = fetch(url) if url else (None, "")
        body = strip_html(html or "")

        reasons: list[str] = []
        keep = True
        label = "ok"

        content = score_announcement(
            html or "", url, title, unit_zh=unit, status=status
        )
        if not content.get("keep"):
            keep = False
            label = content.get("label") or "bad_content"
            reasons.extend(content.get("reasons") or ["content_fail"])

        expired, why = deadline_passed(body, title, url)
        if expired:
            keep = False
            label = "deadline_passed"
            reasons.append(why or "deadline_passed")

        if year_2025_without_rolling(title, url, body):
            keep = False
            label = "stale_2025_no_rolling"
            reasons.append("published_2025_without_rolling")

        # Require 2026+ or rolling if no clear 2026 signal
        blob = " ".join([title, url, body[:4000]])
        if keep and not YEAR_2026_PLUS.search(blob) and not (
            YEAR_2025.search(blob) and (ROLLING.search(blob) or ROLLING_OK.search(blob))
        ):
            if not (ROLLING.search(blob) or ROLLING_OK.search(blob)):
                keep = False
                label = "not_2026_or_rolling"
                reasons.append("needs_2026_or_rolling_2025")

        row = dict(j)
        row["_advanced"] = {
            "keep": keep,
            "label": label,
            "reasons": reasons or ["ok"],
            "content_label": content.get("label"),
        }
        if keep:
            kept.append(row)
            mark = "KEEP"
        else:
            rejected.append(row)
            mark = "DROP"
        print(
            "  [%d/%d] %s %s | %s"
            % (i, len(jobs), mark, label, (title or "")[:52])
        )

    kept_clean = [{k: v for k, v in j.items() if not k.startswith("_")} for j in kept]
    audit = {
        "date": TODAY,
        "region": region,
        "full_list": len(jobs),
        "kept": len(kept_clean),
        "rejected": len(rejected),
        "rows": [
            {
                "title": j.get("JobPostTitle"),
                "url": j.get("JobPostUrl"),
                "keep": j["_advanced"]["keep"],
                "label": j["_advanced"]["label"],
                "reasons": j["_advanced"]["reasons"],
                "job_type": j.get("JobType"),
            }
            for j in kept + rejected
        ],
    }
    (base / "advanced_filter_audit.json").write_text(
        json.dumps(audit, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (base / "discovered_jobs_verified.json").write_text(
        json.dumps(kept_clean, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (base / "discovered_jobs_fresh.json").write_text(
        json.dumps(kept_clean, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (base / "discovered_jobs_validated.json").write_text(
        json.dumps(kept_clean, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (base / "discovered_jobs.json").write_text(
        json.dumps(kept_clean, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    rows = [job_to_row(j) for j in kept_clean]
    CLEAN.mkdir(parents=True, exist_ok=True)
    CLEAN_ECON.mkdir(parents=True, exist_ok=True)
    xlsx = CLEAN_ECON / ("%s_listings_verified_%s.xlsx" % (region, TODAY))
    pdf = CLEAN / ("%s_listings_verified_%s.pdf" % (region, TODAY))
    write_xlsx(rows, base / ("listings_verified_%s.xlsx" % TODAY), region.title())
    write_xlsx(rows, xlsx, region.title())
    export_one(
        xlsx,
        pdf,
        "%s Economics - Faculty + Postdoc (advanced filter)" % region.title(),
        use_cjk=True,
    )
    export_one(
        xlsx,
        CLEAN / ("%s_listings_validated_%s.pdf" % (region, TODAY)),
        "%s Economics - Faculty + Postdoc (validated)" % region.title(),
        use_cjk=True,
    )
    print(
        "-> %s advanced kept=%d/%d PDF -> %s"
        % (region, len(kept_clean), len(jobs), pdf)
    )
    return {
        "full_list": len(jobs),
        "kept": len(kept_clean),
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
        summary["regions"][region] = advanced_filter_region(region)
    out = DATA_ECON / ("advanced_filter_summary_%s.json" % TODAY)
    out.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
