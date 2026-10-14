# -*- coding: utf-8 -*-
"""Download every AEA / EJM / CICE listing page; set Job Type + Deadline from page.

Uses cached page text under data/economics/aea_job_type_page_cache.json.

Usage:
  python enrich_website_from_pages.py
  python enrich_website_from_pages.py --asia-pacific-only
  python enrich_website_from_pages.py --channels AEA EJM CICE
"""
from __future__ import annotations

import argparse
import json
import re
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font

from classify_aea_job_type import (
    classify_with_optional_fetch,
    fetch_listing_text,
    load_page_cache,
    save_page_cache,
)
from columns import COLUMNS
from combine_clean import extract_url, normalize_row
from deadline_extract import extract_deadline, has_deadline
from paths import APP_XLSX, DATA_ECON, ensure_dirs

TODAY = date.today().isoformat()
JOE_URL_RE = re.compile(
    r"https?://www\.aeaweb\.org/joe/listing\.php\?JOE_ID=[^\s|&]+",
    re.I,
)
ASIA_COUNTRIES = {
    "CN",
    "CHINA",
    "HK",
    "HONG KONG",
    "MO",
    "MACAU",
    "MACAO",
    "MACAO; CHINA",
    "TW",
    "TAIWAN",
    "MY",
    "MALAYSIA",
    "JP",
    "JAPAN",
    "AU",
    "AUSTRALIA",
    "NZ",
    "NEW ZEALAND",
}

SOURCE_SHEETS = [
    ("JOE Listings", "AEA"),
    ("EJM Listings", "EJM"),
    ("CICE Listings", "CICE"),
]


def _norm_country(val) -> str:
    return re.sub(r"\s+", " ", str(val or "").strip()).upper()


def asia_pacific_keep(row: dict) -> bool:
    return _norm_country(row.get("Country")) in ASIA_COUNTRIES


def _title_from_row(row: dict) -> str:
    jp = str(row.get("Job Post") or "")
    if " | " in jp:
        left = jp.split(" | ", 1)[0].strip()
        if left and not left.startswith("http"):
            return left
    return left if " | " in jp else (jp if not jp.startswith("http") else "")


def listing_urls_for_row(row: dict) -> list[str]:
    """Primary listing page first (JOE listing / EJM ad / CICE detail)."""
    channel = (row.get("Submission Channel") or "").upper()
    jp = str(row.get("Job Post") or "")
    app = extract_url(row.get("Application Link")) or ""
    urls: list[str] = []

    if channel == "AEA":
        m = JOE_URL_RE.search(jp) or JOE_URL_RE.search(app or "")
        if m:
            urls.append(m.group(0).rstrip(").,;\"'"))
        if app and app not in urls and app.startswith("http"):
            urls.append(app)
    else:
        if app and app.startswith("http"):
            urls.append(app)
        for u in re.findall(r"https?://[^\s|]+", jp):
            u = u.rstrip(").,;\"'")
            if u not in urls:
                urls.append(u)
    return urls


def _deadline_str(val) -> str | None:
    if val in (None, "", "None"):
        return None
    if isinstance(val, datetime):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, date):
        return val.isoformat()
    s = str(val).strip()
    if re.match(r"20\d{2}-\d{2}-\d{2}", s):
        return s[:10]
    return None


def enrich_row(
    row: dict,
    *,
    cache: dict,
    force_deadline: bool = False,
) -> tuple[dict, dict]:
    """Return (updated_row, stats_flags)."""
    flags = {"fetched": 0, "type_changed": 0, "deadline_filled": 0, "pages": 0}
    urls = listing_urls_for_row(row)
    if not urls:
        return row, flags

    title = _title_from_row(row)
    channel = (row.get("Submission Channel") or "").upper()
    section = ""
    pages: list[str] = []
    for url in urls:
        text = fetch_listing_text(url, cache=cache)
        flags["fetched"] += 1
        if text:
            pages.append(text)
            flags["pages"] += 1

    page_blob = "\n".join(pages)
    primary = urls[0]
    job_type, reason, _amb = classify_with_optional_fetch(
        section=section,
        title=title,
        full_text=page_blob,
        url=primary,
        fetch_always=False,  # already fetched above
        cache=cache,
    )
    # Re-classify with merged page text explicitly
    from classify_aea_job_type import _classify_text

    jt2, reason2 = _classify_text(title, page_blob, section)
    if jt2:
        job_type, reason = jt2, "page_" + reason2
    elif page_blob:
        job_type, reason = job_type, reason

    old_type = row.get("Job Type") or ""
    if old_type == "Academic":
        old_type = "Faculty"
    new_row = dict(row)
    new_row["Job Type"] = job_type
    new_row["_job_type_reason"] = reason
    if job_type != old_type:
        flags["type_changed"] = 1

    # Deadline from listing page(s)
    need_dl = force_deadline or not has_deadline(
        {"Deadline": new_row.get("Deadline")}
    )
    page_deadline = extract_deadline(title + "\n" + page_blob) if page_blob else None
    if page_deadline and (need_dl or force_deadline):
        if _deadline_str(new_row.get("Deadline")) != page_deadline:
            new_row["Deadline"] = page_deadline
            flags["deadline_filled"] = 1
    elif page_deadline and not _deadline_str(new_row.get("Deadline")):
        new_row["Deadline"] = page_deadline
        flags["deadline_filled"] = 1

    # CICE often has no structured deadline in JSON  page is the source of truth
    if channel == "CICE" and page_deadline:
        new_row["Deadline"] = page_deadline
        if flags["deadline_filled"] == 0 and _deadline_str(row.get("Deadline")) != page_deadline:
            flags["deadline_filled"] = 1

    return new_row, flags


def cell_value_or_link(cell) -> str | None:
    if cell.hyperlink and cell.hyperlink.target:
        return str(cell.hyperlink.target)
    if cell.value is None:
        return None
    return str(cell.value)


def read_sheet(wb, sheet_name: str, default_channel: str) -> list[dict]:
    if sheet_name not in wb.sheetnames:
        return []
    ws = wb[sheet_name]
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    rows = []
    for r in range(2, ws.max_row + 1):
        raw = {}
        for c, name in enumerate(headers, 1):
            if not name:
                continue
            cell = ws.cell(r, c)
            if name in ("Job Post", "Application Link"):
                raw[name] = cell_value_or_link(cell)
            else:
                raw[name] = cell.value
        if not any(raw.get(h) for h in ("University", "Job Post", "Application Link")):
            continue
        rows.append(normalize_row(raw, default_channel))
    return rows


def write_sheet(wb, sheet_name: str, rows: list[dict]) -> None:
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    link_font = Font(color="0563C1", underline="single")
    # Prefer shared COLUMNS; fall back to union of keys present
    headers = list(COLUMNS)
    for c, name in enumerate(headers, 1):
        ws.cell(1, c, name).font = Font(bold=True)
    for r_i, row in enumerate(rows, 2):
        for c, name in enumerate(headers, 1):
            val = row.get(name)
            if name == "Deadline":
                val = _deadline_str(val) or val
            cell = ws.cell(r_i, c, val)
            if name == "Job Post":
                title = _title_from_row(row) or "Listing"
                # Prefer JOE listing URL in hyperlink when present
                jp = str(row.get("Job Post") or "")
                m = JOE_URL_RE.search(jp)
                href = (
                    (m.group(0) if m else None)
                    or extract_url(jp)
                    or row.get("Application Link")
                )
                cell.value = title
                if href:
                    cell.hyperlink = str(href)
                    cell.font = link_font
            elif name == "Application Link" and val:
                cell.hyperlink = str(val)
                cell.font = link_font


def run(
    *,
    channels: list[str] | None = None,
    asia_pacific_only: bool = False,
    force_deadline: bool = False,
) -> dict:
    ensure_dirs()
    want = {c.upper() for c in (channels or ["AEA", "EJM", "CICE"])}
    cache = load_page_cache()
    cache_before = len(cache)

    wb = load_workbook(APP_XLSX)
    all_enriched: list[dict] = []
    summary = {
        "date": TODAY,
        "channels": {},
        "fetched": 0,
        "pages_ok": 0,
        "type_changed": 0,
        "deadline_filled": 0,
        "rows": 0,
    }

    for sheet_name, channel in SOURCE_SHEETS:
        if channel not in want:
            continue
        rows = read_sheet(wb, sheet_name, channel)
        targets = [
            i
            for i, r in enumerate(rows)
            if (not asia_pacific_only) or asia_pacific_keep(r)
        ]
        print(
            "=== ENRICH %s from listing pages (%d / %d rows) ==="
            % (channel, len(targets), len(rows))
        )
        ch_stats = {
            "rows": len(targets),
            "sheet_rows": len(rows),
            "fetched": 0,
            "pages_ok": 0,
            "type_changed": 0,
            "deadline_filled": 0,
            "faculty": 0,
            "postdoc": 0,
            "non_academic": 0,
            "with_deadline": 0,
        }
        done = 0
        for i in targets:
            row = rows[i]
            new_row, flags = enrich_row(
                row, cache=cache, force_deadline=force_deadline
            )
            rows[i] = new_row
            done += 1
            ch_stats["fetched"] += flags["fetched"]
            ch_stats["pages_ok"] += flags["pages"]
            ch_stats["type_changed"] += flags["type_changed"]
            ch_stats["deadline_filled"] += flags["deadline_filled"]
            if done % 25 == 0 or done == len(targets):
                print(
                    "  [%d/%d] type_changed=%d deadline_filled=%d cache=%d"
                    % (
                        done,
                        len(targets),
                        ch_stats["type_changed"],
                        ch_stats["deadline_filled"],
                        len(cache),
                    )
                )
            if done % 50 == 0:
                save_page_cache(cache)

        for new_row in rows:
            jt = new_row.get("Job Type") or ""
            if jt == "Academic":
                new_row["Job Type"] = "Faculty"
                jt = "Faculty"
            if jt == "Postdoc":
                ch_stats["postdoc"] += 1
            elif jt == "Non-Academic":
                ch_stats["non_academic"] += 1
            else:
                ch_stats["faculty"] += 1
            if has_deadline({"Deadline": new_row.get("Deadline")}):
                ch_stats["with_deadline"] += 1

        write_sheet(wb, sheet_name, rows)
        all_enriched.extend(rows)
        summary["channels"][channel] = ch_stats
        summary["fetched"] += ch_stats["fetched"]
        summary["pages_ok"] += ch_stats["pages_ok"]
        summary["type_changed"] += ch_stats["type_changed"]
        summary["deadline_filled"] += ch_stats["deadline_filled"]
        summary["rows"] += ch_stats["rows"]
        print(
            "  -> Faculty=%d Postdoc=%d Non-Academic=%d with_deadline=%d/%d"
            % (
                ch_stats["faculty"],
                ch_stats["postdoc"],
                ch_stats["non_academic"],
                ch_stats["with_deadline"],
                ch_stats["sheet_rows"],
            )
        )

    wb.save(APP_XLSX)
    save_page_cache(cache)
    summary["cache_entries"] = len(cache)
    summary["cache_new"] = len(cache) - cache_before

    # Rebuild clean/economics all_listings from workbook (full sheets, not AP-only)
    from combine_clean import main as combine_clean_main

    print("=== REBUILD clean/economics/all_listings ===")
    combine_clean_main()

    (DATA_ECON / ("website_page_enrich_summary_%s.json" % TODAY)).write_text(
        json.dumps(summary, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(
        "Done. rows=%d pages_ok=%d type_changed=%d deadline_filled=%d cache=%d"
        % (
            summary["rows"],
            summary["pages_ok"],
            summary["type_changed"],
            summary["deadline_filled"],
            summary["cache_entries"],
        )
    )
    return summary


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument(
        "--channels",
        nargs="+",
        default=["AEA", "EJM", "CICE"],
        help="Which sources to enrich",
    )
    ap.add_argument(
        "--asia-pacific-only",
        action="store_true",
        help="Only CN/MY/JP/AU/NZ/HK/TW/MO rows (faster)",
    )
    ap.add_argument(
        "--force-deadline",
        action="store_true",
        help="Overwrite existing deadlines when page has a date",
    )
    args = ap.parse_args()
    print(
        json.dumps(
            run(
                channels=args.channels,
                asia_pacific_only=args.asia_pacific_only,
                force_deadline=args.force_deadline,
            ),
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
