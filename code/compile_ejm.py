"""
Compile EJM JSON ads into Application-list columns (chunk-by-chunk).

Job Type: Faculty | Postdoc | Non-Academic (from position_types + title).
"""
from __future__ import annotations

import html
import json
import re
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font

from classify_aea_job_type import classify_ejm_position_types
from columns import COLUMNS

from paths import APP_XLSX, DATA_ECON, ensure_dirs

DATA = DATA_ECON

MATERIAL_LABELS = [
    "Curriculum Vitae (CV)",
    "CV",
    "Cover Letter",
    "Job Market Paper",
    "Research Papers",
    "Research Paper",
    "Teaching Evaluations",
    "Teaching Statement",
    "Research Statement",
    "Letters of Reference",
    "Letter of Reference",
    "Transcripts",
    "Transcript",
    "Writing Sample",
]


def clean_text(value: str) -> str:
    value = html.unescape(html.unescape(value or ""))
    value = re.sub(r"<[^>]+>", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def parse_deadline(raw):
    if raw is None or raw == "":
        return None
    if isinstance(raw, (int, float)):
        return None
    s = str(raw).strip()
    for fmt in ("%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d"):
        try:
            return datetime.strptime(s[:19] if " " in s else s[:10], fmt if " " in s or len(s) > 10 else "%Y-%m-%d")
        except ValueError:
            continue
    try:
        return datetime.strptime(s[:10], "%Y-%m-%d")
    except ValueError:
        return s


def extract_required_materials(full_text: str) -> str:
    if not full_text:
        return ""
    found = []
    for label in MATERIAL_LABELS:
        if re.search(rf"\b{re.escape(label)}\b", full_text, flags=re.I):
            norm = "CV" if label.lower().startswith("curriculum vitae") or label == "CV" else label
            if norm == "Letter of Reference":
                norm = "Letters of Reference"
            if norm == "Research Paper":
                norm = "Research Papers"
            if norm not in found:
                found.append(norm)
    return "; ".join(found)


def extract_email(full_text: str) -> str:
    emails = re.findall(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", full_text or "")
    out = []
    for e in emails:
        e = e.rstrip(".")
        if e.lower() not in {x.lower() for x in out}:
            out.append(e)
    return "; ".join(out[:3])


def location_fields(ad: dict) -> tuple[str, str]:
    locs = ad.get("locations") or []
    countries, places = [], []
    for loc in locs:
        if not isinstance(loc, dict):
            continue
        cc = (loc.get("country_code") or "").strip()
        city = (loc.get("city") or "").strip()
        state = (loc.get("state") or "").strip()
        if cc and cc not in countries:
            countries.append(cc)
        place = ", ".join(p for p in (city, state) if p)
        if place and place not in places:
            places.append(place)
    return "; ".join(countries), "; ".join(places)


def ad_to_row(ad: dict) -> dict:
    title = clean_text(ad.get("adtitle") or "")
    url = (ad.get("url") or "").strip()
    plain = clean_text(ad.get("adtext") or "")
    country, location = location_fields(ad)
    job_post = " | ".join(p for p in (title, url) if p)

    job_type, _reason = classify_ejm_position_types(
        ad.get("position_types"),
        title=title,
        body=plain,
    )

    return {
        "Country": country or None,
        "Location": location or None,
        "University": clean_text(ad.get("name") or "") or None,
        "School": clean_text(ad.get("department") or "") or None,
        "Deadline": parse_deadline(ad.get("deadline_date") or ad.get("enddate")),
        "Job Post": job_post or None,
        "Required Material": extract_required_materials(plain) or None,
        "Submission Channel": "EJM",
        "Application Link": url or None,  # direct EJM listing URL
        "Email": extract_email(plain) or None,
        "Submission Date": None,
        "Reference Status": None,
        "Interview Date": None,
        "Job Type": job_type,
    }


def iter_ads_chunked(json_path: Path, chunk_size: int = 50):
    """Load JSON once, yield ads in chunks (memory-friendly iteration API)."""
    with json_path.open("r", encoding="utf-8") as f:
        ads = json.load(f)
    if not isinstance(ads, list):
        raise ValueError("EJM JSON root is not a list")
    for i in range(0, len(ads), chunk_size):
        yield ads[i : i + chunk_size]


def load_ejm_rows(json_path: Path | None = None) -> list[dict]:
    if json_path is None:
        files = sorted(DATA.glob("ejm_listings_*.json"), reverse=True)
        # prefer dated full downloads over probe file
        files = [p for p in files if "probe" not in p.name] or files
        if not files:
            raise FileNotFoundError("No ejm_listings_*.json in data/. Run download_ejm.py first.")
        json_path = files[0]

    print(f"Parsing chunk-by-chunk: {json_path}")
    rows: list[dict] = []
    n = 0
    for chunk in iter_ads_chunked(json_path):
        for ad in chunk:
            rows.append(ad_to_row(ad))
            n += 1
        print(f"  ... {n} ads")
    return rows


def update_application_workbook(rows: list[dict]) -> None:
    wb = load_workbook(APP_XLSX)
    if "EJM Listings" in wb.sheetnames:
        del wb["EJM Listings"]
    ws = wb.create_sheet("EJM Listings")
    header_font = Font(bold=True)
    for c, name in enumerate(COLUMNS, 1):
        cell = ws.cell(1, c, name)
        cell.font = header_font
    for r_i, row in enumerate(rows, 2):
        for c, name in enumerate(COLUMNS, 1):
            ws.cell(r_i, c, row.get(name))
    wb.save(APP_XLSX)


def _job_type_rank(job_type: str) -> int:
    if job_type in ("Faculty", "Academic"):
        return 0
    if job_type == "Postdoc":
        return 1
    return 2


def main() -> None:
    rows = load_ejm_rows()

    def sort_key(r):
        d = r.get("Deadline")
        dkey = d if isinstance(d, datetime) else datetime.max
        return (
            _job_type_rank(r.get("Job Type") or ""),
            dkey,
            r.get("University") or "",
        )

    rows.sort(key=sort_key)
    update_application_workbook(rows)
    n_fac = sum(1 for r in rows if r["Job Type"] in ("Faculty", "Academic"))
    n_post = sum(1 for r in rows if r["Job Type"] == "Postdoc")
    n_non = sum(1 for r in rows if r["Job Type"] == "Non-Academic")
    print(f"Compiled {len(rows)} EJM listings")
    print(f"Columns ({len(COLUMNS)}): {COLUMNS}")
    print(f"Faculty: {n_fac} | Postdoc: {n_post} | Non-Academic: {n_non}")
    print(f"Updated sheet EJM Listings in {APP_XLSX}")


if __name__ == "__main__":
    main()
