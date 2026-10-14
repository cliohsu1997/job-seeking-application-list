"""
Read JOE XML chunk-by-chunk and write Application-list format.

Columns match the Application sheet exactly (no extra columns).
Position title is stored inside Job Post as: "<title> | <joe url>"
Job Type: Faculty | Postdoc | Non-Academic (see classify_aea_job_type).
"""
from __future__ import annotations

import argparse
import html
import re
from datetime import datetime
from pathlib import Path
import xml.etree.ElementTree as ET

from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from classify_aea_job_type import (
    classify_with_optional_fetch,
    load_page_cache,
    save_page_cache,
)
from joe_urls import listing_url, read_year_issue

from paths import APP_XLSX, DATA_ECON, ensure_dirs

DATA = DATA_ECON

# Must match Application sheet headers exactly (order included)
COLUMNS = [
    "Country",
    "Location",
    "University",
    "School",
    "Deadline",
    "Job Post",
    "Required Material",
    "Submission Channel",
    "Application Link",
    "Email",
    "Submission Date",
    "Reference Status",
    "Interview Date",
    "Job Type",
]

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
    return re.sub(r"\s+", " ", value).strip()


def text(el: ET.Element | None) -> str:
    if el is None:
        return ""
    return clean_text("".join(el.itertext()))


def local_tag(tag: str) -> str:
    return tag.split("}")[-1].lower() if tag else ""


def parse_deadline(raw: str):
    raw = (raw or "").strip()
    if not raw:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%m/%d/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            continue
    return raw


def parse_locations(pos: ET.Element) -> tuple[str, str]:
    countries, places = [], []
    locs = None
    for child in pos:
        if local_tag(child.tag) == "locations":
            locs = child
            break
    if locs is None:
        return "", ""
    for loc in locs:
        if local_tag(loc.tag) != "location":
            continue
        fields = {local_tag(c.tag): text(c) for c in loc}
        country = fields.get("country", "")
        state = fields.get("state", "")
        city = fields.get("city", "")
        if country and country not in countries:
            countries.append(country)
        place = ", ".join(p for p in (city.title() if city else "", state) if p)
        if place and place not in places:
            places.append(place)
    return "; ".join(countries), "; ".join(places)


def extract_required_materials(full_text: str) -> str:
    if not full_text:
        return ""
    found = []
    m = re.search(
        r"Application Requirements?\s*:?\s*(.+?)(?:Application deadline|Application Instructions|To apply|About |$)",
        full_text,
        flags=re.I | re.S,
    )
    block = m.group(1) if m else full_text
    for label in MATERIAL_LABELS:
        if re.search(rf"\b{re.escape(label)}\b", block, flags=re.I):
            norm = "CV" if label.lower().startswith("curriculum vitae") or label == "CV" else label
            if norm == "Letter of Reference":
                norm = "Letters of Reference"
            if norm == "Research Paper":
                norm = "Research Papers"
            if norm not in found:
                found.append(norm)
    return "; ".join(found)


def extract_email(full_text: str) -> str:
    if not full_text:
        return ""
    emails = re.findall(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", full_text)
    out = []
    for e in emails:
        e = e.rstrip(".")
        if e.lower() not in {x.lower() for x in out}:
            out.append(e)
    return "; ".join(out[:3])


def child_text(pos: ET.Element, name: str) -> str:
    for child in pos:
        if local_tag(child.tag) == name:
            return text(child)
    return ""


def position_to_row(
    pos: ET.Element,
    year: str,
    issue: str,
    *,
    fetch_ambiguous: bool = False,
    page_cache: dict | None = None,
) -> dict:
    jp_id = pos.attrib.get("jp_id", "")
    section = child_text(pos, "jp_section")
    full_text = child_text(pos, "jp_full_text")
    country, location = parse_locations(pos)
    department = child_text(pos, "jp_department")
    division = child_text(pos, "jp_division")
    school = department or division
    title = child_text(pos, "jp_title")
    # Canonical JOE_ID (e.g. 2026-02_111477855) — bare numeric ids often break in Excel
    joe_url = listing_url(jp_id, year, issue) if jp_id else ""

    job_type, reason, was_ambiguous = classify_with_optional_fetch(
        section=section,
        title=title,
        full_text=full_text,
        url=joe_url,
        fetch_ambiguous=fetch_ambiguous,
        cache=page_cache,
    )

    return {
        "Country": country or None,
        "Location": location or None,
        "University": child_text(pos, "jp_institution") or None,
        "School": school or None,
        "Deadline": parse_deadline(child_text(pos, "jp_application_deadline")),
        "Job Post": joe_url or None,
        "Required Material": extract_required_materials(full_text) or None,
        "Submission Channel": "AEA",
        "Application Link": joe_url or None,  # replaced by verify_aea_links.py
        "Email": extract_email(full_text) or None,
        "Submission Date": None,
        "Reference Status": None,
        "Interview Date": None,
        "Job Type": job_type,
        "_title": title,
        "_section": section,
        "_job_type_reason": reason,
        "_ambiguous": was_ambiguous,
    }


def iter_positions_chunked(xml_path: Path):
    """Yield each <position> using iterparse; clear nodes to keep memory low."""
    for _event, elem in ET.iterparse(xml_path, events=("end",)):
        if local_tag(elem.tag) != "position":
            continue
        yield elem
        elem.clear()


def load_joe_rows(
    xml_path: Path | None = None,
    *,
    fetch_ambiguous: bool = False,
) -> list[dict]:
    if xml_path is None:
        xml_files = sorted(DATA.glob("joe_listings_*.xml"), reverse=True)
        if not xml_files:
            raise FileNotFoundError("No joe_listings_*.xml in data/. Run download_joe.py first.")
        xml_path = xml_files[0]

    year, issue = read_year_issue(xml_path)
    print(f"Parsing chunk-by-chunk: {xml_path} (JOE cycle {year}-{issue})")
    page_cache = load_page_cache() if fetch_ambiguous else None
    cache_size_before = len(page_cache) if page_cache is not None else 0
    rows: list[dict] = []
    for i, pos in enumerate(iter_positions_chunked(xml_path), 1):
        rows.append(
            position_to_row(
                pos,
                year,
                issue,
                fetch_ambiguous=fetch_ambiguous,
                page_cache=page_cache,
            )
        )
        if i % 50 == 0:
            print(f"  ... {i} positions")
    if page_cache is not None and len(page_cache) != cache_size_before:
        save_page_cache(page_cache)
        print(f"  page cache saved ({len(page_cache)} urls)")
    print(f"  parsed {len(rows)} positions")
    return rows


def read_application_columns() -> list[str]:
    wb = load_workbook(APP_XLSX, read_only=True)
    ws = wb["Application"]
    headers = []
    for c in range(1, ws.max_column + 1):
        headers.append(ws.cell(1, c).value)
    wb.close()
    return headers


def update_application_workbook(rows: list[dict]) -> None:
    wb = load_workbook(APP_XLSX)
    ws = wb["Application"]
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]

    if headers != COLUMNS:
        # keep Application as source of truth if it already has these cols
        if "Job Type" not in headers:
            col = ws.max_column + 1
            ws.cell(1, col, "Job Type")
            headers.append("Job Type")
        for r in range(2, ws.max_row + 1):
            if ws.cell(r, 3).value and not ws.cell(r, headers.index("Job Type") + 1).value:
                ws.cell(r, headers.index("Job Type") + 1, "Faculty")

    if "JOE Listings" in wb.sheetnames:
        del wb["JOE Listings"]
    js = wb.create_sheet("JOE Listings")
    header_font = Font(bold=True)
    for c, name in enumerate(COLUMNS, 1):
        cell = js.cell(1, c, name)
        cell.font = header_font
    for r_i, row in enumerate(rows, 2):
        for c, name in enumerate(COLUMNS, 1):
            js.cell(r_i, c, row.get(name))
    for c in range(1, len(COLUMNS) + 1):
        js.column_dimensions[get_column_letter(c)].width = 16

    wb.save(APP_XLSX)


def _job_type_rank(job_type: str) -> int:
    if job_type in ("Faculty", "Academic"):
        return 0
    if job_type == "Postdoc":
        return 1
    return 2


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description="Compile JOE XML into Application list")
    parser.add_argument(
        "--fetch-ambiguous",
        action="store_true",
        help="Fetch JOE listing pages for ambiguous academic rows (rate-limited, cached)",
    )
    args = parser.parse_args(argv)

    ensure_dirs()
    app_cols = read_application_columns()
    if app_cols != COLUMNS:
        print("WARNING: Application headers differ from script COLUMNS")
        print("  Application:", app_cols)
        print("  Script:     ", COLUMNS)

    rows = load_joe_rows(fetch_ambiguous=args.fetch_ambiguous)

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
    ambiguous = [
        r
        for r in rows
        if r.get("_ambiguous")
        or (r.get("_job_type_reason") or "").startswith("ambiguous")
    ]
    print(f"Compiled {len(rows)} JOE listings")
    print(f"Columns ({len(COLUMNS)}): {COLUMNS}")
    print(f"Faculty: {n_fac} | Postdoc: {n_post} | Non-Academic: {n_non}")
    if ambiguous:
        print(f"Ambiguous (defaulted Faculty): {len(ambiguous)}")
        for r in ambiguous[:25]:
            print(
                f"  - {r.get('University')}: {r.get('_title')} "
                f"[{r.get('_section', '')[:40]}]"
            )
        if len(ambiguous) > 25:
            print(f"  ... and {len(ambiguous) - 25} more")
    print(f"Updated sheet JOE Listings in {APP_XLSX}")


if __name__ == "__main__":
    main()
