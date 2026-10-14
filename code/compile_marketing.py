"""
Compile Marketing Job Market spreadsheet into Application-list columns.

Primary source: data/marketing/Marketing Job Market.xlsx
  - sheet "2026 Job Postings" (tenure-track / open academic)
  - sheet "2026 NTTPDInd" (NTT / clinical / industry)
"""
from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from columns import COLUMNS
from paths import APP_XLSX, DATA_MKT, ensure_dirs

SOURCE_XLSX_NAME = "Marketing Job Market.xlsx"
CONTINENT_COUNTRY = {
    "NA": "North America",
    "AS": "Asia",
    "EU": "Europe",
    "OC": "Oceania",
    "AF": "Africa",
    "SA": "South America",
    "LA": "Latin America",
    "OTHER": "Other",
}


def parse_deadline(raw):
    if raw is None or raw == "":
        return None
    if isinstance(raw, datetime):
        return raw
    if type(raw) is date:
        return datetime(raw.year, raw.month, raw.day)
    s = str(raw).strip()
    if not s or s.lower() in {"until filled", "open", "n/a", "na", "-"}:
        return None
    for fmt, n in (("%Y-%m-%d", 10), ("%m/%d/%Y", 10), ("%m-%d-%Y", 10), ("%b %d, %Y", 20), ("%B %d, %Y", 20)):
        try:
            return datetime.strptime(s[:n], fmt)
        except ValueError:
            continue
    # embedded date
    m = re.search(r"(\d{4}-\d{2}-\d{2})", s)
    if m:
        return datetime.strptime(m.group(1), "%Y-%m-%d")
    return None


def extract_url(cell, fallback_text=None) -> str | None:
    if cell is not None and cell.hyperlink and cell.hyperlink.target:
        return str(cell.hyperlink.target).strip()
    text = fallback_text if fallback_text is not None else (cell.value if cell is not None else None)
    if text is None:
        return None
    s = str(text).strip()
    if s.startswith("http://") or s.startswith("https://"):
        return s.split()[0].rstrip(").,;\"'")
    urls = re.findall(r"https?://[^\s|]+", s)
    return urls[-1].rstrip(").,;\"'") if urls else None


def job_type_from_track(track: str | None, sheet: str) -> str:
    if sheet == "NTT":
        t = (track or "").lower()
        if any(x in t for x in ("industry", "ind", "private", "non-academic")):
            return "Non-Academic"
        return "Academic"
    return "Academic"


def title_from_link_cell(cell, notes) -> str:
    """Prefer display text if it is not a bare URL."""
    val = cell.value if cell is not None else None
    if val and not str(val).startswith("http"):
        return str(val).strip()
    if notes:
        first = str(notes).split("\n")[0].strip()
        if 8 <= len(first) <= 160:
            return first
    return "Marketing faculty position"


def read_job_postings(wb) -> list[dict]:
    ws = wb["2026 Job Postings"]
    rows = []
    for r in range(5, ws.max_row + 1):
        school = ws.cell(r, 2).value
        if not school or str(school).strip() in {"School", ""}:
            continue
        continent = (ws.cell(r, 3).value or "").strip().upper() if ws.cell(r, 3).value else ""
        link_cell = ws.cell(r, 5)
        url = extract_url(link_cell)
        deadline = parse_deadline(ws.cell(r, 6).value)
        notes = ws.cell(r, 7).value
        title = title_from_link_cell(link_cell, notes)
        materials = None
        if notes:
            # keep short note as required-material hint only if it looks like materials list
            note_s = str(notes).replace("\n", " ").strip()
            if any(k in note_s.lower() for k in ("cv", "cover letter", "letters", "packet", "materials")):
                materials = note_s[:300]
        rows.append(
            {
                "Country": CONTINENT_COUNTRY.get(continent, continent or None),
                "Location": None,
                "University": str(school).strip(),
                "School": None,
                "Deadline": deadline,
                "Job Post": f"{title} | {url}" if url else title,
                "Required Material": materials,
                "Submission Channel": "MJM",
                "Application Link": url,
                "Email": None,
                "Submission Date": None,
                "Reference Status": None,
                "Interview Date": None,
                "Job Type": "Academic",
                "_title": title,
            }
        )
    return rows


def read_ntt(wb) -> list[dict]:
    if "2026 NTTPDInd" not in wb.sheetnames:
        return []
    ws = wb["2026 NTTPDInd"]
    rows = []
    for r in range(3, ws.max_row + 1):
        school = ws.cell(r, 2).value
        if not school or str(school).strip() in {"School", ""}:
            continue
        continent = (ws.cell(r, 3).value or "").strip().upper() if ws.cell(r, 3).value else ""
        track = ws.cell(r, 4).value
        link_cell = ws.cell(r, 5)
        url = extract_url(link_cell)
        deadline = parse_deadline(ws.cell(r, 7).value)
        notes = ws.cell(r, 8).value
        title = title_from_link_cell(link_cell, notes)
        if track and not str(title).lower().startswith(str(track).lower()):
            title = f"{track}: {title}" if title == "Marketing faculty position" else title
        rows.append(
            {
                "Country": CONTINENT_COUNTRY.get(continent, continent or None),
                "Location": None,
                "University": str(school).strip(),
                "School": None,
                "Deadline": deadline,
                "Job Post": f"{title} | {url}" if url else title,
                "Required Material": str(notes).replace("\n", " ").strip()[:300] if notes else None,
                "Submission Channel": "MJM-NTT",
                "Application Link": url,
                "Email": None,
                "Submission Date": None,
                "Reference Status": None,
                "Interview Date": None,
                "Job Type": job_type_from_track(str(track) if track else None, "NTT"),
                "_title": title,
            }
        )
    return rows


def write_xlsx(rows: list[dict], path: Path, sheet: str = "Marketing Listings") -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    link_font = Font(color="0563C1", underline="single")
    for c, name in enumerate(COLUMNS, 1):
        ws.cell(1, c, name).font = Font(bold=True)
    for r_i, row in enumerate(rows, 2):
        for c, name in enumerate(COLUMNS, 1):
            val = row.get(name)
            cell = ws.cell(r_i, c, val)
            if name == "Job Post":
                url = row.get("Application Link")
                title = row.get("_title") or "Marketing listing"
                if url:
                    cell.value = title
                    cell.hyperlink = str(url)
                    cell.font = link_font
            elif name == "Application Link" and val:
                cell.hyperlink = str(val)
                cell.font = link_font
    for c in range(1, len(COLUMNS) + 1):
        ws.column_dimensions[get_column_letter(c)].width = 16
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def update_app_sheet(rows: list[dict]) -> None:
    if not APP_XLSX.exists():
        return
    wb = load_workbook(APP_XLSX)
    if "Marketing Listings" in wb.sheetnames:
        del wb["Marketing Listings"]
    ws = wb.create_sheet("Marketing Listings")
    link_font = Font(color="0563C1", underline="single")
    for c, name in enumerate(COLUMNS, 1):
        ws.cell(1, c, name).font = Font(bold=True)
    for r_i, row in enumerate(rows, 2):
        for c, name in enumerate(COLUMNS, 1):
            val = row.get(name)
            cell = ws.cell(r_i, c, val)
            if name == "Job Post" and row.get("Application Link"):
                cell.value = row.get("_title") or "Marketing listing"
                cell.hyperlink = str(row["Application Link"])
                cell.font = link_font
            elif name == "Application Link" and val:
                cell.hyperlink = str(val)
                cell.font = link_font
    wb.save(APP_XLSX)


def find_source() -> Path:
    ensure_dirs()
    candidates = [
        DATA_MKT / SOURCE_XLSX_NAME,
        DATA_MKT / "Marketing Job Market - 2026 Job Postings.xlsx",
    ]
    for p in candidates:
        if p.exists():
            return p
    raise FileNotFoundError(
        f"Put '{SOURCE_XLSX_NAME}' in {DATA_MKT} (copied from Downloads)"
    )


def main() -> None:
    ensure_dirs()
    src = find_source()
    print(f"Reading {src}")
    wb = load_workbook(src, data_only=False)
    rows = read_job_postings(wb) + read_ntt(wb)
    wb.close()

    # de-dupe by university + url
    seen, unique = set(), []
    for r in rows:
        key = (r.get("University"), r.get("Application Link"), r.get("Submission Channel"))
        if key in seen:
            continue
        seen.add(key)
        unique.append(r)

    today = date.today().isoformat()
    structured = DATA_MKT / f"marketing_structured_{today}.xlsx"
    write_xlsx(unique, structured)
    update_app_sheet(unique)

    n_tt = sum(1 for r in unique if r.get("Submission Channel") == "MJM")
    n_ntt = sum(1 for r in unique if r.get("Submission Channel") == "MJM-NTT")
    print(f"Compiled {len(unique)} marketing listings (TT/open {n_tt}, NTT/Ind {n_ntt})")
    print(f"  {structured}")
    print(f"Updated Marketing Listings sheet in {APP_XLSX}")


if __name__ == "__main__":
    main()
