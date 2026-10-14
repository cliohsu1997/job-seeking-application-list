# -*- coding: utf-8 -*-
"""Filter AEA/EJM/CICE to Asia-Pacific targets and combine with provincial list.

Keeps website rows whose Country is China (incl. HK/Taiwan/Macau), Malaysia,
Japan, Australia, or New Zealand. Merges with china_macau_taiwan provincial
jobs. Writes both Excel and PDF:

  clean/economics/asia_pacific_listings_YYYY-MM-DD.xlsx
  clean/asia_pacific_listings_YYYY-MM-DD.pdf

Usage:
  python combine_asia_pacific.py
  python combine_asia_pacific.py --fetch-year
"""
from __future__ import annotations

import argparse
import json
import re
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from china_universities_data import REGIONS
from columns import COLUMNS
from download_china_provinces import job_to_row
from enrich_job_metadata import (
    build_ranking_index,
    enrich_jobs,
    published_year_from_text,
    ranking_for_university,
)
from export_clean_pdf import export_one, latest
from paths import CLEAN, CLEAN_ECON, DATA_ECON, ensure_dirs
from summarize_china_jobs import _dedupe, load_macau_taiwan, load_region_jobs

TODAY = date.today().isoformat()

_ALLOWED = {
    "CHINA",
    "CN",
    "PRC",
    "HONG KONG",
    "HONGKONG",
    "HK",
    "TAIWAN",
    "TW",
    "MACAU",
    "MACAO",
    "MACAO CHINA",
    "MACAO; CHINA",
    "MO",
    "MALAYSIA",
    "MY",
    "JAPAN",
    "JP",
    "AUSTRALIA",
    "AU",
    "NEW ZEALAND",
    "NEWZEALAND",
    "NZ",
}


def _norm_country(value) -> str:
    s = str(value or "").strip().upper()
    s = re.sub(r"[\s._]+", " ", s)
    s = s.replace(",", " ").replace(";", " ")
    s = re.sub(r"\s+", " ", s).strip()
    if "MACAO" in s or "MACAU" in s:
        return "MACAO; CHINA" if "CHINA" in s else "MACAU"
    if "HONG" in s and "KONG" in s:
        return "HONG KONG"
    if "NEW" in s and "ZEALAND" in s:
        return "NEW ZEALAND"
    return s


def country_allowed(value) -> bool:
    n = _norm_country(value)
    if n in _ALLOWED:
        return True
    compact = n.replace(" ", "")
    return compact in {a.replace(" ", "") for a in _ALLOWED}


def load_website_rows(path: Path | None = None) -> list[dict]:
    src = path or latest(CLEAN_ECON, "all_listings_*.xlsx")
    if not src:
        raise FileNotFoundError(
            "No clean/economics/all_listings_*.xlsx - run combine_clean.py first"
        )
    wb = load_workbook(src, data_only=True)
    ws = wb.active
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    rows = []
    for r in range(2, ws.max_row + 1):
        item = {h: ws.cell(r, c).value for c, h in enumerate(headers, 1) if h}
        for col in ("Job Post", "Application Link"):
            if col not in headers:
                continue
            cell = ws.cell(r, headers.index(col) + 1)
            if cell.hyperlink and cell.hyperlink.target:
                href = cell.hyperlink.target
                val = cell.value
                if col == "Job Post" and val and href:
                    item[col] = "%s | %s" % (val, href)
                elif col == "Application Link":
                    item[col] = href
        rows.append(item)
    wb.close()
    return rows


def website_row_to_internal(row: dict) -> dict:
    link = row.get("Application Link") or ""
    job_post = str(row.get("Job Post") or "")
    title = job_post
    if " | " in job_post:
        left, right = job_post.rsplit(" | ", 1)
        if right.strip().startswith("http"):
            title = left.strip()
            if not link:
                link = right.strip()
    deadline = row.get("Deadline")
    if isinstance(deadline, datetime):
        deadline = deadline.strftime("%Y-%m-%d")
    return {
        "region": "website",
        "Country": row.get("Country"),
        "Location": row.get("Location"),
        "University": row.get("University"),
        "School": row.get("School"),
        "Deadline": deadline,
        "JobPostTitle": title,
        "JobPostUrl": link,
        "RequiredMaterial": row.get("Required Material"),
        "SubmissionChannel": row.get("Submission Channel"),
        "ApplicationLink": link,
        "Email": row.get("Email"),
        "JobType": (
            "Faculty"
            if (row.get("Job Type") or "") == "Academic"
            else (row.get("Job Type") or "Faculty")
        ),
        "PublishedYear": row.get("Published Year") or "",
        "CollegeRanking": row.get("College Ranking") or "",
        "_source": "aea_ejm_cice",
    }


def load_provincial(*, prefer: str = "verified") -> list[dict]:
    coastal: list[dict] = []
    for region in REGIONS:
        coastal.extend(load_region_jobs(region, prefer))
    macau, taiwan = load_macau_taiwan()
    return _dedupe(coastal + macau + taiwan)


def ensure_metadata(jobs: list[dict], *, fetch_page: bool) -> list[dict]:
    idx = build_ranking_index()
    out = []
    for j in jobs:
        year = j.get("PublishedYear") or j.get("Published Year") or ""
        rank = j.get("CollegeRanking") or j.get("College Ranking") or ""
        if not year:
            year = published_year_from_text(
                j.get("JobPostTitle") or "",
                j.get("JobPostUrl") or j.get("ApplicationLink") or "",
                j.get("RequiredMaterial") or "",
            )
        if not rank:
            rank = ranking_for_university(j.get("University") or "", idx)
        j = dict(j)
        j["PublishedYear"] = year
        j["CollegeRanking"] = rank
        out.append(j)
    if fetch_page:
        need = [j for j in out if not j.get("PublishedYear")]
        if need:
            print(
                "  fetching pages for %d jobs missing Published Year..."
                % len(need)
            )
            fixed = enrich_jobs(need, fetch_page=True, ranking_index=idx)
            by_url = {
                (j.get("JobPostUrl") or j.get("ApplicationLink") or ""): j
                for j in fixed
            }
            merged = []
            for j in out:
                u = j.get("JobPostUrl") or j.get("ApplicationLink") or ""
                merged.append(by_url.get(u, j) if u in by_url else j)
            out = merged
    return out


def to_app_row(job: dict) -> dict:
    row = job_to_row(job)
    row["Published Year"] = job.get("PublishedYear") or job.get("Published Year") or ""
    row["College Ranking"] = (
        job.get("CollegeRanking") or job.get("College Ranking") or ""
    )
    # Finals: Faculty | Postdoc | Non-Academic (Academic -> Faculty)
    jt = row.get("Job Type") or job.get("JobType") or ""
    if jt == "Academic":
        jt = "Faculty"
    row["Job Type"] = jt or "Faculty"
    for gone in ("Submission Date", "Reference Status", "Interview Date"):
        row.pop(gone, None)
    return row


def _sort_rows(rows: list[dict]) -> list[dict]:
    def sort_key(r):
        d = r.get("Deadline")
        dkey = d if isinstance(d, datetime) else datetime.max
        return (
            str(r.get("Country") or ""),
            str(r.get("University") or ""),
            dkey,
            str(r.get("Job Post") or ""),
        )

    return sorted(rows, key=sort_key)


def _write_sheet(ws, rows: list[dict]) -> None:
    link_font = Font(color="0563C1", underline="single")
    for c, name in enumerate(COLUMNS, 1):
        ws.cell(1, c, name).font = Font(bold=True)
    for r_i, row in enumerate(rows, 2):
        for c, name in enumerate(COLUMNS, 1):
            val = row.get(name)
            cell = ws.cell(r_i, c, val)
            if name == "Job Post" and row.get("Application Link"):
                cell.value = row.get("_title") or "Faculty"
                cell.hyperlink = str(row["Application Link"])
                cell.font = link_font
            elif name == "Application Link" and val:
                cell.hyperlink = str(val)
                cell.font = link_font
    for c in range(1, len(COLUMNS) + 1):
        ws.column_dimensions[get_column_letter(c)].width = 16


SHEET_WEBSITE = "AEA EJM CICE Asia-Pacific"
SHEET_PROVINCIAL = "Provincial China+MT"


def write_two_section_xlsx(
    web_rows: list[dict],
    prov_rows: list[dict],
    path: Path,
) -> Path:
    """Two sheets, identical columns: top AEA/EJM/CICE, bottom provincial."""
    wb = Workbook()
    ws1 = wb.active
    ws1.title = SHEET_WEBSITE[:31]
    _write_sheet(ws1, web_rows)
    ws2 = wb.create_sheet(SHEET_PROVINCIAL[:31])
    _write_sheet(ws2, prov_rows)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def write_one_sheet_xlsx(rows: list[dict], path: Path, sheet: str) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = sheet[:31]
    _write_sheet(ws, rows)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def _export_pair(
    rows: list[dict],
    *,
    stem: str,
    sheet: str,
    title: str,
) -> tuple[Path, Path]:
    xlsx = CLEAN / ("%s_%s.xlsx" % (stem, TODAY))
    pdf = CLEAN / ("%s_%s.pdf" % (stem, TODAY))
    write_one_sheet_xlsx(rows, xlsx, sheet)
    export_one(xlsx, pdf, title, use_cjk=True)
    return xlsx, pdf


def combine(
    *,
    prefer: str = "verified",
    fetch_page: bool = False,
    website_xlsx: Path | None = None,
) -> dict:
    """Compile AEA + provincial(+Macau/Taiwan) finals into clean/ root.

    Run enrich_website_from_pages + enrich_provincial_from_pages first for
    Faculty/Postdoc + Deadline from listing pages.
    """
    ensure_dirs()

    print("\n########## ENRICH Published Year + College Ranking ##########")
    from enrich_job_metadata import enrich_region

    for region in REGIONS:
        try:
            enrich_region(region, fetch_page=False)
        except FileNotFoundError:
            pass

    web_raw = load_website_rows(website_xlsx)
    web_keep = [r for r in web_raw if country_allowed(r.get("Country"))]
    print(
        "Website rows: %d total -> %d Asia-Pacific keep"
        % (len(web_raw), len(web_keep))
    )
    web_jobs = [website_row_to_internal(r) for r in web_keep]
    provincial = load_provincial(prefer=prefer)
    print("Provincial + Macau/Taiwan: %d" % len(provincial))

    web_jobs = ensure_metadata(web_jobs, fetch_page=fetch_page)
    provincial = ensure_metadata(provincial, fetch_page=fetch_page)

    web_rows = _sort_rows([to_app_row(j) for j in web_jobs])
    prov_rows = _sort_rows([to_app_row(j) for j in provincial])

    CLEAN.mkdir(parents=True, exist_ok=True)

    aea_xlsx, aea_pdf = _export_pair(
        web_rows,
        stem="aea_listings",
        sheet=SHEET_WEBSITE,
        title="AEA/EJM/CICE Asia-Pacific (CN/MY/JP/AU/NZ)",
    )
    prov_xlsx, prov_pdf = _export_pair(
        prov_rows,
        stem="provincial_listings",
        sheet=SHEET_PROVINCIAL,
        title="Provincial China + Macau/Taiwan - Faculty + Postdoc",
    )

    xlsx = CLEAN / ("asia_pacific_listings_%s.xlsx" % TODAY)
    pdf = CLEAN / ("asia_pacific_listings_%s.pdf" % TODAY)
    write_two_section_xlsx(web_rows, prov_rows, xlsx)
    export_one(
        xlsx,
        pdf,
        "Asia-Pacific Economics - AEA/EJM/CICE + Provincial",
        use_cjk=True,
    )

    all_jobs = web_jobs + provincial
    (DATA_ECON / ("asia_pacific_jobs_%s.json" % TODAY)).write_text(
        json.dumps(all_jobs, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    summary = {
        "date": TODAY,
        "website_total": len(web_raw),
        "website_kept": len(web_keep),
        "provincial": len(provincial),
        "macau_taiwan": sum(
            1
            for j in provincial
            if str(j.get("region") or "").lower() in ("macau", "taiwan")
            or str(j.get("Country") or "").upper()
            in ("MACAU", "MACAO", "TAIWAN", "TW", "MO")
        ),
        "combined": len(web_rows) + len(prov_rows),
        "columns": COLUMNS,
        "sheets": [SHEET_WEBSITE, SHEET_PROVINCIAL],
        "aea_xlsx": str(aea_xlsx),
        "aea_pdf": str(aea_pdf),
        "provincial_xlsx": str(prov_xlsx),
        "provincial_pdf": str(prov_pdf),
        "xlsx": str(xlsx),
        "pdf": str(pdf),
        "countries_website": sorted(
            {_norm_country(r.get("Country")) for r in web_keep if r.get("Country")}
        ),
    }
    (DATA_ECON / ("asia_pacific_summary_%s.json" % TODAY)).write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(
        "Combined website=%d provincial(+MT)=%d -> %s | %s"
        % (len(web_rows), len(prov_rows), pdf, xlsx)
    )
    print("AEA ->", aea_pdf)
    print("Provincial ->", prov_pdf)
    return summary


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--prefer", choices=("verified", "validated"), default="verified")
    ap.add_argument(
        "--fetch-year",
        action="store_true",
        help="Re-fetch pages when Published Year is missing (slow).",
    )
    ap.add_argument("--website-xlsx", type=Path, default=None)
    args = ap.parse_args()
    print(
        json.dumps(
            combine(
                prefer=args.prefer,
                fetch_page=args.fetch_year,
                website_xlsx=args.website_xlsx,
            ),
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
