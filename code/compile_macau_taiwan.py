"""
Compile Macau + Taiwan economics faculty openings found via department career pages.

Outputs:
  data/economics/macau_taiwan_departments_*.json  (dept -> career portal map)
  data/economics/macau_taiwan_listings_*.json
  clean/economics/macau_taiwan_listings_*.xlsx
"""
from __future__ import annotations

import json
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from columns import COLUMNS
from paths import CLEAN_ECON, DATA_ECON, ensure_dirs

# Department directory (economics / closely related units) + career portal
DEPARTMENTS = [
    # --- Macau ---
    {
        "region": "Macau",
        "university": "University of Macau",
        "unit": "Department of Economics (FSS)",
        "dept_url": "https://econ.fss.um.edu.mo/",
        "career_portal": "https://career.admo.um.edu.mo/",
    },
    {
        "region": "Macau",
        "university": "University of Macau",
        "unit": "Department of Finance and Business Economics (FBA)",
        "dept_url": "https://fba.um.edu.mo/fbe/",
        "career_portal": "https://career.admo.um.edu.mo/",
    },
    {
        "region": "Macau",
        "university": "City University of Macau",
        "unit": "Faculty of Finance / Economics & Digital Economics",
        "dept_url": "https://hro.cityu.edu.mo/en/category/job-application/teaching-en/",
        "career_portal": "https://career.cityu.edu.mo",
    },
    {
        "region": "Macau",
        "university": "City University of Macau",
        "unit": "Institute of Urban and Sustainable Development (Economics)",
        "dept_url": "https://hro.cityu.edu.mo/en/category/job-application/teaching-en/",
        "career_portal": "https://career.cityu.edu.mo",
    },
    {
        "region": "Macau",
        "university": "Macau University of Science and Technology",
        "unit": "School of Business — Department of Applied Economics",
        "dept_url": "https://msb.must.edu.mo/",
        "career_portal": "https://www.must.edu.mo/en/careers",
    },
    {
        "region": "Macau",
        "university": "Macao Polytechnic University",
        "unit": "Faculty of Business (no standalone Economics dept)",
        "dept_url": "https://mpusite.mpu.edu.mo/esce/en/about_faculty.php",
        "career_portal": "https://www.mpu.edu.mo/vacancy",
        "note": "No open Economics faculty posting found at check time",
    },
    # --- Taiwan ---
    {
        "region": "Taiwan",
        "university": "National Taiwan University",
        "unit": "Department of Economics",
        "dept_url": "https://econ.ntu.edu.tw/",
        "career_portal": "https://econ.ntu.edu.tw/archives/18834",
    },
    {
        "region": "Taiwan",
        "university": "National Chengchi University",
        "unit": "Department of Economics",
        "dept_url": "https://econo.nccu.edu.tw/",
        "career_portal": "https://econo.nccu.edu.tw/eng/PageDoc/Detail?fid=1052&id=42260",
    },
    {
        "region": "Taiwan",
        "university": "National Tsing Hua University",
        "unit": "Taipei School of Economics and Political Science (TSE)",
        "dept_url": "http://www.tse.nthu.edu.tw/",
        "career_portal": "http://www.tse.nthu.edu.tw/post/82-tse-job-opening",
    },
    {
        "region": "Taiwan",
        "university": "National Cheng Kung University",
        "unit": "Department of Economics",
        "dept_url": "https://econ.ncku.edu.tw/",
        "career_portal": "https://ann.ncku.edu.tw/p/404-1013-305836.php?Lang=zh-tw",
    },
    {
        "region": "Taiwan",
        "university": "Soochow University",
        "unit": "Department of Economics",
        "dept_url": "http://www.scu.edu.tw/econ/",
        "career_portal": "https://web-ch.scu.edu.tw/econ/opinion/16118/31786",
    },
    {
        "region": "Taiwan",
        "university": "National Sun Yat-sen University",
        "unit": "Department of Political Economy",
        "dept_url": "https://pe.nsysu.edu.tw/",
        "career_portal": "https://www.nstc.gov.tw/folksonomy/detail/4b115db7-4cd9-4774-8ba4-31335f5ea07c?l=en",
    },
    {
        "region": "Taiwan",
        "university": "National Chung Hsing University",
        "unit": "Department of Applied Economics",
        "dept_url": "https://www.nchu.edu.tw/",
        "career_portal": "https://www2.nchu.edu.tw/en-news-detail/id/1254",
        "note": "Prior call deadline 2026-07-02 (likely closed)",
    },
]

# Open faculty jobs verified from career pages (as of scrape date)
JOBS = [
    {
        "Country": "Macau",
        "Location": "Taipa",
        "University": "University of Macau",
        "School": "Faculty of Social Sciences — Department of Economics",
        "Deadline": None,
        "Job Post": "Assistant Professor in Economics (macro/monetary) | https://career.admo.um.edu.mo/fss_deco_ap_04_2026/",
        "Required Material": "Online application via Career@UM; PhD in Economics",
        "Submission Channel": "UM-Career",
        "Application Link": "https://career.admo.um.edu.mo/fss_deco_ap_04_2026/",
        "Email": "fss.economics@um.edu.mo",
        "Job Type": "Academic",
        "_title": "Assistant Professor in Economics (macro/monetary)",
    },
    {
        "Country": "Macau",
        "Location": "Taipa",
        "University": "University of Macau",
        "School": "Faculty of Business Administration — Finance & Business Economics",
        "Deadline": None,
        "Job Post": "Associate/Assistant Professor in Microeconometrics | https://career.admo.um.edu.mo/fba_fbe_mec_aap_08_2026/",
        "Required Material": "Online application; 3 referees; PhD in Economics",
        "Submission Channel": "UM-Career",
        "Application Link": "https://career.admo.um.edu.mo/fba_fbe_mec_aap_08_2026/",
        "Email": "vacancy@um.edu.mo",
        "Job Type": "Academic",
        "_title": "Associate/Assistant Professor in Microeconometrics",
    },
    {
        "Country": "Macau",
        "Location": "Taipa",
        "University": "University of Macau",
        "School": "Faculty of Business Administration — Integrated Resort & Tourism Management",
        "Deadline": None,
        "Job Post": "Assistant Professor in Tourism Economics | https://career.admo.um.edu.mo/fba_drtm_te_ap_05_2026/",
        "Required Material": "Online application via Career@UM",
        "Submission Channel": "UM-Career",
        "Application Link": "https://career.admo.um.edu.mo/fba_drtm_te_ap_05_2026/",
        "Email": "vacancy@um.edu.mo",
        "Job Type": "Academic",
        "_title": "Assistant Professor in Tourism Economics",
    },
    {
        "Country": "Macau",
        "Location": "Taipa",
        "University": "City University of Macau",
        "School": "Faculty of Finance",
        "Deadline": None,
        "Job Post": "Professor/Associate/Assistant Professor — Economics & Digital Economics | https://hro.cityu.edu.mo/en/2026/10/05/professor-associate-professor-assistant-professor-faculty-of-finance/",
        "Required Material": "Resume; ID; certificates; cover letter; 3 referees — apply online",
        "Submission Channel": "CityU-Career",
        "Application Link": "https://career.cityu.edu.mo",
        "Email": "jobs@cityu.edu.mo",
        "Job Type": "Academic",
        "_title": "Professor/Associate/Assistant — Economics & Digital Economics",
    },
    {
        "Country": "Macau",
        "Location": "Taipa",
        "University": "City University of Macau",
        "School": "Institute of Urban and Sustainable Development",
        "Deadline": None,
        "Job Post": "Professor/Associate/Assistant Professor — Economics | https://hro.cityu.edu.mo/en/2026/10/05/professor-associate-professor-assistant-professor-institute-of-urban-and-sustainable-development-2/",
        "Required Material": "Resume; ID; certificates; cover letter; 3 referees — apply online",
        "Submission Channel": "CityU-Career",
        "Application Link": "https://career.cityu.edu.mo",
        "Email": "jobs@cityu.edu.mo",
        "Job Type": "Academic",
        "_title": "Professor/Associate/Assistant — Economics (IUSD)",
    },
    {
        "Country": "Macau",
        "Location": "Taipa",
        "University": "Macau University of Science and Technology",
        "School": "School of Business — Applied Economics (among areas)",
        "Deadline": None,
        "Job Post": "Professor/Associate/Assistant/Lecturer — Economics (School of Business rolling call) | https://www.must.edu.mo/en/careers",
        "Required Material": "CV; certificates; publications; apply via MUST careers",
        "Submission Channel": "MUST-Career",
        "Application Link": "https://www.must.edu.mo/en/careers",
        "Email": "hro@must.edu.mo",
        "Job Type": "Academic",
        "_title": "Faculty (Economics area) — School of Business",
    },
    {
        "Country": "Taiwan",
        "Location": "Taipei",
        "University": "National Taiwan University",
        "School": "Department of Economics",
        "Deadline": datetime(2027, 1, 31),
        "Job Post": "Assistant Professor or above (up to 3; trade/finance/public/urban priority) | https://econ.ntu.edu.tw/archives/18834",
        "Required Material": "CV; cover letter; JMP; research papers; transcripts; letters of reference",
        "Submission Channel": "Dept/JOE",
        "Application Link": "https://econ.ntu.edu.tw/archives/18834",
        "Email": "kuanmingchen@ntu.edu.tw",
        "Job Type": "Academic",
        "_title": "Assistant Professor or above (up to 3)",
    },
    {
        "Country": "Taiwan",
        "Location": "Taipei",
        "University": "National Chengchi University",
        "School": "Department of Economics",
        "Deadline": datetime(2026, 11, 30),
        "Job Post": "Open-rank tenure-track faculty (1) | https://econo.nccu.edu.tw/eng/PageDoc/Detail?fid=1052&id=42260",
        "Required Material": "Cover letter; CV; diploma/expected grad letter; transcripts; JMP/pubs; syllabi; 3 letters",
        "Submission Channel": "Email",
        "Application Link": "https://econo.nccu.edu.tw/eng/PageDoc/Detail?fid=1052&id=42260",
        "Email": "econ@nccu.edu.tw",
        "Job Type": "Academic",
        "_title": "Open-rank tenure-track faculty (1)",
    },
    {
        "Country": "Taiwan",
        "Location": "Hsinchu",
        "University": "National Tsing Hua University",
        "School": "Taipei School of Economics and Political Science (TSE)",
        "Deadline": datetime(2026, 11, 15),
        "Job Post": "Faculty — Macroeconomics / International Trade (1 of 2 TSE posts) | http://www.tse.nthu.edu.tw/post/82-tse-job-opening",
        "Required Material": "Apply via AEA JOE (Economics field); English instruction",
        "Submission Channel": "JOE",
        "Application Link": "http://www.tse.nthu.edu.tw/post/82-tse-job-opening",
        "Email": "tse@my.nthu.edu.tw",
        "Job Type": "Academic",
        "_title": "Faculty — Macroeconomics / International Trade",
    },
    {
        "Country": "Taiwan",
        "Location": "Tainan",
        "University": "National Cheng Kung University",
        "School": "Department of Economics",
        "Deadline": datetime(2026, 11, 28),
        "Job Post": "Assistant/Associate/Professor (any field) | https://ann.ncku.edu.tw/p/404-1013-305836.php?Lang=zh-tw",
        "Required Material": "Cover letter; CV; JMP; transcripts (AP); 3 letters; diploma — via AEA JOE",
        "Submission Channel": "JOE",
        "Application Link": "https://ann.ncku.edu.tw/p/404-1013-305836.php?Lang=zh-tw",
        "Email": "em56300@email.ncku.edu.tw",
        "Job Type": "Academic",
        "_title": "Assistant/Associate/Professor (any field)",
    },
    {
        "Country": "Taiwan",
        "Location": "Taipei",
        "University": "Soochow University",
        "School": "Department of Economics",
        "Deadline": datetime(2026, 10, 30),
        "Job Post": "Assistant Professor or above (2; macro empirics priority) | https://web-ch.scu.edu.tw/econ/opinion/16118/31786",
        "Required Material": "Diploma; transcripts; CV; thesis/pubs; 3 letters; syllabi — paper + email PDF",
        "Submission Channel": "Email/Mail",
        "Application Link": "https://web-ch.scu.edu.tw/econ/opinion/16118/31786",
        "Email": "waccla@scu.edu.tw",
        "Job Type": "Academic",
        "_title": "Assistant Professor or above (2)",
    },
    {
        "Country": "Taiwan",
        "Location": "Kaohsiung",
        "University": "National Sun Yat-sen University",
        "School": "Department of Political Economy",
        "Deadline": datetime(2026, 12, 31),
        "Job Post": "Faculty (2) — Political Economy / Economics | https://www.nstc.gov.tw/folksonomy/detail/4b115db7-4cd9-4774-8ba4-31335f5ea07c?l=en",
        "Required Material": "Cover letter; CV; research & teaching statements; pubs; degree certs; 3 letters",
        "Submission Channel": "Email",
        "Application Link": "https://www.nstc.gov.tw/folksonomy/detail/4b115db7-4cd9-4774-8ba4-31335f5ea07c?l=en",
        "Email": "poleco@mail.nsysu.edu.tw",
        "Job Type": "Academic",
        "_title": "Faculty (2) — Political Economy / Economics",
    },
]


def write_xlsx(rows: list[dict], path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Macau Taiwan Econ"
    link_font = Font(color="0563C1", underline="single")
    for c, name in enumerate(COLUMNS, 1):
        ws.cell(1, c, name).font = Font(bold=True)
    for r_i, row in enumerate(rows, 2):
        for c, name in enumerate(COLUMNS, 1):
            val = row.get(name)
            cell = ws.cell(r_i, c, val)
            if name == "Job Post" and row.get("Application Link"):
                cell.value = row.get("_title") or "Faculty position"
                cell.hyperlink = str(row["Application Link"])
                cell.font = link_font
            elif name == "Application Link" and val:
                cell.hyperlink = str(val)
                cell.font = link_font
    for c in range(1, len(COLUMNS) + 1):
        ws.column_dimensions[get_column_letter(c)].width = 16
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def main() -> None:
    ensure_dirs()
    today = date.today().isoformat()

    dept_path = DATA_ECON / f"macau_taiwan_departments_{today}.json"
    dept_path.write_text(json.dumps(DEPARTMENTS, ensure_ascii=False, indent=2), encoding="utf-8")

    jobs_path = DATA_ECON / f"macau_taiwan_listings_{today}.json"
    serializable = []
    for j in JOBS:
        item = {k: v for k, v in j.items() if not k.startswith("_")}
        if isinstance(item.get("Deadline"), datetime):
            item["Deadline"] = item["Deadline"].strftime("%Y-%m-%d")
        serializable.append(item)
    jobs_path.write_text(json.dumps(serializable, ensure_ascii=False, indent=2), encoding="utf-8")

    clean_path = CLEAN_ECON / f"macau_taiwan_listings_{today}.xlsx"
    # normalize rows for COLUMNS
    rows = []
    for j in JOBS:
        row = {c: j.get(c) for c in COLUMNS}
        row["_title"] = j.get("_title")
        rows.append(row)
    write_xlsx(rows, clean_path)

    n_mac = sum(1 for j in JOBS if j["Country"] == "Macau")
    n_tw = sum(1 for j in JOBS if j["Country"] == "Taiwan")
    print(f"Departments mapped: {len(DEPARTMENTS)}")
    print(f"Open jobs: {len(JOBS)} (Macau {n_mac}, Taiwan {n_tw})")
    print(f"  {dept_path}")
    print(f"  {jobs_path}")
    print(f"  {clean_path}")


if __name__ == "__main__":
    main()
