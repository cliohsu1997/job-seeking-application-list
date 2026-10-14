"""
Build full Macau + Taiwan university catalogs, fetch department/career pages,
and extract economics faculty openings into:

  data/economics/macau/
  data/economics/taiwan/
"""
from __future__ import annotations

import csv
import json
import re
import ssl
import time
from datetime import date, datetime
from html import unescape
from io import StringIO
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from columns import COLUMNS
from paths import CLEAN_ECON, DATA_ECON, ensure_dirs

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
CTX = ssl._create_unverified_context()
TODAY = date.today().isoformat()
SLEEP = 0.4

MACAU_DIR = DATA_ECON / "macau"
TAIWAN_DIR = DATA_ECON / "taiwan"
MOE_CSV_URL = (
    "https://udb.moe.edu.tw/ulist/Resource/download/"
    "ulistdepartmentlist_college?extension=csv"
)

# Official Macau HEIs (DSEDJ / Wikipedia)
MACAU_UNIVERSITIES = [
    {
        "name": "University of Macau",
        "chinese": "澳門大學",
        "type": "Public",
        "home": "https://www.um.edu.mo/",
        "career": "https://career.admo.um.edu.mo/",
        "econ_units": [
            {
                "unit": "Department of Economics (FSS)",
                "url": "https://econ.fss.um.edu.mo/",
            },
            {
                "unit": "Department of Finance and Business Economics (FBA)",
                "url": "https://fba.um.edu.mo/fbe/",
            },
        ],
    },
    {
        "name": "Macao Polytechnic University",
        "chinese": "澳門理工大學",
        "type": "Public",
        "home": "https://www.mpu.edu.mo/",
        "career": "https://www.mpu.edu.mo/vacancy",
        "econ_units": [
            {"unit": "Faculty of Business", "url": "https://mpusite.mpu.edu.mo/esce/en/"}
        ],
    },
    {
        "name": "Macao University of Tourism",
        "chinese": "澳門旅遊大學",
        "type": "Public",
        "home": "https://www.utm.edu.mo/",
        "career": "https://www.utm.edu.mo/en/about-utm/career/",
        "econ_units": [],
    },
    {
        "name": "Academy of Public Security Forces of Macao",
        "chinese": "澳門保安部隊高等學校",
        "type": "Public",
        "home": "https://www.fsm.gov.mo/",
        "career": None,
        "econ_units": [],
    },
    {
        "name": "Macau University of Science and Technology",
        "chinese": "澳門科技大學",
        "type": "Private",
        "home": "https://www.must.edu.mo/",
        "career": "https://www.must.edu.mo/en/careers",
        "econ_units": [
            {
                "unit": "Department of Applied Economics (School of Business)",
                "url": "https://msb.must.edu.mo/",
            }
        ],
    },
    {
        "name": "City University of Macau",
        "chinese": "澳門城市大學",
        "type": "Private",
        "home": "https://www.cityu.edu.mo/",
        "career": "https://hro.cityu.edu.mo/en/category/job-application/teaching-en/",
        "econ_units": [
            {
                "unit": "Faculty of Finance (Economics / Digital Economics)",
                "url": "https://hro.cityu.edu.mo/en/category/job-application/teaching-en/",
            }
        ],
    },
    {
        "name": "University of Saint Joseph",
        "chinese": "聖若瑟大學",
        "type": "Private",
        "home": "https://www.usj.edu.mo/",
        "career": "https://www.usj.edu.mo/en/careers/",
        "econ_units": [],
    },
    {
        "name": "Kiang Wu Nursing College of Macau",
        "chinese": "澳門鏡湖護理學院",
        "type": "Private",
        "home": "https://www.kwnc.edu.mo/",
        "career": "https://www.kwnc.edu.mo/",
        "econ_units": [],
    },
    {
        "name": "Macau Institute of Management",
        "chinese": "澳門管理學院",
        "type": "Private",
        "home": "https://www.mma.org.mo/",
        "career": "https://www.mma.org.mo/",
        "econ_units": [],
    },
    {
        "name": "Macau Millennium College",
        "chinese": "中西創新學院",
        "type": "Private",
        "home": "https://www.mmc.edu.mo/",
        "career": "https://www.mmc.edu.mo/",
        "econ_units": [],
    },
]

# Curated open jobs (filled from career pages) — Application columns
KNOWN_JOBS = [
    # --- Macau ---
    dict(
        Country="Macau",
        Location="Taipa",
        University="University of Macau",
        School="Faculty of Social Sciences — Department of Economics",
        Deadline=None,
        JobPostTitle="Assistant Professor in Economics (macro/monetary)",
        JobPostUrl="https://career.admo.um.edu.mo/fss_deco_ap_04_2026/",
        RequiredMaterial="Online application via Career@UM; PhD in Economics",
        SubmissionChannel="UM-Career",
        ApplicationLink="https://career.admo.um.edu.mo/fss_deco_ap_04_2026/",
        Email="fss.economics@um.edu.mo",
        JobType="Academic",
    ),
    dict(
        Country="Macau",
        Location="Taipa",
        University="University of Macau",
        School="FBA — Finance & Business Economics",
        Deadline=None,
        JobPostTitle="Associate/Assistant Professor in Microeconometrics",
        JobPostUrl="https://career.admo.um.edu.mo/fba_fbe_mec_aap_08_2026/",
        RequiredMaterial="Online application; 3 referees; PhD in Economics",
        SubmissionChannel="UM-Career",
        ApplicationLink="https://career.admo.um.edu.mo/fba_fbe_mec_aap_08_2026/",
        Email="vacancy@um.edu.mo",
        JobType="Academic",
    ),
    dict(
        Country="Macau",
        Location="Taipa",
        University="University of Macau",
        School="FBA — Integrated Resort & Tourism Management",
        Deadline=None,
        JobPostTitle="Assistant Professor in Tourism Economics",
        JobPostUrl="https://career.admo.um.edu.mo/fba_drtm_te_ap_05_2026/",
        RequiredMaterial="Online application; 3 referees",
        SubmissionChannel="UM-Career",
        ApplicationLink="https://career.admo.um.edu.mo/fba_drtm_te_ap_05_2026/",
        Email="vacancy@um.edu.mo",
        JobType="Academic",
    ),
    dict(
        Country="Macau",
        Location="Taipa",
        University="City University of Macau",
        School="Faculty of Finance",
        Deadline=None,
        JobPostTitle="Professor/Associate/Assistant — Economics & Digital Economics",
        JobPostUrl="https://hro.cityu.edu.mo/en/2026/10/05/professor-associate-professor-assistant-professor-faculty-of-finance/",
        RequiredMaterial="Resume; ID; certificates; cover letter; 3 referees",
        SubmissionChannel="CityU-Career",
        ApplicationLink="https://career.cityu.edu.mo",
        Email="jobs@cityu.edu.mo",
        JobType="Academic",
    ),
    dict(
        Country="Macau",
        Location="Taipa",
        University="City University of Macau",
        School="Institute of Urban and Sustainable Development",
        Deadline=None,
        JobPostTitle="Professor/Associate/Assistant — Economics (IUSD)",
        JobPostUrl="https://hro.cityu.edu.mo/en/2026/10/05/professor-associate-professor-assistant-professor-institute-of-urban-and-sustainable-development-2/",
        RequiredMaterial="Resume; ID; certificates; cover letter; 3 referees",
        SubmissionChannel="CityU-Career",
        ApplicationLink="https://career.cityu.edu.mo",
        Email="jobs@cityu.edu.mo",
        JobType="Academic",
    ),
    dict(
        Country="Macau",
        Location="Taipa",
        University="Macau University of Science and Technology",
        School="School of Business — Applied Economics",
        Deadline=None,
        JobPostTitle="Faculty (Economics among areas) — School of Business",
        JobPostUrl="https://www.must.edu.mo/en/careers",
        RequiredMaterial="CV; certificates; publications via MUST careers",
        SubmissionChannel="MUST-Career",
        ApplicationLink="https://www.must.edu.mo/en/careers",
        Email="hro@must.edu.mo",
        JobType="Academic",
    ),
    # --- Taiwan ---
    dict(
        Country="Taiwan",
        Location="Taipei",
        University="National Taiwan University",
        School="Department of Economics",
        Deadline="2027-01-31",
        JobPostTitle="Assistant Professor or above (up to 3)",
        JobPostUrl="https://econ.ntu.edu.tw/archives/18834",
        RequiredMaterial="CV; cover letter; JMP; papers; transcripts; letters",
        SubmissionChannel="Dept/JOE",
        ApplicationLink="https://econ.ntu.edu.tw/archives/18834",
        Email="kuanmingchen@ntu.edu.tw",
        JobType="Academic",
    ),
    dict(
        Country="Taiwan",
        Location="Taipei",
        University="National Chengchi University",
        School="Department of Economics",
        Deadline="2026-11-30",
        JobPostTitle="Open-rank tenure-track faculty (1)",
        JobPostUrl="https://econo.nccu.edu.tw/eng/PageDoc/Detail?fid=1052&id=42260",
        RequiredMaterial="Cover letter; CV; diploma; transcripts; JMP; syllabi; 3 letters",
        SubmissionChannel="Email",
        ApplicationLink="https://econo.nccu.edu.tw/eng/PageDoc/Detail?fid=1052&id=42260",
        Email="econ@nccu.edu.tw",
        JobType="Academic",
    ),
    dict(
        Country="Taiwan",
        Location="Hsinchu",
        University="National Tsing Hua University",
        School="Taipei School of Economics and Political Science (TSE)",
        Deadline="2026-11-15",
        JobPostTitle="Faculty — Macroeconomics / International Trade",
        JobPostUrl="http://www.tse.nthu.edu.tw/post/82-tse-job-opening",
        RequiredMaterial="AEA JOE application (Economics field)",
        SubmissionChannel="JOE",
        ApplicationLink="http://www.tse.nthu.edu.tw/post/82-tse-job-opening",
        Email="tse@my.nthu.edu.tw",
        JobType="Academic",
    ),
    dict(
        Country="Taiwan",
        Location="Tainan",
        University="National Cheng Kung University",
        School="Department of Economics",
        Deadline="2026-11-28",
        JobPostTitle="Assistant/Associate/Professor (any field)",
        JobPostUrl="https://ann.ncku.edu.tw/p/404-1013-305836.php?Lang=zh-tw",
        RequiredMaterial="Cover letter; CV; JMP; transcripts; 3 letters; diploma via JOE",
        SubmissionChannel="JOE",
        ApplicationLink="https://ann.ncku.edu.tw/p/404-1013-305836.php?Lang=zh-tw",
        Email="em56300@email.ncku.edu.tw",
        JobType="Academic",
    ),
    dict(
        Country="Taiwan",
        Location="Taipei",
        University="Soochow University",
        School="Department of Economics",
        Deadline="2026-10-30",
        JobPostTitle="Assistant Professor or above (2)",
        JobPostUrl="https://web-ch.scu.edu.tw/econ/opinion/16118/31786",
        RequiredMaterial="Diploma; transcripts; CV; thesis; 3 letters; syllabi",
        SubmissionChannel="Email/Mail",
        ApplicationLink="https://web-ch.scu.edu.tw/econ/opinion/16118/31786",
        Email="waccla@scu.edu.tw",
        JobType="Academic",
    ),
    dict(
        Country="Taiwan",
        Location="Kaohsiung",
        University="National Sun Yat-sen University",
        School="Department of Political Economy",
        Deadline="2026-12-31",
        JobPostTitle="Faculty (2) — Political Economy / Economics",
        JobPostUrl="https://www.nstc.gov.tw/folksonomy/detail/4b115db7-4cd9-4774-8ba4-31335f5ea07c?l=en",
        RequiredMaterial="Cover letter; CV; statements; pubs; degree certs; 3 letters",
        SubmissionChannel="Email",
        ApplicationLink="https://www.nstc.gov.tw/folksonomy/detail/4b115db7-4cd9-4774-8ba4-31335f5ea07c?l=en",
        Email="poleco@mail.nsysu.edu.tw",
        JobType="Academic",
    ),
]


def fetch(url: str) -> tuple[int | None, str]:
    if not url:
        return None, ""
    req = Request(url, headers={"User-Agent": UA, "Accept": "text/html,*/*"})
    try:
        with urlopen(req, timeout=45, context=CTX) as r:
            return r.status, r.read().decode("utf-8", "replace")
    except HTTPError as e:
        try:
            body = e.read().decode("utf-8", "replace")
        except Exception:
            body = ""
        return e.code, body
    except Exception as e:
        return None, str(e)


def slug(s: str) -> str:
    s = re.sub(r"[^\w\-]+", "_", s, flags=re.U)
    return s.strip("_")[:80] or "item"


def ensure_region_dirs() -> None:
    ensure_dirs()
    for d in (
        MACAU_DIR,
        TAIWAN_DIR,
        MACAU_DIR / "raw",
        TAIWAN_DIR / "raw",
        MACAU_DIR / "departments",
        TAIWAN_DIR / "departments",
    ):
        d.mkdir(parents=True, exist_ok=True)


def download_moe_csv() -> Path:
    dest = TAIWAN_DIR / f"moe_departments_college_{TODAY}.csv"
    print(f"Downloading MOE department CSV...")
    req = Request(MOE_CSV_URL, headers={"User-Agent": UA})
    with urlopen(req, timeout=120, context=CTX) as r:
        data = r.read()
    dest.write_bytes(data)
    # also keep stable name
    (TAIWAN_DIR / "moe_departments_college.csv").write_bytes(data)
    print(f"  saved {dest} ({len(data):,} bytes)")
    return dest


def parse_moe_economics(csv_path: Path) -> tuple[list[dict], list[dict]]:
    text = csv_path.read_text(encoding="utf-8-sig", errors="replace")
    lines = text.splitlines()
    # skip title row
    body = "\n".join(lines[1:])
    reader = csv.DictReader(StringIO(body))
    econ_units = []
    universities = {}
    for row in reader:
        uni = (row.get("學校名稱") or "").strip()
        if not uni:
            continue
        code = (row.get("學校代碼") or "").strip()
        if code not in universities:
            universities[code] = {
                "code": code,
                "name_zh": uni,
                "public_private": (row.get("公私立") or "").strip(),
                "system": (row.get("體制") or "").strip(),
            }
        dept = (row.get("系所名稱") or "").strip()
        if "經濟" in dept:
            econ_units.append(
                {
                    "school_code": code,
                    "university": uni,
                    "public_private": (row.get("公私立") or "").strip(),
                    "system": (row.get("體制") or "").strip(),
                    "college": (row.get("學院名稱") or "").strip(),
                    "dept_code": (row.get("系所代碼") or "").strip(),
                    "dept_name": dept,
                    "phone": (row.get("電話") or "").strip(),
                    "ext": (row.get("電話分機") or "").strip(),
                    "fax": (row.get("傳真") or "").strip(),
                    "email": (row.get("電子郵件") or "").strip(),
                    "url": (row.get("系所網址") or "").strip(),
                }
            )
    # de-dupe econ by university+dept+url
    seen, unique = set(), []
    for u in econ_units:
        key = (u["university"], u["dept_name"], u["url"])
        if key in seen:
            continue
        seen.add(key)
        unique.append(u)
    return list(universities.values()), unique


def save_fetch(folder: Path, name: str, url: str) -> dict:
    status, body = fetch(url)
    safe = slug(name)
    path = folder / f"{safe}.html"
    if body and len(body) > 50:
        path.write_text(body, encoding="utf-8", errors="replace")
    meta = {
        "name": name,
        "url": url,
        "status": status,
        "bytes": len(body or ""),
        "saved": str(path.name) if path.exists() else None,
        "fetched_at": datetime.now().isoformat(timespec="seconds"),
    }
    print(f"  [{status}] {name[:50]} ({meta['bytes']} bytes)")
    time.sleep(SLEEP)
    return meta


def scrape_macau() -> dict:
    print("=== Macau universities ===")
    catalog_path = MACAU_DIR / f"universities_{TODAY}.json"
    catalog_path.write_text(
        json.dumps(MACAU_UNIVERSITIES, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (MACAU_DIR / "universities.json").write_text(
        json.dumps(MACAU_UNIVERSITIES, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    fetch_log = []
    for uni in MACAU_UNIVERSITIES:
        if uni.get("career"):
            fetch_log.append(
                save_fetch(MACAU_DIR / "raw", f"career_{uni['name']}", uni["career"])
            )
        for unit in uni.get("econ_units") or []:
            if unit.get("url"):
                fetch_log.append(
                    save_fetch(
                        MACAU_DIR / "departments",
                        f"{uni['name']}_{unit['unit']}",
                        unit["url"],
                    )
                )
    # key job pages
    for job in KNOWN_JOBS:
        if job["Country"] != "Macau":
            continue
        fetch_log.append(
            save_fetch(MACAU_DIR / "raw", f"job_{job['JobPostTitle']}", job["JobPostUrl"])
        )

    log_path = MACAU_DIR / f"fetch_log_{TODAY}.json"
    log_path.write_text(json.dumps(fetch_log, ensure_ascii=False, indent=2), encoding="utf-8")
    return {"universities": len(MACAU_UNIVERSITIES), "fetches": len(fetch_log)}


def scrape_taiwan(universities: list[dict], econ_units: list[dict]) -> dict:
    print("=== Taiwan universities + economics units ===")
    (TAIWAN_DIR / f"universities_{TODAY}.json").write_text(
        json.dumps(universities, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (TAIWAN_DIR / "universities.json").write_text(
        json.dumps(universities, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (TAIWAN_DIR / f"economics_units_{TODAY}.json").write_text(
        json.dumps(econ_units, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (TAIWAN_DIR / "economics_units.json").write_text(
        json.dumps(econ_units, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    fetch_log = []
    for unit in econ_units:
        url = unit.get("url")
        if not url:
            continue
        name = f"{unit['university']}_{unit['dept_name']}"
        fetch_log.append(save_fetch(TAIWAN_DIR / "departments", name, url))

    for job in KNOWN_JOBS:
        if job["Country"] != "Taiwan":
            continue
        fetch_log.append(
            save_fetch(TAIWAN_DIR / "raw", f"job_{job['JobPostTitle']}", job["JobPostUrl"])
        )

    log_path = TAIWAN_DIR / f"fetch_log_{TODAY}.json"
    log_path.write_text(json.dumps(fetch_log, ensure_ascii=False, indent=2), encoding="utf-8")
    return {
        "universities": len(universities),
        "econ_units": len(econ_units),
        "fetches": len(fetch_log),
    }


def parse_deadline(raw):
    if raw is None or raw == "":
        return None
    if isinstance(raw, datetime):
        return raw
    s = str(raw).strip()
    try:
        return datetime.strptime(s[:10], "%Y-%m-%d")
    except ValueError:
        return None


def job_to_row(job: dict) -> dict:
    url = job.get("ApplicationLink") or job.get("JobPostUrl")
    title = job.get("JobPostTitle") or "Faculty position"
    return {
        "Country": job.get("Country"),
        "Location": job.get("Location"),
        "University": job.get("University"),
        "School": job.get("School"),
        "Deadline": parse_deadline(job.get("Deadline")),
        "Job Post": f"{title} | {url}" if url else title,
        "Required Material": job.get("RequiredMaterial"),
        "Submission Channel": job.get("SubmissionChannel"),
        "Application Link": url,
        "Email": job.get("Email"),
        "Submission Date": None,
        "Reference Status": None,
        "Interview Date": None,
        "Job Type": job.get("JobType") or "Academic",
        "_title": title,
    }


def write_xlsx(rows: list[dict], path: Path, sheet: str) -> None:
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


def compile_listings() -> None:
    macau_jobs = [job_to_row(j) for j in KNOWN_JOBS if j["Country"] == "Macau"]
    taiwan_jobs = [job_to_row(j) for j in KNOWN_JOBS if j["Country"] == "Taiwan"]
    all_jobs = macau_jobs + taiwan_jobs

    for region, rows in (("macau", macau_jobs), ("taiwan", taiwan_jobs)):
        folder = MACAU_DIR if region == "macau" else TAIWAN_DIR
        json_path = folder / f"listings_{TODAY}.json"
        serial = []
        for r in rows:
            item = {k: v for k, v in r.items() if not k.startswith("_")}
            if isinstance(item.get("Deadline"), datetime):
                item["Deadline"] = item["Deadline"].strftime("%Y-%m-%d")
            serial.append(item)
        json_path.write_text(json.dumps(serial, ensure_ascii=False, indent=2), encoding="utf-8")
        write_xlsx(rows, folder / f"listings_{TODAY}.xlsx", region.title())

    write_xlsx(all_jobs, CLEAN_ECON / f"macau_taiwan_listings_{TODAY}.xlsx", "Macau Taiwan")
    # also region clean files
    write_xlsx(macau_jobs, CLEAN_ECON / f"macau_listings_{TODAY}.xlsx", "Macau")
    write_xlsx(taiwan_jobs, CLEAN_ECON / f"taiwan_listings_{TODAY}.xlsx", "Taiwan")
    print(f"Compiled listings: Macau {len(macau_jobs)}, Taiwan {len(taiwan_jobs)}")


def main() -> None:
    ensure_region_dirs()
    csv_path = download_moe_csv()
    universities, econ_units = parse_moe_economics(csv_path)
    print(f"Taiwan universities in MOE file: {len(universities)}")
    print(f"Taiwan economics-related units: {len(econ_units)}")

    m = scrape_macau()
    t = scrape_taiwan(universities, econ_units)
    compile_listings()

    summary = {
        "date": TODAY,
        "macau": m,
        "taiwan": t,
        "open_jobs": {
            "macau": sum(1 for j in KNOWN_JOBS if j["Country"] == "Macau"),
            "taiwan": sum(1 for j in KNOWN_JOBS if j["Country"] == "Taiwan"),
        },
    }
    (DATA_ECON / f"macau_taiwan_summary_{TODAY}.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(summary, indent=2))
    print("Done.")


if __name__ == "__main__":
    main()
