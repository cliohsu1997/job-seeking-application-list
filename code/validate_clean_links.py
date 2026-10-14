"""
Visit every Application Link in clean/ one-by-one.
Keep only pages that look like real job-application postings.
Rewrite clean/ without the rejected rows.
"""
from __future__ import annotations

import json
import re
import time
from datetime import date, datetime
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlparse
from urllib.request import Request, urlopen

from openpyxl import load_workbook

from columns import COLUMNS

ROOT = Path(__file__).resolve().parents[1]
from paths import CLEAN_ECON, DATA_ECON, ensure_dirs

CLEAN = CLEAN_ECON
DATA = DATA_ECON
PROGRESS = DATA / "clean_validate_progress.json"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
SLEEP = 0.35

# Hosts / path patterns that are themselves application boards
TRUSTED_APPLY_PATTERNS = (
    r"aeaweb\.org/joe/listing\.php",
    r"econjobmarket\.org/positions/\d+",
    r"cicecon\.org\.cn/careers/detail\.html",
    r"apply\.interfolio\.com/\d+",
    r"\.myworkdayjobs\.com/.+/job/",
    r"boards\.greenhouse\.io/.+/jobs/\d+",
    r"job-boards\.greenhouse\.io/.+/jobs/\d+",
    r"jobs\.lever\.co/.+/",
    r"\.icims\.com/jobs/\d+",
    r"usajobs\.gov/job/",
    r"academicjobsonline\.org/ajo/jobs/\d+",
    r"applytojob\.com/apply/",
    r"apptrkr\.com/\d+",
    r"applicantpro\.com/jobs/\d+",
    r"pageuppeople\.com/.+/job/\d+",
    r"careers\.bowdoin\.edu/postings/\d+",
    r"cuny\.jobs/.+/job/",
    r"jobmarket\.unibocconi\.eu/\?id=\d+",
    r"facultypositions\.ahduni\.edu\.in",
    r"dayforcehcm\.com/.+/jobs/\d+",
    r"grnh\.se/",
)

POSITIVE_PHRASES = (
    "application deadline",
    "application requirements",
    "how to apply",
    "to apply",
    "apply now",
    "apply online",
    "submit your application",
    "submit an application",
    "cover letter",
    "curriculum vitae",
    "letters of reference",
    "letter of recommendation",
    "job market paper",
    "faculty position",
    "job opening",
    "job posting",
    "position announcement",
    "we invite applications",
    "applications are invited",
    "apply for this",
    "candidate portal",
    "requisition",
    "vacancy",
    "recruitment",
    "hiring",
    "assistant professor",
    "associate professor",
    "postdoctoral",
    "joe listings",
    "econ job market",
)

NEGATIVE_PHRASES = (
    "position has been filled",
    "job has been filled",
    "listing is inactive",
    "this listing is inactive",
    "no longer accepting applications",
    "no longer available",
    "this job is no longer",
    "page not found",
    "404 not found",
    "error 404",
)


def latest_clean_xlsx() -> Path:
    files = sorted(CLEAN.glob("all_listings_*.xlsx"), reverse=True)
    if not files:
        raise FileNotFoundError(
            "No clean/economics/all_listings_*.xlsx — run combine_clean.py first"
        )
    return files[0]


def fetch(url: str) -> tuple[int, str, str]:
    req = Request(url, headers={"User-Agent": UA, "Accept": "text/html,*/*"})
    try:
        with urlopen(req, timeout=30) as resp:
            body = resp.read(80000).decode("utf-8", "replace")
            return resp.status, resp.geturl(), body
    except HTTPError as e:
        try:
            body = e.read(20000).decode("utf-8", "replace")
        except Exception:
            body = ""
        return e.code, url, body
    except URLError as e:
        return 0, url, str(e.reason)
    except Exception as e:
        return 0, url, str(e)


def is_trusted_apply_url(url: str) -> bool:
    u = (url or "").lower()
    return any(re.search(p, u) for p in TRUSTED_APPLY_PATTERNS)


def is_generic_homepage(url: str) -> bool:
    """True for bare domain / school root with no job-like path."""
    try:
        p = urlparse(url)
    except Exception:
        return False
    path = (p.path or "/").rstrip("/")
    if path in ("", "/"):
        return True
    # shallow generic paths
    generic = {
        "/en",
        "/home",
        "/index",
        "/index.html",
        "/index.php",
        "/about",
        "/economics",
        "/econ",
        "/academic",
        "/colleges-schools-and-departments",
    }
    if path.lower() in generic:
        return True
    # usajobs root without /job/
    if "usajobs.gov" in (p.netloc or "").lower() and "/job/" not in path.lower():
        return True
    return False


def page_is_application(url: str, status: int, body: str, final_url: str) -> tuple[bool, str]:
    if not status or status >= 400:
        return False, f"http_{status or 'error'}"

    text = re.sub(r"<script[\s\S]*?</script>", " ", body or "", flags=re.I)
    text = re.sub(r"<style[\s\S]*?</style>", " ", text, flags=re.I)
    text = re.sub(r"<[^>]+>", " ", text)
    text = html_unescape_ws(text).lower()
    head = text[:8000]

    for phrase in NEGATIVE_PHRASES:
        if phrase in head:
            return False, f"negative:{phrase}"

    # Trusted job-board URLs that loaded OK
    if is_trusted_apply_url(final_url) or is_trusted_apply_url(url):
        # still require some application-ish content when body is HTML
        if any(p in head for p in POSITIVE_PHRASES) or "application" in head or "apply" in head:
            return True, "trusted_board"
        # some ATS pages are JS-heavy; trust URL pattern if status 200
        if status == 200 and len(body) > 500:
            return True, "trusted_board_js"

    pos_hits = [p for p in POSITIVE_PHRASES if p in head]
    if len(pos_hits) >= 2:
        return True, f"content:{pos_hits[0]}"
    if len(pos_hits) == 1 and ("apply" in head or "application" in head or "job" in head):
        return True, f"content:{pos_hits[0]}"

    if is_generic_homepage(final_url) or is_generic_homepage(url):
        return False, "generic_homepage"

    # news / research article pages
    if any(
        x in head
        for x in ("doi.org", "abstract", "references", "cite this article", "nature portfolio")
    ) and "application deadline" not in head:
        return False, "not_job_page"

    if "apply" in head and ("position" in head or "faculty" in head or "job" in head):
        return True, "content:apply+job"

    return False, "no_application_signal"


def html_unescape_ws(s: str) -> str:
    import html

    s = html.unescape(s or "")
    return re.sub(r"\s+", " ", s).strip()


def load_progress() -> dict:
    if PROGRESS.exists():
        return json.loads(PROGRESS.read_text(encoding="utf-8"))
    return {"checked": {}}


def save_progress(prog: dict) -> None:
    DATA.mkdir(exist_ok=True)
    PROGRESS.write_text(json.dumps(prog, ensure_ascii=False, indent=2), encoding="utf-8")


def read_clean_rows(path: Path) -> list[dict]:
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    rows = []
    for r in range(2, ws.max_row + 1):
        item = {}
        for c, h in enumerate(headers, 1):
            cell = ws.cell(r, c)
            if h in ("Job Post", "Application Link") and cell.hyperlink and cell.hyperlink.target:
                item[h] = cell.hyperlink.target
            else:
                item[h] = cell.value
        if item.get("Application Link") or item.get("Job Post"):
            rows.append(item)
    wb.close()
    return rows


def write_clean(rows: list[dict], today: str) -> Path:
    from combine_clean import write_xlsx, clear_old_econ_outputs

    ensure_dirs()
    clear_old_econ_outputs()
    xlsx_path = CLEAN / f"all_listings_{today}.xlsx"
    write_xlsx(rows, xlsx_path)
    for channel in ("AEA", "EJM", "CICE"):
        subset = [r for r in rows if r.get("Submission Channel") == channel]
        write_xlsx(subset, CLEAN / f"{channel.lower()}_listings_{today}.xlsx")
    return xlsx_path


def _safe_print(msg: str) -> None:
    try:
        print(msg)
    except UnicodeEncodeError:
        print(msg.encode("ascii", "replace").decode("ascii"))


def main() -> None:
    path = latest_clean_xlsx()
    rows = read_clean_rows(path)
    prog = load_progress()
    checked: dict = prog.setdefault("checked", {})

    _safe_print(f"Validating {len(rows)} Application Links from {path.name} one-by-one...")
    kept: list[dict] = []
    rejected: list[dict] = []

    for i, row in enumerate(rows, 1):
        url = str(row.get("Application Link") or "").strip()
        key = url
        uni = row.get("University")
        ch = row.get("Submission Channel")
        _safe_print(f"[{i}/{len(rows)}] {ch} | {uni}")
        _safe_print(f"     {url}")

        if key in checked:
            cached = checked[key]
            ok = cached.get("ok")
            reason = cached.get("reason")
            status = cached.get("status")
            final = cached.get("final_url")
            _safe_print(f"     cached -> {'KEEP' if ok else 'REMOVE'} ({reason})")
        else:
            status, final, body = fetch(url)
            ok, reason = page_is_application(url, status, body, final)
            checked[key] = {
                "ok": ok,
                "reason": reason,
                "status": status,
                "final_url": final,
                "university": uni,
                "channel": ch,
            }
            _safe_print(f"     {'KEEP' if ok else 'REMOVE'} status={status} ({reason})")
            if i % 10 == 0:
                save_progress(prog)
            time.sleep(SLEEP)

        # Keep the row always; only clear Application Link when page is not an application
        if checked[key]["ok"]:
            kept.append(row)
        else:
            cleared = dict(row)
            cleared["Application Link"] = None
            kept.append(cleared)
            rejected.append(
                {
                    "reason": checked[key].get("reason"),
                    "status": checked[key].get("status"),
                    "Submission Channel": ch,
                    "University": uni,
                    "Application Link": url,
                    "final_url": checked[key].get("final_url"),
                    "Job Post": row.get("Job Post"),
                }
            )

    save_progress(prog)
    today = date.today().isoformat()
    out = write_clean(kept, today)
    _safe_print(
        f"Rows kept: {len(kept)} | Application Links cleared: {len(rejected)}"
    )
    _safe_print(f"Clean rewritten: {out}")


if __name__ == "__main__":
    main()
