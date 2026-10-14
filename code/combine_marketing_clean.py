"""
Clean marketing listings: drop outdated deadlines, write clean/marketing/.

Uses Marketing Listings sheet in application list.xlsx if present,
else latest data/marketing/marketing_structured_*.xlsx.
"""
from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from columns import COLUMNS
from paths import APP_XLSX, CLEAN_MKT, DATA_MKT, ensure_dirs


def extract_url(text) -> str | None:
    if text is None:
        return None
    s = str(text).strip()
    if not s:
        return None
    urls = re.findall(r"https?://[^\s|]+", s)
    if s.startswith("http"):
        return urls[-1].rstrip(").,;\"'") if urls else s
    return urls[-1].rstrip(").,;\"'") if urls else None


def deadline_date(value) -> date | None:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if type(value) is date:
        return value
    s = str(value).strip()
    for fmt, n in (("%Y-%m-%d %H:%M:%S", 19), ("%Y-%m-%d", 10), ("%m/%d/%Y", 10)):
        try:
            return datetime.strptime(s[:n], fmt).date()
        except ValueError:
            continue
    return None


def is_outdated(row: dict, today: date) -> bool:
    d = deadline_date(row.get("Deadline"))
    if d is None:
        return False
    return d < today


def cell_value_or_link(cell):
    if cell.hyperlink and cell.hyperlink.target:
        return cell.hyperlink.target
    return cell.value


def read_rows_from_wb(wb, sheet_name: str) -> list[dict]:
    if sheet_name not in wb.sheetnames:
        return []
    ws = wb[sheet_name]
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    rows = []
    for r in range(2, ws.max_row + 1):
        item = {}
        for c, h in enumerate(headers, 1):
            if h in ("Job Post", "Application Link"):
                item[h] = cell_value_or_link(ws.cell(r, c))
            else:
                item[h] = ws.cell(r, c).value
        if not any(item.get(h) for h in ("University", "Job Post", "Application Link")):
            continue
        link = extract_url(item.get("Application Link")) or extract_url(item.get("Job Post"))
        item["Application Link"] = link
        rows.append(item)
    return rows


def write_xlsx(rows: list[dict], path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Marketing Listings"
    for c, name in enumerate(COLUMNS, 1):
        ws.cell(1, c, name).font = Font(bold=True)
    for r_i, row in enumerate(rows, 2):
        for c, name in enumerate(COLUMNS, 1):
            ws.cell(r_i, c, row.get(name))
    for c in range(1, len(COLUMNS) + 1):
        ws.column_dimensions[get_column_letter(c)].width = 16
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def clear_old() -> None:
    if not CLEAN_MKT.exists():
        return
    for p in CLEAN_MKT.glob("*"):
        if p.is_file():
            p.unlink()


def main() -> None:
    ensure_dirs()
    clear_old()

    rows: list[dict] = []
    if APP_XLSX.exists():
        wb = load_workbook(APP_XLSX, data_only=False)
        rows = read_rows_from_wb(wb, "Marketing Listings")
        wb.close()

    if not rows:
        files = sorted(DATA_MKT.glob("marketing_structured_*.xlsx"), reverse=True)
        if not files:
            raise FileNotFoundError("No marketing structured file — run compile_marketing.py")
        wb = load_workbook(files[0], data_only=False)
        rows = read_rows_from_wb(wb, wb.sheetnames[0])
        wb.close()

    # MJM sheet is the living open list — keep all rows with a direct link.
    # (Unlike economics boards, past "deadline" cells often mean review start.)
    today_d = date.today()
    with_link = [r for r in rows if extract_url(r.get("Application Link"))]
    past_deadline = sum(1 for r in with_link if is_outdated(r, today_d))
    active = with_link

    today = today_d.isoformat()
    xlsx_path = CLEAN_MKT / f"marketing_listings_{today}.xlsx"
    write_xlsx(active, xlsx_path)

    print(f"Marketing rows: {len(rows)}")
    print(f"Skipped without link: {len(rows) - len(with_link)}")
    print(f"Note: {past_deadline} have deadline < {today} (kept; MJM living list)")
    print(f"Clean listings: {len(active)}")
    print(f"Wrote {xlsx_path}")


if __name__ == "__main__":
    main()
