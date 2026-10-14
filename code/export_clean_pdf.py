"""Export clean economics + marketing listings to readable PDFs in clean/ (root)."""
from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import A3, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from columns import COLUMNS
from paths import CLEAN, CLEAN_ECON, CLEAN_MKT, ensure_dirs

# Windows CJK fonts so Chinese job titles render in coastal PDFs.
_CJK_FONT_CANDIDATES = (
    Path(r"C:\Windows\Fonts\msyh.ttc"),
    Path(r"C:\Windows\Fonts\msyh.ttf"),
    Path(r"C:\Windows\Fonts\simsun.ttc"),
    Path(r"C:\Windows\Fonts\simhei.ttf"),
)
_CJK_FONT_NAME: str | None = None


def ensure_cjk_font() -> str | None:
    global _CJK_FONT_NAME
    if _CJK_FONT_NAME:
        return _CJK_FONT_NAME
    for path in _CJK_FONT_CANDIDATES:
        if not path.exists():
            continue
        try:
            name = "AppListCJK"
            # .ttc needs subfontIndex; try 0 first
            if path.suffix.lower() == ".ttc":
                pdfmetrics.registerFont(TTFont(name, str(path), subfontIndex=0))
            else:
                pdfmetrics.registerFont(TTFont(name, str(path)))
            _CJK_FONT_NAME = name
            return name
        except Exception:
            continue
    return None

WEIGHTS = {
    "Country": 6,
    "Location": 8,
    "University": 12,
    "School": 10,
    "Deadline": 6,
    "Job Post": 10,
    "Required Material": 10,
    "Submission Channel": 5,
    "Application Link": 12,
    "Email": 8,
    "Submission Date": 5,
    "Reference Status": 5,
    "Interview Date": 5,
    "Job Type": 5,
}

_URL_RE = re.compile(r"https?://[^\s|]+")


def latest(folder: Path, pattern: str) -> Path | None:
    if not folder.exists():
        return None
    files = sorted(folder.glob(pattern), reverse=True)
    return files[0] if files else None


def xml_escape(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def extract_url(text) -> str | None:
    if text is None:
        return None
    s = str(text).strip()
    if not s:
        return None
    m = _URL_RE.findall(s)
    if m:
        return m[-1].rstrip(").,;\"'")
    if s.startswith("http://") or s.startswith("https://"):
        return s
    return None


def fmt(value, col: str, max_len: int = 90) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    text = str(value).replace("\n", " ").strip()
    if len(text) > max_len:
        text = text[: max_len - 1] + "…"
    return xml_escape(text)


def fmt_link_cell(value, col: str) -> str:
    """Render Job Post / Application Link as a clickable PDF link with the full URL shown."""
    if value is None or value == "":
        return ""
    raw = str(value).replace("\n", " ").strip()
    url = extract_url(raw)
    if not url:
        return fmt(raw, col)

    title = None
    if " | " in raw:
        left, right = raw.rsplit(" | ", 1)
        if extract_url(right):
            title = left.strip() or None

    # Job Post: prefer human title when present; otherwise show the full URL.
    # Application Link: always show the full URL (no truncation / no substitute label).
    if col == "Job Post" and title:
        label = title
    else:
        label = url

    href = xml_escape(url)
    return f'<link href="{href}">{xml_escape(label)}</link>'

def load_rows(path: Path) -> list[dict]:
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    rows = []
    for r in range(2, ws.max_row + 1):
        item = {}
        for c, h in enumerate(headers, 1):
            cell = ws.cell(r, c)
            if h in ("Job Post", "Application Link"):
                # Prefer cell text (may be "Title | url"); fall back to hyperlink target.
                val = cell.value
                href = cell.hyperlink.target if cell.hyperlink else None
                if val and href and extract_url(str(val)) is None:
                    item[h] = f"{val} | {href}"
                elif href and not val:
                    item[h] = href
                else:
                    item[h] = val or href
            else:
                item[h] = cell.value
        rows.append(item)
    wb.close()
    return rows


def export_one(src: Path, pdf_path: Path, title: str, use_cjk: bool = False) -> Path:
    rows = load_rows(src)
    today = date.today().isoformat()
    page = landscape(A3)
    margin = 6 * mm
    usable = page[0] - 2 * margin
    weights = [WEIGHTS.get(c, 6) for c in COLUMNS]
    total_w = sum(weights)
    col_widths = [usable * w / total_w for w in weights]

    font = ensure_cjk_font() if use_cjk else None
    base = getSampleStyleSheet()["Normal"]
    cell_kw = {"fontName": font} if font else {}
    cell_style = ParagraphStyle(
        "Cell", parent=base, fontSize=5.5, leading=7, **cell_kw
    )
    # CJK wrap lets long URLs break mid-string without inserting zero-width chars
    # (Helvetica has no ZWSP glyph, which previously showed as boxes).
    link_style = ParagraphStyle(
        "LinkCell",
        parent=cell_style,
        fontSize=5.5,
        leading=7,
        wordWrap="CJK",
        **cell_kw,
    )
    header_style = ParagraphStyle(
        "Hdr",
        parent=base,
        fontSize=6.5,
        leading=8,
        textColor=colors.white,
        **cell_kw,
    )
    title_style = ParagraphStyle(
        "Title", parent=base, fontSize=11, leading=14, spaceAfter=2, **cell_kw
    )
    sub_style = ParagraphStyle(
        "Sub", parent=base, fontSize=7.5, leading=9, **cell_kw
    )

    data = [[Paragraph(f"<b>{h}</b>", header_style) for h in COLUMNS]]
    for row in rows:
        cells = []
        for c in COLUMNS:
            if c in ("Job Post", "Application Link"):
                cells.append(Paragraph(fmt_link_cell(row.get(c), c), link_style))
            else:
                cells.append(Paragraph(fmt(row.get(c), c), cell_style))
        data.append(cells)
    table = Table(data, colWidths=col_widths, repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f4e79")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTSIZE", (0, 0), (-1, -1), 5.5),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f2f5f9")]),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#b0b0b0")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 2),
                ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                ("TOPPADDING", (0, 0), (-1, -1), 1.5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
            ]
        )
    )

    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    doc = SimpleDocTemplate(
        str(pdf_path),
        pagesize=page,
        leftMargin=margin,
        rightMargin=margin,
        topMargin=8 * mm,
        bottomMargin=8 * mm,
        title=title,
    )
    story = [
        Paragraph(title, title_style),
        Paragraph(
            f"Source: {src.name} · {len(rows)} listings · Application columns · {today}",
            sub_style,
        ),
        Spacer(1, 3 * mm),
        table,
    ]
    doc.build(story)
    return pdf_path


def main() -> None:
    ensure_dirs()
    today = date.today().isoformat()
    written = []

    econ = latest(CLEAN_ECON, "all_listings_*.xlsx")
    if econ:
        out = export_one(
            econ,
            CLEAN / f"economics_listings_{today}.pdf",
            "Economics Job Listings — Application List Format",
        )
        written.append(out)
        print(f"Wrote {out}")
    else:
        print("No clean/economics/all_listings_*.xlsx — skip econ PDF")

    mkt = latest(CLEAN_MKT, "marketing_listings_*.xlsx")
    if mkt:
        out = export_one(
            mkt,
            CLEAN / f"marketing_listings_{today}.pdf",
            "Marketing Job Listings — Application List Format",
        )
        written.append(out)
        print(f"Wrote {out}")
    else:
        print("No clean/marketing/marketing_listings_*.xlsx — skip marketing PDF")

    coastal = latest(CLEAN_ECON, "china_coastal_listings_mined_*.xlsx")
    if not coastal:
        coastal = latest(CLEAN_ECON, "china_coastal_listings_*.xlsx")
    if coastal:
        out = export_one(
            coastal,
            CLEAN / f"china_coastal_listings_{today}.pdf",
            "China Coastal (GD/FJ/ZJ/SZ) — Faculty + Postdoc",
            use_cjk=True,
        )
        written.append(out)
        print(f"Wrote {out}")
    else:
        print("No clean/economics/china_coastal_listings*.xlsx — skip coastal PDF")

    if not written:
        raise FileNotFoundError("No clean listings to export")


if __name__ == "__main__":
    main()
