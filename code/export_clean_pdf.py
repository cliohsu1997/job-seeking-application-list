# -*- coding: utf-8 -*-
"""Export final PDFs/XLSX into clean/ root only.

Allowed in clean/ (outside subfolders):
  - marketing_listings_*.pdf
  - aea_listings_*.pdf (+ .xlsx)
  - provincial_listings_*.pdf (+ .xlsx)
  - asia_pacific_listings_*.pdf (+ .xlsx)  # combined two-section

Everything else moves to clean/economics/province/.
"""
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
from paths import CLEAN, CLEAN_ECON, CLEAN_ECON_PROVINCE, CLEAN_MKT, ensure_dirs

_CJK_FONT_CANDIDATES = (
    Path(r"C:\Windows\Fonts\simhei.ttf"),
    Path(r"C:\Windows\Fonts\simkai.ttf"),
    Path(r"C:\Windows\Fonts\simfang.ttf"),
    Path(r"C:\Windows\Fonts\msyh.ttf"),
    Path(r"C:\Windows\Fonts\msyh.ttc"),
    Path(r"C:\Windows\Fonts\simsun.ttc"),
)
_CJK_FONT_NAME: str | None = None
_CJK_PROBE = "\u5317\u4eac\u7ecf\u6d4e"


def ensure_cjk_font() -> str | None:
    global _CJK_FONT_NAME
    if _CJK_FONT_NAME:
        return _CJK_FONT_NAME
    for path in _CJK_FONT_CANDIDATES:
        if not path.exists():
            continue
        try:
            name = "AppListCJK"
            if path.suffix.lower() == ".ttc":
                pdfmetrics.registerFont(TTFont(name, str(path), subfontIndex=0))
            else:
                pdfmetrics.registerFont(TTFont(name, str(path)))
            if pdfmetrics.stringWidth(_CJK_PROBE, name, 10) <= 0:
                continue
            _CJK_FONT_NAME = name
            return name
        except Exception:
            continue
    return None


WEIGHTS = {
    "Country": 5,
    "Location": 6,
    "University": 10,
    "College Ranking": 4,
    "School": 8,
    "Deadline": 5,
    "Published Year": 4,
    "Job Post": 10,
    "Required Material": 8,
    "Submission Channel": 5,
    "Application Link": 11,
    "Email": 7,
    "Job Type": 4,
}

_URL_RE = re.compile(r"https?://[^\s|]+")

FINAL_STEMS = (
    "marketing_listings_",
    "aea_listings_",
    "provincial_listings_",
    "asia_pacific_listings_",
)


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
        text = text[: max_len - 1] + "..."
    return xml_escape(text)


def fmt_link_cell(value, col: str) -> str:
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
    if col == "Job Post" and title:
        label = title
    else:
        label = url
    href = xml_escape(url)
    return '<link href="%s">%s</link>' % (href, xml_escape(label))


def _sheet_rows(ws) -> list[dict]:
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    rows = []
    for r in range(2, ws.max_row + 1):
        item = {}
        empty = True
        for c, h in enumerate(headers, 1):
            if not h:
                continue
            cell = ws.cell(r, c)
            if h in ("Job Post", "Application Link"):
                val = cell.value
                href = cell.hyperlink.target if cell.hyperlink else None
                if val and href and extract_url(str(val)) is None:
                    item[h] = "%s | %s" % (val, href)
                elif href and not val:
                    item[h] = href
                else:
                    item[h] = val or href
            else:
                item[h] = cell.value
            if item.get(h) not in (None, ""):
                empty = False
        if not empty:
            rows.append(item)
    return rows


def load_rows(path: Path) -> list[dict]:
    wb = load_workbook(path, data_only=True)
    rows = _sheet_rows(wb.active)
    wb.close()
    return rows


def load_sheets(path: Path) -> list[tuple[str, list[dict]]]:
    wb = load_workbook(path, data_only=True)
    out = [(name, _sheet_rows(wb[name])) for name in wb.sheetnames]
    wb.close()
    return out


def _build_table(rows, col_widths, cell_style, link_style, header_style):
    data = [[Paragraph("<b>%s</b>" % h, header_style) for h in COLUMNS]]
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
                (
                    "ROWBACKGROUNDS",
                    (0, 1),
                    (-1, -1),
                    [colors.white, colors.HexColor("#f2f5f9")],
                ),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#b0b0b0")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 2),
                ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                ("TOPPADDING", (0, 0), (-1, -1), 1.5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
            ]
        )
    )
    return table


def export_one(src: Path, pdf_path: Path, title: str, use_cjk: bool = False) -> Path:
    sheets = load_sheets(src)
    if len(sheets) > 1:
        return export_sections(
            sheets,
            pdf_path,
            title=title,
            use_cjk=use_cjk,
            source_name=src.name,
        )

    rows = sheets[0][1] if sheets else []
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
    cell_style = ParagraphStyle("Cell", parent=base, fontSize=5.5, leading=7, **cell_kw)
    link_style = ParagraphStyle(
        "LinkCell", parent=cell_style, fontSize=5.5, leading=7, wordWrap="CJK", **cell_kw
    )
    header_style = ParagraphStyle(
        "Hdr", parent=base, fontSize=6.5, leading=8, textColor=colors.white, **cell_kw
    )
    title_style = ParagraphStyle(
        "Title", parent=base, fontSize=11, leading=14, spaceAfter=2, **cell_kw
    )
    sub_style = ParagraphStyle("Sub", parent=base, fontSize=7.5, leading=9, **cell_kw)

    table = _build_table(rows, col_widths, cell_style, link_style, header_style)
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
    doc.build(
        [
            Paragraph(title, title_style),
            Paragraph(
                "Source: %s · %d listings · Application columns · %s"
                % (src.name, len(rows), today),
                sub_style,
            ),
            Spacer(1, 3 * mm),
            table,
        ]
    )
    return pdf_path


def export_sections(
    sections: list[tuple[str, list[dict]]],
    pdf_path: Path,
    *,
    title: str,
    use_cjk: bool = False,
    source_name: str = "",
    gap_mm: float = 18.0,
) -> Path:
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
        "CellSec", parent=base, fontSize=5.5, leading=7, **cell_kw
    )
    link_style = ParagraphStyle(
        "LinkCellSec",
        parent=cell_style,
        fontSize=5.5,
        leading=7,
        wordWrap="CJK",
        **cell_kw,
    )
    header_style = ParagraphStyle(
        "HdrSec",
        parent=base,
        fontSize=6.5,
        leading=8,
        textColor=colors.white,
        **cell_kw,
    )
    title_style = ParagraphStyle(
        "TitleSec", parent=base, fontSize=11, leading=14, spaceAfter=2, **cell_kw
    )
    section_style = ParagraphStyle(
        "SectionSec",
        parent=base,
        fontSize=10,
        leading=13,
        spaceBefore=2,
        spaceAfter=4,
        textColor=colors.HexColor("#1f4e79"),
        **cell_kw,
    )
    sub_style = ParagraphStyle("SubSec", parent=base, fontSize=7.5, leading=9, **cell_kw)

    total = sum(len(rows) for _, rows in sections)
    story = [
        Paragraph(title, title_style),
        Paragraph(
            "Source: %s · %d listings · %d sections · same columns · %s"
            % (source_name or pdf_path.name, total, len(sections), today),
            sub_style,
        ),
        Spacer(1, 3 * mm),
    ]
    for i, (sec_name, rows) in enumerate(sections):
        if i > 0:
            story.append(Spacer(1, gap_mm * mm))
        story.append(
            Paragraph("<b>%s</b>  (%d listings)" % (sec_name, len(rows)), section_style)
        )
        story.append(Spacer(1, 2 * mm))
        story.append(
            _build_table(rows, col_widths, cell_style, link_style, header_style)
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
    doc.build(story)
    return pdf_path


def prune_clean_root(today: str | None = None) -> list[Path]:
    """Keep only marketing / aea / provincial / asia_pacific finals in clean/.

    Also archive older dated copies of the same stem (keep newest per extension).
    """
    ensure_dirs()
    CLEAN_ECON_PROVINCE.mkdir(parents=True, exist_ok=True)
    moved = []
    for path in CLEAN.iterdir():
        if not path.is_file():
            continue
        if any(path.name.startswith(p) for p in FINAL_STEMS):
            continue
        dest = CLEAN_ECON_PROVINCE / path.name
        if dest.exists():
            dest.unlink()
        path.replace(dest)
        moved.append(dest)

    for stem in FINAL_STEMS:
        by_ext: dict[str, list[Path]] = {}
        for path in CLEAN.glob(stem + "*"):
            if path.is_file():
                by_ext.setdefault(path.suffix.lower(), []).append(path)
        for group in by_ext.values():
            group.sort(key=lambda p: p.name, reverse=True)
            for old in group[1:]:
                dest = CLEAN_ECON_PROVINCE / old.name
                if dest.exists():
                    dest.unlink()
                old.replace(dest)
                moved.append(dest)
    return moved


def main(*, prune: bool = True) -> None:
    ensure_dirs()
    today = date.today().isoformat()
    written = []

    mkt = latest(CLEAN_MKT, "marketing_listings_*.xlsx")
    if mkt:
        out = export_one(
            mkt,
            CLEAN / ("marketing_listings_%s.pdf" % today),
            "Marketing Job Listings",
        )
        written.append(out)
        print("Wrote", out)

    for stem, title in (
        ("aea_listings", "AEA/EJM/CICE Asia-Pacific"),
        ("provincial_listings", "Provincial China + Macau/Taiwan"),
        ("asia_pacific_listings", "Asia-Pacific combined"),
    ):
        src = latest(CLEAN, "%s_*.xlsx" % stem)
        if not src:
            print("No clean/%s_*.xlsx — skip" % stem)
            continue
        out = export_one(
            src,
            CLEAN / ("%s_%s.pdf" % (stem, today)),
            title,
            use_cjk=True,
        )
        written.append(out)
        print("Wrote", out)

    if prune:
        for path in prune_clean_root(today):
            print("Moved non-final -> economics/province/%s" % path.name)

    if not written:
        raise FileNotFoundError("No clean listings to export")


if __name__ == "__main__":
    main()
