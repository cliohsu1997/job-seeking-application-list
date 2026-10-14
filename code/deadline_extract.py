# -*- coding: utf-8 -*-
"""Shared deadline parsing from listing-page text."""
from __future__ import annotations

import re
from datetime import date, datetime

from freshness_rules import ROLLING
from validate_job_links import DEADLINE_RE, ROLLING_OK

_ISO = re.compile(r"(20\d{2})[-/.](\d{1,2})[-/.](\d{1,2})")
_US = re.compile(r"\b(\d{1,2})/(\d{1,2})/(20\d{2})\b")
_CN = re.compile(
    r"(20\d{2})\s*[\u5e74/-]\s*(\d{1,2})\s*[\u6708/-]\s*(\d{1,2})\s*\u65e5?"
)
_EN_MONTH = {
    "jan": 1,
    "january": 1,
    "feb": 2,
    "february": 2,
    "mar": 3,
    "march": 3,
    "apr": 4,
    "april": 4,
    "may": 5,
    "jun": 6,
    "june": 6,
    "jul": 7,
    "july": 7,
    "aug": 8,
    "august": 8,
    "sep": 9,
    "sept": 9,
    "september": 9,
    "oct": 10,
    "october": 10,
    "nov": 11,
    "november": 11,
    "dec": 12,
    "december": 12,
}
_MONTH = (
    r"(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|"
    r"Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|"
    r"Nov(?:ember)?|Dec(?:ember)?)"
)
# Month DD[, ]YYYY  (allows Nov. / November 28 th / December 7th?2025)
_EN_MDY = re.compile(
    rf"{_MONTH}\.?\s+(\d{{1,2}})\s*(?:st|nd|rd|th)?"
    rf"\s*[,?]?\s*(20\d{{2}})",
    re.I,
)
# DD Month YYYY  (Closing date: 30 Nov 2024)
_EN_DMY = re.compile(
    rf"(\d{{1,2}})\s*(?:st|nd|rd|th)?\s+{_MONTH}\.?\s*(20\d{{2}})",
    re.I,
)
# Keyword window then a date (CICE: "Deadline Date: Nov. 20, 2019")
_EN_KW = re.compile(
    r"(?:"
    r"deadline(?:\s*date)?|"
    r"closing\s+date|"
    r"due|"
    r"apply\s+by|"
    r"applications?\s+(?:due|close|must\s+be\s+received)|"
    r"must\s+be\s+received\s+by|"
    r"received\s+by|"
    r"no\s+later\s+than|"
    r"full\s+consideration|"
    r"submit(?:ted)?\s+before|"
    r"before\s+the|"
    # "submit ... letters by November 28, 2025"
    r"(?:submit|submitted|send|sent).{0,160}?\b(?:by|before)\b|"
    r"(?:letters|materials|applications?)\s+(?:by|before)\b|"
    r"\b(?:by|before)\s+(?=" + _MONTH + r")"
    r")"
    r".{0,120}",
    re.I,
)
_NEAR_KW = re.compile(
    r"("
    r"\u622a\u6b62\u65e5\u671f|\u62a5\u540d\u622a\u6b62|\u622a\u6b62\u65f6\u95f4|"
    r"\u7533\u8bf7\u622a\u6b62|\u6750\u6599\u63d0\u4ea4|\u8bf7\u4e8e|"
    r"\u62a5\u540d\u65f6\u95f4|\u7533\u8bf7\u65f6\u95f4|"
    r"deadline(?:\s*date)?|closing\s+date|apply\s+by|"
    r"received\s+by|no\s+later\s+than|full\s+consideration"
    r").{0,120}",
    re.I,
)


def _ymd(y, m, d) -> str | None:
    try:
        y, m, d = int(y), int(m), int(d or 1)
        if not (2000 <= y <= 2039 and 1 <= m <= 12 and 1 <= d <= 31):
            return None
        return date(y, m, d).isoformat()
    except (TypeError, ValueError):
        return None


def _month_num(name: str | None) -> int | None:
    if not name:
        return None
    return _EN_MONTH.get(name.lower().rstrip("."))


def _collect_en_dates(window: str) -> list[str]:
    found: list[str] = []
    for m in _EN_MDY.finditer(window):
        mo = _month_num(m.group(1))
        s = _ymd(m.group(3), mo, m.group(2)) if mo else None
        if s:
            found.append(s)
    for m in _EN_DMY.finditer(window):
        mo = _month_num(m.group(2))
        s = _ymd(m.group(3), mo, m.group(1)) if mo else None
        if s:
            found.append(s)
    for m in _US.finditer(window):
        s = _ymd(m.group(3), m.group(1), m.group(2))
        if s:
            found.append(s)
    for cre in (_ISO, _CN):
        for dm in cre.finditer(window):
            s = _ymd(dm.group(1), dm.group(2), dm.group(3))
            if s:
                found.append(s)
    return found


def extract_deadline(text: str) -> str | None:
    """Return YYYY-MM-DD, 'Rolling', or None if unknown."""
    if not text:
        return None

    # Normalize HTML entities / odd spaces so "Nov.\n20" and "28 th" match
    text = (
        text.replace("\xa0", " ")
        .replace("&nbsp;", " ")
        .replace("&#32;", " ")
        .replace("\u3000", " ")
    )
    text = re.sub(r"\s+", " ", text)

    found: list[str] = []
    for m in DEADLINE_RE.finditer(text):
        s = _ymd(m.group(1), m.group(2), m.group(3) or 1)
        if s:
            found.append(s)

    for m in _EN_KW.finditer(text):
        found.extend(_collect_en_dates(m.group(0)))

    for m in _NEAR_KW.finditer(text):
        found.extend(_collect_en_dates(m.group(0)))

    # Prefer keyword-linked dates; if none, still accept clear MDY near apply phrases
    if not found:
        by_pat = (
            r"(?:applications?\s+(?:must\s+be\s+)?received\s+by|"
            r"submit(?:ted)?\s+(?:by|before)|"
            r"to\s+be\s+guaranteed\s+full\s+consideration|"
            r"fully\s+considered|"
            r"\b(?:by|before)\s+" + _MONTH + r")"
            r".{0,40}"
        )
        for m in re.finditer(by_pat, text, flags=re.I):
            found.extend(_collect_en_dates(m.group(0)))

    if found:
        # Concrete dates beat "rolling" language elsewhere on the page
        today = date.today().isoformat()
        future = sorted(d for d in found if d >= today)
        return future[0] if future else sorted(found)[-1]

    if ROLLING.search(text) or ROLLING_OK.search(text):
        return "Rolling"
    return None


def html_to_text(html: str) -> str:
    """Strip tags/scripts for deadline / job-type scanning."""
    if not html:
        return ""
    text = re.sub(r"(?is)<script[^>]*>.*?</script>", " ", html)
    text = re.sub(r"(?is)<style[^>]*>.*?</style>", " ", text)
    text = re.sub(r"<[^>]+>", " ", text)
    text = (
        text.replace("&nbsp;", " ")
        .replace("&#32;", " ")
        .replace("&amp;", "&")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", '"')
    )
    text = re.sub(r"\s+", " ", text).strip()
    return text


def has_deadline(job: dict) -> bool:
    d = job.get("Deadline")
    if d in (None, "", "None"):
        return False
    if isinstance(d, datetime):
        return True
    s = str(d).strip()
    if re.match(r"20\d{2}-\d{2}-\d{2}", s):
        return True
    if s.lower() == "rolling" or "\u5e38\u5e74" in s:
        return True
    return False


# Back-compat alias used by older callers
_has_deadline = has_deadline
