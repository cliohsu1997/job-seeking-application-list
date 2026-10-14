# -*- coding: utf-8 -*-
"""Shared deadline parsing from listing-page text.

Looks for Deadline / closing / apply-by language, then English months
(Jan...Dec), ISO (2026-11-30), slash dates (2026/11/30, 11/30/2026),
and Chinese year-month-day forms.
"""
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
# Compact Chinese without ?: 2026?11?30
_CN_SHORT = re.compile(
    r"(20\d{2})\s*\u5e74\s*(\d{1,2})\s*\u6708\s*(\d{1,2})"
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
# Month DD[, ]YYYY  / Nov. 20, 2019 / January 15th 2026
_EN_MDY = re.compile(
    rf"{_MONTH}\.?\s+(\d{{1,2}})\s*(?:st|nd|rd|th)?"
    rf"\s*[,/]?\s*(20\d{{2}})",
    re.I,
)
# DD Month YYYY
_EN_DMY = re.compile(
    rf"(\d{{1,2}})\s*(?:st|nd|rd|th)?\s+{_MONTH}\.?\s*,?\s*(20\d{{2}})",
    re.I,
)
# Month YYYY (no day) -> last day of month when near deadline keyword only
_EN_MY = re.compile(rf"{_MONTH}\.?\s+(20\d{{2}})\b", re.I)

_KW_CORE = (
    r"deadline(?:\s*date)?|"
    r"closing\s+date|"
    r"application\s+deadline|"
    r"due\s+date|"
    r"due|"
    r"apply\s+by|"
    r"expected\s+by|"
    r"applications?\s+(?:due|close|must\s+be\s+received|accepted\s+until)|"
    r"must\s+be\s+received\s+by|"
    r"received\s+by|"
    r"no\s+later\s+than|"
    r"full\s+consideration|"
    r"review\s+(?:of\s+)?applications?\s+begins|"
    r"applications?\s+will\s+be\s+reviewed|"
    r"submit(?:ted)?\s+before|"
    r"before\s+the|"
    r"(?:submit|submitted|send|sent).{0,160}?\b(?:by|before)\b|"
    r"(?:letters|materials|applications?)\s+(?:by|before)\b|"
    r"\u622a\u6b62\u65e5\u671f|\u62a5\u540d\u622a\u6b62|\u622a\u6b62\u65f6\u95f4|"
    r"\u7533\u8bf7\u622a\u6b62|\u6750\u6599\u63d0\u4ea4|\u8bf7\u4e8e|"
    r"\u62a5\u540d\u65f6\u95f4|\u7533\u8bf7\u65f6\u95f4|\u622a\u6b62"
)

# "expected by DATE" always counts as a deadline keyword.
_EN_KW = re.compile(r"(?:" + _KW_CORE + r").{0,160}", re.I)
_NEAR_KW = _EN_KW


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


def _month_end(y: int, m: int) -> str | None:
    if m == 12:
        return _ymd(y, 12, 31)
    nxt = date(y, m + 1, 1)
    from datetime import timedelta

    last = nxt - timedelta(days=1)
    return last.isoformat()


def _collect_en_dates(window: str, *, allow_month_year: bool = False) -> list[str]:
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
    for cre in (_ISO, _CN, _CN_SHORT):
        for dm in cre.finditer(window):
            s = _ymd(dm.group(1), dm.group(2), dm.group(3))
            if s:
                found.append(s)
    if allow_month_year and not found:
        for m in _EN_MY.finditer(window):
            mo = _month_num(m.group(1))
            if mo:
                s = _month_end(int(m.group(2)), mo)
                if s:
                    found.append(s)
    return found


def _normalize_text(text: str) -> str:
    text = (
        text.replace("\xa0", " ")
        .replace("&nbsp;", " ")
        .replace("&#32;", " ")
        .replace("\u3000", " ")
    )
    return re.sub(r"\s+", " ", text)


def extract_deadline(text: str) -> str | None:
    """Return YYYY-MM-DD, 'Rolling', or None if unknown."""
    if not text:
        return None
    text = _normalize_text(text)

    found: list[str] = []
    for m in DEADLINE_RE.finditer(text):
        s = _ymd(m.group(1), m.group(2), m.group(3) or 1)
        if s:
            found.append(s)

    for m in _EN_KW.finditer(text):
        found.extend(_collect_en_dates(m.group(0), allow_month_year=True))

    # If the page mentions Deadline / closing at all, harvest nearby EN/ISO dates
    # from a wider slice of the document (helps sparse CICE/AEA pages).
    if not found and re.search(r"deadline|\u622a\u6b62|closing\s+date", text, re.I):
        for m in re.finditer(
            r"(?:deadline|\u622a\u6b62|closing\s+date).{0,240}",
            text,
            flags=re.I,
        ):
            found.extend(_collect_en_dates(m.group(0), allow_month_year=True))

    # Last resort: future ISO / Month DD, YYYY when apply language is present.
    if not found and re.search(
        r"\bapply\b|application|\u62db\u8058|\u5e94\u8058|\u7533\u8bf7",
        text,
        re.I,
    ):
        candidates = _collect_en_dates(text)
        today = date.today().isoformat()
        future = sorted({d for d in candidates if d >= today})
        if len(future) == 1:
            found.append(future[0])
        elif 1 < len(future) <= 3:
            found.append(future[0])

    if found:
        today = date.today().isoformat()
        # Never keep previous-year / already-passed deadlines
        future = sorted(d for d in found if d >= today)
        if future:
            return future[0]
        found = []

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


def page_has_only_past_deadlines(text: str) -> bool:
    """True when the page has date(s) but every one is before today (closed ad)."""
    if not text:
        return False
    text = _normalize_text(text)
    if ROLLING.search(text) or ROLLING_OK.search(text):
        return False
    found: list[str] = []
    for m in DEADLINE_RE.finditer(text):
        s = _ymd(m.group(1), m.group(2), m.group(3) or 1)
        if s:
            found.append(s)
    for m in _NEAR_KW.finditer(text):
        found.extend(_collect_en_dates(m.group(0), allow_month_year=True))
    if not found:
        found.extend(_collect_en_dates(text))
    if not found:
        return False
    today = date.today().isoformat()
    return all(d < today for d in found)


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


_has_deadline = has_deadline
