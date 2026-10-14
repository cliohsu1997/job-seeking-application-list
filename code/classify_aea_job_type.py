"""
Classify AEA/JOE (and EJM-style) listings into Job Type:
  Faculty | Postdoc | Non-Academic

Prefer JOE XML fields (section, title, full text), then the live listing page.
Page text is cached under data/economics/aea_job_type_page_cache.json.
"""
from __future__ import annotations

import json
import re
import time
from pathlib import Path
from urllib.request import Request, urlopen

from paths import DATA_ECON, ensure_dirs

CACHE_PATH = DATA_ECON / "aea_job_type_page_cache.json"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
FETCH_SLEEP_SEC = 0.25

# Strong postdoc signals (title preferred).
POSTDOC_RE = re.compile(
    r"("
    r"post[\s\-]?doc(?:toral|torate)?|"
    r"postdoctoral|"
    r"\u535a\u58eb\u540e|"
    r"\u5e08\u8d44\u535a\u58eb\u540e"
    r")",
    re.I,
)

# Research fellow / scholar often = postdoc-track (not Research Associate Professor).
# Bare "research associate" is often pre-doc (e.g. Fed RA); only count with postdoc nearby.
RESEARCH_FELLOW_RE = re.compile(
    r"("
    r"research\s+fellow|"
    r"research\s+scholar|"
    r"international\s+research\s+fellow"
    r")",
    re.I,
)
RESEARCH_ASSOCIATE_RE = re.compile(
    r"research\s+associate(?!\s+professor)",
    re.I,
)

# Asia faculty ranks that look like "research fellow" but are tenure-track.
ASIA_RESEARCH_FACULTY_RE = re.compile(
    r"("
    r"assistant\s+research\s+fellow|"
    r"associate\s+research\s+fellow|"
    r"research\s+professor|"
    r"assistant\s+research\s+professor|"
    r"associate\s+research\s+professor"
    r")",
    re.I,
)

FACULTY_RE = re.compile(
    r"("
    r"\b(?:assistant|associate|full)\s+professor\b|"
    r"\bprofessor\b|"
    r"\blecturer\b|"
    r"\bsenior\s+lecturer\b|"
    r"\binstructor\b|"
    r"\bfaculty\b|"
    r"tenure[\s\-]?track|"
    r"tenured|"
    r"open[\s\-]?rank|"
    r"professor\s+of\s+practice|"
    r"teaching\s+professor|"
    r"adjunct\s+(?:assistant|associate|full)?\s*professor|"
    r"\u6559\u6388|"
    r"\u526f\u6559\u6388|"
    r"\u8bb2\u5e08|"
    r"\u6559\u5e08"
    r")",
    re.I,
)

# Predoc / RA / student - not Postdoc.
PREDOC_RE = re.compile(
    r"("
    r"pre[\s\-]?doc(?:toral)?|"
    r"predoctoral|"
    r"research\s+assistant|"
    r"phd\s+student|"
    r"phd\s+scholarship|"
    r"doctoral\s+student|"
    r"graduate\s+economist"
    r")",
    re.I,
)

NONACADEMIC_SECTION_RE = re.compile(
    r"non[\s\-]?academic",
    re.I,
)
TENURE_SECTION_RE = re.compile(
    r"permanent|tenure",
    re.I,
)


def is_nonacademic_section(section: str) -> bool:
    return bool(NONACADEMIC_SECTION_RE.search(section or ""))


def is_tenure_section(section: str) -> bool:
    s = section or ""
    if is_nonacademic_section(s):
        return False
    return bool(TENURE_SECTION_RE.search(s)) and "full-time academic" in s.lower()


def _classify_text(
    title: str,
    body: str,
    section: str = "",
) -> tuple[str | None, str]:
    """
    Return (job_type or None if ambiguous, reason).
    None means still ambiguous among Faculty vs Postdoc (section already academic).
    """
    title = title or ""
    body = body or ""
    section = section or ""
    title_body = f"{title}\n{body[:4000]}"

    if is_nonacademic_section(section):
        return "Non-Academic", "section_nonacademic"

    # Title-first postdoc (wins over faculty words in mixed titles).
    if POSTDOC_RE.search(title):
        return "Postdoc", "title_postdoc"

    # Asia research-faculty ranks on tenure sections stay Faculty.
    if ASIA_RESEARCH_FACULTY_RE.search(title) and is_tenure_section(section):
        return "Faculty", "title_asia_research_faculty"

    if ASIA_RESEARCH_FACULTY_RE.search(title) and FACULTY_RE.search(title):
        return "Faculty", "title_research_faculty"

    # Research fellow / scholar in title (without professor).
    if RESEARCH_FELLOW_RE.search(title) and not FACULTY_RE.search(title):
        if is_tenure_section(section) and ASIA_RESEARCH_FACULTY_RE.search(title):
            return "Faculty", "title_asia_research_faculty"
        return "Postdoc", "title_research_fellow"

    # Research associate without "postdoc" is ambiguous; treat as postdoc only
    # outside permanent tenure sections (JOE visiting / temporary).
    if (
        RESEARCH_ASSOCIATE_RE.search(title)
        and not FACULTY_RE.search(title)
        and not PREDOC_RE.search(title)
        and not is_tenure_section(section)
        and not is_nonacademic_section(section)
    ):
        return "Postdoc", "title_research_associate"

    if FACULTY_RE.search(title):
        return "Faculty", "title_faculty"

    if PREDOC_RE.search(title):
        return "Faculty", "title_predoc_keep_academic"

    # Permanent tenure-track section without postdoc signal -> faculty.
    if is_tenure_section(section):
        return "Faculty", "section_tenure"

    # Body-only postdoc when title was uninformative.
    if POSTDOC_RE.search(body[:3000]) and not FACULTY_RE.search(title):
        # Avoid body false positives when title is clearly faculty-ish open rank etc.
        if not FACULTY_RE.search(body[:800]):
            return "Postdoc", "body_postdoc"
        # Mixed body: prefer faculty if tenure language near top.
        if re.search(r"tenure|assistant professor|associate professor", body[:1500], re.I):
            return "Faculty", "body_faculty_over_postdoc"
        return "Postdoc", "body_postdoc_mixed"

    if FACULTY_RE.search(body[:2000]):
        return "Faculty", "body_faculty"

    if RESEARCH_FELLOW_RE.search(body[:2000]) and not FACULTY_RE.search(body[:2000]):
        return "Postdoc", "body_research_fellow"

    return None, "ambiguous"


def classify_joe_fields(
    *,
    section: str = "",
    title: str = "",
    full_text: str = "",
) -> tuple[str, str]:
    """Classify from JOE XML fields. Defaults ambiguous academic -> Faculty."""
    job_type, reason = _classify_text(title, full_text, section)
    if job_type:
        return job_type, reason
    return "Faculty", "ambiguous_default_faculty"


def classify_ejm_position_types(
    position_types,
    title: str = "",
    body: str = "",
) -> tuple[str, str]:
    """EJM: use position_types enums first, then title/body."""
    names = []
    for p in position_types or []:
        if isinstance(p, dict):
            names.append((p.get("name") or "").strip())
        else:
            names.append(str(p).strip())
    lower = [n.lower() for n in names if n]

    for n in lower:
        if n in ("other nonacademic", "consultant", "private sector", "industry") or (
            "nonacademic" in n or "non-academic" in n
        ):
            return "Non-Academic", "ejm_position_nonacademic"

    faculty_enums = {
        "assistant professor",
        "associate professor",
        "full professor",
        "tenured professor",
        "untenured professor",
        "lecturer",
        "senior lecturer",
        "visiting assistant professor",
        "visiting associate professor",
        "visiting professor",
        "other academic",
    }
    predoc_enums = {
        "research assistant",
        "research assistant (pre-doc)",
        "doctoral student",
    }
    has_postdoc = any(
        "postdoc" in n or n == "postdoctoral scholar" for n in lower
    )
    has_faculty = any(
        n in faculty_enums or "professor" in n or "lecturer" in n for n in lower
    )
    only_predoc = bool(lower) and all(
        n in predoc_enums or "pre-doc" in n or n == "research assistant"
        for n in lower
    )

    # Postdoc-only (or postdoc + RA) -> Postdoc; mixed faculty+postdoc -> Faculty.
    if has_postdoc and not has_faculty:
        return "Postdoc", "ejm_position_postdoc"
    if has_faculty:
        return "Faculty", "ejm_position_faculty"
    if only_predoc and not has_postdoc:
        return "Faculty", "ejm_position_predoc"

    return classify_joe_fields(section="", title=title, full_text=body)


def load_page_cache() -> dict:
    ensure_dirs()
    if not CACHE_PATH.exists():
        return {}
    try:
        return json.loads(CACHE_PATH.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}


def save_page_cache(cache: dict) -> None:
    ensure_dirs()
    CACHE_PATH.write_text(
        json.dumps(cache, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def _fetch_cice_detail_text(url: str) -> str:
    """CICE detail.html is an empty SPA; load saved HTML or Ajax detail body."""
    try:
        from deadline_extract import html_to_text
        from download_cice import load_detail_html_for_url
    except Exception:
        return ""
    html = load_detail_html_for_url(url)
    text = html_to_text(html)
    return text[:12000] if text else ""


def fetch_listing_text(url: str, cache: dict | None = None) -> str:
    """Fetch JOE listing page text; use cache when present."""
    if not url:
        return ""
    own = cache is None
    if cache is None:
        cache = load_page_cache()
    is_cice = "cicecon.org.cn" in url and "/careers/detail" in url
    hit = cache.get(url)
    if isinstance(hit, dict) and hit.get("text") is not None:
        cached = hit.get("text") or ""
        # Reuse only real CICE Ajax detail text (SPA shell caches are useless)
        if is_cice:
            if hit.get("source") == "cice_detail" and len(cached) >= 200:
                return cached
        elif cached or "error" in hit:
            return cached
    elif isinstance(hit, str) and not is_cice:
        return hit

    text = ""
    source = "http"
    try:
        if is_cice:
            text = _fetch_cice_detail_text(url)
            source = "cice_detail"
        if not text:
            req = Request(url, headers={"User-Agent": UA})
            with urlopen(req, timeout=30) as resp:
                raw = resp.read()
            text = raw.decode("utf-8", errors="replace")
            text = re.sub(r"(?is)<script[^>]*>.*?</script>", " ", text)
            text = re.sub(r"(?is)<style[^>]*>.*?</style>", " ", text)
            text = re.sub(r"<[^>]+>", " ", text)
            text = re.sub(r"\s+", " ", text).strip()[:12000]
            source = "http"
    except Exception as exc:
        cache[url] = {"text": "", "error": str(exc)}
        if own:
            save_page_cache(cache)
        return ""

    cache[url] = {"text": text, "source": source}
    if own:
        save_page_cache(cache)
        time.sleep(FETCH_SLEEP_SEC)
    return text


def classify_with_optional_fetch(
    *,
    section: str = "",
    title: str = "",
    full_text: str = "",
    url: str = "",
    fetch_ambiguous: bool = False,
    fetch_always: bool = False,
    cache: dict | None = None,
) -> tuple[str, str, bool]:
    """
    Returns (job_type, reason, was_ambiguous_before_default).
    fetch_always: always download listing page and merge into classification.
    fetch_ambiguous: fetch only when title/body/section still unclear.
    """
    page = ""
    if url and (fetch_always or fetch_ambiguous):
        # For fetch_ambiguous, only hit the network when needed (below).
        if fetch_always:
            page = fetch_listing_text(url, cache=cache)

    blob = full_text or ""
    if page:
        blob = (blob + "\n" + page).strip()

    job_type, reason = _classify_text(title, blob, section)
    if job_type:
        if page and reason.startswith(("body_", "title_")):
            return job_type, "page_" + reason, False
        return job_type, reason, False

    ambiguous = True
    if url and fetch_ambiguous and not fetch_always:
        page = fetch_listing_text(url, cache=cache)
        if page:
            job_type2, reason2 = _classify_text(title, page, section)
            if job_type2:
                return job_type2, f"page_{reason2}", True

    if page:
        job_type2, reason2 = _classify_text(title, page, section)
        if job_type2:
            return job_type2, f"page_{reason2}", True

    return "Faculty", "ambiguous_default_faculty", ambiguous
