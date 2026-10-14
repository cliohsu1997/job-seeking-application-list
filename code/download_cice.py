"""Download CICE careers listings via AjaxHandler into ../data."""
from __future__ import annotations

import html
import json
import re
from datetime import date
from pathlib import Path
from urllib.parse import unquote
from urllib.request import Request, urlopen

from paths import DATA_ECON, ensure_dirs

DATA = DATA_ECON
BASE = "http://www.cicecon.org.cn"
API = BASE + "/ajax/AjaxHandler_CICE2018.ashx"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"


def clean_text(value: str) -> str:
    value = html.unescape(html.unescape(value or ""))
    value = value.replace("&#32;", " ").replace("%u201C", '"').replace("%u201D", '"')
    return re.sub(r"\s+", " ", value).strip()


def fetch_careers_html() -> str:
    body = b"cmd=getCareersIndex&k="
    req = Request(
        API,
        data=body,
        headers={
            "User-Agent": UA,
            "Content-Type": "application/x-www-form-urlencoded",
            "Referer": BASE + "/careers/index.html",
        },
        method="POST",
    )
    with urlopen(req, timeout=60) as resp:
        raw = resp.read().decode("utf-8", "replace")
    m = re.search(r"list:\s*'([^']*)'", raw) or re.search(r'list:\s*"([^"]*)"', raw)
    if not m:
        raise RuntimeError("Could not parse CICE list field from API response")
    return unquote(m.group(1))


def parse_listings(page_html: str) -> list[dict]:
    items = []
    blocks = re.split(r"<li\b", page_html, flags=re.I)
    for block in blocks[1:]:
        href_m = re.search(
            r'href=["\'](/careers/detail\.html\?k=[^"\']+)["\']',
            block,
            flags=re.I,
        )
        if not href_m:
            continue
        rel = href_m.group(1)
        url = BASE + rel
        key_m = re.search(r"[?&]k=([^&\"']+)", rel)
        key = key_m.group(1) if key_m else ""

        # Institution in h3 > a
        inst_m = re.search(
            r"<h3[^>]*>\s*<a[^>]*>(.*?)</a>\s*</h3>",
            block,
            flags=re.I | re.S,
        )
        institution = clean_text(re.sub(r"<[^>]+>", "", inst_m.group(1))) if inst_m else ""

        # School / unit in the <p> under par_p
        school_m = re.search(
            r'<div class="par_p">.*?<p>(.*?)</p>',
            block,
            flags=re.I | re.S,
        )
        school = clean_text(re.sub(r"<[^>]+>", "", school_m.group(1))) if school_m else ""
        if school.lower() in {"access to details", "to apply for jobs"}:
            school = ""
        # if school duplicates institution, clear
        if school and institution and school.lower() == institution.lower():
            school = ""

        # Description blurb
        desc_m = re.search(
            r'<div class="partner_box_p">\s*<p>(.*?)</p>',
            block,
            flags=re.I | re.S,
        )
        description = clean_text(re.sub(r"<[^>]+>", " ", desc_m.group(1))) if desc_m else ""

        items.append(
            {
                "key": key,
                "url": url,
                "institution": institution,
                "school": school,
                "description": description,
            }
        )

    # de-dupe by url
    seen, out = set(), []
    for it in items:
        if it["url"] in seen:
            continue
        seen.add(it["url"])
        out.append(it)
    return out


def download_cice() -> Path:
    ensure_dirs()
    today = date.today().isoformat()
    print(f"Downloading CICE careers from {API}")
    page_html = fetch_careers_html()
    listings = parse_listings(page_html)
    json_path = DATA / f"cice_listings_{today}.json"
    json_path.write_text(json.dumps(listings, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Saved {json_path} ({len(listings)} listings)")
    return json_path


def main() -> None:
    download_cice()


if __name__ == "__main__":
    main()
