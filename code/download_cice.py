"""Download CICE careers listings + detail HTML into ../data/economics.

Index list comes from AjaxHandler getCareersIndex.
Each detail page is an SPA shell; real content is fetched via
getCareersDetail with TDES-encrypted keys (same as the site JS).

HTML is saved under data/economics/cice_html/ and deadlines are scraped
from phrases like "Deadline Date: Nov. 20, 2019" / "November 1, 2026".
"""
from __future__ import annotations

import hashlib
import html as html_lib
import json
import re
import time
from base64 import b64encode
from datetime import date
from pathlib import Path
from urllib.parse import quote, unquote
from urllib.request import Request, urlopen

from Crypto.Cipher import DES3
from Crypto.Util.Padding import pad

from deadline_extract import extract_deadline, html_to_text
from paths import DATA_ECON, ROOT, ensure_dirs

DATA = DATA_ECON
HTML_DIR = DATA / "cice_html"
BASE = "http://www.cicecon.org.cn"
API = BASE + "/ajax/AjaxHandler_CICE2018.ashx"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
# Keys from /Ajax/TDES.js on cicecon.org.cn
_TDES_KEY = b"hfoui7DzDfzhfieNMJ760CUO"
_TDES_IV = b"oMo3xjFb"
FETCH_SLEEP_SEC = 0.2


def clean_text(value: str) -> str:
    value = html_lib.unescape(html_lib.unescape(value or ""))
    value = value.replace("&#32;", " ").replace("%u201C", '"').replace("%u201D", '"')
    return re.sub(r"\s+", " ", value).strip()


def js_unescape(s: str) -> str:
    """Approximate JS unescape() used by CICE ajax responses."""

    def repl_u(mo: re.Match) -> str:
        return chr(int(mo.group(1), 16))

    def repl(mo: re.Match) -> str:
        return chr(int(mo.group(1), 16))

    s = re.sub(r"%u([0-9a-fA-F]{4})", repl_u, s)
    s = re.sub(r"%([0-9a-fA-F]{2})", repl, s)
    return s


def tdes_encrypt(text: str) -> str:
    """Match TDES.encrypt in /Ajax/TDES.js (md5 suffix + 3DES-CBC + b64)."""
    payload = text + "$" + hashlib.md5(text.encode("utf-8")).hexdigest()
    data = pad(payload.encode("utf-8"), 8)
    ct = DES3.new(_TDES_KEY, DES3.MODE_CBC, _TDES_IV).encrypt(data)
    # JS strips trailing NUL bytes from the ciphertext string before b64
    return b64encode(ct.rstrip(b"\x00")).decode("ascii")


def _post_api(body: bytes, referer: str) -> str:
    req = Request(
        API,
        data=body,
        headers={
            "User-Agent": UA,
            "Content-Type": "application/x-www-form-urlencoded",
            "Referer": referer,
        },
        method="POST",
    )
    with urlopen(req, timeout=60) as resp:
        return resp.read().decode("utf-8", "replace")


def fetch_careers_index_html() -> str:
    raw = _post_api(b"cmd=getCareersIndex&k=", BASE + "/careers/index.html")
    m = re.search(r"list:\s*'([^']*)'", raw) or re.search(
        r'list:\s*"([^"]*)"', raw
    )
    if not m:
        raise RuntimeError("Could not parse CICE list field from API response")
    return unquote(m.group(1))


def fetch_careers_detail_html(key: str) -> str:
    """Return decoded detail HTML fragment for listing key (may be empty)."""
    if not key:
        return ""
    enc = tdes_encrypt(key)
    body = f"cmd=getCareersDetail&k={quote(enc, safe='')}".encode()
    raw = _post_api(body, BASE + f"/careers/detail.html?k={key}")
    m = re.search(r"info:\s*'(.*?)'\s*,", raw, flags=re.S)
    if not m:
        m = re.search(r'info:\s*"(.*?)"\s*,', raw, flags=re.S)
    if not m:
        return ""
    return js_unescape(m.group(1))


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

        inst_m = re.search(
            r"<h3[^>]*>\s*<a[^>]*>(.*?)</a>\s*</h3>",
            block,
            flags=re.I | re.S,
        )
        institution = clean_text(re.sub(r"<[^>]+>", "", inst_m.group(1))) if inst_m else ""

        school_m = re.search(
            r'<div class="par_p">.*?<p>(.*?)</p>',
            block,
            flags=re.I | re.S,
        )
        school = clean_text(re.sub(r"<[^>]+>", "", school_m.group(1))) if school_m else ""
        if school.lower() in {"access to details", "to apply for jobs"}:
            school = ""
        if school and institution and school.lower() == institution.lower():
            school = ""

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

    seen, out = set(), []
    for it in items:
        if it["url"] in seen:
            continue
        seen.add(it["url"])
        out.append(it)
    return out


def enrich_from_detail_html(item: dict, detail_html: str) -> dict:
    """Attach deadline (+ optional text) scraped from saved detail HTML."""
    text = html_to_text(detail_html)
    deadline = extract_deadline(text) if text else None
    item = dict(item)
    item["deadline"] = deadline
    item["detail_html_path"] = None
    if text:
        item["_detail_text_len"] = len(text)
    return item


def has_cice_content(detail_html: str) -> bool:
    """False for empty SPA shells (Job Requirements / Application = 'nothing')."""
    text = html_to_text(detail_html or "")
    if len(text) < 200:
        return False
    # Empty employer templates on CICE repeat the word "nothing"
    nothings = len(re.findall(r"\bnothing\b", text, flags=re.I))
    if nothings >= 3 and len(text) < 400:
        return False
    if re.search(
        r"Job Requirements\s+nothing\s+Application Instruction\s+nothing",
        text,
        flags=re.I,
    ):
        return False
    return bool(
        re.search(
            r"apply|application|deadline|Ph\.?D|professor|"
            r"\u62db\u8058|\u5e94\u8058|\u5c97\u4f4d",
            text,
            flags=re.I,
        )
    )


def download_cice() -> Path:
    ensure_dirs()
    HTML_DIR.mkdir(parents=True, exist_ok=True)
    today = date.today().isoformat()
    print(f"Downloading CICE careers from {API}")
    page_html = fetch_careers_index_html()
    index_path = HTML_DIR / f"index_{today}.html"
    index_path.write_text(page_html, encoding="utf-8")
    print(f"Saved index HTML {index_path} ({len(page_html)} chars)")

    listings = parse_listings(page_html)
    print(f"Found {len(listings)} index listings; fetching detail pages...")

    kept: list[dict] = []
    skipped_empty = 0
    with_deadline = 0
    for i, item in enumerate(listings, 1):
        key = item.get("key") or ""
        detail_html = ""
        try:
            detail_html = fetch_careers_detail_html(key)
        except Exception as exc:
            print(f"  [{i}/{len(listings)}] {key}: fetch error {exc}")
        if detail_html:
            html_path = HTML_DIR / f"{key}.html"
            html_path.write_text(detail_html, encoding="utf-8")
            item["detail_html_path"] = str(html_path.relative_to(ROOT)).replace(
                "\\", "/"
            )
            item["deadline"] = enrich_from_detail_html(item, detail_html).get(
                "deadline"
            )
        else:
            item["deadline"] = None
            item["detail_html_path"] = None

        if not has_cice_content(detail_html):
            skipped_empty += 1
            print(
                "  skip empty CICE page: %s"
                % ((item.get("institution") or key)[:50],)
            )
            time.sleep(FETCH_SLEEP_SEC)
            continue
        from deadline_extract import html_to_text, page_has_only_past_deadlines

        if detail_html and page_has_only_past_deadlines(html_to_text(detail_html)):
            skipped_empty += 1
            print(
                "  skip closed (past deadline) CICE: %s"
                % ((item.get("institution") or key)[:50],)
            )
            time.sleep(FETCH_SLEEP_SEC)
            continue

        kept.append(item)
        if item.get("deadline"):
            with_deadline += 1
        if i % 10 == 0 or i == len(listings):
            print(
                f"  [{i}/{len(listings)}] kept={len(kept)} "
                f"with_deadline={with_deadline} "
                f"last={item.get('institution', '')[:40]!r} "
                f"dl={item.get('deadline')}"
            )
        time.sleep(FETCH_SLEEP_SEC)

    json_path = DATA / f"cice_listings_{today}.json"
    for it in kept:
        it.pop("_detail_text_len", None)
    json_path.write_text(
        json.dumps(kept, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(
        f"Saved {json_path} ({len(kept)} listings, "
        f"{with_deadline} with deadline; skipped_empty={skipped_empty})"
    )
    return json_path


def load_detail_html_for_url(url: str) -> str:
    """Load saved detail HTML for a CICE careers URL, or fetch live."""
    m = re.search(r"[?&]k=([A-Fa-f0-9]+)", url or "")
    if not m:
        return ""
    key = m.group(1)
    path = HTML_DIR / f"{key}.html"
    if path.exists():
        return path.read_text(encoding="utf-8")
    try:
        html = fetch_careers_detail_html(key)
    except Exception:
        return ""
    if html:
        HTML_DIR.mkdir(parents=True, exist_ok=True)
        path.write_text(html, encoding="utf-8")
    return html


def main() -> None:
    download_cice()


if __name__ == "__main__":
    main()
