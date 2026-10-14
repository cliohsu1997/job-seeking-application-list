"""Canonical AEA JOE listing URLs."""
from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from pathlib import Path

JOE_LISTING_TMPL = "https://www.aeaweb.org/joe/listing.php?JOE_ID={joe_id}"


def local_tag(tag: str) -> str:
    return tag.split("}")[-1].lower() if tag else ""


def read_year_issue(xml_path: Path) -> tuple[str, str]:
    """Return (year, zero-padded issue), e.g. ('2026', '02')."""
    # lightweight: only scan start of file
    head = xml_path.read_text(encoding="utf-8", errors="replace")[:2000]
    ym = re.search(r'joe_year_ID="(\d+)"', head)
    im = re.search(r'joe_issue_ID="(\d+)"', head)
    if not ym or not im:
        root = ET.parse(xml_path).getroot()
        year_el = next(root.iter())
        # fallback walk
        year = issue = None
        for el in root.iter():
            if local_tag(el.tag) == "year":
                year = el.attrib.get("joe_year_ID")
            if local_tag(el.tag) == "issue":
                issue = el.attrib.get("joe_issue_ID")
                break
        if not year or not issue:
            raise ValueError(f"Could not read year/issue from {xml_path}")
        return year, f"{int(issue):02d}"
    return ym.group(1), f"{int(im.group(1)):02d}"


def canonical_joe_id(jp_id: str, year: str, issue: str) -> str:
    """Build JOE_ID like 2026-02_111477855."""
    jp_id = str(jp_id).strip()
    if re.match(r"^\d{4}-\d{2}_\d+$", jp_id):
        return jp_id
    # bare numeric id from XML
    return f"{year}-{issue}_{jp_id}"


def listing_url(jp_id: str, year: str, issue: str) -> str:
    return JOE_LISTING_TMPL.format(joe_id=canonical_joe_id(jp_id, year, issue))
