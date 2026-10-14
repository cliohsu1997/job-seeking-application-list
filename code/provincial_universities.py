# -*- coding: utf-8 -*-
"""Provincial university catalogs: 211 + strong econ schools per coastal region.

Source JSON: data/economics/china_universities/provincial_catalog.json
"""
from __future__ import annotations

import json
from pathlib import Path

_CATALOG_PATH = (
    Path(__file__).resolve().parents[1]
    / "data/economics/china_universities/provincial_catalog.json"
)

REGIONS = (
    "beijing",
    "tianjin",
    "hebei",
    "liaoning",
    "shanghai",
    "jiangsu",
    "zhejiang",
    "fujian",
    "shandong",
    "guangdong",
    "shenzhen",
    "guangxi",
    "hainan",
)

PROVINCE_ZH = {
    "beijing": "\u5317\u4eac",
    "tianjin": "\u5929\u6d25",
    "hebei": "\u6cb3\u5317",
    "liaoning": "\u8fbd\u5b81",
    "shanghai": "\u4e0a\u6d77",
    "jiangsu": "\u6c5f\u82cf",
    "zhejiang": "\u6d59\u6c5f",
    "fujian": "\u798f\u5efa",
    "shandong": "\u5c71\u4e1c",
    "guangdong": "\u5e7f\u4e1c",
    "shenzhen": "\u5e7f\u4e1c",
    "guangxi": "\u5e7f\u897f",
    "hainan": "\u6d77\u5357",
}


def load_catalog() -> dict:
    return json.loads(_CATALOG_PATH.read_text(encoding="utf-8"))


def universities_by_region() -> dict[str, list[dict]]:
    """Always reload from disk so catalog edits are picked up mid-session."""
    if not _CATALOG_PATH.exists():
        return {}
    return dict(load_catalog().get("universities_by_region") or {})


def refresh_provincial_universities() -> dict[str, list[dict]]:
    """Refresh module cache; prefer universities_by_region() for live reads."""
    global PROVINCIAL_UNIVERSITIES
    PROVINCIAL_UNIVERSITIES = universities_by_region()
    return PROVINCIAL_UNIVERSITIES


# Cached snapshot (call refresh_provincial_universities after catalog edits)
PROVINCIAL_UNIVERSITIES: dict[str, list[dict]] = universities_by_region()


def clean_provincial_list(rows: list[dict]) -> list[dict]:
    out, seen = [], set()
    for u in rows:
        if "skip" in (u.get("tags") or []):
            continue
        zh = u.get("name_zh") or ""
        if zh and zh not in seen and u.get("home"):
            seen.add(zh)
            out.append(u)
    return out
