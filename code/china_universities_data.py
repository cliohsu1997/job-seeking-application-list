# -*- coding: utf-8 -*-
"""Unified coastal university catalog for the province pipeline."""
from __future__ import annotations

from provincial_universities import (
    PROVINCE_ZH,
    PROVINCIAL_UNIVERSITIES,
    REGIONS,
    clean_provincial_list,
)

__all__ = [
    "REGIONS",
    "PROVINCE_ZH",
    "NATIONAL_985",
    "build_coastal_universities",
    "universities_for_region",
]

NATIONAL_985 = [
    {"name_zh": "\u6e05\u534e\u5927\u5b66", "name_en": "Tsinghua University", "province_zh": "\u5317\u4eac", "tags": ["985", "211"]},
    {"name_zh": "\u5317\u4eac\u5927\u5b66", "name_en": "Peking University", "province_zh": "\u5317\u4eac", "tags": ["985", "211"]},
    {"name_zh": "\u590d\u65e6\u5927\u5b66", "name_en": "Fudan University", "province_zh": "\u4e0a\u6d77", "tags": ["985", "211"]},
    {"name_zh": "\u4e0a\u6d77\u4ea4\u901a\u5927\u5b66", "name_en": "Shanghai Jiao Tong University", "province_zh": "\u4e0a\u6d77", "tags": ["985", "211"]},
    {"name_zh": "\u6d59\u6c5f\u5927\u5b66", "name_en": "Zhejiang University", "province_zh": "\u6d59\u6c5f", "tags": ["985", "211"]},
    {"name_zh": "\u5357\u4eac\u5927\u5b66", "name_en": "Nanjing University", "province_zh": "\u6c5f\u82cf", "tags": ["985", "211"]},
    {"name_zh": "\u5c71\u4e1c\u5927\u5b66", "name_en": "Shandong University", "province_zh": "\u5c71\u4e1c", "tags": ["985", "211"]},
    {"name_zh": "\u4e2d\u5c71\u5927\u5b66", "name_en": "Sun Yat-sen University", "province_zh": "\u5e7f\u4e1c", "tags": ["985", "211"]},
    {"name_zh": "\u53a6\u95e8\u5927\u5b66", "name_en": "Xiamen University", "province_zh": "\u798f\u5efa", "tags": ["985", "211"]},
]

UNIVERSITIES_985_211: list[dict] = []


def build_coastal_universities() -> list[dict]:
    """Prefer provincial_catalog.json; enrich econ_units from legacy UNIVERSITIES."""
    from download_china_provinces import UNIVERSITIES
    from provincial_universities import universities_by_region

    known = {
        u["name_zh"]: u
        for rows in UNIVERSITIES.values()
        for u in rows
        if u.get("name_zh")
    }
    rows, seen = [], set()
    # Live reload so catalog JSON edits apply without process restart
    for region, unis in universities_by_region().items():
        for u in clean_provincial_list(unis):
            zh = u["name_zh"]
            key = (region, zh)
            if key in seen:
                continue
            seen.add(key)
            hit = known.get(zh) or {}
            rows.append(
                {
                    "name_zh": zh,
                    "name_en": u.get("name_en") or hit.get("name_en") or zh,
                    "region": region,
                    "city": u.get("city") or hit.get("city"),
                    "tags": list(u.get("tags") or ["known_catalog"]),
                    "home": u.get("home") or hit.get("home"),
                    "career": u.get("career") or hit.get("career"),
                    "postdoc": u.get("postdoc") or hit.get("postdoc"),
                    "econ_units": list(u.get("econ_units") or hit.get("econ_units") or []),
                }
            )
    # Legacy-only schools (if any) not yet in catalog
    for region, unis in UNIVERSITIES.items():
        for u in unis:
            zh = u.get("name_zh")
            if not zh or (region, zh) in seen:
                continue
            seen.add((region, zh))
            rows.append(
                {
                    "name_zh": zh,
                    "name_en": u.get("name_en") or zh,
                    "region": region,
                    "city": u.get("city"),
                    "tags": ["known_catalog"],
                    "home": u.get("home"),
                    "career": u.get("career"),
                    "postdoc": u.get("postdoc"),
                    "econ_units": list(u.get("econ_units") or []),
                }
            )
    return rows


def universities_for_region(region: str, *, reload: bool = True) -> list[dict]:
    global UNIVERSITIES_985_211
    if reload or not UNIVERSITIES_985_211:
        UNIVERSITIES_985_211 = build_coastal_universities()
    return [dict(u) for u in UNIVERSITIES_985_211 if u["region"] == region]
