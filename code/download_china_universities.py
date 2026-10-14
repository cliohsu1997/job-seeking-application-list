# -*- coding: utf-8 -*-
"""Write per-region university catalogs from provincial_catalog.json.

Usage:
  python download_china_universities.py
  python download_china_universities.py --region beijing
"""
from __future__ import annotations

import argparse
import json
from datetime import date

from china_universities_data import (
    NATIONAL_985,
    PROVINCE_ZH,
    REGIONS,
    build_coastal_universities,
    universities_for_region,
)
from paths import DATA_ECON, ensure_dirs

TODAY = date.today().isoformat()


def write_catalog(regions: list[str] | None = None) -> dict:
    ensure_dirs()
    base = DATA_ECON / "china_universities"
    by_region_dir = base / "by_region"
    for d in (base, by_region_dir):
        d.mkdir(parents=True, exist_ok=True)

    all_rows = build_coastal_universities()
    (base / "all.json").write_text(
        json.dumps(
            {
                "date": TODAY,
                "note": "985/211 + provincial econ schools (coastal + Beijing)",
                "count": len(all_rows),
                "universities": all_rows,
                "national_985_reference": NATIONAL_985,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )

    summary = {"date": TODAY, "regions": {}}
    for region in regions or list(REGIONS):
        rows = universities_for_region(region)
        payload = {
            "date": TODAY,
            "region": region,
            "province_zh": PROVINCE_ZH.get(region, ""),
            "count": len(rows),
            "with_econ_units": sum(1 for r in rows if r.get("econ_units")),
            "tags_985_211": sum(
                1 for r in rows if any(t in (r.get("tags") or []) for t in ("985", "211"))
            ),
            "universities": rows,
        }
        (by_region_dir / f"{region}.json").write_text(
            json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        region_dir = DATA_ECON / region
        region_dir.mkdir(parents=True, exist_ok=True)
        (region_dir / "universities.json").write_text(
            json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        summary["regions"][region] = {
            "count": payload["count"],
            "with_econ_units": payload["with_econ_units"],
            "tags_985_211": payload["tags_985_211"],
        }
        print(
            "-> %s: %d schools (985/211=%d, econ_units=%d)"
            % (
                region,
                payload["count"],
                payload["tags_985_211"],
                payload["with_econ_units"],
            )
        )

    (base / "download_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print("Wrote", base / "all.json")
    return summary


def load_region_universities(region: str) -> list[dict]:
    path = DATA_ECON / region / "universities.json"
    if not path.exists():
        write_catalog([region])
    return json.loads(path.read_text(encoding="utf-8")).get("universities") or []


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--region", action="append", choices=list(REGIONS))
    ap.add_argument("--all", action="store_true", help="Write catalogs for every region")
    args = ap.parse_args()
    regions = list(REGIONS) if args.all or not args.region else args.region
    write_catalog(regions)


if __name__ == "__main__":
    main()
