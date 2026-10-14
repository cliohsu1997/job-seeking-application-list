# -*- coding: utf-8 -*-
"""Province-agnostic seed merge for coastal pipeline.

Loads data/economics/{region}/seeds.json (preferred) and optional Python
fallback modules (seed_guangdong_jobs.SEEDS). Merges into
discovered_jobs_candidates.json before validation.

Usage:
  python seed_jobs.py --region guangdong
  python seed_jobs.py --region guangdong --sync-from-validated
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path

from china_universities_data import REGIONS
from paths import DATA_ECON, ensure_dirs

SEED_FIELDS = (
    "region",
    "Country",
    "Location",
    "University",
    "School",
    "Deadline",
    "JobPostTitle",
    "JobPostUrl",
    "RequiredMaterial",
    "SubmissionChannel",
    "ApplicationLink",
    "Email",
    "JobType",
)


def seeds_path(region: str) -> Path:
    return DATA_ECON / region / "seeds.json"


def load_json_seeds(region: str) -> list[dict]:
    path = seeds_path(region)
    if not path.exists():
        return []
    data = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(data, list):
        return data
    return list(data.get("seeds") or [])


def load_python_fallback(region: str) -> list[dict]:
    if region == "guangdong":
        try:
            from seed_guangdong_jobs import SEEDS

            return list(SEEDS)
        except Exception:
            return []
    return []


def load_seeds(region: str) -> list[dict]:
    """JSON seeds first; Python fallback fills any URLs not already present."""
    by_url: dict[str, dict] = {}
    for s in load_json_seeds(region):
        u = s.get("JobPostUrl")
        if u:
            by_url[u] = s
    for s in load_python_fallback(region):
        u = s.get("JobPostUrl")
        if u and u not in by_url:
            by_url[u] = s
    return list(by_url.values())


def merge_seeds(region: str) -> dict:
    ensure_dirs()
    seeds = load_seeds(region)
    base = DATA_ECON / region
    base.mkdir(parents=True, exist_ok=True)
    path = base / "discovered_jobs_candidates.json"
    existing = []
    if path.exists():
        existing = json.loads(path.read_text(encoding="utf-8"))
    by_url = {j.get("JobPostUrl"): j for j in existing if j.get("JobPostUrl")}
    for s in seeds:
        by_url[s["JobPostUrl"]] = s
    merged = list(by_url.values())
    path.write_text(json.dumps(merged, ensure_ascii=False, indent=2), encoding="utf-8")
    (base / "discovered_jobs.json").write_text(
        json.dumps(merged, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(
        "  seeds merged: json+fallback=%d -> candidates=%d"
        % (len(seeds), len(merged))
    )
    return {"seeds": len(seeds), "candidates": len(merged)}


def sync_seeds_from_validated(region: str) -> int:
    """Rewrite seeds.json from current validated listing (keeps pipeline reproducible)."""
    base = DATA_ECON / region
    src = base / "discovered_jobs_validated.json"
    if not src.exists():
        src = base / "discovered_jobs.json"
    if not src.exists():
        raise FileNotFoundError("No validated jobs for %s" % region)
    jobs = json.loads(src.read_text(encoding="utf-8"))
    seeds = []
    for j in jobs:
        row = {k: j.get(k) for k in SEED_FIELDS}
        row["region"] = region
        if row.get("JobPostUrl"):
            seeds.append(row)
    payload = {
        "region": region,
        "note": "Synced from validated listing; pipeline merges then validates",
        "count": len(seeds),
        "seeds": seeds,
    }
    seeds_path(region).write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print("synced seeds.json: %d rows -> %s" % (len(seeds), seeds_path(region)))
    return len(seeds)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--region", action="append", choices=list(REGIONS))
    ap.add_argument("--all", action="store_true", help="Merge seeds for every region")
    ap.add_argument(
        "--sync-from-validated",
        action="store_true",
        help="Overwrite seeds.json from discovered_jobs_validated.json",
    )
    args = ap.parse_args()
    targets = list(REGIONS) if args.all else (args.region or [])
    if not targets:
        raise SystemExit(
            "Pass --region <name> (repeatable) or --all.\nRegions: %s"
            % ", ".join(REGIONS)
        )
    for region in targets:
        if args.sync_from_validated:
            sync_seeds_from_validated(region)
        merge_seeds(region)


if __name__ == "__main__":
    main()
