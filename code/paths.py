"""Shared project paths (economics vs marketing)."""
from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

DATA = ROOT / "data"
DATA_ECON = DATA / "economics"
DATA_ECON_MACAU = DATA_ECON / "macau"
DATA_ECON_TAIWAN = DATA_ECON / "taiwan"
DATA_MKT = DATA / "marketing"

CLEAN = ROOT / "clean"
CLEAN_ECON = CLEAN / "economics"
CLEAN_ECON_PROVINCE = CLEAN_ECON / "province"
CLEAN_MKT = CLEAN / "marketing"

APP_XLSX = ROOT / "application list.xlsx"
CODE = ROOT / "code"


def coastal_regions() -> tuple[str, ...]:
    from china_universities_data import REGIONS

    return tuple(REGIONS)


def ensure_dirs() -> None:
    dirs = [
        DATA_ECON,
        DATA_ECON_MACAU,
        DATA_ECON_TAIWAN,
        DATA_MKT,
        CLEAN_ECON,
        CLEAN_ECON_PROVINCE,
        CLEAN_MKT,
        DATA_ECON / "china_universities",
        DATA_ECON / "china_universities" / "by_region",
    ]
    for region in coastal_regions():
        dirs.append(DATA_ECON / region)
    for p in dirs:
        p.mkdir(parents=True, exist_ok=True)


# Back-compat aliases used by older scripts
DATA_ECON_GUANGDONG = DATA_ECON / "guangdong"
DATA_ECON_FUJIAN = DATA_ECON / "fujian"
DATA_ECON_ZHEJIANG = DATA_ECON / "zhejiang"
DATA_ECON_SHENZHEN = DATA_ECON / "shenzhen"
