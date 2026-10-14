"""
Ensure marketing source files are present under data/marketing/.

Copies from Downloads if missing:
  - Marketing Job Market.xlsx
  - Marketing Job Market - 2026 Job Postings.pdf
"""
from __future__ import annotations

import shutil
from pathlib import Path

from paths import DATA_MKT, ensure_dirs

FILES = (
    "Marketing Job Market.xlsx",
    "Marketing Job Market - 2026 Job Postings.pdf",
)


def download_marketing() -> Path:
    ensure_dirs()
    downloads = Path.home() / "Downloads"
    out = None
    for name in FILES:
        dest = DATA_MKT / name
        src = downloads / name
        if dest.exists():
            print(f"  present: {dest}")
            out = dest
            continue
        if src.exists():
            shutil.copy2(src, dest)
            print(f"  copied: {src} -> {dest}")
            out = dest
        else:
            print(f"  MISSING: {dest} (also not in Downloads)")
    if out is None:
        raise FileNotFoundError(f"No marketing source in {DATA_MKT} or Downloads")
    return DATA_MKT / FILES[0]


def main() -> None:
    download_marketing()


if __name__ == "__main__":
    main()
