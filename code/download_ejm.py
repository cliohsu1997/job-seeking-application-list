"""Download EconJobMarket (EJM) public ads JSON into ../data (streamed)."""
from __future__ import annotations

from datetime import date
from pathlib import Path
from urllib.request import Request, urlopen

from paths import DATA_ECON, ensure_dirs

DATA = DATA_ECON
EJM_JSON_URL = "https://backend.econjobmarket.org/data/zz_public/json/Ads"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
CHUNK_SIZE = 256 * 1024


def download_ejm_json() -> Path:
    ensure_dirs()
    today = date.today().isoformat()
    path = DATA / f"ejm_listings_{today}.json"

    print(f"Downloading {EJM_JSON_URL}")
    req = Request(
        EJM_JSON_URL,
        headers={"User-Agent": UA, "Accept": "application/json"},
    )
    total = 0
    with urlopen(req, timeout=180) as resp, path.open("wb") as out:
        while True:
            chunk = resp.read(CHUNK_SIZE)
            if not chunk:
                break
            out.write(chunk)
            total += len(chunk)
            print(f"  ... {total:,} bytes", end="\r")

    print(f"\nSaved {path} ({total:,} bytes)")
    return path


def main() -> None:
    download_ejm_json()


if __name__ == "__main__":
    main()
