"""Download AEA JOE full XML export into ../data (streamed in chunks)."""
from __future__ import annotations

from datetime import date
from pathlib import Path
from urllib.request import Request, urlopen

from paths import DATA_ECON, ensure_dirs

DATA = DATA_ECON
JOE_XML_URL = "https://www.aeaweb.org/joe/resultset_output.php?mode=full_xml"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
CHUNK_SIZE = 256 * 1024


def download_joe_xml() -> Path:
    ensure_dirs()
    today = date.today().isoformat()
    xml_path = DATA / f"joe_listings_{today}.xml"

    print(f"Downloading {JOE_XML_URL}")
    req = Request(JOE_XML_URL, headers={"User-Agent": UA})
    total = 0
    with urlopen(req, timeout=180) as resp, xml_path.open("wb") as out:
        while True:
            chunk = resp.read(CHUNK_SIZE)
            if not chunk:
                break
            out.write(chunk)
            total += len(chunk)
            print(f"  ... {total:,} bytes", end="\r")

    print(f"\nSaved {xml_path} ({total:,} bytes)")
    return xml_path


def main() -> None:
    download_joe_xml()


if __name__ == "__main__":
    main()
