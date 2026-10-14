"""
One-shot pipeline -> clean/ finals only:

  marketing_listings_*.pdf
  aea_listings_*.pdf + .xlsx
  provincial_listings_*.pdf + .xlsx   (coastal + Macau/Taiwan)
  asia_pacific_listings_*.pdf + .xlsx

Usage:
    python code/update.py

Provincial discovery (slow, per-region) is separate:
    python code/run_province_pipeline.py --all --continue-on-error
Macau/Taiwan catalog refresh (optional):
    python code/download_macau_taiwan.py
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from download_joe import download_joe_xml
from compile_joe import main as compile_joe
from verify_aea_links import main as verify_aea_links
from download_ejm import download_ejm_json
from compile_ejm import main as compile_ejm
from download_cice import download_cice
from compile_cice import main as compile_cice
from combine_clean import main as combine_clean
from validate_clean_links import main as validate_clean_links
from download_marketing import download_marketing
from compile_marketing import main as compile_marketing
from combine_marketing_clean import main as combine_marketing_clean
from export_clean_pdf import main as export_clean_pdf
from paths import ensure_dirs


def main() -> None:
    ensure_dirs()

    # ----- A. Website track (AEA / EJM / CICE) -----
    print("=== A1 JOE download + compile ===")
    download_joe_xml()
    compile_joe()

    print("=== A2 AEA application-link verify ===")
    verify_aea_links()

    print("=== A3 EJM ===")
    download_ejm_json()
    compile_ejm()

    print("=== A4 CICE ===")
    download_cice()
    compile_cice()

    print("=== A5 listing pages -> Faculty/Postdoc + Deadline ===")
    try:
        from enrich_website_from_pages import run as enrich_website_pages

        enrich_website_pages()
    except Exception as exc:
        print("!! website page enrich failed, falling back to combine_clean:", exc)
        combine_clean()

    print("=== A6 validate Application Links ===")
    validate_clean_links()

    # ----- Marketing -----
    print("=== M1 marketing ===")
    download_marketing()
    compile_marketing()
    combine_marketing_clean()

    # ----- B. Provincial + Macau/Taiwan page enrich -----
    print("=== B0 ensure Macau/Taiwan listings JSON ===")
    try:
        from paths import DATA_ECON

        need_mt = any(
            not list((DATA_ECON / r).glob("listings_*.json"))
            for r in ("macau", "taiwan")
        )
        if need_mt:
            from download_macau_taiwan import main as download_macau_taiwan

            download_macau_taiwan()
        else:
            print("  macau/taiwan listings_*.json present")
    except Exception as exc:
        print("!! macau/taiwan ensure skipped:", exc)

    print("=== B provincial(+Macau/Taiwan) pages -> Faculty/Postdoc + Deadline ===")
    try:
        from enrich_provincial_from_pages import run as enrich_provincial_pages

        enrich_provincial_pages()
    except Exception as exc:
        print("!! provincial page enrich skipped:", exc)

    # ----- Combine finals + prune clean/ root -----
    print("=== C combine asia_pacific (aea + provincial+MT) ===")
    try:
        from combine_asia_pacific import combine as combine_asia_pacific

        combine_asia_pacific(prefer="verified", fetch_page=False)
    except Exception as exc:
        print("!! asia-pacific combine skipped:", exc)

    print("=== D export PDFs + prune clean/ root ===")
    export_clean_pdf()
    print("Done.")


if __name__ == "__main__":
    main()
