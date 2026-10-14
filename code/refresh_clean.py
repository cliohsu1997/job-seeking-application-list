"""
Refresh sources, rebuild clean economics + marketing, export PDFs.
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from download_ejm import download_ejm_json
from compile_ejm import main as compile_ejm
from download_cice import download_cice
from compile_cice import main as compile_cice
from rebuild_joe_from_progress import main as rebuild_joe
from combine_clean import main as combine_clean
from validate_clean_links import main as validate_clean_links
from download_marketing import download_marketing
from compile_marketing import main as compile_marketing
from combine_marketing_clean import main as combine_marketing_clean
from export_clean_pdf import main as export_clean_pdf
from paths import ensure_dirs


def main() -> None:
    ensure_dirs()

    print("=== Refresh EJM ===")
    download_ejm_json()
    compile_ejm()

    print("=== Refresh CICE ===")
    download_cice()
    compile_cice()

    print("=== Rebuild AEA/JOE (progress overlay) ===")
    rebuild_joe()

    print("=== Rewrite clean/economics/ ===")
    combine_clean()

    print("=== Validate economics Application Links ===")
    validate_clean_links()

    print("=== Marketing compile + clean ===")
    download_marketing()
    compile_marketing()
    combine_marketing_clean()

    print("=== PDFs in clean/ ===")
    export_clean_pdf()
    print("Done.")


if __name__ == "__main__":
    main()
