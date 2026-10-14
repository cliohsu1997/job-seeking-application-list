"""
Update pipeline: economics (AEA/JOE + EJM + CICE) and marketing (MJM spreadsheet).

Usage:
    python code/update.py
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

    print("=== ECONOMICS 1/6 AEA / JOE download + compile ===")
    download_joe_xml()
    compile_joe()

    print("=== ECONOMICS 2/6 AEA link verify ===")
    verify_aea_links()

    print("=== ECONOMICS 3/6 EJM ===")
    download_ejm_json()
    compile_ejm()

    print("=== ECONOMICS 4/6 CICE ===")
    download_cice()
    compile_cice()

    print("=== ECONOMICS 5/6 combine -> clean/economics/ ===")
    combine_clean()

    print("=== ECONOMICS 6/6 validate Application Links ===")
    validate_clean_links()

    print("=== MARKETING 1/3 ensure source files ===")
    download_marketing()

    print("=== MARKETING 2/3 compile spreadsheet ===")
    compile_marketing()

    print("=== MARKETING 3/3 clean/marketing/ ===")
    combine_marketing_clean()

    print("=== PDF -> clean/ (root, not subfolders) ===")
    export_clean_pdf()
    print("Done.")


if __name__ == "__main__":
    main()
