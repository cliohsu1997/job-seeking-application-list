# -*- coding: utf-8 -*-
"""Back-compat wrapper: coastal-only summarize.

Prefer: python summarize_china_jobs.py
"""
from __future__ import annotations

import argparse
import json

from summarize_china_jobs import summarize


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--prefer", choices=("verified", "validated"), default="verified")
    args = ap.parse_args()
    print(
        json.dumps(
            summarize(prefer=args.prefer, coastal_only=True),
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
