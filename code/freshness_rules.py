# -*- coding: utf-8 -*-
"""Shared freshness / rolling-deadline helpers for advanced filtering."""
from __future__ import annotations

import re

YEAR_2026_PLUS = re.compile(r"202[6-9]|20[3-9]\d")
YEAR_2025 = re.compile(r"2025")
YEAR_OLD = re.compile(r"201[0-9]|202[0-4]")
ROLLING = re.compile(
    "("
    "\u957f\u671f\u6709\u6548|\u5e38\u5e74\u62db\u8058|\u5e38\u5e74\u63a5\u6536|"
    "\u62db\u6ee1\u5373\u6b62|\u6eda\u52a8\u62db\u8058|\u968f\u65f6\u7533\u8bf7|"
    "\u5168\u5e74\u5f00\u653e|\u957f\u671f\u62db\u8058|until\\s+filled|open\\s+until|"
    "rolling|year[- ]?round"
    ")",
    re.I,
)


def strip_html(html: str) -> str:
    text = re.sub(r"(?is)<script[^>]*>.*?</script>", " ", html or "")
    text = re.sub(r"(?is)<style[^>]*>.*?</style>", " ", text)
    text = re.sub(r"<[^>]+>", " ", text)
    return re.sub(r"\s+", " ", text).strip()
