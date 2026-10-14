# How to run

```bash
pip install -r requirements.txt
python code/update.py
```

That regenerates the four `clean/` root finals (see below).

## Pipeline (necessary and sufficient)

| Track | Steps | Module |
|-------|--------|--------|
| **A. Website** | download/compile JOE → verify links → EJM → CICE → **page enrich** (Faculty/Postdoc + Deadline) → validate links | `update.py` A1–A6 |
| **Marketing** | download → compile → clean | `update.py` M1 |
| **B. Provincial** | (discovery separately) then **page enrich** for coastal + **Macau/Taiwan** | `enrich_provincial_from_pages.py` |
| **C. Combine** | AEA Asia-Pacific + provincial(+MT) → three listing sets | `combine_asia_pacific.py` |
| **D. Export** | PDFs + prune non-finals out of `clean/` | `export_clean_pdf.py` |

Provincial discovery (only when refreshing seeds / 2× schools):

```bash
python code/run_province_pipeline.py --region shanghai
python code/run_province_pipeline.py --all --continue-on-error
```

Macau/Taiwan catalog refresh (optional; listings already under `data/economics/{macau,taiwan}/`):

```bash
python code/download_macau_taiwan.py
```

Rebuild finals from existing data (no re-download):

```bash
python code/enrich_website_from_pages.py
python code/enrich_provincial_from_pages.py
python code/combine_asia_pacific.py
python code/export_clean_pdf.py
```

## `clean/` root only

| File | Content |
|------|---------|
| `marketing_listings_*.pdf` | Marketing |
| `aea_listings_*.pdf` + `.xlsx` | AEA/EJM/CICE (CN/MY/JP/AU/NZ/HK/TW/MO) |
| `provincial_listings_*.pdf` + `.xlsx` | Coastal China **+ Macau + Taiwan** |
| `asia_pacific_listings_*.pdf` + `.xlsx` | Combined (AEA sheet top, provincial+MT bottom) |

Everything else under `clean/economics/` or `clean/economics/province/`.

**Job Type** on finals: `Faculty` | `Postdoc` | `Non-Academic`.  
Also: **Published Year**, **College Ranking** (`985` / `211` / `985/211`).  
Deadline is a date (`YYYY-MM-DD`) or `Rolling` when the page says open/常年有效; blank if the page has no extractable date. CICE detail HTML is saved under `data/economics/cice_html/` and scraped for phrases like `Deadline Date: Nov. …`.
