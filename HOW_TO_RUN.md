# How to run

```bash
pip install -r requirements.txt
python code/update.py
```

## Folder layout

| Path | Contents |
|------|----------|
| `data/economics/` | AEA/JOE, EJM, CICE raw downloads + progress |
| `data/marketing/` | Marketing Job Market.xlsx (+ optional PDF) |
| `clean/economics/` | Economics Application-format xlsx |
| `clean/marketing/` | Marketing Application-format xlsx |
| `clean/*.pdf` | Compiled readable PDFs (**not** in subfolders) |

## Marketing source

Copy into `data/marketing/` (or leave in Downloads; `download_marketing.py` copies them):

- `Marketing Job Market.xlsx`
- `Marketing Job Market - 2026 Job Postings.pdf` (optional archive)

## Let Cursor Fetch without clicking Run

1. Open **Settings → Cursor Settings → Agents → Approvals & Execution**
2. Set Run Mode to **Run Everything** (zero prompts)  
   or **Auto-review** (fewer prompts; this repo steers Fetch via `.cursor/permissions.json`)
3. Start a new Agent chat after changing the mode

Refresh clean + validate only:

```bash
python code/refresh_clean.py
```

Marketing only:

```bash
python code/download_marketing.py
python code/compile_marketing.py
python code/combine_marketing_clean.py
python code/export_clean_pdf.py
```

PDFs:

```bash
python code/export_clean_pdf.py
```

Outputs: `clean/economics_listings_YYYY-MM-DD.pdf`, `clean/marketing_listings_YYYY-MM-DD.pdf`

## Macau + Taiwan

```bash
python code/download_macau_taiwan.py
python code/compile_macau_taiwan.py
```

Outputs:
- `data/economics/macau/listings_*.json`
- `data/economics/taiwan/listings_*.json`
- `clean/economics/macau_listings_*.xlsx`, `taiwan_listings_*.xlsx`, `macau_taiwan_listings_*.xlsx`

## China provincial pipeline (coastal + Beijing)

Province folders under `data/economics/{region}/`. Shenzhen stays under
`shenzhen/` (not duplicated in Guangdong). Catalog:
`data/economics/china_universities/provincial_catalog.json`.

| Region key | Area |
|------------|------|
| `beijing` | Beijing |
| `tianjin` | Tianjin |
| `hebei` | Hebei |
| `liaoning` | Liaoning |
| `shanghai` | Shanghai |
| `jiangsu` | Jiangsu |
| `zhejiang` | Zhejiang |
| `fujian` | Fujian |
| `shandong` | Shandong |
| `guangdong` | Guangdong (excl. Shenzhen) |
| `shenzhen` | Shenzhen |
| `guangxi` | Guangxi |
| `hainan` | Hainan |

### Stages

1. **Preliminary** — discover + seeds + validate → `discovered_jobs_full.json`  
   Target: full list **≥ 2 × schools** in the province  
2. **Advanced** — one-by-one re-fetch; drop wrong content, past deadlines,  
   and **2025 without rolling/招满即止**  
3. **Summarize** — merge coastal verified jobs; optionally add Macau/Taiwan  

```bash
python code/run_province_pipeline.py --list
python code/init_province_pipeline.py --all
python code/run_province_pipeline.py --region shanghai
python code/run_province_pipeline.py --all --continue-on-error
python code/run_province_pipeline.py --check-ready --require-advanced
python code/run_province_pipeline.py --summarize-only --with-macau-taiwan
```

Skip slow scrape (seed + validate only):

```bash
python code/run_province_pipeline.py --region beijing --skip-discover
```

### Step-by-step (same for every region)

```bash
python code/download_china_universities.py --region beijing
python code/discover_by_university.py --region beijing
python code/seed_jobs.py --region beijing
python code/validate_job_links.py --region beijing
python code/advanced_filter_jobs.py --region beijing
python code/check_province_ready.py --require-advanced
python code/summarize_china_jobs.py
```

Coastal-only (no Macau/Taiwan):

```bash
python code/summarize_china_jobs.py --coastal-only
```

Lock a good validated set into seeds:

```bash
python code/seed_jobs.py --region guangdong --sync-from-validated
```

### Outputs

- `data/economics/{region}/seeds.json`
- `data/economics/{region}/discovered_jobs_full.json` — preliminary
- `data/economics/{region}/discovered_jobs_verified.json` — advanced keep
- `clean/{region}_listings_validated_YYYY-MM-DD.pdf`
- `clean/china_coastal_listings_validated_YYYY-MM-DD.pdf`
- `clean/china_macau_taiwan_listings_YYYY-MM-DD.pdf` — coastal + Macau + Taiwan

### Task plan

See `.cursor/plan/202610/20261014/expand_coastal_provincial_pipeline.md`.
