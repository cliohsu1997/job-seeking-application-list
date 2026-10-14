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

## Macau + Taiwan universities (full catalog + econ units)

Downloads MOE Taiwan department list, saves all Macau HEIs + Taiwan universities / economics units, fetches department & career pages:

```bash
python code/download_macau_taiwan.py
```

Outputs:
- `data/economics/macau/` — university list, raw career/job HTML, listings
- `data/economics/taiwan/` — MOE CSV, all universities, economics units, dept HTML, listings
- `clean/economics/macau_listings_*.xlsx`, `taiwan_listings_*.xlsx`, `macau_taiwan_listings_*.xlsx`

## China provincial economics pipeline (coastal + Beijing)

Province-level folders under `data/economics/{region}/`. Shenzhen schools stay
under `shenzhen/` (not duplicated in Guangdong).

School lists = 985/211/double-first-class **plus** provincial econ hirers
(same strategy as Guangdong). Catalog:
`data/economics/china_universities/provincial_catalog.json`.

### Available regions

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

```bash
python code/run_province_pipeline.py --list
python code/init_province_pipeline.py --all
```

### One-command pipeline (any province)

Catalog → one-by-one school scrape → seeds → validate/polish → PDF:

```bash
python code/run_province_pipeline.py --region guangdong
python code/run_province_pipeline.py --region beijing
python code/run_province_pipeline.py --region shanghai
python code/run_province_pipeline.py --region fujian
python code/run_province_pipeline.py --region zhejiang
python code/run_province_pipeline.py --region shandong
python code/run_province_pipeline.py --region jiangsu
python code/run_province_pipeline.py --region tianjin
python code/run_province_pipeline.py --region hebei
python code/run_province_pipeline.py --region liaoning
python code/run_province_pipeline.py --region shenzhen
python code/run_province_pipeline.py --region guangxi
python code/run_province_pipeline.py --region hainan
python code/run_province_pipeline.py --all
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
```

Or `--all` on discover / seed / validate:

```bash
python code/discover_by_university.py --all
python code/seed_jobs.py --all
python code/validate_job_links.py --all
```

Lock a good validated set into seeds (reproducible keep list):

```bash
python code/seed_jobs.py --region guangdong --sync-from-validated
```

### Outputs

- `data/economics/china_universities/all.json` — full catalog + national 985 reference
- `data/economics/{region}/universities.json` — detailed school list for that province
- `data/economics/{region}/seeds.json` — known-good announcement URLs
- `data/economics/{region}/discovered_jobs_candidates.json` — pre-validation
- `data/economics/{region}/discovered_jobs_validated.json` — kept rows
- `clean/{region}_listings_validated_YYYY-MM-DD.pdf`

### Task plan

See `.cursor/plan/202610/20261014/expand_coastal_provincial_pipeline.md`.

### Legacy (Guangdong/Fujian/Zhejiang/Shenzhen dept pages only)

```bash
python code/download_china_provinces.py
python code/discover_dept_jobs.py --region guangdong --validate
```
