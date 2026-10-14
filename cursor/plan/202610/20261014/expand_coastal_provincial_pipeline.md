# Expand coastal provincial economics job pipeline

**Date:** 2026-10-14  
**Repo:** https://github.com/cliohsu1997/job-seeking-application-list.git

## Goal

Run the Guangdong-quality faculty + postdoc pipeline at **province level** for
Beijing, Shanghai, Shandong, Zhejiang, and all mainland coastal provinces,
with a detailed school scrape list per province (985/211/??? + provincial
econ hirers).

## Regions (folder keys)

| Region | Province/city | Notes |
|--------|---------------|--------|
| beijing | ?? | PKU / Tsinghua / RUC / CUFE / UIBE / BNU |
| tianjin | ?? | Nankai / TJU / TJUFE |
| hebei | ?? | Hebei U / HEBUT (+ coastal adjacency) |
| liaoning | ?? | DUFE / LNU / DUT / NEU |
| shanghai | ?? | Fudan / SJTU / Tongji / ECNU / SUFE / SUIBE / SHU |
| jiangsu | ?? | NJU / SEU / NUFE / Soochow / Jiangsu U |
| zhejiang | ?? | ZJU + provincial finance/business schools |
| fujian | ?? | XMU + provincial |
| shandong | ?? | SDU / OUC / SDUFE / Qingdao U |
| guangdong | ?? | Validated reference province (excl. Shenzhen) |
| shenzhen | ?? | City split; not duplicated under guangdong |
| guangxi | ?? | GXU / GXUFE |
| hainan | ?? | Hainan U |

## Pipeline (per province)

1. `download_china_universities.py` ù write `universities.json` for the region  
2. `discover_by_university.py` ù one-by-one home / econ unit / HR scrape  
3. `seed_jobs.py` ù merge `seeds.json` (+ Python fallback)  
4. `validate_job_links.py` ù keep real open faculty/postdoc; dedupe; drop dean/stale  
5. Export `clean/{region}_listings_validated_YYYY-MM-DD.pdf`

Or: `python code/run_province_pipeline.py --region {region}`

## Implementation status (2026-10-14)

- [x] Remove stale debug / one-off province scrapers  
- [x] Web-checked Project 211 lists (gk100 / eol.cn) + ??? / finance extras  
- [x] `provincial_catalog.json` ó 122 schools across 13 regions (full 211 per province)  
- [x] Concise loaders: `provincial_universities.py` + `china_universities_data.py`  
- [x] Unified `REGIONS` + `paths.ensure_dirs` for all provinces  
- [x] Guangdong seeds locked (`seeds.json`) + polish in validator  
- [x] `HOW_TO_RUN.md` updated  
- [x] Audit: all coastal 211 present; GD double-first-class complete (incl. GZCMU/GMU); Fujian 10; Zhejiang 13  
- [ ] Run discover+validate for beijing / shanghai / shandong / zhejiang  
- [ ] Run remaining coastal provinces (`--all`)  
- [ ] Sync seeds per province after first good validate  

## Run order (recommended)

1. beijing, shanghai (largest econ markets)  
2. shandong, zhejiang, jiangsu  
3. fujian, tianjin, liaoning  
4. guangxi, hainan, hebei  
5. guangdong / shenzhen refresh as needed  

## Quality rules (from Guangdong)

- Prefer ???? career / announcement pages  
- Keep faculty + ??? only; reject dean-only / student career / stale cycles  
- After a good keep set:  
  `python code/seed_jobs.py --region {region} --sync-from-validated`
