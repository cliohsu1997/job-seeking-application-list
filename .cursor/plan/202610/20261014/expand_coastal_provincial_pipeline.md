# Expand coastal provincial economics job pipeline

**Date:** 2026-10-14  
**Repo:** https://github.com/cliohsu1997/job-seeking-application-list.git

## Goal

Run the Guangdong-quality faculty + postdoc pipeline at **province level** for
Beijing, Shanghai, Shandong, Zhejiang, and all mainland coastal provinces,
with a detailed school scrape list per province (985/211/双一流 + provincial
econ hirers).

## Regions (folder keys)

| Region | Province/city | Notes |
|--------|---------------|--------|
| beijing | 北京 | PKU / Tsinghua / RUC / CUFE / UIBE / BNU |
| tianjin | 天津 | Nankai / TJU / TJUFE |
| hebei | 河北 | Hebei U / HEBUT (+ coastal adjacency) |
| liaoning | 辽宁 | DUFE / LNU / DUT / NEU |
| shanghai | 上海 | Fudan / SJTU / Tongji / ECNU / SUFE / SUIBE / SHU |
| jiangsu | 江苏 | NJU / SEU / NUFE / Soochow / Jiangsu U |
| zhejiang | 浙江 | ZJU + provincial finance/business schools |
| fujian | 福建 | XMU / FZU + provincial |
| shandong | 山东 | SDU / OUC / SDUFE / Qingdao U |
| guangdong | 广东 | Validated reference province (excl. Shenzhen) |
| shenzhen | 深圳 | City split; not duplicated under guangdong |
| guangxi | 广西 | GXU / GXUFE |
| hainan | 海南 | Hainan U |

## Pipeline (per province)

1. `init_province_pipeline.py` / ensure workspace (`seeds.json`, candidates stub)
2. `download_china_universities.py` - write `universities.json`
3. `discover_by_university.py` - one-by-one home / econ unit / HR scrape
4. `seed_jobs.py` - merge `seeds.json` (+ Python fallback)
5. `validate_job_links.py` - keep real open faculty/postdoc; dedupe; drop dean/stale
6. Export `clean/{region}_listings_validated_YYYY-MM-DD.pdf`

One command:

```bash
python code/run_province_pipeline.py --list
python code/run_province_pipeline.py --region {region}
python code/run_province_pipeline.py --all --continue-on-error
```

## Implementation status (2026-10-14)

- [x] Remove stale debug / one-off province scrapers  
- [x] Web-checked Project 211 lists + 双一流 / finance extras  
- [x] `provincial_catalog.json` - 122 schools / 13 regions (full 211 per province)  
- [x] Concise loaders: `provincial_universities.py` + `china_universities_data.py`  
- [x] Unified `REGIONS` + `paths.ensure_dirs` for all provinces  
- [x] Guangdong seeds locked (`seeds.json`) + polish in validator  
- [x] `HOW_TO_RUN.md` updated for all 13 region keys  
- [x] Audit: coastal 211 complete; GD 双一流 complete (GZCMU/GMU); Fujian 10; Zhejiang 13  
- [x] Per-province workspaces via `init_province_pipeline.py`  
- [x] `run_province_pipeline.py --region|--all|--list|--init-only|--continue-on-error`  
- [x] Catalog reloads from disk; empty validate sets handled  
- [x] Two-stage filter: preliminary full list (~2x schools) then advanced one-by-one  
- [x] Advanced drops: wrong content, deadline passed, 2025 without rolling, other stale  
- [x] `check_province_ready.py` / `--summarize-when-ready` before coastal merge  
- [ ] All provinces advanced-ready (>=2x verified) then coastal merge  
- [x] Plan lives under `.cursor/plan/` (tracked via gitignore exceptions)  
- [x] First Shanghai pipeline run (12 schools -> 3 kept Donghua after stale purge)  
- [x] Shanghai coverage fix: hand seeds + catalog career hubs -> **25 kept (>=2x 12 schools)**  
- [x] Validator: hard-reject title years 2010-2023; CAREER_KW + jobs.* hubs  
- [ ] Run discover+validate for beijing / shanghai (enriched seeds) / shandong / zhejiang  
- [ ] Run remaining coastal provinces (`--all`)  
- [ ] Sync seeds per province after first good validate  

## Run order (recommended)

1. beijing, shanghai (largest econ markets)  
2. shandong, zhejiang, jiangsu  
3. fujian, tianjin, liaoning  
4. guangxi, hainan, hebei  
5. guangdong / shenzhen refresh as needed  

## Quality rules (from Guangdong)

- Prefer official career / announcement pages  
- Keep faculty + 博士后 only; reject dean-only / student career / stale cycles  
- After a good keep set:  
  `python code/seed_jobs.py --region {region} --sync-from-validated`
