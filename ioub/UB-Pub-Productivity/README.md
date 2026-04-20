# UB Publication Productivity

This repository stores lecturer master data, yearly history, and monthly snapshots using CSV files only.

## Folder layout

- `data/master/lecturers.csv`: lecturer master data
- `data/master/faculties.csv`: faculty reference data
- `data/history/author_metrics_yearly.csv`: yearly historical metrics
- `data/snapshots/author_metrics_YYYY-MM.csv`: one monthly snapshot file
- `data/derived/`: generated dashboard-ready CSV outputs
- `scripts/update_yearly_data.py`: updates `author_metrics_yearly.csv` from the Scopus APIs
- `scripts/create_monthly_snapshots.py`: creates monthly snapshot CSV files from the Scopus APIs

## Update yearly data

Update `author_metrics_yearly.csv` from Scopus. The script uses:
- `Scopus Search API` for publication counts by year
- `Citation Overview API` for citations by year
- `Author Retrieval API` for current h-index

Set your API key first:

```bash
export SCOPUS_API_KEY="your_api_key"
python3 scripts/update_yearly_data.py --start-year 2016 --end-year 2025
```

## Create monthly snapshots

Create a monthly snapshot directly from Scopus:

```bash
export SCOPUS_API_KEY="your_api_key"
python3 scripts/create_monthly_snapshots.py --month 2026-04
```

If you already have the previous month's snapshot file, `citation_count_month` will be
calculated from the difference in `citation_count_total`. You can also point to a specific
previous file with `--previous-snapshot`.
