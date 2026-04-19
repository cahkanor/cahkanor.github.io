# UB Publication Productivity

This repository stores lecturer master data, yearly history, and monthly snapshots using CSV files only.

## Folder layout

- `data/master/lecturers.csv`: lecturer master data
- `data/master/faculties.csv`: faculty reference data
- `data/history/author_metrics_yearly.csv`: yearly historical metrics
- `data/snapshots/author_metrics_YYYY-MM.csv`: one monthly snapshot file
- `data/derived/`: generated dashboard-ready CSV outputs
- `scripts/update_yearly_data.py`: updates `author_metrics_yearly.csv` from a year-specific source CSV
- `scripts/create_monthly_snapshots.py`: creates monthly snapshot CSV files in the current format

## Update yearly data

Update `author_metrics_yearly.csv` from a year-specific source file such as `author_metrics_by_year_2020_2025_from_2026.csv`:

```bash
python3 scripts/update_yearly_data.py
```

## Create monthly snapshots

Create monthly snapshot CSV files from a workbook that follows the current 2026-style layout:

```bash
python3 scripts/create_monthly_snapshots.py \
  --input "Produktivitas Dosen UB 2026_clean.xlsx" \
  --months 2026-01,2026-02,2026-03 \
  --canonical-month 2026-03
```
