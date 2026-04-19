# Snapshot Files

Put one CSV file per month in this folder.

Recommended naming:
- `author_metrics_YYYY-MM.csv`

Required columns:
- `snapshot_month`
- `lecturer_id`
- `lecturer_name`
- `faculty_code`
- `scopus_author_id`
- `publication_count_total`
- `citation_count_total`
- `h_index`
- `api_status`
- `notes`

Example:

```csv
snapshot_month,lecturer_id,lecturer_name,faculty_code,scopus_author_id,publication_count_total,citation_count_total,h_index,api_status,notes
2026-04,L0001,Prof. Ir. WAYAN FIRDAUS MAHMUDY,FILKOM,55779190600,120,950,18,OK,
2026-04,L0002,Prof. Dr.Eng. FITRI UTAMININGRUM,FILKOM,55488741100,90,720,15,OK,
```
