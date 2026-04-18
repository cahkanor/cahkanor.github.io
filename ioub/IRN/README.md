# IRN Analytics Dashboard

Static client-side dashboard for Universitas Brawijaya publication, international collaboration, IRN, partner university, partner country, and lecturer collaboration analysis.

## Run

Open `index.html` directly, or serve the folder with any static server.

## Data Source

The dashboard automatically loads these bundled files at startup:

1. `data/SCOPUS ID DATA - UB.xlsx`
2. `data/Publications_at_Brawijaya_University.xlsx`
3. `data/Institution Country File.xlsx`

## Features

- Year range filter and faculty scope filter
- KPI cards for publications, international collaboration, IRN universities, and partner countries
- Charts for yearly trends, domestic vs international mix, top universities, top countries, and top faculties
- Tables for universities, IRN universities, countries, lecturer collaboration, and data quality
- Drill-down detail panel for a selected university or country
- CSV export for filtered datasets

## Assumptions

- A publication belongs to every faculty represented by matched UB lecturer Scopus IDs.
- Partner universities are deduplicated after normalization and UB is excluded from partner counts.
- Country analysis is inferred from institution strings plus alias dictionaries in `app.js`.
- IRN qualification defaults to at least 3 distinct joint publications in the selected IRN window.
