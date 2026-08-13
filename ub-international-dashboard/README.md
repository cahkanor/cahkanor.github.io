# UB International Collaboration Dashboard

A professional, fully static, read-only dashboard for monitoring Universitas Brawijaya international agreements and activities—with or without a formal agreement.

## What is included

- Executive, agreement, activity, expiration, utilization, no-agreement, explorer, data-quality, and printable-report views.
- Query-based partner, agreement, and activity detail views.
- Client-side global search, filters, sorting, pagination, CSV/JSON downloads, copy, and browser print/PDF.
- Responsive UB blue/gold visual design with accessible status text and empty/error states.
- Excel activity template with lookup lists, data validation, instructions, and a flat-table design.
- Python preprocessing and validation that generates optimized local JSON.
- No backend, database server, login, or mandatory external service.

## Required software

- Python 3.10+
- `openpyxl` 3.1+
- Node.js 18+ (only for the small static build copier; there are no npm dependencies)

Install Python requirements:

```bash
python -m pip install -r requirements.txt
```

## Project structure

```text
ub-international-dashboard/
├── data-source/
│   ├── Database_Kerjasama_Luar_Negeri_master.xlsx
│   └── Database_Aktivitas_Kerjasama_Internasional.xlsx
├── public/data/             Generated JSON files
├── scripts/
│   ├── prepare_data.py      Conversion, cleaning, matching, derivation
│   ├── validate_data.py     Concise validation report
│   └── build.mjs            Dependency-free production build
├── src/
│   ├── configuration/       Category colors and navigation
│   ├── services/            Cached JSON loader
│   ├── utilities/           Formatting and export helpers
│   ├── app.js               Static SPA views and interactions
│   └── styles.css           Responsive UB visual system
├── index.html
├── package.json
└── requirements.txt
```

## Activity workbook columns

Each row represents one activity. The workbook includes all 41 requested columns: identity/title/description; category/status/dates; agreement availability and MoU reference; UB faculties, programs, unit and PIC; partner and partner PIC; mode/location/mobility; participant counts and roles; funding; expected/actual outputs and links; evidence; SDGs; remarks; and last-updated date.

Important entry rules:

- Use a unique stable `Activity ID`, such as `ACT-2026-001`.
- Use `|` between multiple faculties, study programs, roles, or SDGs.
- `With Agreement` should normally have a valid `MoU ID`.
- `Without Agreement` may leave MoU ID/name/number blank and is valid—not a quality error.
- Do not publish sensitive personal information in public-facing JSON.

## Update workflow

1. Update the Excel workbooks in `data-source/`.
2. Regenerate JSON:

   ```bash
   python scripts/prepare_data.py
   ```

3. Review validation:

   ```bash
   python scripts/validate_data.py
   ```

4. Build the site:

   ```bash
   npm run build
   ```

The production site is written to `dist/`.

## Run locally

Browsers block `fetch()` from `file://`, so use a local static server:

### Windows — easiest method

Double-click `START_DASHBOARD.cmd`. It starts a local server and opens the dashboard automatically. Do not double-click `index.html` directly.

```bash
npm run dev
```

Open `http://localhost:4173`.

To preview the production build:

```bash
npm run preview
```

## Generated data

`prepare_data.py` creates:

- `mous.json`: normalized agreements and expiration/utilization fields.
- `activities.json`: normalized activities, arrays from pipe-delimited fields, agreement matching, and validation flags.
- `partners.json`: partner-level agreement, activity, participant, faculty, program, and category metrics.
- `countries.json`: country footprint totals.
- `faculties.json`: activity and participant totals by faculty.
- `summary.json`: KPI totals and rates.
- `data-quality.json`: duplicates, missing fields, invalid dates, unknown MoUs, and other checks.

Agreement status, remaining days, expiration windows, duration years, activity year/duration/current status, participant totals, agreement matching, validity during activity, and MoU utilization are derived during preprocessing. Original Excel files are never overwritten.

## Configuration

- Activity categories are stored in `public/data/categories.json`.
- Category badge colors and navigation are in `src/configuration/categories.js`.
- Utilization thresholds are `1` activity for Low, `2–3` for Moderate, and `4+` for High. Change them in `prepare_data.py` and `categories.json`, then regenerate data.

## Deployment

### GitHub Pages

Build locally and publish the contents of `dist/` using a Pages workflow or a `gh-pages` branch. All asset paths are relative, so repository subpaths are supported.

### Netlify

- Build command: `python scripts/prepare_data.py && npm run build`
- Publish directory: `dist`

### Vercel

- Framework preset: Other
- Build command: `python scripts/prepare_data.py && npm run build`
- Output directory: `dist`

If the hosting environment does not include Python/openpyxl, commit the generated `public/data/*.json` and use only `npm run build`.

## Data quality and relationship rules

`MoU ID` is the preferred key. One MoU may have zero or many activities; one activity may have zero or one MoU. Unknown MoU references remain in the activity dataset and are labeled `Agreement Reference Not Found`. Activities explicitly entered without agreements remain separate from invalid references and are included in activity totals.

## Static-site limitations

- The site is intended for viewing, filtering, analysis, export, and reporting.
- Users cannot permanently add or edit records in the browser without an external storage service.
- Data changes must be made in the Excel sources, followed by JSON regeneration and site rebuild/redeployment.
- Document/evidence links must remain accessible to intended users; restricted documents are linked, not embedded.
- Review personal data before public deployment. Public JSON should not contain sensitive information.
- Country dots use a lightweight stylized map suitable for a dependency-free static build; they are analytical markers, not authoritative geographic boundaries.
