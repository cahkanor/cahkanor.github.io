#!/usr/bin/env python3
"""Update `author_metrics_yearly.csv` from a year-specific source CSV.

Target schema assumptions:
- `year`
- `lecturer_id`
- `publication_count_total_cum`
- `publication_count` (inserted if missing)
- `citation_count`

Source schema assumptions:
- `year`
- `lecturer_id`
- `publication_count_year`
- `citation_count_year`
"""

from __future__ import annotations

import argparse
import csv
from pathlib import Path
from typing import Iterable


BASE_DIR = Path(__file__).resolve().parents[1]
DEFAULT_TARGET = BASE_DIR / "data" / "history" / "author_metrics_yearly.csv"
DEFAULT_SOURCE = BASE_DIR / "data" / "history" / "author_metrics_by_year_2020_2025_from_2026.csv"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Update author_metrics_yearly.csv from a year-specific source CSV."
    )
    parser.add_argument(
        "--target",
        default=str(DEFAULT_TARGET),
        help=f"Path to target yearly metrics CSV. Default: {DEFAULT_TARGET}",
    )
    parser.add_argument(
        "--source",
        default=str(DEFAULT_SOURCE),
        help=f"Path to source yearly-history CSV. Default: {DEFAULT_SOURCE}",
    )
    return parser.parse_args()


def normalize(value: str | None) -> str:
    return (value or "").strip()


def to_int(value: str | None) -> int | None:
    text = normalize(value)
    if not text:
        return None
    return int(float(text))


def read_csv(path: Path) -> tuple[list[str], list[dict[str, str]]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        return list(reader.fieldnames or []), list(reader)


def write_csv(path: Path, fieldnames: Iterable[str], rows: Iterable[dict[str, str]]) -> None:
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(fieldnames))
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def ensure_column(fieldnames: list[str], after_column: str, new_column: str) -> list[str]:
    if new_column in fieldnames:
        return fieldnames
    index = fieldnames.index(after_column) + 1
    return fieldnames[:index] + [new_column] + fieldnames[index:]


def build_source_index(rows: list[dict[str, str]]) -> dict[tuple[str, str], dict[str, str]]:
    index: dict[tuple[str, str], dict[str, str]] = {}
    for row in rows:
        lecturer_id = normalize(row.get("lecturer_id"))
        year = normalize(row.get("year"))
        if lecturer_id and year:
            index[(lecturer_id, year)] = row
    return index


def main() -> int:
    args = parse_args()
    target_path = Path(args.target).resolve()
    source_path = Path(args.source).resolve()

    if not target_path.exists():
        raise FileNotFoundError(f"Target file not found: {target_path}")
    if not source_path.exists():
        raise FileNotFoundError(f"Source file not found: {source_path}")

    fieldnames, target_rows = read_csv(target_path)
    _, source_rows = read_csv(source_path)

    if "publication_count_total_cum" not in fieldnames:
        raise ValueError("Target file must contain 'publication_count_total_cum'.")
    if "citation_count" not in fieldnames:
        raise ValueError("Target file must contain 'citation_count'.")

    fieldnames = ensure_column(fieldnames, "publication_count_total_cum", "publication_count")
    source_index = build_source_index(source_rows)

    # Fill year-specific values from the source file.
    updated = 0
    for row in target_rows:
        key = (normalize(row.get("lecturer_id")), normalize(row.get("year")))
        source_row = source_index.get(key)
        if not source_row:
            continue

        pub_year = normalize(source_row.get("publication_count_year"))
        cite_year = normalize(source_row.get("citation_count_year"))
        row["publication_count"] = pub_year
        row["citation_count"] = cite_year

        data_source = normalize(row.get("data_source"))
        if "2026_workbook_march_history" not in data_source:
            row["data_source"] = f"{data_source};2026_workbook_march_history" if data_source else "2026_workbook_march_history"
        updated += 1

    # Recompute cumulative publication totals per lecturer based on the current
    # file structure. The first row for a lecturer keeps its existing cumulative
    # value if present; later rows use prior cumulative + publication_count.
    grouped: dict[str, list[dict[str, str]]] = {}
    for row in target_rows:
        grouped.setdefault(normalize(row.get("lecturer_id")), []).append(row)

    for lecturer_id, rows in grouped.items():
        rows.sort(key=lambda item: int(normalize(item.get("year"))))
        previous_cumulative: int | None = None

        for row in rows:
            current_cumulative = to_int(row.get("publication_count_total_cum"))
            publication_count = to_int(row.get("publication_count"))

            if previous_cumulative is None:
                if current_cumulative is not None:
                    previous_cumulative = current_cumulative
                elif publication_count is not None:
                    previous_cumulative = publication_count
                    row["publication_count_total_cum"] = str(publication_count)
                continue

            if publication_count is not None:
                previous_cumulative = previous_cumulative + publication_count
                row["publication_count_total_cum"] = str(previous_cumulative)
            elif current_cumulative is not None:
                previous_cumulative = current_cumulative
            else:
                row["publication_count_total_cum"] = str(previous_cumulative)

    target_rows.sort(key=lambda row: (normalize(row.get("lecturer_id")), int(normalize(row.get("year")))))
    write_csv(target_path, fieldnames, target_rows)
    print(f"Updated {updated} rows in: {target_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
