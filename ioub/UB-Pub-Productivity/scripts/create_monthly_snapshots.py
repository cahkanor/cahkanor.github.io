#!/usr/bin/env python3
"""Create monthly snapshot CSV files from the Scopus APIs.

This script keeps the current snapshot schema and fills:
- `publication_count_month`
- `publication_count_total`
- `citation_count_month`
- `citation_count_total`
- `h_index`

Expected environment variables:
- `SCOPUS_API_KEY` (required)
- `SCOPUS_INSTTOKEN` (optional)
"""

from __future__ import annotations

import argparse
import csv
import os
import sys
import time
from calendar import monthrange
from pathlib import Path
from typing import Any, Iterable

import requests


BASE_DIR = Path(__file__).resolve().parents[1]
DEFAULT_MASTER = BASE_DIR / "data" / "master" / "lecturers.csv"
DEFAULT_OUTPUT_DIR = BASE_DIR / "data" / "snapshots"

AUTHOR_RETRIEVAL_URL = "https://api.elsevier.com/content/author/author_id/{author_id}"
SCOPUS_SEARCH_URL = "https://api.elsevier.com/content/search/scopus"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Create monthly lecturer snapshots from the Scopus APIs."
    )
    parser.add_argument("--month", required=True, help="Snapshot month in YYYY-MM format.")
    parser.add_argument(
        "--master",
        default=str(DEFAULT_MASTER),
        help=f"Path to master lecturers CSV. Default: {DEFAULT_MASTER}",
    )
    parser.add_argument(
        "--output-dir",
        default=str(DEFAULT_OUTPUT_DIR),
        help=f"Snapshot output directory. Default: {DEFAULT_OUTPUT_DIR}",
    )
    parser.add_argument(
        "--previous-snapshot",
        help="Optional previous snapshot CSV path. If omitted, the script will look for the previous month automatically.",
    )
    parser.add_argument(
        "--sleep",
        type=float,
        default=0.2,
        help="Delay in seconds after successful API requests. Default: 0.2",
    )
    parser.add_argument(
        "--timeout",
        type=float,
        default=45.0,
        help="Request timeout in seconds. Default: 45",
    )
    parser.add_argument(
        "--search-count",
        type=int,
        default=25,
        help="Scopus Search page size. Keep this modest to avoid service-level errors. Default: 25",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=0,
        help="Optional limit for lecturers to process, useful for testing.",
    )
    return parser.parse_args()


def normalize(value: str | None) -> str:
    return (value or "").strip()


def normalize_key(value: str | None) -> str:
    return "".join(char for char in normalize(value).upper() if char.isalnum())


def to_int_or_blank(value: int | None) -> int | str:
    return "" if value is None else value


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def write_csv(path: Path, fieldnames: Iterable[str], rows: Iterable[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(fieldnames))
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def recursive_find_first(payload: Any, target_key: str) -> Any:
    if isinstance(payload, dict):
        for key, value in payload.items():
            if key == target_key:
                return value
            found = recursive_find_first(value, target_key)
            if found is not None:
                return found
    elif isinstance(payload, list):
        for item in payload:
            found = recursive_find_first(item, target_key)
            if found is not None:
                return found
    return None


def to_scalar(value: Any) -> Any:
    if isinstance(value, dict):
        if "$" in value:
            return value["$"]
        if "@_fa" in value:
            return value["@_fa"]
    return value


def to_int(value: Any) -> int | None:
    text = normalize("" if value is None else str(value))
    if not text:
        return None
    try:
        return int(float(text))
    except ValueError:
        return None


def extract_scopus_id(entry: dict[str, Any]) -> str:
    for key in ("dc:identifier", "eid"):
        raw = entry.get(key)
        if not raw:
            continue
        text = normalize(str(to_scalar(raw)))
        if "SCOPUS_ID:" in text:
            return normalize(text.split("SCOPUS_ID:", 1)[1])
        if text.startswith("2-s2.0-"):
            return normalize(text.replace("2-s2.0-", "", 1))
    return ""


def extract_cover_date(entry: dict[str, Any]) -> tuple[int | None, int | None]:
    raw = entry.get("prism:coverDate") or entry.get("prism:coverDisplayDate")
    text = normalize(str(to_scalar(raw))) if raw is not None else ""
    if len(text) >= 7 and text[0:4].isdigit() and text[5:7].isdigit():
        return int(text[0:4]), int(text[5:7])
    cover_year = entry.get("prism:coverYear")
    if cover_year is not None:
        year = to_int(to_scalar(cover_year))
        return year, None
    return None, None


def month_parts(snapshot_month: str) -> tuple[int, int]:
    if len(snapshot_month) != 7 or snapshot_month[4] != "-":
        raise ValueError("--month must be in YYYY-MM format")
    year = int(snapshot_month[:4])
    month = int(snapshot_month[5:7])
    if month < 1 or month > 12:
        raise ValueError("--month must be in YYYY-MM format")
    return year, month


def previous_month(snapshot_month: str) -> str:
    year, month = month_parts(snapshot_month)
    if month == 1:
        return f"{year - 1}-12"
    return f"{year:04d}-{month - 1:02d}"


def build_previous_snapshot_index(path: Path | None) -> dict[str, dict[str, str]]:
    if path is None or not path.exists():
        return {}

    index: dict[str, dict[str, str]] = {}
    for row in read_csv(path):
        lecturer_id = normalize(row.get("lecturer_id"))
        scopus_author_id = normalize(row.get("scopus_author_id"))
        key = lecturer_id or f"SCOPUS_{scopus_author_id}" if scopus_author_id else ""
        if key:
            index[key] = row
    return index


class ElsevierClient:
    def __init__(self, api_key: str, insttoken: str | None, timeout: float, sleep_seconds: float) -> None:
        self.timeout = timeout
        self.sleep_seconds = sleep_seconds
        self.session = requests.Session()
        self.session.headers.update(
            {
                "Accept": "application/json",
                "X-ELS-APIKey": api_key,
            }
        )
        if insttoken:
            self.session.headers["X-ELS-Insttoken"] = insttoken

    def _request(self, url: str, params: dict[str, Any]) -> requests.Response:
        response = self.session.get(url, params=params, timeout=self.timeout)
        response.raise_for_status()
        if self.sleep_seconds:
            time.sleep(self.sleep_seconds)
        return response

    def fetch_author_totals(self, author_id: str) -> tuple[int | None, int | None, int | None]:
        response = self._request(
            AUTHOR_RETRIEVAL_URL.format(author_id=author_id),
            {
                "view": "ENHANCED",
                "httpAccept": "application/json",
            },
        )
        payload = response.json()
        document_count = to_int(to_scalar(recursive_find_first(payload, "document-count")))
        citation_count = to_int(to_scalar(recursive_find_first(payload, "citation-count")))
        h_index = to_int(to_scalar(recursive_find_first(payload, "h-index")))
        return document_count, citation_count, h_index

    def fetch_author_documents(self, author_id: str, count: int) -> list[tuple[str, int | None, int | None]]:
        documents: list[tuple[str, int | None, int | None]] = []
        cursor = "*"

        while cursor:
            response = self._request(
                SCOPUS_SEARCH_URL,
                {
                    "query": f"AU-ID({author_id})",
                    "cursor": cursor,
                    "count": count,
                    "view": "COMPLETE",
                    "httpAccept": "application/json",
                },
            )
            payload = response.json()
            search_results = payload.get("search-results", {})
            entries = search_results.get("entry", []) or []

            if isinstance(entries, dict):
                entries = [entries]

            if not entries:
                break

            for entry in entries:
                scopus_id = extract_scopus_id(entry)
                year, month = extract_cover_date(entry)
                if scopus_id:
                    documents.append((scopus_id, year, month))

            cursor_block = search_results.get("cursor", {})
            cursor = cursor_block.get("@next", "") if isinstance(cursor_block, dict) else ""

        return documents


def main() -> int:
    args = parse_args()
    snapshot_month = normalize(args.month)
    snapshot_year, snapshot_month_number = month_parts(snapshot_month)
    month_end_day = monthrange(snapshot_year, snapshot_month_number)[1]

    master_path = Path(args.master).resolve()
    output_dir = Path(args.output_dir).resolve()
    previous_snapshot_path = (
        Path(args.previous_snapshot).resolve()
        if args.previous_snapshot
        else output_dir / f"author_metrics_{previous_month(snapshot_month)}.csv"
    )

    if not master_path.exists():
        raise FileNotFoundError(f"Master lecturers CSV not found: {master_path}")

    api_key = normalize(os.getenv("SCOPUS_API_KEY"))
    if not api_key:
        print("Missing SCOPUS_API_KEY environment variable.", file=sys.stderr)
        return 1

    insttoken = normalize(os.getenv("SCOPUS_INSTTOKEN")) or None
    client = ElsevierClient(
        api_key=api_key,
        insttoken=insttoken,
        timeout=args.timeout,
        sleep_seconds=args.sleep,
    )

    master_rows = read_csv(master_path)
    if args.limit > 0:
        master_rows = master_rows[: args.limit]

    previous_snapshot = build_previous_snapshot_index(previous_snapshot_path)
    rows: list[dict[str, Any]] = []
    total = len(master_rows)

    for index, master_row in enumerate(master_rows, start=1):
        lecturer_id = normalize(master_row.get("lecturer_id"))
        lecturer_name = normalize(master_row.get("lecturer_name") or master_row.get("Nama"))
        faculty_code = normalize(master_row.get("faculty_code") or master_row.get("Homebase"))
        faculty_name = normalize(master_row.get("faculty_name") or master_row.get("Homebase"))
        scopus_author_id = normalize(master_row.get("scopus_author_id") or master_row.get("SCOPUS ID"))
        nip_nik = normalize(master_row.get("NIP/NIK") or master_row.get("NIK/NIP"))
        source_sheet = normalize(master_row.get("source_sheet"))
        source_row = normalize(master_row.get("source_row"))

        row: dict[str, Any] = {
            "snapshot_month": snapshot_month,
            "lecturer_id": lecturer_id,
            "lecturer_name": lecturer_name,
            "faculty_code": faculty_code,
            "faculty_name": faculty_name,
            "scopus_author_id": scopus_author_id,
            "monthly_scopus_id": scopus_author_id,
            "nip_nik": nip_nik,
            "publication_count_month": "",
            "publication_count_total": "",
            "citation_count_month": "",
            "citation_count_total": "",
            "h_index": "",
            "api_status": "",
            "notes": "",
            "source_sheet": source_sheet,
            "source_row": source_row,
            "match_status": "matched" if lecturer_id else "unmatched",
            "match_key": "scopus_author_id" if scopus_author_id else "",
        }

        if not scopus_author_id:
            row["api_status"] = "SKIPPED"
            row["notes"] = "Empty SCOPUS author ID"
            rows.append(row)
            print(f"[{index}/{total}] {lecturer_id or lecturer_name}: skipped (empty Scopus ID)")
            continue

        try:
            documents = client.fetch_author_documents(scopus_author_id, count=args.search_count)
            _, citation_count_total, h_index = client.fetch_author_totals(scopus_author_id)
        except requests.HTTPError as exc:
            response = exc.response
            row["api_status"] = "HTTP_ERROR"
            row["notes"] = f"HTTP {response.status_code}: {response.text}" if response is not None else str(exc)
            rows.append(row)
            print(f"[{index}/{total}] {lecturer_id or lecturer_name}: HTTP error")
            continue
        except requests.RequestException as exc:
            row["api_status"] = "REQUEST_ERROR"
            row["notes"] = str(exc)
            rows.append(row)
            print(f"[{index}/{total}] {lecturer_id or lecturer_name}: request error")
            continue

        publication_count_month = 0
        publication_count_total = 0
        for _, year, month in documents:
            if year is None:
                continue
            if year < snapshot_year or (year == snapshot_year and month is not None and month <= snapshot_month_number):
                if year < snapshot_year or month is None or month <= snapshot_month_number:
                    publication_count_total += 1
            if year == snapshot_year and month == snapshot_month_number:
                publication_count_month += 1

        previous_key = lecturer_id or f"SCOPUS_{scopus_author_id}"
        previous_row = previous_snapshot.get(previous_key, {})
        previous_citation_total = to_int(previous_row.get("citation_count_total"))
        citation_count_month = None
        if citation_count_total is not None and previous_citation_total is not None:
            citation_count_month = citation_count_total - previous_citation_total

        row["publication_count_month"] = publication_count_month
        row["publication_count_total"] = publication_count_total
        row["citation_count_total"] = to_int_or_blank(citation_count_total)
        row["citation_count_month"] = to_int_or_blank(citation_count_month)
        row["h_index"] = to_int_or_blank(h_index)
        row["api_status"] = "OK"
        row["notes"] = f"Snapshot as of {snapshot_year:04d}-{snapshot_month_number:02d}-{month_end_day:02d}"
        rows.append(row)

        print(
            f"[{index}/{total}] {lecturer_id or lecturer_name}: "
            f"pub_month={publication_count_month} pub_total={publication_count_total} "
            f"cit_total={'' if citation_count_total is None else citation_count_total}"
        )

    rows.sort(
        key=lambda row: (
            normalize(row.get("faculty_code")),
            normalize(row.get("lecturer_name")),
            normalize(row.get("nip_nik")),
        )
    )

    fieldnames = [
        "snapshot_month",
        "lecturer_id",
        "lecturer_name",
        "faculty_code",
        "faculty_name",
        "scopus_author_id",
        "monthly_scopus_id",
        "nip_nik",
        "publication_count_month",
        "publication_count_total",
        "citation_count_month",
        "citation_count_total",
        "h_index",
        "api_status",
        "notes",
        "source_sheet",
        "source_row",
        "match_status",
        "match_key",
    ]

    output_path = output_dir / f"author_metrics_{snapshot_month}.csv"
    write_csv(output_path, fieldnames, rows)
    print(f"Wrote {len(rows)} rows to: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
