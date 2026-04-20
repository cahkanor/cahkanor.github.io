#!/usr/bin/env python3
"""Update yearly lecturer metrics from the Scopus APIs.

This script keeps the current `author_metrics_yearly.csv` structure and fills:
- `publication_count`: publications in that year only
- `publication_count_total_cum`: cumulative publications up to that year
- `citation_count`: citations in that year only
- `h_index`: current h-index, written to the latest updated year row

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
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any, Iterable

import requests


BASE_DIR = Path(__file__).resolve().parents[1]
DEFAULT_TARGET = BASE_DIR / "data" / "history" / "author_metrics_yearly.csv"
DEFAULT_MASTER = BASE_DIR / "data" / "master" / "lecturers.csv"

AUTHOR_RETRIEVAL_URL = "https://api.elsevier.com/content/author/author_id/{author_id}"
SCOPUS_SEARCH_URL = "https://api.elsevier.com/content/search/scopus"
CITATION_OVERVIEW_URL = "https://api.elsevier.com/content/abstract/citations"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Update yearly metrics in author_metrics_yearly.csv from the Scopus APIs."
    )
    parser.add_argument(
        "--target",
        default=str(DEFAULT_TARGET),
        help=f"Path to target yearly metrics CSV. Default: {DEFAULT_TARGET}",
    )
    parser.add_argument(
        "--master",
        default=str(DEFAULT_MASTER),
        help=f"Path to master lecturers CSV. Default: {DEFAULT_MASTER}",
    )
    parser.add_argument("--start-year", type=int, help="First year to update.")
    parser.add_argument("--end-year", type=int, help="Last year to update.")
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
        "--citation-batch-size",
        type=int,
        default=5,
        help="Citation Overview batch size. Default: 5",
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


def to_int(value: Any) -> int | None:
    text = normalize("" if value is None else str(value))
    if not text:
        return None
    try:
        return int(float(text))
    except ValueError:
        return None


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


def extract_cover_year(entry: dict[str, Any]) -> int | None:
    for key in ("prism:coverDate", "prism:coverDisplayDate", "prism:coverYear"):
        raw = entry.get(key)
        text = normalize(str(to_scalar(raw))) if raw is not None else ""
        if len(text) >= 4 and text[:4].isdigit():
            return int(text[:4])
    return None


def find_scopus_id_in_block(block: Any) -> str:
    if isinstance(block, dict):
        for key in ("scopus-id", "dc:identifier", "identifier", "eid"):
            if key in block:
                value = normalize(str(to_scalar(block[key])))
                if "SCOPUS_ID:" in value:
                    return normalize(value.split("SCOPUS_ID:", 1)[1])
                if value.startswith("2-s2.0-"):
                    return normalize(value.replace("2-s2.0-", "", 1))
                if value.isdigit():
                    return value
        for value in block.values():
            found = find_scopus_id_in_block(value)
            if found:
                return found
    elif isinstance(block, list):
        for item in block:
            found = find_scopus_id_in_block(item)
            if found:
                return found
    return ""


def extract_year_count_pairs(block: Any) -> dict[int, int]:
    pairs: dict[int, int] = {}

    def walk(node: Any) -> None:
        if isinstance(node, dict):
            year = None
            count = None
            for year_key in ("@year", "year", "@_year"):
                if year_key in node:
                    year = to_int(to_scalar(node[year_key]))
                    break
            for count_key in ("$t", "$", "count", "@count", "@_count", "citation-count"):
                if count_key in node:
                    count = to_int(to_scalar(node[count_key]))
                    if count is not None:
                        break

            if year is not None and count is not None:
                pairs[year] = count

            for value in node.values():
                walk(value)
        elif isinstance(node, list):
            for item in node:
                walk(item)

    walk(block)
    return pairs


def find_candidate_citation_blocks(payload: Any) -> list[Any]:
    candidates: list[Any] = []

    def walk(node: Any) -> None:
        if isinstance(node, dict):
            keys = set(node.keys())
            if {"cc", "range", "pcc"}.intersection(keys):
                candidates.append(node)
            for value in node.values():
                walk(value)
        elif isinstance(node, list):
            for item in node:
                walk(item)

    walk(payload)
    return candidates


def parse_citation_overview_payload(payload: dict[str, Any]) -> dict[str, dict[int, int]]:
    by_scopus_id: dict[str, dict[int, int]] = {}
    for block in find_candidate_citation_blocks(payload):
        scopus_id = find_scopus_id_in_block(block)
        if not scopus_id:
            continue
        year_map = extract_year_count_pairs(block)
        if year_map:
            by_scopus_id[scopus_id] = year_map
    return by_scopus_id


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

    def fetch_h_index(self, author_id: str) -> int | None:
        response = self._request(
            AUTHOR_RETRIEVAL_URL.format(author_id=author_id),
            {
                "view": "ENHANCED",
                "httpAccept": "application/json",
            },
        )
        payload = response.json()
        value = recursive_find_first(payload, "h-index")
        return to_int(to_scalar(value))

    def fetch_author_documents(self, author_id: str, count: int) -> list[tuple[str, int]]:
        documents: list[tuple[str, int]] = []
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
                cover_year = extract_cover_year(entry)
                if scopus_id and cover_year is not None:
                    documents.append((scopus_id, cover_year))

            cursor_block = search_results.get("cursor", {})
            cursor = cursor_block.get("@next", "") if isinstance(cursor_block, dict) else ""

        return documents

    def fetch_citation_year_map(
        self,
        scopus_ids: list[str],
        start_year: int,
        end_year: int,
        batch_size: int,
    ) -> tuple[dict[str, dict[int, int]], str]:
        if not scopus_ids:
            return {}, ""

        aggregated: dict[str, dict[int, int]] = {}
        last_error = ""

        def fetch_chunk(ids: list[str]) -> None:
            nonlocal last_error
            if not ids:
                return

            id_param = ",".join(ids)
            try:
                response = self._request(
                    CITATION_OVERVIEW_URL,
                    {
                        "scopus_id": id_param,
                        "start": start_year,
                        "end": end_year,
                        "httpAccept": "application/json",
                    },
                )
            except requests.HTTPError as exc:
                response = exc.response
                body = response.text if response is not None else str(exc)
                if response is not None and response.status_code == 400 and len(ids) > 1:
                    if "Exceeds the maximum number allowed for the service level" in body:
                        midpoint = max(1, len(ids) // 2)
                        fetch_chunk(ids[:midpoint])
                        fetch_chunk(ids[midpoint:])
                        return
                last_error = f"HTTP {response.status_code}: {body}" if response is not None else str(exc)
                return
            except requests.RequestException as exc:
                last_error = str(exc)
                return

            parsed = parse_citation_overview_payload(response.json())
            if parsed:
                aggregated.update(parsed)

        for offset in range(0, len(scopus_ids), batch_size):
            fetch_chunk(scopus_ids[offset : offset + batch_size])

        return aggregated, last_error


def main() -> int:
    args = parse_args()
    target_path = Path(args.target).resolve()
    master_path = Path(args.master).resolve()

    if not target_path.exists():
        raise FileNotFoundError(f"Target file not found: {target_path}")
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

    fieldnames, target_rows = read_csv(target_path)
    _, master_rows = read_csv(master_path)

    master_by_lecturer_id = {
        normalize(row.get("lecturer_id")): row
        for row in master_rows
        if normalize(row.get("lecturer_id"))
    }

    rows_by_lecturer: dict[str, list[dict[str, str]]] = defaultdict(list)
    for row in target_rows:
        lecturer_id = normalize(row.get("lecturer_id"))
        year = to_int(row.get("year"))
        if not lecturer_id or year is None:
            continue
        if args.start_year is not None and year < args.start_year:
            continue
        if args.end_year is not None and year > args.end_year:
            continue
        rows_by_lecturer[lecturer_id].append(row)

    lecturer_ids = sorted(rows_by_lecturer)
    if args.limit > 0:
        lecturer_ids = lecturer_ids[: args.limit]

    total = len(lecturer_ids)
    for index, lecturer_id in enumerate(lecturer_ids, start=1):
        rows = sorted(rows_by_lecturer[lecturer_id], key=lambda item: to_int(item.get("year")) or 0)
        master_row = master_by_lecturer_id.get(lecturer_id, {})
        scopus_author_id = normalize(
            master_row.get("scopus_author_id")
            or master_row.get("SCOPUS ID")
            or rows[0].get("scopus_author_id")
        )

        if not scopus_author_id:
            for row in rows:
                row["api_status"] = "SKIPPED"
                row["api_error"] = "Empty SCOPUS author ID"
                row["data_source"] = "scopus_api"
            print(f"[{index}/{total}] {lecturer_id}: skipped (empty Scopus ID)")
            continue

        start_year = min(to_int(row.get("year")) or 0 for row in rows)
        end_year = max(to_int(row.get("year")) or 0 for row in rows)

        status = "OK"
        error_messages: list[str] = []

        try:
            documents = client.fetch_author_documents(scopus_author_id, count=args.search_count)
        except requests.HTTPError as exc:
            response = exc.response
            message = f"HTTP {response.status_code}: {response.text}" if response is not None else str(exc)
            for row in rows:
                row["api_status"] = "HTTP_ERROR"
                row["api_error"] = message
                row["data_source"] = "scopus_api"
            print(f"[{index}/{total}] {lecturer_id}: HTTP error during document fetch")
            continue
        except requests.RequestException as exc:
            message = str(exc)
            for row in rows:
                row["api_status"] = "REQUEST_ERROR"
                row["api_error"] = message
                row["data_source"] = "scopus_api"
            print(f"[{index}/{total}] {lecturer_id}: request error during document fetch")
            continue

        publication_year_counts = Counter(
            cover_year for _, cover_year in documents if start_year <= cover_year <= end_year
        )
        document_ids = [doc_id for doc_id, _ in documents]

        citation_by_doc, citation_error = client.fetch_citation_year_map(
            scopus_ids=document_ids,
            start_year=start_year,
            end_year=end_year,
            batch_size=max(1, args.citation_batch_size),
        )

        citation_year_counts: Counter[int] = Counter()
        if citation_by_doc:
            for year_map in citation_by_doc.values():
                citation_year_counts.update(year_map)
        elif document_ids and citation_error:
            status = "PARTIAL"
            error_messages.append(citation_error)

        try:
            h_index = client.fetch_h_index(scopus_author_id)
        except requests.RequestException as exc:
            h_index = None
            status = "PARTIAL" if status == "OK" else status
            error_messages.append(f"h-index unavailable: {exc}")

        cumulative_publications = 0
        latest_year = max(to_int(row.get("year")) or 0 for row in rows)

        for row in rows:
            year = to_int(row.get("year")) or 0
            pub_count = publication_year_counts.get(year, 0)
            cumulative_publications += pub_count

            row["publication_count"] = str(pub_count)
            row["publication_count_total_cum"] = str(cumulative_publications)
            row["citation_count"] = (
                str(citation_year_counts.get(year, 0))
                if citation_by_doc or year in citation_year_counts
                else ""
            )
            row["h_index"] = str(h_index) if h_index is not None and year == latest_year else row.get("h_index", "")
            row["scopus_author_id"] = scopus_author_id
            row["data_source"] = "scopus_api"
            row["api_status"] = status
            row["api_error"] = " | ".join(dict.fromkeys(msg for msg in error_messages if msg))

        print(
            f"[{index}/{total}] {lecturer_id}: documents={len(documents)} "
            f"h_index={'' if h_index is None else h_index} status={status}"
        )

    target_rows.sort(
        key=lambda row: (
            normalize(row.get("lecturer_id")),
            to_int(row.get("year")) or 0,
        )
    )
    write_csv(target_path, fieldnames, target_rows)
    print(f"Updated yearly metrics in: {target_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
