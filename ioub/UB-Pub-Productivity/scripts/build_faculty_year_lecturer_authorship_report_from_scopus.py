#!/usr/bin/env python3
"""Build lecturer-level faculty/year authorship workbook directly from Scopus APIs.

The output workbook mirrors the local publication-based report logic:
- one sheet per faculty-year
- one row per lecturer in that faculty
- counts by `first_author_papers`, `corresponding_author_papers`, `co_author_papers`, `total_papers`

Per lecturer, each paper is counted in exactly one category using this precedence:
1. single author OR first author -> `first_author_papers`
2. otherwise corresponding author -> `corresponding_author_papers`
3. otherwise author list member -> `co_author_papers`

Required environment variables:
- SCOPUS_API_KEY

Optional environment variables:
- SCOPUS_INSTTOKEN

Important caveat:
- Corresponding-author detection is inferred from Abstract Retrieval metadata and may vary
  depending on entitlement/view availability in your Elsevier subscription.
"""

from __future__ import annotations

import argparse
import csv
import os
import sys
import time
import xml.etree.ElementTree as ET
from collections import defaultdict
from pathlib import Path
from typing import Any

import requests


BASE_DIR = Path(__file__).resolve().parents[1]
DEFAULT_MASTER = BASE_DIR / "data" / "master" / "lecturers.csv"
DEFAULT_FACULTIES = BASE_DIR / "data" / "master" / "faculties.csv"
DEFAULT_OUTPUT = BASE_DIR / "data" / "derived" / "faculty_year_lecturer_authorship_report_from_scopus.xlsx"

SCOPUS_SEARCH_URL = "https://api.elsevier.com/content/search/scopus"
ABSTRACT_RETRIEVAL_URL = "https://api.elsevier.com/content/abstract/scopus_id/{scopus_id}"

NS = {
    "abstracts": "http://www.elsevier.com/xml/svapi/abstract/dtd",
    "dc": "http://purl.org/dc/elements/1.1/",
    "prism": "http://prismstandard.org/namespaces/basic/2.0/",
    "ce": "http://www.elsevier.com/xml/common/dtd",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build lecturer-level faculty/year authorship workbook directly from Scopus APIs."
    )
    parser.add_argument("--master", default=str(DEFAULT_MASTER), help=f"Master lecturers CSV. Default: {DEFAULT_MASTER}")
    parser.add_argument("--faculties", default=str(DEFAULT_FACULTIES), help=f"Faculties CSV. Default: {DEFAULT_FACULTIES}")
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT), help=f"Output workbook path. Default: {DEFAULT_OUTPUT}")
    parser.add_argument("--start-year", type=int, help="Optional first year to include.")
    parser.add_argument("--end-year", type=int, help="Optional last year to include.")
    parser.add_argument("--limit", type=int, default=0, help="Optional limit of lecturers to process for testing.")
    parser.add_argument("--sleep", type=float, default=0.5, help="Delay after successful requests. Default: 0.5")
    parser.add_argument("--timeout", type=float, default=60.0, help="Request timeout in seconds. Default: 60")
    parser.add_argument("--search-count", type=int, default=25, help="Scopus Search page size. Default: 25")
    parser.add_argument("--max-retries", type=int, default=6, help="Maximum retries for transient/rate-limit errors. Default: 6")
    parser.add_argument("--backoff-seconds", type=float, default=3.0, help="Base retry backoff seconds. Default: 3")
    return parser.parse_args()


def normalize(value: Any) -> str:
    return str(value or "").strip()


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def column_name(index: int) -> str:
    name = ""
    current = index
    while current > 0:
        current -= 1
        name = chr(65 + (current % 26)) + name
        current //= 26
    return name


def xml_escape(value: str) -> str:
    return (
        normalize(value)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def build_grouped_sheet_xml(headers: list[str], data_rows: list[dict[str, Any]], header_rows: list[list[str]], merge_refs: list[str]) -> str:
    rows = header_rows + [[row.get(header, "") for header in headers] for row in data_rows]
    row_xml_parts: list[str] = []

    for row_index, cells in enumerate(rows, start=1):
        cell_xml_parts: list[str] = []
        for col_index, value in enumerate(cells, start=1):
            ref = f"{column_name(col_index)}{row_index}"
            text = normalize(value)
            if text and text.replace(".", "", 1).replace("-", "", 1).isdigit():
                cell_xml_parts.append(f'<c r="{ref}"><v>{text}</v></c>')
            else:
                cell_xml_parts.append(
                    f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">{xml_escape(text)}</t></is></c>'
                )
        row_xml_parts.append(f'<row r="{row_index}">{"".join(cell_xml_parts)}</row>')

    merge_xml = ""
    if merge_refs:
        merge_xml = "<mergeCells count=\"{}\">{}</mergeCells>".format(
            len(merge_refs),
            "".join(f'<mergeCell ref="{ref}"/>' for ref in merge_refs),
        )

    last_col = column_name(len(headers))
    last_row = len(rows)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<dimension ref="A1:{last_col}{last_row}"/>'
        '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        f'<sheetData>{"".join(row_xml_parts)}</sheetData>'
        f"{merge_xml}"
        "</worksheet>"
    )


def write_xlsx(output_path: Path, sheets: list[tuple[str, str]]) -> None:
    import zipfile

    build_dir = output_path.parent / "tmp_scopus_lecturer_xlsx"
    if build_dir.exists():
        for child in sorted(build_dir.rglob("*"), reverse=True):
            if child.is_file():
                child.unlink()
            else:
                child.rmdir()
        build_dir.rmdir()

    (build_dir / "_rels").mkdir(parents=True, exist_ok=True)
    (build_dir / "xl" / "_rels").mkdir(parents=True, exist_ok=True)
    (build_dir / "xl" / "worksheets").mkdir(parents=True, exist_ok=True)

    worksheet_overrides: list[str] = []
    sheet_entries: list[str] = []
    sheet_rels: list[str] = []

    for index, (sheet_name, sheet_xml) in enumerate(sheets, start=1):
        (build_dir / "xl" / "worksheets" / f"sheet{index}.xml").write_text(sheet_xml, encoding="utf-8")
        worksheet_overrides.append(
            f'<Override PartName="/xl/worksheets/sheet{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
        sheet_entries.append(f'<sheet name="{xml_escape(sheet_name)}" sheetId="{index}" r:id="rId{index}"/>')
        sheet_rels.append(
            f'<Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{index}.xml"/>'
        )

    (build_dir / "[Content_Types].xml").write_text(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        + "".join(worksheet_overrides)
        + "</Types>",
        encoding="utf-8",
    )

    (build_dir / "_rels" / ".rels").write_text(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        "</Relationships>",
        encoding="utf-8",
    )

    (build_dir / "xl" / "workbook.xml").write_text(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<sheets>{"".join(sheet_entries)}</sheets></workbook>',
        encoding="utf-8",
    )

    (build_dir / "xl" / "_rels" / "workbook.xml.rels").write_text(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(sheet_rels)
        + "</Relationships>",
        encoding="utf-8",
    )

    if output_path.exists():
        output_path.unlink()
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for file_path in build_dir.rglob("*"):
            if file_path.is_file():
                archive.write(file_path, file_path.relative_to(build_dir))

    for child in sorted(build_dir.rglob("*"), reverse=True):
        if child.is_file():
            child.unlink()
        else:
            child.rmdir()
    build_dir.rmdir()


class ElsevierClient:
    def __init__(self, api_key: str, insttoken: str | None, timeout: float, sleep_seconds: float, max_retries: int, backoff_seconds: float) -> None:
        self.timeout = timeout
        self.sleep_seconds = sleep_seconds
        self.max_retries = max_retries
        self.backoff_seconds = backoff_seconds
        self.session = requests.Session()
        self.session.headers.update({"X-ELS-APIKey": api_key, "Accept": "application/json"})
        if insttoken:
            self.session.headers["X-ELS-Insttoken"] = insttoken

    def _request(self, url: str, *, params: dict[str, Any], accept: str) -> requests.Response:
        headers = {"Accept": accept}
        for attempt in range(self.max_retries + 1):
            response = self.session.get(url, params=params, headers=headers, timeout=self.timeout)
            if response.status_code in (429, 500, 502, 503, 504):
                if attempt >= self.max_retries:
                    response.raise_for_status()
                retry_after = response.headers.get("Retry-After")
                if retry_after:
                    try:
                        wait_seconds = max(float(retry_after), self.backoff_seconds)
                    except ValueError:
                        wait_seconds = self.backoff_seconds * (2**attempt)
                else:
                    wait_seconds = self.backoff_seconds * (2**attempt)
                print(
                    f"Rate/server limit hit ({response.status_code}) for {url}. Waiting {wait_seconds:.1f}s before retry...",
                    file=sys.stderr,
                )
                time.sleep(wait_seconds)
                continue
            response.raise_for_status()
            if self.sleep_seconds:
                time.sleep(self.sleep_seconds)
            return response
        raise RuntimeError("Unreachable retry loop exit")

    def fetch_author_documents(self, author_id: str, count: int) -> list[str]:
        scopus_ids: list[str] = []
        cursor = "*"
        while cursor:
            response = self._request(
                SCOPUS_SEARCH_URL,
                params={
                    "query": f"AU-ID({author_id})",
                    "cursor": cursor,
                    "count": count,
                    "view": "COMPLETE",
                    "httpAccept": "application/json",
                },
                accept="application/json",
            )
            payload = response.json()
            search_results = payload.get("search-results", {})
            entries = search_results.get("entry", []) or []
            if isinstance(entries, dict):
                entries = [entries]
            if not entries:
                break
            for entry in entries:
                for key in ("dc:identifier", "eid"):
                    raw = normalize(entry.get(key))
                    if "SCOPUS_ID:" in raw:
                        scopus_ids.append(raw.split("SCOPUS_ID:", 1)[1].strip())
                        break
                    if raw.startswith("2-s2.0-"):
                        scopus_ids.append(raw.replace("2-s2.0-", "", 1).strip())
                        break
            cursor_block = search_results.get("cursor", {})
            cursor = cursor_block.get("@next", "") if isinstance(cursor_block, dict) else ""
        return list(dict.fromkeys(scopus_ids))

    def fetch_abstract_xml(self, scopus_id: str) -> str:
        response = self._request(
            ABSTRACT_RETRIEVAL_URL.format(scopus_id=scopus_id),
            params={"view": "FULL", "httpAccept": "application/xml"},
            accept="application/xml",
        )
        return response.text


def parse_document(xml_text: str) -> dict[str, Any]:
    root = ET.fromstring(xml_text)
    year_text = normalize(root.findtext(".//prism:coverDate", namespaces=NS))
    if year_text[:4].isdigit():
        year = int(year_text[:4])
    else:
        alt_year = normalize(root.findtext(".//prism:coverDisplayDate", namespaces=NS))
        year = int(alt_year[:4]) if alt_year[:4].isdigit() else None

    authors: list[str] = []
    corresponding_ids: set[str] = set()

    for author in root.findall(".//abstracts:authors/abstracts:author", NS):
        author_id = normalize(author.get("auid") or author.get("author-id") or author.findtext(".//dc:identifier", namespaces=NS))
        if author_id.startswith("SCOPUS_ID:"):
            author_id = author_id.split("SCOPUS_ID:", 1)[1].strip()
        if author_id:
            authors.append(author_id)
            corr_attr = normalize(author.get("corresponding-author") or author.get("corresponding"))
            if corr_attr.lower() in {"y", "yes", "true", "1"}:
                corresponding_ids.add(author_id)

        for tag in ("ce:indexed-name", "ce:surname"):
            # just iterate child nodes to find any embedded corresponding flags if present
            pass

    if not authors:
        for author in root.findall(".//author", {}):
            author_id = normalize(author.get("auid") or author.get("author-id"))
            if author_id:
                authors.append(author_id)
                corr_attr = normalize(author.get("corresponding-author") or author.get("corresponding"))
                if corr_attr.lower() in {"y", "yes", "true", "1"}:
                    corresponding_ids.add(author_id)

    if not corresponding_ids:
        for node in root.iter():
            attrs = {key.lower(): normalize(value).lower() for key, value in node.attrib.items()}
            if attrs.get("corresponding-author") in {"y", "yes", "true", "1"} or attrs.get("corresponding") in {"y", "yes", "true", "1"}:
                author_id = normalize(node.attrib.get("auid") or node.attrib.get("author-id"))
                if author_id:
                    corresponding_ids.add(author_id)

    title = normalize(root.findtext(".//dc:title", namespaces=NS))
    eid = normalize(root.findtext(".//abstracts:coredata/dc:identifier", namespaces=NS))
    if eid.startswith("SCOPUS_ID:"):
        eid = eid.split("SCOPUS_ID:", 1)[1].strip()
    doi = normalize(root.findtext(".//prism:doi", namespaces=NS))

    return {
        "year": year,
        "title": title,
        "eid": eid,
        "doi": doi,
        "author_ids": authors,
        "first_author_id": authors[0] if authors else "",
        "single_author_id": authors[0] if len(authors) == 1 else "",
        "corresponding_author_ids": sorted(corresponding_ids),
    }


def main() -> int:
    args = parse_args()
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
        max_retries=max(0, args.max_retries),
        backoff_seconds=max(0.1, args.backoff_seconds),
    )

    lecturers = read_csv(Path(args.master).resolve())
    faculties = read_csv(Path(args.faculties).resolve())

    faculty_names = {normalize(row.get("faculty_code")): normalize(row.get("faculty_name")) for row in faculties}
    lecturers_by_faculty: dict[str, list[dict[str, str]]] = defaultdict(list)
    lecturer_by_scopus: dict[str, dict[str, str]] = {}

    for lecturer in lecturers:
        faculty_code = normalize(lecturer.get("faculty_code"))
        lecturer_row = {
            "faculty_code": faculty_code,
            "faculty_name": normalize(lecturer.get("faculty_name")) or faculty_names.get(faculty_code, faculty_code),
            "lecturer_name": normalize(lecturer.get("lecturer_name") or lecturer.get("Nama")),
            "scopus_author_id": normalize(lecturer.get("scopus_author_id") or lecturer.get("SCOPUS ID")),
        }
        lecturers_by_faculty[faculty_code].append(lecturer_row)
        if lecturer_row["scopus_author_id"] and lecturer_row["scopus_author_id"] not in lecturer_by_scopus:
            lecturer_by_scopus[lecturer_row["scopus_author_id"]] = lecturer_row

    for rows in lecturers_by_faculty.values():
        rows.sort(key=lambda row: row["lecturer_name"])

    all_lecturers: list[dict[str, str]] = []
    for faculty_code in sorted(lecturers_by_faculty):
        all_lecturers.extend(lecturers_by_faculty[faculty_code])
    if args.limit > 0:
        all_lecturers = all_lecturers[: args.limit]

    document_cache: dict[str, dict[str, Any]] = {}
    lecturer_docs: dict[str, set[str]] = defaultdict(set)

    total = len(all_lecturers)
    for index, lecturer in enumerate(all_lecturers, start=1):
        scopus_id = lecturer["scopus_author_id"]
        if not scopus_id:
            print(f"[{index}/{total}] {lecturer['lecturer_name']}: skipped (empty Scopus ID)")
            continue
        try:
            doc_ids = client.fetch_author_documents(scopus_id, count=args.search_count)
        except requests.RequestException as exc:
            print(f"[{index}/{total}] {lecturer['lecturer_name']}: request error {exc}", file=sys.stderr)
            continue
        lecturer_docs[scopus_id].update(doc_ids)
        print(f"[{index}/{total}] {lecturer['lecturer_name']}: documents={len(doc_ids)}")

    all_doc_ids = sorted({doc_id for ids in lecturer_docs.values() for doc_id in ids})
    for index, doc_id in enumerate(all_doc_ids, start=1):
        try:
            xml_text = client.fetch_abstract_xml(doc_id)
            parsed = parse_document(xml_text)
            if parsed["year"] is None:
                continue
            if args.start_year is not None and parsed["year"] < args.start_year:
                continue
            if args.end_year is not None and parsed["year"] > args.end_year:
                continue
            document_cache[doc_id] = parsed
        except requests.RequestException as exc:
            print(f"[doc {index}/{len(all_doc_ids)}] {doc_id}: request error {exc}", file=sys.stderr)
        except ET.ParseError as exc:
            print(f"[doc {index}/{len(all_doc_ids)}] {doc_id}: XML parse error {exc}", file=sys.stderr)

    years = sorted({doc["year"] for doc in document_cache.values() if doc.get("year") is not None})
    counts_by_faculty_year: dict[tuple[str, int], dict[str, dict[str, int | str]]] = {}

    def ensure_counter(faculty_code: str, year: int, lecturer_row: dict[str, str]) -> dict[str, int | str]:
        key = (faculty_code, year)
        if key not in counts_by_faculty_year:
            counts_by_faculty_year[key] = {}
        lecturer_key = lecturer_row["scopus_author_id"] or lecturer_row["lecturer_name"]
        if lecturer_key not in counts_by_faculty_year[key]:
            counts_by_faculty_year[key][lecturer_key] = {
                "lecturer_name": lecturer_row["lecturer_name"],
                "scopus_author_id": lecturer_row["scopus_author_id"],
                "first_author_papers": 0,
                "corresponding_author_papers": 0,
                "co_author_papers": 0,
                "total_papers": 0,
            }
        return counts_by_faculty_year[key][lecturer_key]

    for faculty_code, rows in lecturers_by_faculty.items():
        for year in years:
            for lecturer_row in rows:
                ensure_counter(faculty_code, year, lecturer_row)

    for doc in document_cache.values():
        year = doc["year"]
        author_ids = [author_id for author_id in doc["author_ids"] if author_id in lecturer_by_scopus]
        for scopus_id in dict.fromkeys(author_ids):
            lecturer_row = lecturer_by_scopus[scopus_id]
            counter = ensure_counter(lecturer_row["faculty_code"], year, lecturer_row)
            if doc["single_author_id"] == scopus_id or doc["first_author_id"] == scopus_id:
                counter["first_author_papers"] = int(counter["first_author_papers"]) + 1
            elif scopus_id in doc["corresponding_author_ids"]:
                counter["corresponding_author_papers"] = int(counter["corresponding_author_papers"]) + 1
            else:
                counter["co_author_papers"] = int(counter["co_author_papers"]) + 1
            counter["total_papers"] = int(counter["total_papers"]) + 1

    sheets: list[tuple[str, str]] = []
    for faculty_code in sorted(lecturers_by_faculty):
        header_top = ["Lecturer Name", "Scopus Author ID"]
        header_bottom = ["", ""]
        headers = ["lecturer_name", "scopus_author_id"]
        merge_refs = ["A1:A2", "B1:B2"]
        current_col = 3
        for year in years:
            header_top.extend([str(year), "", "", ""])
            header_bottom.extend(["First Author", "Corresponding Author", "Co-author", "Total"])
            headers.extend([
                f"{year} First Author",
                f"{year} Corresponding Author",
                f"{year} Co-author",
                f"{year} Total",
            ])
            merge_refs.append(f"{column_name(current_col)}1:{column_name(current_col + 3)}1")
            current_col += 4

        data_rows: list[dict[str, Any]] = []
        for lecturer_row in lecturers_by_faculty[faculty_code]:
            out_row: dict[str, Any] = {
                "lecturer_name": lecturer_row["lecturer_name"],
                "scopus_author_id": lecturer_row["scopus_author_id"],
            }
            for year in years:
                counter = ensure_counter(faculty_code, year, lecturer_row)
                out_row[f"{year} First Author"] = counter["first_author_papers"]
                out_row[f"{year} Corresponding Author"] = counter["corresponding_author_papers"]
                out_row[f"{year} Co-author"] = counter["co_author_papers"]
                out_row[f"{year} Total"] = counter["total_papers"]
            data_rows.append(out_row)

        sheet_name = faculty_code[:31]
        sheet_xml = build_grouped_sheet_xml(headers, data_rows, [header_top, header_bottom], merge_refs)
        sheets.append((sheet_name, sheet_xml))

    output_path = Path(args.output).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    write_xlsx(output_path, sheets)
    print(f"Created {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
