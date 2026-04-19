#!/usr/bin/env python3
"""Create monthly snapshot CSV files in the current repository format.

This script is designed for workbook layouts like `Produktivitas Dosen UB 2026_clean.xlsx`,
where each target month has a block starting with `Id Scopus <Month><YY>` and then:
- 2020..2025 yearly paper/sitasi pairs
- monthly `Paper` and `Sitasi`
- `ttl ctt ...`
- `ttl paper ...`
- `H-Index`
"""

from __future__ import annotations

import argparse
import csv
import re
import zipfile
from collections import defaultdict
from pathlib import Path
from typing import Any, Iterable
from xml.etree import ElementTree as ET


BASE_DIR = Path(__file__).resolve().parents[1]
DEFAULT_MASTER = BASE_DIR / "data" / "master" / "lecturers.csv"
DEFAULT_OUTPUT_DIR = BASE_DIR / "data" / "snapshots"

NS_MAIN = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
NS_REL_DOC = {"r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
NS_REL_PKG = {"pr": "http://schemas.openxmlformats.org/package/2006/relationships"}

MONTH_ALIASES = {
    "01": ("jan", "januari", "january"),
    "02": ("feb", "februari", "february"),
    "03": ("mar", "maret", "march"),
    "04": ("apr", "april"),
    "05": ("mei", "may"),
    "06": ("jun", "juni", "june"),
    "07": ("jul", "juli", "july"),
    "08": ("agu", "agust", "agustus", "aug", "august"),
    "09": ("sep", "sept", "september"),
    "10": ("okt", "oktober", "oct", "october"),
    "11": ("nov", "november"),
    "12": ("des", "desember", "dec", "december"),
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Create monthly snapshot CSV files from a workbook."
    )
    parser.add_argument("--input", required=True, help="Source workbook path")
    parser.add_argument(
        "--master",
        default=str(DEFAULT_MASTER),
        help=f"Master lecturers CSV path. Default: {DEFAULT_MASTER}",
    )
    parser.add_argument(
        "--output-dir",
        default=str(DEFAULT_OUTPUT_DIR),
        help=f"Snapshot output directory. Default: {DEFAULT_OUTPUT_DIR}",
    )
    parser.add_argument(
        "--months",
        required=True,
        help="Comma-separated months to export, for example 2026-01,2026-02,2026-03",
    )
    parser.add_argument(
        "--canonical-month",
        required=True,
        help="Month whose Scopus ID should be treated as canonical, for example 2026-03",
    )
    return parser.parse_args()


def normalize(value: str | None) -> str:
    return (value or "").strip()


def normalize_key(value: str | None) -> str:
    return re.sub(r"[^0-9A-Za-z]+", "", normalize(value)).upper()


def normalize_label(value: str | None) -> str:
    return re.sub(r"[^a-z0-9]+", "", normalize(value).lower())


def normalize_number_text(value: str | None) -> str:
    text = normalize(value)
    if not text:
        return ""
    match = re.fullmatch(r"([0-9]+)\.0+", text)
    return match.group(1) if match else text


def text_content(node: ET.Element | None) -> str:
    if node is None:
        return ""
    return "".join(part for part in node.itertext() if part).strip()


def parse_cell_value(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t", "")
    if cell_type == "s":
        value = text_content(cell.find("x:v", NS_MAIN))
        return shared_strings[int(value)] if value else ""
    if cell_type == "inlineStr":
        return text_content(cell.find("x:is", NS_MAIN))
    return text_content(cell.find("x:v", NS_MAIN))


def column_letters(cell_ref: str) -> str:
    match = re.match(r"([A-Z]+)", cell_ref)
    return match.group(1) if match else ""


def col_to_num(column: str) -> int:
    result = 0
    for char in column:
        result = result * 26 + (ord(char.upper()) - ord("A") + 1)
    return result


def get_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    try:
        xml_bytes = archive.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ET.fromstring(xml_bytes)
    return [text_content(item) for item in root.findall("x:si", NS_MAIN)]


def get_sheet_targets(archive: zipfile.ZipFile) -> list[tuple[str, str]]:
    workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
    rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))

    rel_map: dict[str, str] = {}
    for rel in rels_root.findall("pr:Relationship", NS_REL_PKG):
        rel_map[rel.attrib["Id"]] = rel.attrib["Target"]

    targets: list[tuple[str, str]] = []
    for sheet in workbook_root.findall("x:sheets/x:sheet", NS_MAIN):
        name = sheet.attrib.get("name", "")
        rel_id = sheet.attrib.get(f"{{{NS_REL_DOC['r']}}}id", "")
        target = rel_map.get(rel_id, "")
        if name and target:
            targets.append((name, f"xl/{target.lstrip('/')}"))
    return targets


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


def to_int_or_blank(value: str) -> int | str:
    text = normalize(value)
    if not text:
        return ""
    try:
        return int(float(text))
    except ValueError:
        return text


def build_master_indexes(master_rows: list[dict[str, str]]) -> tuple[dict[str, dict[str, str]], dict[str, dict[str, str]]]:
    by_scopus: dict[str, dict[str, str]] = {}
    by_nip: dict[str, dict[str, str]] = {}
    for row in master_rows:
        scopus = normalize_key(row.get("scopus_author_id") or row.get("SCOPUS ID"))
        nip = normalize_key(row.get("NIP/NIK") or row.get("NIK/NIP"))
        if scopus:
            by_scopus[scopus] = row
        if nip:
            by_nip[nip] = row
    return by_scopus, by_nip


def get_base_metadata(row_cells: dict[int, str]) -> dict[str, str]:
    return {
        "nip_nik": normalize_number_text(row_cells.get(2)),
        "lecturer_name_workbook": normalize(row_cells.get(3)),
        "homebase": normalize(row_cells.get(4)),
        "status": normalize(row_cells.get(5)),
        "jabatan": normalize(row_cells.get(6)),
        "pangkat": normalize(row_cells.get(7)),
        "pendidikan": normalize(row_cells.get(8)),
        "keterangan": normalize(row_cells.get(9)),
        "jk": normalize(row_cells.get(10)),
        "scopus_id_initial": normalize_number_text(row_cells.get(11)),
    }


def lookup_master_row(
    canonical_scopus_id: str,
    nip_nik: str,
    by_scopus: dict[str, dict[str, str]],
    by_nip: dict[str, dict[str, str]],
) -> tuple[dict[str, str] | None, str]:
    if canonical_scopus_id:
        row = by_scopus.get(normalize_key(canonical_scopus_id))
        if row:
            return row, "scopus_id_canonical"
    if nip_nik:
        row = by_nip.get(normalize_key(nip_nik))
        if row:
            return row, "nip_nik"
    return None, ""


def month_tokens(snapshot_month: str) -> tuple[str, ...]:
    year = snapshot_month[:4]
    month = snapshot_month[5:7]
    yy = year[2:]
    aliases = MONTH_ALIASES[month]
    return tuple(f"{alias}{yy}" for alias in aliases)


def find_month_block_starts(row2_cells: dict[int, str], months: list[str]) -> dict[str, int]:
    starts: dict[str, int] = {}
    labels = {col: normalize_label(text) for col, text in row2_cells.items()}

    for snapshot_month in months:
        tokens = month_tokens(snapshot_month)
        for col, label in labels.items():
            if "idscop" in label and any(token in label for token in tokens):
                starts[snapshot_month] = col
                break
    return starts


def parse_sheet_rows(
    sheet_name: str,
    target: str,
    archive: zipfile.ZipFile,
    shared_strings: list[str],
    months: list[str],
    canonical_month: str,
    master_by_scopus: dict[str, dict[str, str]],
    master_by_nip: dict[str, dict[str, str]],
) -> dict[str, list[dict[str, Any]]]:
    root = ET.fromstring(archive.read(target))
    xml_rows = root.findall("x:sheetData/x:row", NS_MAIN)
    if len(xml_rows) < 4:
        return {}

    parsed_rows: list[dict[int, str]] = []
    for row in xml_rows:
        parsed: dict[int, str] = {}
        for cell in row.findall("x:c", NS_MAIN):
            col = col_to_num(column_letters(cell.attrib.get("r", "")))
            if col:
                parsed[col] = parse_cell_value(cell, shared_strings)
        parsed_rows.append(parsed)

    block_starts = find_month_block_starts(parsed_rows[1], months)
    missing = [month for month in months if month not in block_starts]
    if missing:
        raise ValueError(f"Sheet {sheet_name} is missing month block(s): {', '.join(missing)}")

    snapshots_by_month: dict[str, list[dict[str, Any]]] = defaultdict(list)
    canonical_start = block_starts[canonical_month]

    for row_index, row_cells in enumerate(parsed_rows[3:], start=4):
        base = get_base_metadata(row_cells)
        if not any(base.values()):
            continue

        canonical_scopus_id = normalize_number_text(row_cells.get(canonical_start)) or base["scopus_id_initial"]
        master_row, matched_by = lookup_master_row(
            canonical_scopus_id=canonical_scopus_id,
            nip_nik=base["nip_nik"],
            by_scopus=master_by_scopus,
            by_nip=master_by_nip,
        )

        lecturer_id = normalize(master_row.get("lecturer_id")) if master_row else ""
        lecturer_name = normalize(master_row.get("lecturer_name") or master_row.get("Nama")) if master_row else base["lecturer_name_workbook"]
        faculty_code = normalize(master_row.get("faculty_code") or master_row.get("Homebase")) if master_row else (normalize(base["homebase"]) or sheet_name)
        faculty_name = normalize(master_row.get("faculty_name") or master_row.get("Homebase")) if master_row else (normalize(base["homebase"]) or sheet_name)
        match_status = "matched" if master_row else "unmatched"
        match_key = matched_by or ""

        for snapshot_month in months:
            start_col = block_starts[snapshot_month]
            month_scopus_id = normalize_number_text(row_cells.get(start_col))

            snapshots_by_month[snapshot_month].append(
                {
                    "snapshot_month": snapshot_month,
                    "lecturer_id": lecturer_id,
                    "lecturer_name": lecturer_name,
                    "faculty_code": faculty_code,
                    "faculty_name": faculty_name,
                    "scopus_author_id": canonical_scopus_id,
                    "monthly_scopus_id": month_scopus_id,
                    "nip_nik": base["nip_nik"],
                    "publication_count_month": to_int_or_blank(row_cells.get(start_col + 13, "")),
                    "publication_count_total": to_int_or_blank(row_cells.get(start_col + 16, "")),
                    "citation_count_month": to_int_or_blank(row_cells.get(start_col + 14, "")),
                    "citation_count_total": to_int_or_blank(row_cells.get(start_col + 15, "")),
                    "h_index": to_int_or_blank(row_cells.get(start_col + 17, "")),
                    "api_status": "IMPORTED",
                    "notes": "",
                    "source_sheet": sheet_name,
                    "source_row": row_index,
                    "match_status": match_status,
                    "match_key": match_key,
                }
            )

    return snapshots_by_month


def main() -> int:
    args = parse_args()
    input_path = Path(args.input).resolve()
    master_path = Path(args.master).resolve()
    output_dir = Path(args.output_dir).resolve()
    months = [normalize(month) for month in args.months.split(",") if normalize(month)]
    canonical_month = normalize(args.canonical_month)

    if not input_path.exists():
        raise FileNotFoundError(f"Workbook not found: {input_path}")
    if not master_path.exists():
        raise FileNotFoundError(f"Master lecturers CSV not found: {master_path}")
    if canonical_month not in months:
        raise ValueError("--canonical-month must also be included in --months")

    master_rows = read_csv(master_path)
    master_by_scopus, master_by_nip = build_master_indexes(master_rows)

    all_snapshots: dict[str, list[dict[str, Any]]] = defaultdict(list)
    with zipfile.ZipFile(input_path) as archive:
        shared_strings = get_shared_strings(archive)
        for sheet_name, target in get_sheet_targets(archive):
            sheet_snapshots = parse_sheet_rows(
                sheet_name=sheet_name,
                target=target,
                archive=archive,
                shared_strings=shared_strings,
                months=months,
                canonical_month=canonical_month,
                master_by_scopus=master_by_scopus,
                master_by_nip=master_by_nip,
            )
            for snapshot_month, rows in sheet_snapshots.items():
                all_snapshots[snapshot_month].extend(rows)

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

    total_rows = 0
    for snapshot_month in sorted(all_snapshots):
        rows = sorted(
            all_snapshots[snapshot_month],
            key=lambda row: (
                normalize(row.get("faculty_code")),
                normalize(row.get("lecturer_name")),
                normalize(row.get("nip_nik")),
            ),
        )
        output_path = output_dir / f"author_metrics_{snapshot_month}.csv"
        write_csv(output_path, fieldnames, rows)
        total_rows += len(rows)
        print(f"Wrote {len(rows)} rows to: {output_path}")

    print(f"Processed {total_rows} rows across {len(all_snapshots)} months.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
