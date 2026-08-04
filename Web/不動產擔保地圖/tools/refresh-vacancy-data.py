"""Convert the vacancy/unsold-home workbook into browser-ready data files."""

from __future__ import annotations

import json
import re
import sys
import zipfile
from xml.etree import ElementTree as ET
from pathlib import Path

SOURCE_FILE = "全台空屋率餘屋率總表_空屋率排序.xlsx"
COUNTY_OUTPUT_FILE = "vacancy-rate.js"
DISTRICT_OUTPUT_FILE = "vacancy-district-rate.js"
XML_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PACKAGE_REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def required_column(headers: list[object], pattern: str, description: str) -> str:
    for value in headers:
        if isinstance(value, str) and re.search(pattern, value):
            return value
    raise ValueError(f"Missing {description} column.")


def to_number(value: object) -> int | float:
    if value is None or value == "":
        return 0
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, (int, float)):
        return value
    try:
        return float(str(value).replace(",", ""))
    except ValueError as exc:
        raise ValueError(f"Invalid numeric value: {value!r}") from exc


def cell_column_index(cell_ref: str) -> int:
    column = 0
    for char in re.match(r"[A-Z]+", cell_ref).group(0):
        column = column * 26 + ord(char) - ord("A") + 1
    return column - 1


def cell_value(cell: ET.Element, shared_strings: list[str]) -> object:
    cell_type = cell.get("t")
    if cell_type == "inlineStr":
        return "".join(cell.itertext())
    value = cell.find(f"{XML_NS}v")
    if value is None or value.text is None:
        return ""
    if cell_type == "s":
        return shared_strings[int(value.text)]
    if cell_type == "b":
        return value.text == "1"
    try:
        numeric = float(value.text)
        return int(numeric) if numeric.is_integer() else numeric
    except ValueError:
        return value.text


def workbook_rows(source_path: Path) -> list[list[tuple[object, ...]]]:
    with zipfile.ZipFile(source_path) as archive:
        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            shared_strings = ["".join(item.itertext()) for item in root.findall(f"{XML_NS}si")]

        relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        targets = {
            relation.get("Id"): relation.get("Target")
            for relation in relationships.findall(f"{PACKAGE_REL_NS}Relationship")
        }
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        sheet_paths = []
        for sheet in workbook.findall(f"{XML_NS}sheets/{XML_NS}sheet"):
            target = targets.get(sheet.get(f"{REL_NS}id"))
            if target is None:
                raise ValueError("Workbook sheet relationship is missing.")
            sheet_paths.append("xl/" + target.lstrip("/"))

        result: list[list[tuple[object, ...]]] = []
        for sheet_path in sheet_paths:
            root = ET.fromstring(archive.read(sheet_path))
            rows: list[tuple[object, ...]] = []
            for row in root.findall(f"{XML_NS}sheetData/{XML_NS}row"):
                cells: dict[int, object] = {}
                for cell in row.findall(f"{XML_NS}c"):
                    ref = cell.get("r")
                    if ref:
                        cells[cell_column_index(ref)] = cell_value(cell, shared_strings)
                if cells:
                    rows.append(tuple(cells.get(index, "") for index in range(max(cells) + 1)))
            result.append(rows)
        return result


def load_rows(values: list[tuple[object, ...]]) -> tuple[list[dict[str, object]], str]:
    if not values:
        raise ValueError("Vacancy worksheet is empty.")
    headers = [str(value).strip() if value is not None else "" for value in values[0]]
    header_index = {header: index for index, header in enumerate(headers)}
    county_col = required_column(headers, r"^縣市$", "county")
    district_col = required_column(headers, r"^鄉鎮市區$", "district")
    vacant_col = required_column(headers, r"空屋宅數$", "vacant homes")
    vacancy_rate_col = required_column(headers, r"空屋比率\(%\)$", "vacancy rate")
    unsold_col = required_column(headers, r"待售新成屋宅數$", "unsold homes")
    total_col = required_column(headers, r"^總宅數$", "total homes")
    unsold_rate_col = required_column(headers, r"^待售新成屋比率$", "unsold rate")
    quadrant_col = required_column(headers, r"^象限$", "quadrant")

    def value(row: tuple[object, ...], column: str) -> object:
        return row[header_index[column]] if header_index[column] < len(row) else None

    rows: list[dict[str, object]] = []
    for row in values[1:]:
        county = str(value(row, county_col) or "").strip()
        district = str(value(row, district_col) or "").strip()
        if not county or not district:
            continue
        rows.append({
            "縣市": county,
            "鄉鎮市區": district,
            "vacantHomes": to_number(value(row, vacant_col)),
            "vacancyRate": to_number(value(row, vacancy_rate_col)),
            "unsoldHomes": to_number(value(row, unsold_col)),
            "totalHomes": to_number(value(row, total_col)),
            "unsoldRate": to_number(value(row, unsold_rate_col)),
            "quadrant": str(value(row, quadrant_col) or "").strip(),
        })

    vacancy_period = re.search(r"(\d+H\d)空屋", vacant_col)
    unsold_period = re.search(r"(\d+Q\d)待售", unsold_col)
    if not vacancy_period or not unsold_period:
        raise ValueError("Could not determine vacancy data periods from headers.")
    return rows, f"{vacancy_period.group(1)}／{unsold_period.group(1)}"


def write_js(path: Path, variable: str, rows: list[dict[str, object]], *, period_label: str | None = None) -> None:
    contents: list[str] = []
    if period_label is not None:
        contents.append(f"window.VACANCY_DATA_DATE_LABEL = {json.dumps(period_label, ensure_ascii=False)};")
    contents.append(f"window.{variable} = {json.dumps(rows, ensure_ascii=False, separators=(',', ':'))};")
    path.write_text("\n".join(contents) + "\n", encoding="utf-8", newline="\n")


def main() -> int:
    project_dir = Path(__file__).resolve().parent.parent
    data_dir = project_dir / "data"
    source_path = data_dir / SOURCE_FILE
    if not source_path.exists():
        raise FileNotFoundError(f"Missing vacancy workbook: {source_path}")

    worksheets = workbook_rows(source_path)
    if len(worksheets) < 2:
        raise ValueError("Vacancy workbook must contain county and district worksheets.")
    county_rows, period_label = load_rows(worksheets[0])
    district_rows, district_period_label = load_rows(worksheets[1])
    if period_label != district_period_label:
        raise ValueError("County and district worksheets use different data periods.")
    if not county_rows or not district_rows:
        raise ValueError("Vacancy workbook does not contain usable county and district rows.")

    write_js(data_dir / COUNTY_OUTPUT_FILE, "VACANCY_RATE_ROWS", county_rows, period_label=period_label)
    write_js(data_dir / DISTRICT_OUTPUT_FILE, "VACANCY_DISTRICT_RATE_ROWS", district_rows)
    print(f"Refreshed {COUNTY_OUTPUT_FILE} ({len(county_rows)} rows) and {DISTRICT_OUTPUT_FILE} ({len(district_rows)} rows) from {SOURCE_FILE}: {period_label}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"Vacancy data refresh failed: {exc}", file=sys.stderr)
        raise SystemExit(1)
