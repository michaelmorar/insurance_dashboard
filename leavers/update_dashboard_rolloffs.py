#!/usr/bin/env python3
"""Set Resource Dashboard Active values to N for monthly tracker roll-offs.

The program reads roll-offs from the Aviva Resource Tracker and updates only the
matching "Active" cells in the selected DATA_MmmYY tab of the Aviva Resource
Dashboard. By default it updates the dashboard file in place.

Examples:
    python update_dashboard_rolloffs.py tracker.xlsx dashboard.xlsx
    python update_dashboard_rolloffs.py tracker.xlsx dashboard.xlsx --month Jul26
    python update_dashboard_rolloffs.py tracker.xlsx dashboard.xlsx --month 2026-02 --output updated_dashboard.xlsx
    python update_dashboard_rolloffs.py tracker.xlsx dashboard.xlsx --month February-2026 --dry-run
"""

from __future__ import annotations

import argparse
import json
import os
import posixpath
import re
import tempfile
import zipfile
from dataclasses import asdict, dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel
from openpyxl.utils import get_column_letter


WORKBOOK_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
MONTH_NUMBERS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}
TRACKER_SHEET_CANDIDATES = ("Change Sourcing", "Rolled Off")
ROLLOFF_COLUMN_CANDIDATES = (
    "Resource Roll Off Date",
    "Resource Roll off",
    "Resource Roll-Off",
    "Date rolled off",
)
NAME_COLUMN_CANDIDATES = ("Candidate Name", "Resource Full Name", "Resource Name", "Name")
ACTIVE_COLUMN_CANDIDATES = ("Active",)


@dataclass
class Update:
    resource: str
    row: int
    old_value: str
    new_value: str = "No"


@dataclass
class RunReport:
    tracker: str
    dashboard: str
    target_month: str
    tracker_sheet: str
    target_sheet: str
    roll_off_count: int
    updated: list[Update]
    already_n: list[str]
    unmatched: list[str]
    output: str | None
    dry_run: bool


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Update a Resource Dashboard DATA tab with tracker roll-offs."
    )
    parser.add_argument("tracker", type=Path, help="Aviva Resource Tracker workbook")
    parser.add_argument("dashboard", type=Path, help="Aviva Resource Dashboard workbook to update")
    parser.add_argument(
        "--month",
        help="Month to process; accepts Jul26, 2026-07, July 2026, or February-2026. Defaults to today.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        help="Write an updated dashboard copy here instead of replacing the dashboard input.",
    )
    parser.add_argument("--dry-run", action="store_true", help="Report the proposed updates without writing a file.")
    parser.add_argument(
        "--tracker-sheet",
        help="Override automatic selection of Change Sourcing / Rolled Off.",
    )
    parser.add_argument(
        "--dashboard-sheet",
        help="Dashboard worksheet to update. Defaults to DATA_MmmYY for --month.",
    )
    return parser.parse_args()


def normalise_header(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").casefold())


def normalise_name(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip()).casefold()


def parse_month(value: str | None) -> tuple[int, int]:
    if not value:
        today = date.today()
        return today.year, today.month
    text = value.strip()
    for fmt in ("%Y-%m", "%Y/%m", "%b%y", "%b %y", "%b %Y", "%B %Y", "%B-%Y", "%b-%Y"):
        try:
            parsed = datetime.strptime(text, fmt)
            return parsed.year, parsed.month
        except ValueError:
            continue
    raise ValueError(f"Cannot parse month '{value}'. Use Jul26, 2026-07, or July 2026.")


def data_sheet_name(year: int, month: int) -> str:
    month_code = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")[month - 1]
    return f"DATA_{month_code}{year % 100:02d}"


def parse_date(value: object, epoch) -> date | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        try:
            parsed = from_excel(value, epoch=epoch)
            return parsed.date() if isinstance(parsed, datetime) else parsed
        except (TypeError, ValueError, OverflowError):
            return None
    text = str(value).strip()
    if not text or text.casefold() in {"n/a", "na", "none", "-"}:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y", "%d.%m.%Y", "%d.%m.%y", "%d-%m-%Y", "%d-%m-%y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def find_header_row_and_columns(ws, required_candidates: dict[str, Iterable[str]]) -> tuple[int, dict[str, int]]:
    for row_number in range(1, min(ws.max_row, 10) + 1):
        headers = {
            normalise_header(cell.value): cell.column
            for cell in ws[row_number]
            if cell.value is not None
        }
        resolved: dict[str, int] = {}
        for key, candidates in required_candidates.items():
            for candidate in candidates:
                column = headers.get(normalise_header(candidate))
                if column:
                    resolved[key] = column
                    break
        if len(resolved) == len(required_candidates):
            return row_number, resolved
    expected = "; ".join(" / ".join(values) for values in required_candidates.values())
    raise ValueError(f"Could not find a header row containing: {expected}")


def choose_tracker_sheet(wb, explicit_name: str | None) -> str:
    if explicit_name:
        if explicit_name not in wb.sheetnames:
            raise ValueError(f"Tracker sheet '{explicit_name}' does not exist.")
        return explicit_name
    for name in TRACKER_SHEET_CANDIDATES:
        if name in wb.sheetnames:
            return name
    raise ValueError("No Change Sourcing or Rolled Off sheet found in the tracker.")


def find_roll_off_names(tracker_path: Path, year: int, month: int, tracker_sheet: str) -> list[str]:
    workbook = load_workbook(tracker_path, read_only=True, data_only=True)
    try:
        worksheet = workbook[tracker_sheet]
        header_row, columns = find_header_row_and_columns(
            worksheet,
            {"name": NAME_COLUMN_CANDIDATES, "rolloff": ROLLOFF_COLUMN_CANDIDATES},
        )
        names: dict[str, str] = {}
        empty_rows = 0
        last_required_column = max(columns.values())
        for row in worksheet.iter_rows(
            min_row=header_row + 1,
            max_col=last_required_column,
            values_only=False,
        ):
            # Some tracker files contain a formatted cell at Excel's final row,
            # which makes max_row 1,048,576. Stop after the real data ends.
            if not any(cell.value is not None for cell in row):
                empty_rows += 1
                if empty_rows >= 100:
                    break
                continue
            empty_rows = 0
            resource = row[columns["name"] - 1].value
            rolloff = parse_date(row[columns["rolloff"] - 1].value, workbook.epoch)
            if resource and rolloff and (rolloff.year, rolloff.month) == (year, month):
                names.setdefault(normalise_name(resource), str(resource).strip())
        return sorted(names.values(), key=str.casefold)
    finally:
        workbook.close()


def dashboard_update_targets(
    dashboard_path: Path,
    year: int,
    month: int,
    names: list[str],
    dashboard_sheet: str | None = None,
) -> tuple[str, int, list[Update], list[str], list[str]]:
    workbook = load_workbook(dashboard_path, read_only=True, data_only=False)
    try:
        sheet_name = dashboard_sheet or data_sheet_name(year, month)
        if sheet_name not in workbook.sheetnames:
            available = ", ".join(workbook.sheetnames)
            raise ValueError(
                f"Dashboard tab '{sheet_name}' does not exist. "
                f"Available tabs: {available}"
            )
        worksheet = workbook[sheet_name]
        header_row, columns = find_header_row_and_columns(
            worksheet,
            {"name": NAME_COLUMN_CANDIDATES, "active": ACTIVE_COLUMN_CANDIDATES},
        )
        rows_by_name: dict[str, list[tuple[int, object]]] = {}
        for row in worksheet.iter_rows(min_row=header_row + 1, values_only=False):
            resource = row[columns["name"] - 1].value
            if resource:
                rows_by_name.setdefault(normalise_name(resource), []).append(
                    (row[0].row, row[columns["active"] - 1].value)
                )
        updates: list[Update] = []
        already_n: list[str] = []
        unmatched: list[str] = []
        for resource in names:
            matches = rows_by_name.get(normalise_name(resource), [])
            if not matches:
                unmatched.append(resource)
                continue
            for row_number, active_value in matches:
                if normalise_name(active_value) == "n":
                    already_n.append(resource)
                else:
                    updates.append(Update(resource=resource, row=row_number, old_value=str(active_value or "")))
        return sheet_name, columns["active"], updates, sorted(set(already_n)), unmatched
    finally:
        workbook.close()


def sheet_part_name(archive: zipfile.ZipFile, sheet_name: str) -> str:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    target_by_id = {rel.attrib["Id"]: rel.attrib["Target"] for rel in relationships.findall(f"{{{PACKAGE_REL_NS}}}Relationship")}
    for sheet in workbook.findall(f".//{{{WORKBOOK_NS}}}sheet"):
        if sheet.attrib.get("name") == sheet_name:
            relationship_id = sheet.attrib[f"{{{REL_NS}}}id"]
            target = target_by_id[relationship_id]
            return posixpath.normpath(posixpath.join("xl", target))
    raise ValueError(f"Could not locate XML for dashboard tab '{sheet_name}'.")


def replace_cell_value(sheet_xml: bytes, coordinate: str, value: str) -> bytes:
    """Replace one existing cell without reserializing the worksheet XML."""
    coordinate_bytes = re.escape(coordinate.encode("ascii"))
    pattern = re.compile(
        rb'<c\b(?=[^>]*\br="' + coordinate_bytes + rb'")'
        rb'(?P<attributes>[^>]*?)(?:/>|>.*?</c>)',
        re.DOTALL,
    )
    matches = list(pattern.finditer(sheet_xml))
    if len(matches) != 1:
        raise ValueError(
            f"Expected exactly one cell at {coordinate}, found {len(matches)}."
        )

    attributes = matches[0].group("attributes")
    attributes = re.sub(rb'\s+t="[^"]*"', b"", attributes)
    encoded_value = escape(value).encode("utf-8")
    replacement = (
        b"<c"
        + attributes
        + b' t="inlineStr"><is><t>'
        + encoded_value
        + b"</t></is></c>"
    )
    return sheet_xml[: matches[0].start()] + replacement + sheet_xml[matches[0].end() :]


def patch_dashboard(source_path: Path, destination_path: Path, sheet_name: str, active_column: int, updates: list[Update]) -> None:
    coordinates = {f"{get_column_letter(active_column)}{update.row}": update.new_value for update in updates}
    with zipfile.ZipFile(source_path, "r") as source:
        target_part = sheet_part_name(source, sheet_name)
        with zipfile.ZipFile(destination_path, "w") as destination:
            for item in source.infolist():
                payload = source.read(item.filename)
                if item.filename == target_part:
                    for coordinate, value in coordinates.items():
                        payload = replace_cell_value(payload, coordinate, value)
                destination.writestr(item, payload)


def run(cli: argparse.Namespace) -> RunReport:
    year, month = parse_month(cli.month)
    if not cli.tracker.is_file():
        raise FileNotFoundError(f"Tracker workbook not found: {cli.tracker}")
    if not cli.dashboard.is_file():
        raise FileNotFoundError(f"Dashboard workbook not found: {cli.dashboard}")
    tracker = load_workbook(cli.tracker, read_only=True, data_only=True)
    try:
        tracker_sheet = choose_tracker_sheet(tracker, cli.tracker_sheet)
    finally:
        tracker.close()
    roll_off_names = find_roll_off_names(cli.tracker, year, month, tracker_sheet)
    target_sheet, active_column, updates, already_n, unmatched = dashboard_update_targets(
        cli.dashboard,
        year,
        month,
        roll_off_names,
        dashboard_sheet=cli.dashboard_sheet,
    )

    output: Path | None = None
    if not cli.dry_run:
        if cli.output:
            output = cli.output
            output.parent.mkdir(parents=True, exist_ok=True)
            patch_dashboard(cli.dashboard, output, target_sheet, active_column, updates)
        else:
            with tempfile.NamedTemporaryFile(dir=cli.dashboard.parent, suffix=".xlsx", delete=False) as temp:
                temporary_path = Path(temp.name)
            try:
                patch_dashboard(cli.dashboard, temporary_path, target_sheet, active_column, updates)
                os.replace(temporary_path, cli.dashboard)
                output = cli.dashboard
            finally:
                if temporary_path.exists():
                    temporary_path.unlink()

    return RunReport(
        tracker=str(cli.tracker),
        dashboard=str(cli.dashboard),
        target_month=f"{year:04d}-{month:02d}",
        tracker_sheet=tracker_sheet,
        target_sheet=target_sheet,
        roll_off_count=len(roll_off_names),
        updated=updates,
        already_n=already_n,
        unmatched=unmatched,
        output=str(output) if output else None,
        dry_run=cli.dry_run,
    )


if __name__ == "__main__":
    report = run(parse_args())
    print(json.dumps(asdict(report), indent=2))
