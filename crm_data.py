from __future__ import annotations

import json
import threading
import zipfile
from collections import Counter
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any
from urllib.parse import unquote
import xml.etree.ElementTree as ET

EXCEL_NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
OWNER_COLUMN = "AC"
STATUS_PRIORITY = {"overdue": 0, "due": 1, "upcoming": 2}


@dataclass
class WorkbookData:
    headers: dict[str, str]
    records: list[dict[str, Any]]


def parse_excel_date(value: Any) -> date | None:
    if value in (None, ""):
        return None

    if isinstance(value, date):
        return value

    text = str(value).strip()
    if not text:
        return None

    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue

    try:
        serial = float(text)
    except ValueError:
        return None

    if serial <= 0:
        return None

    return (datetime(1899, 12, 30) + timedelta(days=serial)).date()


def shift_months(value: date, months: int) -> date:
    month_index = value.month - 1 + months
    year = value.year + month_index // 12
    month = month_index % 12 + 1

    last_day = 31
    while True:
        try:
            return date(year, month, min(value.day, last_day))
        except ValueError:
            last_day -= 1


def classify_status(today_value: date, remind_date: date, renewal_date: date) -> str:
    if today_value > renewal_date:
        return "overdue"
    if today_value >= remind_date:
        return "due"
    return "upcoming"


def choose_owner(cells: dict[str, str]) -> tuple[str, str]:
    value = (cells.get(OWNER_COLUMN) or "").strip()
    if value:
        return value, OWNER_COLUMN
    return "未分配", "system"


def select_latest_customer_rows(rows: list[tuple[int, dict[str, str]]]) -> list[tuple[int, dict[str, str], date]]:
    latest_by_customer: dict[str, tuple[int, dict[str, str], date]] = {}

    for row_number, cells in rows:
        customer_name = (cells.get("A") or "").strip()
        report_date = parse_excel_date(cells.get("S"))
        if not customer_name or report_date is None:
            continue

        existing = latest_by_customer.get(customer_name)
        if existing is None or (report_date, row_number) >= (existing[2], existing[0]):
            latest_by_customer[customer_name] = (row_number, cells, report_date)

    return list(latest_by_customer.values())


def column_sort_key(column_name: str) -> int:
    total = 0
    for char in column_name:
        total = total * 26 + (ord(char.upper()) - 64)
    return total


def cell_value(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")

    if cell_type == "inlineStr":
        parts = [node.text or "" for node in cell.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")]
        return "".join(parts).strip()

    value_node = cell.find("a:v", EXCEL_NS)
    value = value_node.text.strip() if value_node is not None and value_node.text else ""

    if cell_type == "s" and value:
        return shared_strings[int(value)]

    return value


def read_workbook(excel_path: Path) -> WorkbookData:
    with zipfile.ZipFile(excel_path) as archive:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}

        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in archive.namelist():
            shared_xml = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in shared_xml:
                texts = [node.text or "" for node in item.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")]
                shared_strings.append("".join(texts))

        sheets = workbook.find("a:sheets", EXCEL_NS)
        if sheets is None or not list(sheets):
            raise ValueError("Excel 文件中没有可用工作表。")

        first_sheet = list(sheets)[0]
        target = rel_map[first_sheet.attrib[REL_NS]].lstrip("/")
        sheet_xml = ET.fromstring(archive.read(f"xl/{target}" if not target.startswith("xl/") else target))
        sheet_data = sheet_xml.find("a:sheetData", EXCEL_NS)
        if sheet_data is None:
            raise ValueError("Excel 工作表中没有数据。")

        rows: list[dict[str, str]] = []
        for row in sheet_data.findall("a:row", EXCEL_NS):
            row_values: dict[str, str] = {}
            for cell in row.findall("a:c", EXCEL_NS):
                reference = cell.attrib["r"]
                column = "".join(char for char in reference if char.isalpha())
                row_values[column] = cell_value(cell, shared_strings)
            rows.append(row_values)

    if not rows:
        raise ValueError("Excel 中没有读取到任何行。")

    headers = rows[0]
    return WorkbookData(headers=headers, records=rows[1:])


class ReminderStore:
    def __init__(self, excel_path: str | Path, today_override: date | None = None) -> None:
        self.excel_path = Path(unquote(str(excel_path))).expanduser()
        self.today_override = today_override
        self.lock = threading.Lock()
        self.records: list[dict[str, Any]] = []
        self.recipients: list[str] = []
        self.last_loaded_at: str | None = None
        self.reload()

    @property
    def today(self) -> date:
        return self.today_override or date.today()

    def reload(self) -> None:
        workbook = read_workbook(self.excel_path)
        headers = workbook.headers
        records: list[dict[str, Any]] = []

        latest_rows = select_latest_customer_rows(list(enumerate(workbook.records, start=2)))
        for row_number, cells, report_date in latest_rows:
            customer_name = (cells.get("A") or "").strip()
            owner_name, owner_column = choose_owner(cells)
            renewal_date = shift_months(report_date, 12)
            remind_date = shift_months(renewal_date, -3)
            status = classify_status(self.today, remind_date, renewal_date)

            fields = []
            for column in sorted(headers.keys(), key=column_sort_key):
                label = headers.get(column, column)
                value = (cells.get(column) or "").strip()
                fields.append(
                    {
                        "column": column,
                        "label": label,
                        "value": value,
                    }
                )

            record = {
                "id": str(row_number),
                "row_number": row_number,
                "customer_name": customer_name,
                "project_name": (cells.get("B") or "").strip(),
                "project_type": (cells.get("C") or "").strip(),
                "project_status": (cells.get("E") or "").strip(),
                "project_code": (cells.get("F") or "").strip(),
                "department": (cells.get("G") or "").strip(),
                "owner_name": owner_name,
                "owner_source": owner_column,
                "responsible_person": (cells.get("H") or "").strip(),
                "primary_writer": (cells.get("J") or "").strip(),
                "signer": (cells.get("AC") or "").strip(),
                "branch_company": (cells.get("AD") or "").strip(),
                "report_date": report_date.isoformat(),
                "renewal_date": renewal_date.isoformat(),
                "remind_date": remind_date.isoformat(),
                "status": status,
                "days_until_renewal": (renewal_date - self.today).days,
                "days_until_remind": (remind_date - self.today).days,
                "fields": fields,
                "search_blob": " ".join(
                    filter(
                        None,
                        [
                            customer_name,
                            cells.get("B", ""),
                            cells.get("F", ""),
                            owner_name,
                            cells.get("H", ""),
                            cells.get("J", ""),
                            cells.get("AC", ""),
                            cells.get("AD", ""),
                        ],
                    )
                ).lower(),
            }
            records.append(record)

        records.sort(
            key=lambda item: (
                STATUS_PRIORITY[item["status"]],
                item["renewal_date"],
                item["owner_name"],
                item["customer_name"],
            )
        )

        recipients = sorted({record["owner_name"] for record in records})

        with self.lock:
            self.records = records
            self.recipients = recipients
            self.last_loaded_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def meta(self) -> dict[str, Any]:
        with self.lock:
            recipient_counts = Counter(record["owner_name"] for record in self.records if record["status"] != "upcoming")
            top_recipients = [
                {"name": name, "count": count}
                for name, count in recipient_counts.most_common(8)
            ]
            return {
                "excel_path": str(self.excel_path),
                "today": self.today.isoformat(),
                "last_loaded_at": self.last_loaded_at,
                "total_records": len(self.records),
                "recipients": self.recipients,
                "top_recipients": top_recipients,
            }

    def query(
        self,
        *,
        search: str = "",
        recipient: str = "",
        status: str = "all",
        page: int = 1,
        page_size: int = 50,
    ) -> dict[str, Any]:
        search_text = search.strip().lower()
        recipient_text = recipient.strip()
        status_text = status.strip().lower() or "all"
        page = max(page, 1)
        page_size = min(max(page_size, 1), 100)

        with self.lock:
            filtered = list(self.records)

        if recipient_text:
            filtered = [record for record in filtered if record["owner_name"] == recipient_text]

        if status_text != "all":
            filtered = [record for record in filtered if record["status"] == status_text]

        if search_text:
            filtered = [record for record in filtered if search_text in record["search_blob"]]

        counts = Counter(record["status"] for record in filtered)
        total = len(filtered)
        start = (page - 1) * page_size
        end = start + page_size

        return {
            "filters": {
                "search": search,
                "recipient": recipient,
                "status": status_text,
                "page": page,
                "page_size": page_size,
            },
            "stats": {
                "total": total,
                "overdue": counts.get("overdue", 0),
                "due": counts.get("due", 0),
                "upcoming": counts.get("upcoming", 0),
            },
            "items": [self._summary(record) for record in filtered[start:end]],
            "page_count": max((total + page_size - 1) // page_size, 1),
            "top_recipients": self._top_recipients(filtered),
        }

    def get_record(self, record_id: str) -> dict[str, Any] | None:
        with self.lock:
            for record in self.records:
                if record["id"] == record_id:
                    return json.loads(json.dumps(record))
        return None

    def _summary(self, record: dict[str, Any]) -> dict[str, Any]:
        return {
            "id": record["id"],
            "customer_name": record["customer_name"],
            "project_name": record["project_name"],
            "project_code": record["project_code"],
            "owner_name": record["owner_name"],
            "department": record["department"],
            "report_date": record["report_date"],
            "renewal_date": record["renewal_date"],
            "remind_date": record["remind_date"],
            "status": record["status"],
            "days_until_renewal": record["days_until_renewal"],
            "project_status": record["project_status"],
        }

    def _top_recipients(self, filtered: list[dict[str, Any]]) -> list[dict[str, Any]]:
        counts = Counter(record["owner_name"] for record in filtered if record["status"] != "upcoming")
        return [{"name": name, "count": count} for name, count in counts.most_common(8)]
