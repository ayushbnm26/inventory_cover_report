"""Shared schemas and constants for PO Items reporting."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any


HEADER_VARIANTS: dict[str, tuple[str, ...]] = {
    "PO": ("po", "purchase order", "purchase order number", "po number"),
    "Vendor Code": ("vendor", "vendor code", "vendor id", "vendor number"),
    "ASIN": ("asin",),
    "External ID": ("external id", "externalid", "external ID", "ean", "upc", "isbn"),
    "External ID Type": ("external id type", "externalidtype", "id type", "external type"),
    "Model Number": ("model number", "model no", "model numebr", "model", "model #"),
    "Title": ("title", "item title", "product title", "description"),
    "Availability": ("availability", "availibility"),
    "Window Type": ("window type", "windowtype"),
    "Window Start": ("window start", "window start date", "start date"),
    "Window End": ("window end", "window end date", "end date"),
    "Expected Date": ("expected date", "expected delivery date", "expected ship date"),
    "Quantity Requested": ("quantity requested", "requested quantity", "qty requested"),
    "Quantity Accepted": ("accepted quantity", "quantity accepted", "qty accepted"),
    "Quantity Received": ("quantity received", "received quantity", "qty received"),
    "Quantity Outstanding": (
        "quantity outstanding",
        "outstanding quantity",
        "qty outstanding",
    ),
    "Unit Cost": ("unit cost", "unit price", "cost"),
    "Total Cost": ("total cost", "extended cost", "line cost"),
    "Ship to location": ("ship to location", "ship-to location", "shipto location"),
    "Backordered": ("backordered", "back ordered", "backorder"),
    "SKU": ("sku", "merchant sku", "vendor sku"),
}

REQUIRED_FIELDS: tuple[str, ...] = (
    "PO",
    "Vendor Code",
    "ASIN",
    "External ID",
    "Model Number",
    "Title",
    "Availability",
    "Window Type",
    "Window Start",
    "Window End",
    "Expected Date",
    "Quantity Requested",
    "Quantity Accepted",
    "Quantity Received",
    "Quantity Outstanding",
    "Unit Cost",
)

OPTIONAL_FIELDS: tuple[str, ...] = (
    "External ID Type",
    "Ship to location",
    "Backordered",
    "SKU",
)

TEXT_FIELDS: frozenset[str] = frozenset(
    {
        "PO",
        "Vendor Code",
        "ASIN",
        "External ID",
        "External ID Type",
        "Model Number",
        "SKU",
        "Ship to location",
        "Title",
        "Availability",
        "Window Type",
        "Backordered",
    }
)

DATE_FIELDS: frozenset[str] = frozenset({"Window Start", "Window End", "Expected Date"})

QUANTITY_FIELDS: frozenset[str] = frozenset(
    {
        "Quantity Requested",
        "Quantity Accepted",
        "Quantity Received",
        "Quantity Outstanding",
    }
)

COST_FIELDS: frozenset[str] = frozenset({"Unit Cost", "Total Cost"})

NUMERIC_FIELDS: frozenset[str] = QUANTITY_FIELDS | COST_FIELDS

MASTER_HEADERS: tuple[str, ...] = (
    "Run ID",
    "Source File",
    "Source Sheet",
    "Source Row",
    "PO",
    "Vendor Code",
    "ASIN",
    "External ID",
    "External ID Type",
    "Ship to location",
    "Model Number",
    "Title",
    "Backordered",
    "Availability",
    "Window Type",
    "Window Start",
    "Window End",
    "Expected Date",
    "Quantity Requested",
    "Quantity Accepted",
    "Quantity Received",
    "Quantity Outstanding",
    "Quantity Received Normalized",
    "Quantity Outstanding Normalized",
    "Open PO Qty - Source",
    "Open PO Qty - Derived",
    "Open PO Qty - Final",
    "Unit Cost",
    "Unit Cost Currency",
    "Total Cost",
    "Total Cost Currency",
    "Open PO Value - Final",
    "Row Validation Status",
    "Row Validation Notes",
)

TEAM_WORKBOOK_HEADERS: tuple[str, ...] = (
    "PO",
    "Vendor Code",
    "ASIN",
    "External ID",
    "External ID Type",
    "Ship to location",
    "Model Number",
    "Title",
    "Backordered",
    "Availability",
    "Window Type",
    "Window Start",
    "Window End",
    "Expected Date",
    "Quantity Requested",
    "Quantity Accepted",
    "Quantity Received",
    "Quantity Outstanding",
    "Quantity Received Normalized",
    "Quantity Outstanding Normalized",
    "Open PO Qty - Source",
    "Open PO Qty - Derived",
    "Open PO Qty - Final",
    "Unit Cost",
    "Unit Cost Currency",
    "Total Cost",
    "Total Cost Currency",
    "Open PO Value - Final",
    "Remarks",
)

RUN_SUMMARY_HEADERS: tuple[str, ...] = ("Metric", "Value")

FILE_AUDIT_HEADERS: tuple[str, ...] = (
    "File Name",
    "Full Path",
    "Copied Run Path",
    "Sheet Used",
    "Header Row Found",
    "Rows Read",
    "Rows Accepted",
    "Rows Rejected",
    "Status",
    "Error/Warning Notes",
)

VALIDATION_ISSUE_HEADERS: tuple[str, ...] = (
    "Severity",
    "Issue Type",
    "Source File",
    "Source Sheet",
    "Source Row",
    "PO",
    "ASIN",
    "Field Name",
    "Raw Value",
    "Issue Detail",
    "Action Taken",
)

DUPLICATE_HEADERS: tuple[str, ...] = (
    "Duplicate Type",
    "Duplicate Group ID",
    "Source File",
    "Source Row",
    "PO",
    "ASIN",
    "Full Row Fingerprint",
    "Action Taken",
)

DUPLICATE_FINGERPRINT_FIELDS: tuple[str, ...] = (
    "PO",
    "Vendor Code",
    "ASIN",
    "External ID",
    "External ID Type",
    "Ship to location",
    "Model Number",
    "Title",
    "Backordered",
    "Availability",
    "Window Type",
    "Window Start",
    "Window End",
    "Expected Date",
    "Quantity Requested",
    "Quantity Accepted",
    "Quantity Received",
    "Quantity Outstanding",
    "Unit Cost",
    "Unit Cost Currency",
    "Total Cost",
    "Total Cost Currency",
)


@dataclass(frozen=True)
class CellValue:
    value: Any
    number_format: str = ""
    data_type: str = ""


@dataclass(frozen=True)
class RawPoItemRow:
    source_file: str
    source_path: Path
    source_sheet: str
    source_row: int
    values: dict[str, CellValue]


@dataclass
class ValidationIssue:
    severity: str
    issue_type: str
    source_file: str = ""
    source_sheet: str = ""
    source_row: int | None = None
    po: str = ""
    asin: str = ""
    field_name: str = ""
    raw_value: Any = None
    issue_detail: str = ""
    action_taken: str = ""

    def as_row(self) -> list[Any]:
        return [
            self.severity,
            self.issue_type,
            self.source_file,
            self.source_sheet,
            self.source_row,
            self.po,
            self.asin,
            self.field_name,
            self.raw_value,
            self.issue_detail,
            self.action_taken,
        ]

    def as_json(self) -> dict[str, Any]:
        return {
            "severity": self.severity,
            "issue_type": self.issue_type,
            "source_file": self.source_file,
            "source_sheet": self.source_sheet,
            "source_row": self.source_row,
            "po": self.po,
            "asin": self.asin,
            "field_name": self.field_name,
            "raw_value": str(self.raw_value) if self.raw_value is not None else None,
            "issue_detail": self.issue_detail,
            "action_taken": self.action_taken,
        }


@dataclass
class NormalizedPoItemRow:
    data: dict[str, Any]
    issues: list[ValidationIssue] = field(default_factory=list)
    fingerprint: str = ""
    dropped: bool = False

    def refresh_validation_status(self) -> None:
        if not self.issues:
            self.data["Row Validation Status"] = "OK"
            self.data["Row Validation Notes"] = ""
            return
        severities = {issue.severity.upper() for issue in self.issues}
        if "ERROR" in severities:
            status = "ERROR"
        else:
            status = "WARNING"
        notes = "; ".join(dict.fromkeys(issue.issue_type for issue in self.issues))
        self.data["Row Validation Status"] = status
        self.data["Row Validation Notes"] = notes

    def as_master_row(self) -> list[Any]:
        self.refresh_validation_status()
        return [self.data.get(header) for header in MASTER_HEADERS]


@dataclass
class FileAuditRecord:
    file_name: str
    full_path: str
    copied_run_path: str = ""
    sheet_used: str = ""
    header_row_found: int | None = None
    rows_read: int = 0
    rows_accepted: int = 0
    rows_rejected: int = 0
    status: str = "PENDING"
    notes: str = ""

    def as_row(self) -> list[Any]:
        return [
            self.file_name,
            self.full_path,
            self.copied_run_path,
            self.sheet_used,
            self.header_row_found,
            self.rows_read,
            self.rows_accepted,
            self.rows_rejected,
            self.status,
            self.notes,
        ]

    def as_json(self) -> dict[str, Any]:
        return {
            "file_name": self.file_name,
            "full_path": self.full_path,
            "copied_run_path": self.copied_run_path,
            "sheet_used": self.sheet_used,
            "header_row_found": self.header_row_found,
            "rows_read": self.rows_read,
            "rows_accepted": self.rows_accepted,
            "rows_rejected": self.rows_rejected,
            "status": self.status,
            "notes": self.notes,
        }


@dataclass
class DuplicateRecord:
    duplicate_type: str
    duplicate_group_id: str
    source_file: str
    source_row: int
    po: str
    asin: str
    full_row_fingerprint: str
    action_taken: str

    def as_row(self) -> list[Any]:
        return [
            self.duplicate_type,
            self.duplicate_group_id,
            self.source_file,
            self.source_row,
            self.po,
            self.asin,
            self.full_row_fingerprint,
            self.action_taken,
        ]

    def as_json(self) -> dict[str, Any]:
        return {
            "duplicate_type": self.duplicate_type,
            "duplicate_group_id": self.duplicate_group_id,
            "source_file": self.source_file,
            "source_row": self.source_row,
            "po": self.po,
            "asin": self.asin,
            "full_row_fingerprint": self.full_row_fingerprint,
            "action_taken": self.action_taken,
        }


@dataclass(frozen=True)
class WorkbookReadResult:
    source_path: Path
    sheet_name: str
    header_row: int
    rows: list[RawPoItemRow]
    warnings: list[str] = field(default_factory=list)


@dataclass(frozen=True)
class PipelineRunResult:
    run_id: str
    run_dir: Path
    output_file: Path
    latest_file: Path
    backend_output_file: Path
    backend_latest_file: Path
    metadata_file: Path
    log_file: Path
    rows_written: int
    validation_issue_count: int
    duplicate_count: int


def json_safe(value: Any) -> Any:
    """Convert common runtime objects into JSON-friendly values."""

    if isinstance(value, Path):
        return str(value)
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, list):
        return [json_safe(item) for item in value]
    if isinstance(value, dict):
        return {str(key): json_safe(item) for key, item in value.items()}
    return value
