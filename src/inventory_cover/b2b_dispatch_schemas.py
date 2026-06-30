"""Schemas and constants for the B2B Dispatch Tracker pipeline."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class B2BTargetSheetSpec:
    expected_sheet_name: str
    source_channel: str
    aliases: tuple[str, ...]
    google_sheet_name: str = ""
    google_header_row: int = 0
    google_dispatch_date_column: str = ""


B2B_TARGET_SHEETS: tuple[B2BTargetSheetSpec, ...] = (
    B2BTargetSheetSpec(
        expected_sheet_name="RK PO 007GK",
        source_channel="RK",
        aliases=("RK PO 007GK",),
        google_sheet_name="RK PO 007GK",
        google_header_row=3,
        google_dispatch_date_column="DATE",
    ),
    B2BTargetSheetSpec(
        expected_sheet_name="CLICKTECK DISPATCH",
        source_channel="CLICKTECH",
        aliases=("CLICKTECK DISPATCH", "CLICKTECH DISPATCH"),
        google_sheet_name="CLICKTECK DISPATCH ",
        google_header_row=3,
        google_dispatch_date_column="Date",
    ),
    B2BTargetSheetSpec(
        expected_sheet_name="ETRADE DISPATCH",
        source_channel="ETRADE",
        aliases=("ETRADE DISPATCH",),
        google_sheet_name="ETRADE DISPATCH ",
        google_header_row=4,
        google_dispatch_date_column="Date",
    ),
)

B2B_HEADER_VARIANTS: dict[str, tuple[str, ...]] = {
    "Appointment ID": ("appointment id", "appointmentid", "appt id"),
    "Invoice No": ("invoice no", "invoice number", "invoice", "inv no"),
    "Boxes": ("boxes", "box"),
    "PO": ("po", "purchase order", "purchase order number"),
    "Ship To Location": ("loc.", "loc", "ship to location", "ship-to location", "shipto location"),
    "ASIN": ("asin",),
    "PO ASIN Key": ("po+asin", "po asin", "po asin key", "poasin", "poasinkey"),
    "Model Number": ("sku", "model name", "model number", "model no"),
    "PO Date": ("po date", "purchase order date"),
    "PO Qty": ("po qty", "po quantity", "purchase order qty"),
    "Dispatch Qty": ("dispatch qty", "dispatch quantity", "qty dispatch"),
    "Unit Value": ("unit value", "unit price", "unit rate"),
    "Dispatch Value Source": ("total value", "po dispatch value", "dispatch value"),
    "Dispatch Date": ("date", "dispatch date"),
    "Dispatch Location": ("location", "dispatch location"),
}

B2B_SOURCE_FIELDS: tuple[str, ...] = (
    "Appointment ID",
    "Invoice No",
    "Boxes",
    "PO",
    "Ship To Location",
    "ASIN",
    "PO ASIN Key",
    "Model Number",
    "PO Date",
    "PO Qty",
    "Dispatch Qty",
    "Unit Value",
    "Dispatch Value Source",
    "Dispatch Date",
    "Dispatch Location",
)

B2B_CRITICAL_FIELDS: tuple[str, ...] = (
    "PO",
    "ASIN",
    "Dispatch Qty",
    "Dispatch Date",
)

B2B_RECOMMENDED_FIELDS: tuple[str, ...] = (
    "Invoice No",
    "Model Number",
    "Dispatch Location",
    "Unit Value",
    "Dispatch Value Source",
    "PO ASIN Key",
)

B2B_TEXT_FIELDS: frozenset[str] = frozenset(
    {
        "Appointment ID",
        "Invoice No",
        "PO",
        "Ship To Location",
        "ASIN",
        "PO ASIN Key",
        "Model Number",
        "Dispatch Location",
    }
)

B2B_DATE_FIELDS: frozenset[str] = frozenset({"PO Date", "Dispatch Date"})
B2B_NUMERIC_FIELDS: frozenset[str] = frozenset(
    {"Boxes", "PO Qty", "Dispatch Qty", "Unit Value", "Dispatch Value Source"}
)

B2B_MASTER_HEADERS: tuple[str, ...] = (
    "Run ID",
    "Source File",
    "Source Sheet",
    "Source Row",
    "Source Channel",
    "Appointment ID",
    "Invoice No",
    "Boxes",
    "PO",
    "Ship To Location",
    "ASIN",
    "PO ASIN Key",
    "Model Number",
    "PO Date",
    "PO Qty",
    "Dispatch Qty",
    "Unit Value",
    "Dispatch Value Source",
    "Dispatch Value Derived",
    "Dispatch Value Difference",
    "Dispatch Date",
    "Dispatch Location",
    "Included In Lookback Window",
    "Row Validation Status",
    "Row Validation Notes",
)

B2B_RUN_SUMMARY_HEADERS: tuple[str, ...] = (
    "Run ID",
    "Run timestamp",
    "Input folder",
    "Output folder",
    "As of date",
    "Lookback days",
    "Lookback start date",
    "Lookback end date",
    "Files discovered",
    "Files processed successfully",
    "Files skipped/failed",
    "Target sheets expected",
    "Target sheets found",
    "Target sheets missing",
    "Total source rows scanned",
    "Rows with valid dispatch date",
    "Rows included in lookback window",
    "Rows excluded outside date window",
    "Rows rejected due to invalid critical fields",
    "Rows written",
    "Warning count",
    "Error count",
    "Duplicate count",
    "Output file name",
    "Latest backend file path",
    "Log file path",
)

B2B_SHEET_AUDIT_HEADERS: tuple[str, ...] = (
    "Run ID",
    "Source File",
    "Expected Sheet Name",
    "Actual Sheet Name",
    "Source Channel",
    "Sheet Found",
    "Header Row Found",
    "Rows Scanned",
    "Rows With Valid Dispatch Date",
    "Rows Included",
    "Rows Excluded Outside Window",
    "Rows Rejected",
    "Status",
    "Notes",
)

B2B_VALIDATION_ISSUE_HEADERS: tuple[str, ...] = (
    "Run ID",
    "Severity",
    "Issue Type",
    "Source File",
    "Source Sheet",
    "Source Row",
    "Source Channel",
    "PO",
    "ASIN",
    "Invoice No",
    "Field Name",
    "Raw Value",
    "Issue Detail",
    "Action Taken",
)

B2B_DUPLICATE_HEADERS: tuple[str, ...] = (
    "Run ID",
    "Duplicate Group ID",
    "Duplicate Key",
    "Source File",
    "Source Sheet",
    "Source Row",
    "Action Taken",
)

B2B_DUPLICATE_KEY_FIELDS: tuple[str, ...] = (
    "Source Channel",
    "PO",
    "ASIN",
    "Invoice No",
    "Dispatch Date",
    "Dispatch Qty",
    "Dispatch Location",
)


@dataclass(frozen=True)
class B2BCellValue:
    value: Any
    number_format: str = ""
    data_type: str = ""


@dataclass(frozen=True)
class RawB2BDispatchRow:
    source_file: str
    source_path: Path
    source_sheet: str
    source_row: int
    source_channel: str
    values: dict[str, B2BCellValue]


@dataclass
class B2BValidationIssue:
    run_id: str
    severity: str
    issue_type: str
    source_file: str = ""
    source_sheet: str = ""
    source_row: int | None = None
    source_channel: str = ""
    po: str = ""
    asin: str = ""
    invoice_no: str = ""
    field_name: str = ""
    raw_value: Any = None
    issue_detail: str = ""
    action_taken: str = ""

    def as_row(self) -> list[Any]:
        return [
            self.run_id,
            self.severity,
            self.issue_type,
            self.source_file,
            self.source_sheet,
            self.source_row,
            self.source_channel,
            self.po,
            self.asin,
            self.invoice_no,
            self.field_name,
            self.raw_value,
            self.issue_detail,
            self.action_taken,
        ]

    def as_json(self) -> dict[str, Any]:
        return {
            "run_id": self.run_id,
            "severity": self.severity,
            "issue_type": self.issue_type,
            "source_file": self.source_file,
            "source_sheet": self.source_sheet,
            "source_row": self.source_row,
            "source_channel": self.source_channel,
            "po": self.po,
            "asin": self.asin,
            "invoice_no": self.invoice_no,
            "field_name": self.field_name,
            "raw_value": _json_value(self.raw_value),
            "issue_detail": self.issue_detail,
            "action_taken": self.action_taken,
        }


@dataclass
class NormalizedB2BDispatchRow:
    data: dict[str, Any]
    issues: list[B2BValidationIssue] = field(default_factory=list)
    duplicate_key: str = ""
    dropped: bool = False

    def refresh_validation_status(self) -> None:
        if not self.issues:
            self.data["Row Validation Status"] = "OK"
            self.data["Row Validation Notes"] = ""
            return
        severities = {issue.severity.upper() for issue in self.issues}
        if "ERROR" in severities:
            status = "ERROR"
        elif "WARNING" in severities:
            status = "WARNING"
        else:
            status = "INFO"
        notes = "; ".join(dict.fromkeys(issue.issue_type for issue in self.issues))
        self.data["Row Validation Status"] = status
        self.data["Row Validation Notes"] = notes

    def as_master_row(self) -> list[Any]:
        self.refresh_validation_status()
        return [self.data.get(header) for header in B2B_MASTER_HEADERS]


@dataclass
class B2BSheetAuditRecord:
    run_id: str
    source_file: str
    expected_sheet_name: str
    source_channel: str
    actual_sheet_name: str = ""
    sheet_found: bool = False
    header_row_found: int | None = None
    rows_scanned: int = 0
    rows_with_valid_dispatch_date: int = 0
    rows_included: int = 0
    rows_excluded_outside_window: int = 0
    rows_rejected: int = 0
    status: str = "PENDING"
    notes: str = ""

    def as_row(self) -> list[Any]:
        return [
            self.run_id,
            self.source_file,
            self.expected_sheet_name,
            self.actual_sheet_name,
            self.source_channel,
            self.sheet_found,
            self.header_row_found,
            self.rows_scanned,
            self.rows_with_valid_dispatch_date,
            self.rows_included,
            self.rows_excluded_outside_window,
            self.rows_rejected,
            self.status,
            self.notes,
        ]

    def as_json(self) -> dict[str, Any]:
        return {
            "run_id": self.run_id,
            "source_file": self.source_file,
            "expected_sheet_name": self.expected_sheet_name,
            "actual_sheet_name": self.actual_sheet_name,
            "source_channel": self.source_channel,
            "sheet_found": self.sheet_found,
            "header_row_found": self.header_row_found,
            "rows_scanned": self.rows_scanned,
            "rows_with_valid_dispatch_date": self.rows_with_valid_dispatch_date,
            "rows_included": self.rows_included,
            "rows_excluded_outside_window": self.rows_excluded_outside_window,
            "rows_rejected": self.rows_rejected,
            "status": self.status,
            "notes": self.notes,
        }


@dataclass
class B2BDuplicateRecord:
    run_id: str
    duplicate_group_id: str
    duplicate_key: str
    source_file: str
    source_sheet: str
    source_row: int
    action_taken: str

    def as_row(self) -> list[Any]:
        return [
            self.run_id,
            self.duplicate_group_id,
            self.duplicate_key,
            self.source_file,
            self.source_sheet,
            self.source_row,
            self.action_taken,
        ]

    def as_json(self) -> dict[str, Any]:
        return {
            "run_id": self.run_id,
            "duplicate_group_id": self.duplicate_group_id,
            "duplicate_key": self.duplicate_key,
            "source_file": self.source_file,
            "source_sheet": self.source_sheet,
            "source_row": self.source_row,
            "action_taken": self.action_taken,
        }


@dataclass(frozen=True)
class B2BWorkbookReadResult:
    source_path: Path
    rows: list[RawB2BDispatchRow]
    sheet_audit: list[B2BSheetAuditRecord]


@dataclass(frozen=True)
class B2BRowProcessingResult:
    normalized_row: NormalizedB2BDispatchRow | None
    issues: list[B2BValidationIssue]
    has_valid_dispatch_date: bool
    included_in_window: bool
    rejected: bool
    excluded_outside_window: bool


@dataclass(frozen=True)
class B2BPipelineRunResult:
    run_id: str
    run_dir: Path
    backend_output_file: Path
    backend_latest_file: Path
    metadata_file: Path
    log_file: Path
    rows_written: int
    validation_issue_count: int
    duplicate_count: int


def _json_value(value: Any) -> Any:
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if value is None:
        return None
    return str(value)
