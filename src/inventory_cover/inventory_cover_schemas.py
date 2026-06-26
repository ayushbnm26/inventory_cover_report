"""Schemas, constants, and interface contracts for the Inventory Cover engine.

The engine consumes stable backend workbook artifacts produced by the source
pipelines. The sheet names and column headers declared here are the *interface
contract*: the engine reads these and does not depend on source-pipeline
internals.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any

from inventory_cover.schemas import json_safe


PIPELINE_NAME = "Final Inventory Cover Calculation Engine"

# ---------------------------------------------------------------------------
# Source interface contract: backend workbook sheet names.
# ---------------------------------------------------------------------------
SALES_SHEET = "Sales_Master"
INVENTORY_SHEET = "Inventory_Master"
B2B_SHEET = "B2B_Dispatch_Master"
PO_SHEET = "PO_Items_Master"

SOURCE_PO = "PO Items"
SOURCE_B2B = "B2B Dispatch"
SOURCE_SALES = "Sales"
SOURCE_INVENTORY = "Inventory"
SOURCE_ASIN_MASTER = "ASIN Master"

# ---------------------------------------------------------------------------
# Team-facing report column contract.
# ---------------------------------------------------------------------------
TEAM_REPORT_HEADERS: tuple[str, ...] = (
    "Model Number / SKU",
    "ASIN",
    "Brand",
    "Brand Name",
    "Vendor",
    "Main Category",
    "Sub Category",
    "Aligned DOH Target",
    "Sales Period Start",
    "Sales Period End",
    "Sales Days",
    "Sales Units",
    "Daily Run Rate",
    "Sellable On Hand Units",
    "Amazon In-Transit Units",
    "Own In-Transit Units",
    "Open PO Quantity",
    "Current Stock DOH",
    "Stock + Amazon Transit DOH",
    "Stock + Own Transit DOH",
    "Total Transit DOH",
    "DOC Including Open PO",
    "Total Supply Cover DOH",
    "Target DOH",
    "Gap to Target Units",
    "Cover Bucket",
    "Cover Alert",
    "Remarks",
)

# Backend master keeps diagnostic flags that are intentionally omitted from the
# team workbook.
BACKEND_MASTER_REPORT_HEADERS: tuple[str, ...] = (
    TEAM_REPORT_HEADERS[:-1] + ("Data Quality Flag", "Remarks")
)

# Columns whose values are written into Excel cells as live formulas.
FORMULA_COLUMNS: tuple[str, ...] = (
    "Sales Days",
    "Daily Run Rate",
    "Target DOH",
    "Current Stock DOH",
    "Stock + Amazon Transit DOH",
    "Stock + Own Transit DOH",
    "Total Transit DOH",
    "DOC Including Open PO",
    "Total Supply Cover DOH",
    "Gap to Target Units",
    "Cover Bucket",
    "Cover Alert",
)

# Numeric input columns to which the blank->zero calculation policy applies.
BLANK_TO_ZERO_INPUT_COLUMNS: tuple[str, ...] = (
    "Sales Units",
    "Sellable On Hand Units",
    "Amazon In-Transit Units",
    "Own In-Transit Units",
    "Open PO Quantity",
)

TEXT_COLUMNS: frozenset[str] = frozenset(
    {
        "Model Number / SKU",
        "ASIN",
        "Brand",
        "Brand Name",
        "Vendor",
        "Main Category",
        "Sub Category",
        "Cover Bucket",
        "Cover Alert",
        "Data Quality Flag",
        "Remarks",
    }
)

DATE_COLUMNS: frozenset[str] = frozenset({"Sales Period Start", "Sales Period End"})

INTEGER_COLUMNS: frozenset[str] = frozenset(
    {
        "Sales Days",
        "Sales Units",
        "Sellable On Hand Units",
        "Amazon In-Transit Units",
        "Own In-Transit Units",
        "Open PO Quantity",
        "Gap to Target Units",
    }
)

DECIMAL_COLUMNS: frozenset[str] = frozenset(
    {
        "Aligned DOH Target",
        "Daily Run Rate",
        "Current Stock DOH",
        "Stock + Amazon Transit DOH",
        "Stock + Own Transit DOH",
        "Total Transit DOH",
        "DOC Including Open PO",
        "Total Supply Cover DOH",
        "Target DOH",
    }
)

# ---------------------------------------------------------------------------
# Cover bucket / alert definitions.
# ---------------------------------------------------------------------------
NO_SALES_LABEL = "No Sales"
NO_SALES_ALERT = "No sales in selected period"

COVER_BUCKETS: tuple[str, ...] = (
    "Critical",
    "High Risk",
    "Watch",
    "Near Target",
    "Healthy",
    NO_SALES_LABEL,
)

# Bucket -> annexure worksheet name.
BUCKET_SHEET_NAMES: dict[str, str] = {
    "Critical": "Critical",
    "High Risk": "High_Risk",
    "Watch": "Watch",
    "Near Target": "Near_Target",
    "Healthy": "Healthy",
    NO_SALES_LABEL: "No_Sales",
}

COVER_ALERTS: dict[str, str] = {
    "Critical": "Immediate action",
    "High Risk": "Urgent replenishment",
    "Watch": "Plan replenishment",
    "Near Target": "Monitor",
    "Healthy": "No immediate action",
    NO_SALES_LABEL: NO_SALES_ALERT,
}

# ---------------------------------------------------------------------------
# Backend audit workbook headers.
# ---------------------------------------------------------------------------
MASTER_BACKEND_EXTRA_HEADERS: tuple[str, ...] = (
    "Run ID",
    "Product Key",
    "Product Key Type",
    "Source Presence Flags",
    "Sales Source Rows",
    "Inventory Source Rows",
    "B2B Source Rows",
    "PO Source Rows",
    "ASIN Master Match Status",
    "Identifier Conflict Notes",
    "Calculation Warning Notes",
)

MASTER_BACKEND_HEADERS: tuple[str, ...] = BACKEND_MASTER_REPORT_HEADERS + MASTER_BACKEND_EXTRA_HEADERS

SOURCE_ROW_TRACE_HEADERS: tuple[str, ...] = (
    "Run ID",
    "Product Key",
    "Product Key Type",
    "Final ASIN",
    "Final Model Number / SKU",
    "Source Type",
    "Source File",
    "Source Sheet",
    "Source Row",
    "Source Business Key",
    "Quantity Used",
    "Value Used",
    "Date Used",
    "Trace Notes",
)

CALCULATION_AUDIT_HEADERS: tuple[str, ...] = (
    "Run ID",
    "Product Key",
    "ASIN",
    "Model Number / SKU",
    "Sales Units Raw",
    "Sales Days Raw",
    "Sales Units Used",
    "Sales Days Used",
    "DRR Formula",
    "Stock Units Used",
    "Amazon Transit Units Used",
    "Own Transit Units Used",
    "Open PO Units Used",
    "Target DOH Used",
    "Total Supply Units",
    "Cover Formula Used",
    "Bucket Formula Used",
    "Warnings",
)

VALIDATION_ISSUE_HEADERS: tuple[str, ...] = (
    "Run ID",
    "Severity",
    "Issue Type",
    "Product Key",
    "ASIN",
    "Model Number / SKU",
    "Source Type",
    "Source File",
    "Source Sheet",
    "Source Row",
    "Field Name",
    "Raw Value",
    "Issue Detail",
    "Action Taken",
)

SOURCE_SUMMARY_HEADERS: tuple[str, ...] = (
    "Run ID",
    "Run Timestamp",
    "Source Type",
    "Source Latest Path",
    "Copied Run Path",
    "Workbook Exists",
    "Sheet Used",
    "Rows Read",
    "Rows Accepted",
    "Rows Used In Calculation",
    "Report Period Start",
    "Report Period End",
    "Report Updated Date",
    "Freshness Status",
    "Warnings",
)

FORMULA_GUIDE_HEADERS: tuple[str, ...] = (
    "Section",
    "Column / Topic",
    "Source / Formula",
    "Explanation",
    "Operational Meaning",
)

RUN_METADATA_HEADERS: tuple[str, ...] = ("Key", "Value")


# ---------------------------------------------------------------------------
# Dataclasses.
# ---------------------------------------------------------------------------
@dataclass
class InventoryCoverValidationIssue:
    run_id: str
    severity: str
    issue_type: str
    product_key: str = ""
    asin: str = ""
    model_number: str = ""
    source_type: str = ""
    source_file: str = ""
    source_sheet: str = ""
    source_row: int | None = None
    field_name: str = ""
    raw_value: Any = None
    issue_detail: str = ""
    action_taken: str = ""

    def as_row(self) -> list[Any]:
        return [
            self.run_id,
            self.severity,
            self.issue_type,
            self.product_key,
            self.asin,
            self.model_number,
            self.source_type,
            self.source_file,
            self.source_sheet,
            self.source_row,
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
            "product_key": self.product_key,
            "asin": self.asin,
            "model_number": self.model_number,
            "source_type": self.source_type,
            "source_file": self.source_file,
            "source_sheet": self.source_sheet,
            "source_row": self.source_row,
            "field_name": self.field_name,
            "raw_value": json_safe(self.raw_value),
            "issue_detail": self.issue_detail,
            "action_taken": self.action_taken,
        }


@dataclass
class SourceTraceRecord:
    run_id: str
    product_key: str
    product_key_type: str
    final_asin: str
    final_model: str
    source_type: str
    source_file: str
    source_sheet: str
    source_row: Any
    source_business_key: str
    quantity_used: Any = None
    value_used: Any = None
    date_used: Any = None
    trace_notes: str = ""

    def as_row(self) -> list[Any]:
        return [
            self.run_id,
            self.product_key,
            self.product_key_type,
            self.final_asin,
            self.final_model,
            self.source_type,
            self.source_file,
            self.source_sheet,
            self.source_row,
            self.source_business_key,
            self.quantity_used,
            self.value_used,
            self.date_used,
            self.trace_notes,
        ]


@dataclass
class SourceSummaryRecord:
    run_id: str
    run_timestamp: str
    source_type: str
    source_latest_path: str
    copied_run_path: str = ""
    workbook_exists: bool = False
    sheet_used: str = ""
    rows_read: int = 0
    rows_accepted: int = 0
    rows_used: int = 0
    report_period_start: Any = None
    report_period_end: Any = None
    report_updated_date: Any = None
    freshness_status: str = ""
    warnings: str = ""

    def as_row(self) -> list[Any]:
        return [
            self.run_id,
            self.run_timestamp,
            self.source_type,
            self.source_latest_path,
            self.copied_run_path,
            "Yes" if self.workbook_exists else "No",
            self.sheet_used,
            self.rows_read,
            self.rows_accepted,
            self.rows_used,
            self.report_period_start,
            self.report_period_end,
            self.report_updated_date,
            self.freshness_status,
            self.warnings,
        ]

    def as_json(self) -> dict[str, Any]:
        return {
            "run_id": self.run_id,
            "source_type": self.source_type,
            "source_latest_path": self.source_latest_path,
            "copied_run_path": self.copied_run_path,
            "workbook_exists": self.workbook_exists,
            "sheet_used": self.sheet_used,
            "rows_read": self.rows_read,
            "rows_accepted": self.rows_accepted,
            "rows_used": self.rows_used,
            "report_period_start": json_safe(self.report_period_start),
            "report_period_end": json_safe(self.report_period_end),
            "report_updated_date": json_safe(self.report_updated_date),
            "freshness_status": self.freshness_status,
            "warnings": self.warnings,
        }


@dataclass
class ProductCoverRow:
    """Holds all computed and raw values for one product universe row."""

    # Identity.
    product_key: str
    product_key_type: str
    asin: str = ""
    model_number: str = ""

    # Descriptive fields.
    brand: str = ""
    brand_name: str = ""
    vendor: str = ""
    main_category: str = ""
    sub_category: str = ""

    # Targets.
    aligned_doh_target: float | None = None
    target_doh: float = 30.0

    # Sales window.
    sales_period_start: date | None = None
    sales_period_end: date | None = None
    sales_days: int = 0
    sales_units: float = 0.0
    sales_units_raw: Any = None
    sales_days_raw: Any = None

    # Supply inputs.
    on_hand_units: float = 0.0
    amazon_transit_units: float = 0.0
    own_transit_units: float = 0.0
    open_po_units: float = 0.0

    # Computed outputs (Python mirror of Excel formulas).
    daily_run_rate: float = 0.0
    current_stock_doh: Any = NO_SALES_LABEL
    stock_amazon_doh: Any = NO_SALES_LABEL
    stock_own_doh: Any = NO_SALES_LABEL
    total_transit_doh: Any = NO_SALES_LABEL
    doc_including_open_po: Any = NO_SALES_LABEL
    total_supply_cover_doh: Any = NO_SALES_LABEL
    total_supply_units: float = 0.0
    gap_to_target_units: float = 0.0
    cover_bucket: str = NO_SALES_LABEL
    cover_alert: str = NO_SALES_ALERT
    data_quality_flag: str = "OK"

    # Backend audit detail.
    source_presence: set[str] = field(default_factory=set)
    sales_source_rows: int = 0
    inventory_source_rows: int = 0
    b2b_source_rows: int = 0
    po_source_rows: int = 0
    asin_master_match_status: str = "Not Available"
    identifier_conflict_notes: list[str] = field(default_factory=list)
    calculation_warning_notes: list[str] = field(default_factory=list)

    def source_presence_flags(self) -> str:
        order = [SOURCE_SALES, SOURCE_INVENTORY, SOURCE_B2B, SOURCE_PO, SOURCE_ASIN_MASTER]
        return ", ".join(name for name in order if name in self.source_presence)

    def team_value(self, header: str) -> Any:
        mapping = {
            "Model Number / SKU": self.model_number,
            "ASIN": self.asin,
            "Brand": self.brand,
            "Brand Name": self.brand_name,
            "Vendor": self.vendor,
            "Main Category": self.main_category,
            "Sub Category": self.sub_category,
            "Aligned DOH Target": self.aligned_doh_target if self.aligned_doh_target is not None else self.target_doh,
            "Sales Period Start": self.sales_period_start,
            "Sales Period End": self.sales_period_end,
            "Sales Days": self.sales_days,
            "Sales Units": self.sales_units,
            "Daily Run Rate": round(self.daily_run_rate, 4),
            "Sellable On Hand Units": self.on_hand_units,
            "Amazon In-Transit Units": self.amazon_transit_units,
            "Own In-Transit Units": self.own_transit_units,
            "Open PO Quantity": self.open_po_units,
            "Current Stock DOH": self._round(self.current_stock_doh),
            "Stock + Amazon Transit DOH": self._round(self.stock_amazon_doh),
            "Stock + Own Transit DOH": self._round(self.stock_own_doh),
            "Total Transit DOH": self._round(self.total_transit_doh),
            "DOC Including Open PO": self._round(self.doc_including_open_po),
            "Total Supply Cover DOH": self._round(self.total_supply_cover_doh),
            "Target DOH": self.target_doh,
            "Gap to Target Units": round(self.gap_to_target_units),
            "Cover Bucket": self.cover_bucket,
            "Cover Alert": self.cover_alert,
            "Data Quality Flag": self.data_quality_flag,
            "Remarks": "",
        }
        return mapping.get(header)

    @staticmethod
    def _round(value: Any) -> Any:
        if isinstance(value, (int, float)):
            return round(float(value), 2)
        return value

    def backend_extra_values(self, run_id: str) -> list[Any]:
        return [
            run_id,
            self.product_key,
            self.product_key_type,
            self.source_presence_flags(),
            self.sales_source_rows,
            self.inventory_source_rows,
            self.b2b_source_rows,
            self.po_source_rows,
            self.asin_master_match_status,
            "; ".join(self.identifier_conflict_notes),
            "; ".join(self.calculation_warning_notes),
        ]


@dataclass(frozen=True)
class InventoryCoverPipelineRunResult:
    run_id: str
    run_dir: Path
    team_output_file: Path
    team_latest_file: Path
    backend_output_file: Path
    backend_latest_file: Path
    metadata_file: Path
    validation_file: Path
    log_file: Path
    product_count: int
    validation_issue_count: int
    warning_count: int


def coerce_date(value: Any) -> date | None:
    """Best-effort conversion of a backend cell value to a date."""

    if value is None or (isinstance(value, str) and value.strip() == ""):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    from inventory_cover.utils.date_parsing import parse_vendor_central_date

    result = parse_vendor_central_date(value)
    if result.ok and isinstance(result.value, date):
        return result.value
    return None
