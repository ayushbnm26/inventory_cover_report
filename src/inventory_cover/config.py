"""Configuration defaults for local reporting pipelines."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[2]


@dataclass(frozen=True)
class PipelineConfig:
    """Runtime configuration for the PO Items consolidation pipeline."""

    project_root: Path = PROJECT_ROOT
    input_dir: Path = PROJECT_ROOT / "data" / "incoming" / "po_items"
    run_root: Path = PROJECT_ROOT / "runs"
    processed_dir: Path = PROJECT_ROOT / "data" / "processed" / "po_items"
    min_files: int = 2
    max_files: int = 10
    allow_single_file: bool = False
    allow_more_than_max_files: bool = False
    dedupe_exact_rows: bool = False
    log_level: str = "INFO"
    header_scan_rows: int = 30
    default_currency: str = "INR"

    def resolved(self) -> "PipelineConfig":
        """Return a copy with filesystem paths resolved against the project root."""

        def resolve(path: Path) -> Path:
            path = Path(path)
            if path.is_absolute():
                return path
            return (self.project_root / path).resolve()

        return PipelineConfig(
            project_root=Path(self.project_root).resolve(),
            input_dir=resolve(self.input_dir),
            run_root=resolve(self.run_root),
            processed_dir=resolve(self.processed_dir),
            min_files=self.min_files,
            max_files=self.max_files,
            allow_single_file=self.allow_single_file,
            allow_more_than_max_files=self.allow_more_than_max_files,
            dedupe_exact_rows=self.dedupe_exact_rows,
            log_level=self.log_level,
            header_scan_rows=self.header_scan_rows,
            default_currency=self.default_currency,
        )


@dataclass(frozen=True)
class B2BDispatchPipelineConfig:
    """Runtime configuration for the B2B Dispatch Tracker pipeline."""

    project_root: Path = PROJECT_ROOT
    input_dir: Path = PROJECT_ROOT / "data" / "incoming" / "b2b_dispatch"
    run_root: Path = PROJECT_ROOT / "runs"
    processed_dir: Path = PROJECT_ROOT / "data" / "processed" / "b2b_dispatch"
    lookback_days: int = 2
    as_of_date: date | None = None
    source_mode: str = "excel"
    allow_multiple_files: bool = False
    allow_missing_target_sheets: bool = False
    dedupe_exact_rows: bool = False
    log_level: str = "INFO"
    header_scan_rows: int = 30
    value_difference_tolerance: float = 1.0
    google_spreadsheet_id: str = ""
    google_credentials_path: Path = PROJECT_ROOT / "secrets" / "google_oauth" / "credentials.json"
    google_token_path: Path = PROJECT_ROOT / "secrets" / "google_oauth" / "token.json"
    google_readonly_scope: tuple[str, ...] = ("https://www.googleapis.com/auth/spreadsheets.readonly",)
    google_values_max_rows: int = 20000
    google_values_max_column: str = "AZ"

    def resolved(self) -> "B2BDispatchPipelineConfig":
        """Return a copy with filesystem paths resolved against the project root."""

        def resolve(path: Path) -> Path:
            path = Path(path)
            if path.is_absolute():
                return path
            return (self.project_root / path).resolve()

        return B2BDispatchPipelineConfig(
            project_root=Path(self.project_root).resolve(),
            input_dir=resolve(self.input_dir),
            run_root=resolve(self.run_root),
            processed_dir=resolve(self.processed_dir),
            lookback_days=self.lookback_days,
            as_of_date=self.as_of_date,
            source_mode=self.source_mode,
            allow_multiple_files=self.allow_multiple_files,
            allow_missing_target_sheets=self.allow_missing_target_sheets,
            dedupe_exact_rows=self.dedupe_exact_rows,
            log_level=self.log_level,
            header_scan_rows=self.header_scan_rows,
            value_difference_tolerance=self.value_difference_tolerance,
            google_spreadsheet_id=self.google_spreadsheet_id,
            google_credentials_path=resolve(self.google_credentials_path),
            google_token_path=resolve(self.google_token_path),
            google_readonly_scope=tuple(self.google_readonly_scope),
            google_values_max_rows=self.google_values_max_rows,
            google_values_max_column=self.google_values_max_column,
        )


@dataclass(frozen=True)
class SalesInventoryPipelineConfig:
    """Runtime configuration for the Vendor Central Sales & Inventory pipeline."""

    project_root: Path = PROJECT_ROOT
    sales_input_dir: Path = PROJECT_ROOT / "data" / "incoming" / "sales"
    inventory_input_dir: Path = PROJECT_ROOT / "data" / "incoming" / "inventory"
    mapping_input_dir: Path = PROJECT_ROOT / "data" / "reference" / "sales_inventory_mapping"
    run_root: Path = PROJECT_ROOT / "runs"
    processed_dir: Path = PROJECT_ROOT / "data" / "processed" / "sales_inventory"
    require_sales: bool = False
    require_inventory: bool = False
    allow_multiple_sales_files: bool = False
    allow_multiple_inventory_files: bool = False
    dedupe_exact_rows: bool = False
    log_level: str = "INFO"
    header_scan_rows: int = 20

    def resolved(self) -> "SalesInventoryPipelineConfig":
        """Return a copy with filesystem paths resolved against the project root."""

        def resolve(path: Path) -> Path:
            path = Path(path)
            if path.is_absolute():
                return path
            return (self.project_root / path).resolve()

        return SalesInventoryPipelineConfig(
            project_root=Path(self.project_root).resolve(),
            sales_input_dir=resolve(self.sales_input_dir),
            inventory_input_dir=resolve(self.inventory_input_dir),
            mapping_input_dir=resolve(self.mapping_input_dir),
            run_root=resolve(self.run_root),
            processed_dir=resolve(self.processed_dir),
            require_sales=self.require_sales,
            require_inventory=self.require_inventory,
            allow_multiple_sales_files=self.allow_multiple_sales_files,
            allow_multiple_inventory_files=self.allow_multiple_inventory_files,
            dedupe_exact_rows=self.dedupe_exact_rows,
            log_level=self.log_level,
            header_scan_rows=self.header_scan_rows,
        )


@dataclass(frozen=True)
class InventoryCoverPipelineConfig:
    """Runtime configuration for the Final Inventory Cover Calculation Engine.

    This engine consumes the latest backend artifacts produced by the source
    pipelines (PO Items, B2B Dispatch, Sales & Inventory). It does not depend on
    the internal reading logic of those pipelines; the stable backend workbook
    sheet and column headers are the interface contract.
    """

    project_root: Path = PROJECT_ROOT
    po_backend_path: Path = (
        PROJECT_ROOT / "data" / "processed" / "po_items" / "latest" / "PO_Items_Backend_Audit_latest.xlsx"
    )
    b2b_backend_path: Path = (
        PROJECT_ROOT / "data" / "processed" / "b2b_dispatch" / "latest" / "B2B_Dispatch_Backend_Audit_latest.xlsx"
    )
    sales_backend_path: Path = (
        PROJECT_ROOT / "data" / "processed" / "sales_inventory" / "latest" / "Sales_Backend_Audit_latest.xlsx"
    )
    inventory_backend_path: Path = (
        PROJECT_ROOT / "data" / "processed" / "sales_inventory" / "latest" / "Inventory_Backend_Audit_latest.xlsx"
    )
    asin_master_path: Path = PROJECT_ROOT / "data" / "reference" / "master_data" / "ASIN_Master.xlsx"
    run_root: Path = PROJECT_ROOT / "runs"
    processed_dir: Path = PROJECT_ROOT / "data" / "processed" / "inventory_cover"
    sales_window_days: int = 30
    default_target_doh: float = 30.0
    blank_numeric_policy: str = "zero_for_calculation"
    strict_freshness: bool = False
    sales_staleness_days: int = 45
    inventory_staleness_days: int = 7
    log_level: str = "INFO"

    def resolved(self) -> "InventoryCoverPipelineConfig":
        """Return a copy with filesystem paths resolved against the project root."""

        def resolve(path: Path) -> Path:
            path = Path(path)
            if path.is_absolute():
                return path
            return (self.project_root / path).resolve()

        return InventoryCoverPipelineConfig(
            project_root=Path(self.project_root).resolve(),
            po_backend_path=resolve(self.po_backend_path),
            b2b_backend_path=resolve(self.b2b_backend_path),
            sales_backend_path=resolve(self.sales_backend_path),
            inventory_backend_path=resolve(self.inventory_backend_path),
            asin_master_path=resolve(self.asin_master_path),
            run_root=resolve(self.run_root),
            processed_dir=resolve(self.processed_dir),
            sales_window_days=self.sales_window_days,
            default_target_doh=self.default_target_doh,
            blank_numeric_policy=self.blank_numeric_policy,
            strict_freshness=self.strict_freshness,
            sales_staleness_days=self.sales_staleness_days,
            inventory_staleness_days=self.inventory_staleness_days,
            log_level=self.log_level,
        )
