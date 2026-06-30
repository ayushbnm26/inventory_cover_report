"""Source acquisition providers for the B2B Dispatch pipeline."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
import shutil
from typing import Any

from inventory_cover.b2b_dispatch_schemas import (
    B2BSheetAuditRecord,
    B2BValidationIssue,
    RawB2BDispatchRow,
)
from inventory_cover.config import B2BDispatchPipelineConfig
from inventory_cover.exceptions import CatastrophicPipelineError, FileValidationError
from inventory_cover.io.b2b_dispatch_excel_io import read_b2b_dispatch_workbook
from inventory_cover.io.b2b_dispatch_file_discovery import discover_b2b_dispatch_files
from inventory_cover.io.b2b_dispatch_google_sheets_io import (
    B2BGoogleSheetsError,
    read_b2b_dispatch_google_sheet,
)


@dataclass(frozen=True)
class B2BDispatchSourceReadResult:
    source_mode: str
    source_label: str
    source_count: int
    source_identifiers: list[str]
    discovered_files: list[Path] = field(default_factory=list)
    failed_files: set[str] = field(default_factory=set)
    rows: list[RawB2BDispatchRow] = field(default_factory=list)
    sheet_audit: list[B2BSheetAuditRecord] = field(default_factory=list)
    validation_issues: list[B2BValidationIssue] = field(default_factory=list)
    metadata: dict[str, Any] = field(default_factory=dict)


class B2BSourceAcquisitionFailure(CatastrophicPipelineError):
    """Source acquisition failure with validation and metadata context."""

    def __init__(
        self,
        message: str,
        *,
        validation_issues: list[B2BValidationIssue] | None = None,
        metadata: dict[str, Any] | None = None,
    ):
        super().__init__(message)
        self.validation_issues = validation_issues or []
        self.metadata = metadata or {}


def acquire_b2b_dispatch_source(
    config: B2BDispatchPipelineConfig,
    run_id: str,
    input_copy_dir: Path,
    logger: Any,
) -> B2BDispatchSourceReadResult:
    """Acquire B2B dispatch rows from the configured source mode."""

    source_mode = normalize_b2b_source_mode(config.source_mode)
    if source_mode == "excel":
        return _acquire_from_excel(config, run_id, input_copy_dir, logger)
    if source_mode == "google_sheets":
        return _acquire_from_google_sheets(config, run_id, logger)
    raise B2BSourceAcquisitionFailure(
        f"Unsupported B2B source mode {config.source_mode!r}. Expected 'excel' or 'google-sheets'.",
        validation_issues=[
            B2BValidationIssue(
                run_id=run_id,
                severity="ERROR",
                issue_type="UNSUPPORTED_B2B_SOURCE_MODE",
                field_name="source_mode",
                raw_value=config.source_mode,
                issue_detail="Expected 'excel' or 'google-sheets'.",
                action_taken="Run failed.",
            )
        ],
        metadata={"b2b_source_mode": config.source_mode},
    )


def normalize_b2b_source_mode(value: str | None) -> str:
    normalized = str(value or "excel").strip().lower().replace("-", "_")
    if normalized in {"google", "sheets", "google_sheet", "google_sheets"}:
        return "google_sheets"
    return normalized or "excel"


def _acquire_from_excel(
    config: B2BDispatchPipelineConfig,
    run_id: str,
    input_copy_dir: Path,
    logger: Any,
) -> B2BDispatchSourceReadResult:
    discovered_files = discover_b2b_dispatch_files(config)
    logger.info("Files discovered: %s", len(discovered_files))
    _copy_inputs(discovered_files, input_copy_dir)

    raw_rows: list[RawB2BDispatchRow] = []
    sheet_audit: list[B2BSheetAuditRecord] = []
    validation_issues: list[B2BValidationIssue] = []
    failed_files: set[str] = set()

    for source_path in discovered_files:
        try:
            read_result = read_b2b_dispatch_workbook(source_path, config, run_id)
            raw_rows.extend(read_result.rows)
            sheet_audit.extend(read_result.sheet_audit)
            found_count = sum(1 for record in read_result.sheet_audit if record.sheet_found)
            missing_count = sum(1 for record in read_result.sheet_audit if not record.sheet_found)
            logger.info(
                "Workbook scanned: %s target_sheets_found=%s target_sheets_missing=%s raw_rows=%s",
                source_path.name,
                found_count,
                missing_count,
                len(read_result.rows),
            )
            for record in read_result.sheet_audit:
                if record.header_row_found:
                    logger.info(
                        "Header row detected: file=%s sheet=%s channel=%s row=%s",
                        record.source_file,
                        record.actual_sheet_name,
                        record.source_channel,
                        record.header_row_found,
                    )
        except FileValidationError as exc:
            failed_files.add(source_path.name)
            validation_issues.append(
                B2BValidationIssue(
                    run_id=run_id,
                    severity="ERROR",
                    issue_type="FILE_OPEN_FAILED",
                    source_file=source_path.name,
                    issue_detail=str(exc),
                    action_taken="File skipped; run continued with other discovered files.",
                )
            )
            logger.error("Skipped unreadable workbook %s: %s", source_path.name, exc)

    return B2BDispatchSourceReadResult(
        source_mode="excel",
        source_label=str(config.input_dir),
        source_count=len(discovered_files),
        source_identifiers=[str(path) for path in discovered_files],
        discovered_files=discovered_files,
        failed_files=failed_files,
        rows=raw_rows,
        sheet_audit=sheet_audit,
        validation_issues=validation_issues,
        metadata={
            "b2b_source_mode": "excel",
            "source_traceability": "local_excel_workbook",
            "input_directory": str(config.input_dir),
        },
    )


def _acquire_from_google_sheets(
    config: B2BDispatchPipelineConfig,
    run_id: str,
    logger: Any,
) -> B2BDispatchSourceReadResult:
    try:
        read_result = read_b2b_dispatch_google_sheet(config, run_id)
    except B2BGoogleSheetsError as exc:
        issue = B2BValidationIssue(
            run_id=run_id,
            severity="ERROR",
            issue_type=exc.issue_type,
            source_file=f"google_sheets:{config.google_spreadsheet_id}" if config.google_spreadsheet_id else "",
            field_name="Google Sheets API",
            raw_value=exc.classification,
            issue_detail=str(exc),
            action_taken="Run failed.",
        )
        raise B2BSourceAcquisitionFailure(
            str(exc),
            validation_issues=[issue],
            metadata={**exc.metadata, "b2b_source_mode": "google_sheets"},
        ) from exc

    found_count = sum(1 for record in read_result.sheet_audit if record.sheet_found)
    missing_count = sum(1 for record in read_result.sheet_audit if not record.sheet_found)
    logger.info(
        "Google Sheet scanned: spreadsheet_id=%s target_sheets_found=%s target_sheets_missing=%s raw_rows=%s",
        config.google_spreadsheet_id,
        found_count,
        missing_count,
        len(read_result.rows),
    )
    for record in read_result.sheet_audit:
        if record.header_row_found:
            logger.info(
                "Header row detected: source=%s sheet=%s channel=%s row=%s",
                record.source_file,
                record.actual_sheet_name,
                record.source_channel,
                record.header_row_found,
            )

    return B2BDispatchSourceReadResult(
        source_mode="google_sheets",
        source_label=f"Google Sheets spreadsheet {config.google_spreadsheet_id}",
        source_count=1,
        source_identifiers=[f"google_sheets:{config.google_spreadsheet_id}"],
        rows=read_result.rows,
        sheet_audit=read_result.sheet_audit,
        metadata={
            **read_result.metadata,
            "source_traceability": "google_sheets_api",
            "source_identifiers": [f"google_sheets:{config.google_spreadsheet_id}"],
        },
    )


def _copy_inputs(files: list[Path], destination: Path) -> dict[Path, Path]:
    copied: dict[Path, Path] = {}
    destination.mkdir(parents=True, exist_ok=True)
    for path in files:
        target = destination / path.name
        shutil.copy2(path, target)
        copied[path] = target
    return copied
