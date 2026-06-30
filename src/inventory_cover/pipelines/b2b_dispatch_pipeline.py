"""Orchestration for the B2B Dispatch Tracker backend pipeline."""

from __future__ import annotations

from datetime import date, datetime, timedelta
from pathlib import Path
import shutil
import time
from typing import Any

from inventory_cover.b2b_dispatch_schemas import (
    B2BPipelineRunResult,
    B2BSheetAuditRecord,
    B2BValidationIssue,
    B2B_TARGET_SHEETS,
    B2BDuplicateRecord,
    NormalizedB2BDispatchRow,
    RawB2BDispatchRow,
)
from inventory_cover.config import B2BDispatchPipelineConfig
from inventory_cover.exceptions import CatastrophicPipelineError
from inventory_cover.io.b2b_dispatch_source_provider import (
    B2BSourceAcquisitionFailure,
    acquire_b2b_dispatch_source,
    normalize_b2b_source_mode,
)
from inventory_cover.logging_utils import setup_run_logger, write_json_file
from inventory_cover.normalization.b2b_dispatch_normalizer import normalize_b2b_dispatch_row
from inventory_cover.reports.b2b_dispatch_excel_writer import write_b2b_dispatch_report
from inventory_cover.validation.b2b_dispatch_validators import attach_b2b_duplicate_findings


class B2BDispatchPipeline:
    """Coordinates discovery, reading, normalization, validation, and reporting."""

    def __init__(self, config: B2BDispatchPipelineConfig):
        self.config = config.resolved()

    def run(self) -> B2BPipelineRunResult:
        if self.config.lookback_days < 1:
            raise CatastrophicPipelineError("lookback_days must be at least 1.")

        run_id = _create_run_id()
        run_dir = self.config.run_root / run_id
        paths = _create_run_dirs(run_dir)
        run_dir = paths["run"]
        log_file = paths["logs"] / "b2b_dispatch_pipeline.log"
        logger = setup_run_logger(
            log_file,
            self.config.log_level,
            logger_name="inventory_cover.b2b_dispatch",
        )
        start_time = datetime.now()
        as_of_date = self.config.as_of_date or date.today()
        lookback_start = as_of_date - timedelta(days=self.config.lookback_days)
        lookback_end = as_of_date
        source_mode = normalize_b2b_source_mode(self.config.source_mode)

        output_file = paths["outputs_b2b"] / f"B2B_Dispatch_Backend_Audit_{run_id}.xlsx"
        latest_file = self.config.processed_dir / "latest" / "B2B_Dispatch_Backend_Audit_latest.xlsx"
        metadata_file = paths["metadata"] / "run_metadata.json"
        validation_file = paths["validation"] / "b2b_dispatch_validation_issues.json"
        duplicates_file = paths["validation"] / "b2b_dispatch_duplicates.json"

        metadata: dict[str, Any] = {
            "run_id": run_id,
            "start_time": start_time.isoformat(),
            "input_directory": str(self.config.input_dir),
            "run_directory": str(run_dir),
            "b2b_source_mode": source_mode,
            "as_of_date": as_of_date.isoformat(),
            "lookback_days": self.config.lookback_days,
            "lookback_start_date": lookback_start.isoformat(),
            "lookback_end_date": lookback_end.isoformat(),
            "status": "STARTED",
        }

        discovered_files: list[Path] = []
        raw_rows: list[RawB2BDispatchRow] = []
        rows: list[NormalizedB2BDispatchRow] = []
        rows_to_write: list[NormalizedB2BDispatchRow] = []
        sheet_audit: list[B2BSheetAuditRecord] = []
        validation_issues: list[B2BValidationIssue] = []
        duplicate_records: list[B2BDuplicateRecord] = []
        failed_files: set[str] = set()
        source_count = 0
        source_label = str(self.config.input_dir)
        source_identifiers: list[str] = []

        logger.info("B2B Dispatch pipeline started. run_id=%s", run_id)
        logger.info(
            "Config: source_mode=%s input_dir=%s run_root=%s processed_dir=%s as_of_date=%s lookback_days=%s "
            "allow_multiple_files=%s allow_missing_target_sheets=%s dedupe_exact_rows=%s",
            source_mode,
            self.config.input_dir,
            self.config.run_root,
            self.config.processed_dir,
            as_of_date,
            self.config.lookback_days,
            self.config.allow_multiple_files,
            self.config.allow_missing_target_sheets,
            self.config.dedupe_exact_rows,
        )

        try:
            source_result = acquire_b2b_dispatch_source(self.config, run_id, paths["inputs_b2b"], logger)
            discovered_files = source_result.discovered_files
            raw_rows.extend(source_result.rows)
            sheet_audit.extend(source_result.sheet_audit)
            validation_issues.extend(source_result.validation_issues)
            failed_files = set(source_result.failed_files)
            source_count = source_result.source_count
            source_label = source_result.source_label
            source_identifiers = source_result.source_identifiers
            metadata.update(source_result.metadata)
            metadata.update(
                {
                    "b2b_source_mode": source_result.source_mode,
                    "source_label": source_label,
                    "source_count": source_count,
                    "source_identifiers": source_identifiers,
                    "date_window_used": {
                        "as_of_date": as_of_date.isoformat(),
                        "lookback_days": self.config.lookback_days,
                        "lookback_start": lookback_start.isoformat(),
                        "lookback_end": lookback_end.isoformat(),
                        "inclusive": True,
                    },
                }
            )

            _apply_sheet_validation_policy(
                run_id=run_id,
                sheet_audit=sheet_audit,
                validation_issues=validation_issues,
                allow_missing_target_sheets=self.config.allow_missing_target_sheets,
            )

            audit_lookup = {
                (record.source_file, record.actual_sheet_name, record.source_channel): record
                for record in sheet_audit
                if record.sheet_found
            }
            for raw_row in raw_rows:
                result = normalize_b2b_dispatch_row(
                    raw_row,
                    run_id=run_id,
                    lookback_start=lookback_start,
                    lookback_end=lookback_end,
                    value_difference_tolerance=self.config.value_difference_tolerance,
                )
                validation_issues.extend(result.issues)
                audit = audit_lookup.get((raw_row.source_file, raw_row.source_sheet, raw_row.source_channel))
                if audit is not None:
                    if result.has_valid_dispatch_date:
                        audit.rows_with_valid_dispatch_date += 1
                    if result.included_in_window:
                        audit.rows_included += 1
                    if result.excluded_outside_window:
                        audit.rows_excluded_outside_window += 1
                    if result.rejected:
                        audit.rows_rejected += 1
                if result.normalized_row is not None:
                    rows.append(result.normalized_row)

            _finalize_sheet_audit(sheet_audit, validation_issues)
            logger.info(
                "Rows scanned=%s valid_dispatch_dates=%s included=%s outside_window=%s rejected=%s",
                sum(record.rows_scanned for record in sheet_audit),
                sum(record.rows_with_valid_dispatch_date for record in sheet_audit),
                sum(record.rows_included for record in sheet_audit),
                sum(record.rows_excluded_outside_window for record in sheet_audit),
                sum(record.rows_rejected for record in sheet_audit),
            )

            if rows:
                rows_to_write, duplicate_issues, duplicate_records = attach_b2b_duplicate_findings(
                    rows,
                    dedupe_exact_rows=self.config.dedupe_exact_rows,
                )
                validation_issues.extend(duplicate_issues)
                _annotate_dedupe_drops(sheet_audit, rows, rows_to_write)
            else:
                rows_to_write = []
                validation_issues.append(
                    B2BValidationIssue(
                        run_id=run_id,
                        severity="WARNING",
                        issue_type="NO_ROWS_INCLUDED_IN_WINDOW",
                        field_name="Dispatch Date",
                        raw_value=f"{lookback_start.isoformat()} through {lookback_end.isoformat()}",
                        issue_detail="No valid dispatch rows were included for the configured inclusive lookback window.",
                        action_taken="Empty latest backend workbook written so downstream transit is zero.",
                    )
                )
                logger.warning(
                    "No dispatch rows were included for %s through %s; writing an empty latest backend "
                    "so downstream transit is zero for this run.",
                    lookback_start,
                    lookback_end,
                )

            summary = _build_run_summary(
                run_id=run_id,
                start_time=start_time,
                source_label=source_label,
                source_count=source_count,
                output_folder=paths["outputs_b2b"],
                as_of_date=as_of_date,
                lookback_days=self.config.lookback_days,
                lookback_start=lookback_start,
                lookback_end=lookback_end,
                discovered_files=discovered_files,
                failed_files=failed_files,
                sheet_audit=sheet_audit,
                rows_written=len(rows_to_write),
                validation_issues=validation_issues,
                duplicate_count=len(duplicate_records),
                output_file=output_file,
                latest_file=latest_file,
                log_file=log_file,
            )

            write_json_file(validation_file, {"issues": [issue.as_json() for issue in validation_issues]})
            write_json_file(duplicates_file, {"duplicates": [record.as_json() for record in duplicate_records]})
            logger.info("Validation JSON written: %s", validation_file)
            logger.info("Duplicate JSON written: %s", duplicates_file)

            write_b2b_dispatch_report(
                output_path=output_file,
                rows=rows_to_write,
                run_summary=summary,
                sheet_audit=sheet_audit,
                validation_issues=validation_issues,
                duplicates=duplicate_records,
            )
            logger.info("Backend audit workbook written: %s", output_file)

            _copy_latest_or_warn(
                source=output_file,
                target=latest_file,
                run_id=run_id,
                validation_issues=validation_issues,
                logger=logger,
            )

            end_time = datetime.now()
            metadata.update(
                {
                    "status": "SUCCESS",
                    "end_time": end_time.isoformat(),
                    "duration_seconds": round((end_time - start_time).total_seconds(), 3),
                    "files_discovered": [str(path) for path in discovered_files],
                    "source_identifiers": source_identifiers,
                    "source_count": source_count,
                    "files_processed_successfully": summary["Files processed successfully"],
                    "files_skipped_or_failed": summary["Files skipped/failed"],
                    "target_sheets_expected": summary["Target sheets expected"],
                    "target_sheets_found": summary["Target sheets found"],
                    "target_sheets_missing": summary["Target sheets missing"],
                    "total_source_rows_scanned": summary["Total source rows scanned"],
                    "rows_with_valid_dispatch_date": summary["Rows with valid dispatch date"],
                    "rows_included_in_lookback_window": summary["Rows included in lookback window"],
                    "rows_excluded_outside_date_window": summary["Rows excluded outside date window"],
                    "rows_rejected_due_to_invalid_critical_fields": summary[
                        "Rows rejected due to invalid critical fields"
                    ],
                    "rows_written": len(rows_to_write),
                    "row_counts": {
                        "fetched": _metadata_rows_fetched(metadata, sheet_audit),
                        "scanned": sum(record.rows_scanned for record in sheet_audit),
                        "included": summary["Rows included in lookback window"],
                        "excluded": summary["Rows excluded outside date window"],
                        "rejected": summary["Rows rejected due to invalid critical fields"],
                    },
                    "warning_count": _count_warnings(validation_issues),
                    "error_count": _count_errors(validation_issues),
                    "duplicate_count": len(duplicate_records),
                    "backend_audit_workbook": str(output_file),
                    "latest_backend_audit_workbook": str(latest_file),
                    "metadata_file": str(metadata_file),
                    "log_file": str(log_file),
                    "validation_issues_file": str(validation_file),
                    "duplicates_file": str(duplicates_file),
                    "sheet_audit": [record.as_json() for record in sheet_audit],
                    "duplicates": [record.as_json() for record in duplicate_records],
                }
            )
            write_json_file(validation_file, {"issues": [issue.as_json() for issue in validation_issues]})
            write_json_file(metadata_file, metadata)
            logger.info("Run metadata written: %s", metadata_file)
            logger.info(
                "B2B Dispatch pipeline completed successfully. rows_written=%s validation_issues=%s duplicates=%s",
                len(rows_to_write),
                len(validation_issues),
                len(duplicate_records),
            )

            return B2BPipelineRunResult(
                run_id=run_id,
                run_dir=run_dir,
                backend_output_file=output_file,
                backend_latest_file=latest_file,
                metadata_file=metadata_file,
                log_file=log_file,
                rows_written=len(rows_to_write),
                validation_issue_count=len(validation_issues),
                duplicate_count=len(duplicate_records),
            )
        except B2BSourceAcquisitionFailure as exc:
            validation_issues.extend(exc.validation_issues)
            metadata.update(exc.metadata)
            logger.error("B2B Dispatch source acquisition failed: %s", exc)
            _write_failure_artifacts(
                metadata=metadata,
                metadata_file=metadata_file,
                validation_file=validation_file,
                duplicates_file=duplicates_file,
                validation_issues=validation_issues,
                duplicate_records=duplicate_records,
                sheet_audit=sheet_audit,
                discovered_files=discovered_files,
                start_time=start_time,
                error=str(exc),
            )
            raise CatastrophicPipelineError(str(exc)) from exc
        except CatastrophicPipelineError as exc:
            logger.error("B2B Dispatch pipeline failed: %s", exc)
            _write_failure_artifacts(
                metadata=metadata,
                metadata_file=metadata_file,
                validation_file=validation_file,
                duplicates_file=duplicates_file,
                validation_issues=validation_issues,
                duplicate_records=duplicate_records,
                sheet_audit=sheet_audit,
                discovered_files=discovered_files,
                start_time=start_time,
                error=str(exc),
            )
            raise


def _apply_sheet_validation_policy(
    run_id: str,
    sheet_audit: list[B2BSheetAuditRecord],
    validation_issues: list[B2BValidationIssue],
    allow_missing_target_sheets: bool,
) -> None:
    if not any(record.sheet_found for record in sheet_audit):
        validation_issues.append(
            B2BValidationIssue(
                run_id=run_id,
                severity="ERROR",
                issue_type="NO_TARGET_SHEETS_FOUND",
                issue_detail="None of the RK, Clicktech, or Etrade target sheets were found.",
                action_taken="Run failed.",
            )
        )
        raise CatastrophicPipelineError("No target sheets found in discovered workbook(s).")

    missing_records = [record for record in sheet_audit if not record.sheet_found]
    failed_records = [record for record in sheet_audit if record.sheet_found and record.status == "FAILED"]

    for record in missing_records:
        severity = "WARNING" if allow_missing_target_sheets else "ERROR"
        record.status = "MISSING_ALLOWED" if allow_missing_target_sheets else "MISSING"
        record.notes = (
            "Missing target sheet allowed by configuration."
            if allow_missing_target_sheets
            else "Missing required target sheet."
        )
        validation_issues.append(
            B2BValidationIssue(
                run_id=run_id,
                severity=severity,
                issue_type="MISSING_TARGET_SHEET",
                source_file=record.source_file,
                source_sheet=record.expected_sheet_name,
                source_channel=record.source_channel,
                field_name="Worksheet",
                raw_value=record.expected_sheet_name,
                issue_detail=record.notes,
                action_taken="Run continued." if allow_missing_target_sheets else "Run failed.",
            )
        )

    for record in failed_records:
        if "Date column missing" in record.notes:
            issue_type = "DATE_COLUMN_MISSING"
        elif "Critical headers missing" in record.notes:
            issue_type = "CRITICAL_HEADERS_MISSING"
        else:
            issue_type = "HEADER_ROW_NOT_FOUND"
        validation_issues.append(
            B2BValidationIssue(
                run_id=run_id,
                severity="ERROR",
                issue_type=issue_type,
                source_file=record.source_file,
                source_sheet=record.actual_sheet_name,
                source_channel=record.source_channel,
                field_name="Header Row",
                raw_value=record.header_row_found,
                issue_detail=record.notes,
                action_taken="Run failed.",
            )
        )

    failures: list[str] = []
    if missing_records and not allow_missing_target_sheets:
        failures.extend(f"{record.source_file}:{record.expected_sheet_name}" for record in missing_records)
    if failed_records:
        failures.extend(f"{record.source_file}:{record.actual_sheet_name} ({record.notes})" for record in failed_records)
    if failures:
        raise CatastrophicPipelineError("Target sheet validation failed: " + "; ".join(failures))


def _finalize_sheet_audit(
    sheet_audit: list[B2BSheetAuditRecord],
    validation_issues: list[B2BValidationIssue],
) -> None:
    warning_keys = {
        (issue.source_file, issue.source_sheet, issue.source_channel)
        for issue in validation_issues
        if issue.severity.upper() == "WARNING"
    }
    for record in sheet_audit:
        if record.status in {"MISSING", "MISSING_ALLOWED", "FAILED"}:
            continue
        if record.rows_rejected:
            record.status = "SUCCESS_WITH_REJECTIONS"
            record.notes = f"{record.rows_rejected} row(s) rejected due to invalid critical fields."
        elif (record.source_file, record.actual_sheet_name, record.source_channel) in warning_keys:
            record.status = "SUCCESS_WITH_WARNINGS"
            record.notes = "One or more included rows has warnings."
        elif record.rows_included == 0:
            record.status = "NO_ROWS_IN_WINDOW"
            record.notes = "No rows were included for the configured lookback window."
        else:
            record.status = "SUCCESS"
            record.notes = ""


def _annotate_dedupe_drops(
    sheet_audit: list[B2BSheetAuditRecord],
    all_rows: list[NormalizedB2BDispatchRow],
    kept_rows: list[NormalizedB2BDispatchRow],
) -> None:
    kept_ids = {id(row) for row in kept_rows}
    dropped_counts: dict[tuple[str, str, str], int] = {}
    for row in all_rows:
        if id(row) in kept_ids:
            continue
        key = (
            str(row.data.get("Source File") or ""),
            str(row.data.get("Source Sheet") or ""),
            str(row.data.get("Source Channel") or ""),
        )
        dropped_counts[key] = dropped_counts.get(key, 0) + 1
    for record in sheet_audit:
        dropped = dropped_counts.get((record.source_file, record.actual_sheet_name, record.source_channel), 0)
        if not dropped:
            continue
        note = f"{dropped} duplicate row(s) dropped by dedupe."
        record.notes = f"{record.notes} {note}".strip()


def _build_run_summary(
    run_id: str,
    start_time: datetime,
    source_label: str,
    source_count: int,
    output_folder: Path,
    as_of_date: date,
    lookback_days: int,
    lookback_start: date,
    lookback_end: date,
    discovered_files: list[Path],
    failed_files: set[str],
    sheet_audit: list[B2BSheetAuditRecord],
    rows_written: int,
    validation_issues: list[B2BValidationIssue],
    duplicate_count: int,
    output_file: Path,
    latest_file: Path,
    log_file: Path,
) -> dict[str, Any]:
    active_source_count = source_count if source_count else len(discovered_files)
    return {
        "Run ID": run_id,
        "Run timestamp": start_time.isoformat(timespec="seconds"),
        "Input folder": source_label,
        "Output folder": str(output_folder),
        "As of date": as_of_date,
        "Lookback days": lookback_days,
        "Lookback start date": lookback_start,
        "Lookback end date": lookback_end,
        "Files discovered": active_source_count,
        "Files processed successfully": active_source_count - len(failed_files),
        "Files skipped/failed": len(failed_files),
        "Target sheets expected": len(B2B_TARGET_SHEETS) * active_source_count,
        "Target sheets found": sum(1 for record in sheet_audit if record.sheet_found),
        "Target sheets missing": sum(1 for record in sheet_audit if not record.sheet_found),
        "Total source rows scanned": sum(record.rows_scanned for record in sheet_audit),
        "Rows with valid dispatch date": sum(record.rows_with_valid_dispatch_date for record in sheet_audit),
        "Rows included in lookback window": sum(record.rows_included for record in sheet_audit),
        "Rows excluded outside date window": sum(record.rows_excluded_outside_window for record in sheet_audit),
        "Rows rejected due to invalid critical fields": sum(record.rows_rejected for record in sheet_audit),
        "Rows written": rows_written,
        "Warning count": _count_warnings(validation_issues),
        "Error count": _count_errors(validation_issues),
        "Duplicate count": duplicate_count,
        "Output file name": output_file.name,
        "Latest backend file path": str(latest_file),
        "Log file path": str(log_file),
    }


def _write_failure_artifacts(
    metadata: dict[str, Any],
    metadata_file: Path,
    validation_file: Path,
    duplicates_file: Path,
    validation_issues: list[B2BValidationIssue],
    duplicate_records: list[B2BDuplicateRecord],
    sheet_audit: list[B2BSheetAuditRecord],
    discovered_files: list[Path],
    start_time: datetime,
    error: str,
) -> None:
    end_time = datetime.now()
    metadata.update(
        {
            "status": "FAILED",
            "end_time": end_time.isoformat(),
            "duration_seconds": round((end_time - start_time).total_seconds(), 3),
            "error": error,
            "files_discovered": [str(path) for path in discovered_files],
            "warning_count": _count_warnings(validation_issues),
            "error_count": _count_errors(validation_issues),
            "duplicate_count": len(duplicate_records),
            "sheet_audit": [record.as_json() for record in sheet_audit],
        }
    )
    write_json_file(validation_file, {"issues": [issue.as_json() for issue in validation_issues]})
    write_json_file(duplicates_file, {"duplicates": [record.as_json() for record in duplicate_records]})
    write_json_file(metadata_file, metadata)


def _create_run_id() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def _create_run_dirs(run_dir: Path) -> dict[str, Path]:
    if run_dir.exists():
        time.sleep(1.05)
        run_dir = run_dir.parent / _create_run_id()
    paths = {
        "run": run_dir,
        "inputs_b2b": run_dir / "inputs" / "b2b_dispatch",
        "outputs_b2b": run_dir / "outputs" / "b2b_dispatch",
        "logs": run_dir / "logs",
        "validation": run_dir / "validation",
        "metadata": run_dir / "metadata",
    }
    for path in paths.values():
        path.mkdir(parents=True, exist_ok=True)
    return paths


def _copy_inputs(files: list[Path], destination: Path) -> dict[Path, Path]:
    copied: dict[Path, Path] = {}
    destination.mkdir(parents=True, exist_ok=True)
    for path in files:
        target = destination / path.name
        shutil.copy2(path, target)
        copied[path] = target
    return copied


def _copy_latest(source: Path, target: Path) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    temp_latest = target.with_suffix(target.suffix + ".tmp")
    shutil.copy2(source, temp_latest)
    temp_latest.replace(target)


def _copy_latest_or_warn(
    source: Path,
    target: Path,
    run_id: str,
    validation_issues: list[B2BValidationIssue],
    logger: Any,
) -> bool:
    try:
        _copy_latest(source, target)
        logger.info("Latest backend audit workbook updated: %s", target)
        return True
    except PermissionError as exc:
        temp_latest = target.with_suffix(target.suffix + ".tmp")
        if temp_latest.exists():
            temp_latest.unlink(missing_ok=True)
        detail = f"Latest backend audit workbook could not be replaced, likely because the file is open: {target}"
        validation_issues.append(
            B2BValidationIssue(
                run_id=run_id,
                severity="WARNING",
                issue_type="LATEST_COPY_FAILED",
                field_name="Latest backend audit workbook",
                raw_value=str(target),
                issue_detail=f"{detail}. {exc}",
                action_taken="Timestamped run output was kept; close the latest workbook and rerun to refresh latest copy.",
            )
        )
        logger.warning("%s", detail)
        return False


def _count_warnings(issues: list[B2BValidationIssue]) -> int:
    return sum(1 for issue in issues if issue.severity.upper() == "WARNING")


def _count_errors(issues: list[B2BValidationIssue]) -> int:
    return sum(1 for issue in issues if issue.severity.upper() == "ERROR")


def _metadata_rows_fetched(metadata: dict[str, Any], sheet_audit: list[B2BSheetAuditRecord]) -> int:
    fetched_by_sheet = metadata.get("google_rows_fetched_by_sheet")
    if isinstance(fetched_by_sheet, dict):
        return sum(int(value or 0) for value in fetched_by_sheet.values())
    return sum(record.rows_scanned for record in sheet_audit)
