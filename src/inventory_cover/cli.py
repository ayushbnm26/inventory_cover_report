"""Command line interface for inventory_cover."""

from __future__ import annotations

import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from datetime import date
import os
from pathlib import Path
import shutil
import sys
from typing import Any, Callable

from inventory_cover.config import (
    B2BDispatchPipelineConfig,
    InventoryCoverPipelineConfig,
    PipelineConfig,
    SalesInventoryPipelineConfig,
)
from inventory_cover.exceptions import PipelineError
from inventory_cover.pipelines.b2b_dispatch_pipeline import B2BDispatchPipeline
from inventory_cover.pipelines.inventory_cover_pipeline import InventoryCoverPipeline
from inventory_cover.pipelines.po_items_pipeline import PoItemsPipeline
from inventory_cover.pipelines.sales_inventory_pipeline import SalesInventoryPipeline
from inventory_cover.notifications import (
    EmailConfigError,
    EmailDeliveryConfig,
    EmailDeliveryError,
    deliver_inventory_cover_report,
)
from inventory_cover.notifications.email_config import load_dotenv_values


_DOTENV_CACHE: dict[str, str] | None = None


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="inventory-cover")
    subparsers = parser.add_subparsers(dest="command")

    cleanup_parser = subparsers.add_parser(
        "stage-latest-inputs",
        help="Replace incoming folders with the latest source workbooks before running the pipelines.",
    )
    add_stage_latest_inputs_args(cleanup_parser)

    po_parser = subparsers.add_parser("run-po-items", help="Run the PO Items consolidation pipeline.")
    add_po_items_args(po_parser)

    b2b_parser = subparsers.add_parser("run-b2b-dispatch", help="Run the B2B Dispatch Tracker pipeline.")
    add_b2b_dispatch_args(b2b_parser)

    sales_inventory_parser = subparsers.add_parser(
        "run-sales-inventory",
        help="Run the Vendor Central Sales & Inventory backend pipeline.",
    )
    add_sales_inventory_args(sales_inventory_parser)

    source_parser = subparsers.add_parser(
        "run-source-pipelines",
        help="Run PO Items, B2B Dispatch, and Sales & Inventory from one command.",
    )
    add_source_pipelines_args(source_parser)

    inventory_cover_parser = subparsers.add_parser(
        "run-inventory-cover",
        help="Run the final Inventory Cover calculation engine on latest source outputs.",
    )
    add_inventory_cover_args(inventory_cover_parser)

    full_parser = subparsers.add_parser(
        "run-full-inventory-cover",
        help="Run all source pipelines and then the inventory cover engine end-to-end.",
    )
    add_full_inventory_cover_args(full_parser)

    subparsers.add_parser("list-pipelines", help="List available pipeline commands.")
    return parser


def add_po_items_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--input-dir", type=Path, default=PipelineConfig.input_dir)
    parser.add_argument("--run-root", type=Path, default=PipelineConfig.run_root)
    parser.add_argument("--processed-dir", type=Path, default=PipelineConfig.processed_dir)
    parser.add_argument("--min-files", type=int, default=2)
    parser.add_argument("--max-files", type=int, default=10)
    parser.add_argument("--allow-single-file", action="store_true")
    parser.add_argument("--allow-more-than-max-files", action="store_true")
    parser.add_argument("--dedupe-exact-rows", action="store_true")
    parser.add_argument("--log-level", default="INFO")


def add_b2b_dispatch_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--source", choices=("excel", "google-sheets"), default=None)
    parser.add_argument("--input-dir", type=Path, default=B2BDispatchPipelineConfig.input_dir)
    parser.add_argument("--run-root", type=Path, default=B2BDispatchPipelineConfig.run_root)
    parser.add_argument("--processed-dir", type=Path, default=B2BDispatchPipelineConfig.processed_dir)
    parser.add_argument("--as-of-date", "--b2b-as-of-date", dest="as_of_date", type=_parse_iso_date)
    parser.add_argument("--lookback-days", "--b2b-lookback-days", dest="lookback_days", type=int, default=2)
    parser.add_argument("--google-spreadsheet-id", default=None)
    parser.add_argument("--google-credentials-path", type=Path, default=None)
    parser.add_argument("--google-token-path", type=Path, default=None)
    parser.add_argument("--allow-multiple-files", action="store_true")
    parser.add_argument("--allow-missing-target-sheets", action="store_true")
    parser.add_argument("--dedupe-exact-rows", action="store_true")
    parser.add_argument("--log-level", default="INFO")


def add_sales_inventory_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--sales-input-dir", type=Path, default=SalesInventoryPipelineConfig.sales_input_dir)
    parser.add_argument("--inventory-input-dir", type=Path, default=SalesInventoryPipelineConfig.inventory_input_dir)
    parser.add_argument("--mapping-input-dir", type=Path, default=SalesInventoryPipelineConfig.mapping_input_dir)
    parser.add_argument("--run-root", type=Path, default=SalesInventoryPipelineConfig.run_root)
    parser.add_argument("--processed-dir", type=Path, default=SalesInventoryPipelineConfig.processed_dir)
    parser.add_argument("--require-sales", action="store_true")
    parser.add_argument("--require-inventory", action="store_true")
    parser.add_argument("--allow-multiple-sales-files", action="store_true")
    parser.add_argument("--allow-multiple-inventory-files", action="store_true")
    parser.add_argument("--dedupe-exact-rows", action="store_true")
    parser.add_argument("--log-level", default="INFO")


def add_source_pipelines_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--parallel", action="store_true", help="Run independent source pipelines concurrently.")
    parser.add_argument("--fail-fast", action="store_true", help="Stop after the first failed pipeline in sequential mode.")
    parser.add_argument("--min-free-gb", type=float, default=1.0, help="Minimum free disk space required before starting.")
    parser.add_argument("--run-root", type=Path, default=PipelineConfig.run_root)
    parser.add_argument("--log-level", default="INFO")
    parser.add_argument("--dedupe-exact-rows", action="store_true", help="Apply exact-row dedupe in every source pipeline.")

    parser.add_argument("--skip-po-items", action="store_true")
    parser.add_argument("--skip-b2b-dispatch", action="store_true")
    parser.add_argument("--skip-sales-inventory", action="store_true")

    parser.add_argument("--po-input-dir", type=Path, default=PipelineConfig.input_dir)
    parser.add_argument("--po-processed-dir", type=Path, default=PipelineConfig.processed_dir)
    parser.add_argument("--po-min-files", type=int, default=2)
    parser.add_argument("--po-max-files", type=int, default=10)
    parser.add_argument("--po-allow-single-file", action="store_true")
    parser.add_argument("--po-allow-more-than-max-files", action="store_true")

    parser.add_argument("--b2b-input-dir", type=Path, default=B2BDispatchPipelineConfig.input_dir)
    parser.add_argument("--b2b-processed-dir", type=Path, default=B2BDispatchPipelineConfig.processed_dir)
    parser.add_argument("--b2b-source", choices=("excel", "google-sheets"), default=None)
    parser.add_argument("--b2b-as-of-date", type=_parse_iso_date)
    parser.add_argument("--b2b-lookback-days", type=int, default=2)
    parser.add_argument("--b2b-google-spreadsheet-id", default=None)
    parser.add_argument("--b2b-google-credentials-path", type=Path, default=None)
    parser.add_argument("--b2b-google-token-path", type=Path, default=None)
    parser.add_argument("--b2b-allow-multiple-files", action="store_true")
    parser.add_argument("--b2b-allow-missing-target-sheets", action="store_true")

    parser.add_argument("--sales-input-dir", type=Path, default=SalesInventoryPipelineConfig.sales_input_dir)
    parser.add_argument("--inventory-input-dir", type=Path, default=SalesInventoryPipelineConfig.inventory_input_dir)
    parser.add_argument("--mapping-input-dir", type=Path, default=SalesInventoryPipelineConfig.mapping_input_dir)
    parser.add_argument("--sales-inventory-processed-dir", type=Path, default=SalesInventoryPipelineConfig.processed_dir)
    parser.add_argument("--require-sales", action="store_true")
    parser.add_argument("--require-inventory", action="store_true")
    parser.add_argument("--allow-multiple-sales-files", action="store_true")
    parser.add_argument("--allow-multiple-inventory-files", action="store_true")


def add_inventory_cover_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--po-backend-path", type=Path, default=InventoryCoverPipelineConfig.po_backend_path)
    parser.add_argument("--b2b-backend-path", type=Path, default=InventoryCoverPipelineConfig.b2b_backend_path)
    parser.add_argument("--sales-backend-path", type=Path, default=InventoryCoverPipelineConfig.sales_backend_path)
    parser.add_argument(
        "--inventory-backend-path", type=Path, default=InventoryCoverPipelineConfig.inventory_backend_path
    )
    parser.add_argument("--asin-master-path", type=Path, default=InventoryCoverPipelineConfig.asin_master_path)
    parser.add_argument("--output-dir", type=Path, default=None, help="Override processed/latest output directory.")
    parser.add_argument("--run-root", type=Path, default=InventoryCoverPipelineConfig.run_root)
    parser.add_argument("--processed-dir", type=Path, default=InventoryCoverPipelineConfig.processed_dir)
    parser.add_argument("--sales-window-days", type=int, default=InventoryCoverPipelineConfig.sales_window_days)
    parser.add_argument("--default-target-doh", type=float, default=InventoryCoverPipelineConfig.default_target_doh)
    parser.add_argument("--blank-numeric-policy", default=InventoryCoverPipelineConfig.blank_numeric_policy)
    parser.add_argument("--strict-freshness", action="store_true")
    parser.add_argument("--log-level", default="INFO")
    parser.add_argument("--send-email", action="store_true", help="Email the generated team workbook after success.")
    parser.add_argument("--email-dry-run", action="store_true", help="Build and audit the email without SMTP.")
    parser.add_argument("--email-env-file", type=Path, default=Path(".env"), help="Local .env file for email settings.")


def add_full_inventory_cover_args(parser: argparse.ArgumentParser) -> None:
    add_inventory_cover_args(parser)
    parser.add_argument("--parallel", action="store_true", help="Run independent source pipelines concurrently.")
    parser.add_argument("--fail-fast", action="store_true", help="Stop after the first failed pipeline in sequential mode.")
    parser.add_argument("--min-free-gb", type=float, default=1.0, help="Minimum free disk space required before starting.")
    parser.add_argument("--dedupe-exact-rows", action="store_true", help="Apply exact-row dedupe in every source pipeline.")

    parser.add_argument("--po-input-dir", type=Path, default=PipelineConfig.input_dir)
    parser.add_argument("--po-processed-dir", type=Path, default=PipelineConfig.processed_dir)
    parser.add_argument("--po-min-files", type=int, default=2)
    parser.add_argument("--po-max-files", type=int, default=10)
    parser.add_argument("--po-allow-single-file", action="store_true")
    parser.add_argument("--po-allow-more-than-max-files", action="store_true")

    parser.add_argument("--b2b-input-dir", type=Path, default=B2BDispatchPipelineConfig.input_dir)
    parser.add_argument("--b2b-processed-dir", type=Path, default=B2BDispatchPipelineConfig.processed_dir)
    parser.add_argument("--b2b-source", choices=("excel", "google-sheets"), default=None)
    parser.add_argument("--b2b-as-of-date", type=_parse_iso_date)
    parser.add_argument("--b2b-lookback-days", type=int, default=2)
    parser.add_argument("--b2b-google-spreadsheet-id", default=None)
    parser.add_argument("--b2b-google-credentials-path", type=Path, default=None)
    parser.add_argument("--b2b-google-token-path", type=Path, default=None)
    parser.add_argument("--b2b-allow-multiple-files", action="store_true")
    parser.add_argument("--b2b-allow-missing-target-sheets", action="store_true")

    parser.add_argument("--sales-input-dir", type=Path, default=SalesInventoryPipelineConfig.sales_input_dir)
    parser.add_argument("--inventory-input-dir", type=Path, default=SalesInventoryPipelineConfig.inventory_input_dir)
    parser.add_argument("--mapping-input-dir", type=Path, default=SalesInventoryPipelineConfig.mapping_input_dir)
    parser.add_argument("--sales-inventory-processed-dir", type=Path, default=SalesInventoryPipelineConfig.processed_dir)
    parser.add_argument("--require-sales", action="store_true")
    parser.add_argument("--require-inventory", action="store_true")
    parser.add_argument("--allow-multiple-sales-files", action="store_true")
    parser.add_argument("--allow-multiple-inventory-files", action="store_true")
    parser.add_argument("--skip-po-items", action="store_true")
    parser.add_argument("--skip-b2b-dispatch", action="store_true")
    parser.add_argument("--skip-sales-inventory", action="store_true")

    parser.add_argument("--skip-source-pipelines", action="store_true")
    parser.add_argument("--continue-on-source-warning", action="store_true")


def add_stage_latest_inputs_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--po-source", type=Path, nargs="+", required=True, help="Latest PO workbook(s) to stage.")
    parser.add_argument("--b2b-source", type=Path, required=True, help="Latest B2B dispatch workbook to stage.")
    parser.add_argument("--sales-source", type=Path, required=True, help="Latest sales workbook to stage.")
    parser.add_argument("--inventory-source", type=Path, required=True, help="Latest inventory workbook to stage.")
    parser.add_argument("--po-input-dir", type=Path, default=PipelineConfig.input_dir)
    parser.add_argument("--b2b-input-dir", type=Path, default=B2BDispatchPipelineConfig.input_dir)
    parser.add_argument("--sales-input-dir", type=Path, default=SalesInventoryPipelineConfig.sales_input_dir)
    parser.add_argument("--inventory-input-dir", type=Path, default=SalesInventoryPipelineConfig.inventory_input_dir)
    parser.add_argument("--clean-first", action="store_true", default=True, help="Remove existing .xlsx files first.")
    parser.add_argument("--keep-backups", action="store_true", help="Keep a copy of removed files in a backup folder.")
    parser.add_argument("--backup-root", type=Path, default=Path(".tmp") / "input_backups")


def inventory_cover_config_from_args(args: argparse.Namespace) -> InventoryCoverPipelineConfig:
    processed_dir = args.output_dir if getattr(args, "output_dir", None) else args.processed_dir
    return InventoryCoverPipelineConfig(
        po_backend_path=args.po_backend_path,
        b2b_backend_path=args.b2b_backend_path,
        sales_backend_path=args.sales_backend_path,
        inventory_backend_path=args.inventory_backend_path,
        asin_master_path=args.asin_master_path,
        run_root=args.run_root,
        processed_dir=processed_dir,
        sales_window_days=args.sales_window_days,
        default_target_doh=args.default_target_doh,
        blank_numeric_policy=args.blank_numeric_policy,
        strict_freshness=args.strict_freshness,
        log_level=args.log_level,
    )


def stage_latest_inputs_from_args(args: argparse.Namespace) -> int:
    staged = [
        _stage_input_folder(Path(args.po_input_dir), [Path(p) for p in args.po_source], args),
        _stage_input_folder(Path(args.b2b_input_dir), [Path(args.b2b_source)], args),
        _stage_input_folder(Path(args.sales_input_dir), [Path(args.sales_source)], args),
        _stage_input_folder(Path(args.inventory_input_dir), [Path(args.inventory_source)], args),
    ]
    for folder, copied in staged:
        print(f"Staged {len(copied)} file(s) into {folder}")
        for source, target in copied:
            print(f"  {source} -> {target}")
    return 0


def run_inventory_cover_from_args(args: argparse.Namespace) -> int:
    result = InventoryCoverPipeline(inventory_cover_config_from_args(args)).run()
    print(f"Run ID: {result.run_id}")
    print(f"Team workbook: {result.team_output_file}")
    print(f"Latest team workbook: {result.team_latest_file}")
    print(f"Backend audit workbook: {result.backend_output_file}")
    print(f"Latest backend audit workbook: {result.backend_latest_file}")
    print(f"Metadata: {result.metadata_file}")
    print(f"Validation issues file: {result.validation_file}")
    print(f"Log: {result.log_file}")
    print(f"Products: {result.product_count}")
    print(f"Validation issues: {result.validation_issue_count}")
    print(f"Warnings: {result.warning_count}")
    if getattr(args, "send_email", False):
        try:
            email_config = EmailDeliveryConfig.from_environment(env_file=args.email_env_file)
            delivery = deliver_inventory_cover_report(
                result,
                email_config,
                dry_run=getattr(args, "email_dry_run", False),
            )
            print(f"Email delivery status: {delivery.status}")
            print(f"Email audit: {delivery.audit_file}")
            print(f"Email log: {delivery.log_file}")
        except (EmailConfigError, EmailDeliveryError) as exc:
            print(f"ERROR: Email delivery failed: {exc}", file=sys.stderr)
            return 1
    return 0


def _stage_input_folder(folder: Path, sources: list[Path], args: argparse.Namespace) -> tuple[Path, list[tuple[Path, Path]]]:
    folder.mkdir(parents=True, exist_ok=True)
    if args.clean_first:
        existing = [path for path in folder.iterdir() if path.is_file() and path.suffix.lower() == ".xlsx"]
        backup_dir = None
        if args.keep_backups and existing:
            backup_dir = Path(args.backup_root) / folder.name
            backup_dir.mkdir(parents=True, exist_ok=True)
        for path in existing:
            if backup_dir is not None:
                shutil.copy2(path, backup_dir / path.name)
            path.unlink()

    copied: list[tuple[Path, Path]] = []
    for source in sources:
        if not source.exists():
            raise PipelineError(f"Source workbook does not exist: {source}")
        target = folder / source.name
        shutil.copy2(source, target)
        copied.append((source, target))
    return folder, copied


def run_full_inventory_cover_from_args(args: argparse.Namespace) -> int:
    source_ok = True
    if not args.skip_source_pipelines:
        print("Step 1/2: Running source pipelines...")
        source_outcomes = _run_full_source_pipelines(args)
        _print_source_pipeline_outcomes(source_outcomes)
        source_ok = all(outcome.status == "SUCCESS" for outcome in source_outcomes)
        if not source_ok and args.fail_fast:
            print("ERROR: A source pipeline failed and --fail-fast is set; stopping before the engine.")
            return 1
        if not source_ok and not args.continue_on_source_warning:
            print(
                "ERROR: A source pipeline failed. Re-run with --continue-on-source-warning to run the "
                "engine on whatever latest outputs are available, or --skip-source-pipelines."
            )
            return 1
    else:
        print("Step 1/2: Skipping source pipelines (--skip-source-pipelines).")

    print("Step 2/2: Running inventory cover calculation engine...")
    rc = run_inventory_cover_from_args(args)
    if rc == 0 and not source_ok:
        print("NOTE: Engine ran on existing latest outputs because some source pipelines did not succeed.")
    return rc


def _run_full_source_pipelines(args: argparse.Namespace) -> list["SourcePipelineOutcome"]:
    tasks = _build_source_pipeline_tasks(args)
    if args.parallel:
        return _run_tasks_parallel(tasks)
    return _run_tasks_sequential(tasks, fail_fast=args.fail_fast)


def config_from_args(args: argparse.Namespace) -> PipelineConfig:
    return PipelineConfig(
        input_dir=args.input_dir,
        run_root=args.run_root,
        processed_dir=args.processed_dir,
        min_files=args.min_files,
        max_files=args.max_files,
        allow_single_file=args.allow_single_file,
        allow_more_than_max_files=args.allow_more_than_max_files,
        dedupe_exact_rows=args.dedupe_exact_rows,
        log_level=args.log_level,
    )


def b2b_config_from_args(args: argparse.Namespace) -> B2BDispatchPipelineConfig:
    return B2BDispatchPipelineConfig(
        input_dir=args.input_dir,
        run_root=args.run_root,
        processed_dir=args.processed_dir,
        as_of_date=args.as_of_date,
        lookback_days=args.lookback_days,
        source_mode=_b2b_source_mode_from_args(getattr(args, "source", None)),
        allow_multiple_files=args.allow_multiple_files,
        allow_missing_target_sheets=args.allow_missing_target_sheets,
        dedupe_exact_rows=args.dedupe_exact_rows,
        log_level=args.log_level,
        google_spreadsheet_id=_arg_or_env(
            getattr(args, "google_spreadsheet_id", None),
            "B2B_GOOGLE_SPREADSHEET_ID",
            B2BDispatchPipelineConfig.google_spreadsheet_id,
        ),
        google_credentials_path=_path_arg_or_env(
            getattr(args, "google_credentials_path", None),
            "B2B_GOOGLE_CREDENTIALS_PATH",
            B2BDispatchPipelineConfig.google_credentials_path,
        ),
        google_token_path=_path_arg_or_env(
            getattr(args, "google_token_path", None),
            "B2B_GOOGLE_TOKEN_PATH",
            B2BDispatchPipelineConfig.google_token_path,
        ),
    )


def sales_inventory_config_from_args(args: argparse.Namespace) -> SalesInventoryPipelineConfig:
    return SalesInventoryPipelineConfig(
        sales_input_dir=args.sales_input_dir,
        inventory_input_dir=args.inventory_input_dir,
        mapping_input_dir=args.mapping_input_dir,
        run_root=args.run_root,
        processed_dir=args.processed_dir,
        require_sales=args.require_sales,
        require_inventory=args.require_inventory,
        allow_multiple_sales_files=args.allow_multiple_sales_files,
        allow_multiple_inventory_files=args.allow_multiple_inventory_files,
        dedupe_exact_rows=args.dedupe_exact_rows,
        log_level=args.log_level,
    )


def run_po_items_from_args(args: argparse.Namespace) -> int:
    result = PoItemsPipeline(config_from_args(args)).run()
    print(f"Run ID: {result.run_id}")
    print(f"Team workbook: {result.output_file}")
    print(f"Latest team workbook: {result.latest_file}")
    print(f"Backend audit workbook: {result.backend_output_file}")
    print(f"Latest backend audit workbook: {result.backend_latest_file}")
    print(f"Metadata: {result.metadata_file}")
    print(f"Log: {result.log_file}")
    return 0


def run_b2b_dispatch_from_args(args: argparse.Namespace) -> int:
    result = B2BDispatchPipeline(b2b_config_from_args(args)).run()
    print(f"Run ID: {result.run_id}")
    print(f"Backend audit workbook: {result.backend_output_file}")
    print(f"Latest backend audit workbook: {result.backend_latest_file}")
    print(f"Metadata: {result.metadata_file}")
    print(f"Log: {result.log_file}")
    print(f"Rows written: {result.rows_written}")
    print(f"Validation issues: {result.validation_issue_count}")
    print(f"Duplicate count: {result.duplicate_count}")
    return 0


def run_sales_inventory_from_args(args: argparse.Namespace) -> int:
    result = SalesInventoryPipeline(sales_inventory_config_from_args(args)).run()
    print(f"Run ID: {result.run_id}")
    print(f"Sales backend workbook: {result.sales_output_file or 'NOT GENERATED'}")
    print(f"Inventory backend workbook: {result.inventory_output_file or 'NOT GENERATED'}")
    print(f"Run summary workbook: {result.summary_output_file}")
    print(f"Latest sales backend workbook: {result.latest_sales_backend_file or 'NOT GENERATED'}")
    print(f"Latest inventory backend workbook: {result.latest_inventory_backend_file or 'NOT GENERATED'}")
    print(f"Latest run summary workbook: {result.latest_run_summary_file}")
    print(f"Metadata: {result.metadata_file}")
    print(f"Log: {result.log_file}")
    print(f"Sales rows written: {result.sales_rows_written}")
    print(f"Inventory rows written: {result.inventory_rows_written}")
    print(f"Validation issues: {result.validation_issue_count}")
    print(f"Duplicate count: {result.duplicate_count}")
    return 0


@dataclass(frozen=True)
class SourcePipelineTask:
    name: str
    command: str
    runner: Callable[[], Any]


@dataclass
class SourcePipelineOutcome:
    name: str
    command: str
    status: str
    run_id: str = ""
    rows: str = ""
    validation_issues: int | None = None
    duplicate_count: int | None = None
    output_paths: list[tuple[str, Path]] = field(default_factory=list)
    error: str = ""


def run_source_pipelines_from_args(args: argparse.Namespace) -> int:
    free_gb = _ensure_free_space(args.run_root, args.min_free_gb)
    tasks = _build_source_pipeline_tasks(args)
    if not tasks:
        print("No source pipelines selected.")
        return 2

    mode = "parallel" if args.parallel else "sequential"
    print(f"Central source pipeline run mode: {mode}")
    print(f"Free disk space at run root: {free_gb:.2f} GB")
    if args.parallel:
        print("Parallel run roots are isolated under: " + str(Path(args.run_root) / "source_pipelines"))

    outcomes = _run_tasks_parallel(tasks) if args.parallel else _run_tasks_sequential(tasks, fail_fast=args.fail_fast)
    _print_source_pipeline_outcomes(outcomes)
    return 0 if all(outcome.status == "SUCCESS" for outcome in outcomes) else 1


def list_pipelines_from_args(args: argparse.Namespace) -> int:
    print("Available pipelines:")
    print("1. stage-latest-inputs")
    print("2. run-po-items")
    print("3. run-b2b-dispatch")
    print("4. run-sales-inventory")
    print("5. run-source-pipelines")
    print("6. run-inventory-cover")
    print("7. run-full-inventory-cover")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    if args.command is None:
        parser.print_help()
        return 2
    try:
        if args.command == "run-po-items":
            return run_po_items_from_args(args)
        if args.command == "stage-latest-inputs":
            return stage_latest_inputs_from_args(args)
        if args.command == "run-b2b-dispatch":
            return run_b2b_dispatch_from_args(args)
        if args.command == "run-sales-inventory":
            return run_sales_inventory_from_args(args)
        if args.command == "run-source-pipelines":
            return run_source_pipelines_from_args(args)
        if args.command == "run-inventory-cover":
            return run_inventory_cover_from_args(args)
        if args.command == "run-full-inventory-cover":
            return run_full_inventory_cover_from_args(args)
        if args.command == "list-pipelines":
            return list_pipelines_from_args(args)
    except PipelineError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    parser.error(f"Unknown command: {args.command}")
    return 2


def _parse_iso_date(value: str) -> date:
    try:
        return date.fromisoformat(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("Expected date in YYYY-MM-DD format.") from exc


def _b2b_source_mode_from_args(explicit_source: str | None) -> str:
    if explicit_source:
        return explicit_source.strip().lower().replace("-", "_")
    if _env_flag_enabled("B2B_GOOGLE_SHEETS_ENABLED"):
        return "google_sheets"
    return "excel"


def _arg_or_env(value: str | None, env_key: str, default: str) -> str:
    if value is not None and str(value).strip():
        return str(value).strip()
    env_value = _env_config_value(env_key)
    if env_value is not None and env_value.strip():
        return env_value.strip()
    return default


def _path_arg_or_env(value: Path | None, env_key: str, default: Path) -> Path:
    if value is not None:
        return value
    env_value = _env_config_value(env_key)
    if env_value is not None and env_value.strip():
        return Path(env_value.strip())
    return default


def _env_flag_enabled(env_key: str) -> bool:
    value = _env_config_value(env_key) or ""
    return value.strip().lower() in {"1", "true", "yes", "y", "on"}


def _env_config_value(env_key: str) -> str | None:
    if env_key in os.environ:
        return os.environ.get(env_key)
    return _project_dotenv_values().get(env_key)


def _project_dotenv_values() -> dict[str, str]:
    global _DOTENV_CACHE
    if _DOTENV_CACHE is None:
        env_file = Path(".env")
        _DOTENV_CACHE = load_dotenv_values(env_file) if env_file.exists() else {}
    return _DOTENV_CACHE


def _build_source_pipeline_tasks(args: argparse.Namespace) -> list[SourcePipelineTask]:
    tasks: list[SourcePipelineTask] = []
    if not args.skip_po_items:
        po_config = PipelineConfig(
            input_dir=args.po_input_dir,
            run_root=_run_root_for(args, "po_items"),
            processed_dir=args.po_processed_dir,
            min_files=args.po_min_files,
            max_files=args.po_max_files,
            allow_single_file=args.po_allow_single_file,
            allow_more_than_max_files=args.po_allow_more_than_max_files,
            dedupe_exact_rows=args.dedupe_exact_rows,
            log_level=args.log_level,
        )
        tasks.append(
            SourcePipelineTask(
                name="Pipeline 1: PO Items",
                command="run-po-items",
                runner=lambda config=po_config: PoItemsPipeline(config).run(),
            )
        )
    if not args.skip_b2b_dispatch:
        b2b_config = B2BDispatchPipelineConfig(
            input_dir=args.b2b_input_dir,
            run_root=_run_root_for(args, "b2b_dispatch"),
            processed_dir=args.b2b_processed_dir,
            as_of_date=args.b2b_as_of_date,
            lookback_days=args.b2b_lookback_days,
            source_mode=_b2b_source_mode_from_args(getattr(args, "b2b_source", None)),
            allow_multiple_files=args.b2b_allow_multiple_files,
            allow_missing_target_sheets=args.b2b_allow_missing_target_sheets,
            dedupe_exact_rows=args.dedupe_exact_rows,
            log_level=args.log_level,
            google_spreadsheet_id=_arg_or_env(
                getattr(args, "b2b_google_spreadsheet_id", None),
                "B2B_GOOGLE_SPREADSHEET_ID",
                B2BDispatchPipelineConfig.google_spreadsheet_id,
            ),
            google_credentials_path=_path_arg_or_env(
                getattr(args, "b2b_google_credentials_path", None),
                "B2B_GOOGLE_CREDENTIALS_PATH",
                B2BDispatchPipelineConfig.google_credentials_path,
            ),
            google_token_path=_path_arg_or_env(
                getattr(args, "b2b_google_token_path", None),
                "B2B_GOOGLE_TOKEN_PATH",
                B2BDispatchPipelineConfig.google_token_path,
            ),
        )
        tasks.append(
            SourcePipelineTask(
                name="Pipeline 2: B2B Dispatch",
                command="run-b2b-dispatch",
                runner=lambda config=b2b_config: B2BDispatchPipeline(config).run(),
            )
        )
    if not args.skip_sales_inventory:
        sales_inventory_config = SalesInventoryPipelineConfig(
            sales_input_dir=args.sales_input_dir,
            inventory_input_dir=args.inventory_input_dir,
            mapping_input_dir=args.mapping_input_dir,
            run_root=_run_root_for(args, "sales_inventory"),
            processed_dir=args.sales_inventory_processed_dir,
            require_sales=args.require_sales,
            require_inventory=args.require_inventory,
            allow_multiple_sales_files=args.allow_multiple_sales_files,
            allow_multiple_inventory_files=args.allow_multiple_inventory_files,
            dedupe_exact_rows=args.dedupe_exact_rows,
            log_level=args.log_level,
        )
        tasks.append(
            SourcePipelineTask(
                name="Pipeline 3: Sales & Inventory",
                command="run-sales-inventory",
                runner=lambda config=sales_inventory_config: SalesInventoryPipeline(config).run(),
            )
        )
    return tasks


def _run_root_for(args: argparse.Namespace, pipeline_name: str) -> Path:
    run_root = Path(args.run_root)
    if args.parallel:
        return run_root / "source_pipelines" / pipeline_name
    return run_root


def _run_tasks_sequential(tasks: list[SourcePipelineTask], fail_fast: bool) -> list[SourcePipelineOutcome]:
    outcomes: list[SourcePipelineOutcome] = []
    for task in tasks:
        outcome = _run_one_task(task)
        outcomes.append(outcome)
        if fail_fast and outcome.status != "SUCCESS":
            break
    return outcomes


def _run_tasks_parallel(tasks: list[SourcePipelineTask]) -> list[SourcePipelineOutcome]:
    outcomes: list[SourcePipelineOutcome] = []
    with ThreadPoolExecutor(max_workers=len(tasks)) as executor:
        future_to_task = {executor.submit(task.runner): task for task in tasks}
        for future in as_completed(future_to_task):
            task = future_to_task[future]
            try:
                result = future.result()
                outcomes.append(_outcome_from_result(task, result))
            except Exception as exc:
                outcomes.append(
                    SourcePipelineOutcome(
                        name=task.name,
                        command=task.command,
                        status="FAILED",
                        error=str(exc),
                    )
                )
    order = {task.command: index for index, task in enumerate(tasks)}
    return sorted(outcomes, key=lambda outcome: order.get(outcome.command, 999))


def _run_one_task(task: SourcePipelineTask) -> SourcePipelineOutcome:
    try:
        return _outcome_from_result(task, task.runner())
    except Exception as exc:
        return SourcePipelineOutcome(
            name=task.name,
            command=task.command,
            status="FAILED",
            error=str(exc),
        )


def _outcome_from_result(task: SourcePipelineTask, result: Any) -> SourcePipelineOutcome:
    output_paths: list[tuple[str, Path]] = []
    for label, attr in (
        ("Team workbook", "output_file"),
        ("Backend audit workbook", "backend_output_file"),
        ("Sales backend workbook", "sales_output_file"),
        ("Inventory backend workbook", "inventory_output_file"),
        ("Run summary workbook", "summary_output_file"),
        ("Metadata", "metadata_file"),
        ("Log", "log_file"),
    ):
        value = getattr(result, attr, None)
        if value:
            output_paths.append((label, Path(value)))

    rows = ""
    if hasattr(result, "sales_rows_written") or hasattr(result, "inventory_rows_written"):
        rows = (
            f"sales={getattr(result, 'sales_rows_written', 0)}, "
            f"inventory={getattr(result, 'inventory_rows_written', 0)}"
        )
    elif hasattr(result, "rows_written"):
        rows = str(getattr(result, "rows_written"))

    return SourcePipelineOutcome(
        name=task.name,
        command=task.command,
        status="SUCCESS",
        run_id=str(getattr(result, "run_id", "")),
        rows=rows,
        validation_issues=getattr(result, "validation_issue_count", None),
        duplicate_count=getattr(result, "duplicate_count", None),
        output_paths=output_paths,
    )


def _print_source_pipeline_outcomes(outcomes: list[SourcePipelineOutcome]) -> None:
    print("Source pipeline results:")
    for outcome in outcomes:
        print(f"- {outcome.name} ({outcome.command}): {outcome.status}")
        if outcome.run_id:
            print(f"  Run ID: {outcome.run_id}")
        if outcome.rows:
            print(f"  Rows written: {outcome.rows}")
        if outcome.validation_issues is not None:
            print(f"  Validation issues: {outcome.validation_issues}")
        if outcome.duplicate_count is not None:
            print(f"  Duplicate count: {outcome.duplicate_count}")
        for label, path in outcome.output_paths:
            print(f"  {label}: {path}")
        if outcome.error:
            print(f"  ERROR: {outcome.error}")


def _ensure_free_space(run_root: Path, min_free_gb: float) -> float:
    probe = Path(run_root).resolve()
    while not probe.exists() and probe.parent != probe:
        probe = probe.parent
    usage = shutil.disk_usage(probe)
    free_gb = usage.free / (1024**3)
    if min_free_gb > 0 and free_gb < min_free_gb:
        raise PipelineError(
            f"Only {free_gb:.2f} GB free at {probe}; need at least {min_free_gb:.2f} GB."
        )
    return free_gb


if __name__ == "__main__":
    raise SystemExit(main())
