"""Google Sheets reading for the B2B Dispatch Tracker source tabs."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Protocol

from inventory_cover.b2b_dispatch_schemas import (
    B2BCellValue,
    B2BSheetAuditRecord,
    B2B_SOURCE_FIELDS,
    B2B_TARGET_SHEETS,
    B2BTargetSheetSpec,
    RawB2BDispatchRow,
)
from inventory_cover.config import B2BDispatchPipelineConfig
from inventory_cover.exceptions import CatastrophicPipelineError
from inventory_cover.io.b2b_dispatch_excel_io import (
    B2BHeaderDetection,
    MIN_B2B_HEADER_SCORE,
    map_b2b_headers,
)
from inventory_cover.utils.text_cleaning import is_blank, normalize_header


class B2BGoogleSheetsError(CatastrophicPipelineError):
    """Actionable Google Sheets acquisition failure."""

    def __init__(
        self,
        message: str,
        *,
        issue_type: str,
        classification: str,
        metadata: dict[str, Any] | None = None,
    ):
        super().__init__(message)
        self.issue_type = issue_type
        self.classification = classification
        self.metadata = metadata or {}


class B2BGoogleSheetsClient(Protocol):
    """Small protocol used by production OAuth client and unit-test fakes."""

    def list_sheet_titles(self, spreadsheet_id: str) -> list[str]:
        """Return spreadsheet tab titles."""

    def get_values(self, spreadsheet_id: str, range_name: str) -> list[list[Any]]:
        """Return row-major cell values for one bounded range."""


@dataclass(frozen=True)
class B2BGoogleSheetsReadResult:
    rows: list[RawB2BDispatchRow]
    sheet_audit: list[B2BSheetAuditRecord]
    metadata: dict[str, Any]


def read_b2b_dispatch_google_sheet(
    config: B2BDispatchPipelineConfig,
    run_id: str,
    client: B2BGoogleSheetsClient | None = None,
) -> B2BGoogleSheetsReadResult:
    """Read the configured B2B dispatch tabs from Google Sheets."""

    if not config.google_spreadsheet_id.strip():
        raise B2BGoogleSheetsError(
            "B2B Google Sheets mode requires B2B_GOOGLE_SPREADSHEET_ID or --google-spreadsheet-id.",
            issue_type="GOOGLE_SPREADSHEET_ID_MISSING",
            classification="configuration",
            metadata={"google_api_failure_classification": "configuration"},
        )

    active_client = client or build_google_sheets_client(config)
    metadata = _base_metadata(config)
    rows: list[RawB2BDispatchRow] = []
    sheet_audit: list[B2BSheetAuditRecord] = []

    try:
        sheet_titles = active_client.list_sheet_titles(config.google_spreadsheet_id)
    except Exception as exc:
        raise _google_api_error(exc, "Could not list Google Sheet tabs.") from exc

    title_lookup = {normalize_header(title): title for title in sheet_titles}
    metadata["google_target_sheet_names_found"] = []

    for spec in B2B_TARGET_SHEETS:
        audit = B2BSheetAuditRecord(
            run_id=run_id,
            source_file=_google_source_file(config.google_spreadsheet_id),
            expected_sheet_name=spec.expected_sheet_name,
            source_channel=spec.source_channel,
        )
        actual_title = _find_google_sheet_title(title_lookup, spec)
        if actual_title is None:
            audit.status = "MISSING"
            audit.notes = "Expected Google Sheet tab was not found."
            sheet_audit.append(audit)
            continue

        audit.sheet_found = True
        audit.actual_sheet_name = actual_title
        metadata["google_target_sheet_names_found"].append(actual_title)

        range_name = _bounded_range(actual_title, config)
        try:
            values = active_client.get_values(config.google_spreadsheet_id, range_name)
        except Exception as exc:
            raise _google_api_error(exc, f"Could not fetch Google Sheet tab {actual_title!r}.") from exc

        metadata.setdefault("google_rows_fetched_by_sheet", {})[actual_title] = len(values)
        detection = detect_b2b_header_row_in_values(
            actual_title,
            values,
            configured_header_row=spec.google_header_row,
            max_scan_rows=config.header_scan_rows,
        )
        if detection is None:
            audit.status = "FAILED"
            audit.notes = f"Header row could not be detected; expected row {spec.google_header_row}."
            sheet_audit.append(audit)
            continue

        audit.header_row_found = detection.header_row
        if "Dispatch Date" not in detection.mapping:
            audit.status = "FAILED"
            audit.notes = f"Date column missing: expected {spec.google_dispatch_date_column}."
            sheet_audit.append(audit)
            continue
        if detection.missing_critical:
            audit.status = "FAILED"
            audit.notes = "Critical headers missing: " + ", ".join(detection.missing_critical)
            sheet_audit.append(audit)
            continue

        sheet_rows = _read_google_sheet_rows(config, values, actual_title, detection, spec)
        audit.rows_scanned = len(sheet_rows)
        audit.status = "READ"
        rows.extend(sheet_rows)
        sheet_audit.append(audit)

    metadata["google_header_rows_used"] = {
        record.actual_sheet_name or record.expected_sheet_name: record.header_row_found for record in sheet_audit
    }
    metadata["google_rows_scanned"] = sum(record.rows_scanned for record in sheet_audit)
    return B2BGoogleSheetsReadResult(rows=rows, sheet_audit=sheet_audit, metadata=metadata)


def detect_b2b_header_row_in_values(
    sheet_name: str,
    values: list[list[Any]],
    *,
    configured_header_row: int,
    max_scan_rows: int,
) -> B2BHeaderDetection | None:
    """Detect a B2B header row from Google Sheets values."""

    configured = _header_detection_for_row(sheet_name, values, configured_header_row)
    if configured is not None and configured.score >= MIN_B2B_HEADER_SCORE:
        return configured

    best: B2BHeaderDetection | None = None
    for row_number in range(1, min(max_scan_rows, len(values)) + 1):
        detection = _header_detection_for_row(sheet_name, values, row_number)
        if detection is None:
            continue
        if best is None or detection.score > best.score:
            best = detection
    if best is not None and best.score < MIN_B2B_HEADER_SCORE:
        return None
    return best


def build_google_sheets_client(config: B2BDispatchPipelineConfig) -> B2BGoogleSheetsClient:
    """Build an OAuth Desktop readonly Google Sheets client."""

    if not config.google_credentials_path.exists():
        raise B2BGoogleSheetsError(
            f"Google OAuth credentials file is missing: {_sanitize_path(config.google_credentials_path, config)}",
            issue_type="GOOGLE_CREDENTIALS_FILE_MISSING",
            classification="configuration",
            metadata={"google_api_failure_classification": "configuration"},
        )

    try:
        from google.auth.exceptions import RefreshError, TransportError
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from googleapiclient.discovery import build
    except ImportError as exc:
        raise B2BGoogleSheetsError(
            "Google Sheets dependencies are not installed. Run `python -m pip install -e .` first.",
            issue_type="GOOGLE_DEPENDENCY_MISSING",
            classification="dependency",
            metadata={"google_api_failure_classification": "dependency"},
        ) from exc

    scopes = list(config.google_readonly_scope)
    credentials = None
    try:
        if config.google_token_path.exists():
            credentials = Credentials.from_authorized_user_file(str(config.google_token_path), scopes)
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        if not credentials or not credentials.valid:
            flow = InstalledAppFlow.from_client_secrets_file(str(config.google_credentials_path), scopes)
            credentials = flow.run_local_server(port=0)
        config.google_token_path.parent.mkdir(parents=True, exist_ok=True)
        config.google_token_path.write_text(credentials.to_json(), encoding="utf-8")
    except RefreshError as exc:
        raise B2BGoogleSheetsError(
            "Google OAuth token refresh failed. Delete token.json and re-authorize if access changed.",
            issue_type="GOOGLE_OAUTH_TOKEN_ERROR",
            classification="oauth",
            metadata={"google_api_failure_classification": "oauth"},
        ) from exc
    except TransportError as exc:
        raise B2BGoogleSheetsError(
            f"Google OAuth transport failed: {exc}",
            issue_type="GOOGLE_TRANSPORT_FAILURE",
            classification="transport",
            metadata={"google_api_failure_classification": "transport"},
        ) from exc
    except Exception as exc:
        raise B2BGoogleSheetsError(
            f"Google OAuth consent/token flow failed: {exc}",
            issue_type="GOOGLE_OAUTH_CONSENT_ERROR",
            classification="oauth",
            metadata={"google_api_failure_classification": "oauth"},
        ) from exc

    try:
        service = build("sheets", "v4", credentials=credentials, cache_discovery=False)
    except Exception as exc:
        raise B2BGoogleSheetsError(
            f"Google Sheets API client initialization failed: {exc}",
            issue_type="GOOGLE_API_CLIENT_ERROR",
            classification="api",
            metadata={"google_api_failure_classification": "api"},
        ) from exc
    return _GoogleSheetsApiClient(service)


class _GoogleSheetsApiClient:
    def __init__(self, service: Any):
        self._service = service

    def list_sheet_titles(self, spreadsheet_id: str) -> list[str]:
        response = (
            self._service.spreadsheets()
            .get(spreadsheetId=spreadsheet_id, fields="sheets.properties.title")
            .execute()
        )
        return [
            str(sheet.get("properties", {}).get("title", ""))
            for sheet in response.get("sheets", [])
            if sheet.get("properties", {}).get("title")
        ]

    def get_values(self, spreadsheet_id: str, range_name: str) -> list[list[Any]]:
        response = (
            self._service.spreadsheets()
            .values()
            .get(
                spreadsheetId=spreadsheet_id,
                range=range_name,
                majorDimension="ROWS",
                valueRenderOption="UNFORMATTED_VALUE",
                dateTimeRenderOption="SERIAL_NUMBER",
            )
            .execute()
        )
        return [list(row) for row in response.get("values", [])]


def _header_detection_for_row(
    sheet_name: str,
    values: list[list[Any]],
    row_number: int,
) -> B2BHeaderDetection | None:
    if row_number < 1 or row_number > len(values):
        return None
    mapping = map_b2b_headers(list(values[row_number - 1]))
    if not mapping:
        return None
    from inventory_cover.b2b_dispatch_schemas import B2B_CRITICAL_FIELDS

    return B2BHeaderDetection(
        sheet_name=sheet_name,
        header_row=row_number,
        mapping=mapping,
        missing_critical=[field for field in B2B_CRITICAL_FIELDS if field not in mapping],
        score=len(mapping),
    )


def _read_google_sheet_rows(
    config: B2BDispatchPipelineConfig,
    values: list[list[Any]],
    actual_title: str,
    detection: B2BHeaderDetection,
    spec: B2BTargetSheetSpec,
) -> list[RawB2BDispatchRow]:
    rows: list[RawB2BDispatchRow] = []
    for source_row, row in enumerate(values[detection.header_row :], start=detection.header_row + 1):
        if _is_blank_values_row(row):
            continue
        cell_values: dict[str, B2BCellValue] = {}
        for field in B2B_SOURCE_FIELDS:
            idx = detection.mapping.get(field)
            value = row[idx] if idx is not None and idx < len(row) else None
            cell_values[field] = B2BCellValue(value=value)
        rows.append(
            RawB2BDispatchRow(
                source_file=_google_source_file(config.google_spreadsheet_id),
                source_path=Path("google_sheets") / config.google_spreadsheet_id,
                source_sheet=actual_title,
                source_row=source_row,
                source_channel=spec.source_channel,
                values=cell_values,
            )
        )
    return rows


def _find_google_sheet_title(
    title_lookup: dict[str, str],
    spec: B2BTargetSheetSpec,
) -> str | None:
    names = [spec.google_sheet_name or spec.expected_sheet_name, spec.expected_sheet_name, *spec.aliases]
    for name in names:
        match = title_lookup.get(normalize_header(name))
        if match:
            return match
    return None


def _bounded_range(sheet_title: str, config: B2BDispatchPipelineConfig) -> str:
    escaped_title = sheet_title.replace("'", "''")
    return f"'{escaped_title}'!A1:{config.google_values_max_column}{config.google_values_max_rows}"


def _base_metadata(config: B2BDispatchPipelineConfig) -> dict[str, Any]:
    return {
        "b2b_source_mode": "google_sheets",
        "google_spreadsheet_id": config.google_spreadsheet_id,
        "google_target_sheet_names_requested": [
            spec.google_sheet_name or spec.expected_sheet_name for spec in B2B_TARGET_SHEETS
        ],
        "google_dispatch_date_columns": {
            spec.google_sheet_name or spec.expected_sheet_name: spec.google_dispatch_date_column
            for spec in B2B_TARGET_SHEETS
        },
        "google_configured_header_rows": {
            spec.google_sheet_name or spec.expected_sheet_name: spec.google_header_row for spec in B2B_TARGET_SHEETS
        },
        "google_authentication_mode": "oauth_desktop_readonly",
        "google_readonly_scope": list(config.google_readonly_scope),
        "google_credentials_path": _sanitize_path(config.google_credentials_path, config),
        "google_token_path": _sanitize_path(config.google_token_path, config),
        "google_values_max_rows": config.google_values_max_rows,
        "google_values_max_column": config.google_values_max_column,
    }


def _google_api_error(exc: Exception, prefix: str) -> B2BGoogleSheetsError:
    status = getattr(getattr(exc, "resp", None), "status", None)
    classification = _classify_google_status(status)
    issue_type = {
        "oauth": "GOOGLE_OAUTH_TOKEN_ERROR",
        "access": "GOOGLE_SPREADSHEET_ACCESS_DENIED",
        "not_found": "GOOGLE_SPREADSHEET_NOT_FOUND",
        "quota": "GOOGLE_API_QUOTA_FAILURE",
        "transport": "GOOGLE_TRANSPORT_FAILURE",
    }.get(classification, "GOOGLE_API_FAILURE")
    detail = f"{prefix} {classification.upper()}: {exc}"
    return B2BGoogleSheetsError(
        detail,
        issue_type=issue_type,
        classification=classification,
        metadata={"google_api_failure_classification": classification},
    )


def _classify_google_status(status: Any) -> str:
    if status in {401}:
        return "oauth"
    if status in {403}:
        return "access"
    if status in {404}:
        return "not_found"
    if status in {429}:
        return "quota"
    if isinstance(status, int) and status >= 500:
        return "transport"
    return "api"


def _sanitize_path(path: Path, config: B2BDispatchPipelineConfig) -> str:
    resolved = Path(path).resolve()
    try:
        return str(resolved.relative_to(config.project_root.resolve()))
    except ValueError:
        return str(Path("<external>") / resolved.name)


def _google_source_file(spreadsheet_id: str) -> str:
    return f"google_sheets:{spreadsheet_id}"


def _is_blank_values_row(row: list[Any]) -> bool:
    return all(is_blank(value) for value in row)
