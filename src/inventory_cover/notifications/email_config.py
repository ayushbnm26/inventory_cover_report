"""Email delivery configuration loaded from environment or a local .env file."""

from __future__ import annotations

from dataclasses import dataclass
import os
from pathlib import Path
from typing import Mapping


class EmailConfigError(ValueError):
    """Raised when email delivery configuration is missing or invalid."""


@dataclass(frozen=True)
class EmailDeliveryConfig:
    """Validated SMTP and report-recipient settings."""

    smtp_host: str
    smtp_port: int | None
    smtp_username: str
    smtp_password: str
    smtp_from: str
    smtp_use_tls: bool
    smtp_use_ssl: bool
    smtp_timeout_seconds: float
    to: tuple[str, ...]
    cc: tuple[str, ...]
    bcc: tuple[str, ...]
    reply_to: str
    subject_prefix: str
    validation_errors: tuple[str, ...] = ()

    @classmethod
    def from_environment(
        cls,
        *,
        env: Mapping[str, str] | None = None,
        env_file: Path | None = Path(".env"),
    ) -> "EmailDeliveryConfig":
        """Build config from .env values overlaid by process environment."""

        file_values = load_dotenv_values(env_file) if env_file is not None else {}
        env_values = dict(os.environ if env is None else env)
        values = {**file_values, **env_values}
        validation_errors: list[str] = []

        return cls(
            smtp_host=_clean(values.get("SMTP_HOST")),
            smtp_port=_parse_optional_port(values.get("SMTP_PORT"), errors=validation_errors),
            smtp_username=_clean(values.get("SMTP_USERNAME")),
            smtp_password=_clean(values.get("SMTP_PASSWORD")),
            smtp_from=_clean(values.get("SMTP_FROM")),
            smtp_use_tls=_parse_bool(
                values.get("SMTP_USE_TLS"), default=True, key="SMTP_USE_TLS", errors=validation_errors
            ),
            smtp_use_ssl=_parse_bool(
                values.get("SMTP_USE_SSL"), default=False, key="SMTP_USE_SSL", errors=validation_errors
            ),
            smtp_timeout_seconds=_parse_timeout(values.get("SMTP_TIMEOUT_SECONDS"), errors=validation_errors),
            to=_parse_recipients(values.get("REPORT_EMAIL_TO")),
            cc=_parse_recipients(values.get("REPORT_EMAIL_CC")),
            bcc=_parse_recipients(values.get("REPORT_EMAIL_BCC")),
            reply_to=_clean(values.get("REPORT_EMAIL_REPLY_TO")),
            subject_prefix=_clean(values.get("REPORT_EMAIL_SUBJECT_PREFIX")) or "Inventory Cover Report",
            validation_errors=tuple(validation_errors),
        )

    def validate(self, *, dry_run: bool = False) -> None:
        """Validate fields required for dry-run or real SMTP delivery."""

        errors: list[str] = list(self.validation_errors)
        if not self.smtp_host:
            errors.append("SMTP_HOST is required")
        if self.smtp_port is None and not _has_error_for(errors, "SMTP_PORT"):
            errors.append("SMTP_PORT is required")
        elif not 1 <= self.smtp_port <= 65535:
            errors.append("SMTP_PORT must be between 1 and 65535")
        if not self.smtp_username:
            errors.append("SMTP_USERNAME is required")
        if not dry_run and not self.smtp_password:
            errors.append("SMTP_PASSWORD is required for network email delivery")
        if not self.smtp_from:
            errors.append("SMTP_FROM is required")
        if not self.to:
            errors.append("REPORT_EMAIL_TO is required")
        if self.smtp_use_tls and self.smtp_use_ssl:
            errors.append("SMTP_USE_TLS and SMTP_USE_SSL cannot both be true")
        if self.smtp_timeout_seconds <= 0:
            errors.append("SMTP_TIMEOUT_SECONDS must be greater than zero")
        if errors:
            raise EmailConfigError("; ".join(errors))

    @property
    def recipients(self) -> tuple[str, ...]:
        """All SMTP envelope recipients."""

        return self.to + self.cc + self.bcc


def load_dotenv_values(path: Path) -> dict[str, str]:
    """Load simple KEY=VALUE entries from a .env file without shell parsing."""

    values: dict[str, str] = {}
    if not path.exists():
        return values

    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except OSError as exc:
        raise EmailConfigError(f"Could not read email env file {path}: {exc}") from exc

    for line_number, raw_line in enumerate(lines, start=1):
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if "=" not in line:
            raise EmailConfigError(f"Invalid .env line {line_number}: expected KEY=VALUE")
        key, value = line.split("=", 1)
        key = key.strip()
        if not key:
            raise EmailConfigError(f"Invalid .env line {line_number}: empty key")
        values[key] = _unquote(value.strip())
    return values


def _clean(value: str | None) -> str:
    return "" if value is None else str(value).strip()


def _unquote(value: str) -> str:
    if len(value) >= 2 and value[0] == value[-1] and value[0] in {"'", '"'}:
        inner = value[1:-1]
        if value[0] == '"':
            inner = inner.replace(r"\"", '"').replace(r"\\", "\\")
        return inner
    return value


def _parse_optional_port(value: str | None, *, errors: list[str]) -> int | None:
    value = _clean(value)
    if not value:
        return None
    try:
        return int(value)
    except ValueError:
        errors.append("SMTP_PORT must be an integer")
        return None


def _parse_bool(value: str | None, *, default: bool, key: str, errors: list[str]) -> bool:
    value = _clean(value)
    if not value:
        return default
    lowered = value.lower()
    if lowered in {"1", "true", "yes", "y", "on"}:
        return True
    if lowered in {"0", "false", "no", "n", "off"}:
        return False
    errors.append(f"{key} must be true or false")
    return default


def _parse_timeout(value: str | None, *, errors: list[str]) -> float:
    value = _clean(value)
    if not value:
        return 30.0
    try:
        return float(value)
    except ValueError:
        errors.append("SMTP_TIMEOUT_SECONDS must be numeric")
        return 30.0


def _parse_recipients(value: str | None) -> tuple[str, ...]:
    raw = _clean(value)
    if not raw:
        return ()
    normalized = raw.replace(";", ",")
    return tuple(part.strip() for part in normalized.split(",") if part.strip())


def _has_error_for(errors: list[str], key: str) -> bool:
    return any(error.startswith(key) for error in errors)
