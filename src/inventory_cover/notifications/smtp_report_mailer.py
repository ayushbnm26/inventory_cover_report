"""SMTP delivery for generated Inventory Cover team workbooks."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from email.message import EmailMessage
import logging
from pathlib import Path
import smtplib
import time
from typing import Any, Callable

from inventory_cover.exceptions import PipelineError
from inventory_cover.inventory_cover_schemas import InventoryCoverPipelineRunResult
from inventory_cover.logging_utils import setup_run_logger, write_json_file
from inventory_cover.notifications.email_config import EmailConfigError, EmailDeliveryConfig
from inventory_cover.notifications.email_context import (
    NOT_AVAILABLE,
    EmailReportContext,
    build_email_report_context,
)


SMTPFactory = Callable[..., Any]


class EmailDeliveryError(PipelineError):
    """Raised when explicitly requested email delivery cannot complete."""


@dataclass(frozen=True)
class EmailDeliveryResult:
    """Traceable outcome for a dry-run or SMTP send attempt."""

    run_id: str
    status: str
    dry_run: bool
    audit_file: Path
    log_file: Path
    subject: str
    attachment_path: Path
    mail_attempted_at: str
    mail_sent_at: str
    duration_seconds: float
    error_type: str = ""
    error_message_sanitized: str = ""


def deliver_inventory_cover_report(
    result: InventoryCoverPipelineRunResult,
    config: EmailDeliveryConfig,
    *,
    dry_run: bool = False,
    smtp_factory: SMTPFactory | None = None,
) -> EmailDeliveryResult:
    """Send or dry-run the final team workbook email after Pipeline 4 succeeds."""

    notifications_dir = result.run_dir / "notifications"
    audit_file = notifications_dir / "email_delivery.json"
    log_file = result.run_dir / "logs" / "email_delivery.log"
    logger = setup_run_logger(log_file, logger_name=f"inventory_cover.email_delivery.{result.run_id}")
    started = time.monotonic()
    attempted_at = _now_iso()
    subject = ""
    context = build_email_report_context(result, logger=logger)

    logger.info(
        "Email delivery requested. dry_run=%s smtp_host=%s smtp_port=%s use_tls=%s use_ssl=%s "
        "from=%s to=%s cc=%s bcc=%s attachment=%s",
        dry_run,
        config.smtp_host,
        config.smtp_port,
        config.smtp_use_tls,
        config.smtp_use_ssl,
        config.smtp_from,
        ", ".join(config.to),
        ", ".join(config.cc),
        ", ".join(config.bcc),
        result.team_output_file,
    )

    try:
        config.validate(dry_run=dry_run)
        _ensure_attachment_exists(context.team_workbook_path)
        mail_timestamp = "Dry run - not sent" if dry_run else attempted_at
        subject = render_subject(context, config)
        message = build_email_message(context, config, mail_timestamp=mail_timestamp)

        if dry_run:
            duration = round(time.monotonic() - started, 3)
            logger.info(
                "Dry run complete. Would send subject=%s recipients=%s attachment=%s",
                subject,
                ", ".join(config.recipients),
                context.team_workbook_path,
            )
            payload = _audit_payload(
                context=context,
                config=config,
                status="DRY_RUN",
                dry_run=True,
                subject=subject,
                attempted_at=attempted_at,
                sent_at="",
                duration_seconds=duration,
            )
            _write_audit_or_raise(audit_file, payload, logger=logger)
            return EmailDeliveryResult(
                run_id=result.run_id,
                status="DRY_RUN",
                dry_run=True,
                audit_file=audit_file,
                log_file=log_file,
                subject=subject,
                attachment_path=context.team_workbook_path,
                mail_attempted_at=attempted_at,
                mail_sent_at="",
                duration_seconds=duration,
            )

        _send_message(config, message, smtp_factory=smtp_factory)
        sent_at = _now_iso()
        duration = round(time.monotonic() - started, 3)
        logger.info("Email sent successfully. sent_at=%s subject=%s", sent_at, subject)
        payload = _audit_payload(
            context=context,
            config=config,
            status="SUCCESS",
            dry_run=False,
            subject=subject,
            attempted_at=attempted_at,
            sent_at=sent_at,
            duration_seconds=duration,
        )
        _write_audit_or_raise(audit_file, payload, logger=logger)
        return EmailDeliveryResult(
            run_id=result.run_id,
            status="SUCCESS",
            dry_run=False,
            audit_file=audit_file,
            log_file=log_file,
            subject=subject,
            attachment_path=context.team_workbook_path,
            mail_attempted_at=attempted_at,
            mail_sent_at=sent_at,
            duration_seconds=duration,
        )
    except Exception as exc:
        error_type, sanitized = _classify_error(exc, config)
        duration = round(time.monotonic() - started, 3)
        logger.error("Email delivery failed. error_type=%s error=%s", error_type, sanitized)
        payload = _audit_payload(
            context=context,
            config=config,
            status="FAILED",
            dry_run=dry_run,
            subject=subject,
            attempted_at=attempted_at,
            sent_at="",
            duration_seconds=duration,
            error_type=error_type,
            error_message_sanitized=sanitized,
        )
        _write_audit_or_raise(audit_file, payload, logger=logger)
        raise EmailDeliveryError(f"{error_type}: {sanitized}") from exc


def build_email_message(
    context: EmailReportContext,
    config: EmailDeliveryConfig,
    *,
    mail_timestamp: str,
) -> EmailMessage:
    """Construct the plain-text business email with the team workbook attached."""

    message = EmailMessage()
    message["Subject"] = render_subject(context, config)
    message["From"] = config.smtp_from
    message["To"] = ", ".join(config.to)
    if config.cc:
        message["Cc"] = ", ".join(config.cc)
    if config.reply_to:
        message["Reply-To"] = config.reply_to
    message.set_content(render_body(context, mail_timestamp=mail_timestamp))

    attachment = context.team_workbook_path
    message.add_attachment(
        attachment.read_bytes(),
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=attachment.name,
    )
    return message


def render_subject(context: EmailReportContext, config: EmailDeliveryConfig) -> str:
    sales_start = context.report_context["sales_period_start"]
    sales_end = context.report_context["sales_period_end"]
    inventory_date = (
        context.report_context["inventory_report_updated_date"]
        if context.report_context["inventory_report_updated_date"] != NOT_AVAILABLE
        else context.report_context["inventory_period_end"]
    )
    if sales_start == NOT_AVAILABLE or sales_end == NOT_AVAILABLE or inventory_date == NOT_AVAILABLE:
        return f"{config.subject_prefix} | Run {context.run_id} | Report Dates Not Fully Available"
    return (
        f"{config.subject_prefix} | Run {context.run_id} | Sales {sales_start} to {sales_end} "
        f"| Inventory {inventory_date}"
    )


def render_body(context: EmailReportContext, *, mail_timestamp: str) -> str:
    rc = context.report_context
    return "\n".join(
        [
            "Inventory Cover Report generated successfully.",
            "",
            "Run ID:",
            context.run_id,
            "Generated at:",
            context.generated_at,
            "Email sent at:",
            mail_timestamp,
            "Attached workbook:",
            context.team_workbook_name,
            "",
            "Report context:",
            f"- Sales period start: {rc['sales_period_start']}",
            f"- Sales period end: {rc['sales_period_end']}",
            f"- Sales report updated date: {rc['sales_report_updated_date']}",
            f"- Inventory period start: {rc['inventory_period_start']}",
            f"- Inventory period end: {rc['inventory_period_end']}",
            f"- Inventory report updated date: {rc['inventory_report_updated_date']}",
            f"- B2B dispatch as-of date: {rc['b2b_dispatch_as_of_date']}",
            f"- B2B dispatch lookback start: {rc['b2b_dispatch_lookback_start']}",
            f"- B2B dispatch lookback end: {rc['b2b_dispatch_lookback_end']}",
            "",
            "Run summary:",
            f"- Product count: {context.product_count}",
            f"- Validation issue count: {context.validation_issue_count}",
            f"- Warning count: {context.warning_count}",
            f"- Team workbook: {context.team_workbook_path}",
            f"- Backend audit workbook: {context.backend_workbook_path}",
            "",
            "Note:",
            "This email was generated automatically after the Inventory Cover engine completed successfully.",
            "",
        ]
    )


def _send_message(
    config: EmailDeliveryConfig,
    message: EmailMessage,
    *,
    smtp_factory: SMTPFactory | None,
) -> None:
    factory = smtp_factory or (smtplib.SMTP_SSL if config.smtp_use_ssl else smtplib.SMTP)
    assert config.smtp_port is not None
    with factory(config.smtp_host, config.smtp_port, timeout=config.smtp_timeout_seconds) as smtp:
        if config.smtp_use_tls:
            smtp.starttls()
        smtp.login(config.smtp_username, config.smtp_password)
        smtp.send_message(message, from_addr=config.smtp_from, to_addrs=list(config.recipients))


def _ensure_attachment_exists(path: Path) -> None:
    if not path.exists():
        raise EmailDeliveryError(f"Attachment does not exist: {path}")
    if not path.is_file():
        raise EmailDeliveryError(f"Attachment is not a file: {path}")


def _audit_payload(
    *,
    context: EmailReportContext,
    config: EmailDeliveryConfig,
    status: str,
    dry_run: bool,
    subject: str,
    attempted_at: str,
    sent_at: str,
    duration_seconds: float,
    error_type: str = "",
    error_message_sanitized: str = "",
) -> dict[str, Any]:
    return {
        "run_id": context.run_id,
        "status": status,
        "dry_run": dry_run,
        "smtp_host": config.smtp_host,
        "smtp_port": config.smtp_port,
        "use_tls": config.smtp_use_tls,
        "use_ssl": config.smtp_use_ssl,
        "from_email": config.smtp_from,
        "to": list(config.to),
        "cc": list(config.cc),
        "bcc": list(config.bcc),
        "subject": subject,
        "attachment_path": str(context.team_workbook_path),
        "attachment_exists": context.team_workbook_path.exists(),
        "generated_at": context.generated_at,
        "mail_attempted_at": attempted_at,
        "mail_sent_at": sent_at,
        "duration_seconds": duration_seconds,
        "report_context": {
            **context.report_context,
            "missing_fields": list(context.missing_fields),
        },
        "error_type": error_type,
        "error_message_sanitized": error_message_sanitized,
    }


def _write_audit_or_raise(path: Path, payload: dict[str, Any], *, logger: logging.Logger) -> None:
    try:
        write_json_file(path, payload)
        logger.info("Email delivery audit written: %s", path)
    except OSError as exc:
        logger.error("Could not write email delivery audit JSON %s: %s", path, exc)
        raise EmailDeliveryError(f"Could not write email delivery audit JSON {path}: {exc}") from exc


def _classify_error(exc: Exception, config: EmailDeliveryConfig) -> tuple[str, str]:
    if isinstance(exc, EmailConfigError):
        error_type = "EMAIL_CONFIG_INVALID"
    elif isinstance(exc, EmailDeliveryError):
        error_type = "EMAIL_DELIVERY_ERROR"
    elif isinstance(exc, smtplib.SMTPAuthenticationError):
        error_type = "SMTP_AUTHENTICATION_FAILED"
    elif isinstance(exc, smtplib.SMTPRecipientsRefused):
        error_type = "SMTP_RECIPIENT_REJECTED"
    elif isinstance(exc, TimeoutError):
        error_type = "SMTP_CONNECTION_TIMEOUT"
    elif isinstance(exc, smtplib.SMTPConnectError):
        error_type = "SMTP_CONNECTION_FAILED"
    elif isinstance(exc, smtplib.SMTPException):
        error_type = "SMTP_ERROR"
    elif isinstance(exc, OSError):
        error_type = "SMTP_CONNECTION_FAILED"
    else:
        error_type = exc.__class__.__name__
    return error_type, _sanitize_error_message(str(exc), config)


def _sanitize_error_message(message: str, config: EmailDeliveryConfig) -> str:
    sanitized = message
    for secret in (config.smtp_password, config.smtp_username):
        if secret:
            sanitized = sanitized.replace(secret, "[REDACTED]")
    return sanitized


def _now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")
