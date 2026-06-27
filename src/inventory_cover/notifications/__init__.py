"""Notification integrations for inventory-cover reports."""

from inventory_cover.notifications.email_config import EmailConfigError, EmailDeliveryConfig
from inventory_cover.notifications.email_context import EmailReportContext, build_email_report_context
from inventory_cover.notifications.smtp_report_mailer import (
    EmailDeliveryError,
    EmailDeliveryResult,
    build_email_message,
    deliver_inventory_cover_report,
)

__all__ = [
    "EmailConfigError",
    "EmailDeliveryConfig",
    "EmailDeliveryError",
    "EmailDeliveryResult",
    "EmailReportContext",
    "build_email_message",
    "build_email_report_context",
    "deliver_inventory_cover_report",
]
