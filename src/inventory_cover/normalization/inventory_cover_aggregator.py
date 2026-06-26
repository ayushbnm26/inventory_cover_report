"""Consolidate backend source rows into a product-level cover universe."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Any

from inventory_cover.io.inventory_cover_input_loader import LoadedSheet, first_present_column
from inventory_cover.inventory_cover_schemas import (
    SOURCE_ASIN_MASTER,
    SOURCE_B2B,
    SOURCE_INVENTORY,
    SOURCE_PO,
    SOURCE_SALES,
    InventoryCoverValidationIssue,
    ProductCoverRow,
    SourceTraceRecord,
    coerce_date,
)
from inventory_cover.utils.numbers import parse_number, numbers_differ
from inventory_cover.utils.text_cleaning import clean_text, normalize_header


# Logical field -> candidate source headers.
ASIN_CANDIDATES = ("ASIN",)
MODEL_CANDIDATES = ("Model Number", "Model Number / SKU", "SKU", "Model")
BRAND_CANDIDATES = ("Brand",)
BRAND_CODE_CANDIDATES = ("Brand Code",)
CATEGORY_CANDIDATES = ("Category", "Main Category", "M-Category")
SUBCATEGORY_CANDIDATES = ("Subcategory", "Sub Category", "S-Category")
VENDOR_CANDIDATES = ("Child Vendor Code", "Vendor Code", "Vendor")

SALES_UNITS_CANDIDATES = ("Shipped Units",)
SALES_START_CANDIDATES = ("Viewing Range Start",)
SALES_END_CANDIDATES = ("Viewing Range End",)
REPORT_UPDATED_CANDIDATES = ("Report Updated Date",)

ON_HAND_CANDIDATES = ("Sellable On Hand Units", "Sellable On-Hand Units")
AMAZON_TRANSIT_PRIMARY = ("Sellable In Transit Units",)
AMAZON_TRANSIT_FALLBACK = ("In Transit Quantity",)

B2B_QTY_CANDIDATES = ("Dispatch Qty", "Dispatch Quantity")
B2B_LOOKBACK_CANDIDATES = ("Included In Lookback Window",)

PO_FINAL_CHAIN = (
    "Open PO Qty - Final",
    "Quantity Outstanding",
    "Open PO Qty - Source",
    "Open PO Qty - Derived",
)

# ASIN Master flexible columns.
AM_SKU_CANDIDATES = ("SKU", "Model Number", "Model Number / SKU")
AM_BRAND_CANDIDATES = ("Brand",)
AM_BRAND_NAME_CANDIDATES = ("Brand Name",)
AM_VENDOR_CANDIDATES = ("Vendor",)
AM_MAIN_CAT_CANDIDATES = ("Main Category", "M-Category")
AM_SUB_CAT_CANDIDATES = ("Sub Category", "S-Category")
AM_TARGET_CANDIDATES = ("Aligned DOH Target",)


@dataclass
class SourceStats:
    rows_read: int = 0
    rows_used: int = 0
    period_start: date | None = None
    period_end: date | None = None
    report_updated_date: date | None = None
    warnings: list[str] = field(default_factory=list)


@dataclass
class AggregationResult:
    products: list[ProductCoverRow]
    traces: list[SourceTraceRecord]
    issues: list[InventoryCoverValidationIssue]
    stats: dict[str, SourceStats]


def _cell(row: dict[str, Any], column: str | None) -> Any:
    if column is None:
        return None
    return row.get(column)


def _text(row: dict[str, Any], column: str | None) -> str:
    return clean_text(_cell(row, column))


def _number(value: Any) -> tuple[float | None, bool]:
    result = parse_number(value)
    if result.was_blank:
        return None, True
    if not result.ok or result.value is None:
        return None, False
    return float(result.value), False


def build_asin_master_lookup(sheet: LoadedSheet) -> dict[str, dict[str, Any]]:
    if not sheet.rows:
        return {}
    headers = sheet.headers
    cols = {
        "asin": first_present_column(headers, ASIN_CANDIDATES),
        "sku": first_present_column(headers, AM_SKU_CANDIDATES),
        "brand": first_present_column(headers, AM_BRAND_CANDIDATES),
        "brand_name": first_present_column(headers, AM_BRAND_NAME_CANDIDATES),
        "vendor": first_present_column(headers, AM_VENDOR_CANDIDATES),
        "main_category": first_present_column(headers, AM_MAIN_CAT_CANDIDATES),
        "sub_category": first_present_column(headers, AM_SUB_CAT_CANDIDATES),
        "target": first_present_column(headers, AM_TARGET_CANDIDATES),
    }
    lookup: dict[str, dict[str, Any]] = {}
    for row in sheet.rows:
        asin = _text(row, cols["asin"])
        if not asin:
            continue
        target_value, _ = _number(_cell(row, cols["target"]))
        lookup[asin] = {
            "sku": _text(row, cols["sku"]),
            "brand": _text(row, cols["brand"]),
            "brand_name": _text(row, cols["brand_name"]),
            "vendor": _text(row, cols["vendor"]),
            "main_category": _text(row, cols["main_category"]),
            "sub_category": _text(row, cols["sub_category"]),
            "target": target_value,
        }
    return lookup


def aggregate_products(
    run_id: str,
    sales: LoadedSheet,
    inventory: LoadedSheet,
    b2b: LoadedSheet,
    po: LoadedSheet,
    asin_master: LoadedSheet,
    config: Any,
) -> AggregationResult:
    issues: list[InventoryCoverValidationIssue] = []
    traces: list[SourceTraceRecord] = []
    stats: dict[str, SourceStats] = {
        SOURCE_SALES: SourceStats(),
        SOURCE_INVENTORY: SourceStats(),
        SOURCE_B2B: SourceStats(),
        SOURCE_PO: SourceStats(),
        SOURCE_ASIN_MASTER: SourceStats(),
    }

    asin_lookup = build_asin_master_lookup(asin_master)
    stats[SOURCE_ASIN_MASTER].rows_read = asin_master.row_count
    stats[SOURCE_ASIN_MASTER].rows_used = len(asin_lookup)

    # Pass A: build a deterministic model -> ASIN map to reduce duplicate rows.
    model_to_asin = _build_model_to_asin(sales, inventory, b2b, po, asin_lookup)

    products: dict[str, ProductCoverRow] = {}

    def get_product(asin: str, model: str) -> ProductCoverRow:
        key, key_type, resolved_asin = _resolve_key(asin, model, model_to_asin)
        if key not in products:
            products[key] = ProductCoverRow(
                product_key=key,
                product_key_type=key_type,
                asin=resolved_asin,
                model_number=model,
            )
        product = products[key]
        if not product.asin and resolved_asin:
            product.asin = resolved_asin
        return product

    _aggregate_inventory(run_id, inventory, get_product, traces, issues, stats[SOURCE_INVENTORY])
    _aggregate_sales(run_id, sales, get_product, traces, issues, stats[SOURCE_SALES], config)
    _aggregate_b2b(run_id, b2b, get_product, traces, issues, stats[SOURCE_B2B])
    _aggregate_po(run_id, po, get_product, traces, issues, stats[SOURCE_PO])
    _apply_global_sales_window(products.values(), stats[SOURCE_SALES])

    # Enrich identity and descriptive fields, applying ASIN master priority.
    for product in products.values():
        _finalize_identity(run_id, product, asin_lookup, config, issues)

    ordered = sorted(
        products.values(),
        key=lambda p: (p.asin or "~", p.model_number or "~"),
    )
    return AggregationResult(products=ordered, traces=traces, issues=issues, stats=stats)


def _build_model_to_asin(
    sales: LoadedSheet,
    inventory: LoadedSheet,
    b2b: LoadedSheet,
    po: LoadedSheet,
    asin_lookup: dict[str, dict[str, Any]],
) -> dict[str, str]:
    mapping: dict[str, set[str]] = {}

    def collect(sheet: LoadedSheet) -> None:
        if not sheet.rows:
            return
        asin_col = first_present_column(sheet.headers, ASIN_CANDIDATES)
        model_col = first_present_column(sheet.headers, MODEL_CANDIDATES)
        for row in sheet.rows:
            asin = _text(row, asin_col)
            model = _text(row, model_col)
            if asin and model:
                mapping.setdefault(normalize_header(model), set()).add(asin)

    collect(inventory)
    collect(sales)
    collect(po)
    collect(b2b)
    for asin, attrs in asin_lookup.items():
        sku = attrs.get("sku") or ""
        if sku:
            mapping.setdefault(normalize_header(sku), set()).add(asin)

    return {model: next(iter(asins)) for model, asins in mapping.items() if len(asins) == 1}


def _resolve_key(asin: str, model: str, model_to_asin: dict[str, str]) -> tuple[str, str, str]:
    if asin:
        return f"ASIN::{asin}", "ASIN", asin
    if model:
        resolved = model_to_asin.get(normalize_header(model))
        if resolved:
            return f"ASIN::{resolved}", "ASIN", resolved
        return f"MODEL::{normalize_header(model)}", "Model Number / SKU", ""
    return "UNKNOWN::blank", "Unknown", ""


def _aggregate_inventory(
    run_id: str,
    sheet: LoadedSheet,
    get_product,
    traces: list[SourceTraceRecord],
    issues: list[InventoryCoverValidationIssue],
    stats: SourceStats,
) -> None:
    if not sheet.rows:
        return
    headers = sheet.headers
    asin_col = first_present_column(headers, ASIN_CANDIDATES)
    model_col = first_present_column(headers, MODEL_CANDIDATES)
    on_hand_col = first_present_column(headers, ON_HAND_CANDIDATES)
    amz_primary_col = first_present_column(headers, AMAZON_TRANSIT_PRIMARY)
    amz_fallback_col = first_present_column(headers, AMAZON_TRANSIT_FALLBACK)
    end_col = first_present_column(headers, SALES_END_CANDIDATES)
    updated_col = first_present_column(headers, REPORT_UPDATED_CANDIDATES)
    brand_col = first_present_column(headers, BRAND_CANDIDATES)
    category_col = first_present_column(headers, CATEGORY_CANDIDATES)
    subcategory_col = first_present_column(headers, SUBCATEGORY_CANDIDATES)
    vendor_col = first_present_column(headers, VENDOR_CANDIDATES)

    stats.rows_read = sheet.row_count
    for index, row in enumerate(sheet.rows, start=2):
        asin = _text(row, asin_col)
        model = _text(row, model_col)
        if not asin and not model:
            continue
        product = get_product(asin, model)
        product.source_presence.add(SOURCE_INVENTORY)
        product.inventory_source_rows += 1
        stats.rows_used += 1

        on_hand, _ = _number(_cell(row, on_hand_col))
        on_hand = on_hand or 0.0
        primary, _ = _number(_cell(row, amz_primary_col))
        fallback, _ = _number(_cell(row, amz_fallback_col))
        amazon_transit = primary
        if amazon_transit is None:
            amazon_transit = fallback
        elif fallback is not None and numbers_differ(primary, fallback, 0.5):
            product.calculation_warning_notes.append(
                "Amazon transit: 'Sellable In Transit Units' preferred over differing 'In Transit Quantity'."
            )
        amazon_transit = amazon_transit or 0.0

        if on_hand < 0:
            issues.append(_issue(run_id, "WARNING", "NEGATIVE_INVENTORY", product, SOURCE_INVENTORY,
                                 sheet, index, on_hand_col or "Sellable On Hand Units", on_hand,
                                 "Negative sellable on-hand units in source.", "Value kept as-is."))

        product.on_hand_units += on_hand
        product.amazon_transit_units += amazon_transit

        _set_if_blank(product, "brand", _text(row, brand_col))
        _set_if_blank(product, "main_category", _text(row, category_col))
        _set_if_blank(product, "sub_category", _text(row, subcategory_col))
        _set_if_blank(product, "vendor", _text(row, vendor_col))

        end = coerce_date(_cell(row, end_col))
        updated = coerce_date(_cell(row, updated_col))
        stats.period_end = _max_date(stats.period_end, end)
        stats.report_updated_date = _max_date(stats.report_updated_date, updated)

        traces.append(SourceTraceRecord(
            run_id=run_id, product_key=product.product_key, product_key_type=product.product_key_type,
            final_asin=product.asin, final_model=product.model_number, source_type=SOURCE_INVENTORY,
            source_file=sheet.path.name, source_sheet=sheet.sheet_name, source_row=index,
            source_business_key=asin or model, quantity_used=on_hand, value_used=amazon_transit,
            date_used=end, trace_notes="On-hand and Amazon transit units.",
        ))


def _aggregate_sales(
    run_id: str,
    sheet: LoadedSheet,
    get_product,
    traces: list[SourceTraceRecord],
    issues: list[InventoryCoverValidationIssue],
    stats: SourceStats,
    config: Any,
) -> None:
    if not sheet.rows:
        return
    headers = sheet.headers
    asin_col = first_present_column(headers, ASIN_CANDIDATES)
    model_col = first_present_column(headers, MODEL_CANDIDATES)
    units_col = first_present_column(headers, SALES_UNITS_CANDIDATES)
    start_col = first_present_column(headers, SALES_START_CANDIDATES)
    end_col = first_present_column(headers, SALES_END_CANDIDATES)
    brand_col = first_present_column(headers, BRAND_CANDIDATES)
    category_col = first_present_column(headers, CATEGORY_CANDIDATES)
    subcategory_col = first_present_column(headers, SUBCATEGORY_CANDIDATES)
    vendor_col = first_present_column(headers, VENDOR_CANDIDATES)

    stats.rows_read = sheet.row_count
    # Group sales rows per product and keep only the latest sales period window.
    per_product: dict[str, list[dict[str, Any]]] = {}
    product_handles: dict[str, ProductCoverRow] = {}
    for index, row in enumerate(sheet.rows, start=2):
        asin = _text(row, asin_col)
        model = _text(row, model_col)
        if not asin and not model:
            continue
        product = get_product(asin, model)
        product_handles[product.product_key] = product
        units, _ = _number(_cell(row, units_col))
        units = units or 0.0
        if units < 0:
            issues.append(_issue(run_id, "WARNING", "NEGATIVE_SALES", product, SOURCE_SALES,
                                 sheet, index, units_col or "Shipped Units", units,
                                 "Negative shipped units in source.", "Value kept as-is."))
        per_product.setdefault(product.product_key, []).append({
            "row": index,
            "asin": asin,
            "model": model,
            "units": units,
            "start": coerce_date(_cell(row, start_col)),
            "end": coerce_date(_cell(row, end_col)),
            "brand": _text(row, brand_col),
            "category": _text(row, category_col),
            "subcategory": _text(row, subcategory_col),
            "vendor": _text(row, vendor_col),
        })

    for key, entries in per_product.items():
        product = product_handles[key]
        product.source_presence.add(SOURCE_SALES)
        ends = [e["end"] for e in entries if e["end"] is not None]
        latest_end = max(ends) if ends else None
        if latest_end is not None:
            window_entries = [e for e in entries if e["end"] == latest_end]
        else:
            window_entries = entries
        total_units = sum(e["units"] for e in window_entries)
        starts = [e["start"] for e in window_entries if e["start"] is not None]
        period_start = min(starts) if starts else None

        product.sales_units = total_units
        product.sales_units_raw = total_units
        product.sales_period_start = period_start
        product.sales_period_end = latest_end
        product.sales_source_rows += len(window_entries)
        stats.rows_used += len(window_entries)
        stats.period_start = _min_date(stats.period_start, period_start)
        stats.period_end = _max_date(stats.period_end, latest_end)

        first = window_entries[0]
        _set_if_blank(product, "brand", first["brand"])
        _set_if_blank(product, "main_category", first["category"])
        _set_if_blank(product, "sub_category", first["subcategory"])
        _set_if_blank(product, "vendor", first["vendor"])

        if total_units > 0 and latest_end is None:
            product.calculation_warning_notes.append(
                "Sales period missing; defaulted Sales Days to window cap."
            )
            issues.append(_issue(run_id, "WARNING", "SALES_PERIOD_MISSING", product, SOURCE_SALES,
                                 sheet, first["row"], "Viewing Range End", None,
                                 "Sales units exist without a sales period.",
                                 f"Sales Days defaulted to {config.sales_window_days}."))

        for e in window_entries:
            traces.append(SourceTraceRecord(
                run_id=run_id, product_key=product.product_key, product_key_type=product.product_key_type,
                final_asin=product.asin, final_model=product.model_number, source_type=SOURCE_SALES,
                source_file=sheet.path.name, source_sheet=sheet.sheet_name, source_row=e["row"],
                source_business_key=e["asin"] or e["model"], quantity_used=e["units"], value_used=None,
                date_used=e["end"], trace_notes="Shipped units in latest sales window.",
            ))


def _aggregate_b2b(
    run_id: str,
    sheet: LoadedSheet,
    get_product,
    traces: list[SourceTraceRecord],
    issues: list[InventoryCoverValidationIssue],
    stats: SourceStats,
) -> None:
    if not sheet.rows:
        return
    headers = sheet.headers
    asin_col = first_present_column(headers, ASIN_CANDIDATES)
    model_col = first_present_column(headers, MODEL_CANDIDATES)
    qty_col = first_present_column(headers, B2B_QTY_CANDIDATES)
    lookback_col = first_present_column(headers, B2B_LOOKBACK_CANDIDATES)

    stats.rows_read = sheet.row_count
    if lookback_col is None:
        stats.warnings.append("No 'Included In Lookback Window' column; all dispatch rows used.")

    for index, row in enumerate(sheet.rows, start=2):
        asin = _text(row, asin_col)
        model = _text(row, model_col)
        if not asin and not model:
            continue
        if lookback_col is not None and not _is_true(_cell(row, lookback_col)):
            continue
        qty, _ = _number(_cell(row, qty_col))
        qty = qty or 0.0
        product = get_product(asin, model)
        product.source_presence.add(SOURCE_B2B)
        product.b2b_source_rows += 1
        product.own_transit_units += qty
        stats.rows_used += 1
        traces.append(SourceTraceRecord(
            run_id=run_id, product_key=product.product_key, product_key_type=product.product_key_type,
            final_asin=product.asin, final_model=product.model_number, source_type=SOURCE_B2B,
            source_file=sheet.path.name, source_sheet=sheet.sheet_name, source_row=index,
            source_business_key=asin or model, quantity_used=qty, value_used=None, date_used=None,
            trace_notes="Own in-transit (B2B dispatch) quantity.",
        ))


def _aggregate_po(
    run_id: str,
    sheet: LoadedSheet,
    get_product,
    traces: list[SourceTraceRecord],
    issues: list[InventoryCoverValidationIssue],
    stats: SourceStats,
) -> None:
    if not sheet.rows:
        return
    headers = sheet.headers
    asin_col = first_present_column(headers, ASIN_CANDIDATES)
    model_col = first_present_column(headers, MODEL_CANDIDATES)
    open_po_col = first_present_column(headers, PO_FINAL_CHAIN)
    end_col = first_present_column(headers, ("Window End", "Expected Date"))

    stats.rows_read = sheet.row_count
    if open_po_col is None:
        stats.warnings.append("No reliable open PO quantity column found; Open PO Quantity set to 0.")
    for index, row in enumerate(sheet.rows, start=2):
        asin = _text(row, asin_col)
        model = _text(row, model_col)
        if not asin and not model:
            continue
        qty = 0.0
        if open_po_col is not None:
            value, _ = _number(_cell(row, open_po_col))
            qty = value or 0.0
        product = get_product(asin, model)
        product.source_presence.add(SOURCE_PO)
        product.po_source_rows += 1
        product.open_po_units += qty
        stats.rows_used += 1
        end = coerce_date(_cell(row, end_col))
        stats.period_end = _max_date(stats.period_end, end)
        traces.append(SourceTraceRecord(
            run_id=run_id, product_key=product.product_key, product_key_type=product.product_key_type,
            final_asin=product.asin, final_model=product.model_number, source_type=SOURCE_PO,
            source_file=sheet.path.name, source_sheet=sheet.sheet_name, source_row=index,
            source_business_key=asin or model, quantity_used=qty, value_used=None, date_used=end,
            trace_notes=f"Open PO from '{open_po_col}'." if open_po_col else "No open PO column.",
        ))


def _apply_global_sales_window(products, sales_stats: SourceStats) -> None:
    """Show the evaluated sales window even for products with no sales row."""

    if sales_stats.period_start is None or sales_stats.period_end is None:
        return
    for product in products:
        if product.sales_period_start is None:
            product.sales_period_start = sales_stats.period_start
        if product.sales_period_end is None:
            product.sales_period_end = sales_stats.period_end


def _finalize_identity(
    run_id: str,
    product: ProductCoverRow,
    asin_lookup: dict[str, dict[str, Any]],
    config: Any,
    issues: list[InventoryCoverValidationIssue],
) -> None:
    attrs = asin_lookup.get(product.asin) if product.asin else None
    if attrs is not None:
        product.source_presence.add(SOURCE_ASIN_MASTER)
        product.asin_master_match_status = "Matched"
        if attrs.get("sku"):
            product.model_number = attrs["sku"]
        if attrs.get("brand"):
            product.brand = attrs["brand"]
        if attrs.get("brand_name"):
            product.brand_name = attrs["brand_name"]
        if attrs.get("vendor"):
            product.vendor = attrs["vendor"]
        if attrs.get("main_category"):
            product.main_category = attrs["main_category"]
        if attrs.get("sub_category"):
            product.sub_category = attrs["sub_category"]
        if attrs.get("target") is not None:
            product.aligned_doh_target = attrs["target"]
    else:
        product.asin_master_match_status = "No ASIN Master" if not asin_lookup else "Unmatched"

    if product.aligned_doh_target is None or product.aligned_doh_target <= 0:
        product.target_doh = float(config.default_target_doh)
        if product.aligned_doh_target is None:
            product.calculation_warning_notes.append("Aligned DOH Target missing; default target used.")
    else:
        product.target_doh = float(product.aligned_doh_target)

    if not product.brand_name:
        product.brand_name = product.brand

    if not product.asin:
        issues.append(_issue(run_id, "WARNING", "MISSING_ASIN", product, "", None, None, None,
                             None, "Product has no ASIN; kept by Model Number / SKU.",
                             "Row retained with Model Number key."))


# ---------------------------------------------------------------------------
# Small helpers.
# ---------------------------------------------------------------------------
def _set_if_blank(product: ProductCoverRow, attr: str, value: str) -> None:
    if value and not getattr(product, attr):
        setattr(product, attr, value)


def _is_true(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return False
    return str(value).strip().lower() in {"true", "1", "yes", "y"}


def _max_date(current: date | None, candidate: date | None) -> date | None:
    if candidate is None:
        return current
    if current is None:
        return candidate
    return max(current, candidate)


def _min_date(current: date | None, candidate: date | None) -> date | None:
    if candidate is None:
        return current
    if current is None:
        return candidate
    return min(current, candidate)


def _issue(
    run_id: str,
    severity: str,
    issue_type: str,
    product: ProductCoverRow,
    source_type: str,
    sheet: LoadedSheet | None,
    row: int | None,
    field_name: str | None,
    raw_value: Any,
    detail: str | None,
    action: str | None,
) -> InventoryCoverValidationIssue:
    return InventoryCoverValidationIssue(
        run_id=run_id,
        severity=severity,
        issue_type=issue_type,
        product_key=product.product_key,
        asin=product.asin,
        model_number=product.model_number,
        source_type=source_type,
        source_file=sheet.path.name if sheet else "",
        source_sheet=sheet.sheet_name if sheet else "",
        source_row=row,
        field_name=field_name or "",
        raw_value=raw_value,
        issue_detail=detail or "",
        action_taken=action or "",
    )
