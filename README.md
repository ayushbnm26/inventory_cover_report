# Inventory Cover Reporting Engine

This project is a local Python reporting engine for Amazon Vendor Central inventory-cover automation.

The implemented pipelines are:

1. Pipeline 1: Amazon PO Items consolidation.
2. Pipeline 2: B2B Dispatch Tracker backend pipeline.
3. Pipeline 3: Amazon Vendor Central Sales & Inventory backend pipeline.
4. Pipeline 4: Final Inventory Cover Calculation Engine.

Pipeline 4 is the core calculation layer. It consumes the latest backend outputs of the three independent source pipelines and produces the final inventory cover report with Daily Run Rate (DRR), Days On Hand (DOH), transit cover, Days Of Cover (DOC), total supply cover, target gap, cover bucket, and cover alert. It is loosely coupled: it reads the source backend artifacts through a stable sheet/column interface contract and never mutates them.

## Pipeline Overview

### Pipeline 1: PO Items

Pipeline 1 reads multiple Amazon PO Items Excel exports, standardizes the rows, validates the run, preserves backend traceability, and writes:

- A simple team workbook for sharing and PO item clubbing.
- A backend audit workbook for logs, validation, duplicates, source traceability, and run review.

Input folder:

```text
data/incoming/po_items/
```

Latest outputs:

```text
data/processed/po_items/latest/PO_Items_Team_Workbook_latest.xlsx
data/processed/po_items/latest/PO_Items_Backend_Audit_latest.xlsx
```

### Pipeline 2: B2B Dispatch Tracker

Pipeline 2 reads the daily B2B Dispatch Tracker workbook, extracts only RK, Clicktech, and Etrade dispatch sheets, filters rows by dispatch date, and writes a backend-only audit workbook for later inventory-cover reporting.

Input folder:

```text
data/incoming/b2b_dispatch/
```

Latest output:

```text
data/processed/b2b_dispatch/latest/B2B_Dispatch_Backend_Audit_latest.xlsx
```

Pipeline 2 does not create a team-facing workbook.

### Pipeline 3: Sales & Inventory

Pipeline 3 reads Amazon Vendor Central sales and inventory Excel exports, standardizes their columns, parses report metadata, validates gently, preserves source traceability, and writes backend-only audit artifacts for the future calculation engine.

Input folders:

```text
data/incoming/sales/
data/incoming/inventory/
data/reference/sales_inventory_mapping/
```

The mapping folder is optional. If exactly one `.xlsx` workbook is present, Pipeline 3 uses it as an ASIN/SKU reference for blank identifiers only. The expected mapping columns are `ASIN`, `SKU`, and optionally `Master SKU`.

Run outputs:

```text
runs/<RUN_ID>/outputs/sales_inventory/Sales_Backend_Audit_<RUN_ID>.xlsx
runs/<RUN_ID>/outputs/sales_inventory/Inventory_Backend_Audit_<RUN_ID>.xlsx
runs/<RUN_ID>/outputs/sales_inventory/Sales_Inventory_Run_Summary_<RUN_ID>.xlsx
```

Latest outputs:

```text
data/processed/sales_inventory/latest/Sales_Backend_Audit_latest.xlsx
data/processed/sales_inventory/latest/Inventory_Backend_Audit_latest.xlsx
data/processed/sales_inventory/latest/Sales_Inventory_Run_Summary_latest.xlsx
```

Pipeline 3 is backend-only. It does not create a team-facing workbook and does not calculate inventory-cover metrics.

### Pipeline 4: Final Inventory Cover Calculation Engine

Pipeline 4 reads the latest backend artifacts of Pipelines 1-3, builds a product-level universe (keyed by ASIN, falling back to Model Number / SKU), calculates inventory cover, and writes two professional workbooks. The visible team report contains Excel formulas in the calculation cells and also embeds cached calculated results, so it does not appear blank in Excel previewers or sessions that have not recalculated formulas yet. The team workbook also includes a hidden `Formula_Audit` sheet with the same formula view for traceability.

Source latest files consumed (interface contract):

```text
data/processed/po_items/latest/PO_Items_Backend_Audit_latest.xlsx        (sheet: PO_Items_Master)
data/processed/b2b_dispatch/latest/B2B_Dispatch_Backend_Audit_latest.xlsx (sheet: B2B_Dispatch_Master)
data/processed/sales_inventory/latest/Sales_Backend_Audit_latest.xlsx     (sheet: Sales_Master)
data/processed/sales_inventory/latest/Inventory_Backend_Audit_latest.xlsx (sheet: Inventory_Master)
```

Optional reference. If absent, the team workbook displays the configured default target (30 by default) instead of leaving the target column blank, and the backend audit records the missing-reference warning:

```text
data/reference/master_data/ASIN_Master.xlsx                               (sheet: ASIN_Master)
```

Output workbooks:

```text
runs/<RUN_ID>/outputs/inventory_cover/Inventory_Cover_Report_<RUN_ID>.xlsx
runs/<RUN_ID>/outputs/inventory_cover/Inventory_Cover_Backend_Audit_<RUN_ID>.xlsx
data/processed/inventory_cover/latest/Inventory_Cover_Report_latest.xlsx
data/processed/inventory_cover/latest/Inventory_Cover_Backend_Audit_latest.xlsx
```

Team workbook sheets. The team-facing report intentionally omits backend-only diagnostic columns such as `Data Quality Flag`; those remain in the audit workbook. `Formula_Audit` is hidden by default and can be unhidden when formula traceability is needed:

```text
Inventory_Cover_Report
Formula_Audit
Critical
High_Risk
Watch
Near_Target
Healthy
No_Sales
Formula_Guide
Source_Summary
```

Backend audit workbook sheets:

```text
Inventory_Cover_Master
Source_Row_Trace
Source_Summary
Validation_Issues
Calculation_Audit
Formula_Guide
Run_Metadata
```

Formula logic (engraved into the main report cells and `Formula_Audit`, all divide-by-zero safe):

```text
Sales Days              = MIN(window, Sales Period End - Sales Period Start + 1)  (window default 30)
Daily Run Rate          = Sales Units / Sales Days
Current Stock DOH       = Sellable On Hand Units / DRR
Stock + Amazon Transit  = (On Hand + Amazon In-Transit) / DRR
Stock + Own Transit     = (On Hand + Own In-Transit) / DRR
Total Transit DOH       = (On Hand + Amazon In-Transit + Own In-Transit) / DRR
DOC Including Open PO    = (On Hand + Open PO + Own In-Transit) / DRR
Total Supply Cover DOH  = (On Hand + Open PO + Amazon In-Transit + Own In-Transit) / DRR
Gap to Target Units     = MAX(Target DOH * DRR - (On Hand + Amazon + Own + Open PO), 0)
Target DOH              = Aligned DOH Target if available else configured default target (30)
```

When DRR is zero, the DOH/DOC columns show `No Sales` instead of dividing by zero.

Cover bucket thresholds (based on Total Supply Cover DOH):

```text
< 5 days      Critical      Immediate action
5–15 days     High Risk     Urgent replenishment
15–25 days    Watch         Plan replenishment
25–30 days    Near Target   Monitor
> 30 days     Healthy       No immediate action
No DRR        No Sales      No sales in selected period
```

Blank numeric policy: blank `Sales Units`, `Sellable On Hand Units`, `Amazon In-Transit Units`, `Own In-Transit Units`, and `Open PO Quantity` are treated as zero for calculation only. Raw source values are preserved in the source backend workbooks and traced in the backend audit workbook.

Products with no sales row still show the latest selected sales period when a sales source exists, so the team can see the period being evaluated. They remain classified as `No Sales` because DRR is zero.

Run the engine on the latest source outputs:

```bash
python -m inventory_cover.cli run-inventory-cover
python scripts/run_inventory_cover_pipeline.py
```

Run the entire project end-to-end (all source pipelines, then the engine):

```bash
python -m inventory_cover.cli run-full-inventory-cover
```

## Automated Email Delivery

The engine can optionally email the run-specific team workbook after Pipeline 4 completes successfully. Email delivery is an explicit final layer: it does not change calculations, source artifacts, workbook formulas, sheet names, or latest-copy behavior. No email is sent unless `--send-email` is provided.

Only the team-facing workbook is attached by default:

```text
runs/<RUN_ID>/outputs/inventory_cover/Inventory_Cover_Report_<RUN_ID>.xlsx
```

Create a local `.env` file from `.env.example` and keep it private. Never commit `.env` or real SMTP credentials.

Required email settings:

```text
SMTP_HOST=
SMTP_PORT=
SMTP_USERNAME=
SMTP_PASSWORD=
SMTP_FROM=
SMTP_USE_TLS=true
SMTP_USE_SSL=false
SMTP_TIMEOUT_SECONDS=30

REPORT_EMAIL_TO=
REPORT_EMAIL_CC=
REPORT_EMAIL_BCC=
REPORT_EMAIL_REPLY_TO=
REPORT_EMAIL_SUBJECT_PREFIX=Inventory Cover Report
```

Example `.env` with placeholders only:

```text
SMTP_HOST=smtp.example.com
SMTP_PORT=587
SMTP_USERNAME=<smtp-username>
SMTP_PASSWORD=<smtp-password-or-app-password>
SMTP_FROM=<sender@example.com>
SMTP_USE_TLS=true
SMTP_USE_SSL=false
SMTP_TIMEOUT_SECONDS=30

REPORT_EMAIL_TO=<recipient@example.com>
REPORT_EMAIL_CC=
REPORT_EMAIL_BCC=
REPORT_EMAIL_REPLY_TO=
REPORT_EMAIL_SUBJECT_PREFIX=Inventory Cover Report
```

Dry-run the email layer without connecting to SMTP:

```bash
python -m inventory_cover.cli run-inventory-cover --send-email --email-dry-run
python -m inventory_cover.cli run-full-inventory-cover --send-email --email-dry-run
```

Send through SMTP after a successful run:

```bash
python -m inventory_cover.cli run-inventory-cover --send-email
python -m inventory_cover.cli run-full-inventory-cover --send-email
```

Gmail and some hosted mail providers usually require an app password or SMTP-specific credential instead of the normal account password. Use the provider's current SMTP host, port, and TLS/SSL guidance.

Each attempted send writes traceability artifacts under the Pipeline 4 run folder:

```text
runs/<RUN_ID>/notifications/email_delivery.json
runs/<RUN_ID>/logs/email_delivery.log
```

The audit JSON records the run ID, recipients, sanitized SMTP settings, subject, attachment path, report context dates, timestamps, dry-run status, and any sanitized error. It never records SMTP passwords.

Troubleshooting:

| Symptom | Likely cause | Action |
| --- | --- | --- |
| `SMTP_HOST is required` | `.env` missing or incomplete | Set `SMTP_HOST` in `.env` or environment. |
| `SMTP_PORT must be an integer` | Invalid port value | Use a numeric port such as `587` or `465`. |
| `SMTP_USE_TLS and SMTP_USE_SSL cannot both be true` | Conflicting security mode | Use STARTTLS or SSL, not both. |
| `SMTP_PASSWORD is required for network email delivery` | Real send requested without password | Add the SMTP password or app password outside Git. |
| `SMTP_AUTHENTICATION_FAILED` | Bad username/password or provider policy | Verify credentials, app-password setup, and sender account. |
| `SMTP_RECIPIENT_REJECTED` | Recipient blocked or malformed | Check `REPORT_EMAIL_TO`, `REPORT_EMAIL_CC`, and `REPORT_EMAIL_BCC`. |
| `Attachment does not exist` | Report generation did not produce the run workbook | Inspect the Pipeline 4 output path and run log. |

Source-failure behaviour for the full command is configurable:

```bash
python -m inventory_cover.cli run-full-inventory-cover --fail-fast
python -m inventory_cover.cli run-full-inventory-cover --continue-on-source-warning
python -m inventory_cover.cli run-full-inventory-cover --skip-source-pipelines
```

Useful engine options:

```bash
python -m inventory_cover.cli run-inventory-cover --sales-window-days 30
python -m inventory_cover.cli run-inventory-cover --default-target-doh 30
python -m inventory_cover.cli run-inventory-cover --asin-master-path data/reference/master_data/ASIN_Master.xlsx
python -m inventory_cover.cli run-inventory-cover --strict-freshness
python -m inventory_cover.cli run-inventory-cover --blank-numeric-policy zero_for_calculation
```

What the report does and does not do: it is a cover snapshot that prioritises replenishment and escalation from the latest available source data. It does not forecast demand, create purchase orders automatically, model lead times, or set pricing.

## Commands

List available pipelines:

```bash
python -m inventory_cover.cli list-pipelines
```

Run all source pipelines from one centralized command:

```bash
python -m inventory_cover.cli run-source-pipelines
python scripts/run_source_pipelines.py
```

Run all source pipelines in parallel:

```bash
python -m inventory_cover.cli run-source-pipelines --parallel
python scripts/run_source_pipelines.py --parallel
```

Parallel mode isolates run folders under:

```text
runs/source_pipelines/po_items/
runs/source_pipelines/b2b_dispatch/
runs/source_pipelines/sales_inventory/
```

This avoids metadata filename collisions while still refreshing the normal latest processed outputs.

Pipeline 1:

```bash
python -m inventory_cover.cli run-po-items --input-dir data/incoming/po_items
python scripts/run_po_items_pipeline.py --input-dir data/incoming/po_items
```

Pipeline 2:

```bash
python -m inventory_cover.cli run-b2b-dispatch --input-dir data/incoming/b2b_dispatch
python scripts/run_b2b_dispatch_pipeline.py --input-dir data/incoming/b2b_dispatch
```

Pipeline 3:

```bash
python -m inventory_cover.cli run-sales-inventory \
  --sales-input-dir data/incoming/sales \
  --inventory-input-dir data/incoming/inventory \
  --mapping-input-dir data/reference/sales_inventory_mapping
```

```bash
python scripts/run_sales_inventory_pipeline.py \
  --sales-input-dir data/incoming/sales \
  --inventory-input-dir data/incoming/inventory \
  --mapping-input-dir data/reference/sales_inventory_mapping
```

Useful Pipeline 3 options:

```bash
python scripts/run_sales_inventory_pipeline.py --require-sales
python scripts/run_sales_inventory_pipeline.py --require-inventory
python scripts/run_sales_inventory_pipeline.py --allow-multiple-sales-files
python scripts/run_sales_inventory_pipeline.py --allow-multiple-inventory-files
python scripts/run_sales_inventory_pipeline.py --dedupe-exact-rows
python scripts/run_sales_inventory_pipeline.py --log-level DEBUG
```

Useful centralized options:

```bash
python -m inventory_cover.cli run-source-pipelines --min-free-gb 2
python -m inventory_cover.cli run-source-pipelines --parallel --min-free-gb 2
python -m inventory_cover.cli run-source-pipelines --skip-po-items
python -m inventory_cover.cli run-source-pipelines --skip-b2b-dispatch
python -m inventory_cover.cli run-source-pipelines --skip-sales-inventory
python -m inventory_cover.cli run-source-pipelines --fail-fast
python -m inventory_cover.cli run-source-pipelines --dedupe-exact-rows
```

The centralized command also accepts source-specific paths and options:

```bash
python -m inventory_cover.cli run-source-pipelines \
  --po-input-dir data/incoming/po_items \
  --b2b-input-dir data/incoming/b2b_dispatch \
  --sales-input-dir data/incoming/sales \
  --inventory-input-dir data/incoming/inventory \
  --mapping-input-dir data/reference/sales_inventory_mapping
```

## Run Folder Pattern

Every pipeline creates a timestamped run folder:

```text
runs/<RUN_ID_YYYYMMDD_HHMMSS>/
```

Pipeline 3 creates:

```text
runs/<RUN_ID>/inputs/sales_inventory/sales/
runs/<RUN_ID>/inputs/sales_inventory/inventory/
runs/<RUN_ID>/inputs/sales_inventory/mapping/
runs/<RUN_ID>/outputs/sales_inventory/
runs/<RUN_ID>/logs/sales_inventory_pipeline.log
runs/<RUN_ID>/metadata/run_metadata.json
runs/<RUN_ID>/validation/sales_inventory_validation_issues.json
runs/<RUN_ID>/validation/sales_inventory_duplicates.json
```

Pipeline 4 creates:

```text
runs/<RUN_ID>/inputs/inventory_cover/
runs/<RUN_ID>/outputs/inventory_cover/
runs/<RUN_ID>/logs/inventory_cover_pipeline.log
runs/<RUN_ID>/metadata/run_metadata.json
runs/<RUN_ID>/validation/inventory_cover_validation_issues.json
```

The exact source backend workbooks consumed by Pipeline 4 are copied into `runs/<RUN_ID>/inputs/inventory_cover/` for traceability.

Raw input workbooks are copied into the run folder. Original source workbooks are not mutated.

## Pipeline 3 Expected Columns

Expected sales columns:

```text
ASIN
Child Vendor Code
Product Title
Brand Code
Brand
Category
Subcategory
Parent ASIN
UPC
EAN
ISBN
Model Number
Store Code
MSRP
Binding
Colour
Release Date
Replenishment Code
Shipped Revenue
Shipped COGS
Shipped Units
Customer Returns
Confirmed Units
Sales Discount
Contra-COGS
Net PPM %
ASIN Confirmation %
```

Expected inventory columns:

```text
ASIN
Child Vendor Code
Product Title
Brand Code
Brand
Category
Subcategory
Parent ASIN
UPC
EAN
ISBN
Model Number
Store Code
MSRP
Binding
Colour
Release Date
Replenishment Code
Vendor Confirmation %
Net Received
Net Received Units
Open Purchase Order Quantity
Receive Fill %
Overall Vendor Lead Time (days)
Aged 90+ Days Sellable Inventory
Aged 90+ Days Sellable Units
Sellable On-Hand Inventory
Sellable On Hand Units
Unsellable On-Hand Inventory
Unsellable On-Hand Units
Confirmed Units
Sales Discount
Contra-COGS
In Transit Quantity
Sellable In Transit Units
Unsellable In Transit Units
```

## Validation Philosophy

Validation is intentionally gentle. Critical structural failures stop the run, but harmless Amazon export changes are warned and preserved where possible.

Pipeline 3:

- Scans the first 20 rows for headers instead of hardcoding a header row.
- Ignores temporary Excel files beginning with `~$`.
- Processes sales-only or inventory-only runs by default when one source is missing.
- Fails when both sources are missing.
- Fails on multiple files in one source folder unless the matching allow-multiple flag is passed.
- Warns for missing expected optional columns.
- Logs unknown extra source columns and ignores them.
- Preserves ASIN, UPC, EAN, ISBN, Model Number, and Child Vendor Code as text.
- Keeps rows with ASIN.
- Fills blank Model Number from ASIN when a unique mapping exists.
- Fills blank ASIN from Model Number/SKU when a unique mapping exists.
- Keeps rows without ASIN when Model Number and Product Title are present and no unique mapping exists.
- Rejects only rows where ASIN, Model Number, and Product Title are all blank.
- Parses most numeric blanks as blank, not zero.
- Defaults blank inventory stock/in-transit fields to zero for `Sellable On-Hand Inventory`, `Sellable On Hand Units`, `In Transit Quantity`, `Sellable In Transit Units`, and `Unsellable In Transit Units`.
- Audits duplicates and keeps them by default.
- Drops only exact normalized duplicate rows when `--dedupe-exact-rows` is enabled.

## Output Workbooks

Pipeline 3 Sales backend workbook sheets:

```text
Sales_Master
Run_Summary
File_Audit
Validation_Issues
Duplicates
Mapping_Audit
Processing_Guide
```

Pipeline 3 Inventory backend workbook sheets:

```text
Inventory_Master
Run_Summary
File_Audit
Validation_Issues
Duplicates
Mapping_Audit
Processing_Guide
```

The `Mapping_Audit` sheet records the optional ASIN/SKU mapping workbook that was loaded, how many rows were usable, and whether any ambiguous keys were skipped.

The `Processing_Guide` sheet explains what the workbook is, which raw report was processed, where the raw input was copied, how headers were detected, how identifiers may be enriched from mapping, how dates and numbers were parsed, how rows and duplicates were handled, and how the future calculation engine will consume the artifact.

## Tests

Run:

```bash
python -m pytest
```

The tests create synthetic Excel fixtures and do not depend on private business files.

## Git Hygiene

Business inputs and generated artifacts are intentionally ignored:

```text
data/incoming/
data/reference/
data/processed/
runs/
logs/
context_dumps/
*.xlsx
*.xls
*.xlsm
*.log
.env
```

Do not commit raw business Excel files.

## Context Dump

To create a pasteable project context file for review:

```bash
python scripts/dump_py_to_txt.py
```

The dump is written to `context_dumps/` and excludes business data, run outputs, caches, logs, and binary Excel files.
