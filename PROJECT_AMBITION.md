# Project Ambition

This project is the foundation of an Amazon Vendor Central Inventory Cover automation system.

The system will combine purchase order, dispatch, sales, and inventory data into a final inventory-cover report. The design principle is loose coupling: each source pipeline produces a clean, validated backend artifact that later calculation code can consume without depending on the internal details of the source export.

## Pipeline Roadmap

1. Amazon PO Items consolidation. — implemented
2. B2B Dispatch Tracker backend pipeline. — implemented
3. Amazon Vendor Central Sales & Inventory backend pipeline. — implemented
4. Final inventory-cover calculation engine. — implemented

Pipelines 1–3 form the source-ingestion foundation. Pipeline 4, the core calculation engine, consumes their latest backend outputs through a stable sheet/column interface contract and produces the final inventory cover report.

## Final Reporting Goal

The final report calculates and explains inventory health using metrics such as:

- DRR: daily run rate.
- DOH: days on hand (current stock, plus Amazon transit, plus own transit, plus total transit).
- DOC: days of cover including open PO.
- Total supply cover DOH.
- Gap to target units.
- Cover bucket and cover alert flags.
- Backend data quality flags and team remarks.

Pipeline 4 implements these calculations. The team workbook engraves the formulas directly into Excel cells (divide-by-zero safe, with explicit `No Sales` handling) and ships annexure sheets per cover bucket, a Formula_Guide, and a Source_Summary without backend-only diagnostic clutter. A backend audit workbook adds data quality flags, full source-row traceability, a calculation audit, validation issues, and run metadata.

Future enhancements may add demand forecasting, recommended PO quantity, and lead-time modelling on top of this engine.
