# Python exporters (ODS) — Status and resume guide

This folder contains the experimental LibreOffice ODS exporters and the Style Factory used to format `.ods` outputs. Work on the Cash Statement ODS is paused.

## Current status
- Style Factory: implemented and documented in `style_factory/README.md`.
- Cash Statement ODS: generation works structurally but has unresolved:
  - load-time formula caching (e.g., `SUM(P26:P39)*-1` showing last ref until user edit),
  - immediate negative formatting for formula cells,
  - totals column border persistence on spacer rows (addressed via column defaults and post-process).
- Excel (XLSX) Cash Statement: operational via C# handler.

## Where to pick up
- Exporter: `exporters/cash_statement_ods.py`
- Post-process: `_post_process_totals_borders()` handles:
  - totals column default borders,
  - cloning bordered variants for styled cells,
  - spacer row normalization,
  - cached value stamping for common formula patterns,
  - enforcing POS/NEG bordered CASH styles based on cached values.
- Test harness: `exporters/test_ods_export.py` demonstrates column defaults, bordered clones, and formula cached-value stamping.

## Known issues to resolve
- Formula cached values: ensure all SUM down-column cases set `office:value-type="float"` + `office:value` correctly at load time.
- Negative display: after caching values, enforce applied CASH `*_POS_CELL`/`*_NEG_CELL` (bordered variants) so parentheses/red appear immediately.
- Keep Style Factory maps intact: generators should use neutral `CASHx_CELL` where possible; Style Factory provides `style:map` and number styles.

## Recommended approach when resuming
1) Validate with `test_ods_export.py`
   - Confirm borders across spacer rows remain intact.
   - Confirm negatives render immediately for formula cells once cached values are stamped.

2) Apply fixes in `cash_statement_ods.py`
   - Limit post-process to:
     - cached value stamping (direct refs, SUM across row, SUM down column, +/- refs),
     - style enforcement for CASH cells only,
     - totals column defaults and bordered clones.
   - Avoid row-level formatting until rendering and borders are stable.

3) Re-enable ODS export path in the web app behind a feature flag, then test with representative data.

