# Cash Statement model-to-XLSX traceability

**Status:** Accepted for Phase 3 implementation  
**Reference:** accepted Phase 0 all-options XLSX baselines and Phase 1 comparator  
**Date:** 15 October 2026

Accepted by Ian Monnox on 15 October 2026. Newly discovered current-XLSX behaviour must be added here and reviewed before dependent implementation continues.

This matrix makes every characterized current Excel element the responsibility of either the shared model or the XLSX renderer. `Mixed` means the model supplies semantic intent and authoritative content while the renderer supplies native spreadsheet representation.

| Current XLSX element | Current implementation | Shared-model source | XLSX renderer responsibility | Ownership |
|---|---|---|---|---|
| Worksheet order: Cash Flow, Settings, Reconciliation | Main handler and sheet builders | Ordered `StatementSheet` values | Create sheets in model order | Mixed |
| Cash Flow title | `RenderHeader` | Metadata title row/value | Font and placement | Mixed |
| Company name | `RenderHeader` | `CashStatementMetadata.CompanyName` | Placement and font | Mixed |
| Generation date/time | `RenderHeader` | `GeneratedAt` | Current display format; volatile comparison rule | Mixed |
| Code/name headings | `RenderHeader` | Semantic identifier/label columns | A/B placement and header style | Mixed |
| Hidden category marker | Category/expression renderers | Hidden-marker column and business codes | Column C placement, font masking, hidden state | Mixed |
| Sheet protection and locked regions | `RenderHeader` plus section methods | Presentation roles identify protected structure | Native worksheet/cell protection | Renderer |
| Year labels and cash status | `RenderPeriodGrid` | Period grid year label/status | Header placement | Mixed |
| Month labels | `RenderPeriodGrid` | Monthly periods | Header placement | Mixed |
| Active-period highlight | `RenderPeriodGrid` | `StatementPeriod.IsActive` | Yellow native style | Mixed |
| Annual columns | `RenderPeriodGrid` | Annual period entries | Bold, borders, locked state, width | Mixed |
| Trade/Money/Tax section order | `HandleAsync` | Ordered sections | Allocate in model order | Model |
| Category headings | `RenderCategoriesAsync` | Flow sections/category grouping | Bold/top-border role | Mixed |
| Cash-code rows | `RenderCategoriesAsync` | Stable cash-code rows and values | Physical cells and money format | Mixed |
| Cash-code annual formulas | `RenderCategoriesAsync` | Authoritative annual value plus semantic sum | Translate to native `SUM` | Mixed |
| Empty category zero | `RenderCategoriesAsync` | Explicit present zero cells | Serialize numeric zero | Model |
| Category polarity | `RenderCategoriesAsync` | Builder-calculated polarized total | Expense/income/neutral style role | Model |
| Category total formulas | `RenderCategoriesAsync` | Authoritative total plus dependency sum | Translate sum and optional negation | Mixed |
| Flow summary references | `RenderSummaryAfterCategories` | Stable summary-reference rows | Translate direct references | Mixed |
| Flow period total | `RenderSummaryAfterCategories` | Authoritative period total | Translate sum | Mixed |
| Calculated-total rows | `RenderSummaryTotalsBlockAsync` | Stable calculated-total rows | Green-yellow header/row layout | Mixed |
| Total-category dependency formulas | `RenderTotalsFormulaAsync` | Validated category graph and authoritative values | Translate semantic sums | Mixed |
| Analysis labels/codes/formats | `RenderExpressionsAsync` | Expression rows, codes, numeric roles | Native style and conditional red-negative display | Mixed |
| Analysis formulas | `RenderExpressionsAsync` | Parsed formula trees and authoritative results | Native formula translation | Mixed |
| Expression validation errors | `RenderExpressionsAsync` | Builder/graph validation diagnostics | No formula substitution; request fails | Model |
| Bank account rows | `RenderClosingBankBalanceAsync` | Bank rows and values | Physical layout | Mixed |
| Bank annual values | `RenderClosingBankBalanceAsync` | Authoritative December/defined closing value | Optional native direct reference | Mixed |
| Company bank total | `RenderClosingBankBalanceAsync` | Authoritative sum | Optional native sum formula | Mixed |
| VAT recurrence title/type | `RenderVatRecurrenceTotalsAsync` | VAT section title/type | Orchid header style | Mixed |
| VAT recurrence metrics | Same | Nine stable metric rows and values | Layout and numeric style | Mixed |
| VAT recurrence accruals | Same | Builder-applied authoritative values | None beyond serialization | Model |
| VAT recurrence annual totals | Same | Authoritative annual values | Optional native sums | Mixed |
| VAT monthly metrics/accruals/totals | `RenderVatPeriodTotalsAsync` | VAT-period rows and authoritative values | Layout and optional formulas | Mixed |
| Balance-sheet assets | `RenderBalanceSheetAsync` | Stable asset rows and period values | Layout | Mixed |
| Balance-sheet annual value | Same | Authoritative last-populated-period value | Optional native reference | Mixed |
| Capital row | Same | Authoritative asset sum | Native sum and capital style | Mixed |
| Column widths A/B/periods | End of `HandleAsync` | Column semantic kinds | Exact accepted widths | Renderer |
| Freeze panes D5 | End of `HandleAsync` | Cash Flow sheet role | Freeze 4 rows and 3 columns | Renderer |
| Number formats and negative red | Grid and expression methods | Money/percentage/ratio roles | Exact native formats and conditional rules | Renderer |
| Borders and fills | Section methods | Presentation roles | Exact accepted native styles | Renderer |
| Used range and blank styled cells | Emergent ClosedXML output | Complete ordered model/layout | Preserve accepted structure under comparator | Renderer |
| Print settings | ClosedXML defaults/current package | None | Preserve accepted settings exactly | Renderer |
| Defined names/merged ranges | None in baselines | None | Must remain absent unless approved | Renderer |
| Settings document type | `SettingsSheetBuilder` | Safe request metadata | Two-column Settings layout | Mixed |
| Settings format | Same | Renderer output format | Two-column Settings layout | Renderer |
| Settings user | Same | `RequestedBy` | Two-column Settings layout | Mixed |
| Settings SQL connection | Same | Deliberately absent | Remove only as an approved security correction | Approval required |
| Settings parameters | Same | Typed request v1 | Stable names and normalized values | Mixed |
| Reconciliation tolerance | `ReconciliationSheetBuilder` | Reconciliation model/config | Numeric format and location | Mixed |
| Reconciliation yearly inputs | Same | Builder values | Table layout | Mixed |
| Bridge total | Same | Authoritative `ProfitAfterTax + CapitalMovement + OpeningAccountPosition` | Optional native formula | Mixed |
| Difference | Same | Authoritative `CapitalDelta - BridgeTotal` | Optional native formula | Mixed |
| Year status | Same | PASS/WARN/FAIL from tolerance rules | Native IF formula and conditional style | Mixed |
| Overall status | Same | Authoritative aggregation of yearly statuses | Native formula and conditional style | Mixed |
| Reconciliation freeze pane | Same | Reconciliation sheet role | Freeze top four rows | Renderer |
| File name/timestamp | End of `HandleAsync` | Requester/document metadata | Safe filename and timestamp suffix | Renderer/service boundary |
| XLSX package metadata | ClosedXML | None | Volatile allow-list only; unknown differences fail | Renderer |

## Option-to-output traceability

| Typed option | Data/calculation effect | Presentation effect |
|---|---|---|
| `IncludeActivePeriods` | Includes future Trade, Tax, and VAT values | None independently |
| `IncludeOrderBook` | Includes forecast order-book values for Trade cash codes only | None independently |
| `IncludeTaxAccruals` | Includes Tax cash-code and VAT accrual values | None independently |
| `IncludeVatDetails` | Builds VAT recurrence and VAT-period sections | Includes both VAT sections |
| `IncludeBankBalances` | Loads and computes bank account/company totals | Includes bank section |
| `IncludeBalanceSheet` | Loads assets and computes capital | Includes balance-sheet section |
| `IncludeReconciliation` | Loads and computes yearly reconciliation | Includes Reconciliation sheet |
| `CommandTimeoutSeconds` | Bounds each repository command | No workbook effect and must not drive renderer behaviour |

## Formula ownership traceability

| Formula family | Semantic dependencies | Authoritative evaluator | Native formula translator |
|---|---|---|---|
| Annual detail total | Monthly cells in same row/year | Builder | `SUM` over allocated monthly cells |
| Category total | Detail cells in category/period, then polarity | Builder | `SUM`, optional unary negation |
| Summary reference | Category-total cell | Builder | Direct reference |
| Summary period total | Summary rows in period | Builder | `SUM` |
| Calculated category total | Category graph children | Builder after graph validation | References/addition or `SUM` |
| Analysis expression | Parsed category references and operators | Builder expression evaluator | Formula AST translation |
| Bank annual closing | Defined closing month cell | Builder | Direct reference |
| Company bank total | Bank accounts in period | Builder | `SUM` |
| VAT annual total | VAT metric monthly/recurrence cells | Builder | `SUM` |
| Balance-sheet annual closing | Last populated asset period | Builder | Direct reference |
| Capital | Asset cells in period | Builder | `SUM` |
| Reconciliation bridge/difference/status | Named yearly inputs and tolerance | Builder | Arithmetic/`IF`/`ABS`/count translation |

## Completeness rule

Phase 4 may add no workbook element that is absent from this matrix without first updating the Phase 2 contract. Any characterized element not reproduced by the replacement XLSX renderer is a blocking difference unless it has a separately approved correction record.
