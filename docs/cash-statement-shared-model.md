# Cash Statement shared-model contract

**Status:** Accepted for Phase 3 implementation  
**Contract version:** 1  
**Date:** 15 October 2026

Accepted on 15 October 2026. The contract remains amendable when implementation evidence identifies an omitted current behaviour; amendments are reviewed before dependent work continues.

## 1. Decision

The Cash Statement builder produces one immutable, renderer-neutral `CashStatementDocument`. XLSX and ODS renderers receive that document and do not query SQL, interpret request strings, determine financial periods, evaluate categories, or calculate financial results.

The model deliberately describes semantic sheets, sections, rows, columns, values, dependencies, and presentation roles. It does not contain Excel row numbers, Excel column letters, ODF addresses, native formula strings, ClosedXML types, ODF XML, or database connection details.

The initial interface skeleton is in `src/TCExports.Generator/CashStatement/Model`.

## 2. Current workbook inventory

The accepted all-options baselines contain three worksheets in this order:

1. `Cash Flow`;
2. `Settings`; and
3. `Reconciliation`.

The largest Phase 0 workbook is `Cash Flow!A1:BC220`; the widest is `Cash Flow!A1:BP196`. Across the four baselines the current implementation produces at most 11,278 used/styled cells, 4,468 formulas, and an 87,184-byte XLSX package.

### Cash Flow elements

| Presentation family | Current source | Value/formula behaviour | Model owner |
|---|---|---|---|
| Statement title, company, generation date | `RenderHeader` | Text values; generation date is volatile | Metadata section and `CashStatementMetadata` |
| Financial years, months, annual columns, active highlight | `RenderPeriodGrid` | One monthly column per configured month and one annual column per year | `StatementPeriodGrid` and semantic period columns |
| Trade categories and cash-code detail | `RenderCategoriesAsync(Trade)` | Repository values; annual sum; expense polarity applied to category total | Builder-owned flow sections |
| Trade summary | `RenderSummaryAfterCategories` | References category totals, then sums the summary | Builder-owned summary section |
| Money categories and summary | `RenderCategoriesAsync(Money)` | Active periods, order book, and tax accruals are always disabled | Builder-owned flow and summary sections |
| Trade calculated totals | `RenderSummaryTotalsBlockAsync(Trade)` and `RenderTotalsFormulaAsync` | Sum semantic child categories | Category dependency graph |
| Tax categories and summary | `RenderCategoriesAsync(Tax)` | Active periods and tax accrual options apply; order book is disabled | Builder-owned flow and summary sections |
| Tax calculated totals | `RenderSummaryTotalsBlockAsync(Tax)` and `RenderTotalsFormulaAsync` | Sum semantic child categories | Category dependency graph |
| Analysis expressions | `RenderExpressionsAsync` | Expressions refer to statement rows by readable labels | Opaque management formula with stable code bindings and renderer-native evaluation |
| Closing bank balances | `RenderClosingBankBalanceAsync` | Monthly account balances; annual value references December; company sum | Bank section |
| VAT recurrence detail | `RenderVatRecurrenceTotalsAsync` | Nine VAT metrics, optional accrual addition, annual sum | VAT recurrence section |
| VAT monthly totals | `RenderVatPeriodTotalsAsync` | Nine VAT metrics, optional accrual addition, annual sum | VAT period section |
| Balance sheet | `RenderBalanceSheetAsync` | Asset/month values; annual value is last populated month; capital sum | Balance-sheet section |

The disabled `RenderEquityReconciliationAsync` method is not an accepted output element and is not carried into contract version 1. The separate `Reconciliation` worksheet is accepted and is included.

### Supporting-sheet elements

| Sheet | Current content | Model treatment |
|---|---|---|
| Settings | Document type, format, user, SQL connection, normalized string parameters | Safe request metadata and typed options. A SQL connection string is never placed in the shared model. Preserving the current exposed connection string would require a separately approved security exception. |
| Reconciliation | Tolerance, overall status, yearly accounting inputs, bridge, difference, yearly status | Builder computes authoritative bridge, difference, and status values. Native formulas remain optional traceability. |

## 3. Repository call catalogue

The current Excel path makes these repository calls while rendering. Phase 3 moves their orchestration to the builder; no renderer may call them.

| Data family | Current calls | Proposed model destination |
|---|---|---|
| Identity/calendar | `GetCompanyNameAsync`, `GetActivePeriodAsync`, `GetActiveYearsAsync`, `GetMonthsAsync` | Metadata and period grid |
| Flow categories | `GetCategoriesAsync`, `GetCashCodesAsync`, `GetCashCodeValuesAsync` | Flow sections, detail rows, authoritative period cells |
| Category totals | `GetCategoriesByTypeAsync`, `GetCategoryTotalsAsync`, `GetCategoryTotalCodesAsync` | Category graph and calculated-total rows |
| Expressions | `GetCategoryExpressionsAsync`, `GetCategoryCodeFromNameAsync`, `SetCategoryExpressionStatusAsync` | Opaque management expressions, stable reference bindings, and validation diagnostics |
| Bank | `GetBankAccountsAsync`, `GetBankBalancesAsync` | Bank rows and company total |
| VAT | `GetVatRecurrenceTypeAsync`, `GetVatRecurrenceAsync`, `GetVatRecurrenceAccrualsAsync`, `GetVatPeriodTotalsAsync`, `GetVatPeriodAccrualsAsync` | VAT recurrence and period sections |
| Balance sheet | `GetBalanceSheetAsync` | Asset rows and capital total |
| Reconciliation | `GetEquityReconciliationByYearAsync` | Reconciliation sheet |

The current repository method shapes remain inputs, not shared-model types. Phase 3 may extend `ICashFlowRepository` only where an authoritative input is absent.

## 4. Identity and ordering

Every sheet, section, row, and column has a stable semantic identifier. Identifiers are case-insensitive canonical lowercase strings when constructed by the builder. They must not contain a physical row number, column number, or format name.

Canonical patterns are:

| Entity | Pattern | Example |
|---|---|---|
| Sheet | fixed semantic name | `cash-flow`, `settings`, `reconciliation` |
| Section | `<family>:<qualifier>` | `flow:trade`, `summary:tax`, `vat:recurrence` |
| Cash-code row | `flow:<category-code>:cash-code:<cash-code>` | `flow:fees:cash-code:fees-001` |
| Category total | `flow:<category-code>:total` | `flow:fees:total` |
| Summary reference | `summary:<cash-type>:<category-code>` | `summary:trade:fees` |
| Summary period total | `summary:<cash-type>:period-total` | `summary:trade:period-total` |
| Calculated total | `calculated-total:<category-code>` | `calculated-total:gross-profit` |
| Expression | `expression:<category-code>` | `expression:gross-margin` |
| Bank account | `bank:<account-code>` | `bank:current` |
| Bank company total | `bank:company-total` | `bank:company-total` |
| VAT metric | `vat:<recurrence-or-period>:<metric>` | `vat:period:net-vat-due` |
| Balance-sheet asset | `balance-sheet:<asset-code>` | `balance-sheet:cash` |
| Capital | `balance-sheet:capital` | `balance-sheet:capital` |
| Reconciliation year | `reconciliation:<year-number>` | `reconciliation:2026` |

`PresentationOrder` controls display only. Calculation order is produced by topologically sorting semantic cell dependencies. A child may therefore appear after its parent without changing calculation correctness. Renderers allocate physical rows and columns only after the complete model and dependency graph are valid.

Identifiers are unique within their entity scope; row identifiers are unique across the document so diagnostic paths and formula references are unambiguous.

## 5. Values and financial semantics

### Authoritative values

`StatementCell.Value` is authoritative. `StatementCell.Formula` is optional traceability metadata. A renderer must write the authoritative cached value before writing a native formula and must not parse or evaluate the emitted formula to discover that value.

The scalar kinds are decimal, text, date, and Boolean. Each value has one of three states:

- `Present`: the typed value exists;
- `Missing`: an expected source value does not exist; or
- `Unsupported`: the requested calculation has no defined result.

Present decimal zero is a real value and is never represented as missing. Empty enabled categories produce present zero for every period.

### Money, percentages, and ratios

- Money is stored as `decimal` in source currency units. Addition and subtraction occur without intermediate display rounding.
- Percentage and ratio values are dimensionless decimals. A percentage of 15% is stored as `0.15m`.
- `StatementNumericKind` records whether a decimal is money, a percentage, a ratio, or a count. Presentation roles and native number-format strings do not determine business meaning.
- Source precision is preserved. Any business rounding rule must specify its decimal scale and `MidpointRounding` mode at the calculation that requires it.
- The default deterministic rule for an explicitly rounded business result is `MidpointRounding.AwayFromZero`. Display-only rounding is renderer presentation and does not alter the authoritative value.

### Sign and polarity

Repository leaf values retain their source sign. Category polarity is an explicit builder input. To preserve current Excel behaviour, an expense category total negates the sum of its detail rows; income and neutral category totals do not. Downstream totals consume the already-polarized category value and do not apply polarity again.

### Division by zero and invalid operands

A direct division by present zero produces `Unsupported(Decimal, "division-by-zero")`; it is not silently converted to zero or missing. A conditional expression that selects a non-dividing branch is valid and must not evaluate the unselected division. Arithmetic with missing or unsupported inputs propagates the corresponding state with a diagnostic path.

## 6. Formula and dependency contract

Formula dependencies use semantic `StatementCellAddress` values. Version 1 permits:

- typed literals and cell references;
- unary negation and absolute value;
- addition, subtraction, multiplication, and division;
- equality and ordered comparisons;
- sum and count aggregates; and
- conditional expressions.

There is no raw native-formula escape hatch for core financial calculations. Management expressions are a deliberately separate, terminal contract: their spreadsheet-language body remains opaque, bracketed statement-row references are bound to stable category or cash codes, and the target spreadsheet engine evaluates the resulting native formula. A management expression can never feed a core statement calculation.

### Management-expression references

Square brackets are reserved for references to Cash Statement rows. TCWeb presents the current readable category or cash-code description but persists its stable code. Resolution is deterministic and case-insensitive: category code, cash code, unique category description, then unique cash-code description. This preserves the historic readable formulas while allowing existing description-based databases to migrate on their next successful edit.

An unknown, empty, unclosed, or ambiguous tag is reported with a stable diagnostic. The affected management expression is invalid and is not emitted as a zero-valued formula; the core Cash Statement remains valid because management expressions are terminal management information. Spreadsheet function names, operators, separators, and other non-tag text are not parsed or rewritten by the shared binder.

### Graph validation

Before evaluation or rendering, validation must:

1. reject empty or duplicate semantic identifiers;
2. reject duplicate presentation orders within one parent;
3. verify that every cell column exists on its sheet;
4. verify that every formula reference resolves to a cell;
5. reject self references;
6. detect cycles across arbitrary dependency depth and report the full closed path;
7. topologically order all calculated cells independently of presentation order;
8. reject a graph exceeding the configured node, edge-per-cell, or depth limit; and
9. report all independently detectable violations in one result.

Missing category children are blocking core contract errors. Missing management-expression references invalidate and diagnose that expression but do not block the core statement. Neither kind may be replaced with literal zero.

## 7. Request/options contract

`CashStatementRequestV1` is the internal typed request. Version 1 contains:

| Option | Default | Scope |
|---|---:|---|
| `IncludeActivePeriods` | `false` | Trade, tax, and VAT value inclusion |
| `IncludeOrderBook` | `false` | Trade cash-code values only |
| `IncludeTaxAccruals` | `true` | Tax cash-code and VAT accrual values |
| `IncludeVatDetails` | `true` | VAT recurrence and monthly sections |
| `IncludeBankBalances` | `true` | Bank section |
| `IncludeBalanceSheet` | `true` | Balance-sheet section |
| `IncludeReconciliation` | `false` | Reconciliation sheet |
| `CommandTimeoutSeconds` | `30` | Repository operation; valid range 1–600 |

These defaults preserve `CashPayloadValidator` and current handler behaviour. `includeBankTypes` is intentionally absent. External string parameters are parsed exactly once at the request boundary in Phase 3. Unknown option names and unsupported contract versions fail validation rather than being ignored internally.

## 8. Renderer contract

`ICashStatementLayoutAllocator` maps semantic cells to physical sheet, row, and column coordinates under configured limits. `ICashStatementFormulaTranslator` translates a validated semantic formula tree through that completed layout.

Renderers own:

- physical row/column allocation and native addresses;
- native formula syntax;
- style definitions corresponding to presentation roles;
- widths, heights, hidden marker columns, panes, protection, printing, and package metadata; and
- output serialization.

Renderers do not own:

- option defaults or interpretation;
- repository access or tenant credentials;
- financial period construction;
- category/reference resolution;
- polarity, VAT allocation, rounding, or reconciliation decisions; or
- authoritative values.

The allocator must reject duplicate or absent mappings. The translator must reject a reference not present in the completed layout. It must never infer a target from a label or presentation position.

## 9. Presentation roles

The contract uses semantic roles instead of copying ClosedXML or ODF styles into the model. Version 1 includes roles for the document title, metadata, headers, active and annual periods, detail, polarity totals, summaries, calculated totals, analysis, bank, VAT, balance sheet, reconciliation states, and hidden markers. Numeric meaning is carried separately by `StatementNumericKind`.

The XLSX renderer maps roles to the exact accepted current formatting. The ODS renderer maps the same roles to visually and semantically equivalent ODF styles. A renderer may combine a row role and a cell role, but it may not change the cell's financial meaning.

## 10. Admission and output limits

The initial `CashStatementLimits.Default` values are:

| Limit | Value |
|---|---:|
| Fiscal years | 20 |
| Period columns | 260 |
| Sheets | 3 |
| Sections | 256 |
| Rows | 10,000 |
| Cells | 2,000,000 |
| Formula nodes | 1,000,000 |
| Dependencies per cell | 1,024 |
| Dependency depth | 256 |
| Serialized output | 64 MiB |

The largest baseline has five fiscal years, 220 Cash Flow rows, 68 Cash Flow columns, 11,278 used/styled cells, 4,468 formulas, and an XLSX size below 100 KiB. The initial limits therefore provide substantial growth headroom while preventing unbounded tenant requests. Phase 7 load testing may lower them; raising them requires measured memory/time evidence.

Limits are checked before acquisition where request dimensions are known, after model construction, during layout allocation, and while serializing through a bounded stream. Limit failure produces a stable validation error and no partial document.

## 11. Known decisions requiring explicit acceptance

1. The shared model never contains or emits a SQL connection string. The current Settings worksheet does, so removing it is a security correction rather than byte-for-byte parity.
2. Unresolved management-expression references invalidate that expression instead of becoming formula literal zero; they do not block the core statement.
3. Division by zero is explicitly unsupported unless protected by a conditional branch.
4. The disabled inline equity-reconciliation block is excluded; the accepted Reconciliation worksheet remains included.
5. Native formulas are traceability output for core calculations; terminal management expressions are the explicit exception and are evaluated by Excel or LibreOffice.

These decisions must be accepted with the Phase 2 contract. They are not changes to the current production handler in this phase.

## 12. Phase boundary

Phase 2 adds types and documentation only. It does not route requests through the model, alter `CashStatementExcelHandler`, change the accepted XLSX output, or enable ODS. Parsing, graph validation implementation, builder calculations, and baseline value comparison belong to Phase 3.
