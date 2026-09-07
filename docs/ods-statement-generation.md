# ODS Cash Statement Generation

## Status

This document records the September 2026 feasibility investigation into restoring native OpenDocument Spreadsheet (`.ods`) Cash Statement generation in TCExports.

The work is deliberately deferred while the product team concentrates on tax and accounts integration. No production commitment or implementation schedule is implied by this document. The intended outcome is to preserve enough technical evidence and design direction to resume the work for the Accounts Mode launch without repeating the investigation.

## Executive conclusion

Native ODS Cash Statement generation is feasible. The current implementation is substantially beyond a proof of concept: it creates structurally valid ODS documents for all four synthetic fixture scenarios, and LibreOffice can open and convert them.

The remaining problem is not the ODS format or LibreOffice. It is that the repository currently contains two independent Cash Statement engines:

- `CashStatementExcelHandler` implements the statement in C# using the shared C# SQL repository and ClosedXML.
- `cash_statement_ods.py` independently implements database access, period mapping, financial aggregation, category-tree formulas, cached-value calculation, and ODS presentation in Python.

Those implementations have drifted. Several deterministic defects make the ODS figures unreliable. Continuing to patch the Python implementation section by section would preserve that structural risk. The recommended solution is one format-neutral Cash Statement model and calculation pipeline, followed by thin XLSX and ODS renderers.

## Scope of the investigation

The WebHarness was exercised against these synthetic fixture databases:

- `tcNodeDb4-COMIPFVT1-COMIN26` — company, minimal template
- `tcNodeDb4-COSIPFVT1-COSTD26` — company, standard template
- `tcNodeDb4-STMIPFVT1-STMIN26` — sole trader, minimal template
- `tcNodeDb4-STSIPFVT1-STSTD26` — sole trader, standard template

The export options enabled bank balances, tax accruals, VAT detail, the balance sheet, and equity reconciliation. Active future periods and the order book were excluded.

Both the XLSX and ODS results were recalculated through the installed LibreOffice engine and converted to XLSX for a like-for-like structural and numerical comparison. The converted files were diagnostic artefacts only and are not repository deliverables.

## Current implementation

### Dispatch

`CashStatementHandler` selects the renderer from the request format:

- `excel` delegates to `CashStatementExcelHandler`.
- `libre` delegates to `CashStatementLibreHandler`.

### Libre process boundary

`CashStatementLibreHandler`:

1. Serializes the complete export payload to a temporary JSON file.
2. Starts a global `python` executable as a child process.
3. Runs `python/exporters/cash_statement_ods.py`.
4. Receives a filename and base64-encoded ODS document through standard output.
5. Deletes the temporary payload file.

The Python exporter uses:

- `pyodbc` for direct SQL Server access;
- `odfdo` to construct the ODS document;
- `lxml` and ZIP/XML post-processing for styles, borders, cached values, and document normalization;
- a separate Python repository implementation in `python/data/sqlserver_repository.py`.

The exporter is approximately 1,500 lines before the supporting repository, style factory, localization, and rendering modules are counted.

### Existing strengths

The current ODS work should not be discarded. It already demonstrates and implements:

- native ODS package creation;
- a Cash Flow worksheet with the expected period grid;
- category and cash-code sections;
- category summaries and total categories;
- Cash Expressions Analysis rows;
- bank balances;
- VAT recurrence and VAT period sections;
- a balance sheet and capital row;
- semantic number styles;
- negative-value presentation;
- annual-column borders;
- hidden marker columns;
- frozen panes and active-sheet settings;
- localization hooks;
- post-processing of the ODS XML package.

All four test scenarios produced ODS files after the connection string was supplied in ODBC form. LibreOffice accepted and converted all four documents, confirming that the generated packages are structurally viable.

## Confirmed defects

### 1. Connection string incompatibility

The WebHarness request normally carries a `Microsoft.Data.SqlClient` connection string. The C# repository converts this appropriately, but the Python repository passes it directly to `pyodbc.connect`.

The result is an ODBC Driver Manager `IM002` error because a SQL Client connection string does not specify an ODBC driver in ODBC syntax. During the feasibility test, generation succeeded only after the same credentials were translated in memory to an ODBC connection string.

This is an integration contract defect, not an environmental or ODS defect. A renderer should not reinterpret application connection strings or maintain its own database connection stack.

### 2. Empty categories create circular references

The Python category renderer always creates a `SUM` formula for a category total. When an enabled category has no enabled cash codes, its calculated first cash-code row is the total row itself and its last cash-code row is the preceding row.

This produces reversed ranges such as:

```text
SUM([.D19:.D18])
```

LibreOffice treats the range as including the formula cell, producing a circular reference. The same condition was found and corrected independently in the XLSX renderer.

The ODS renderer should preserve an empty enabled category but write a numeric zero for every period total.

### 3. Total-category formulas depend on display order

`render_totals_formula` builds and registers total rows in a single pass. It can reference only category rows already present in `totals_row_by_category`.

This fails for valid forward references. In the standard sole-trader template:

```text
CT-CUMEXP
  CT-CSTSAL
  CT-STAFFC
  CT-OVERHD
```

`CT-CUMEXP` has display order 55 and `CT-OVERHD` has display order 60. When `CT-CUMEXP` is rendered, the `CT-OVERHD` row has not yet been registered, so the generated formula silently omits Overheads.

The SQL Category Tree is acyclic and correct. The renderer must not use display order as calculation order. It should allocate/register every total row first and populate formulas in a second pass, or evaluate the graph topologically while preserving display order separately.

### 4. The handwritten cached-value evaluator corrupts formulas

ODS formula cells are initially stamped with a cached numeric value of zero. `_post_process_totals_borders` then attempts to calculate cached values for annual-total columns by recognizing a limited set of formula shapes:

- direct cell references;
- horizontal and vertical `SUM` ranges;
- simple lists of added or subtracted references.

When a formula does not match those forms, a permissive reference pattern extracts every cell reference and adds them together. This is incorrect for functions, comparisons, division, multiplication, repeated operands, and conditional expressions.

For example, an annual Gross Margin formula equivalent to:

```text
IF(P8=0;0;P87/P8)
```

should evaluate to approximately `0.80`. The cached evaluator instead combined its cell references and produced a value in the hundreds of thousands. Similar corruption was observed in Net Profit, wage, administration, direct-cost, depreciation, and coverage ratios.

This explains much of the previously reported numerical unreliability. It is not safe to extend the regex evaluator. Cached results should come from the authoritative report calculation model, or the document should explicitly request recalculation by LibreOffice.

### 5. VAT period allocation differs from XLSX

The Python VAT period renderer derives a calendar month with `StartOn.month` and then searches the application month collection for a matching `MonthNumber`. This is not the same algorithm as the current C# renderer and is unsafe for a non-January financial year.

The four-fixture comparison found missing values in the second VAT section even in the company scenarios, where the main category graph did not suffer from circular references. The differences included VAT due, net VAT due, sales excluding VAT, and purchases excluding VAT.

All renderers must consume a single authoritative period-to-column mapping based on the application financial calendar. No renderer should infer a financial column from a calendar month independently.

### 6. Total rows differ between formats

The XLSX renderer creates the total-category rows for the requested cash type and later fills their formulas. The Python path attempts to call `get_categories_by_type` with a string category type even though that method is absent from the current Python SQL repository. It then creates another totals block from all rows in `Cash.vwCategoryTotals`.

Consequences include:

- ODS contains `CT-VAT` in the main total block while XLSX does not;
- row positions diverge after that point;
- responsibility is split between `render_summary_totals_block` and `render_totals_formula`;
- missing repository capabilities are silently treated as empty results.

### 7. Equity reconciliation is not implemented

The Libre request accepts `includeReconciliation`, but the Python exporter neither reads the option nor renders `Cash.vwEquityReconciliationByYear`.

The XLSX statement includes the reconciliation when requested. ODS parity therefore cannot currently be claimed even if all existing ODS values were corrected.

### 8. Runtime and deployment are implicit

The Libre handler starts `python` by name and assumes that compatible versions of Python, `pyodbc`, `odfdo`, `lxml`, an ODBC driver, and their native dependencies are installed and discoverable by the web process.

The project copies the Python source into the build output but does not define or provision its interpreter and package environment. The WebHarness happened to locate a Python installation, but the interactive shell did not resolve `python`, demonstrating that execution currently depends on host-specific process configuration.

Cancellation and timeouts also cross a child-process boundary. Database command timeout behaviour is not consistently propagated through the Python repository.

### 9. The implementations can drift silently

The ODS and XLSX paths separately encode:

- SQL queries and stored-procedure calls;
- DTO shaping;
- financial-calendar alignment;
- active-period behaviour;
- order-book and tax-accrual options;
- category selection and polarity;
- category-tree aggregation;
- expressions;
- VAT layout;
- bank and balance-sheet calculations;
- display structure and styles.

Recent repairs to category selection, VAT accrual column alignment, marker lookup, and empty categories in the C# implementation did not automatically benefit ODS. This duplication is the primary long-term risk.

## Numerical comparison summary

The feasibility comparison was not intended to certify every row, but it was broad enough to establish the pattern:

- Detail structure and cash-code coverage were close to XLSX.
- All four ODS packages were accepted by LibreOffice.
- More than 4,800 numeric cells could be matched between formats in the smallest scenario and more than 7,400 in the largest.
- Company scenarios showed material discrepancies in VAT period rows and annual expression caches.
- Sole-trader scenarios also showed category-total failures caused by empty categories and missing forward references.
- Standard sole trader showed the largest number of mismatches because it exercises the deepest Category Tree, including `CT-CUMEXP`.
- Balance-sheet capital totals were materially incorrect in the sole-trader outputs affected by upstream formula failures.

These results confirm feasibility but explicitly rule out treating the current Libre route as financially reliable.

## Recommended architecture

### Principle

There should be one Cash Statement definition and two serialization targets, not two Cash Statement implementations.

### Shared report model

Introduce a format-neutral model owned by the C# generator. A possible shape is:

```text
CashStatementDocument
  Metadata
  PeriodGrid
    FinancialYear
    Period
    AnnualTotalColumn
  Sections
    CategorySection
    CategorySummary
    CategoryTotals
    Expressions
    BankBalances
    VatRecurrence
    VatPeriods
    BalanceSheet
    EquityReconciliation
  Rows
    SemanticId
    RowKind
    Label/code fields
    Style role
    Period cells
      authoritative value
      optional formula expression
```

The exact types should follow normal C# conventions. The important properties are:

- every row has a stable semantic identifier independent of its worksheet position;
- values are calculated before rendering;
- formula relationships refer to semantic identifiers and periods, not temporary row numbers;
- the financial period grid is constructed once;
- presentation order is separate from dependency order;
- renderer-specific formula syntax is produced only after physical rows and columns are allocated.

### Shared data acquisition

Use `ICashFlowRepository` and its C# SQL Server implementation for both formats. Extend that interface where necessary for reconciliation or missing report inputs.

The Python renderer should not:

- receive database credentials;
- connect to SQL Server;
- reproduce SQL queries;
- infer periods;
- apply business rules;
- update expression status or write event-log records.

If Python remains, it should receive a completed, versioned report model as JSON.

### Category graph preparation

Before rendering:

1. Load every enabled leaf, total, and expression category required by the statement.
2. Load every parent-child relationship.
3. Validate that every referenced node exists.
4. Detect cycles and report their complete paths.
5. Allocate all presentation rows.
6. Resolve total values in dependency order.
7. Produce formulas from semantic dependencies after allocation.

This naturally supports `CT-CUMEXP` and future trees with arbitrary depth or forward references.

### Authoritative values and formulas

The shared model should carry authoritative decimal values for every period cell. Those values support:

- parity tests;
- ODS cached values;
- immediate display without depending on a desktop recalculation;
- reconciliation and diagnostic checks;
- consistent rounding.

Formulas should remain in both workbook formats where they provide useful traceability and interactive recalculation. Cached values must never be derived by parsing arbitrary formula strings with regular expressions.

### Renderer options

Two implementation choices are credible.

#### Option A: retain Python as a thin ODS serializer

Keep the existing `odfdo` and style-factory work, but replace the Python repository and financial logic with a JSON report-model reader.

Advantages:

- preserves most existing ODS and style work;
- uses a mature ODF-oriented Python library;
- reduces the initial rewrite.

Costs:

- Python deployment remains part of the product;
- process invocation and dependency provisioning must be formalized;
- typed C# to JSON to Python contracts need versioning and tests.

#### Option B: implement a C# ODS serializer

Port the useful ODF package and styling techniques into a C# renderer operating directly on the shared model.

Advantages:

- one runtime and deployment model;
- direct use of decimal values and shared types;
- simpler cancellation, logging, and error propagation.

Costs:

- more initial rendering work;
- the existing Python style and XML post-processing code must be translated or replaced;
- a suitable maintained ODF library must be evaluated, or a deliberately small ODF package writer maintained locally.

### Recommendation

Use Option A for the first reliable ODS release unless deployment constraints make Python unacceptable. The existing Python renderer and style factory represent valuable work. Restricting Python to deterministic serialization removes the source of the financial discrepancies while minimizing rework.

The report-model contract should make a later move to a C# serializer possible without changing financial logic or acceptance tests.

## Proposed delivery phases

### Phase 1: define parity and freeze the experimental route

- Keep `libre` explicitly experimental and disabled in production.
- Preserve the four current fixture databases as the initial parity suite.
- Capture expected XLSX semantic rows and authoritative values.
- Define the rounding and comparison tolerance for money, percentages, and ratios.
- Add tests that fail on missing rows, duplicate semantic identifiers, cycles, and unresolved references.

### Phase 2: extract the shared statement model

- Move period-grid construction and section composition out of `CashStatementExcelHandler`.
- Build the complete format-neutral document using `ICashFlowRepository`.
- Include all request options, including reconciliation.
- Add category-graph validation and value calculation.
- Keep XLSX output behaviour stable while extraction proceeds.

### Phase 3: make XLSX a renderer

- Render the shared model through ClosedXML.
- Translate semantic references into Excel A1 formulas only after row allocation.
- Confirm that the output matches the current repaired XLSX statement.
- Use this phase to prove the shared model before changing ODS.

### Phase 4: reduce Python to ODS serialization

- Pass the completed report model to the Python process.
- Remove `pyodbc`, the Python SQL repositories, and connection strings from the ODS path.
- Remove duplicated period, VAT, category, balance-sheet, and expression calculations.
- Remove the regex cached-value evaluator.
- Adapt the existing ODS style factory and XML normalization to the shared row model.
- Store authoritative cached values and native OpenFormula expressions.

### Phase 5: operational hardening

- Define and provision a supported Python runtime and locked package versions.
- Avoid relying on an ambient `python` executable.
- Add startup health checks for the optional ODS feature.
- Propagate cancellation and bounded execution time through the child process.
- Capture standard error without leaking connection information.
- Ensure temporary payloads contain no database credentials; preferably send only the report model through a bounded stream or protected temporary file.
- Add deterministic cleanup and concurrent-request tests.

### Phase 6: Accounts Mode release qualification

- Enable ODS only after the full acceptance matrix passes.
- Exercise the export through the actual Accounts Mode UI and authorization path.
- Open the native ODS output in a supported LibreOffice version.
- Verify fresh open, recalculation, save, reopen, printing, and PDF export.
- Confirm that XLSX remains unchanged.

## Acceptance matrix

Each of the four fixtures should be tested with at least these option profiles:

1. Standard launch profile: bank balances, VAT details, balance sheet, reconciliation, and tax accruals enabled.
2. Minimal profile: optional sections disabled.
3. Active-period profile: active periods enabled.
4. Forecast profile: order book enabled where meaningful.

For each profile, verify:

### Structure

- same financial years, months, and annual columns;
- same enabled categories and cash codes;
- same section presence and order;
- stable semantic identifiers;
- no duplicate or missing identifiers;
- no ghost worksheets or malformed repeated rows.

### Calculations

- cash-code values agree exactly at stored decimal precision;
- category totals agree to the defined monetary tolerance;
- nested totals agree, including `CT-CUMEXP`;
- empty enabled categories return zero without circular formulas;
- VAT recurrence and period values use the correct financial columns;
- bank closing balances agree;
- balance-sheet rows and capital agree;
- equity reconciliation agrees and remains within tolerance;
- expression results agree at the defined ratio tolerance;
- no unresolved references, formula errors, or calculation cycles.

### Native LibreOffice behaviour

- opens without a repair warning;
- displays authoritative values before manual recalculation;
- full recalculation preserves the same results;
- save and reopen preserve results and formulas;
- number formats, negative values, totals borders, hidden markers, widths, and frozen panes remain usable;
- printing and PDF output are legible;
- no external links, macros, or prompts are introduced.

### Service behaviour

- normal ASP.NET connection configuration works without ODBC syntax in the request;
- authorization behaviour matches XLSX;
- cancellation stops database and renderer work;
- timeouts are enforced;
- parallel exports do not collide on temporary paths;
- failures return bounded, actionable diagnostics without secrets;
- the ODS feature can be disabled without affecting XLSX.

## Suggested automated tests

### Unit tests

- financial-period to column mapping for non-January year starts;
- empty category rendering;
- polarity application;
- graph cycle detection;
- forward and backward total references;
- arbitrary-depth total trees;
- missing child references;
- expression token resolution;
- division-by-zero behaviour;
- decimal rounding and cached values;
- renderer formula syntax translation.

### Contract tests

- serialize the same small report model to XLSX and ODS;
- compare semantic rows and authoritative values;
- assert formulas reference the intended semantic cells;
- assert the ODS package contains valid required parts and MIME ordering;
- assert no database connection information enters the ODS payload.

### Integration tests

- generate both formats from each fixture database;
- recalculate ODS with headless LibreOffice;
- read the recalculated result and compare every semantic numeric cell with the shared model;
- compare the XLSX result with the same model;
- run equity reconciliation independently against the database;
- retain concise mismatch output keyed by semantic row and financial period.

## Licensing and product positioning

The preference for an open spreadsheet format is aligned with the GNU-licensed product and with users who choose LibreOffice. It also gives Accounts Mode a useful non-Microsoft document path.

Licensing and format should nevertheless be treated as separate engineering decisions. The licensing compatibility of every runtime and library must be recorded before release, including transitive and native dependencies. This document makes no legal determination.

An ODS implementation should be described as supported only after it passes the same financial controls as XLSX. Format openness cannot compensate for uncertain accounting figures.

## Work explicitly not recommended

- Do not continue expanding the regex cached-formula evaluator.
- Do not add another database query to Python when the C# repository already supplies the concept.
- Do not make display order control calculation dependencies.
- Do not compare rows solely by worksheet position.
- Do not hide empty categories merely to avoid zero handling.
- Do not require callers to know whether a renderer uses SQL Client or ODBC.
- Do not certify output by checking only that LibreOffice opens it.
- Do not enable the current Libre route for production financial reporting.

## Resumption checklist

When this project is resumed:

1. Re-read this document and the current `requirements.md`.
2. Confirm the four fixture databases or equivalent reproducible scenarios still exist.
3. Regenerate the repaired XLSX statements as the current behavioural baseline.
4. Decide whether Python will remain as the initial thin ODS serializer.
5. Write the shared report-model contract and parity-test specification before implementation.
6. Extract the C# statement builder while holding XLSX output stable.
7. Replace Python database and calculation logic with report-model serialization.
8. Run the complete acceptance matrix before exposing ODS in Accounts Mode.

## Relevant implementation locations

- `src/TCExports.Generator/Handlers/CashStatementHandler.cs`
- `src/TCExports.Generator/Handlers/CashStatementExcelHandler.cs`
- `src/TCExports.Generator/Handlers/CashStatementLibreHandler.cs`
- `src/TCExports.Generator/Data/ICashFlowRepository.cs`
- `src/TCExports.Generator/Data/SqlServerCashFlowRepository*.cs`
- `src/TCExports.Generator/python/exporters/cash_statement_ods.py`
- `src/TCExports.Generator/python/data/sqlserver_repository.py`
- `src/TCExports.Generator/python/style_factory/`

## Final assessment

The ODS path should be resumed after the tax/accounts integration reaches an appropriate pause point. The existing output and styling work are reusable, the failure modes are understood, and the four fixtures provide a credible test base.

The key condition is architectural: resume it as a shared-model and renderer-parity project. On that basis, a dependable native LibreOffice Cash Statement is a realistic Accounts Mode launch objective.
