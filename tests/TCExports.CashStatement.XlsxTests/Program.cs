using System.Collections.Immutable;
using ClosedXML.Excel;
using TCExports.Generator.CashStatement.Model;
using TCExports.Generator.Excel;

var tests = new (string Name, Action Run)[]
{
    ("cash-flow layout preserves legacy row gaps and period columns", CashFlowLayout),
    ("reconciliation layout preserves its blank spacer column", ReconciliationLayout),
    ("formula translation preserves contiguous range and polarity shapes", AggregateFormulas),
    ("formula translation binds opaque expression references", TerminalFormula),
    ("formula translation rejects an address absent from layout", MissingReference),
    ("renderer emits the accepted sheet skeleton without credentials", RendererSkeleton)
};

var failures = new List<string>();
foreach (var test in tests)
{
    try
    {
        test.Run();
        Console.WriteLine($"PASS {test.Name}");
    }
    catch (Exception ex)
    {
        failures.Add($"FAIL {test.Name}: {ex.Message}");
    }
}

foreach (var failure in failures)
    Console.Error.WriteLine(failure);
if (failures.Count > 0)
    return 1;

Console.WriteLine($"Cash Statement XLSX tests passed: {tests.Length}/{tests.Length}.");
return 0;

static void CashFlowLayout()
{
    var fixture = Fixture();
    var layout = new CashStatementXlsxLayoutAllocator().Allocate(fixture.Document, CashStatementLimits.Default);

    Physical(layout, fixture.DetailMonth1, "Cash Flow", 7, 4);
    Physical(layout, fixture.DetailAnnual, "Cash Flow", 7, 6);
    Physical(layout, fixture.TotalMonth1, "Cash Flow", 8, 4);
    Physical(layout, fixture.SummaryMonth1, "Cash Flow", 11, 4);
}

static void ReconciliationLayout()
{
    var fixture = Fixture();
    var layout = new CashStatementXlsxLayoutAllocator().Allocate(fixture.Document, CashStatementLimits.Default);

    Physical(layout, fixture.ReconciliationTolerance, "Reconciliation", 2, 2);
    Physical(layout, fixture.ReconciliationOverallStatus, "Reconciliation", 2, 5);
    Physical(layout, fixture.ReconciliationYear, "Reconciliation", 5, 1);
    Physical(layout, fixture.ReconciliationBridge, "Reconciliation", 5, 13);
    Physical(layout, fixture.ReconciliationDifference, "Reconciliation", 5, 14);
    Physical(layout, fixture.ReconciliationStatus, "Reconciliation", 5, 15);
}

static void AggregateFormulas()
{
    var fixture = Fixture();
    var layout = new CashStatementXlsxLayoutAllocator().Allocate(fixture.Document, CashStatementLimits.Default);
    var translator = new CashStatementXlsxFormulaTranslator();

    var annual = new StatementFormulaAggregate(
        StatementAggregateOperator.Sum,
        ImmutableArray.Create<StatementFormula>(
            new StatementFormulaReference(fixture.DetailMonth1),
            new StatementFormulaReference(fixture.DetailMonth2)));
    Equal("=SUM(D7:E7)", translator.Translate(annual, layout), "annual formula");

    var expense = new StatementFormulaUnary(
        StatementUnaryOperator.Negate,
        new StatementFormulaAggregate(
            StatementAggregateOperator.Sum,
            ImmutableArray.Create<StatementFormula>(new StatementFormulaReference(fixture.DetailMonth1))));
    Equal("=SUM(D7:D7)*-1", translator.Translate(expense, layout), "expense total formula");

    var nonContiguous = new StatementFormulaAggregate(
        StatementAggregateOperator.Sum,
        ImmutableArray.Create<StatementFormula>(
            new StatementFormulaReference(fixture.DetailMonth1),
            new StatementFormulaReference(fixture.SummaryMonth1)));
    Equal("=D7+D11", translator.Translate(nonContiguous, layout), "calculated total formula");
}

static void TerminalFormula()
{
    var fixture = Fixture();
    var layout = new CashStatementXlsxLayoutAllocator().Allocate(fixture.Document, CashStatementLimits.Default);
    var terminal = new StatementFormulaTerminal(
        "IF([CC-SALES]=0,0,[CC-SALES]/2)",
        ImmutableDictionary<string, StatementCellAddress>.Empty
            .Add("CC-SALES", fixture.DetailMonth1));

    Equal(
        "=IF(D7=0,0,D7/2)",
        new CashStatementXlsxFormulaTranslator().Translate(terminal, layout),
        "terminal formula");
}

static void MissingReference()
{
    var missing = new StatementCellAddress(
        new StatementSheetId("cash-flow"),
        new StatementRowId("missing"),
        new StatementColumnId("period:2026:01"));
    var exception = Catch<InvalidOperationException>(() =>
        new CashStatementXlsxFormulaTranslator().Translate(
            new StatementFormulaReference(missing),
            new StatementLayout(ImmutableDictionary<StatementCellAddress, PhysicalCellAddress>.Empty)));
    Require(exception.Message.Contains("absent from the XLSX layout", StringComparison.Ordinal),
        "missing-reference diagnostic was not stable");
}

static void RendererSkeleton()
{
    var bytes = new CashStatementXlsxRenderer().Render(Fixture().Document);
    Require(bytes.Length > 0, "renderer returned an empty package");

    using var stream = new MemoryStream(bytes);
    using var workbook = new XLWorkbook(stream);
    Equal("Cash Flow", workbook.Worksheet(1).Name, "first sheet");
    Equal("Settings", workbook.Worksheet(2).Name, "second sheet");
    Equal("Reconciliation", workbook.Worksheet(3).Name, "third sheet");

    var cashFlow = workbook.Worksheet("Cash Flow");
    Equal("Cash Statement", cashFlow.Cell("A1").GetString(), "statement title");
    Equal("Example", cashFlow.Cell("A2").GetString(), "company name");
    Equal("CC-SALES", cashFlow.Cell("A7").GetString(), "detail code");
    Equal(10m, cashFlow.Cell("D7").GetValue<decimal>(), "detail value");
    Equal(0m, cashFlow.Cell("D8").GetValue<decimal>(), "empty category value");
    Require(!cashFlow.Cell("D8").HasFormula, "empty category emitted an invalid native formula");
    Equal("Reference not found: [CC-LABOUR]", cashFlow.Cell("D14").GetString(), "expression diagnostic");
    Require(cashFlow.Cell("D14").Style.Font.FontColor == XLColor.Red, "expression diagnostic is not red");
    Require(cashFlow.Cell("E14").IsEmpty(), "expression diagnostic was repeated across the row");
    Require(cashFlow.Column(3).IsHidden, "marker column is visible");

    var settings = workbook.Worksheet("Settings");
    Require(settings.Cell("A4").IsEmpty() && settings.Cell("B4").IsEmpty(),
        "credential-removal spacer row is not blank");
    Require(!settings.CellsUsed().Any(cell => cell.GetString().Contains("Server=", StringComparison.OrdinalIgnoreCase)),
        "settings sheet contains a SQL connection string");

    var reconciliation = workbook.Worksheet("Reconciliation");
    Equal("Submission checks (Does it add up?)", reconciliation.Cell("A1").GetString(), "reconciliation title");
    Require(reconciliation.Cell("L5").IsEmpty(), "reconciliation spacer column is populated");
    Equal("G5+H5+J5", reconciliation.Cell("M5").FormulaA1, "bridge formula");
    Equal("K5-M5", reconciliation.Cell("N5").FormulaA1, "difference formula");
    Require(reconciliation.Cell("O5").FormulaA1.StartsWith("IF(ABS(N5)", StringComparison.Ordinal),
        "status formula is absent");
}

static FixtureData Fixture()
{
    var cashSheetId = new StatementSheetId("cash-flow");
    var month1 = StatementPeriodId.Month(2026, 1);
    var month2 = StatementPeriodId.Month(2026, 2);
    var annual = StatementPeriodId.Annual(2026);
    var month1Column = new StatementColumnId($"period:{month1}");
    var month2Column = new StatementColumnId($"period:{month2}");
    var annualColumn = new StatementColumnId($"period:{annual}");
    var detailId = new StatementRowId("flow:sales:cash-code:cc-sales");
    var totalId = new StatementRowId("flow:sales:total");
    var summaryId = new StatementRowId("summary:trade:sales");

    ImmutableDictionary<StatementColumnId, StatementCell> PeriodCells() =>
        ImmutableDictionary<StatementColumnId, StatementCell>.Empty
            .Add(month1Column, Money(10m))
            .Add(month2Column, Money(20m))
            .Add(annualColumn, Money(30m));

    ImmutableDictionary<StatementColumnId, StatementCell> EmptyExpenseCells() =>
        new[] { month1Column, month2Column, annualColumn }.ToImmutableDictionary(
            column => column,
            _ => Money(0m) with
            {
                Formula = new StatementFormulaUnary(
                    StatementUnaryOperator.Negate,
                    new StatementFormulaAggregate(
                        StatementAggregateOperator.Sum,
                        ImmutableArray<StatementFormula>.Empty))
            });

    ImmutableDictionary<StatementColumnId, StatementCell> InvalidExpressionCells() =>
        new[] { month1Column, month2Column, annualColumn }.ToImmutableDictionary(
            column => column,
            _ => new StatementCell(
                StatementValue.Unsupported(StatementScalarKind.Decimal, "Reference not found: [CC-LABOUR]"),
                StatementNumericKind.Percentage,
                StatementPresentationRole.AnalysisValue,
                NumberFormat: "0%"));

    var cashSheet = new StatementSheet(
        cashSheetId,
        StatementSheetKind.CashFlow,
        0,
        "Cash Flow",
        ImmutableArray.Create(
            new StatementColumn(new StatementColumnId("code"), "Code", StatementColumnKind.Identifier, 0),
            new StatementColumn(new StatementColumnId("label"), "Name", StatementColumnKind.Label, 1),
            new StatementColumn(new StatementColumnId("marker"), string.Empty, StatementColumnKind.HiddenMarker, 2),
            new StatementColumn(month1Column, "JAN", StatementColumnKind.Period, 3, PeriodId: month1),
            new StatementColumn(month2Column, "FEB", StatementColumnKind.Period, 4, PeriodId: month2),
            new StatementColumn(annualColumn, "TOTALS", StatementColumnKind.AnnualPeriod, 5, PeriodId: annual)),
        ImmutableArray.Create(
            new StatementSection(
                new StatementSectionId("flow:trade"),
                StatementSectionKind.FlowCategories,
                0,
                "Trade",
                StatementPresentationRole.SectionHeader,
                ImmutableArray.Create(
                    new StatementRow(new StatementRowId("flow:trade:sales:header"), StatementRowKind.Header, 0, "Sales", "SALES", StatementPresentationRole.SectionHeader, ImmutableDictionary<StatementColumnId, StatementCell>.Empty),
                    new StatementRow(detailId, StatementRowKind.Detail, 1, "Sales", "CC-SALES", StatementPresentationRole.Detail, PeriodCells()),
                    new StatementRow(totalId, StatementRowKind.CategoryTotal, 2, "Totals", "SALES", StatementPresentationRole.ExpenseTotal, EmptyExpenseCells()))),
            new StatementSection(
                new StatementSectionId("summary:trade"),
                StatementSectionKind.FlowSummary,
                1,
                "Summary",
                StatementPresentationRole.SummaryHeader,
                ImmutableArray.Create(
                    new StatementRow(summaryId, StatementRowKind.SummaryReference, 0, "Sales", "SALES", StatementPresentationRole.SummaryTotal, PeriodCells()))),
            new StatementSection(
                new StatementSectionId("analysis"),
                StatementSectionKind.Analysis,
                2,
                "Analysis",
                StatementPresentationRole.AnalysisHeader,
                ImmutableArray.Create(
                    new StatementRow(new StatementRowId("expression:ce-lm"), StatementRowKind.Expression, 0, "Labour Margin %", "CE-LM", StatementPresentationRole.AnalysisValue, InvalidExpressionCells())))));

    var reconciliationId = new StatementSheetId("reconciliation");
    var summaryRowId = new StatementRowId("reconciliation:summary");
    var yearRowId = new StatementRowId("reconciliation:2026");
    var profitAfterTax = new StatementCellAddress(reconciliationId, yearRowId, new StatementColumnId("profit-after-tax"));
    var capitalMovement = new StatementCellAddress(reconciliationId, yearRowId, new StatementColumnId("capital-movement"));
    var openingAccount = new StatementCellAddress(reconciliationId, yearRowId, new StatementColumnId("opening-account-position"));
    var capitalDelta = new StatementCellAddress(reconciliationId, yearRowId, new StatementColumnId("capital-delta"));
    var bridge = new StatementCellAddress(reconciliationId, yearRowId, new StatementColumnId("bridge-total"));
    var difference = new StatementCellAddress(reconciliationId, yearRowId, new StatementColumnId("difference"));
    var reconciliationCells = ImmutableDictionary<StatementColumnId, StatementCell>.Empty
        .Add(new StatementColumnId("year"), Count(2026m))
        .Add(new StatementColumnId("description"), Text("2026-27"))
        .Add(profitAfterTax.ColumnId, Money(90m))
        .Add(capitalMovement.ColumnId, Money(5m))
        .Add(openingAccount.ColumnId, Money(5m))
        .Add(capitalDelta.ColumnId, Money(100m))
        .Add(bridge.ColumnId, Money(100m) with
        {
            Formula = new StatementFormulaAggregate(
                StatementAggregateOperator.Sum,
                ImmutableArray.Create<StatementFormula>(
                    new StatementFormulaReference(profitAfterTax),
                    new StatementFormulaReference(capitalMovement),
                    new StatementFormulaReference(openingAccount)))
        })
        .Add(difference.ColumnId, Money(0m) with
        {
            Formula = new StatementFormulaBinary(
                StatementBinaryOperator.Subtract,
                new StatementFormulaReference(capitalDelta),
                new StatementFormulaReference(bridge))
        })
        .Add(new StatementColumnId("status"), Text("PASS"));
    var reconciliationSheet = new StatementSheet(
        reconciliationId,
        StatementSheetKind.Reconciliation,
        1,
        "Reconciliation",
        ImmutableArray.Create(
            ReconColumn("year", "Year", 0),
            ReconColumn("description", "Description", 1),
            ReconColumn("opening-capital", "OpeningCapital", 2),
            ReconColumn("closing-capital", "ClosingCapital", 3),
            ReconColumn("profit", "Profit", 4),
            ReconColumn("business-tax", "BusinessTax", 5),
            ReconColumn("profit-after-tax", "ProfitAfterTax", 6),
            ReconColumn("capital-movement", "CapitalMovement", 7),
            ReconColumn("opening-subject-position", "OpeningSubjectPosition", 8),
            ReconColumn("opening-account-position", "OpeningAccountPosition", 9),
            ReconColumn("capital-delta", "CapitalDelta", 10),
            ReconColumn("bridge-total", "BridgeTotal", 11),
            ReconColumn("difference", "Difference", 12),
            ReconColumn("tolerance", "Tolerance", 13),
            ReconColumn("status", "Status", 14)),
        ImmutableArray.Create(new StatementSection(
            new StatementSectionId("reconciliation"),
            StatementSectionKind.Reconciliation,
            0,
            "Reconciliation",
            StatementPresentationRole.ReconciliationHeader,
            ImmutableArray.Create(
                new StatementRow(
                    summaryRowId,
                    StatementRowKind.Metadata,
                    0,
                    "Submission checks (Does it add up?)",
                    null,
                    StatementPresentationRole.ReconciliationPass,
                    ImmutableDictionary<StatementColumnId, StatementCell>.Empty
                        .Add(new StatementColumnId("tolerance"), Money(0.10m))
                        .Add(new StatementColumnId("status"), Text("PASS"))),
                new StatementRow(yearRowId, StatementRowKind.Reconciliation, 1, "2026-27", "2026", StatementPresentationRole.ReconciliationPass, reconciliationCells)))));

    var document = new CashStatementDocument(
        new CashStatementMetadata("Example", "Cash Statement", "owner", DateTimeOffset.UnixEpoch, month1),
        new CashStatementRequestV1(IncludeReconciliation: true),
        new StatementPeriodGrid(ImmutableArray.Create(
            new StatementPeriod(month1, 0, "JAN", "2026-27", "Active", new DateOnly(2026, 1, 1), true),
            new StatementPeriod(month2, 1, "FEB", "2026-27", "Active", new DateOnly(2026, 2, 1), false),
            new StatementPeriod(annual, 2, "TOTALS", "2026-27", "Active", null, false))),
        ImmutableArray.Create(cashSheet, reconciliationSheet));

    return new FixtureData(
        document,
        new(cashSheetId, detailId, month1Column),
        new(cashSheetId, detailId, month2Column),
        new(cashSheetId, detailId, annualColumn),
        new(cashSheetId, totalId, month1Column),
        new(cashSheetId, summaryId, month1Column),
        new(reconciliationId, summaryRowId, new StatementColumnId("tolerance")),
        new(reconciliationId, summaryRowId, new StatementColumnId("status")),
        new(reconciliationId, yearRowId, new StatementColumnId("year")),
        new(reconciliationId, yearRowId, new StatementColumnId("bridge-total")),
        new(reconciliationId, yearRowId, new StatementColumnId("difference")),
        new(reconciliationId, yearRowId, new StatementColumnId("status")));

    static StatementColumn ReconColumn(string id, string label, int order) =>
        new(new StatementColumnId(id), label, StatementColumnKind.Attribute, order);
}

static StatementCell Money(decimal value) => new(StatementValue.FromDecimal(value), StatementNumericKind.Money);
static StatementCell Count(decimal value) => new(StatementValue.FromDecimal(value), StatementNumericKind.Count);
static StatementCell Text(string value) => new(StatementValue.FromText(value));

static void Physical(StatementLayout layout, StatementCellAddress semantic, string sheet, int row, int column)
{
    Require(layout.Cells.TryGetValue(semantic, out var actual), $"mapping absent: {semantic}");
    Equal(new PhysicalCellAddress(sheet, row, column), actual, $"physical mapping {semantic}");
}

static TException Catch<TException>(Action action) where TException : Exception
{
    try { action(); }
    catch (TException exception) { return exception; }
    throw new InvalidOperationException($"Expected {typeof(TException).Name}.");
}

static void Equal<T>(T expected, T actual, string message)
{
    if (!EqualityComparer<T>.Default.Equals(expected, actual))
        throw new InvalidOperationException($"{message}: expected '{expected}', actual '{actual}'");
}

static void Require(bool condition, string message)
{
    if (!condition)
        throw new InvalidOperationException(message);
}

sealed record FixtureData(
    CashStatementDocument Document,
    StatementCellAddress DetailMonth1,
    StatementCellAddress DetailMonth2,
    StatementCellAddress DetailAnnual,
    StatementCellAddress TotalMonth1,
    StatementCellAddress SummaryMonth1,
    StatementCellAddress ReconciliationTolerance,
    StatementCellAddress ReconciliationOverallStatus,
    StatementCellAddress ReconciliationYear,
    StatementCellAddress ReconciliationBridge,
    StatementCellAddress ReconciliationDifference,
    StatementCellAddress ReconciliationStatus);
