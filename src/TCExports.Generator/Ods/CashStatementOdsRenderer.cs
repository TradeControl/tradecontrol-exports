using System.Collections.Immutable;
using System.Globalization;
using TCExports.Generator.CashStatement.Model;

namespace TCExports.Generator.Ods;

/// <summary>
/// Maps the immutable Cash Statement document to a physical ODS workbook without
/// repository access or financial calculation.
/// </summary>
public sealed class CashStatementOdsRenderer
{
    private static readonly IReadOnlyDictionary<string, int> ReconciliationColumns =
        new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["year"] = 1, ["description"] = 2, ["opening-capital"] = 3,
            ["closing-capital"] = 4, ["profit"] = 5, ["business-tax"] = 6,
            ["profit-after-tax"] = 7, ["capital-movement"] = 8,
            ["opening-subject-position"] = 9, ["opening-account-position"] = 10,
            ["capital-delta"] = 11, ["bridge-total"] = 13, ["difference"] = 14,
            ["status"] = 15
        };

    private readonly CashStatementLimits _limits;
    private readonly ICashStatementLayoutAllocator _allocator;
    private readonly OdsPackageWriter _writer;

    public CashStatementOdsRenderer(
        CashStatementLimits? limits = null,
        ICashStatementLayoutAllocator? allocator = null,
        OdsPackageWriter? writer = null)
    {
        _limits = limits ?? CashStatementLimits.Default;
        _allocator = allocator ?? new CashStatementOdsLayoutAllocator();
        _writer = writer ?? new OdsPackageWriter();
    }

    public async Task<byte[]> RenderAsync(CashStatementDocument document, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(document);
        cancellationToken.ThrowIfCancellationRequested();

        // The complete layout is deliberately allocated before any formula is handed
        // to the package writer. Forward references therefore resolve deterministically.
        var layout = _allocator.Allocate(document, _limits);
        var worksheets = ImmutableArray.CreateBuilder<OdsWorksheet>();
        var cashFlow = document.Sheets.Single(sheet => sheet.Kind == StatementSheetKind.CashFlow);
        worksheets.Add(RenderCashFlow(document, cashFlow, layout, cancellationToken));
        worksheets.Add(RenderSettings(document));
        var reconciliation = document.Sheets.SingleOrDefault(sheet => sheet.Kind == StatementSheetKind.Reconciliation);
        if (reconciliation is not null)
            worksheets.Add(RenderReconciliation(reconciliation, layout, cancellationToken));

        var package = new OdsSpreadsheetDocument(worksheets.ToImmutable());
        var odsLimits = new OdsWriteLimits(
            MaximumWorksheets: _limits.MaxSheets,
            MaximumColumnsPerWorksheet: Math.Max(1_024, _limits.MaxPeriodColumns + 3),
            MaximumRows: _limits.MaxRows,
            MaximumCells: _limits.MaxCells,
            MaximumOutputBytes: _limits.MaxOutputBytes);
        using var output = new MemoryStream();
        await _writer.WriteAsync(output, package, odsLimits, cancellationToken).ConfigureAwait(false);
        return output.ToArray();
    }

    private static OdsWorksheet RenderCashFlow(
        CashStatementDocument document,
        StatementSheet sheet,
        StatementLayout layout,
        CancellationToken cancellationToken)
    {
        var columns = sheet.Columns.OrderBy(column => column.PresentationOrder).ToArray();
        var rowCount = Math.Max(4, layout.Cells
            .Where(pair => pair.Key.SheetId == sheet.Id)
            .Select(pair => pair.Value.RowNumber)
            .DefaultIfEmpty(4)
            .Max());
        var grid = new OdsGrid(rowCount, columns.Length);

        grid.Set(1, 1, Text(document.Metadata.Title, StatementPresentationRole.DocumentTitle));
        grid.Set(2, 1, Text(document.Metadata.CompanyName, StatementPresentationRole.Metadata));
        grid.Set(3, 1, Text(document.Metadata.DateLabel, StatementPresentationRole.Metadata));
        grid.Set(3, 2, Text(document.Metadata.GeneratedAt.ToString("dd MMM HH:mm:ss", CultureInfo.InvariantCulture),
            StatementPresentationRole.Metadata));

        foreach (var column in columns)
        {
            var columnNumber = column.PresentationOrder + 1;
            grid.Set(4, columnNumber, Text(column.Label, StatementPresentationRole.ColumnHeader));
            if (column.PeriodId is not { } periodId)
                continue;
            var period = document.PeriodGrid.Periods.Single(item => item.Id == periodId);
            if (period.Id.IsAnnual)
                grid.Set(3, columnNumber, Text(period.YearLabel, StatementPresentationRole.ColumnHeader));
            else if (document.PeriodGrid.Periods
                     .Where(item => item.Id.YearNumber == period.Id.YearNumber && !item.Id.IsAnnual)
                     .OrderBy(item => item.PresentationOrder)
                     .First().Id == period.Id)
                grid.Set(3, columnNumber, Text($"{period.YearLabel} ({period.CashStatus})", StatementPresentationRole.ColumnHeader));
        }
        grid.ApplyRoleToRow(4, StatementPresentationRole.ColumnHeader);
        grid.ApplyFormatToRow(4, topBorder: OdsBorderWeight.Thin, bottomBorder: OdsBorderWeight.Thick);

        var currentRow = 4;
        foreach (var section in sheet.Sections.OrderBy(section => section.PresentationOrder))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (section.Kind == StatementSectionKind.FlowCategories)
            {
                foreach (var row in section.Rows.OrderBy(row => row.PresentationOrder))
                {
                    if (row.Kind == StatementRowKind.Header)
                    {
                        currentRow += 2;
                        grid.Set(currentRow, 1, Text(row.Label, StatementPresentationRole.SectionHeader));
                        grid.ApplyRoleToRow(currentRow, StatementPresentationRole.SectionHeader);
                        grid.ApplyFormatToRow(currentRow, bold: true, topBorder: OdsBorderWeight.Medium);
                        continue;
                    }
                    currentRow++;
                    RenderCashFlowRow(grid, sheet, row, currentRow, layout);
                }
                continue;
            }

            currentRow += 2;
            grid.Set(currentRow, 1, Text(section.Title.ToUpperInvariant(), section.PresentationRole));
            grid.ApplyRoleToRow(currentRow, section.PresentationRole);
            grid.ApplyFormatToRow(currentRow,
                bold: true,
                locked: section.Kind != StatementSectionKind.CategoryTotals,
                topBorder: OdsBorderWeight.Thin,
                bottomBorder: OdsBorderWeight.Thin);
            var firstSectionRow = currentRow + 1;
            foreach (var row in section.Rows.OrderBy(row => row.PresentationOrder))
            {
                currentRow++;
                RenderCashFlowRow(grid, sheet, row, currentRow, layout);
            }
            FinishSection(grid, section, firstSectionRow, currentRow);
        }

        grid.ApplyFormatToColumn(2, rightBorder: OdsBorderWeight.Thick);
        for (var column = 1; column <= 3; column++)
            grid.ApplyFormatToColumn(column, locked: true);
        for (var row = 1; row <= 4; row++)
            grid.ApplyFormatToRow(row, locked: true);
        foreach (var column in columns)
        {
            var columnNumber = column.PresentationOrder + 1;
            grid.ApplyColumnContext(
                columnNumber,
                activePeriod: column.PresentationRole == StatementPresentationRole.CurrentPeriod,
                annualPeriod: column.Kind == StatementColumnKind.AnnualPeriod);
        }

        var odsColumns = columns.Select(column => column.Kind switch
        {
            StatementColumnKind.Identifier => new OdsColumn(2.4m),
            StatementColumnKind.Label => new OdsColumn(6.6m),
            StatementColumnKind.HiddenMarker => new OdsColumn(0.5m, Hidden: true),
            _ => new OdsColumn(2.4m)
        }).ToImmutableArray();
        return grid.ToWorksheet(sheet.Id, sheet.Name, odsColumns, frozenColumns: 3, frozenRows: 4, protectedSheet: true);
    }

    private static void RenderCashFlowRow(
        OdsGrid grid,
        StatementSheet sheet,
        StatementRow row,
        int rowNumber,
        StatementLayout layout)
    {
        switch (row.Kind)
        {
            case StatementRowKind.Detail:
            case StatementRowKind.BankAccount:
            case StatementRowKind.BalanceSheetAsset:
            case StatementRowKind.SummaryReference:
                grid.Set(rowNumber, 1, Text(row.BusinessCode ?? string.Empty, row.PresentationRole));
                grid.Set(rowNumber, 2, Text(row.Label, row.PresentationRole));
                break;
            case StatementRowKind.CategoryTotal:
                grid.Set(rowNumber, 1, Text(row.Label, row.PresentationRole));
                grid.Set(rowNumber, 3, Text(row.BusinessCode ?? string.Empty, StatementPresentationRole.HiddenMarker));
                grid.ApplyRoleToRow(rowNumber, row.PresentationRole);
                grid.ApplyFormatToRow(rowNumber, bold: true,
                    topBorder: OdsBorderWeight.Thin, bottomBorder: OdsBorderWeight.Thin);
                break;
            case StatementRowKind.CalculatedTotal:
                grid.Set(rowNumber, 1, Text(row.BusinessCode ?? string.Empty));
                grid.Set(rowNumber, 2, Text(row.Label));
                grid.Set(rowNumber, 3, Text(row.BusinessCode ?? string.Empty, StatementPresentationRole.HiddenMarker));
                grid.ApplyFormatToRow(rowNumber, locked: true);
                break;
            case StatementRowKind.Expression:
                grid.Set(rowNumber, 1, Text(row.Label, row.PresentationRole));
                grid.Set(rowNumber, 3, Text(row.BusinessCode ?? string.Empty, StatementPresentationRole.HiddenMarker));
                grid.ApplyFormatToRow(rowNumber, locked: true);
                break;
            case StatementRowKind.PeriodTotal:
            case StatementRowKind.BankTotal:
            case StatementRowKind.VatMetric:
            case StatementRowKind.Capital:
                grid.Set(rowNumber, 1, Text(
                    row.Kind is StatementRowKind.BankTotal or StatementRowKind.Capital
                        ? row.Label.ToUpperInvariant()
                        : row.Label,
                    row.PresentationRole));
                if (row.Kind == StatementRowKind.PeriodTotal)
                    grid.ApplyFormatToRow(rowNumber, bold: true, locked: true,
                        topBorder: OdsBorderWeight.Thin, bottomBorder: OdsBorderWeight.Thick);
                else if (row.Kind == StatementRowKind.BankTotal)
                    grid.ApplyFormatToRow(rowNumber, bold: true,
                        topBorder: OdsBorderWeight.Thin, bottomBorder: OdsBorderWeight.Thin);
                else if (row.Kind == StatementRowKind.Capital)
                    grid.ApplyFormatToRow(rowNumber, bold: true, locked: true,
                        topBorder: OdsBorderWeight.Thin, bottomBorder: OdsBorderWeight.Thick);
                break;
        }

        if (row.Kind == StatementRowKind.SummaryReference)
            grid.ApplyFormatToRow(rowNumber, locked: true);

        var firstPeriod = sheet.Columns.Where(column => column.PeriodId is not null)
            .MinBy(column => column.PresentationOrder)!;
        foreach (var pair in row.Cells)
        {
            var semantic = new StatementCellAddress(sheet.Id, row.Id, pair.Key);
            var physical = layout.Cells[semantic];
            var source = pair.Value;
            if (row.Kind == StatementRowKind.Expression &&
                source.Value.State == StatementValueState.Unsupported && source.Formula is null)
            {
                grid.Set(physical.RowNumber, physical.ColumnNumber,
                    pair.Key == firstPeriod.Id
                        ? new OdsCell(StatementValue.FromText(source.Value.Reason ?? "Expression error"),
                            FontColor: "#ff0000",
                            HorizontalAlignment: OdsHorizontalAlignment.Left,
                            SemanticAddress: semantic)
                        : new OdsCell(StatementValue.Missing(source.Value.Kind), SemanticAddress: semantic));
                continue;
            }

            var renderedSource = row.Kind == StatementRowKind.Capital
                ? source with
                {
                    Formula = new StatementFormulaAggregate(
                        StatementAggregateOperator.Sum,
                        sheet.Sections.Single(section => section.Kind == StatementSectionKind.BalanceSheet)
                            .Rows.Where(asset => asset.Kind == StatementRowKind.BalanceSheetAsset)
                            .Select(asset => (StatementFormula)new StatementFormulaReference(
                                new StatementCellAddress(sheet.Id, asset.Id, pair.Key)))
                            .ToImmutableArray())
                }
                : source;
            grid.Set(physical.RowNumber, physical.ColumnNumber,
                FromModel(renderedSource, semantic,
                    row.Kind == StatementRowKind.CalculatedTotal ? StatementPresentationRole.Default : null));
        }
    }

    private static void FinishSection(
        OdsGrid grid,
        StatementSection section,
        int firstRow,
        int lastRow)
    {
        if (section.Rows.IsEmpty)
            return;
        if (section.Kind == StatementSectionKind.FlowSummary)
        {
            var ordered = section.Rows.OrderBy(row => row.PresentationOrder).ToArray();
            var lastSummary = Array.FindLastIndex(ordered, row => row.Kind == StatementRowKind.SummaryReference);
            if (lastSummary >= 0)
                grid.ApplyFormatToRow(firstRow + lastSummary, bottomBorder: OdsBorderWeight.Thick);
        }
        if (section.Kind is StatementSectionKind.FlowSummary or StatementSectionKind.CategoryTotals or
            StatementSectionKind.Analysis or StatementSectionKind.VatRecurrence or StatementSectionKind.VatPeriods)
            grid.ApplyFormatToRow(lastRow, bottomBorder: OdsBorderWeight.Thick);
        if (section.Kind is StatementSectionKind.VatRecurrence or StatementSectionKind.VatPeriods)
        {
            grid.ApplyFormatToRow(firstRow, topBorder: OdsBorderWeight.Thin);
            grid.ApplyFormatToRow(lastRow, bold: true);
        }
    }

    private static OdsWorksheet RenderSettings(CashStatementDocument document)
    {
        var rows = new List<OdsRow>
        {
            Row(Text("Document Type", StatementPresentationRole.ColumnHeader), Text("cashflow")),
            Row(Text("Format", StatementPresentationRole.ColumnHeader), Text("ods")),
            Row(Text("User", StatementPresentationRole.ColumnHeader), Text(document.Metadata.RequestedBy)),
            Row(Blank(), Blank()),
            Row(Blank(), Blank()),
            Row(Text("Parameters", StatementPresentationRole.SectionHeader), Blank())
        };
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["commandTimeout"] = document.Request.CommandTimeoutSeconds.ToString(CultureInfo.InvariantCulture),
            ["includeActivePeriods"] = Boolean(document.Request.IncludeActivePeriods),
            ["includeBalanceSheet"] = Boolean(document.Request.IncludeBalanceSheet),
            ["includeBankBalances"] = Boolean(document.Request.IncludeBankBalances),
            ["includeOrderBook"] = Boolean(document.Request.IncludeOrderBook),
            ["includeReconciliation"] = Boolean(document.Request.IncludeReconciliation),
            ["includeTaxAccruals"] = Boolean(document.Request.IncludeTaxAccruals),
            ["includeVatDetails"] = Boolean(document.Request.IncludeVatDetails)
        };
        rows.AddRange(values.OrderBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase)
            .Select(pair => Row(Text(pair.Key), Text(pair.Value))));
        return new OdsWorksheet(
            new StatementSheetId("settings"),
            "Settings",
            ImmutableArray.Create(new OdsColumn(6.4m), new OdsColumn(21m)),
            rows.ToImmutableArray(),
            FrozenRows: 1);
    }

    private static OdsWorksheet RenderReconciliation(
        StatementSheet sheet,
        StatementLayout layout,
        CancellationToken cancellationToken)
    {
        var rows = sheet.Sections.SelectMany(section => section.Rows).ToArray();
        var dataRows = rows.Count(row => row.Kind == StatementRowKind.Reconciliation);
        var grid = new OdsGrid(Math.Max(5, 4 + dataRows), 15);
        var summary = rows.Single(row => row.Kind == StatementRowKind.Metadata);
        grid.Set(1, 1, Text(summary.Label, StatementPresentationRole.ReconciliationHeader));
        grid.Set(2, 1, Text(sheet.Columns.Single(column => column.Id.Value == "tolerance").Label,
            StatementPresentationRole.ColumnHeader));
        grid.Set(2, 4, Text(sheet.Columns.Single(column => column.Id.Value == "status").Label,
            StatementPresentationRole.ColumnHeader));

        foreach (var pair in ReconciliationColumns)
            grid.Set(4, pair.Value, Text(sheet.Columns.Single(column => column.Id.Value == pair.Key).Label,
                StatementPresentationRole.ColumnHeader));
        grid.ApplyRoleToRow(4, StatementPresentationRole.ColumnHeader);
        grid.ApplyFormatToRow(4, bold: true, locked: true, bottomBorder: OdsBorderWeight.Thick);

        foreach (var row in rows)
        {
            cancellationToken.ThrowIfCancellationRequested();
            foreach (var pair in row.Cells)
            {
                var semantic = new StatementCellAddress(sheet.Id, row.Id, pair.Key);
                var physical = layout.Cells[semantic];
                var role = pair.Value.PresentationRole == StatementPresentationRole.ReconciliationHeader
                    ? StatementPresentationRole.Default
                    : pair.Value.PresentationRole;
                var numberFormat = pair.Value.NumericKind == StatementNumericKind.Money ? "Cash2" : pair.Value.NumberFormat;
                grid.Set(physical.RowNumber, physical.ColumnNumber,
                    FromModel(pair.Value, semantic, role, numberFormat));
            }
        }
        grid.ApplyFormatToCell(2, 2, locked: true, numberFormat: "Cash2");
        grid.ApplyFormatToCell(2, 5, bold: true);
        for (var column = 3; column <= 14; column++)
            grid.ApplyFormatToColumn(column, numberFormat: "Cash2");

        var columns = Enumerable.Range(1, 15)
            .Select(column => new OdsColumn(column switch { 1 => 2.4m, 2 => 6.6m, 12 => 0.8m, _ => 3m }))
            .ToImmutableArray();
        return grid.ToWorksheet(sheet.Id, sheet.Name, columns, frozenColumns: 2, frozenRows: 4);
    }

    private static OdsCell FromModel(
        StatementCell source,
        StatementCellAddress semantic,
        StatementPresentationRole? presentationRole = null,
        string? numberFormat = null)
    {
        var formula = source.Formula is not null && ContainsEmptyAggregate(source.Formula)
            ? null
            : source.Formula;
        return new OdsCell(
            source.Value,
            source.NumericKind,
            presentationRole ?? source.PresentationRole,
            formula,
            semantic,
            NumberFormat: numberFormat ?? source.NumberFormat);
    }

    private static bool ContainsEmptyAggregate(StatementFormula formula) => formula switch
    {
        StatementFormulaAggregate aggregate => aggregate.Operands.IsEmpty || aggregate.Operands.Any(ContainsEmptyAggregate),
        StatementFormulaUnary unary => ContainsEmptyAggregate(unary.Operand),
        StatementFormulaBinary binary => ContainsEmptyAggregate(binary.Left) || ContainsEmptyAggregate(binary.Right),
        StatementFormulaConditional conditional => ContainsEmptyAggregate(conditional.Condition) ||
                                                   ContainsEmptyAggregate(conditional.WhenTrue) ||
                                                   ContainsEmptyAggregate(conditional.WhenFalse),
        _ => false
    };

    private static OdsCell Text(string value, StatementPresentationRole role = StatementPresentationRole.Default) =>
        new(StatementValue.FromText(value), PresentationRole: role);
    private static OdsCell Blank() => new(StatementValue.Missing(StatementScalarKind.Text));
    private static OdsRow Row(params OdsCell[] cells) => new(cells.ToImmutableArray());
    private static string Boolean(bool value) => value ? "true" : "false";

    private sealed class OdsGrid(int rowCount, int columnCount)
    {
        private readonly OdsCell[,] _cells = Create(rowCount, columnCount);

        public void Set(int rowNumber, int columnNumber, OdsCell cell)
        {
            if (rowNumber < 1 || rowNumber > rowCount || columnNumber < 1 || columnNumber > columnCount)
                throw new InvalidOperationException($"ODS cell {rowNumber},{columnNumber} is outside the allocated grid.");
            _cells[rowNumber - 1, columnNumber - 1] = cell;
        }

        public OdsWorksheet ToWorksheet(
            StatementSheetId id,
            string name,
            ImmutableArray<OdsColumn> columns,
            int frozenColumns = 0,
            int frozenRows = 0,
            bool protectedSheet = false)
        {
            var rows = ImmutableArray.CreateBuilder<OdsRow>(rowCount);
            for (var row = 0; row < rowCount; row++)
            {
                var cells = ImmutableArray.CreateBuilder<OdsCell>(columnCount);
                for (var column = 0; column < columnCount; column++)
                    cells.Add(_cells[row, column]);
                rows.Add(new OdsRow(cells.MoveToImmutable()));
            }
            return new OdsWorksheet(id, name, columns, rows.MoveToImmutable(), frozenColumns, frozenRows,
                Protected: protectedSheet);
        }

        public void ApplyRoleToRow(int rowNumber, StatementPresentationRole role)
        {
            if (rowNumber < 1 || rowNumber > rowCount)
                throw new InvalidOperationException($"ODS row {rowNumber} is outside the allocated grid.");
            for (var column = 0; column < columnCount; column++)
                _cells[rowNumber - 1, column] = _cells[rowNumber - 1, column] with { PresentationRole = role };
        }

        public void ApplyFormatToRow(
            int rowNumber,
            bool? bold = null,
            bool? locked = null,
            OdsBorderWeight? topBorder = null,
            OdsBorderWeight? bottomBorder = null,
            OdsBorderWeight? leftBorder = null,
            OdsBorderWeight? rightBorder = null,
            string? numberFormat = null)
        {
            if (rowNumber < 1 || rowNumber > rowCount)
                throw new InvalidOperationException($"ODS row {rowNumber} is outside the allocated grid.");
            for (var column = 1; column <= columnCount; column++)
                ApplyFormatToCell(rowNumber, column, bold, locked, topBorder, bottomBorder, leftBorder,
                    rightBorder, numberFormat);
        }

        public void ApplyFormatToColumn(
            int columnNumber,
            bool? bold = null,
            bool? locked = null,
            OdsBorderWeight? topBorder = null,
            OdsBorderWeight? bottomBorder = null,
            OdsBorderWeight? leftBorder = null,
            OdsBorderWeight? rightBorder = null,
            string? numberFormat = null)
        {
            if (columnNumber < 1 || columnNumber > columnCount)
                throw new InvalidOperationException($"ODS column {columnNumber} is outside the allocated grid.");
            for (var row = 1; row <= rowCount; row++)
                ApplyFormatToCell(row, columnNumber, bold, locked, topBorder, bottomBorder, leftBorder,
                    rightBorder, numberFormat);
        }

        public void ApplyFormatToCell(
            int rowNumber,
            int columnNumber,
            bool? bold = null,
            bool? locked = null,
            OdsBorderWeight? topBorder = null,
            OdsBorderWeight? bottomBorder = null,
            OdsBorderWeight? leftBorder = null,
            OdsBorderWeight? rightBorder = null,
            string? numberFormat = null)
        {
            if (rowNumber < 1 || rowNumber > rowCount || columnNumber < 1 || columnNumber > columnCount)
                throw new InvalidOperationException($"ODS cell {rowNumber},{columnNumber} is outside the allocated grid.");
            var cell = _cells[rowNumber - 1, columnNumber - 1];
            _cells[rowNumber - 1, columnNumber - 1] = cell with
            {
                Bold = bold ?? cell.Bold,
                Locked = locked ?? cell.Locked,
                TopBorder = topBorder ?? cell.TopBorder,
                BottomBorder = bottomBorder ?? cell.BottomBorder,
                LeftBorder = leftBorder ?? cell.LeftBorder,
                RightBorder = rightBorder ?? cell.RightBorder,
                NumberFormat = numberFormat ?? cell.NumberFormat
            };
        }

        public void ApplyColumnContext(int columnNumber, bool activePeriod, bool annualPeriod)
        {
            if (columnNumber < 1 || columnNumber > columnCount)
                throw new InvalidOperationException($"ODS column {columnNumber} is outside the allocated grid.");
            for (var row = 1; row <= rowCount; row++)
            {
                var cell = _cells[row - 1, columnNumber - 1];
                _cells[row - 1, columnNumber - 1] = cell with
                {
                    ActivePeriod = activePeriod,
                    Bold = annualPeriod || cell.Bold,
                    Locked = annualPeriod || cell.Locked,
                    LeftBorder = annualPeriod ? OdsBorderWeight.Thin : cell.LeftBorder,
                    RightBorder = annualPeriod ? OdsBorderWeight.Thick : cell.RightBorder
                };
            }
        }

        private static OdsCell[,] Create(int rows, int columns)
        {
            var cells = new OdsCell[rows, columns];
            for (var row = 0; row < rows; row++)
                for (var column = 0; column < columns; column++)
                    cells[row, column] = Blank();
            return cells;
        }
    }
}
