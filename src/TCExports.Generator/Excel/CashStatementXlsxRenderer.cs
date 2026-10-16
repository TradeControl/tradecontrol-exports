using System.Globalization;
using ClosedXML.Excel;
using TCExports.Generator.CashStatement.Model;

namespace TCExports.Generator.Excel;

/// <summary>
/// Renders the immutable Cash Statement document to XLSX without repository access
/// or financial calculation.
/// </summary>
public sealed class CashStatementXlsxRenderer
{
    private const string MoneyFormat = "#,##0;[Red](#,##0);_-";
    private const string ReconciliationMoneyFormat = "#,##0.00;[Red](#,##0.00);_-";

    private readonly CashStatementLimits _limits;
    private readonly ICashStatementLayoutAllocator _allocator;
    private readonly ICashStatementFormulaTranslator _formulaTranslator;

    public CashStatementXlsxRenderer(
        CashStatementLimits? limits = null,
        ICashStatementLayoutAllocator? allocator = null,
        ICashStatementFormulaTranslator? formulaTranslator = null)
    {
        _limits = limits ?? CashStatementLimits.Default;
        _allocator = allocator ?? new CashStatementXlsxLayoutAllocator();
        _formulaTranslator = formulaTranslator ?? new CashStatementXlsxFormulaTranslator();
    }

    public byte[] Render(CashStatementDocument document, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(document);
        ct.ThrowIfCancellationRequested();
        var layout = _allocator.Allocate(document, _limits);

        using var workbook = new XLWorkbook();
        var cashFlow = document.Sheets.Single(sheet => sheet.Kind == StatementSheetKind.CashFlow);
        RenderCashFlow(workbook, cashFlow, document, layout, ct);
        RenderSettings(workbook, document);

        var reconciliation = document.Sheets.SingleOrDefault(
            sheet => sheet.Kind == StatementSheetKind.Reconciliation);
        if (reconciliation is not null)
            RenderReconciliation(workbook, reconciliation, layout, ct);

        using var output = new MemoryStream();
        workbook.SaveAs(output);
        if (output.Length > _limits.MaxOutputBytes)
        {
            throw new InvalidOperationException(
                $"XLSX output is {output.Length} bytes; limit is {_limits.MaxOutputBytes}.");
        }
        return output.ToArray();
    }

    private void RenderCashFlow(
        XLWorkbook workbook,
        StatementSheet sheet,
        CashStatementDocument document,
        StatementLayout layout,
        CancellationToken ct)
    {
        var worksheet = workbook.Worksheets.Add(sheet.Name);
        var columns = sheet.Columns.OrderBy(column => column.PresentationOrder).ToArray();

        worksheet.Cell(1, 1).Value = document.Metadata.Title;
        worksheet.Cell(1, 1).Style.Font.Bold = true;
        worksheet.Cell(1, 1).Style.Font.FontSize = 12;
        worksheet.Cell(2, 1).Value = document.Metadata.CompanyName;
        worksheet.Cell(2, 1).Style.Font.Bold = true;
        worksheet.Cell(2, 1).Style.Font.FontSize = 10;
        worksheet.Cell(3, 1).Value = document.Metadata.DateLabel;
        worksheet.Cell(3, 2).Value = document.Metadata.GeneratedAt.ToString("dd MMM HH:mm:ss", CultureInfo.InvariantCulture);
        worksheet.Row(3).Style.Font.Bold = true;
        worksheet.Row(3).Style.Font.FontSize = 10;
        worksheet.Range("B2:B3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

        worksheet.Cell(4, 1).Value = columns.Single(column => column.Kind == StatementColumnKind.Identifier).Label;
        worksheet.Cell(4, 2).Value = columns.Single(column => column.Kind == StatementColumnKind.Label).Label;
        worksheet.Row(4).Style.Font.Bold = true;
        worksheet.Row(4).Style.Font.FontSize = 8;
        worksheet.Row(4).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
        worksheet.Row(4).Style.Border.TopBorder = XLBorderStyleValues.Thin;
        worksheet.Column(2).Style.Border.RightBorder = XLBorderStyleValues.Thick;

        // Match the legacy writer's operation order. Applying column protection after
        // the data rows exist causes ClosedXML to serialize otherwise missing cells at
        // every row/column intersection.
        worksheet.Protect();
        worksheet.Style.Protection.Locked = false;
        worksheet.Columns(1, 3).Style.Protection.Locked = true;
        worksheet.Rows(1, 4).Style.Protection.Locked = true;

        foreach (var column in columns.Where(column => column.PeriodId is not null))
        {
            var columnNumber = column.PresentationOrder + 1;
            var period = document.PeriodGrid.Periods.Single(item => item.Id == column.PeriodId);
            worksheet.Cell(4, columnNumber).Value = column.Label;
            worksheet.Column(columnNumber).Style.NumberFormat.Format = MoneyFormat;
            if (!period.Id.IsAnnual && document.PeriodGrid.Periods
                    .Where(item => item.Id.YearNumber == period.Id.YearNumber && !item.Id.IsAnnual)
                    .OrderBy(item => item.PresentationOrder)
                    .First().Id == period.Id)
            {
                worksheet.Cell(3, columnNumber).Value = $"{period.YearLabel} ({period.CashStatus})";
            }
            if (period.Id.IsAnnual)
            {
                worksheet.Cell(3, columnNumber).Value = period.YearLabel;
                worksheet.Column(columnNumber).Style.Border.RightBorder = XLBorderStyleValues.Thick;
                worksheet.Column(columnNumber).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                worksheet.Column(columnNumber).Style.Protection.Locked = true;
                worksheet.Column(columnNumber).Style.Font.Bold = true;
            }
            if (period.IsActive)
                worksheet.Column(columnNumber).Style.Fill.BackgroundColor = XLColor.Yellow;
        }

        var currentRow = 4;
        foreach (var section in sheet.Sections.OrderBy(section => section.PresentationOrder))
        {
            ct.ThrowIfCancellationRequested();
            if (section.Kind == StatementSectionKind.FlowCategories)
            {
                foreach (var row in section.Rows.OrderBy(row => row.PresentationOrder))
                {
                    if (row.Kind == StatementRowKind.Header)
                    {
                        currentRow += 2;
                        worksheet.Cell(currentRow, 1).Value = row.Label;
                        worksheet.Row(currentRow).Style.Font.Bold = true;
                        worksheet.Row(currentRow).Style.Border.TopBorder = XLBorderStyleValues.Medium;
                        continue;
                    }
                    currentRow++;
                    RenderCashFlowRow(worksheet, sheet, row, currentRow, layout);
                }
                continue;
            }

            currentRow += 2;
            RenderSectionHeader(worksheet, section, currentRow);
            foreach (var row in section.Rows.OrderBy(row => row.PresentationOrder))
            {
                currentRow++;
                RenderCashFlowRow(worksheet, sheet, row, currentRow, layout);
            }
            FinishSection(worksheet, section, currentRow);
        }

        worksheet.Column(3).Hide();
        worksheet.Column(1).Width = 9;
        worksheet.Column(2).Width = 25;
        foreach (var column in columns.Where(column => column.PeriodId is not null))
            worksheet.Column(column.PresentationOrder + 1).Width = 9;
        worksheet.SheetView.FreezeRows(4);
        worksheet.SheetView.FreezeColumns(3);
    }

    private void RenderCashFlowRow(
        IXLWorksheet worksheet,
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
                worksheet.Cell(rowNumber, 1).Value = row.BusinessCode ?? string.Empty;
                worksheet.Cell(rowNumber, 2).Value = row.Label;
                break;
            case StatementRowKind.CategoryTotal:
                worksheet.Cell(rowNumber, 1).Value = row.Label;
                worksheet.Cell(rowNumber, 3).Value = row.BusinessCode ?? string.Empty;
                StyleTotalRow(worksheet.Row(rowNumber), row.PresentationRole);
                worksheet.Cell(rowNumber, 3).Style.Font.FontColor = worksheet.Row(rowNumber).Style.Fill.BackgroundColor;
                break;
            case StatementRowKind.SummaryReference:
                worksheet.Cell(rowNumber, 1).Value = row.BusinessCode ?? string.Empty;
                worksheet.Cell(rowNumber, 2).Value = row.Label;
                worksheet.Row(rowNumber).Style.Protection.Locked = true;
                break;
            case StatementRowKind.PeriodTotal:
                worksheet.Cell(rowNumber, 1).Value = row.Label;
                worksheet.Row(rowNumber).Style.Font.Bold = true;
                worksheet.Row(rowNumber).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Row(rowNumber).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
                worksheet.Row(rowNumber).Style.Protection.Locked = true;
                break;
            case StatementRowKind.CalculatedTotal:
                worksheet.Cell(rowNumber, 1).FormulaA1 = $"=\"{EscapeFormulaText(row.BusinessCode ?? string.Empty)}\"";
                worksheet.Cell(rowNumber, 2).Value = row.Label;
                worksheet.Cell(rowNumber, 3).Value = row.BusinessCode ?? string.Empty;
                worksheet.Row(rowNumber).Style.Protection.Locked = true;
                break;
            case StatementRowKind.Expression:
                worksheet.Cell(rowNumber, 1).Value = row.Label;
                worksheet.Cell(rowNumber, 3).Value = row.BusinessCode ?? string.Empty;
                worksheet.Cell(rowNumber, 3).Style.Font.FontColor = XLColor.LimeGreen;
                worksheet.Row(rowNumber).Style.Protection.Locked = true;
                if (row.Cells.Values.Select(cell => cell.NumberFormat).FirstOrDefault(format => !string.IsNullOrWhiteSpace(format)) is { } format)
                    worksheet.Row(rowNumber).Style.NumberFormat.Format = format;
                var firstPeriodColumn = sheet.Columns.Where(column => column.PeriodId is not null).Min(column => column.PresentationOrder + 1);
                var lastPeriodColumn = sheet.Columns.Where(column => column.PeriodId is not null).Max(column => column.PresentationOrder + 1);
                worksheet.Range(rowNumber, firstPeriodColumn, rowNumber, lastPeriodColumn)
                    .AddConditionalFormat().WhenLessThan(0).Font.SetFontColor(XLColor.Red);
                break;
            case StatementRowKind.BankTotal:
                worksheet.Cell(rowNumber, 1).Value = row.Label.ToUpperInvariant();
                worksheet.Row(rowNumber).Style.Font.Bold = true;
                worksheet.Row(rowNumber).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Row(rowNumber).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                break;
            case StatementRowKind.VatMetric:
                worksheet.Cell(rowNumber, 1).Value = row.Label;
                break;
            case StatementRowKind.Capital:
                worksheet.Cell(rowNumber, 1).Value = row.Label.ToUpperInvariant();
                worksheet.Row(rowNumber).Style.Font.Bold = true;
                worksheet.Row(rowNumber).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Row(rowNumber).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
                worksheet.Row(rowNumber).Style.Protection.Locked = true;
                break;
        }

        WriteCells(worksheet, sheet, row, layout);
        if (row.Kind == StatementRowKind.Capital)
            WriteCapitalFormulas(worksheet, sheet, row, layout);
    }

    private void WriteCells(
        IXLWorksheet worksheet,
        StatementSheet sheet,
        StatementRow row,
        StatementLayout layout)
    {
        foreach (var pair in row.Cells)
        {
            var semantic = new StatementCellAddress(sheet.Id, row.Id, pair.Key);
            var cell = pair.Value;
            if (cell.Value.State == StatementValueState.Missing && cell.Formula is null)
                continue;

            if (row.Kind == StatementRowKind.Expression &&
                cell.Value.State == StatementValueState.Unsupported &&
                cell.Formula is null)
            {
                var firstPeriodColumn = sheet.Columns
                    .Where(column => column.PeriodId is not null)
                    .MinBy(column => column.PresentationOrder)!;
                if (pair.Key != firstPeriodColumn.Id)
                    continue;

                var errorAddress = layout.Cells[semantic];
                var errorCell = worksheet.Cell(errorAddress.RowNumber, errorAddress.ColumnNumber);
                errorCell.Value = cell.Value.Reason ?? "Expression error";
                errorCell.Style.Font.FontColor = XLColor.Red;
                errorCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                continue;
            }

            var physical = layout.Cells[semantic];
            var target = worksheet.Cell(physical.RowNumber, physical.ColumnNumber);

            if (cell.Value.State == StatementValueState.Present)
                WriteValue(target, cell.Value);

            var formula = SimplifyCalculatedTotal(cell.Formula, row.Kind);
            if (formula is not null && ContainsEmptyAggregate(formula))
                continue;
            if (formula is not null &&
                (cell.Value.State != StatementValueState.Unsupported || formula is StatementFormulaTerminal))
            {
                target.FormulaA1 = _formulaTranslator.Translate(formula, layout);
            }
            if (!string.IsNullOrWhiteSpace(cell.NumberFormat))
                target.Style.NumberFormat.Format = cell.NumberFormat;
            else if (cell.NumericKind == StatementNumericKind.Percentage)
                target.Style.NumberFormat.Format = "0.00%";
        }
    }

    private static StatementFormula? SimplifyCalculatedTotal(StatementFormula? formula, StatementRowKind rowKind) =>
        rowKind == StatementRowKind.CalculatedTotal &&
        formula is StatementFormulaAggregate
        {
            Operator: StatementAggregateOperator.Sum,
            Operands.Length: 1
        } aggregate
            ? aggregate.Operands[0]
            : formula;

    private static void WriteCapitalFormulas(
        IXLWorksheet worksheet,
        StatementSheet sheet,
        StatementRow capital,
        StatementLayout layout)
    {
        var balanceSection = sheet.Sections.Single(section => section.Kind == StatementSectionKind.BalanceSheet);
        var assets = balanceSection.Rows.Where(row => row.Kind == StatementRowKind.BalanceSheetAsset).ToArray();
        if (assets.Length == 0)
            return;

        foreach (var pair in capital.Cells)
        {
            var capitalAddress = layout.Cells[new StatementCellAddress(sheet.Id, capital.Id, pair.Key)];
            var firstAddress = layout.Cells[new StatementCellAddress(sheet.Id, assets[0].Id, pair.Key)];
            var lastAddress = layout.Cells[new StatementCellAddress(sheet.Id, assets[^1].Id, pair.Key)];
            var column = ColumnName(capitalAddress.ColumnNumber);
            worksheet.Cell(capitalAddress.RowNumber, capitalAddress.ColumnNumber).FormulaA1 =
                $"=SUM({column}{firstAddress.RowNumber}:{column}{lastAddress.RowNumber})";
        }
    }

    private static void RenderSectionHeader(IXLWorksheet worksheet, StatementSection section, int rowNumber)
    {
        worksheet.Cell(rowNumber, 1).Value = section.Title.ToUpperInvariant();
        var row = worksheet.Row(rowNumber);
        row.Style.Font.Bold = true;
        row.Style.Border.TopBorder = XLBorderStyleValues.Thin;
        row.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        if (section.Kind != StatementSectionKind.CategoryTotals)
            row.Style.Protection.Locked = true;
        row.Style.Fill.BackgroundColor = section.Kind switch
        {
            StatementSectionKind.FlowSummary => XLColor.LightGreen,
            StatementSectionKind.CategoryTotals => XLColor.GreenYellow,
            StatementSectionKind.Analysis => XLColor.LimeGreen,
            StatementSectionKind.BankBalances => XLColor.LightSteelBlue,
            StatementSectionKind.VatRecurrence or StatementSectionKind.VatPeriods => XLColor.Orchid,
            StatementSectionKind.BalanceSheet => XLColor.Red,
            _ => XLColor.NoColor
        };
        if (section.Kind == StatementSectionKind.BalanceSheet)
            row.Style.Font.FontColor = XLColor.White;
    }

    private static void FinishSection(IXLWorksheet worksheet, StatementSection section, int lastRow)
    {
        if (section.Rows.IsEmpty)
            return;
        if (section.Kind == StatementSectionKind.FlowSummary)
        {
            var orderedRows = section.Rows.OrderBy(row => row.PresentationOrder).ToArray();
            var lastSummaryIndex = Array.FindLastIndex(
                orderedRows,
                row => row.Kind == StatementRowKind.SummaryReference);
            if (lastSummaryIndex >= 0)
            {
                var firstRow = lastRow - orderedRows.Length + 1;
                worksheet.Row(firstRow + lastSummaryIndex).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
            }
        }
        if (section.Kind is StatementSectionKind.FlowSummary or StatementSectionKind.CategoryTotals or
            StatementSectionKind.Analysis or StatementSectionKind.VatRecurrence or StatementSectionKind.VatPeriods)
        {
            worksheet.Row(lastRow).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
        }
        if (section.Kind is StatementSectionKind.VatRecurrence or StatementSectionKind.VatPeriods)
            worksheet.Row(lastRow).Style.Font.Bold = true;
        if (section.Kind is StatementSectionKind.VatRecurrence or StatementSectionKind.VatPeriods && section.Rows.Length > 0)
        {
            var firstRow = lastRow - section.Rows.Length + 1;
            worksheet.Row(firstRow).Style.Border.TopBorder = XLBorderStyleValues.Thin;
        }
    }

    private static void StyleTotalRow(IXLRow row, StatementPresentationRole role)
    {
        row.Style.Font.Bold = true;
        row.Style.Border.TopBorder = XLBorderStyleValues.Thin;
        row.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        row.Style.Fill.BackgroundColor = role switch
        {
            StatementPresentationRole.ExpenseTotal => XLColor.LightSalmon,
            StatementPresentationRole.IncomeTotal => XLColor.CornflowerBlue,
            _ => XLColor.LightGray
        };
    }

    private static void RenderSettings(XLWorkbook workbook, CashStatementDocument document)
    {
        var worksheet = workbook.Worksheets.Add("Settings");
        worksheet.Cell(1, 1).Value = "Document Type";
        worksheet.Cell(1, 2).Value = "cashflow";
        worksheet.Cell(2, 1).Value = "Format";
        worksheet.Cell(2, 2).Value = "excel";
        worksheet.Cell(3, 1).Value = "User";
        worksheet.Cell(3, 2).Value = document.Metadata.RequestedBy;
        // Row 4 deliberately remains blank: SQL credentials are never rendered.
        worksheet.Cell(6, 1).Value = "Parameters";
        worksheet.Row(6).Style.Font.Bold = true;

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
        var row = 7;
        foreach (var pair in values.OrderBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase))
        {
            worksheet.Cell(row, 1).Value = pair.Key;
            worksheet.Cell(row, 2).Value = pair.Value;
            row++;
        }

        worksheet.Column(1).Width = 24;
        worksheet.Column(2).Width = 80;
        worksheet.Range(1, 1, 1, 2).Style.Font.Bold = true;
        worksheet.Range(1, 1, row - 1, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        worksheet.Range(1, 1, row - 1, 2).Style.Border.InsideBorder = XLBorderStyleValues.Hair;
        worksheet.SheetView.FreezeRows(1);

        static string Boolean(bool value) => value ? "true" : "false";
    }

    private void RenderReconciliation(
        XLWorkbook workbook,
        StatementSheet sheet,
        StatementLayout layout,
        CancellationToken ct)
    {
        var worksheet = workbook.Worksheets.Add(sheet.Name);
        var rows = sheet.Sections.Single().Rows.OrderBy(row => row.PresentationOrder).ToArray();
        var summary = rows.Single(row => row.Kind == StatementRowKind.Metadata);
        worksheet.Cell(1, 1).Value = summary.Label;
        worksheet.Row(1).Style.Font.Bold = true;
        worksheet.Cell(2, 1).Value = sheet.Columns.Single(column => column.Id.Value == "tolerance").Label;
        worksheet.Cell(2, 4).Value = sheet.Columns.Single(column => column.Id.Value == "status").Label;
        WriteCells(worksheet, sheet, summary, layout);
        worksheet.Cell(2, 2).Style.NumberFormat.Format = ReconciliationMoneyFormat;
        worksheet.Cell(2, 5).Style.Font.Bold = true;

        var headerColumns = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["year"] = 1, ["description"] = 2, ["opening-capital"] = 3, ["closing-capital"] = 4,
            ["profit"] = 5, ["business-tax"] = 6, ["profit-after-tax"] = 7,
            ["capital-movement"] = 8, ["opening-subject-position"] = 9,
            ["opening-account-position"] = 10, ["capital-delta"] = 11,
            ["bridge-total"] = 13, ["difference"] = 14, ["status"] = 15
        };
        foreach (var pair in headerColumns)
            worksheet.Cell(4, pair.Value).Value = sheet.Columns.Single(column => column.Id.Value == pair.Key).Label;
        worksheet.Row(4).Style.Font.Bold = true;
        worksheet.Row(4).Style.Border.BottomBorder = XLBorderStyleValues.Thick;

        var dataRows = rows.Where(row => row.Kind == StatementRowKind.Reconciliation).ToArray();
        foreach (var row in dataRows)
        {
            ct.ThrowIfCancellationRequested();
            WriteCells(worksheet, sheet, row, layout);
            var rowNumber = layout.Cells[new StatementCellAddress(sheet.Id, row.Id, new StatementColumnId("status"))].RowNumber;
            worksheet.Cell(rowNumber, 15).FormulaA1 =
                $"=IF(ABS(N{rowNumber})<= $B$2, \"PASS\", IF(ABS(N{rowNumber})<= ($B$2*10), \"WARN\", \"FAIL\"))";
        }

        var firstDataRow = 5;
        var lastDataRow = Math.Max(firstDataRow, firstDataRow + dataRows.Length - 1);
        ApplyStatusFormats(worksheet, firstDataRow, lastDataRow);

        worksheet.Row(2).Style.Protection.Locked = true;
        worksheet.Row(4).Style.Protection.Locked = true;
        worksheet.Columns(3, 14).Style.NumberFormat.Format = ReconciliationMoneyFormat;
        worksheet.Column(1).Width = 9;
        worksheet.Column(2).Width = 25;
        worksheet.Columns(3, 15).AdjustToContents();
        worksheet.SheetView.FreezeRows(4);
    }

    private static void ApplyStatusFormats(IXLWorksheet worksheet, int firstDataRow, int lastDataRow)
    {
        var ranges = new[] { worksheet.Range($"O{firstDataRow}:O{lastDataRow}"), worksheet.Range("E2:E2") };
        foreach (var range in ranges)
        {
            range.AddConditionalFormat().WhenEquals("FAIL").Font.SetFontColor(XLColor.Red).Fill.SetBackgroundColor(XLColor.LightPink);
            range.AddConditionalFormat().WhenEquals("WARN").Font.SetFontColor(XLColor.DarkOrange).Fill.SetBackgroundColor(XLColor.LightYellow);
            range.AddConditionalFormat().WhenEquals("PASS").Font.SetFontColor(XLColor.DarkGreen).Fill.SetBackgroundColor(XLColor.LightGreen);
        }
    }

    private static void WriteValue(IXLCell target, StatementValue value)
    {
        switch (value.Kind)
        {
            case StatementScalarKind.Decimal when value.DecimalValue is { } number:
                target.Value = number == 0m ? 0m : number;
                break;
            case StatementScalarKind.Text:
                target.Value = value.TextValue ?? string.Empty;
                break;
            case StatementScalarKind.Date when value.DateValue is { } date:
                target.Value = date.ToDateTime(TimeOnly.MinValue);
                break;
            case StatementScalarKind.Boolean when value.BooleanValue is { } boolean:
                target.Value = boolean;
                break;
            default:
                throw new InvalidOperationException($"Incomplete present value of kind '{value.Kind}'.");
        }
    }

    private static bool ContainsEmptyAggregate(StatementFormula formula) => formula switch
    {
        StatementFormulaAggregate aggregate =>
            aggregate.Operands.IsEmpty || aggregate.Operands.Any(ContainsEmptyAggregate),
        StatementFormulaUnary unary => ContainsEmptyAggregate(unary.Operand),
        StatementFormulaBinary binary =>
            ContainsEmptyAggregate(binary.Left) || ContainsEmptyAggregate(binary.Right),
        StatementFormulaConditional conditional =>
            ContainsEmptyAggregate(conditional.Condition) ||
            ContainsEmptyAggregate(conditional.WhenTrue) ||
            ContainsEmptyAggregate(conditional.WhenFalse),
        _ => false
    };

    private static string ColumnName(int columnNumber)
    {
        var result = string.Empty;
        for (var value = columnNumber; value > 0; value = (value - 1) / 26)
            result = (char)('A' + ((value - 1) % 26)) + result;
        return result;
    }

    private static string EscapeFormulaText(string value) => value.Replace("\"", "\"\"");
}
