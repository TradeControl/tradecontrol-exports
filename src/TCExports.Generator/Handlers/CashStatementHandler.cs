using ClosedXML.Excel;
using TCExports.Generator.Contracts;
using TCExports.Generator.Data;
using TCExports.Generator.Excel;

namespace TCExports.Generator.Handlers;

public sealed class CashStatementHandler : IDocumentHandler
{
    private readonly ICashFlowRepository _repo;
    public CashStatementHandler(ICashFlowRepository repo) => _repo = repo;

    public bool CanHandle(ExportPayload payload) =>
        payload.DocumentType.Equals("cashflow", StringComparison.OrdinalIgnoreCase)
        && payload.Format.Equals("excel", StringComparison.OrdinalIgnoreCase);

    public async Task<ExportResult> HandleAsync(ExportPayload payload, CancellationToken ct)
    {
        var active = await _repo.GetActivePeriodAsync(payload.SqlConnection, ct: ct);
        var years = await _repo.GetActiveYearsAsync(payload.SqlConnection, ct: ct);
        var months = await _repo.GetMonthsAsync(payload.SqlConnection, ct: ct);

        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Cash Flow");

        RenderHeader(ws, active, payload.UserName);
        RenderPeriodGrid(ws, years, months, active);

        bool includeActivePeriods = payload.Params.TryGetValue("includeActivePeriods", out var ap) && ap == "true";
        bool includeOrderBook = payload.Params.TryGetValue("includeOrderBook", out var ob) && ob == "true";
        bool includeTaxAccruals = payload.Params.TryGetValue("includeTaxAccruals", out var ta) && ta == "true";

        // Slice sequence:
        await RenderCategoriesAsync(ws, payload.SqlConnection, years, months, CashType.Trade, includeActivePeriods, includeOrderBook, false, ct);
        await RenderCategoriesAsync(ws, payload.SqlConnection, years, months, CashType.Money, false, false, false, ct);
        await RenderSummaryTotalsBlockAsync(ws, payload.SqlConnection, CashType.Trade, ct);

        await RenderCategoriesAsync(ws, payload.SqlConnection, years, months, CashType.Tax, includeActivePeriods, false, includeTaxAccruals, ct);
        await RenderSummaryTotalsBlockAsync(ws, payload.SqlConnection, CashType.Tax, ct);

        ws.SheetView.FreezeRows(4);
        ws.SheetView.FreezeColumns(3);
        ws.Columns().AdjustToContents();

        SettingsSheetBuilder.AddSettingsWorksheet(wb, payload);

        using var ms = new MemoryStream();
        wb.SaveAs(ms);

        var ts = DateTime.UtcNow.ToString("yyyyMMdd_HHmmss");
        var safeDoc = System.Text.RegularExpressions.Regex.Replace(payload.DocumentType, @"[^A-Za-z0-9_\-]+", "_");
        var fileName = $"{payload.UserName}_{safeDoc}_{ts}.xlsx";

        return new ExportResult
        {
            Status = "success",
            FileName = fileName,
            FileContent = Convert.ToBase64String(ms.ToArray())
        };
    }

    private static void RenderHeader(IXLWorksheet ws, ActivePeriodDto active, string userName)
    {
        ws.Cell(1, 1).Value = $"Cash Flow Statement: {active.MonthName} {active.Description}";
        ws.Cell(1, 1).Style.Font.Bold = true;
        ws.Cell(1, 1).Style.Font.FontSize = 12;

        ws.Cell(2, 1).Value = userName;
        ws.Cell(2, 1).Style.Font.Bold = true;
        ws.Cell(2, 1).Style.Font.FontSize = 10;

        ws.Cell(3, 1).Value = "Date";
        ws.Cell(3, 2).Value = DateTime.Now.ToString("dd MMM HH:mm:ss");
        ws.Row(3).Style.Font.Bold = true;
        ws.Row(3).Style.Font.FontSize = 10;
        ws.Range("B2:B3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

        ws.Cell(4, 1).Value = "Code";
        ws.Cell(4, 2).Value = "Name";
        ws.Row(4).Style.Font.Bold = true;
        ws.Row(4).Style.Font.FontSize = 8;
        ws.Row(4).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
        ws.Row(4).Style.Border.TopBorder = XLBorderStyleValues.Thin;
        ws.Column(2).Style.Border.RightBorder = XLBorderStyleValues.Thick;

        ws.Protect();
        ws.Style.Protection.Locked = false;
        ws.Columns(1, 3).Style.Protection.Locked = true;
        ws.Rows(1, 4).Style.Protection.Locked = true;
    }

    private static int RenderPeriodGrid(IXLWorksheet ws, IReadOnlyList<ActiveYearDto> years, IReadOnlyList<MonthDto> months, ActivePeriodDto active)
    {
        int firstCol = 4;
        int curCol = firstCol;

        foreach (var y in years)
        {
            ws.Cell(3, curCol).Value = $"{y.Description} ({y.CashStatus})";
            curCol -= 1;

            foreach (var m in months)
            {
                curCol++;
                ws.Cell(4, curCol).Value = m.MonthName;
                if (y.YearNumber == active.YearNumber && m.MonthNumber == active.MonthNumber)
                    ws.Column(curCol).Style.Fill.BackgroundColor = XLColor.Yellow;

                ws.Column(curCol).Style.NumberFormat.Format = "#,##0;[Red](#,##0);_-";
                ws.Column(curCol).Width = 11;
            }

            curCol++;
            ws.Cell(3, curCol).Value = y.Description;
            ws.Cell(4, curCol).Value = "Totals";
            ws.Column(curCol).Style.Border.RightBorder = XLBorderStyleValues.Thick;
            ws.Column(curCol).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            ws.Column(curCol).Style.Protection.Locked = true;
            ws.Column(curCol).Style.NumberFormat.Format = "#,##0;[Red](#,##0);_-";
            ws.Column(curCol).Width = 11;
            ws.Column(curCol).Style.Font.Bold = true;
            curCol++;
        }
        return curCol - 1;
    }

    private async Task RenderCategoriesAsync(
        IXLWorksheet ws,
        string connectionString,
        IReadOnlyList<ActiveYearDto> years,
        IReadOnlyList<MonthDto> months,
        CashType cashType,
        bool includeActivePeriods,
        bool includeOrderBook,
        bool includeTaxAccruals,
        CancellationToken ct)
    {
        const int firstCol = 4;
        int curRow = ws.LastRowUsed()?.RowNumber() ?? 3;

        var categories = await _repo.GetCategoriesAsync(connectionString, cashType, ct: ct);
        foreach (var category in categories)
        {
            // Category header
            curRow += 2;
            ws.Cell(curRow, 1).Value = category.Category;
            ws.Row(curRow).Style.Font.Bold = true;
            ws.Row(curRow).Style.Border.TopBorder = XLBorderStyleValues.Medium;

            int startRow = curRow;

            // Cash codes and values
            var cashCodes = await _repo.GetCashCodesAsync(connectionString, category.CategoryCode, ct: ct);
            foreach (var code in cashCodes)
            {
                curRow++;
                ws.Cell(curRow, 1).Value = code.CashCode;
                ws.Cell(curRow, 2).Value = code.CashDescription;

                int yearCount = 0;
                foreach (var year in years)
                {
                    yearCount++;
                    int startCol = firstCol + ((yearCount - 1) * (months.Count + 1));
                    int curCol = startCol;

                    var values = await _repo.GetCashCodeValuesAsync(
                        connectionString,
                        code.CashCode,
                        year.YearNumber,
                        includeActivePeriods,
                        includeOrderBook,
                        includeTaxAccruals,
                        ct: ct);

                    for (int mIndex = 0; mIndex < months.Count; mIndex++)
                    {
                        var v = (mIndex < values.Count) ? values[mIndex] : null;
                        var rounded = v is null ? 0m : Math.Round(v.InvoiceValue, 0, MidpointRounding.AwayFromZero);
                        ws.Cell(curRow, curCol).Value = rounded;
                        curCol++;
                    }

                    // Yearly total
                    ws.Cell(curRow, curCol).FormulaA1 = $"=SUM({GetColLetter(startCol)}{curRow}:{GetColLetter(curCol - 1)}{curRow})";
                }
            }

            // Per-category totals row styling and formulas
            curRow++;
            ws.Cell(curRow, 1).Value = "Totals";
            ws.Row(curRow).Style.Font.Bold = true;
            ws.Row(curRow).Style.Border.TopBorder = XLBorderStyleValues.Thin;
            ws.Row(curRow).Style.Border.BottomBorder = XLBorderStyleValues.Thin;

            var totalsRow = ws.Row(curRow);
            if (category.CashPolarityCode == 0) totalsRow.Style.Fill.BackgroundColor = XLColor.LightSalmon;
            else if (category.CashPolarityCode == 1) totalsRow.Style.Fill.BackgroundColor = XLColor.CornflowerBlue;
            else totalsRow.Style.Fill.BackgroundColor = XLColor.LightGray;
            ws.Cell(curRow, 3).Style.Font.FontColor = totalsRow.Style.Fill.BackgroundColor;

            int totalCols = years.Count * (months.Count + 1);
            for (int i = 0; i < totalCols; i++)
            {
                int col = firstCol + i;
                var sum = $"=SUM({GetColLetter(col)}{startRow + 1}:{GetColLetter(col)}{curRow - 1})";
                if (category.CashPolarityCode == 0) sum += "*-1";
                ws.Cell(curRow, col).FormulaA1 = sum;
            }
        }
    }

    private async Task RenderSummaryTotalsBlockAsync(IXLWorksheet ws, string connectionString, CashType cashType, CancellationToken ct)
    {
        var totals = await _repo.GetCategoriesByTypeAsync(connectionString, cashType, CategoryType.Total, ct: ct);
        if (totals.Count < 2) return;

        int curRow = (ws.LastRowUsed()?.RowNumber() ?? 3) + 2;

        ws.Row(curRow).Style.Border.TopBorder = XLBorderStyleValues.Thin;
        ws.Row(curRow).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        ws.Cell(curRow, 1).Value = $"{totals.First().CashType} Totals";
        ws.Row(curRow).Style.Font.Bold = true;

        foreach (var total in totals)
        {
            curRow++;
            ws.Row(curRow).Style.Protection.Locked = true;
            ws.Cell(curRow, 1).FormulaA1 = $"=\"{total.CategoryCode}\"";
            ws.Cell(curRow, 2).Value = total.Category;
            ws.Cell(curRow, 3).FormulaA1 = $"=\"{total.CategoryCode}\"";
        }

        ws.Row(curRow).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
    }

    private static string GetColLetter(int col)
    {
        int dividend = col;
        string name = string.Empty;
        while (dividend > 0)
        {
            int modulo = (dividend - 1) % 26;
            name = Convert.ToChar('A' + modulo) + name;
            dividend = (dividend - modulo) / 26;
        }
        return name;
    }
}