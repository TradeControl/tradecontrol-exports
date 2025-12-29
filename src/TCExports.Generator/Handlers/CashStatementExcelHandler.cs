using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using TCExports.Generator.Contracts;
using TCExports.Generator.Data;
using TCExports.Generator.Excel;

namespace TCExports.Generator.Handlers;

/// <summary>
/// Generates the Cash Statement Excel workbook with optional sections:
/// categories, summaries, VAT details, bank balances, and balance sheet.
/// </summary>
/// <remarks>
/// Layout:
/// - Header rows 1..4
/// - Period grid by year/month starting at column D
/// - Category sections (Trade, Money, Tax) + summaries and totals
/// - Optional VAT recurrence and period totals
/// - Optional bank balances and company total
/// - Optional balance sheet (point-in-time asset balances) and capital row
/// </remarks>
public sealed class CashStatementExcelHandler : IDocumentHandler
{
    private readonly ICashFlowRepository _repo;

    /// <summary>
    /// Initializes a new instance of the <see cref="CashStatementExcelHandler"/> class.
    /// </summary>
    /// <param name="repo">Repository for cash flow data access.</param>
    public CashStatementExcelHandler(ICashFlowRepository repo) => _repo = repo;

    /// <summary>
    /// Indicates whether the handler can generate the requested document.
    /// </summary>
    /// <param name="payload">The export payload containing document type and format.</param>
    /// <returns>true if the payload targets a cashflow Excel export; otherwise false.</returns>
    public bool CanHandle(ExportPayload payload) =>
        payload.DocumentType.Equals("cashflow", StringComparison.OrdinalIgnoreCase)
        && payload.Format.Equals("excel", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Builds and returns the Cash Statement Excel workbook as a base64 string.
    /// </summary>
    /// <param name="payload">Export parameters including SQL connection and feature flags.</param>
    /// <param name="ct">Cancellation token.</param>
    /// <returns>An <see cref="ExportResult"/> containing the generated file metadata and contents.</returns>
    /// <remarks>
    /// Feature flags (payload.Params):
    /// - includeActivePeriods: include future periods when computing values
    /// - includeOrderBook: include forecasted order book values
    /// - includeTaxAccruals: include tax accrual adjustments
    /// - includeVatDetails: render VAT recurrence and monthly totals blocks
    /// - includeBankBalances: render bank balances section
    /// - includeBalanceSheet: render balance sheet section
    /// </remarks>
    public async Task<ExportResult> HandleAsync(ExportPayload payload, CancellationToken ct)
    {
        var active = await _repo.GetActivePeriodAsync(payload.SqlConnection, ct: ct);
        var years = await _repo.GetActiveYearsAsync(payload.SqlConnection, ct: ct);
        var months = await _repo.GetMonthsAsync(payload.SqlConnection, ct: ct);
        var companyName = await _repo.GetCompanyNameAsync(payload.SqlConnection, ct: ct);

        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Cash Flow");

        RenderHeader(ws, active, companyName);
        var lastCol = RenderPeriodGrid(ws, years, months, active);

        bool includeActivePeriods = payload.Params.TryGetValue("includeActivePeriods", out var ap) && ap == "true";
        bool includeOrderBook = payload.Params.TryGetValue("includeOrderBook", out var ob) && ob == "true";
        bool includeTaxAccruals = payload.Params.TryGetValue("includeTaxAccruals", out var ta) && ta == "true";
        bool includeVatDetails = payload.Params.TryGetValue("includeVatDetails", out var ivd) && ivd == "true";
        bool includeBalanceSheet = payload.Params.TryGetValue("includeBalanceSheet", out var ibs) && ibs == "true";
        bool includeBankBalances = payload.Params.TryGetValue("includeBankBalances", out var ibb) && ibb == "true";

        var tradeCategories = await RenderCategoriesAsync(ws, payload.SqlConnection, years, months, CashType.Trade, includeActivePeriods, includeOrderBook, false, ct);
        RenderSummaryAfterCategories(ws, tradeCategories, years, months);

        var moneyCategories = await RenderCategoriesAsync(ws, payload.SqlConnection, years, months, CashType.Money, false, false, false, ct);
        RenderSummaryAfterCategories(ws, moneyCategories, years, months);

        await RenderSummaryTotalsBlockAsync(ws, payload.SqlConnection, CashType.Trade, ct);

        var taxCategories = await RenderCategoriesAsync(ws, payload.SqlConnection, years, months, CashType.Tax, includeActivePeriods, false, includeTaxAccruals, ct);
        RenderSummaryAfterCategories(ws, taxCategories, years, months);

        await RenderSummaryTotalsBlockAsync(ws, payload.SqlConnection, CashType.Tax, ct);
        await RenderTotalsFormulaAsync(ws, payload.SqlConnection, ct);
        await RenderExpressionsAsync(ws, payload.SqlConnection, ct);

        if (includeBankBalances)
            await RenderClosingBankBalanceAsync(ws, payload.SqlConnection, years, months, ct);

        // VAT reporting (resource-driven) when requested
        if (includeVatDetails)
        {
            await RenderVatRecurrenceTotalsAsync(ws, payload.SqlConnection, years, months, includeActivePeriods, includeTaxAccruals, ct);
            await RenderVatPeriodTotalsAsync(ws, payload.SqlConnection, years, months, includeActivePeriods, includeTaxAccruals, ct);
        }

        if (includeBalanceSheet)
        {
            await RenderBalanceSheetAsync(ws, payload.SqlConnection, years, months, ct);
            // Equity reconciliation intentionally disabled until opening balances are introduced.
            // await RenderEquityReconciliationAsync(ws, years, months);
        }

        // Hide helper column C and enforce widths
        ws.Column(3).Hide();
        ws.Column(1).Width = 9;   // A: Code
        ws.Column(2).Width = 25;  // B: Name
        for (int col = 4; col <= lastCol; col++)
            ws.Column(col).Width = 9;

        ws.SheetView.FreezeRows(4);
        ws.SheetView.FreezeColumns(3);

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

    /// <summary>
    /// Renders the header (title, company, date) and sets protection on header cells.
    /// </summary>
    private static void RenderHeader(IXLWorksheet ws, ActivePeriodDto active, string companyName)
    {
        ws.Cell(1, 1).Value = string.Format(Properties.Resources.TextStatementTitle, active.MonthName, active.Description);
        ws.Cell(1, 1).Style.Font.Bold = true;
        ws.Cell(1, 1).Style.Font.FontSize = 12;

        // A2: Company Name
        ws.Cell(2, 1).Value = companyName;
        ws.Cell(2, 1).Style.Font.Bold = true;
        ws.Cell(2, 1).Style.Font.FontSize = 10;

        ws.Cell(3, 1).Value = Properties.Resources.TextDate;
        ws.Cell(3, 2).Value = DateTime.Now.ToString("dd MMM HH:mm:ss");
        ws.Row(3).Style.Font.Bold = true;
        ws.Row(3).Style.Font.FontSize = 10;
        ws.Range("B2:B3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

        ws.Cell(4, 1).Value = Properties.Resources.TextCode;
        ws.Cell(4, 2).Value = Properties.Resources.TextName;
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

    /// <summary>
    /// Renders the year/month period grid headers and returns the last period column index.
    /// </summary>
    /// <param name="ws">Worksheet to write to.</param>
    /// <param name="years">Active years.</param>
    /// <param name="months">Months in a year (ordered).</param>
    /// <param name="active">Active period for highlighting.</param>
    /// <returns>Last used period column index.</returns>
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
            }

            curCol++;
            ws.Cell(3, curCol).Value = y.Description;
            ws.Cell(4, curCol).Value = "TOTALS";
            ws.Column(curCol).Style.Border.RightBorder = XLBorderStyleValues.Thick;
            ws.Column(curCol).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            ws.Column(curCol).Style.Protection.Locked = true;
            ws.Column(curCol).Style.NumberFormat.Format = "#,##0;[Red](#,##0);_-";
            ws.Column(curCol).Style.Font.Bold = true;
            curCol++;
        }
        return curCol - 1;
    }

    /// <summary>
    /// Renders category sections (rows per cash code, monthly values, and yearly totals) for the specified cash type.
    /// </summary>
    /// <param name="ws">Worksheet to write to.</param>
    /// <param name="connectionString">SQL connection string.</param>
    /// <param name="years">Active years.</param>
    /// <param name="months">Months in a year.</param>
    /// <param name="cashType">Cash type (Trade, Money, Tax).</param>
    /// <param name="includeActivePeriods">Include future periods when computing values.</param>
    /// <param name="includeOrderBook">Include forecasted order book values.</param>
    /// <param name="includeTaxAccruals">Include tax accrual adjustments.</param>
    /// <param name="ct">Cancellation token.</param>
    /// <returns>The list of rendered categories (for follow-up summary).</returns>
    private async Task<IReadOnlyList<FlowCategoryDto>> RenderCategoriesAsync(
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
            curRow += 2;
            ws.Cell(curRow, 1).Value = category.Category;
            ws.Row(curRow).Style.Font.Bold = true;
            ws.Row(curRow).Style.Border.TopBorder = XLBorderStyleValues.Medium;

            int startRow = curRow;

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
                        ws.Cell(curRow, curCol).Value = v?.InvoiceValue ?? 0m;
                        curCol++;
                    }

                    ws.Cell(curRow, curCol).FormulaA1 = $"=SUM({GetColLetter(startCol)}{curRow}:{GetColLetter(curCol - 1)}{curRow})";
                }
            }

            curRow++;
            ws.Cell(curRow, 1).Value = "TOTALS";
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

            // Category code marker in column C (for lookup).
            ws.Cell(curRow, 3).FormulaA1 = $"=\"{category.CategoryCode}\"";
        }

        return categories;
    }

    /// <summary>
    /// Renders the totals block listing category codes marked as totals for the specified cash type.
    /// </summary>
    private async Task RenderSummaryTotalsBlockAsync(IXLWorksheet ws, string connectionString, CashType cashType, CancellationToken ct)
    {
        var totals = await _repo.GetCategoriesByTypeAsync(connectionString, cashType, CategoryType.Total, ct: ct);
        if (totals.Count < 2) return;

        int curRow = (ws.LastRowUsed()?.RowNumber() ?? 3) + 2;

        ws.Row(curRow).Style.Border.TopBorder = XLBorderStyleValues.Thin;
        ws.Row(curRow).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        ws.Row(curRow).Style.Fill.BackgroundColor = XLColor.GreenYellow;
        ws.Row(curRow).Style.Font.Bold = true;
        ws.Cell(curRow, 1).Value = $"{totals.First().CashType} {Properties.Resources.TextTotals}";

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

    /// <summary>
    /// Renders a summary block referencing each category's totals row followed by a per-period total.
    /// </summary>
    private void RenderSummaryAfterCategories(
        IXLWorksheet ws,
        IReadOnlyList<FlowCategoryDto> categories,
        IReadOnlyList<ActiveYearDto> years,
        IReadOnlyList<MonthDto> months)
    {
        if (categories.Count < 2) return;

        int curRow = (ws.LastRowUsed()?.RowNumber() ?? 3) + 2;

        ws.Cell(curRow, 1).Value = Properties.Resources.TextSummary;
        ws.Row(curRow).Style.Font.Bold = true;
        ws.Row(curRow).Style.Border.TopBorder = XLBorderStyleValues.Thin;
        ws.Row(curRow).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        ws.Row(curRow).Style.Fill.BackgroundColor = XLColor.LightGreen;
        ws.Row(curRow).Style.Protection.Locked = true;

        int firstCol = 4;
        int lastCol = firstCol + (years.Count * (months.Count + 1)) - 1;

        int startRow = curRow;

        foreach (var category in categories)
        {
            curRow++;
            ws.Row(curRow).Style.Protection.Locked = true;

            ws.Cell(curRow, 1).Value = category.CategoryCode;
            ws.Cell(curRow, 2).Value = category.Category;

            int targetRow = FindCategoryTotalsRow(ws, category.CategoryCode);
            for (int col = firstCol; col <= lastCol; col++)
            {
                string colLetter = GetColLetter(col);
                ws.Cell(curRow, col).FormulaA1 = $"={colLetter}{targetRow}";
            }
        }

        ws.Row(curRow).Style.Border.BottomBorder = XLBorderStyleValues.Thick;

        curRow++;
        ws.Cell(curRow, 1).Value = Properties.Resources.TextPeriodTotal;
        ws.Row(curRow).Style.Font.Bold = true;
        for (int col = firstCol; col <= lastCol; col++)
        {
            string colLetter = GetColLetter(col);
            ws.Cell(curRow, col).FormulaA1 = $"=SUM({colLetter}{startRow + 1}:{colLetter}{curRow - 1})";
        }
        ws.Row(curRow).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
        ws.Row(curRow).Style.Border.TopBorder = XLBorderStyleValues.Thin;
        ws.Row(curRow).Style.Protection.Locked = true;
    }

    /// <summary>
    /// Populates formulas for total categories based on their component category codes.
    /// </summary>
    private async Task RenderTotalsFormulaAsync(
        IXLWorksheet ws,
        string connectionString,
        CancellationToken ct)
    {
        const int firstCol = 4;
        int lastCol = GetLastPeriodColumn(ws);

        var totalCategories = await _repo.GetCategoryTotalsAsync(connectionString, ct: ct);
        if (totalCategories.Count == 0) return;

        foreach (var total in totalCategories)
        {
            int targetRow = FindCategoryTotalsRow(ws, total.CategoryCode);
            if (targetRow <= 0) continue;

            var sumCodes = await _repo.GetCategoryTotalCodesAsync(connectionString, total.CategoryCode, ct: ct);
            if (sumCodes.Count == 0) continue;

            for (int col = firstCol; col <= lastCol; col++)
            {
                string colLetter = GetColLetter(col);
                string formula = string.Empty;

                foreach (var code in sumCodes)
                {
                    int srcRow = FindCategoryTotalsRow(ws, code);
                    if (srcRow <= 0) continue;
                    formula = formula.Length == 0
                        ? $"={colLetter}{srcRow}"
                        : $"{formula}+{colLetter}{srcRow}";
                }

                if (!string.IsNullOrEmpty(formula))
                    ws.Cell(targetRow, col).FormulaA1 = formula;
            }
        }
    }

    /// <summary>
    /// Renders analysis expressions (custom formulas referencing category totals),
    /// and reports syntax errors back to the repository.
    /// </summary>
    private async Task RenderExpressionsAsync(IXLWorksheet ws, string connectionString, CancellationToken ct)
    {
        var exprs = await _repo.GetCategoryExpressionsAsync(connectionString, ct: ct);
        if (exprs.Count == 0) return;

        int curRow = (ws.LastRowUsed()?.RowNumber() ?? 3) + 2;
        ws.Cell(curRow, 1).Value = Properties.Resources.TextAnalysis;
        ws.Row(curRow).Style.Font.Bold = true;
        ws.Row(curRow).Style.Border.TopBorder = XLBorderStyleValues.Thin;
        ws.Row(curRow).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        ws.Row(curRow).Style.Fill.BackgroundColor = XLColor.LimeGreen;
        ws.Row(curRow).Style.Protection.Locked = true;

        const int firstCol = 4;
        int lastCol = GetLastPeriodColumn(ws);

        foreach (var expr in exprs)
        {
            curRow++;
            ws.Row(curRow).Style.Protection.Locked = true;

            ws.Cell(curRow, 1).Value = expr.Category ?? string.Empty;

            string exprCategoryCode = await _repo.GetCategoryCodeFromNameAsync(connectionString, expr.Category ?? string.Empty, ct: ct);
            if (!string.IsNullOrWhiteSpace(exprCategoryCode))
            {
                ws.Cell(curRow, 3).FormulaA1 = $"=\"{exprCategoryCode}\"";
                ws.Cell(curRow, 3).Style.Font.FontColor = XLColor.LimeGreen;
            }

            if (!string.IsNullOrWhiteSpace(expr.Format))
                ws.Row(curRow).Style.NumberFormat.Format = expr.Format;

            ws.Range(curRow, firstCol, curRow, lastCol)
              .AddConditionalFormat()
              .WhenLessThan(0)
              .Font.SetFontColor(XLColor.Red);

            var template = expr.Expression ?? string.Empty;
            var tokens = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int i = 0;
            while (i < template.Length)
            {
                int lb = template.IndexOf('[', i);
                if (lb < 0) break;
                int rb = template.IndexOf(']', lb + 1);
                if (rb < 0) break;
                var token = template.Substring(lb + 1, rb - lb - 1).Trim();
                if (!string.IsNullOrEmpty(token)) tokens.Add(token);
                i = rb + 1;
            }

            var nameToCode = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var name in tokens)
            {
                var code = await _repo.GetCategoryCodeFromNameAsync(connectionString, name, ct: ct);
                if (string.IsNullOrWhiteSpace(code)) code = name;
                nameToCode[name] = code;
            }

            string normalizedTemplate = template;
            foreach (var kvp in nameToCode)
                normalizedTemplate = normalizedTemplate.Replace($"[{kvp.Key}]", $"[{kvp.Value}]");

            var errors = new List<string>();

            for (int col = firstCol; col <= lastCol; col++)
            {
                string colLetter = GetColLetter(col);
                string formula = normalizedTemplate;

                foreach (var code in nameToCode.Values)
                {
                    int srcRow = FindCategoryTotalsRow(ws, code);
                    if (srcRow <= 0)
                        errors.Add($"Reference not found: [{code}]");
                    var replacement = (srcRow > 0) ? $"{colLetter}{srcRow}" : "0";
                    formula = formula.Replace($"[{code}]", replacement);
                }

                try
                {
                    ws.Cell(curRow, col).FormulaA1 = $"={formula}";
                }
                catch (Exception ex)
                {
                    var cell = ws.Cell(curRow, col);
                    try { cell.FormulaA1 = "=\"\""; } catch { /* ignore */ }
                    cell.Style.Fill.BackgroundColor = XLColor.Red;
                    errors.Add($"Excel formula error: {ex.Message}");
                }
            }

            if (!string.IsNullOrWhiteSpace(exprCategoryCode))
            {
                bool isError = errors.Count > 0;
                string? errorMessage = isError ? string.Join("; ", errors.Distinct()) : null;
                await _repo.SetCategoryExpressionStatusAsync(connectionString, exprCategoryCode, isError, errorMessage, ct: ct);
            }
        }

        ws.Row(curRow).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
    }

    /// <summary>
    /// Renders closing bank balances per account and sums into a company total.
    /// </summary>
    private async Task RenderClosingBankBalanceAsync(
        IXLWorksheet ws,
        string connectionString,
        IReadOnlyList<ActiveYearDto> years,
        IReadOnlyList<MonthDto> months,
        CancellationToken ct)
    {
        int curRow = (ws.LastRowUsed()?.RowNumber() ?? 3) + 2;
        ws.Cell(curRow, 1).Value = Properties.Resources.TextClosingBalances.ToUpperInvariant();
        ws.Row(curRow).Style.Font.Bold = true;
        ws.Row(curRow).Style.Border.TopBorder = XLBorderStyleValues.Thin;
        ws.Row(curRow).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        ws.Row(curRow).Style.Fill.BackgroundColor = XLColor.LightSteelBlue;
        ws.Row(curRow).Style.Protection.Locked = true;

        const int firstCol = 4;
        int lastCol = GetLastPeriodColumn(ws);

        var accounts = await _repo.GetBankAccountsAsync(connectionString, ct: ct);
        if (accounts.Count == 0)
        {
            curRow++;
            ws.Cell(curRow, 1).Value = "(no bank accounts)";
            return;
        }

        var perAccount = new Dictionary<string, Dictionary<(short Year, byte Month), decimal>>(StringComparer.OrdinalIgnoreCase);
        foreach (var acct in accounts)
        {
            var map = new Dictionary<(short, byte), decimal>();
            var balances = await _repo.GetBankBalancesAsync(connectionString, acct.AccountCode, ct: ct);
            foreach (var b in balances)
            {
                var key = (b.YearNumber, b.MonthNumber);
                map[key] = b.Balance;
            }
            perAccount[acct.AccountCode] = map;
        }

        int m12Index = months.Select((m, idx) => new { m.MonthNumber, idx }).FirstOrDefault(x => x.MonthNumber == 12)?.idx ?? -1;

        foreach (var acct in accounts)
        {
            curRow++;
            ws.Cell(curRow, 1).Value = acct.AccountCode;
            ws.Cell(curRow, 2).Value = acct.AccountName;

            int curCol = firstCol;
            foreach (var y in years)
            {
                int yearStartCol = curCol;

                foreach (var m in months)
                {
                    var key = (y.YearNumber, (byte)m.MonthNumber);
                    decimal value = 0m;
                    if (perAccount.TryGetValue(acct.AccountCode, out var map) && map.TryGetValue(key, out var bal))
                        value = bal;

                    ws.Cell(curRow, curCol).Value = value;
                    curCol++;
                }

                if (m12Index >= 0)
                {
                    int m12Col = yearStartCol + m12Index;
                    ws.Cell(curRow, curCol).FormulaA1 = $"={GetColLetter(m12Col)}{curRow}";
                }
                curCol++;
            }
        }

        curRow++;
        ws.Cell(curRow, 1).Value = Properties.Resources.TextCompanyBalance.ToUpperInvariant();
        ws.Row(curRow).Style.Font.Bold = true;
        ws.Row(curRow).Style.Border.TopBorder = XLBorderStyleValues.Thin;
        ws.Row(curRow).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        for (int col = firstCol; col <= lastCol; col++)
        {
            var colLetter = GetColLetter(col);
            int firstAcctRow = curRow - accounts.Count;
            ws.Cell(curRow, col).FormulaA1 = $"=SUM({colLetter}{firstAcctRow}:{colLetter}{curRow - 1})";
        }
    }

    /// <summary>
    /// Renders VAT recurrence totals by period with optional accrual adjustments.
    /// </summary>
    private async Task RenderVatRecurrenceTotalsAsync(
        IXLWorksheet ws,
        string connectionString,
        IReadOnlyList<ActiveYearDto> years,
        IReadOnlyList<MonthDto> months,
        bool includeActivePeriods,
        bool includeTaxAccruals,
        CancellationToken ct)
    {
        int curRow = (ws.LastRowUsed()?.RowNumber() ?? 3) + 2;
        var vatType = await _repo.GetVatRecurrenceTypeAsync(connectionString, ct: ct);
        ws.Cell(curRow, 1).Value = $"{Properties.Resources.TextVatDueTitle} {vatType}".ToUpperInvariant();
        ws.Row(curRow).Style.Font.Bold = true;
        ws.Row(curRow).Style.Border.TopBorder = XLBorderStyleValues.Thin;
        ws.Row(curRow).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        ws.Row(curRow).Style.Fill.BackgroundColor = XLColor.Orchid;
        ws.Row(curRow).Style.Protection.Locked = true;

        int startRow = curRow + 1;
        var labels = new[]
        {
            Properties.Resources.TextVatHomeSales,
            Properties.Resources.TextVatHomePurchases,
            Properties.Resources.TextVatExportSales,
            Properties.Resources.TextVatExportPurchases,
            Properties.Resources.TextVatHomeSalesVat,
            Properties.Resources.TextVatHomePurchasesVat,
            Properties.Resources.TextVatExportSalesVat,
            Properties.Resources.TextVatExportPurchasesVat,
            Properties.Resources.TextVatAdjustment,
            Properties.Resources.TextVatDue
        };
        for (int i = 0; i < labels.Length; i++)
            ws.Cell(startRow + i, 1).Value = labels[i].ToUpperInvariant();
        ws.Row(startRow + labels.Length - 1).Style.Font.Bold = true;
        ws.Row(startRow + labels.Length - 1).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
        ws.Row(startRow).Style.Border.TopBorder = XLBorderStyleValues.Thin;

        const int firstCol = 4;
        int curCol = firstCol;
        short currentYear = 0;
        int blockStartCol = firstCol;

        var recurrence = await _repo.GetVatRecurrenceAsync(connectionString, ct: ct);
        foreach (var p in recurrence)
        {
            if (currentYear == 0)
            {
                curCol = firstCol;
                currentYear = p.YearNumber;
                blockStartCol = curCol;
            }

            if (currentYear != p.YearNumber)
            {
                for (int offset = 0; offset < labels.Length; offset++)
                    ws.Cell(startRow + offset, curCol).FormulaA1 =
                        $"=SUM({GetColLetter(blockStartCol)}{startRow + offset}:{GetColLetter(curCol - 1)}{startRow + offset})";
                currentYear = p.YearNumber;
                curCol++;
                blockStartCol = curCol;
            }

            if (includeActivePeriods || p.StartOn <= DateTime.UtcNow)
            {
                ws.Cell(startRow + 0, curCol).Value = p.HomeSales;
                ws.Cell(startRow + 1, curCol).Value = p.HomePurchases;
                ws.Cell(startRow + 2, curCol).Value = p.ExportSales;
                ws.Cell(startRow + 3, curCol).Value = p.ExportPurchases;
                ws.Cell(startRow + 4, curCol).Value = p.HomeSalesVat;
                ws.Cell(startRow + 5, curCol).Value = p.HomePurchasesVat;
                ws.Cell(startRow + 6, curCol).Value = p.ExportSalesVat;
                ws.Cell(startRow + 7, curCol).Value = p.ExportPurchasesVat;
                ws.Cell(startRow + 8, curCol).Value = p.VatAdjustment;
                ws.Cell(startRow + 9, curCol).Value = p.VatDue;
            }
            else
            {
                ws.Range(startRow, curCol, startRow + (labels.Length - 1), curCol).Value = 0;
            }
            curCol++;
        }

        for (int offset = 0; offset < labels.Length; offset++)
            ws.Cell(startRow + offset, curCol).FormulaA1 =
                $"=SUM({GetColLetter(blockStartCol)}{startRow + offset}:{GetColLetter(curCol - 1)}{startRow + offset})";

        if (includeTaxAccruals)
        {
            var accruals = await _repo.GetVatRecurrenceAccrualsAsync(connectionString, ct: ct);
            short accrualYear = 0;
            int accrualCol = firstCol;
            foreach (var a in accruals)
            {
                if (accrualYear == 0)
                    accrualYear = a.YearNumber;

                if (accrualYear != a.YearNumber)
                {
                    accrualYear = a.YearNumber;
                    accrualCol++;
                }

                AddDecimal(ws, startRow + 4, accrualCol, a.HomeSalesVat);
                AddDecimal(ws, startRow + 5, accrualCol, a.HomePurchasesVat);
                AddDecimal(ws, startRow + 6, accrualCol, a.ExportSalesVat);
                AddDecimal(ws, startRow + 7, accrualCol, a.ExportPurchasesVat);
                AddDecimal(ws, startRow + 9, accrualCol, a.VatDue);

                accrualCol++;
            }
        }
    }

    /// <summary>
    /// Renders VAT monthly totals with optional accrual adjustments.
    /// </summary>
    private async Task RenderVatPeriodTotalsAsync(
        IXLWorksheet ws,
        string connectionString,
        IReadOnlyList<ActiveYearDto> years,
        IReadOnlyList<MonthDto> months,
        bool includeActivePeriods,
        bool includeTaxAccruals,
        CancellationToken ct)
    {
        int curRow = (ws.LastRowUsed()?.RowNumber() ?? 3) + 2;

        ws.Cell(curRow, 1).Value = $"{Properties.Resources.TextVatDueTitle} {Properties.Resources.TextTotals}".ToUpperInvariant();
        ws.Row(curRow).Style.Font.Bold = true;
        ws.Row(curRow).Style.Border.TopBorder = XLBorderStyleValues.Thin;
        ws.Row(curRow).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        ws.Row(curRow).Style.Fill.BackgroundColor = XLColor.Orchid;
        ws.Row(curRow).Style.Protection.Locked = true;

        int startRow = curRow + 1;
        var labelsMonthly = new[]
        {
            Properties.Resources.TextVatHomeSales,
            Properties.Resources.TextVatHomePurchases,
            Properties.Resources.TextVatExportSales,
            Properties.Resources.TextVatExportPurchases,
            Properties.Resources.TextVatHomeSalesVat,
            Properties.Resources.TextVatHomePurchasesVat,
            Properties.Resources.TextVatExportSalesVat,
            Properties.Resources.TextVatExportPurchasesVat,
            Properties.Resources.TextVatDue
        };

        for (int i = 0; i < labelsMonthly.Length; i++)
            ws.Cell(startRow + i, 1).Value = labelsMonthly[i].ToUpperInvariant();
        ws.Row(startRow + labelsMonthly.Length - 1).Style.Font.Bold = true;
        ws.Row(startRow + labelsMonthly.Length - 1).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
        ws.Row(startRow).Style.Border.TopBorder = XLBorderStyleValues.Thin;

        const int firstCol = 4;
        int curCol = firstCol;
        short currentYear = 0;
        int blockStartCol = firstCol;

        var monthlyTotals = await _repo.GetVatPeriodTotalsAsync(connectionString, ct: ct);
        foreach (var p in monthlyTotals)
        {
            if (currentYear == 0)
            {
                curCol = firstCol;
                currentYear = p.YearNumber;
                blockStartCol = curCol;
            }

            if (currentYear != p.YearNumber)
            {
                for (int offset = 0; offset < labelsMonthly.Length; offset++)
                    ws.Cell(startRow + offset, curCol).FormulaA1 =
                        $"=SUM({GetColLetter(blockStartCol)}{startRow + offset}:{GetColLetter(curCol - 1)}{startRow + offset})";
                currentYear = p.YearNumber;
                curCol++;
                blockStartCol = curCol;
            }

            if (includeActivePeriods || p.StartOn <= DateTime.UtcNow)
            {
                ws.Cell(startRow + 0, curCol).Value = p.HomeSales;
                ws.Cell(startRow + 1, curCol).Value = p.HomePurchases;
                ws.Cell(startRow + 2, curCol).Value = p.ExportSales;
                ws.Cell(startRow + 3, curCol).Value = p.ExportPurchases;
                ws.Cell(startRow + 4, curCol).Value = p.HomeSalesVat;
                ws.Cell(startRow + 5, curCol).Value = p.HomePurchasesVat;
                ws.Cell(startRow + 6, curCol).Value = p.ExportSalesVat;
                ws.Cell(startRow + 7, curCol).Value = p.ExportPurchasesVat;
                ws.Cell(startRow + 8, curCol).Value = p.VatDue;
            }
            else
            {
                ws.Range(startRow, curCol, startRow + (labelsMonthly.Length - 1), curCol).Value = 0;
            }

            curCol++;
        }

        for (int offset = 0; offset < labelsMonthly.Length; offset++)
            ws.Cell(startRow + offset, curCol).FormulaA1 =
                $"=SUM({GetColLetter(blockStartCol)}{startRow + offset}:{GetColLetter(curCol - 1)}{startRow + offset})";

        if (includeTaxAccruals)
        {
            var accruals = await _repo.GetVatPeriodAccrualsAsync(connectionString, ct: ct);
            short accrualYear = 0;
            int accrualCol = firstCol;
            foreach (var a in accruals)
            {
                if (accrualYear == 0)
                    accrualYear = a.YearNumber;

                if (accrualYear != a.YearNumber)
                {
                    accrualYear = a.YearNumber;
                    accrualCol++;
                }

                AddDecimal(ws, startRow + 4, accrualCol, a.HomeSalesVat);
                AddDecimal(ws, startRow + 5, accrualCol, a.HomePurchasesVat);
                AddDecimal(ws, startRow + 6, accrualCol, a.ExportSalesVat);
                AddDecimal(ws, startRow + 7, accrualCol, a.ExportPurchasesVat);
                AddDecimal(ws, startRow + 8, accrualCol, a.VatDue);

                accrualCol++;
            }
        }
    }

    /// <summary>
    /// Renders the balance sheet section: asset balances per month, carried-forward year totals,
    /// and a capital row that sums assets per column.
    /// </summary>
    private async Task RenderBalanceSheetAsync(
        IXLWorksheet ws,
        string connectionString,
        IReadOnlyList<ActiveYearDto> years,
        IReadOnlyList<MonthDto> months,
        CancellationToken ct)
    {
        var entries = await _repo.GetBalanceSheetAsync(connectionString, ct: ct);
        if (entries.Count == 0) return;

        int curRow = (ws.LastRowUsed()?.RowNumber() ?? 3) + 2;

        // Header
        ws.Cell(curRow, 1).Value = Properties.Resources.TextBalanceSheet.ToUpperInvariant();
        ws.Row(curRow).Style.Font.Bold = true;
        ws.Row(curRow).Style.Border.TopBorder = XLBorderStyleValues.Thin;
        ws.Row(curRow).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        ws.Row(curRow).Style.Fill.BackgroundColor = XLColor.Red;
        ws.Row(curRow).Style.Font.FontColor = XLColor.White;
        ws.Row(curRow).Style.Protection.Locked = true;

        const int firstCol = 4;
        int lastCol = GetLastPeriodColumn(ws);

        int startRow = curRow;

        // Group by (AssetCode, AssetName), preserve first-seen order
        var groups = new Dictionary<string, Dictionary<(short Year, byte Month), decimal>>(StringComparer.OrdinalIgnoreCase);
        var order = new List<(string Code, string Name, string Key)>();

        foreach (var e in entries)
        {
            string key = $"{e.AssetCode}\u001F{e.AssetName}";
            if (!groups.TryGetValue(key, out var map))
            {
                map = new Dictionary<(short, byte), decimal>();
                groups[key] = map;
                order.Add((e.AssetCode, e.AssetName, key));
            }

            var ym = (e.YearNumber, e.MonthNumber);
            map[ym] = (map.TryGetValue(ym, out var existing) ? existing : 0m) + e.Balance;
        }

        int m12Index = months.Select((m, idx) => new { m.MonthNumber, idx }).FirstOrDefault(x => x.MonthNumber == 12)?.idx ?? -1;

        foreach (var item in order)
        {
            curRow++;
            ws.Cell(curRow, 1).Value = item.Code;
            ws.Cell(curRow, 2).Value = item.Name;

            var map = groups[item.Key];

            for (int yIndex = 0; yIndex < years.Count; yIndex++)
            {
                var y = years[yIndex];
                int yearStartCol = firstCol + (yIndex * (months.Count + 1));
                int curCol = yearStartCol;
                int lastNonEmptyCol = -1;

                for (int mIndex = 0; mIndex < months.Count; mIndex++)
                {
                    var m = months[mIndex];
                    if (map.TryGetValue((y.YearNumber, (byte)m.MonthNumber), out var bal))
                    {
                        ws.Cell(curRow, curCol).Value = bal;
                        lastNonEmptyCol = curCol;
                    }
                    curCol++;
                }

                int yearTotalCol = yearStartCol + months.Count;
                if (lastNonEmptyCol > 0)
                {
                    ws.Cell(curRow, yearTotalCol).FormulaA1 = $"={GetColLetter(lastNonEmptyCol)}{curRow}";
                }
            }
        }

        curRow++;
        ws.Cell(curRow, 1).Value = Properties.Resources.TextCapital.ToUpperInvariant();
        ws.Row(curRow).Style.Font.Bold = true;
        ws.Row(curRow).Style.Border.TopBorder = XLBorderStyleValues.Thin;
        ws.Row(curRow).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
        ws.Row(curRow).Style.Protection.Locked = true;

        int firstAssetRow = startRow + 1;
        for (int col = firstCol; col <= lastCol; col++)
        {
            string colLetter = GetColLetter(col);
            ws.Cell(curRow, col).FormulaA1 = $"=SUM({colLetter}{firstAssetRow}:{colLetter}{curRow - 1})";
        }
    }

    /// <summary>
    /// Renders an equity reconciliation block tying P&amp;L net income to change in equity.
    /// Currently disabled until opening balances are introduced.
    /// </summary>
    private async Task RenderEquityReconciliationAsync(
            IXLWorksheet ws,
            IReadOnlyList<ActiveYearDto> years,
            IReadOnlyList<MonthDto> months)
    {
        // Find P&L Period Total row
        var periodTotalRow = ws.Column(1).CellsUsed()
            .FirstOrDefault(c => string.Equals(c.GetString(), Properties.Resources.TextPeriodTotal, StringComparison.OrdinalIgnoreCase))?.Address.RowNumber ?? -1;
        if (periodTotalRow <= 0) return;

        // Find Balance Sheet capital row
        var capitalRow = ws.Column(1).CellsUsed()
            .LastOrDefault(c => string.Equals(c.GetString(), Properties.Resources.TextCapital.ToUpperInvariant(), StringComparison.OrdinalIgnoreCase))?.Address.RowNumber
            ?? ws.Column(1).CellsUsed().LastOrDefault(c => string.Equals(c.GetString(), Properties.Resources.TextCapital, StringComparison.OrdinalIgnoreCase))?.Address.RowNumber
            ?? -1;
        if (capitalRow <= 0) return;

        // Header
        int curRow = (ws.LastRowUsed()?.RowNumber() ?? capitalRow) + 2;
        ws.Cell(curRow, 1).Value = Properties.Resources.TextReconEquityRecon;
        ws.Row(curRow).Style.Font.Bold = true;
        ws.Row(curRow).Style.Border.TopBorder = XLBorderStyleValues.Thin;
        ws.Row(curRow).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        ws.Row(curRow).Style.Fill.BackgroundColor = XLColor.LightGray;
        ws.Row(curRow).Style.Protection.Locked = true;

        // Labels
        curRow++;
        ws.Cell(curRow, 1).Value = Properties.Resources.TextReconNetIncome;
        curRow++;
        ws.Cell(curRow, 1).Value = Properties.Resources.TextReconEquityChange;
        curRow++;
        ws.Cell(curRow, 1).Value = Properties.Resources.TextReconDifference;

        const int firstCol = 4;
        int lastCol = GetLastPeriodColumn(ws);
        int colsPerYear = months.Count + 1;

        for (int col = firstCol; col <= lastCol; col++)
        {
            string colLetter = GetColLetter(col);

            int reconciliationHeaderRow = curRow - 2;
            int netRow = reconciliationHeaderRow + 1;
            int changeEquityRow = reconciliationHeaderRow + 2;
            int diffRow = reconciliationHeaderRow + 3;

            // Net Income from Period Total (always shown)
            ws.Cell(netRow, col).FormulaA1 = $"={colLetter}{periodTotalRow}";

            // Determine year block for this column
            int zeroBased = col - firstCol;
            int yearIndex = zeroBased / colsPerYear;
            int yearStartCol = firstCol + (yearIndex * colsPerYear);
            int yearTotalsCol = yearStartCol + months.Count;

            // For the first year, opening equity is unknown; leave change/difference blank
            if (yearIndex == 0)
            {
                ws.Cell(changeEquityRow, col).Clear();
                ws.Cell(diffRow, col).Clear();
                continue;
            }

            // Opening = first month col of the year block, Ending = year totals col (carried forward month 12)
            string openingLetter = GetColLetter(yearStartCol);
            string endingLetter = GetColLetter(yearTotalsCol);

            // Change in Equity = Ending - Opening (from Balance Sheet capital row)
            ws.Cell(changeEquityRow, col).FormulaA1 = $"={endingLetter}{capitalRow}-{openingLetter}{capitalRow}";

            // Difference = Change in Equity - Net Income
            ws.Cell(diffRow, col).FormulaA1 = $"={GetColLetter(col)}{changeEquityRow}-{GetColLetter(col)}{netRow}";

            // Highlight differences when non-zero
            var addr = $"{GetColLetter(col)}{diffRow}";
            ws.Cell(diffRow, col)
              .AddConditionalFormat()
              .WhenIsTrue($"{addr}<>0")
              .Font.SetFontColor(XLColor.Red);
        }

        ws.Row(curRow).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
        await Task.CompletedTask;
    }

    #region utils
    /// <summary>
    /// Returns the last period column index, preferring the header row.
    /// </summary>
    private static int GetLastPeriodColumn(IXLWorksheet ws)
    {
        var headerLast = ws.Row(4).LastCellUsed();
        if (headerLast != null) return headerLast.Address.ColumnNumber;

        var anyLast = ws.LastColumnUsed();
        return anyLast?.ColumnNumber() ?? 4;
    }

    /// <summary>
    /// Finds the row number of a category totals row by matching its category code marker in column C.
    /// </summary>
    private static int FindCategoryTotalsRow(IXLWorksheet ws, string categoryCode)
    {
        foreach (var cell in ws.Column(3).CellsUsed())
        {
            var val = cell.HasFormula ? cell.Value.ToString() : cell.GetString();
            if (string.Equals(val, categoryCode, StringComparison.OrdinalIgnoreCase))
                return cell.Address.RowNumber;
        }
        return -1;
    }

    /// <summary>
    /// Converts a 1-based column index to an Excel column letter (e.g., 1 -> A, 28 -> AB).
    /// </summary>
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

    /// <summary>
    /// Adds a nullable decimal delta to a cell's current numeric value, treating blank as zero.
    /// </summary>
    private static void AddDecimal(IXLWorksheet ws, int row, int col, decimal? delta)
    {
        if (delta is null) return;
        var cell = ws.Cell(row, col);
        var current = cell.GetValue<decimal>();
        cell.Value = current + delta.Value;
    }
    #endregion
}