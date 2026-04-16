using ClosedXML.Excel;
using TCExports.Generator.Data;

namespace TCExports.Generator.Excel;

public static class ReconciliationSheetBuilder
{
    public static async Task AddReconciliationWorksheetAsync(
        XLWorkbook workbook,
        ICashFlowRepository repo,
        string sqlConnection,
        int commandTimeoutSeconds,
        CancellationToken ct)
    {
        var rows = await repo.GetEquityReconciliationByYearAsync(sqlConnection, commandTimeoutSeconds, ct);

        var ws = workbook.Worksheets.Add("Reconciliation");

        ws.Cell(1, 1).Value = "Submission checks (Does it add up?)";
        ws.Row(1).Style.Font.Bold = true;

        ws.Cell(2, 1).Value = "Tolerance";
        ws.Cell(2, 2).Value = 0.10m;
        ws.Cell(2, 2).Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00);_-";
        ws.Cell(2, 4).Value = "Status";
        ws.Cell(2, 5).Value = "PENDING";
        ws.Cell(2, 5).Style.Font.Bold = true;

        ws.Cell(4, 1).Value = "Year";
        ws.Cell(4, 2).Value = "Description";
        ws.Cell(4, 3).Value = "OpeningCapital";
        ws.Cell(4, 4).Value = "ClosingCapital";
        ws.Cell(4, 5).Value = "Profit";
        ws.Cell(4, 6).Value = "BusinessTax";
        ws.Cell(4, 7).Value = "ProfitAfterTax";
        ws.Cell(4, 8).Value = "CapitalMovement";
        ws.Cell(4, 9).Value = "OpeningPosition";
        ws.Cell(4, 10).Value = "OpeningAccountPosition";
        ws.Cell(4, 11).Value = "CapitalDelta";

        ws.Cell(4, 13).Value = "BridgeTotal";
        ws.Cell(4, 14).Value = "Difference";
        ws.Cell(4, 15).Value = "Status";

        ws.Row(4).Style.Font.Bold = true;
        ws.Row(4).Style.Border.BottomBorder = XLBorderStyleValues.Thick;

        int firstDataRow = 5;
        int r = firstDataRow - 1;

        foreach (var row in rows)
        {
            r++;

            ws.Cell(r, 1).Value = row.YearNumber;
            ws.Cell(r, 2).Value = row.Description ?? string.Empty;

            ws.Cell(r, 3).Value = row.OpeningCapital;
            ws.Cell(r, 4).Value = row.ClosingCapital;
            ws.Cell(r, 5).Value = row.Profit;
            ws.Cell(r, 6).Value = row.BusinessTax;
            ws.Cell(r, 7).Value = row.ProfitAfterTax;
            ws.Cell(r, 8).Value = row.CapitalMovement;
            ws.Cell(r, 9).Value = row.OpeningPosition;
            ws.Cell(r, 10).Value = row.OpeningAccountPosition;
            ws.Cell(r, 11).Value = row.CapitalDelta;

            ws.Cell(r, 13).FormulaA1 = $"=G{r}+H{r}+I{r}+J{r}";
            ws.Cell(r, 14).FormulaA1 = $"=K{r}-M{r}";
            ws.Cell(r, 15).FormulaA1 =
                $"=IF(ABS(N{r})<= $B$2, \"PASS\", IF(ABS(N{r})<= ($B$2*10), \"WARN\", \"FAIL\"))";
        }

        int lastDataRow = Math.Max(firstDataRow, r);

        ws.Cell(2, 5).FormulaA1 =
            $"=IF(COUNTIF(O{firstDataRow}:O{lastDataRow},\"FAIL\")>0,\"FAIL\",IF(COUNTIF(O{firstDataRow}:O{lastDataRow},\"WARN\")>0,\"WARN\",\"PASS\"))";

        ws.Range($"O{firstDataRow}:O{lastDataRow}")
            .AddConditionalFormat()
            .WhenEquals("FAIL")
            .Font.SetFontColor(XLColor.Red)
            .Fill.SetBackgroundColor(XLColor.LightPink);

        ws.Range($"O{firstDataRow}:O{lastDataRow}")
            .AddConditionalFormat()
            .WhenEquals("WARN")
            .Font.SetFontColor(XLColor.DarkOrange)
            .Fill.SetBackgroundColor(XLColor.LightYellow);

        ws.Range($"O{firstDataRow}:O{lastDataRow}")
            .AddConditionalFormat()
            .WhenEquals("PASS")
            .Font.SetFontColor(XLColor.DarkGreen)
            .Fill.SetBackgroundColor(XLColor.LightGreen);

        ws.Cell(2, 5)
            .AddConditionalFormat()
            .WhenEquals("FAIL")
            .Font.SetFontColor(XLColor.Red)
            .Fill.SetBackgroundColor(XLColor.LightPink);

        ws.Cell(2, 5)
            .AddConditionalFormat()
            .WhenEquals("WARN")
            .Font.SetFontColor(XLColor.DarkOrange)
            .Fill.SetBackgroundColor(XLColor.LightYellow);

        ws.Cell(2, 5)
            .AddConditionalFormat()
            .WhenEquals("PASS")
            .Font.SetFontColor(XLColor.DarkGreen)
            .Fill.SetBackgroundColor(XLColor.LightGreen);

        ws.Row(2).Style.Protection.Locked = true;
        ws.Row(4).Style.Protection.Locked = true;

        ws.Columns(3, 14).Style.NumberFormat.Format = "#,##0.00;[Red](#,##0.00);_-";

        ws.Column(1).Width = 9;
        ws.Column(2).Width = 25;

        ws.Columns(3, 15).AdjustToContents();

        ws.SheetView.FreezeRows(4);
    }
}
