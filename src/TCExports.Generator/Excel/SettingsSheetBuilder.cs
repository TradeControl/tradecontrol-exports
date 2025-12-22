using ClosedXML.Excel;
using TCExports.Generator.Contracts;

namespace TCExports.Generator.Excel;

public static class SettingsSheetBuilder
{
    // Adds/updates a worksheet named "Settings" with payload metadata and normalized params
    public static void AddSettingsWorksheet(XLWorkbook wb, ExportPayload payload)
    {
        var ws = wb.Worksheets.FirstOrDefault(w => w.Name == "Settings") ?? wb.Worksheets.Add("Settings");
        ws.Clear();

        int r = 1;

        ws.Cell(r, 1).Value = "Document Type"; ws.Cell(r, 2).Value = payload.DocumentType; r++;
        ws.Cell(r, 1).Value = "Format"; ws.Cell(r, 2).Value = payload.Format; r++;
        ws.Cell(r, 1).Value = "User"; ws.Cell(r, 2).Value = payload.UserName; r++;
        ws.Cell(r, 1).Value = "SqlConnection"; ws.Cell(r, 2).Value = payload.SqlConnection; r++;

        r++; ws.Cell(r, 1).Value = "Parameters"; ws.Row(r).Style.Font.Bold = true; r++;

        if (payload.Params is { Count: > 0 })
        {
            foreach (var kvp in payload.Params.OrderBy(k => k.Key, StringComparer.OrdinalIgnoreCase))
            {
                ws.Cell(r, 1).Value = kvp.Key;
                ws.Cell(r, 2).Value = kvp.Value;
                r++;
            }
        }
        else
        {
            ws.Cell(r, 1).Value = "(none)";
            r++;
        }

        // Formatting
        ws.Column(1).Width = 24;
        ws.Column(2).Width = 80;
        ws.Range(1, 1, 1, 2).Style.Font.Bold = true;
        ws.Range(1, 1, r - 1, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        ws.Range(1, 1, r - 1, 2).Style.Border.InsideBorder = XLBorderStyleValues.Hair;
        ws.SheetView.FreezeRows(1);
    }
}