using System.Diagnostics;
using System.Reflection;
using System.Text;
using TCExports.Generator.Contracts;
using ClosedXML.Excel;
using CsvHelper;
using CsvHelper.Configuration;
using System.Globalization;

namespace TCExports.Generator.Handlers;

public sealed class TestExcelHandler : IDocumentHandler
{
    public bool CanHandle(ExportPayload payload) =>
        payload.DocumentType.Equals("test", StringComparison.OrdinalIgnoreCase)
        && payload.FileType.Equals("spreadsheet", StringComparison.OrdinalIgnoreCase)
        && payload.Format.Equals("excel", StringComparison.OrdinalIgnoreCase);

    public async Task<ExportResult> HandleAsync(ExportPayload payload, CancellationToken ct)
    {
        if (!payload.Params.TryGetValue("sqlQuery", out var sqlQuery) || string.IsNullOrWhiteSpace(sqlQuery))
        {
            return new ExportResult
            {
                Status = "error",
                Code = "VALIDATION_ERROR",
                Message = "params.sqlQuery is required for documentType 'test'.",
                Details = new Dictionary<string, string[]> { ["params.sqlQuery"] = new[] { "required" } }
            };
        }

        var safeDoc = RegexSanitize(payload.DocumentType);
        var timestamp = DateTime.UtcNow.ToString("yyyyMMdd_HHmmss");
        var fileName = $"{payload.UserName}_{safeDoc}_{timestamp}.xlsx";
        var scriptPath = GetScriptPath(Path.Combine("python", "success_export.py"));

        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "python",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            psi.ArgumentList.Add(scriptPath);
            psi.ArgumentList.Add("--conn");
            psi.ArgumentList.Add(payload.SqlConnection);
            psi.ArgumentList.Add("--query");
            psi.ArgumentList.Add(sqlQuery);
            psi.ArgumentList.Add("--filename");
            psi.ArgumentList.Add("temp.csv");

            using var process = new Process { StartInfo = psi };
            process.Start();
            var stdoutTask = process.StandardOutput.ReadToEndAsync();
            var stderrTask = process.StandardError.ReadToEndAsync();
            await Task.WhenAll(process.WaitForExitAsync(ct), stdoutTask, stderrTask);

            var stdout = (await stdoutTask).Trim();
            var stderr = (await stderrTask).Trim();

            if (process.ExitCode != 0)
                return new ExportResult { Status = "error", Code = "EXECUTION_ERROR", Message = stderr };

            var parts = stdout.Split('|', 2);
            if (parts.Length != 2)
                return new ExportResult { Status = "error", Code = "EXECUTION_ERROR", Message = "Unexpected Python output." };

            var csvBase64 = parts[1];
            var csv = Encoding.UTF8.GetString(Convert.FromBase64String(csvBase64));

            // Robust CSV parsing with multiline quoted fields
            var cfg = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HasHeaderRecord = true,
                DetectColumnCountChanges = false,
                IgnoreBlankLines = false,
                BadDataFound = null, // ignore bad data callback
                TrimOptions = TrimOptions.None
            };

            using var reader = new StringReader(csv);
            using var csvReader = new CsvReader(reader, cfg);

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Export");

            int r = 1;
            // Write headers
            csvReader.Read();
            csvReader.ReadHeader();
            var headers = csvReader.HeaderRecord ?? Array.Empty<string>();
            for (int c = 0; c < headers.Length; c++)
                ws.Cell(r, c + 1).Value = headers[c];
            r++;

            // Write rows (CsvHelper handles quoted newlines)
            while (await csvReader.ReadAsync())
            {
                for (int c = 0; c < headers.Length; c++)
                {
                    var value = csvReader.GetField(c);
                    ws.Cell(r, c + 1).Value = value ?? string.Empty;
                }
                r++;
            }

            ws.Columns().AdjustToContents();

            using var ms = new MemoryStream();
            wb.SaveAs(ms);
            var xlsxBase64 = Convert.ToBase64String(ms.ToArray());

            return new ExportResult
            {
                Status = "success",
                FileName = fileName,
                FileContent = xlsxBase64
            };
        }
        catch (OperationCanceledException)
        {
            return new ExportResult { Status = "error", Code = "EXECUTION_CANCELED", Message = "Canceled." };
        }
        catch (Exception ex)
        {
            return new ExportResult { Status = "error", Code = "EXECUTION_ERROR", Message = ex.Message };
        }
    }

    private static string RegexSanitize(string? input) =>
        System.Text.RegularExpressions.Regex.Replace(input ?? "unnamed", @"[^A-Za-z0-9_\-]+", "_");

    private static string GetScriptPath(string relativePath)
    {
        var baseDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)!;
        return Path.Combine(baseDir, relativePath);
    }
}