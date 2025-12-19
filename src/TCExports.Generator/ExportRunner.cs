using System.Diagnostics;
using System.Reflection;

namespace TCExports.Generator;

public static class ExportRunner
{
    public static async Task<ExportResult> ExportDataAsync(ExportPayload payload)
    {
        // Validate first
        if (!ExportPayloadValidator.TryValidate(payload, out var errors))
        {
            return new ExportResult
            {
                Status = "error",
                Code = "VALIDATION_ERROR",
                Message = "Payload validation failed.",
                Details = errors
            };
        }

        // Construct safe filename
        var safeDoc = RegexSanitize(payload.DocumentType);
        var timestamp = DateTime.UtcNow.ToString("yyyyMMdd_HHmmss");
        var ext = payload.Format.Equals("csv", StringComparison.OrdinalIgnoreCase) ? "csv" : "csv";
        var fileName = $"{payload.UserName}_{safeDoc}_{timestamp}.{ext}";
        var scriptPath = GetScriptPath(Path.Combine("python", "success_export.py"));

        try
        {
            // Expect sqlQuery param for 'test' document type
            string? sqlQuery = null;
            if (payload.DocumentType.Equals("test", StringComparison.OrdinalIgnoreCase))
                payload.Params.TryGetValue("sqlQuery", out sqlQuery);

            if (string.IsNullOrWhiteSpace(sqlQuery))
                return new ExportResult
                {
                    Status = "error",
                    Code = "VALIDATION_ERROR",
                    Message = "params.sqlQuery is required for documentType 'test'.",
                    Details = new Dictionary<string, string[]> { ["params.sqlQuery"] = new[] { "required" } }
                };

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

            // Pass the dynamic query from the payload
            psi.ArgumentList.Add("--query");
            psi.ArgumentList.Add(sqlQuery);

            psi.ArgumentList.Add("--filename");
            psi.ArgumentList.Add(fileName);

            using var process = new Process { StartInfo = psi };
            process.Start();

            var stdout = await process.StandardOutput.ReadToEndAsync();
            var stderr = await process.StandardError.ReadToEndAsync();
            await process.WaitForExitAsync();

            if (process.ExitCode != 0)
                return new ExportResult { Status = "error", Code = "EXECUTION_ERROR", Message = stderr };

            var parts = stdout.Split('|', 2);
            return new ExportResult
            {
                Status = "success",
                FileName = parts[0],
                FileContent = parts.Length > 1 ? parts[1] : null
            };
        }
        catch (Exception ex)
        {
            return new ExportResult { Status = "error", Code = "EXECUTION_ERROR", Message = ex.Message };
        }
    }

    private static string RegexSanitize(string input) =>
        System.Text.RegularExpressions.Regex.Replace(input ?? "unnamed", @"[^A-Za-z0-9_\-]+", "_");

    private static string GetScriptPath(string relativePath)
    {
        var baseDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)!;
        return Path.Combine(baseDir, relativePath);
    }
}