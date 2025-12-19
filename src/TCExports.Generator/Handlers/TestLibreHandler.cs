using System.Diagnostics;
using System.Reflection;
using TCExports.Generator.Contracts;

namespace TCExports.Generator.Handlers;

public sealed class TestLibreHandler : IDocumentHandler
{
    public bool CanHandle(ExportPayload payload) =>
        payload.DocumentType.Equals("test", StringComparison.OrdinalIgnoreCase)
        && payload.Format.Equals("libre", StringComparison.OrdinalIgnoreCase);

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
        var ts = DateTime.UtcNow.ToString("yyyyMMdd_HHmmss");
        var fileName = $"{payload.UserName}_{safeDoc}_{ts}.ods";
        var scriptPath = GetScriptPath(Path.Combine("python", "exporters", "test_ods_export.py"));

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
            psi.ArgumentList.Add(fileName);

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

            return new ExportResult
            {
                Status = "success",
                FileName = parts[0],
                FileContent = parts[1]
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