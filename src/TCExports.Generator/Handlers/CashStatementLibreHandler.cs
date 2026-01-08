using System.Diagnostics;
using System.Text;
using System.Text.Json;
using System.Reflection;

using TCExports.Generator.Contracts;

namespace TCExports.Generator.Handlers;

public sealed class CashStatementLibreHandler : IDocumentHandler
{
    public bool CanHandle(ExportPayload payload) =>
        payload.DocumentType.Equals("cashflow", StringComparison.OrdinalIgnoreCase)
        && payload.Format.Equals("libre", StringComparison.OrdinalIgnoreCase);

    public async Task<ExportResult> HandleAsync(ExportPayload payload, CancellationToken ct)
    {
        // Serialize payload to a temp file to avoid stdin pipe issues
        var tempDir = Path.GetTempPath();
        var tempFile = Path.Combine(tempDir, $"cashflow_payload_{Guid.NewGuid():N}.json");
        var json = JsonSerializer.Serialize(payload);
        var utf8NoBom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        await File.WriteAllTextAsync(tempFile, json, utf8NoBom, ct);

        var scriptPath = Path.Combine(AppContext.BaseDirectory, "python", "exporters", "cash_statement_ods.py");
        var pythonRoot = Path.Combine(AppContext.BaseDirectory, "python");

        var psi = new ProcessStartInfo
        {
            FileName = "python",
            // -u for unbuffered stdout/stderr; pass payload file path as argument
            Arguments = $"-u \"{scriptPath}\" \"{tempFile}\"",
            RedirectStandardInput = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
            WorkingDirectory = pythonRoot // ensure 'data' and 'i18n' resolve
        };

        try
        {
            psi.EnvironmentVariables["PYTHONPATH"] = pythonRoot;

            using var proc = new Process { StartInfo = psi };
            proc.Start();

            var stdoutTask = proc.StandardOutput.ReadToEndAsync();
            var stderrTask = proc.StandardError.ReadToEndAsync();

            await proc.WaitForExitAsync(ct);

            var stdout = await stdoutTask;
            var stderr = await stderrTask;

            if (proc.ExitCode != 0)
            {
                return new ExportResult
                {
                    Status = "error",
                    Code = "EXECUTION_ERROR",
                    Message = "Libre generator failed.",
                    Details = new Dictionary<string, string[]>
                    {
                        ["stderr"] = new[] { stderr }
                    }
                };
            }

            var sep = stdout.IndexOf('|');
            if (sep <= 0)
            {
                return new ExportResult
                {
                    Status = "error",
                    Code = "EXECUTION_ERROR",
                    Message = "Unexpected output from Libre generator.",
                    Details = new Dictionary<string, string[]>
                    {
                        ["stdout"] = new[] { stdout }
                    }
                };
            }

            var fileName = System.Text.RegularExpressions.Regex.Replace(stdout[..sep].Trim(), @"[^A-Za-z0-9_\-\.]+", "_");
            var base64 = stdout[(sep + 1)..].Trim();

            return new ExportResult
            {
                Status = "success",
                FileName = fileName,
                FileContent = base64
            };
        }
        finally
        {
            // Cleanup temp file
            try { File.Delete(tempFile); } catch { /* ignore */ }
        }
    }
}