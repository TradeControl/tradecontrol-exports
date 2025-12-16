using System.Diagnostics;
using System.Reflection;

namespace TCExports.Generator;

public static class ExportRunner
{
    public static async Task<ExportResult> ExportDataAsync(ExportPayload payload)
    {
        // Construct safe filename
        var safeDoc = RegexSanitize(payload.DocumentType);
        var timestamp = DateTime.UtcNow.ToString("yyyyMMdd_HHmmss");
        var fileName = $"{payload.UserName}_{safeDoc}_{timestamp}.{payload.FileType}";
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
            psi.ArgumentList.Add("SELECT TOP 100 LogCode, LoggedOn, EventTypeCode, EventMessage, InsertedBy FROM App.tbEventLog ORDER BY LogCode DESC;"); 
            psi.ArgumentList.Add("--filename"); 
            psi.ArgumentList.Add(fileName);

            using var process = new Process { StartInfo = psi };
            process.Start();

            var stdout = await process.StandardOutput.ReadToEndAsync();
            var stderr = await process.StandardError.ReadToEndAsync();
            await process.WaitForExitAsync();

            if (process.ExitCode != 0)
                return new ExportResult { Status = "error", Message = stderr };

            // stdout format: "filename|base64content"
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
            return new ExportResult { Status = "error", Message = ex.Message };
        }
    }

    private static string RegexSanitize(string input) =>
        System.Text.RegularExpressions.Regex.Replace(input ?? "unnamed", @"[^A-Za-z0-9_\-]+", "_");

    private static string GetScriptPath(string relativePath)
    {
        // Get the folder where the DLL is running
        var baseDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)!;
        return Path.Combine(baseDir, relativePath);
    }

}
