using System.Diagnostics;
using System.Security.Cryptography;
using System.Text.Json;
using Microsoft.Data.SqlClient;
using TCExports.Generator.CashStatement;
using TCExports.Generator.CashStatement.Model;
using TCExports.Generator.Data;
using TCExports.Generator.Excel;

var options = Options.Parse(args);
Directory.CreateDirectory(options.OutputDirectory);
var baseConnectionString = File.ReadAllText(options.ConnectionFile).Trim();
if (string.IsNullOrWhiteSpace(baseConnectionString))
    throw new InvalidOperationException("The supplied connection file is empty.");

var scenarios = new[]
{
    new Scenario("company-minimal", "tcNodeDb4-COMIPFVT1-COMIN26"),
    new Scenario("company-standard", "tcNodeDb4-COSIPFVT1-COSTD26"),
    new Scenario("sole-trader-minimal", "tcNodeDb4-STMIPFVT1-STMIN26"),
    new Scenario("sole-trader-standard", "tcNodeDb4-STSIPFVT1-STSTD26")
};
var results = new List<ScenarioResult>();

foreach (var scenario in scenarios)
{
    var stopwatch = Stopwatch.StartNew();
    var memoryBefore = GC.GetTotalMemory(forceFullCollection: true);
    try
    {
        var connection = new SqlConnectionStringBuilder(baseConnectionString)
        {
            InitialCatalog = scenario.Database
        };
        var request = new CashStatementRequestV1(
            IncludeActivePeriods: true,
            IncludeOrderBook: true,
            IncludeTaxAccruals: true,
            IncludeVatDetails: true,
            IncludeBankBalances: true,
            IncludeBalanceSheet: true,
            IncludeReconciliation: true,
            CommandTimeoutSeconds: 120);
        var generatedAt = DateTimeOffset.UtcNow;
        var build = await new CashStatementBuilder(new SqlServerCashFlowRepository()).BuildAsync(
            connection.ConnectionString,
            request,
            "cash-statement-baseline",
            generatedAt);
        if (!build.IsSuccess || build.Value is null)
            throw new InvalidOperationException(string.Join("; ", build.Violations.Select(item => $"{item.Code}: {item.Message}")));

        var bytes = new CashStatementXlsxRenderer().Render(build.Value);
        var fileName = $"{scenario.Name}__all-options__shared-model.xlsx";
        var outputPath = Path.Combine(options.OutputDirectory, fileName);
        await File.WriteAllBytesAsync(outputPath, bytes);
        stopwatch.Stop();
        results.Add(new(
            scenario.Name,
            scenario.Database,
            true,
            fileName,
            bytes.LongLength,
            Convert.ToHexString(SHA256.HashData(bytes)),
            stopwatch.ElapsedMilliseconds,
            GC.GetTotalMemory(false) - memoryBefore,
            Process.GetCurrentProcess().WorkingSet64,
            null));
        Console.WriteLine($"PASS {scenario.Name}: {bytes.Length} bytes in {stopwatch.ElapsedMilliseconds} ms");
    }
    catch (Exception ex)
    {
        stopwatch.Stop();
        results.Add(new(
            scenario.Name,
            scenario.Database,
            false,
            null,
            0,
            null,
            stopwatch.ElapsedMilliseconds,
            GC.GetTotalMemory(false) - memoryBefore,
            Process.GetCurrentProcess().WorkingSet64,
            ex.Message));
        Console.Error.WriteLine($"FAIL {scenario.Name}: {ex.Message}");
    }
}

var manifest = new Manifest(DateTimeOffset.UtcNow, results.All(result => result.Passed), results);
await File.WriteAllTextAsync(
    Path.Combine(options.OutputDirectory, "manifest.json"),
    JsonSerializer.Serialize(manifest, new JsonSerializerOptions { WriteIndented = true }));
return manifest.Passed ? 0 : 1;

sealed record Options(string ConnectionFile, string OutputDirectory)
{
    public static Options Parse(string[] args)
    {
        string Required(string name)
        {
            var index = Array.FindIndex(args, value => value.Equals(name, StringComparison.OrdinalIgnoreCase));
            if (index < 0 || index == args.Length - 1)
                throw new ArgumentException($"Required option missing: {name}");
            return Path.GetFullPath(args[index + 1]);
        }
        return new(Required("--connection-file"), Required("--output"));
    }
}

sealed record Scenario(string Name, string Database);
sealed record ScenarioResult(
    string Scenario,
    string Database,
    bool Passed,
    string? OutputFile,
    long OutputBytes,
    string? Sha256,
    long ElapsedMilliseconds,
    long ManagedMemoryDeltaBytes,
    long ProcessWorkingSetBytes,
    string? Error);
sealed record Manifest(DateTimeOffset GeneratedAtUtc, bool Passed, IReadOnlyList<ScenarioResult> Scenarios);
