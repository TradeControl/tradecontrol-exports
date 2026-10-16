using System.Diagnostics;
using System.Security.Cryptography;
using System.Text.Json;
using Microsoft.Data.SqlClient;
using TCExports.Generator.CashStatement;
using TCExports.Generator.CashStatement.Model;
using TCExports.Generator.Data;
using TCExports.Generator.Excel;
using TCExports.Generator.Ods;

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
    try
    {
        var connection = new SqlConnectionStringBuilder(baseConnectionString) { InitialCatalog = scenario.Database };
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
            "cash-statement-tri-format-acceptance",
            generatedAt);
        if (!build.IsSuccess || build.Value is null)
            throw new InvalidOperationException(string.Join("; ", build.Violations.Select(item => $"{item.Code}: {item.Message}")));

        // Both renderers consume this exact immutable document instance. This prevents a
        // database read or generation timestamp from differing between the two formats.
        var xlsx = new CashStatementXlsxRenderer().Render(build.Value);
        var ods = await new CashStatementOdsRenderer().RenderAsync(build.Value);
        var xlsxName = $"{scenario.Name}__all-options__shared-model.xlsx";
        var odsName = $"{scenario.Name}__all-options__shared-model.ods";
        await File.WriteAllBytesAsync(Path.Combine(options.OutputDirectory, xlsxName), xlsx);
        await File.WriteAllBytesAsync(Path.Combine(options.OutputDirectory, odsName), ods);

        stopwatch.Stop();
        results.Add(new ScenarioResult(
            scenario.Name,
            scenario.Database,
            true,
            generatedAt,
            xlsxName,
            xlsx.LongLength,
            Convert.ToHexString(SHA256.HashData(xlsx)),
            odsName,
            ods.LongLength,
            Convert.ToHexString(SHA256.HashData(ods)),
            stopwatch.ElapsedMilliseconds,
            null));
        Console.WriteLine($"PASS {scenario.Name}: XLSX {xlsx.Length} bytes; ODS {ods.Length} bytes; one shared document");
    }
    catch (Exception exception)
    {
        stopwatch.Stop();
        results.Add(new ScenarioResult(
            scenario.Name, scenario.Database, false, null, null, 0, null, null, 0, null,
            stopwatch.ElapsedMilliseconds, exception.ToString()));
        Console.Error.WriteLine($"FAIL {scenario.Name}: {exception.Message}");
    }
}

var manifest = new Manifest(
    DateTimeOffset.UtcNow,
    "Every XLSX/ODS pair was rendered from the same immutable CashStatementDocument instance with all supported request options enabled.",
    results.All(result => result.Passed),
    results);
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

        return new Options(Required("--connection-file"), Required("--output"));
    }
}

sealed record Scenario(string Name, string Database);
sealed record ScenarioResult(
    string Scenario,
    string Database,
    bool Passed,
    DateTimeOffset? GeneratedAt,
    string? XlsxFile,
    long XlsxBytes,
    string? XlsxSha256,
    string? OdsFile,
    long OdsBytes,
    string? OdsSha256,
    long ElapsedMilliseconds,
    string? Error);
sealed record Manifest(DateTimeOffset GeneratedAtUtc, string GenerationContract, bool Passed, IReadOnlyList<ScenarioResult> Scenarios);
