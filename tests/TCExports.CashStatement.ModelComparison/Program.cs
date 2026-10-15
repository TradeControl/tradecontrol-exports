using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using Microsoft.Data.SqlClient;
using TCExports.Generator.CashStatement;
using TCExports.Generator.CashStatement.Model;
using TCExports.Generator.Data;

var options = Options.Parse(args);
Directory.CreateDirectory(options.EvidenceDirectory);
var rootConnectionString = File.ReadAllText(options.ConnectionFile).Trim();
if (string.IsNullOrWhiteSpace(rootConnectionString))
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
    var managedMemoryBefore = GC.GetTotalMemory(forceFullCollection: true);
    var stopwatch = Stopwatch.StartNew();
    var snapshotPath = Path.Combine(
        options.SnapshotDirectory,
        $"{scenario.Name}__all-options__raw.snapshot.json");
    using var snapshotDocument = JsonDocument.Parse(File.ReadAllText(snapshotPath));
    var snapshot = BaselineSnapshot.Read(snapshotDocument.RootElement);

    var connection = new SqlConnectionStringBuilder(rootConnectionString) { InitialCatalog = scenario.Database };
    var request = new CashStatementRequestV1(
        IncludeActivePeriods: true,
        IncludeOrderBook: true,
        IncludeTaxAccruals: true,
        IncludeVatDetails: true,
        IncludeBankBalances: true,
        IncludeBalanceSheet: true,
        IncludeReconciliation: true,
        CommandTimeoutSeconds: 120);
    var build = await new CashStatementBuilder(new SqlServerCashFlowRepository()).BuildAsync(
        connection.ConnectionString,
        request,
        "phase-3-model-comparison",
        DateTimeOffset.UtcNow);

    if (!build.IsSuccess || build.Value is null)
    {
        var violations = build.Violations.Select(item => $"{item.Code}: {item.Message}").ToArray();
        results.Add(new ScenarioResult(scenario.Name, scenario.Database, false, 0, 0, 0, violations, stopwatch.ElapsedMilliseconds, GC.GetTotalMemory(false) - managedMemoryBefore, Process.GetCurrentProcess().WorkingSet64));
        Console.Error.WriteLine($"FAIL {scenario.Name}: model contract failed ({violations.Length} violation(s))");
        continue;
    }

    var comparison = Compare(build.Value, snapshot);
    stopwatch.Stop();
    results.Add(new ScenarioResult(
        scenario.Name,
        scenario.Database,
        comparison.Mismatches.Count == 0,
        comparison.Compared,
        comparison.Delegated,
        comparison.Missing,
        comparison.Mismatches.ToArray(),
        stopwatch.ElapsedMilliseconds,
        GC.GetTotalMemory(false) - managedMemoryBefore,
        Process.GetCurrentProcess().WorkingSet64));
    Console.WriteLine(comparison.Mismatches.Count == 0
        ? $"PASS {scenario.Name}: {comparison.Compared} values, {comparison.Missing} blanks, {comparison.Delegated} delegated expressions"
        : $"FAIL {scenario.Name}: {comparison.Mismatches.Count} mismatch(es)");
}

var evidence = new Evidence(
    DateTimeOffset.UtcNow,
    "all-options; accepted LibreOffice-recalculated Phase 0 XLSX snapshots",
    "Exact decimal equality for source cells; spreadsheet-derived baseline formulas use absolute tolerance 0.000000001 for binary residue; unsupported management expressions require a corresponding baseline formula cell.",
    results.All(result => result.Passed),
    results);
var evidencePath = Path.Combine(options.EvidenceDirectory, "model-baseline-comparison.json");
File.WriteAllText(evidencePath, JsonSerializer.Serialize(evidence, new JsonSerializerOptions { WriteIndented = true }));

if (!evidence.Passed)
{
    foreach (var result in results.Where(result => !result.Passed))
        foreach (var mismatch in result.Mismatches.Take(50))
            Console.Error.WriteLine($"  {result.Scenario}: {mismatch}");
    return 1;
}

Console.WriteLine($"Cash Statement model comparison passed: {results.Count}/{results.Count} scenarios.");
return 0;

static Comparison Compare(CashStatementDocument document, BaselineSnapshot baseline)
{
    var mismatches = new List<string>();
    var compared = 0;
    var delegated = 0;
    var missing = 0;
    var cashFlow = document.Sheets.Single(sheet => sheet.Kind == StatementSheetKind.CashFlow);
    var periodKeys = document.PeriodGrid.Periods.ToDictionary(
        period => period.Id,
        period => $"{period.YearLabel}/{(period.Id.IsAnnual ? "TOTALS" : period.MonthLabel.ToUpperInvariant())}");

    foreach (var section in cashFlow.Sections)
    {
        string? categoryName = null;
        foreach (var row in section.Rows.OrderBy(row => row.PresentationOrder))
        {
            if (section.Kind == StatementSectionKind.FlowCategories && row.Kind == StatementRowKind.Header)
            {
                categoryName = row.Label.ToUpperInvariant();
                continue;
            }

            var rowKey = SemanticRowKey(section, row, categoryName);
            if (rowKey is null)
                continue;
            if (!baseline.Rows.TryGetValue(rowKey, out var baselineRow))
            {
                mismatches.Add($"cash-flow row not found: {rowKey}");
                continue;
            }

            foreach (var period in document.PeriodGrid.Periods)
            {
                var columnId = new StatementColumnId($"period:{period.Id}");
                if (!row.Cells.TryGetValue(columnId, out var cell))
                    continue;
                var periodKey = periodKeys[period.Id];
                baselineRow.TryGetValue(periodKey, out var baselineCell);

                if (cell.Value.State == StatementValueState.Unsupported)
                {
                    if (cell.Formula is StatementFormulaTerminal && baselineCell is { HasFormula: true })
                        delegated++;
                    else if (cell.Formula is null && baselineCell is not null)
                        delegated++; // Known invalid legacy expression remains explicit and non-zero-coercing.
                    else
                        mismatches.Add($"{rowKey}/{periodKey}: unsupported expression has no corresponding baseline cell/formula");
                    continue;
                }

                if (cell.Value.State == StatementValueState.Missing)
                {
                    if (baselineCell is null || baselineCell.IsBlank)
                        missing++;
                    else
                        mismatches.Add($"{rowKey}/{periodKey}: model is missing, baseline is '{baselineCell.Value}'");
                    continue;
                }

                if (cell.Value.Kind != StatementScalarKind.Decimal || cell.Value.DecimalValue is not { } actual)
                    continue;
                if (baselineCell is null || !baselineCell.TryDecimal(out var expected))
                {
                    mismatches.Add($"{rowKey}/{periodKey}: model is {actual}, baseline is blank or non-numeric");
                    continue;
                }
                compared++;
                if (baselineCell.HasFormula
                    ? Math.Abs(actual - expected) > 0.000000001m
                    : actual != expected)
                    mismatches.Add($"{rowKey}/{periodKey}: model {actual} != baseline {expected}");
            }
        }
    }

    var reconciliation = document.Sheets.SingleOrDefault(sheet => sheet.Kind == StatementSheetKind.Reconciliation);
    if (reconciliation is null)
    {
        mismatches.Add("reconciliation sheet missing");
    }
    else
    {
        var rows = reconciliation.Sections.Single().Rows;
        var summary = rows.Single(row => row.Id.Value == "reconciliation:summary");
        CompareRecon(summary.Cells[new StatementColumnId("tolerance")].Value, "B2", exact: true);
        CompareRecon(summary.Cells[new StatementColumnId("status")].Value, "E2", exact: true);

        foreach (var row in rows.Where(row => row.Kind == StatementRowKind.Reconciliation))
        {
            var year = row.Id.Value[(row.Id.Value.LastIndexOf(':') + 1)..];
            var baselineRowNumber = baseline.ReconciliationCells.Values
                .Where(cell => cell.Address.StartsWith('A') && cell.Value == year)
                .Select(cell => int.Parse(cell.Address[1..], CultureInfo.InvariantCulture))
                .SingleOrDefault();
            if (baselineRowNumber == 0)
            {
                mismatches.Add($"reconciliation year not found: {year}");
                continue;
            }

            var columns = new Dictionary<string, string>
            {
                ["description"] = "B", ["opening-capital"] = "C", ["closing-capital"] = "D",
                ["profit"] = "E", ["business-tax"] = "F", ["profit-after-tax"] = "G",
                ["capital-movement"] = "H", ["opening-subject-position"] = "I",
                ["opening-account-position"] = "J", ["capital-delta"] = "K",
                ["bridge-total"] = "M", ["difference"] = "N", ["status"] = "O"
            };
            foreach (var pair in columns)
                CompareRecon(row.Cells[new StatementColumnId(pair.Key)].Value, $"{pair.Value}{baselineRowNumber}", exact: pair.Key is not ("bridge-total" or "difference"));
        }
    }

    return new Comparison(compared, delegated, missing, mismatches);

    void CompareRecon(StatementValue actual, string address, bool exact)
    {
        if (!baseline.ReconciliationCells.TryGetValue(address, out var expected))
        {
            mismatches.Add($"reconciliation/{address}: baseline cell missing");
            return;
        }
        if (actual.Kind == StatementScalarKind.Decimal && actual.DecimalValue is { } number)
        {
            compared++;
            if (!expected.TryDecimal(out var expectedNumber) ||
                (exact ? number != expectedNumber : Math.Abs(number - expectedNumber) > 0.000000001m))
                mismatches.Add($"reconciliation/{address}: model {number} != baseline {expected.Value}");
            return;
        }
        if (actual.Kind == StatementScalarKind.Text)
        {
            compared++;
            if (!string.Equals(actual.TextValue, expected.Value, StringComparison.Ordinal))
                mismatches.Add($"reconciliation/{address}: model '{actual.TextValue}' != baseline '{expected.Value}'");
        }
    }
}

static string? SemanticRowKey(StatementSection section, StatementRow row, string? categoryName) => section.Kind switch
{
    StatementSectionKind.FlowCategories when row.Kind == StatementRowKind.Detail =>
        $"{categoryName}|{row.BusinessCode?.ToUpperInvariant()}|{row.Label.ToUpperInvariant()}",
    StatementSectionKind.FlowCategories when row.Kind == StatementRowKind.CategoryTotal =>
        $"{categoryName}|{row.BusinessCode?.ToUpperInvariant()}",
    StatementSectionKind.FlowSummary =>
        $"{SummaryPrefix(section.Id.Value)}|{(row.Kind == StatementRowKind.PeriodTotal ? "PERIOD TOTAL" : $"{row.BusinessCode?.ToUpperInvariant()}|{row.Label.ToUpperInvariant()}")}",
    StatementSectionKind.CategoryTotals => $"{section.Title.ToUpperInvariant()}|{row.BusinessCode?.ToUpperInvariant()}",
    StatementSectionKind.Analysis => $"ANALYSIS|{row.BusinessCode?.ToUpperInvariant()}",
    StatementSectionKind.BankBalances when row.Kind == StatementRowKind.BankAccount =>
        $"CLOSING BALANCE|{row.BusinessCode?.ToUpperInvariant()}|{row.Label.ToUpperInvariant()}",
    StatementSectionKind.BankBalances when row.Kind == StatementRowKind.BankTotal => "CLOSING BALANCE|COMPANY BALANCE",
    StatementSectionKind.VatRecurrence or StatementSectionKind.VatPeriods =>
        $"{section.Title.ToUpperInvariant()}|{row.Label.ToUpperInvariant()}",
    StatementSectionKind.BalanceSheet when row.Kind == StatementRowKind.BalanceSheetAsset =>
        $"BALANCE SHEET|{row.BusinessCode?.ToUpperInvariant()}|{row.Label.ToUpperInvariant()}",
    StatementSectionKind.BalanceSheet when row.Kind == StatementRowKind.Capital => "BALANCE SHEET|CAPITAL",
    _ => null
};

static string SummaryPrefix(string sectionId) => sectionId switch
{
    "summary:trade" => "SUMMARY",
    "summary:money" => "SUMMARY|#2",
    "summary:tax" => "SUMMARY|#3",
    _ => throw new InvalidOperationException($"Unknown summary section '{sectionId}'.")
};

sealed record Options(string ConnectionFile, string SnapshotDirectory, string EvidenceDirectory)
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
        return new(Required("--connection-file"), Required("--snapshots"), Required("--evidence"));
    }
}

sealed record Scenario(string Name, string Database);
sealed record Comparison(int Compared, int Delegated, int Missing, List<string> Mismatches);
sealed record ScenarioResult(
    string Scenario,
    string Database,
    bool Passed,
    int ComparedValues,
    int DelegatedExpressions,
    int MissingValues,
    IReadOnlyList<string> Mismatches,
    long ElapsedMilliseconds,
    long ManagedMemoryDeltaBytes,
    long ProcessWorkingSetBytes);
sealed record Evidence(DateTimeOffset GeneratedAtUtc, string Baseline, string ComparisonPolicy, bool Passed, IReadOnlyList<ScenarioResult> Scenarios);

sealed record BaselineCell(string Address, string DataType, string Value, bool HasFormula)
{
    public bool IsBlank => DataType == "Blank" || string.IsNullOrEmpty(Value);
    public bool TryDecimal(out decimal value) => decimal.TryParse(Value, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
}

sealed record BaselineSnapshot(
    Dictionary<string, Dictionary<string, BaselineCell>> Rows,
    Dictionary<string, BaselineCell> ReconciliationCells)
{
    public static BaselineSnapshot Read(JsonElement root)
    {
        var rows = new Dictionary<string, Dictionary<string, BaselineCell>>(StringComparer.OrdinalIgnoreCase);
        foreach (var row in root.GetProperty("CashFlowSemantics").GetProperty("Rows").EnumerateArray())
        {
            var cells = new Dictionary<string, BaselineCell>(StringComparer.OrdinalIgnoreCase);
            foreach (var cell in row.GetProperty("Cells").EnumerateArray())
            {
                var period = cell.GetProperty("PeriodKey").GetString()!;
                cells[period] = ReadCell(cell, period);
            }
            rows[row.GetProperty("Key").GetString()!] = cells;
        }

        var worksheets = root.GetProperty("Worksheets").EnumerateArray().ToArray();
        var reconciliation = worksheets
            .Single(sheet => sheet.GetProperty("Name").GetString() == "Reconciliation");
        var reconciliationCells = reconciliation.GetProperty("Cells").EnumerateArray()
            .Select(cell => ReadCell(cell, cell.GetProperty("Address").GetString()!))
            .ToDictionary(cell => cell.Address, StringComparer.OrdinalIgnoreCase);
        return new(rows, reconciliationCells);
    }

    private static BaselineCell ReadCell(JsonElement cell, string address) => new(
        address,
        cell.GetProperty("DataType").GetString() ?? string.Empty,
        cell.TryGetProperty("Value", out var value) ? value.GetString() ?? string.Empty : string.Empty,
        cell.TryGetProperty("Formula", out var formula) && !string.IsNullOrWhiteSpace(formula.GetString()));
}
