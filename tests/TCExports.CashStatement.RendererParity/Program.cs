using System.Globalization;
using System.Text.Json;
using ClosedXML.Excel;
using Microsoft.Data.SqlClient;
using TCExports.Generator.CashStatement;
using TCExports.Generator.CashStatement.Model;
using TCExports.Generator.Data;
using TCExports.Generator.Excel;
using TCExports.Generator.Ods;

const decimal FormulaTolerance = 0.000000001m;
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
    var mismatches = new List<string>();
    var connection = new SqlConnectionStringBuilder(rootConnectionString) { InitialCatalog = scenario.Database };
    var request = new CashStatementRequestV1(true, true, true, true, true, true, true, 120);
    var build = await new CashStatementBuilder(new SqlServerCashFlowRepository()).BuildAsync(
        connection.ConnectionString, request, "cash-statement-renderer-parity", DateTimeOffset.UtcNow);
    if (!build.IsSuccess || build.Value is null)
    {
        mismatches.AddRange(build.Violations.Select(item => $"{item.Code}: {item.Message}"));
        results.Add(new ScenarioResult(scenario.Name, false, 0, 0, 0, 0, 0, mismatches.Count, mismatches));
        continue;
    }

    var document = build.Value;
    var xlsxLayout = new CashStatementXlsxLayoutAllocator().Allocate(document, CashStatementLimits.Default);
    var odsLayout = new CashStatementOdsLayoutAllocator().Allocate(document, CashStatementLimits.Default);
    var modelCells = document.Sheets
        .SelectMany(sheet => sheet.Sections.SelectMany(section => section.Rows.SelectMany(row =>
            row.Cells.Select(pair => new KeyValuePair<StatementCellAddress, StatementCell>(
                new StatementCellAddress(sheet.Id, row.Id, pair.Key), pair.Value)))))
        .ToDictionary(pair => pair.Key, pair => pair.Value);
    var invalidExpressionCells = document.Sheets
        .SelectMany(sheet => sheet.Sections.SelectMany(section => section.Rows
            .Where(row => row.Kind == StatementRowKind.Expression)
            .SelectMany(row => row.Cells
                .Where(pair => pair.Value.Value.State == StatementValueState.Unsupported && pair.Value.Formula is null)
                .Select(pair => new StatementCellAddress(sheet.Id, row.Id, pair.Key)))))
        .ToHashSet();

    if (!xlsxLayout.Cells.Keys.ToHashSet().SetEquals(odsLayout.Cells.Keys))
        mismatches.Add("XLSX and ODS semantic layout keys differ.");
    if (!modelCells.Keys.ToHashSet().SetEquals(xlsxLayout.Cells.Keys))
        mismatches.Add("The model and renderer semantic layout keys differ.");

    var xlsxPath = Path.Combine(options.XlsxDirectory, $"{scenario.Name}__all-options__shared-model.xlsx");
    var odsPath = Path.Combine(options.OdsAsXlsxDirectory, $"{scenario.Name}__all-options__shared-model.xlsx");
    var originalPath = Path.Combine(options.OriginalDirectory, $"{scenario.Name}__all-options__raw.xlsx");
    using var original = new XLWorkbook(originalPath);
    using var xlsx = new XLWorkbook(xlsxPath);
    using var ods = new XLWorkbook(odsPath);
    var compared = 0;
    var modelCompared = 0;
    var formulaResults = 0;
    var originalCompared = 0;
    var approvedDifferences = 0;

    foreach (var address in xlsxLayout.Cells.Keys.OrderBy(item => item.ToString(), StringComparer.Ordinal))
    {
        if (!odsLayout.Cells.TryGetValue(address, out var odsPhysical) ||
            !xlsxLayout.Cells.TryGetValue(address, out var xlsxPhysical) ||
            !modelCells.TryGetValue(address, out var expected))
            continue;

        var xlsxCell = xlsx.Worksheet(xlsxPhysical.SheetName).Cell(xlsxPhysical.RowNumber, xlsxPhysical.ColumnNumber);
        var odsCell = ods.Worksheet(odsPhysical.SheetName).Cell(odsPhysical.RowNumber, odsPhysical.ColumnNumber);
        var originalCell = original.Worksheet(xlsxPhysical.SheetName).Cell(xlsxPhysical.RowNumber, xlsxPhysical.ColumnNumber);
        var xlsxValue = Read(xlsxCell);
        var odsValue = Read(odsCell);
        CompareValues(xlsxValue, odsValue, address.ToString(), expected.Formula is not null, mismatches);
        compared++;

        var expectsFormula = address.RowId.Value == "balance-sheet:capital" ||
                             expected.Formula is not null && !ContainsEmptyAggregate(expected.Formula);
        if (xlsxCell.HasFormula != expectsFormula)
            mismatches.Add($"{address}: XLSX formula presence {xlsxCell.HasFormula} != expected {expectsFormula}");
        if (odsCell.HasFormula != expectsFormula)
            mismatches.Add($"{address}: ODS formula presence {odsCell.HasFormula} != expected {expectsFormula}");

        if (invalidExpressionCells.Contains(address))
        {
            approvedDifferences++;
        }
        else
        {
            var originalValue = Read(originalCell);
            var calculated = expectsFormula || originalCell.HasFormula;
            CompareValues(originalValue, xlsxValue, $"{address}: Original/XLSX", calculated, mismatches);
            CompareValues(originalValue, odsValue, $"{address}: Original/ODS", calculated, mismatches);
            if (originalCell.HasFormula != expectsFormula)
                mismatches.Add($"{address}: Original formula presence {originalCell.HasFormula} != expected {expectsFormula}");
            originalCompared++;
        }

        if (expected.Value.State == StatementValueState.Present)
        {
            CompareModel(expected.Value, xlsxValue, $"{address}: XLSX", expectsFormula, mismatches);
            CompareModel(expected.Value, odsValue, $"{address}: ODS", expectsFormula, mismatches);
            modelCompared++;
        }
        else if (expected.Value.State == StatementValueState.Missing && !expectsFormula)
        {
            if (xlsxValue.Kind != ActualKind.Missing)
                mismatches.Add($"{address}: XLSX is not missing");
            if (odsValue.Kind != ActualKind.Missing)
                mismatches.Add($"{address}: ODS is not missing");
            modelCompared++;
        }

        if (expected.Formula is StatementFormulaTerminal)
        {
            formulaResults++;
            if (xlsxValue.Kind == ActualKind.Error || odsValue.Kind == ActualKind.Error)
                mismatches.Add($"{address}: management expression contains a spreadsheet error");
        }
    }

    var result = new ScenarioResult(
        scenario.Name, mismatches.Count == 0, originalCompared, compared, modelCompared, formulaResults, approvedDifferences,
        mismatches.Count, mismatches.Take(200).ToArray());
    results.Add(result);
    Console.WriteLine(result.Passed
        ? $"PASS {scenario.Name}: {originalCompared} tri-format cells, {compared} XLSX/ODS cells, {modelCompared} model values, {formulaResults} expression results, {approvedDifferences} approved invalid-expression cells"
        : $"FAIL {scenario.Name}: {mismatches.Count} mismatch(es)");
}

var evidence = new Evidence(
    DateTimeOffset.UtcNow,
    "All XLSX and ODS values are read after LibreOffice recalculation. Source/model values are exact; formula results use absolute tolerance 0.000000001.",
    results.All(result => result.Passed),
    results);
await File.WriteAllTextAsync(
    Path.Combine(options.EvidenceDirectory, "renderer-parity.json"),
    JsonSerializer.Serialize(evidence, new JsonSerializerOptions { WriteIndented = true }));
return evidence.Passed ? 0 : 1;

static ActualValue Read(IXLCell cell)
{
    if (cell.IsEmpty() && !cell.HasFormula)
        return new ActualValue(ActualKind.Missing, null);
    return cell.DataType switch
    {
        XLDataType.Number => new ActualValue(ActualKind.Decimal,
            decimal.Parse(cell.GetDouble().ToString("R", CultureInfo.InvariantCulture), CultureInfo.InvariantCulture)),
        XLDataType.Boolean => new ActualValue(ActualKind.Boolean, cell.GetBoolean()),
        XLDataType.DateTime => new ActualValue(ActualKind.Date, DateOnly.FromDateTime(cell.GetDateTime())),
        XLDataType.Error => new ActualValue(ActualKind.Error, cell.Value.ToString(CultureInfo.InvariantCulture)),
        _ => new ActualValue(ActualKind.Text, cell.Value.ToString(CultureInfo.InvariantCulture))
    };
}

static void CompareValues(ActualValue left, ActualValue right, string path, bool formula, List<string> mismatches)
{
    if (left.Kind != right.Kind)
    {
        mismatches.Add($"{path}: XLSX {left.Kind} != ODS {right.Kind}");
        return;
    }
    if (left.Kind == ActualKind.Decimal)
    {
        var difference = Math.Abs((decimal)left.Value! - (decimal)right.Value!);
        if (formula ? difference > FormulaTolerance : difference != 0)
            mismatches.Add($"{path}: XLSX {left.Value} != ODS {right.Value}");
        return;
    }
    if (!Equals(left.Value, right.Value))
        mismatches.Add($"{path}: XLSX '{left.Value}' != ODS '{right.Value}'");
}

static void CompareModel(StatementValue expected, ActualValue actual, string path, bool formula, List<string> mismatches)
{
    var expectedKind = expected.Kind switch
    {
        StatementScalarKind.Decimal => ActualKind.Decimal,
        StatementScalarKind.Text => ActualKind.Text,
        StatementScalarKind.Date => ActualKind.Date,
        StatementScalarKind.Boolean => ActualKind.Boolean,
        _ => throw new ArgumentOutOfRangeException()
    };
    if (actual.Kind != expectedKind)
    {
        mismatches.Add($"{path} {actual.Kind} != model {expectedKind}");
        return;
    }
    object? expectedValue = expected.Kind switch
    {
        StatementScalarKind.Decimal => expected.DecimalValue,
        StatementScalarKind.Text => expected.TextValue,
        StatementScalarKind.Date => expected.DateValue,
        StatementScalarKind.Boolean => expected.BooleanValue,
        _ => null
    };
    if (expectedKind == ActualKind.Decimal)
    {
        var difference = Math.Abs((decimal)actual.Value! - (decimal)expectedValue!);
        if (formula ? difference > FormulaTolerance : difference != 0)
            mismatches.Add($"{path} {actual.Value} != model {expectedValue}");
    }
    else if (!Equals(actual.Value, expectedValue))
        mismatches.Add($"{path} '{actual.Value}' != model '{expectedValue}'");
}

static bool ContainsEmptyAggregate(StatementFormula formula) => formula switch
{
    StatementFormulaAggregate aggregate => aggregate.Operands.IsEmpty || aggregate.Operands.Any(ContainsEmptyAggregate),
    StatementFormulaUnary unary => ContainsEmptyAggregate(unary.Operand),
    StatementFormulaBinary binary => ContainsEmptyAggregate(binary.Left) || ContainsEmptyAggregate(binary.Right),
    StatementFormulaConditional conditional => ContainsEmptyAggregate(conditional.Condition) ||
                                               ContainsEmptyAggregate(conditional.WhenTrue) ||
                                               ContainsEmptyAggregate(conditional.WhenFalse),
    _ => false
};

sealed record Options(string ConnectionFile, string OriginalDirectory, string XlsxDirectory, string OdsAsXlsxDirectory, string EvidenceDirectory)
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
        return new Options(Required("--connection-file"), Required("--original"), Required("--xlsx"), Required("--ods-as-xlsx"), Required("--evidence"));
    }
}

enum ActualKind { Missing, Decimal, Text, Date, Boolean, Error }
sealed record ActualValue(ActualKind Kind, object? Value);
sealed record Scenario(string Name, string Database);
sealed record ScenarioResult(string Scenario, bool Passed, int TriFormatCellsCompared, int RendererCellsCompared, int ModelValuesCompared, int ManagementExpressionResults, int ApprovedInvalidExpressionCells, int MismatchCount, IReadOnlyList<string> Mismatches);
sealed record Evidence(DateTimeOffset GeneratedAtUtc, string ComparisonContract, bool Passed, IReadOnlyList<ScenarioResult> Scenarios);
