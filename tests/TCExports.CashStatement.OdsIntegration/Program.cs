using System.Collections.Immutable;
using System.Diagnostics;
using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;
using Microsoft.Data.SqlClient;
using TCExports.Generator.CashStatement;
using TCExports.Generator.CashStatement.Model;
using TCExports.Generator.Data;
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
    var allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
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
        var build = await new CashStatementBuilder(new SqlServerCashFlowRepository()).BuildAsync(
            connection.ConnectionString,
            request,
            "cash-statement-phase-6",
            DateTimeOffset.UtcNow);
        if (!build.IsSuccess || build.Value is null)
            throw new InvalidOperationException(string.Join("; ", build.Violations.Select(item => $"{item.Code}: {item.Message}")));

        var bytes = await new CashStatementOdsRenderer().RenderAsync(build.Value);
        var semantic = SemanticOdsVerifier.Verify(build.Value, bytes);
        var fileName = $"{scenario.Name}__all-options__shared-model.ods";
        await File.WriteAllBytesAsync(Path.Combine(options.OutputDirectory, fileName), bytes);
        stopwatch.Stop();
        results.Add(new ScenarioResult(
            scenario.Name, scenario.Database, true, fileName, bytes.LongLength,
            Convert.ToHexString(SHA256.HashData(bytes)), stopwatch.ElapsedMilliseconds,
            GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore,
            Process.GetCurrentProcess().WorkingSet64, semantic.Cells, semantic.Formulas, null));
        Console.WriteLine($"PASS {scenario.Name}: {semantic.Cells} semantic cells, {semantic.Formulas} formulas, " +
                          $"{bytes.Length} bytes in {stopwatch.ElapsedMilliseconds} ms");
    }
    catch (Exception exception)
    {
        stopwatch.Stop();
        results.Add(new ScenarioResult(
            scenario.Name, scenario.Database, false, null, 0, null, stopwatch.ElapsedMilliseconds,
            GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore,
            Process.GetCurrentProcess().WorkingSet64, 0, 0, exception.ToString()));
        Console.Error.WriteLine($"FAIL {scenario.Name}: {exception.Message}");
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
    long AllocatedBytes,
    long ProcessWorkingSetBytes,
    int SemanticCells,
    int FormulaCells,
    string? Error);
sealed record Manifest(DateTimeOffset GeneratedAtUtc, bool Passed, IReadOnlyList<ScenarioResult> Scenarios);

static class SemanticOdsVerifier
{
    private static readonly XNamespace Office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    private static readonly XNamespace Table = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
    private static readonly XNamespace Text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
    private static readonly XNamespace Config = "urn:oasis:names:tc:opendocument:xmlns:config:1.0";

    public static VerificationResult Verify(CashStatementDocument document, byte[] package)
    {
        var layout = new CashStatementOdsLayoutAllocator().Allocate(document, CashStatementLimits.Default);
        var references = layout.Cells.ToDictionary(
            pair => pair.Key,
            pair => new OdsCellReference(pair.Value.SheetName, pair.Value.RowNumber, pair.Value.ColumnNumber));
        var translator = new CashStatementOdsFormulaTranslator(references);
        using var buffer = new MemoryStream(package, writable: false);
        using var archive = new ZipArchive(buffer, ZipArchiveMode.Read);
        VerifyFreezePanes(document, archive);
        using var contentStream = archive.GetEntry("content.xml")?.Open() ??
                                  throw new InvalidDataException("ODS content.xml is missing.");
        var content = XDocument.Load(contentStream);
        var tables = content.Descendants(Table + "table").ToDictionary(
            table => (string?)table.Attribute(Table + "name") ?? string.Empty,
            StringComparer.Ordinal);

        var cellCount = 0;
        var formulaCount = 0;
        foreach (var sheet in document.Sheets.Where(sheet =>
                     sheet.Kind is StatementSheetKind.CashFlow or StatementSheetKind.Reconciliation))
        {
            if (!tables.TryGetValue(sheet.Name, out var table))
                throw new InvalidDataException($"ODS sheet '{sheet.Name}' is missing.");
            var firstPeriod = sheet.Kind == StatementSheetKind.CashFlow
                ? sheet.Columns.Where(column => column.PeriodId is not null).MinBy(column => column.PresentationOrder)?.Id
                : null;
            foreach (var row in sheet.Sections.SelectMany(section => section.Rows))
            foreach (var pair in row.Cells)
            {
                var semantic = new StatementCellAddress(sheet.Id, row.Id, pair.Key);
                if (!layout.Cells.TryGetValue(semantic, out var physical))
                    throw new InvalidDataException($"Semantic cell '{semantic}' has no ODS allocation.");
                var cell = CellAt(table, physical.RowNumber, physical.ColumnNumber);
                var suppressedExpressionError = row.Kind == StatementRowKind.Expression &&
                                                pair.Value.Value.State == StatementValueState.Unsupported &&
                                                pair.Value.Formula is null && pair.Key != firstPeriod;
                var engineEvaluatedExpression = pair.Value.Formula is not null &&
                                                pair.Value.Value is
                                                {
                                                    State: StatementValueState.Unsupported,
                                                    Reason: "spreadsheet-engine-evaluation-required"
                                                };
                if (suppressedExpressionError || engineEvaluatedExpression)
                    Require(cell.Attribute(Office + "value-type") is null,
                        $"Renderer-owned missing cache '{semantic}' contains a value.");
                else
                    VerifyValue(pair.Value.Value, cell, semantic);

                var renderedFormula = row.Kind == StatementRowKind.Capital
                    ? new StatementFormulaAggregate(
                        StatementAggregateOperator.Sum,
                        sheet.Sections.Single(section => section.Kind == StatementSectionKind.BalanceSheet)
                            .Rows.Where(asset => asset.Kind == StatementRowKind.BalanceSheetAsset)
                            .Select(asset => (StatementFormula)new StatementFormulaReference(
                                new StatementCellAddress(sheet.Id, asset.Id, pair.Key)))
                            .ToImmutableArray())
                    : pair.Value.Formula;
                var expectedFormula = renderedFormula is not null && !ContainsEmptyAggregate(renderedFormula)
                    ? translator.Translate(renderedFormula)
                    : null;
                Equal(expectedFormula, (string?)cell.Attribute(Table + "formula"), $"formula {semantic}");
                if (expectedFormula is not null)
                    formulaCount++;
                cellCount++;
            }
        }
        Equal(layout.Cells.Count, cellCount, "semantic allocation count");
        return new VerificationResult(cellCount, formulaCount);
    }

    private static void VerifyFreezePanes(CashStatementDocument document, ZipArchive archive)
    {
        using var settingsStream = archive.GetEntry("settings.xml")?.Open() ??
                                   throw new InvalidDataException("ODS settings.xml is missing.");
        var settings = XDocument.Load(settingsStream);
        Equal("http://openoffice.org/2004/office",
            settings.Root?.GetNamespaceOfPrefix("ooo")?.NamespaceName,
            "OpenOffice settings namespace");
        var tableViews = settings.Descendants(Config + "config-item-map-entry")
            .Where(entry => entry.Attribute(Config + "name") is not null)
            .ToDictionary(entry => (string)entry.Attribute(Config + "name")!, StringComparer.Ordinal);

        var cashFlow = document.Sheets.Single(sheet => sheet.Kind == StatementSheetKind.CashFlow);
        VerifyFreezePane(tableViews, cashFlow.Name, expectedColumns: 3, expectedRows: 4);

        var reconciliation = document.Sheets.SingleOrDefault(sheet => sheet.Kind == StatementSheetKind.Reconciliation);
        if (reconciliation is not null)
            VerifyFreezePane(tableViews, reconciliation.Name, expectedColumns: 2, expectedRows: 4);
    }

    private static void VerifyFreezePane(
        IReadOnlyDictionary<string, XElement> tableViews,
        string sheetName,
        int expectedColumns,
        int expectedRows)
    {
        if (!tableViews.TryGetValue(sheetName, out var tableView))
            throw new InvalidDataException($"ODS view settings for sheet '{sheetName}' are missing.");

        var items = tableView.Elements(Config + "config-item").ToDictionary(
            item => (string?)item.Attribute(Config + "name") ?? string.Empty,
            item => item.Value,
            StringComparer.Ordinal);
        Equal(expectedColumns.ToString(CultureInfo.InvariantCulture), items["HorizontalSplitPosition"],
            $"horizontal freeze for {sheetName}");
        Equal(expectedRows.ToString(CultureInfo.InvariantCulture), items["VerticalSplitPosition"],
            $"vertical freeze for {sheetName}");
        Equal(expectedColumns.ToString(CultureInfo.InvariantCulture), items["CursorPositionX"],
            $"cursor column for {sheetName}");
        Equal(expectedRows.ToString(CultureInfo.InvariantCulture), items["CursorPositionY"],
            $"cursor row for {sheetName}");
        Equal(((expectedColumns > 0 ? 1 : 0) + (expectedRows > 0 ? 2 : 0)).ToString(CultureInfo.InvariantCulture),
            items["ActiveSplitRange"], $"active frozen pane for {sheetName}");
        Equal("0", items["PositionLeft"], $"left pane position for {sheetName}");
        Equal(expectedColumns.ToString(CultureInfo.InvariantCulture), items["PositionRight"],
            $"right pane position for {sheetName}");
        Equal("0", items["PositionTop"], $"top pane position for {sheetName}");
        Equal(expectedRows.ToString(CultureInfo.InvariantCulture), items["PositionBottom"],
            $"bottom pane position for {sheetName}");
    }

    private static XElement CellAt(XElement table, int rowNumber, int columnNumber)
    {
        var rows = table.Elements(Table + "table-row").ToArray();
        if (rowNumber < 1 || rowNumber > rows.Length)
            throw new InvalidDataException($"ODS row {rowNumber} is missing.");
        var cells = rows[rowNumber - 1].Elements()
            .Where(element => element.Name == Table + "table-cell" || element.Name == Table + "covered-table-cell")
            .ToArray();
        if (columnNumber < 1 || columnNumber > cells.Length)
            throw new InvalidDataException($"ODS cell {rowNumber},{columnNumber} is missing.");
        return cells[columnNumber - 1];
    }

    private static void VerifyValue(StatementValue expected, XElement actual, StatementCellAddress address)
    {
        if (expected.State == StatementValueState.Missing)
        {
            Require(actual.Attribute(Office + "value-type") is null, $"missing value {address}");
            return;
        }
        if (expected.State == StatementValueState.Unsupported)
        {
            Equal("string", (string?)actual.Attribute(Office + "value-type"), $"unsupported type {address}");
            Equal(expected.Reason ?? "Unsupported value", (string?)actual.Attribute(Office + "string-value"),
                $"unsupported value {address}");
            return;
        }

        switch (expected.Kind)
        {
            case StatementScalarKind.Decimal:
                Equal("float", (string?)actual.Attribute(Office + "value-type"), $"decimal type {address}");
                var lexical = (string?)actual.Attribute(Office + "value") ?? string.Empty;
                Require(decimal.TryParse(lexical, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) &&
                        number == expected.DecimalValue, $"decimal value {address}: '{lexical}'.");
                break;
            case StatementScalarKind.Text:
                Equal("string", (string?)actual.Attribute(Office + "value-type"), $"text type {address}");
                Equal(expected.TextValue ?? string.Empty, (string?)actual.Attribute(Office + "string-value"),
                    $"text value {address}");
                break;
            case StatementScalarKind.Date:
                Equal("date", (string?)actual.Attribute(Office + "value-type"), $"date type {address}");
                Equal(expected.DateValue?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                    (string?)actual.Attribute(Office + "date-value"), $"date value {address}");
                break;
            case StatementScalarKind.Boolean:
                Equal("boolean", (string?)actual.Attribute(Office + "value-type"), $"Boolean type {address}");
                Equal(expected.BooleanValue == true ? "true" : "false",
                    (string?)actual.Attribute(Office + "boolean-value"), $"Boolean value {address}");
                break;
        }
        var display = string.Concat(actual.Descendants(Text + "p").Select(item => item.Value));
        Require(!display.StartsWith('#') && !display.StartsWith("Err:", StringComparison.Ordinal),
            $"ODS error token at {address}: {display}");
    }

    private static bool ContainsEmptyAggregate(StatementFormula formula) => formula switch
    {
        StatementFormulaAggregate aggregate => aggregate.Operands.IsEmpty || aggregate.Operands.Any(ContainsEmptyAggregate),
        StatementFormulaUnary unary => ContainsEmptyAggregate(unary.Operand),
        StatementFormulaBinary binary => ContainsEmptyAggregate(binary.Left) || ContainsEmptyAggregate(binary.Right),
        StatementFormulaConditional conditional => ContainsEmptyAggregate(conditional.Condition) ||
                                                   ContainsEmptyAggregate(conditional.WhenTrue) ||
                                                   ContainsEmptyAggregate(conditional.WhenFalse),
        _ => false
    };

    private static void Require(bool condition, string message)
    {
        if (!condition) throw new InvalidDataException(message);
    }

    private static void Equal<T>(T expected, T actual, string context)
    {
        if (!EqualityComparer<T>.Default.Equals(expected, actual))
            throw new InvalidDataException($"{context}: expected '{expected}', actual '{actual}'.");
    }
}

sealed record VerificationResult(int Cells, int Formulas);
