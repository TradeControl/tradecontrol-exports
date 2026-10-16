using System.Collections.Immutable;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using TCExports.Generator.CashStatement.Model;
using TCExports.Generator.Ods;

if (args.Length == 3 && args[0] == "--compare")
{
    var expected = Inspect(Path.GetFullPath(args[1]), requireCanonicalPackage: false);
    var actual = Inspect(Path.GetFullPath(args[2]), requireCanonicalPackage: false);
    Equal(expected.FormulaCells.Count, actual.FormulaCells.Count, "round-trip formula count");
    var recalculatedPlaceholders = 0;
    foreach (var item in expected.FormulaCells)
    {
        if (!actual.FormulaCells.TryGetValue(item.Key, out var roundTrip))
            throw new InvalidDataException($"round-trip formula missing at {item.Key}");
        Equal(NormalizeFormulaNumbers(item.Value.Formula), NormalizeFormulaNumbers(roundTrip.Formula),
            $"round-trip formula at {item.Key}");
        if (item.Value.ValueType == "missing")
        {
            Require(roundTrip.ValueType is "float" or "currency" or "percentage",
                $"recalculated expression value type at {item.Key}: '{roundTrip.ValueType}'");
            Require(decimal.TryParse(roundTrip.Value, System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out _),
                $"recalculated expression is not numeric at {item.Key}: '{roundTrip.Value}'");
            recalculatedPlaceholders++;
        }
        else
        {
            Equal(item.Value.ValueType, roundTrip.ValueType, $"round-trip cached value type at {item.Key}");
            EqualCachedValue(item.Value, roundTrip, item.Key);
        }
        Require(!roundTrip.DisplayText.StartsWith('#') && !roundTrip.DisplayText.StartsWith("Err:", StringComparison.Ordinal),
            $"round-trip error token at {item.Key}: {roundTrip.DisplayText}");
    }
    Console.WriteLine($"PASS compared {expected.FormulaCells.Count} formula cells; " +
                      $"Calc evaluated {recalculatedPlaceholders} management-expression cells");
    return;
}

var output = Path.GetFullPath(args.FirstOrDefault() ?? Path.Combine("TestResults", "cash-statement-ods-hostile-fixture.ods"));
Directory.CreateDirectory(Path.GetDirectoryName(output)!);

var tests = new (string Name, Func<Task> Run)[]
{
    ("hostile shared-model fixture", HostileFixtureAsync),
    ("deterministic package", DeterministicPackageAsync),
    ("package contract rejects mutations", PackageMutationTestsAsync),
    ("limits, cancellation, and disposal", LimitsCancellationAndDisposalAsync),
    ("unallocated formula dependency fails", UnallocatedDependencyAsync)
};

var passed = 0;
foreach (var test in tests)
{
    await test.Run();
    passed++;
    Console.WriteLine($"PASS {test.Name}");
}

Console.WriteLine($"{passed}/{tests.Length} ODS tests passed");
Console.WriteLine($"fixture={output}");

async Task HostileFixtureAsync()
{
    await WriteAsync(output, HostileFixture());
    var snapshot = Inspect(output);
    Equal(6, snapshot.Entries.Length, "entry count");
    Equal("mimetype", snapshot.Entries[0], "first entry");
    Equal(20, snapshot.FormulaCount, "formula count");
    Equal(20, snapshot.CachedFormulaCount, "cached formula count");
    Require(snapshot.Formulas.All(formula => formula.StartsWith("of:=", StringComparison.Ordinal)), "non-OpenFormula formula");
    Equal("urn:oasis:names:tc:opendocument:xmlns:of:1.2",
        snapshot.Content.Root!.GetNamespaceOfPrefix("of")!.NamespaceName,
        "OpenFormula namespace");
    Equal("http://openoffice.org/2004/office",
        snapshot.Settings.Root!.GetNamespaceOfPrefix("ooo")!.NamespaceName,
        "OpenOffice settings namespace");
    Require(snapshot.Formulas.Any(formula => formula.Contains("IF(", StringComparison.Ordinal)), "guarded division missing");
    Require(snapshot.Formulas.Any(formula => formula.Contains("SUM(", StringComparison.Ordinal)), "aggregate formula missing");
    Require(snapshot.Content.Descendants(Ns.Table + "table-column").Any(column =>
        (string?)column.Attribute(Ns.Table + "visibility") == "collapse"), "hidden marker column missing");
    Require(snapshot.Settings.Descendants(Ns.Config + "config-item").Any(item =>
        (string?)item.Attribute(Ns.Config + "name") == "HorizontalSplitPosition" && item.Value == "3"), "frozen columns missing");
    Require(snapshot.Settings.Descendants(Ns.Config + "config-item").Any(item =>
        (string?)item.Attribute(Ns.Config + "name") == "VerticalSplitPosition" && item.Value == "2"), "frozen rows missing");
    Require(snapshot.Settings.Descendants(Ns.Config + "config-item").Any(item =>
        (string?)item.Attribute(Ns.Config + "name") == "ActiveSplitRange" && item.Value == "3"), "active frozen pane missing");
    Require(snapshot.Settings.Descendants(Ns.Config + "config-item").Any(item =>
        (string?)item.Attribute(Ns.Config + "name") == "PositionRight" && item.Value == "3"), "right pane position missing");
    Require(snapshot.Settings.Descendants(Ns.Config + "config-item").Any(item =>
        (string?)item.Attribute(Ns.Config + "name") == "PositionBottom" && item.Value == "2"), "bottom pane position missing");
    Require(snapshot.Styles.Descendants(Ns.Style + "map").Any(map =>
        (string?)map.Attribute(Ns.Style + "condition") == "value()<0"), "negative number style missing");
    Require(snapshot.Content.Descendants(Ns.Table + "table").Any(table =>
        (string?)table.Attribute(Ns.Table + "protected") == "true"), "sheet protection missing");
    Require(snapshot.Content.Descendants(Ns.Style + "table-cell-properties").Any(properties =>
        (string?)properties.Attribute(Ns.Style + "cell-protect") == "protected"), "locked cell style missing");
    Require(snapshot.Content.Descendants(Ns.Style + "table-cell-properties").Any(properties =>
        (string?)properties.Attribute(Ns.Style + "cell-protect") == "none"), "unlocked cell style missing");
    Require(snapshot.Content.Descendants(Ns.Style + "table-cell-properties").Any(properties =>
        (string?)properties.Attribute(Ns.Fo + "background-color") == "#ffff00"), "active-period fill missing");
    Require(snapshot.Content.Descendants(Ns.Style + "table-cell-properties").Any(properties =>
        (string?)properties.Attribute(Ns.Fo + "border-right") == "1.5pt solid #000000"), "annual border missing");
}

async Task DeterministicPackageAsync()
{
    var document = HostileFixture();
    using var first = new MemoryStream();
    using var second = new MemoryStream();
    var writer = new OdsPackageWriter();
    await writer.WriteAsync(first, document);
    await writer.WriteAsync(second, document);
    Require(first.ToArray().SequenceEqual(second.ToArray()), "identical inputs did not produce identical packages");
}

async Task PackageMutationTestsAsync()
{
    var directory = Path.Combine(Path.GetDirectoryName(output)!, $"mutations-{Guid.NewGuid():N}");
    Directory.CreateDirectory(directory);
    try
    {
        var missing = Path.Combine(directory, "missing-styles.ods");
        File.Copy(output, missing);
        using (var archive = ZipFile.Open(missing, ZipArchiveMode.Update))
            archive.GetEntry("styles.xml")!.Delete();
        Throws<InvalidDataException>(() => Inspect(missing), "missing required part");

        var formula = Path.Combine(directory, "missing-cache.ods");
        File.Copy(output, formula);
        ReplaceEntry(formula, "content.xml", document =>
        {
            document.Descendants(Ns.Table + "table-cell")
                .First(cell => cell.Attribute(Ns.Table + "formula") is not null)
                .Attribute(Ns.Office + "value")!.Remove();
        });
        Throws<InvalidDataException>(() => Inspect(formula), "formula cache");

        var manifest = Path.Combine(directory, "bad-manifest.ods");
        File.Copy(output, manifest);
        ReplaceEntry(manifest, "META-INF/manifest.xml", document =>
        {
            document.Descendants(Ns.Manifest + "file-entry")
                .First(entry => (string?)entry.Attribute(Ns.Manifest + "full-path") == "content.xml")
                .Remove();
        });
        Throws<InvalidDataException>(() => Inspect(manifest), "manifest relationship");

        var wrongOrder = Path.Combine(directory, "wrong-entry-order.ods");
        ReorderPackage(output, wrongOrder);
        Throws<InvalidDataException>(() => Inspect(wrongOrder), "mimetype ordering");

        var wrongNamespace = Path.Combine(directory, "wrong-table-namespace.ods");
        File.Copy(output, wrongNamespace);
        ReplaceEntry(wrongNamespace, "content.xml", document =>
            document.Root!.SetAttributeValue(XNamespace.Xmlns + "table", "urn:invalid:table"));
        Throws<InvalidDataException>(() => Inspect(wrongNamespace), "table namespace");
    }
    finally
    {
        Directory.Delete(directory, recursive: true);
    }

    await Task.CompletedTask;
}

async Task LimitsCancellationAndDisposalAsync()
{
    var writer = new OdsPackageWriter();
    using var cellsLimited = new MemoryStream();
    await ThrowsAsync<InvalidDataException>(() => writer.WriteAsync(
        cellsLimited,
        HostileFixture(),
        new OdsWriteLimits(MaximumCells: 10)), "cell limit");

    using var columnsLimited = new MemoryStream();
    await ThrowsAsync<InvalidDataException>(() => writer.WriteAsync(
        columnsLimited,
        HostileFixture(),
        new OdsWriteLimits(MaximumColumnsPerWorksheet: 3)), "column limit");

    using var bytesLimited = new MemoryStream();
    await ThrowsAsync<InvalidDataException>(() => writer.WriteAsync(
        bytesLimited,
        HostileFixture(),
        new OdsWriteLimits(MaximumOutputBytes: 512)), "byte limit");

    using var cancelled = new CancellationTokenSource();
    cancelled.Cancel();
    using var cancelledOutput = new MemoryStream();
    await ThrowsAsync<OperationCanceledException>(() => writer.WriteAsync(
        cancelledOutput,
        HostileFixture(),
        cancellationToken: cancelled.Token), "cancellation");

    using var destination = new MemoryStream();
    await writer.WriteAsync(destination, HostileFixture());
    destination.WriteByte(0x00);
    Require(destination.CanWrite, "writer disposed the caller-owned destination");
}

async Task UnallocatedDependencyAsync()
{
    var missingAddress = Address("Z99");
    var cell = Cell("D3", 1m, new StatementFormulaReference(missingAddress));
    var worksheet = new OdsWorksheet(
        FixtureIds.Sheet,
        "Cash Flow",
        Columns(4),
        ImmutableArray.Create(new OdsRow(ImmutableArray.Create(Blank(), Blank(), Blank(), cell))));
    using var stream = new MemoryStream();
    await ThrowsAsync<InvalidDataException>(() => new OdsPackageWriter().WriteAsync(
        stream,
        new OdsSpreadsheetDocument(ImmutableArray.Create(worksheet))), "missing dependency");
}

static async Task WriteAsync(string path, OdsSpreadsheetDocument document)
{
    await using var stream = File.Create(path);
    await new OdsPackageWriter().WriteAsync(stream, document);
}

static OdsSpreadsheetDocument HostileFixture()
{
    var formulas = new Dictionary<string, StatementFormula>(StringComparer.Ordinal)
    {
        ["D3"] = Literal(100m),
        ["E3"] = Literal(-25m),
        ["F3"] = Literal(50m),
        ["P3"] = Sum("D3", "E3", "F3"),
        ["Q3"] = new StatementFormulaTerminal(
            "IF([TOTAL]=0,0,[VALUE]/[TOTAL])",
            new Dictionary<string, StatementCellAddress>
            {
                ["TOTAL"] = Address("P3"),
                ["VALUE"] = Address("D3")
            }.ToImmutableDictionary(StringComparer.Ordinal)),
        ["D5"] = Binary(StatementBinaryOperator.Add, "D3", "D7"),
        ["E5"] = Binary(StatementBinaryOperator.Add, "E3", "E7"),
        ["F5"] = Binary(StatementBinaryOperator.Add, "F3", "F7"),
        ["P5"] = Sum("D5", "E5", "F5"),
        ["P7"] = Sum("D7", "E7", "F7"),
        ["D9"] = Reference("D3"),
        ["D10"] = Sum("D3", "D5", "D7"),
        ["D12"] = Binary(StatementBinaryOperator.Subtract, "D3", "D7"),
        ["E12"] = new StatementFormulaBinary(StatementBinaryOperator.Multiply, Reference("E7"), Literal(2m)),
        ["F12"] = GuardedDivide("F4", "F3", "F4"),
        ["P12"] = Sum("D12", "E12", "F12"),
        ["D13"] = Sum("D14"),
        ["P13"] = Sum("D13"),
        ["D14"] = Sum("D15"),
        ["P14"] = Sum("D14")
    };
    var values = new Dictionary<string, decimal>(StringComparer.Ordinal)
    {
        ["D3"] = 100m, ["E3"] = -25m, ["F3"] = 50m, ["P3"] = 125m, ["Q3"] = 0.8m,
        ["D5"] = 140m, ["E5"] = -15m, ["F5"] = 55m, ["P5"] = 180m,
        ["D7"] = 40m, ["E7"] = 10m, ["F7"] = 5m, ["P7"] = 55m,
        ["D9"] = 100m, ["D10"] = 280m,
        ["D12"] = 60m, ["E12"] = 20m, ["F12"] = 0m, ["P12"] = 80m,
        ["D13"] = 10m, ["P13"] = 10m, ["D14"] = 10m, ["P14"] = 10m, ["D15"] = 10m
    };

    var rows = ImmutableArray.CreateBuilder<OdsRow>(16);
    for (var row = 1; row <= 16; row++)
    {
        var cells = ImmutableArray.CreateBuilder<OdsCell>(17);
        for (var column = 1; column <= 17; column++)
        {
            var address = $"{ColumnName(column)}{row}";
            if (row == 1 && column == 1)
            {
                cells.Add(new OdsCell(StatementValue.FromText("C# ODS hostile fixture — April to March"),
                    PresentationRole: StatementPresentationRole.DocumentTitle,
                    ColumnSpan: 17,
                    Locked: true));
                break;
            }
            if (row == 2)
            {
                var labels = new[] { "Code", "Name", "Marker", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar", "Total", "Ratio" };
                cells.Add(new OdsCell(StatementValue.FromText(labels[column - 1]),
                    PresentationRole: StatementPresentationRole.ColumnHeader));
                continue;
            }

            if (values.TryGetValue(address, out var value))
            {
                var numericKind = column == 17 ? StatementNumericKind.Percentage : StatementNumericKind.Money;
                var role = column == 16 ? StatementPresentationRole.AnnualPeriod :
                    row is 5 or 10 or 13 or 14 ? StatementPresentationRole.NeutralTotal : StatementPresentationRole.Detail;
                cells.Add(new OdsCell(StatementValue.FromDecimal(value), numericKind, role,
                    formulas.GetValueOrDefault(address), Address(address),
                    NumberFormat: column == 17 ? "Pct1" : null,
                    Locked: column == 16,
                    ActivePeriod: column == 4,
                    LeftBorder: column == 16 ? OdsBorderWeight.Thin : OdsBorderWeight.None,
                    RightBorder: column == 16 ? OdsBorderWeight.Thick : OdsBorderWeight.None));
            }
            else if (address == "D16")
            {
                cells.Add(new OdsCell(StatementValue.Missing(StatementScalarKind.Decimal), StatementNumericKind.Money,
                    SemanticAddress: Address(address)));
            }
            else
            {
                cells.Add(new OdsCell(StatementValue.FromDecimal(0m),
                    column == 17 ? StatementNumericKind.Percentage : StatementNumericKind.Money,
                    column == 16 ? StatementPresentationRole.AnnualPeriod : StatementPresentationRole.Detail,
                    SemanticAddress: Address(address)));
            }
        }
        rows.Add(new OdsRow(cells.ToImmutable()));
    }

    var worksheet = new OdsWorksheet(FixtureIds.Sheet, "Cash Flow", Columns(17), rows.MoveToImmutable(), 3, 2,
        Protected: true);
    return new OdsSpreadsheetDocument(ImmutableArray.Create(worksheet), "TCExports C# ODS capability fixture");
}

static ImmutableArray<OdsColumn> Columns(int count) => Enumerable.Range(1, count)
    .Select(column => new OdsColumn(column switch { 1 => 2.4m, 2 => 5.5m, 3 => 1.8m, _ => 2.2m }, column == 3))
    .ToImmutableArray();

static OdsCell Cell(string address, decimal value, StatementFormula formula) =>
    new(StatementValue.FromDecimal(value), StatementNumericKind.Money, StatementPresentationRole.Detail, formula, Address(address));

static OdsCell Blank() => new(StatementValue.Missing(StatementScalarKind.Text));

static StatementFormula Literal(decimal value) => new StatementFormulaLiteral(StatementValue.FromDecimal(value));
static StatementFormula Reference(string address) => new StatementFormulaReference(Address(address));
static StatementFormula Binary(StatementBinaryOperator operation, string left, string right) =>
    new StatementFormulaBinary(operation, Reference(left), Reference(right));
static StatementFormula Sum(params string[] addresses) =>
    new StatementFormulaAggregate(StatementAggregateOperator.Sum, addresses.Select(Reference).ToImmutableArray());
static StatementFormula GuardedDivide(string guard, string numerator, string denominator) =>
    new StatementFormulaConditional(
        new StatementFormulaBinary(StatementBinaryOperator.Equal, Reference(guard), Literal(0m)),
        Literal(0m),
        new StatementFormulaBinary(StatementBinaryOperator.Divide, Reference(numerator), Reference(denominator)));

static StatementCellAddress Address(string a1)
{
    var split = a1.TakeWhile(char.IsLetter).Count();
    return new StatementCellAddress(FixtureIds.Sheet, new StatementRowId($"row:{a1[split..]}"), new StatementColumnId($"column:{a1[..split]}"));
}

static string ColumnName(int number)
{
    var result = string.Empty;
    while (number > 0)
    {
        number--;
        result = (char)('A' + number % 26) + result;
        number /= 26;
    }
    return result;
}

static PackageSnapshot Inspect(string path, bool requireCanonicalPackage = true)
{
    var bytes = File.ReadAllBytes(path);
    Require(bytes.Length >= 30, "package is too short");
    if (requireCanonicalPackage)
    {
        Equal((ushort)0, BitConverter.ToUInt16(bytes, 8), "mimetype compression method");
        Equal((ushort)0, BitConverter.ToUInt16(bytes, 28), "mimetype ZIP extra length");
    }

    using var archive = ZipFile.OpenRead(path);
    var required = new[] { "mimetype", "content.xml", "styles.xml", "settings.xml", "meta.xml", "META-INF/manifest.xml" };
    var entries = archive.Entries.Select(entry => entry.FullName).ToArray();
    if (requireCanonicalPackage)
        Require(entries.SequenceEqual(required), "ODS package entries are missing, additional, or out of order");
    foreach (var name in required)
        Require(archive.GetEntry(name) is not null, $"missing required ODS part: {name}");
    using (var reader = new StreamReader(archive.GetEntry("mimetype")!.Open(), Encoding.ASCII))
        Equal(OdsPackageWriter.MimeType, reader.ReadToEnd(), "mimetype content");

    var content = Load(archive, "content.xml");
    var styles = Load(archive, "styles.xml");
    var settings = Load(archive, "settings.xml");
    var manifest = Load(archive, "META-INF/manifest.xml");
    Equal(Ns.Table.NamespaceName, content.Root!.GetNamespaceOfPrefix("table")!.NamespaceName, "table namespace");
    var formulas = content.Descendants(Ns.Table + "table-cell")
        .Where(cell => cell.Attribute(Ns.Table + "formula") is not null)
        .ToArray();
    if (requireCanonicalPackage && formulas.Any(cell => CachedValue(cell) is null))
        throw new InvalidDataException("Formula cell missing authoritative cached value.");

    foreach (var name in required.Where(name => name.EndsWith(".xml", StringComparison.Ordinal) &&
                                                name != "META-INF/manifest.xml"))
        Require(manifest.Descendants(Ns.Manifest + "file-entry").Any(entry =>
            (string?)entry.Attribute(Ns.Manifest + "full-path") == name), $"manifest missing {name}");

    return new PackageSnapshot(entries, content, styles, settings,
        formulas.Length,
        formulas.Count(cell => CachedValue(cell) is not null),
        formulas.Select(cell => (string)cell.Attribute(Ns.Table + "formula")!).ToArray(),
        FormulaCells(content));
}

static IReadOnlyDictionary<string, FormulaCellSnapshot> FormulaCells(XDocument content)
{
    var result = new Dictionary<string, FormulaCellSnapshot>(StringComparer.Ordinal);
    foreach (var table in content.Descendants(Ns.Table + "table"))
    {
        var sheetName = (string?)table.Attribute(Ns.Table + "name") ?? "Sheet";
        var rowNumber = 1;
        foreach (var row in table.Elements(Ns.Table + "table-row"))
        {
            var rowRepeat = (int?)row.Attribute(Ns.Table + "number-rows-repeated") ?? 1;
            for (var rowOffset = 0; rowOffset < rowRepeat; rowOffset++)
            {
                var columnNumber = 1;
                foreach (var cell in row.Elements())
                {
                    var columnRepeat = (int?)cell.Attribute(Ns.Table + "number-columns-repeated") ?? 1;
                    var formula = (string?)cell.Attribute(Ns.Table + "formula");
                    if (formula is not null)
                    {
                        var valueType = (string?)cell.Attribute(Ns.Office + "value-type") ?? "missing";
                        var value = CachedValue(cell) ?? string.Empty;
                        var display = string.Concat(cell.Descendants(Ns.Text + "p").Select(paragraph => paragraph.Value));
                        for (var repeat = 0; repeat < columnRepeat; repeat++)
                            result.Add($"{sheetName}!{ColumnName(columnNumber + repeat)}{rowNumber + rowOffset}",
                                new FormulaCellSnapshot(formula, valueType, value, display));
                    }
                    columnNumber += columnRepeat;
                }
            }
            rowNumber += rowRepeat;
        }
    }
    return result;
}

static string? CachedValue(XElement cell) => (string?)cell.Attribute(Ns.Office + "value-type") switch
{
    "float" or "currency" or "percentage" => (string?)cell.Attribute(Ns.Office + "value"),
    "string" => (string?)cell.Attribute(Ns.Office + "string-value"),
    "date" => (string?)cell.Attribute(Ns.Office + "date-value"),
    "boolean" => (string?)cell.Attribute(Ns.Office + "boolean-value"),
    _ => null
};

static XDocument Load(ZipArchive archive, string name)
{
    using var stream = archive.GetEntry(name)!.Open();
    return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
}

static void ReplaceEntry(string path, string name, Action<XDocument> mutate)
{
    using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
    var entry = archive.GetEntry(name)!;
    XDocument document;
    using (var stream = entry.Open()) document = XDocument.Load(stream);
    mutate(document);
    entry.Delete();
    var replacement = archive.CreateEntry(name, CompressionLevel.Optimal);
    replacement.LastWriteTime = new DateTimeOffset(2000, 1, 1, 0, 0, 0, TimeSpan.Zero);
    using var writer = new StreamWriter(replacement.Open(), new UTF8Encoding(false));
    document.Save(writer, SaveOptions.DisableFormatting);
}

static void ReorderPackage(string source, string destination)
{
    var entries = new Dictionary<string, byte[]>(StringComparer.Ordinal);
    using (var archive = ZipFile.OpenRead(source))
    {
        foreach (var entry in archive.Entries)
        {
            using var stream = entry.Open();
            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            entries.Add(entry.FullName, buffer.ToArray());
        }
    }

    using var file = File.Create(destination);
    using var reordered = new ZipArchive(file, ZipArchiveMode.Create, leaveOpen: false, Encoding.UTF8);
    foreach (var name in entries.Keys.OrderBy(name => name == "content.xml" ? 0 : name == "mimetype" ? 1 : 2))
    {
        var entry = reordered.CreateEntry(name, name == "mimetype" ? CompressionLevel.NoCompression : CompressionLevel.Optimal);
        entry.LastWriteTime = new DateTimeOffset(2000, 1, 1, 0, 0, 0, TimeSpan.Zero);
        using var stream = entry.Open();
        stream.Write(entries[name]);
    }
}

static void Require(bool condition, string message)
{
    if (!condition) throw new InvalidDataException(message);
}

static void Equal<T>(T expected, T actual, string context) where T : notnull
{
    if (!EqualityComparer<T>.Default.Equals(expected, actual))
        throw new InvalidDataException($"{context}: expected '{expected}', actual '{actual}'.");
}

static void EqualCachedValue(FormulaCellSnapshot expected, FormulaCellSnapshot actual, string cell)
{
    if (expected.ValueType is "float" or "currency" or "percentage")
    {
        if (!decimal.TryParse(expected.Value, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var expectedNumber) ||
            !decimal.TryParse(actual.Value, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var actualNumber) ||
            expectedNumber != actualNumber)
            throw new InvalidDataException(
                $"round-trip cached numeric value at {cell}: expected '{expected.Value}', actual '{actual.Value}'.");
        return;
    }

    Equal(expected.Value, actual.Value, $"round-trip cached value at {cell}");
}

static string NormalizeFormulaNumbers(string formula)
{
    var result = new StringBuilder(formula.Length);
    var segment = new StringBuilder();
    var inString = false;

    void AppendSegment()
    {
        var normalized = System.Text.RegularExpressions.Regex.Replace(segment.ToString(),
            @"(?<![A-Za-z_])\d+(?:\.\d+)?(?:[Ee][+-]?\d+)?(?![A-Za-z_])",
            match => decimal.TryParse(match.Value, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var number)
                    ? number.ToString("G29", System.Globalization.CultureInfo.InvariantCulture)
                    : match.Value);
        result.Append(normalized);
        segment.Clear();
    }

    foreach (var character in formula)
    {
        if (character == '"')
        {
            if (!inString)
                AppendSegment();
            result.Append(character);
            inString = !inString;
        }
        else if (inString)
            result.Append(character);
        else
            segment.Append(character);
    }

    AppendSegment();
    return result.ToString();
}

static void Throws<T>(Action action, string context) where T : Exception
{
    try { action(); }
    catch (T) { return; }
    throw new InvalidDataException($"{context}: expected {typeof(T).Name}.");
}

static async Task ThrowsAsync<T>(Func<Task> action, string context) where T : Exception
{
    try { await action(); }
    catch (T) { return; }
    throw new InvalidDataException($"{context}: expected {typeof(T).Name}.");
}

internal static class Ns
{
    public static readonly XNamespace Office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    public static readonly XNamespace Table = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
    public static readonly XNamespace Text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
    public static readonly XNamespace Style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
    public static readonly XNamespace Fo = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
    public static readonly XNamespace Config = "urn:oasis:names:tc:opendocument:xmlns:config:1.0";
    public static readonly XNamespace Manifest = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
}

internal sealed record PackageSnapshot(
    string[] Entries,
    XDocument Content,
    XDocument Styles,
    XDocument Settings,
    int FormulaCount,
    int CachedFormulaCount,
    string[] Formulas,
    IReadOnlyDictionary<string, FormulaCellSnapshot> FormulaCells);

internal sealed record FormulaCellSnapshot(string Formula, string ValueType, string Value, string DisplayText);

internal static class FixtureIds
{
    public static readonly StatementSheetId Sheet = new("cash-flow");
}
