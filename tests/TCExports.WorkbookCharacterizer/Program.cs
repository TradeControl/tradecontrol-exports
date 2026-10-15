using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using ClosedXML.Excel;

if (args.Length == 0)
{
    PrintUsage();
    return 2;
}

var options = new JsonSerializerOptions
{
    WriteIndented = true,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
};

if (args[0].Equals("inspect", StringComparison.OrdinalIgnoreCase) && args.Length == 3)
{
    var inputDirectory = Path.GetFullPath(args[1]);
    var outputDirectory = Path.GetFullPath(args[2]);
    Directory.CreateDirectory(outputDirectory);

    var inputs = Directory.EnumerateFiles(inputDirectory, "*.xlsx", SearchOption.TopDirectoryOnly)
        .OrderBy(Path.GetFileName, StringComparer.Ordinal)
        .ToArray();

    if (inputs.Length == 0)
    {
        Console.Error.WriteLine($"No XLSX files found in {inputDirectory}");
        return 3;
    }

    var index = new List<WorkbookIndexEntry>();
    foreach (var input in inputs)
    {
        var snapshot = WorkbookInspector.Inspect(input);
        var output = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(input) + ".snapshot.json");
        File.WriteAllText(output, JsonSerializer.Serialize(snapshot, options), new UTF8Encoding(false));
        index.Add(new WorkbookIndexEntry(
            snapshot.FileName,
            snapshot.FileLength,
            snapshot.RawSha256,
            snapshot.NormalizedPackageSha256,
            snapshot.Worksheets.Count,
            snapshot.Worksheets.Sum(sheet => sheet.Cells.Count),
            snapshot.Worksheets.Sum(sheet => sheet.Cells.Count(cell => cell.Formula is not null)),
            Path.GetFileName(output)));
        Console.WriteLine($"{snapshot.FileName}: {snapshot.Worksheets.Count} sheets, " +
                          $"{snapshot.Worksheets.Sum(sheet => sheet.Cells.Count)} cells, " +
                          $"{snapshot.Worksheets.Sum(sheet => sheet.Cells.Count(cell => cell.Formula is not null))} formulas");
    }

    File.WriteAllText(
        Path.Combine(outputDirectory, "index.json"),
        JsonSerializer.Serialize(index, options),
        new UTF8Encoding(false));

    return 0;
}

if (args[0].Equals("compare", StringComparison.OrdinalIgnoreCase) && args.Length == 4)
{
    var expected = WorkbookInspector.Inspect(Path.GetFullPath(args[1]));
    var actual = WorkbookInspector.Inspect(Path.GetFullPath(args[2]));
    var report = WorkbookComparer.Compare(expected, actual);
    var reportPath = Path.GetFullPath(args[3]);
    Directory.CreateDirectory(Path.GetDirectoryName(reportPath) ?? throw new InvalidOperationException("Comparison report has no directory."));
    File.WriteAllText(reportPath, JsonSerializer.Serialize(report, options), new UTF8Encoding(false));
    Console.WriteLine($"Strict: {report.Equivalent}; semantic definition: {report.SemanticDefinitionEquivalent}; semantic result: {report.SemanticResultEquivalent}; differences: {report.TotalDifferenceCount}");
    foreach (var category in report.CategoryCounts.OrderBy(item => item.Key, StringComparer.Ordinal))
        Console.WriteLine($"  {category.Key}: {category.Value}");
    foreach (var difference in report.Differences.GroupBy(item => item.Category, StringComparer.Ordinal).SelectMany(group => group.Take(3)))
        Console.WriteLine($"  sample {difference.Category}: {difference.Location} — {difference.Message}");
    return report.Equivalent ? 0 : 1;
}

if (args[0].Equals("self-test", StringComparison.OrdinalIgnoreCase) && args.Length == 3)
{
    var baseline = Path.GetFullPath(args[1]);
    var outputDirectory = Path.GetFullPath(args[2]);
    Directory.CreateDirectory(outputDirectory);
    var results = MutationProof.Run(baseline, outputDirectory);
    File.WriteAllText(Path.Combine(outputDirectory, "mutation-proof.json"), JsonSerializer.Serialize(results, options), new UTF8Encoding(false));
    foreach (var result in results)
        Console.WriteLine($"{result.Mutation}: {(result.Passed ? "PASS" : "FAIL")} [{string.Join(", ", result.DetectedCategories)}]");
    return results.All(result => result.Passed) ? 0 : 1;
}

PrintUsage();
return 2;

static void PrintUsage()
{
    Console.Error.WriteLine("Usage:");
    Console.Error.WriteLine("  WorkbookCharacterizer inspect <input-directory> <output-directory>");
    Console.Error.WriteLine("  WorkbookCharacterizer compare <expected.xlsx> <actual.xlsx> <report.json>");
    Console.Error.WriteLine("  WorkbookCharacterizer self-test <baseline.xlsx> <output-directory>");
}

internal static class WorkbookInspector
{
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace Relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace PackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace CoreProperties = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    private static readonly XNamespace DublinCoreTerms = "http://purl.org/dc/terms/";

    public static WorkbookSnapshot Inspect(string path)
    {
        var file = new FileInfo(path);
        using var archive = ZipFile.OpenRead(path);
        var package = ReadPackage(archive);
        var sheetParts = ReadSheetParts(archive);
        var styles = ReadStyles(archive);

        using var workbook = new XLWorkbook(path);
        var sheets = new List<WorksheetSnapshot>();
        foreach (var worksheet in workbook.Worksheets.OrderBy(sheet => sheet.Position))
        {
            sheetParts.TryGetValue(worksheet.Name, out var part);
            var cellParts = part is null
                ? new Dictionary<string, CellPartSnapshot>(StringComparer.Ordinal)
                : ReadCellParts(archive, part.Path);

            var cells = new List<CellSnapshot>();
            foreach (var cell in worksheet.CellsUsed(XLCellsUsedOptions.All).OrderBy(cell => cell.Address.RowNumber).ThenBy(cell => cell.Address.ColumnNumber))
            {
                var address = cell.Address.ToString() ?? throw new InvalidDataException("Cell has no address.");
                cellParts.TryGetValue(address, out var cellPart);
                var styleIndex = cellPart?.StyleIndex ?? 0;
                styles.TryGetValue(styleIndex, out var styleSignature);
                cells.Add(new CellSnapshot(
                    address,
                    cell.DataType.ToString(),
                    FormatCellValue(cell),
                    cellPart?.StoredValue,
                    cell.GetFormattedString(CultureInfo.GetCultureInfo("en-GB")),
                    cell.HasFormula ? cell.FormulaA1 : null,
                    styleIndex,
                    styleSignature,
                    FormatEffectiveStyle(cell.Style)));
            }

            sheets.Add(new WorksheetSnapshot(
                worksheet.Position,
                worksheet.Name,
                worksheet.Visibility.ToString(),
                worksheet.RangeUsed(XLCellsUsedOptions.All)?.RangeAddress.ToString(),
                part?.FreezePane,
                part?.Protection,
                part?.PrintSettings ?? string.Empty,
                part?.MergedRanges ?? [],
                part?.Rows ?? [],
                part?.Columns ?? [],
                cells));
        }

        var normalizedPackageHash = HashText(string.Join("\n", package
            .OrderBy(entry => entry.NormalizedPath, StringComparer.Ordinal)
            .Select(entry => $"{entry.NormalizedPath}|{entry.NormalizedSha256}")));

        return new WorkbookSnapshot(
            file.Name,
            file.Length,
            HashFile(path),
            normalizedPackageHash,
            ReadDefinedNames(archive),
            package,
            sheets,
            BuildCashFlowSemantics(sheets.FirstOrDefault(sheet => sheet.Name == "Cash Flow")));
    }

    private static SemanticSheetSnapshot? BuildCashFlowSemantics(WorksheetSnapshot? sheet)
    {
        if (sheet is null)
            return null;

        var cellsByPosition = sheet.Cells.ToDictionary(
            cell => (Row: GetRowNumber(cell.Address), Column: GetColumnNumber(cell.Address)));
        var maxColumn = cellsByPosition.Keys.Select(position => position.Column).DefaultIfEmpty(0).Max();
        var periods = new List<SemanticPeriodSnapshot>();
        string? currentYear = null;
        string? currentStatus = null;
        for (var column = 4; column <= maxColumn; column++)
        {
            if (cellsByPosition.TryGetValue((3, column), out var yearCell))
            {
                var match = Regex.Match(yearCell.DisplayValue, @"^(?<year>\d{4}-\d{2})(?:\s+\((?<status>[^)]+)\))?$");
                if (match.Success)
                {
                    currentYear = match.Groups["year"].Value;
                    if (match.Groups["status"].Success)
                        currentStatus = match.Groups["status"].Value;
                }
            }
            if (currentYear is null || !cellsByPosition.TryGetValue((4, column), out var periodCell) || string.IsNullOrWhiteSpace(periodCell.DisplayValue))
                continue;
            var label = periodCell.DisplayValue.Trim().ToUpperInvariant();
            periods.Add(new SemanticPeriodSnapshot($"{currentYear}/{label}", column, currentYear, currentStatus, label));
        }

        var rows = new List<SemanticRowSnapshot>();
        var duplicateCounts = new Dictionary<string, int>(StringComparer.Ordinal);
        var sectionCounts = new Dictionary<string, int>(StringComparer.Ordinal);
        var section = "HEADER";
        var rowNumbers = sheet.Cells.Select(cell => GetRowNumber(cell.Address)).Where(row => row >= 5).Distinct().OrderBy(row => row);
        foreach (var rowNumber in rowNumbers)
        {
            var a = GetDisplay(1);
            var b = GetDisplay(2);
            var c = GetDisplay(3);
            var periodCells = periods
                .Select(period => (Period: period, Cell: cellsByPosition.GetValueOrDefault((rowNumber, period.Column))))
                .Where(item => item.Cell is not null)
                .ToArray();
            var hasPeriodContent = periodCells.Any(item =>
                !string.IsNullOrEmpty(item.Cell!.Value)
                || !string.IsNullOrEmpty(item.Cell.StoredValue)
                || !string.IsNullOrEmpty(item.Cell.Formula));

            if (!string.IsNullOrWhiteSpace(a) && string.IsNullOrWhiteSpace(b) && string.IsNullOrWhiteSpace(c) && !hasPeriodContent)
            {
                var sectionBase = NormalizeIdentity(a);
                sectionCounts.TryGetValue(sectionBase, out var sectionCount);
                sectionCounts[sectionBase] = ++sectionCount;
                section = sectionCount == 1 ? sectionBase : $"{sectionBase}|#{sectionCount}";
                continue;
            }

            if (string.IsNullOrWhiteSpace(a) && string.IsNullOrWhiteSpace(b) && string.IsNullOrWhiteSpace(c) && !hasPeriodContent)
                continue;

            var identity = !string.IsNullOrWhiteSpace(c)
                ? c
                : !string.IsNullOrWhiteSpace(a) && !string.IsNullOrWhiteSpace(b)
                    ? $"{a}|{b}"
                    : !string.IsNullOrWhiteSpace(a) ? a : b;
            var baseKey = $"{section}|{NormalizeIdentity(identity)}";
            duplicateCounts.TryGetValue(baseKey, out var count);
            duplicateCounts[baseKey] = ++count;
            var key = count == 1 ? baseKey : $"{baseKey}|#{count}";
            var kind = !string.IsNullOrWhiteSpace(c)
                ? (a.Equals("TOTALS", StringComparison.OrdinalIgnoreCase) ? "category-total" : "coded")
                : !string.IsNullOrWhiteSpace(a) && !string.IsNullOrWhiteSpace(b) ? "detail" : "label";
            rows.Add(new SemanticRowSnapshot(
                key,
                section,
                kind,
                a,
                b,
                c,
                rowNumber,
                periodCells.Select(item => new SemanticCellSnapshot(
                    item.Period.Key,
                    item.Cell!.DataType,
                    item.Cell.Value,
                    item.Cell.StoredValue,
                    item.Cell.DisplayValue,
                    item.Cell.Formula,
                    NormalizeFormula(item.Cell.Formula),
                    item.Cell.EffectiveStyleSignature)).ToArray()));

            string GetDisplay(int column) => cellsByPosition.GetValueOrDefault((rowNumber, column))?.DisplayValue?.Trim() ?? string.Empty;
        }

        return new SemanticSheetSnapshot(periods, rows);
    }

    private static int GetRowNumber(string address) =>
        int.Parse(Regex.Match(address, @"\d+$").Value, CultureInfo.InvariantCulture);

    private static int GetColumnNumber(string address)
    {
        var letters = Regex.Match(address, @"^[A-Z]+").Value;
        var result = 0;
        foreach (var character in letters)
            result = checked(result * 26 + character - 'A' + 1);
        return result;
    }

    private static string NormalizeIdentity(string value) =>
        Regex.Replace(value.Trim().ToUpperInvariant(), @"\s+", " ");

    private static string? NormalizeFormula(string? formula) =>
        formula is null
            ? null
            : Regex.Replace(formula, @"SUM\(([A-Z]+[0-9]+):\1\)", "SUM($1)", RegexOptions.CultureInvariant);

    private static string FormatCellValue(IXLCell cell)
    {
        if (cell.IsEmpty())
            return string.Empty;

        return cell.DataType switch
        {
            XLDataType.Boolean => cell.GetBoolean() ? "true" : "false",
            XLDataType.DateTime => cell.GetDateTime().ToString("O", CultureInfo.InvariantCulture),
            XLDataType.Number => cell.GetDouble().ToString("R", CultureInfo.InvariantCulture),
            XLDataType.TimeSpan => cell.GetTimeSpan().ToString("c", CultureInfo.InvariantCulture),
            _ => cell.Value.ToString(CultureInfo.InvariantCulture)
        };
    }

    private static string FormatEffectiveStyle(IXLStyle style)
    {
        var numberFormat = Regex.Replace(style.NumberFormat.Format ?? string.Empty, @"\\([()])", "$1").ToUpperInvariant();
        return string.Join("|",
            $"num:{style.NumberFormat.NumberFormatId}:{numberFormat}",
            $"font:{style.Font.FontName}:{style.Font.FontSize.ToString("R", CultureInfo.InvariantCulture)}:{style.Font.Bold}:{style.Font.Italic}:{style.Font.Underline}:{style.Font.Strikethrough}:{style.Font.VerticalAlignment}:{style.Font.FontColor}",
            $"fill:{style.Fill.PatternType}:{style.Fill.BackgroundColor}:{style.Fill.PatternColor}",
            $"border:L={style.Border.LeftBorder},{style.Border.LeftBorderColor};R={style.Border.RightBorder},{style.Border.RightBorderColor};T={style.Border.TopBorder},{style.Border.TopBorderColor};B={style.Border.BottomBorder},{style.Border.BottomBorderColor};D={style.Border.DiagonalBorder},{style.Border.DiagonalBorderColor},{style.Border.DiagonalUp},{style.Border.DiagonalDown}",
            $"align:{style.Alignment.Horizontal}:{style.Alignment.Vertical}:{style.Alignment.Indent}:{style.Alignment.ShrinkToFit}:{style.Alignment.TextRotation}:{style.Alignment.WrapText}",
            $"protect:{style.Protection.Locked}:{style.Protection.Hidden}");
    }

    private static IReadOnlyList<PackageEntrySnapshot> ReadPackage(ZipArchive archive)
    {
        var result = new List<PackageEntrySnapshot>();
        foreach (var entry in archive.Entries.OrderBy(entry => entry.FullName, StringComparer.Ordinal))
        {
            var bytes = ReadEntryBytes(entry);
            var rawHash = HashBytes(bytes);
            var normalizedHash = rawHash;
            if (IsCoreProperties(entry.FullName))
            {
                var xml = ParseXml(bytes);
                xml.Root?.Element(DublinCoreTerms + "created")?.Remove();
                xml.Root?.Element(DublinCoreTerms + "modified")?.Remove();
                xml.Root?.Element(CoreProperties + "lastModifiedBy")?.Remove();
                normalizedHash = HashText(CanonicalXml(xml));
            }
            else if (entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                     || entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
            {
                var xml = ParseXml(bytes);
                NormalizePackageReferences(entry.FullName, xml);
                normalizedHash = HashText(CanonicalXml(xml));
            }

            result.Add(new PackageEntrySnapshot(entry.FullName, NormalizePackagePath(entry.FullName), entry.Length, rawHash, normalizedHash));
        }
        return result;
    }

    private static Dictionary<string, SheetPartSnapshot> ReadSheetParts(ZipArchive archive)
    {
        var workbook = ReadXml(archive, "xl/workbook.xml");
        var relationships = ReadXml(archive, "xl/_rels/workbook.xml.rels");
        var targets = relationships.Root!
            .Elements(PackageRelationships + "Relationship")
            .ToDictionary(
                element => (string)element.Attribute("Id")!,
                element => NormalizePartPath("xl", (string)element.Attribute("Target")!),
                StringComparer.Ordinal);

        var result = new Dictionary<string, SheetPartSnapshot>(StringComparer.Ordinal);
        foreach (var sheet in workbook.Descendants(Main + "sheet"))
        {
            var name = (string)sheet.Attribute("name")!;
            var relationshipId = (string)sheet.Attribute(Relationships + "id")!;
            if (!targets.TryGetValue(relationshipId, out var target))
                continue;
            result[name] = ReadSheetPart(archive, target);
        }
        return result;
    }

    private static SheetPartSnapshot ReadSheetPart(ZipArchive archive, string path)
    {
        var xml = ReadXml(archive, path);
        var pane = xml.Descendants(Main + "pane").FirstOrDefault();
        var freeze = pane is null ? null : CanonicalElement(pane);
        var protection = xml.Root?.Element(Main + "sheetProtection") is { } sheetProtection
            ? CanonicalElement(sheetProtection)
            : null;
        var printNodes = xml.Root?.Elements()
            .Where(element => element.Name == Main + "printOptions"
                              || element.Name == Main + "pageMargins"
                              || element.Name == Main + "pageSetup"
                              || element.Name == Main + "headerFooter")
            .Select(CanonicalElement)
            .ToArray() ?? [];
        var merges = xml.Descendants(Main + "mergeCell")
            .Select(element => (string)element.Attribute("ref")!)
            .OrderBy(value => value, StringComparer.Ordinal)
            .ToArray();
        var rows = xml.Descendants(Main + "row")
            .Where(element => element.Attribute("customHeight") is not null || element.Attribute("hidden") is not null)
            .Select(CanonicalElement)
            .ToArray();
        var columns = xml.Descendants(Main + "col")
            .Select(CanonicalElement)
            .ToArray();
        return new SheetPartSnapshot(path, freeze, protection, string.Join("", printNodes), merges, rows, columns);
    }

    private static Dictionary<string, CellPartSnapshot> ReadCellParts(ZipArchive archive, string path)
    {
        var xml = ReadXml(archive, path);
        return xml.Descendants(Main + "c")
            .Where(cell => cell.Attribute("r") is not null)
            .ToDictionary(
                cell => (string)cell.Attribute("r")!,
                cell => new CellPartSnapshot(
                    (int?)cell.Attribute("s") ?? 0,
                    cell.Element(Main + "v")?.Value),
                StringComparer.Ordinal);
    }

    private static Dictionary<int, string> ReadStyles(ZipArchive archive)
    {
        var xml = ReadXml(archive, "xl/styles.xml");
        var root = xml.Root!;
        var numberFormats = root.Element(Main + "numFmts")?.Elements(Main + "numFmt")
            .Where(element => element.Attribute("numFmtId") is not null)
            .ToDictionary(
                element => (int)element.Attribute("numFmtId")!,
                element => (string?)element.Attribute("formatCode") ?? string.Empty) ?? [];
        var fonts = root.Element(Main + "fonts")?.Elements(Main + "font").ToArray() ?? [];
        var fills = root.Element(Main + "fills")?.Elements(Main + "fill").ToArray() ?? [];
        var borders = root.Element(Main + "borders")?.Elements(Main + "border").ToArray() ?? [];
        var xfs = root.Element(Main + "cellXfs")?.Elements(Main + "xf").ToArray() ?? [];
        var result = new Dictionary<int, string>();

        for (var index = 0; index < xfs.Length; index++)
        {
            var xf = xfs[index];
            var numberFormatId = (int?)xf.Attribute("numFmtId") ?? 0;
            var fontId = (int?)xf.Attribute("fontId") ?? 0;
            var fillId = (int?)xf.Attribute("fillId") ?? 0;
            var borderId = (int?)xf.Attribute("borderId") ?? 0;
            var signature = new XElement("style",
                new XAttribute("numberFormatId", numberFormatId),
                new XAttribute("numberFormatCode", numberFormats.GetValueOrDefault(numberFormatId, string.Empty)),
                new XElement("xf", xf.Attributes(), xf.Elements()),
                fontId < fonts.Length ? new XElement("font", fonts[fontId].Attributes(), fonts[fontId].Nodes()) : null,
                fillId < fills.Length ? new XElement("fill", fills[fillId].Attributes(), fills[fillId].Nodes()) : null,
                borderId < borders.Length ? new XElement("border", borders[borderId].Attributes(), borders[borderId].Nodes()) : null);
            result[index] = CanonicalElement(signature);
        }
        return result;
    }

    private static IReadOnlyList<DefinedNameSnapshot> ReadDefinedNames(ZipArchive archive)
    {
        var xml = ReadXml(archive, "xl/workbook.xml");
        return xml.Descendants(Main + "definedName")
            .Select(element => new DefinedNameSnapshot(
                (string?)element.Attribute("name") ?? string.Empty,
                (int?)element.Attribute("localSheetId"),
                element.Value))
            .OrderBy(item => item.Name, StringComparer.Ordinal)
            .ThenBy(item => item.LocalSheetId)
            .ToArray();
    }

    private static string NormalizePartPath(string basePath, string target)
    {
        var combined = target.StartsWith('/') ? target[1..] : $"{basePath}/{target}";
        var parts = new List<string>();
        foreach (var part in combined.Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            if (part == "..")
                parts.RemoveAt(parts.Count - 1);
            else if (part != ".")
                parts.Add(part);
        }
        return string.Join('/', parts);
    }

    private static XDocument ReadXml(ZipArchive archive, string path)
    {
        var entry = archive.GetEntry(path) ?? throw new InvalidDataException($"Missing XLSX part: {path}");
        using var stream = entry.Open();
        return XDocument.Load(stream, System.Xml.Linq.LoadOptions.PreserveWhitespace);
    }

    private static byte[] ReadEntryBytes(ZipArchiveEntry entry)
    {
        using var input = entry.Open();
        using var output = new MemoryStream();
        input.CopyTo(output);
        return output.ToArray();
    }

    private static XDocument ParseXml(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        return XDocument.Load(stream, System.Xml.Linq.LoadOptions.PreserveWhitespace);
    }

    private static bool IsCoreProperties(string path) =>
        path.Equals("docProps/core.xml", StringComparison.Ordinal)
        || path.Contains("/metadata/core-properties/", StringComparison.OrdinalIgnoreCase);

    private static string NormalizePackagePath(string path) =>
        IsCoreProperties(path) ? "{core-properties}" : path;

    private static void NormalizePackageReferences(string path, XDocument document)
    {
        if (path.Equals("[Content_Types].xml", StringComparison.Ordinal))
        {
            foreach (var attribute in document.Descendants().Attributes("PartName")
                         .Where(attribute => IsCoreProperties(attribute.Value.TrimStart('/'))))
                attribute.Value = "/{core-properties}";
        }

        if (path.Equals("_rels/.rels", StringComparison.Ordinal))
        {
            foreach (var relationship in document.Descendants()
                         .Where(element => ((string?)element.Attribute("Type"))?.EndsWith("/metadata/core-properties", StringComparison.Ordinal) == true))
                relationship.SetAttributeValue("Target", "{core-properties}");
        }
    }

    private static string CanonicalXml(XDocument document) =>
        document.Root is null ? string.Empty : CanonicalElement(document.Root);

    private static string CanonicalElement(XElement element)
    {
        var normalized = new XElement(
            element.Name,
            element.Attributes()
                .Where(attribute => !attribute.IsNamespaceDeclaration)
                .OrderBy(attribute => attribute.Name.NamespaceName, StringComparer.Ordinal)
                .ThenBy(attribute => attribute.Name.LocalName, StringComparer.Ordinal)
                .Select(attribute => new XAttribute(attribute.Name, attribute.Value)),
            element.Nodes().Select(node => node is XElement child
                ? XElement.Parse(CanonicalElement(child), System.Xml.Linq.LoadOptions.PreserveWhitespace)
                : node is XText text ? new XText(text.Value) : node));
        return normalized.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
    }

    private static string HashFile(string path)
    {
        using var stream = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(stream));
    }

    private static string HashBytes(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes));
    private static string HashText(string value) => HashBytes(Encoding.UTF8.GetBytes(value));
}

internal static class WorkbookComparer
{
    public static ComparisonReport Compare(WorkbookSnapshot expected, WorkbookSnapshot actual)
    {
        var differences = new List<WorkbookDifference>();
        var categoryCounts = new Dictionary<string, int>(StringComparer.Ordinal);
        var sampleCounts = new Dictionary<string, int>(StringComparer.Ordinal);
        var totalDifferenceCount = 0;
        if (expected.Worksheets.Count != actual.Worksheets.Count)
            Add("Structure", "Workbook", $"Worksheet count {expected.Worksheets.Count} != {actual.Worksheets.Count}");

        var sheetCount = Math.Min(expected.Worksheets.Count, actual.Worksheets.Count);
        for (var index = 0; index < sheetCount; index++)
        {
            var left = expected.Worksheets[index];
            var right = actual.Worksheets[index];
            var location = $"Sheet[{index + 1}]";
            CompareField("Structure", location, "name", left.Name, right.Name);
            CompareField("Structure", location, "visibility", left.Visibility, right.Visibility);
            CompareField("Structure", location, "used range", left.UsedRange, right.UsedRange);
            CompareField("Structure", location, "freeze pane", left.FreezePane, right.FreezePane);
            CompareField("Structure", location, "protection", left.Protection, right.Protection);
            CompareField("Structure", location, "print settings", left.PrintSettings, right.PrintSettings);
            CompareSequence("Structure", location, "merged ranges", left.MergedRanges, right.MergedRanges);
            CompareSequence("Structure", location, "row settings", left.Rows, right.Rows);
            CompareSequence("Structure", location, "column settings", left.Columns, right.Columns);

            var leftCells = left.Cells.ToDictionary(cell => cell.Address, StringComparer.Ordinal);
            var rightCells = right.Cells.ToDictionary(cell => cell.Address, StringComparer.Ordinal);
            foreach (var address in leftCells.Keys.Union(rightCells.Keys, StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal))
            {
                if (!leftCells.TryGetValue(address, out var leftCell) || !rightCells.TryGetValue(address, out var rightCell))
                {
                    Add("Structure", $"{left.Name}!{address}", "Cell exists on only one side");
                    continue;
                }
                CompareField("Value", $"{left.Name}!{address}", "data type", leftCell.DataType, rightCell.DataType);
                CompareField("Value", $"{left.Name}!{address}", "value", leftCell.Value, rightCell.Value);
                CompareField("Value", $"{left.Name}!{address}", "stored value", leftCell.StoredValue, rightCell.StoredValue);
                CompareField("Value", $"{left.Name}!{address}", "display value", leftCell.DisplayValue, rightCell.DisplayValue);
                CompareField("Formula", $"{left.Name}!{address}", "formula", leftCell.Formula, rightCell.Formula);
                CompareField("Style", $"{left.Name}!{address}", "effective style", leftCell.EffectiveStyleSignature, rightCell.EffectiveStyleSignature);
            }
        }

        CompareSequence("Structure", "Workbook", "defined names", expected.DefinedNames, actual.DefinedNames);
        CompareSemantics(expected.CashFlowSemantics, actual.CashFlowSemantics);
        var expectedParts = expected.PackageEntries.ToDictionary(entry => entry.NormalizedPath, StringComparer.Ordinal);
        var actualParts = actual.PackageEntries.ToDictionary(entry => entry.NormalizedPath, StringComparer.Ordinal);
        foreach (var path in expectedParts.Keys.Union(actualParts.Keys, StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal))
        {
            if (!expectedParts.TryGetValue(path, out var left) || !actualParts.TryGetValue(path, out var right))
            {
                Add("Package", path, "Package part exists on only one side");
                continue;
            }
            if (!left.NormalizedSha256.Equals(right.NormalizedSha256, StringComparison.Ordinal))
                Add(path.Equals("{core-properties}", StringComparison.Ordinal) ? "Metadata" : "Package", path, "Normalized part hash differs");
        }

        var semanticDefinitionEquivalent = !HasAny("SemanticStructure", "SemanticInputValue", "SemanticFormula", "SemanticStyle");
        var semanticResultEquivalent = semanticDefinitionEquivalent && !HasAny("SemanticFormulaValue");
        return new ComparisonReport(
            totalDifferenceCount == 0,
            semanticDefinitionEquivalent,
            semanticResultEquivalent,
            totalDifferenceCount,
            categoryCounts,
            differences);

        bool HasAny(params string[] categories) => categories.Any(category => categoryCounts.GetValueOrDefault(category) > 0);

        void CompareSemantics(SemanticSheetSnapshot? left, SemanticSheetSnapshot? right)
        {
            if (left is null || right is null)
            {
                if (left != right)
                    Add("SemanticStructure", "Cash Flow", "Semantic sheet exists on only one side");
                return;
            }
            CompareSequence("SemanticStructure", "Cash Flow", "period identities", left.Periods.Select(period => period.Key).ToArray(), right.Periods.Select(period => period.Key).ToArray());
            var leftRows = left.Rows.ToDictionary(row => row.Key, StringComparer.Ordinal);
            var rightRows = right.Rows.ToDictionary(row => row.Key, StringComparer.Ordinal);
            foreach (var key in leftRows.Keys.Union(rightRows.Keys, StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal))
            {
                if (!leftRows.TryGetValue(key, out var leftRow) || !rightRows.TryGetValue(key, out var rightRow))
                {
                    Add("SemanticStructure", key, "Semantic row exists on only one side");
                    continue;
                }
                var leftCells = leftRow.Cells.ToDictionary(cell => cell.PeriodKey, StringComparer.Ordinal);
                var rightCells = rightRow.Cells.ToDictionary(cell => cell.PeriodKey, StringComparer.Ordinal);
                foreach (var period in leftCells.Keys.Union(rightCells.Keys, StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal))
                {
                    if (!leftCells.TryGetValue(period, out var leftCell) || !rightCells.TryGetValue(period, out var rightCell))
                    {
                        Add("SemanticStructure", $"{key}/{period}", "Semantic cell exists on only one side");
                        continue;
                    }
                    var valueCategory = leftCell.NormalizedFormula is not null || rightCell.NormalizedFormula is not null
                        ? "SemanticFormulaValue"
                        : "SemanticInputValue";
                    CompareField(valueCategory, $"{key}/{period}", "data type", leftCell.DataType, rightCell.DataType);
                    CompareField(valueCategory, $"{key}/{period}", "value", leftCell.Value, rightCell.Value);
                    CompareField(valueCategory, $"{key}/{period}", "stored value", leftCell.StoredValue, rightCell.StoredValue);
                    CompareField(valueCategory, $"{key}/{period}", "display value", leftCell.DisplayValue, rightCell.DisplayValue);
                    CompareField("SemanticFormula", $"{key}/{period}", "normalized formula", leftCell.NormalizedFormula, rightCell.NormalizedFormula);
                    CompareField("SemanticStyle", $"{key}/{period}", "style", leftCell.StyleSignature, rightCell.StyleSignature);
                }
            }
        }

        void CompareField(string category, string location, string field, string? left, string? right)
        {
            if (!string.Equals(left, right, StringComparison.Ordinal))
                Add(category, location, $"{field} differs");
        }

        void CompareSequence<T>(string category, string location, string field, IReadOnlyList<T> left, IReadOnlyList<T> right)
        {
            if (!left.SequenceEqual(right))
                Add(category, location, $"{field} differ");
        }

        void Add(string category, string location, string message)
        {
            totalDifferenceCount++;
            categoryCounts[category] = categoryCounts.GetValueOrDefault(category) + 1;
            if (sampleCounts.GetValueOrDefault(category) < 25)
            {
                differences.Add(new WorkbookDifference(category, location, message));
                sampleCounts[category] = sampleCounts.GetValueOrDefault(category) + 1;
            }
        }
    }
}

internal static class MutationProof
{
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace DublinCore = "http://purl.org/dc/elements/1.1/";

    public static IReadOnlyList<MutationProofResult> Run(string baseline, string outputDirectory)
    {
        var expected = WorkbookInspector.Inspect(baseline);
        var corePart = expected.PackageEntries.First(entry => entry.NormalizedPath == "{core-properties}").Path;
        var cases = new (string Name, string Part, Func<XDocument, bool> Mutate, string ExpectedCategory)[]
        {
            ("value", "xl/worksheets/sheet1.xml", MutateValue, "Value"),
            ("formula", "xl/worksheets/sheet1.xml", MutateFormula, "Formula"),
            ("style", "xl/worksheets/sheet1.xml", MutateStyle, "Style"),
            ("structure", "xl/workbook.xml", MutateStructure, "Structure"),
            ("metadata", corePart, MutateMetadata, "Metadata")
        };
        var results = new List<MutationProofResult>();
        foreach (var test in cases)
        {
            var output = Path.Combine(outputDirectory, $"mutation-{test.Name}.xlsx");
            MutatePackage(baseline, output, test.Part, test.Mutate);
            var report = WorkbookComparer.Compare(expected, WorkbookInspector.Inspect(output));
            var categories = report.Differences.Select(difference => difference.Category).Distinct(StringComparer.Ordinal).OrderBy(value => value).ToArray();
            results.Add(new MutationProofResult(test.Name, report.CategoryCounts.ContainsKey(test.ExpectedCategory), report.CategoryCounts.Keys.OrderBy(value => value).ToArray(), report.TotalDifferenceCount));
        }
        if (expected.CashFlowSemantics is not null)
        {
            var reordered = expected with
            {
                CashFlowSemantics = expected.CashFlowSemantics with
                {
                    Rows = expected.CashFlowSemantics.Rows.Reverse().ToArray()
                }
            };
            var report = WorkbookComparer.Compare(expected, reordered);
            results.Add(new MutationProofResult(
                "semantic-row-order-independence",
                report.Equivalent,
                report.CategoryCounts.Keys.OrderBy(value => value).ToArray(),
                report.TotalDifferenceCount));
        }
        return results;
    }

    private static void MutatePackage(string source, string destination, string targetPart, Func<XDocument, bool> mutate)
    {
        if (File.Exists(destination))
            File.Delete(destination);
        using var input = ZipFile.OpenRead(source);
        using var output = ZipFile.Open(destination, ZipArchiveMode.Create);
        var changed = false;
        foreach (var inputEntry in input.Entries)
        {
            var outputEntry = output.CreateEntry(inputEntry.FullName, CompressionLevel.Optimal);
            outputEntry.LastWriteTime = inputEntry.LastWriteTime;
            using var inputStream = inputEntry.Open();
            using var outputStream = outputEntry.Open();
            if (inputEntry.FullName.Equals(targetPart, StringComparison.Ordinal))
            {
                var document = XDocument.Load(inputStream, System.Xml.Linq.LoadOptions.PreserveWhitespace);
                changed = mutate(document);
                document.Save(outputStream, System.Xml.Linq.SaveOptions.DisableFormatting);
            }
            else
            {
                inputStream.CopyTo(outputStream);
            }
        }
        if (!changed)
            throw new InvalidDataException($"Mutation did not change {targetPart}");
    }

    private static bool MutateValue(XDocument document)
    {
        var cell = document.Descendants(Main + "c").FirstOrDefault(candidate =>
            candidate.Element(Main + "f") is null
            && ((string?)candidate.Attribute("t") is null or "n")
            && double.TryParse(candidate.Element(Main + "v")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out _));
        if (cell?.Element(Main + "v") is not { } value || !double.TryParse(value.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
            return false;
        value.Value = (number + 1).ToString("R", CultureInfo.InvariantCulture);
        return true;
    }

    private static bool MutateFormula(XDocument document)
    {
        var formula = document.Descendants(Main + "f").FirstOrDefault(element => !string.IsNullOrWhiteSpace(element.Value));
        if (formula is null)
            return false;
        formula.Value = $"({formula.Value})+0";
        return true;
    }

    private static bool MutateStyle(XDocument document)
    {
        var cell = document.Descendants(Main + "c").FirstOrDefault(candidate => candidate.Attribute("s") is not null);
        if (cell?.Attribute("s") is not { } style || !int.TryParse(style.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value))
            return false;
        style.Value = value == 0 ? "1" : "0";
        return true;
    }

    private static bool MutateStructure(XDocument document)
    {
        var sheet = document.Descendants(Main + "sheet").FirstOrDefault();
        if (sheet?.Attribute("name") is not { } name)
            return false;
        name.Value += " MUTATED";
        return true;
    }

    private static bool MutateMetadata(XDocument document)
    {
        var title = document.Root?.Element(DublinCore + "title");
        if (title is null)
            document.Root?.Add(new XElement(DublinCore + "title", "Mutation proof"));
        else
            title.Value = "Mutation proof";
        return document.Root is not null;
    }
}

internal sealed record WorkbookSnapshot(
    string FileName,
    long FileLength,
    string RawSha256,
    string NormalizedPackageSha256,
    IReadOnlyList<DefinedNameSnapshot> DefinedNames,
    IReadOnlyList<PackageEntrySnapshot> PackageEntries,
    IReadOnlyList<WorksheetSnapshot> Worksheets,
    SemanticSheetSnapshot? CashFlowSemantics);

internal sealed record WorksheetSnapshot(
    int Position,
    string Name,
    string Visibility,
    string? UsedRange,
    string? FreezePane,
    string? Protection,
    string PrintSettings,
    IReadOnlyList<string> MergedRanges,
    IReadOnlyList<string> Rows,
    IReadOnlyList<string> Columns,
    IReadOnlyList<CellSnapshot> Cells);

internal sealed record CellSnapshot(
    string Address,
    string DataType,
    string Value,
    string? StoredValue,
    string DisplayValue,
    string? Formula,
    int StyleIndex,
    string? StyleSignature,
    string EffectiveStyleSignature);

internal sealed record PackageEntrySnapshot(string Path, string NormalizedPath, long Length, string RawSha256, string NormalizedSha256);
internal sealed record DefinedNameSnapshot(string Name, int? LocalSheetId, string Reference);
internal sealed record WorkbookIndexEntry(
    string FileName,
    long FileLength,
    string RawSha256,
    string NormalizedPackageSha256,
    int WorksheetCount,
    int CellCount,
    int FormulaCount,
    string SnapshotFile);

internal sealed record ComparisonReport(
    bool Equivalent,
    bool SemanticDefinitionEquivalent,
    bool SemanticResultEquivalent,
    int TotalDifferenceCount,
    IReadOnlyDictionary<string, int> CategoryCounts,
    IReadOnlyList<WorkbookDifference> Differences);
internal sealed record WorkbookDifference(string Category, string Location, string Message);
internal sealed record MutationProofResult(
    string Mutation,
    bool Passed,
    IReadOnlyList<string> DetectedCategories,
    int DifferenceCount);

internal sealed record CellPartSnapshot(int StyleIndex, string? StoredValue);

internal sealed record SemanticSheetSnapshot(
    IReadOnlyList<SemanticPeriodSnapshot> Periods,
    IReadOnlyList<SemanticRowSnapshot> Rows);

internal sealed record SemanticPeriodSnapshot(string Key, int Column, string FinancialYear, string? Status, string Label);

internal sealed record SemanticRowSnapshot(
    string Key,
    string Section,
    string Kind,
    string Code,
    string Name,
    string Marker,
    int PhysicalRow,
    IReadOnlyList<SemanticCellSnapshot> Cells);

internal sealed record SemanticCellSnapshot(
    string PeriodKey,
    string DataType,
    string Value,
    string? StoredValue,
    string DisplayValue,
    string? Formula,
    string? NormalizedFormula,
    string? StyleSignature);

internal sealed record SheetPartSnapshot(
    string Path,
    string? FreezePane,
    string? Protection,
    string PrintSettings,
    IReadOnlyList<string> MergedRanges,
    IReadOnlyList<string> Rows,
    IReadOnlyList<string> Columns);


