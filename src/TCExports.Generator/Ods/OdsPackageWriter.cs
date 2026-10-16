using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using TCExports.Generator.CashStatement.Model;

namespace TCExports.Generator.Ods;

public sealed class OdsPackageWriter
{
    public const string MimeType = "application/vnd.oasis.opendocument.spreadsheet";
    private static readonly DateTimeOffset StableTimestamp = new(2000, 1, 1, 0, 0, 0, TimeSpan.Zero);

    private static readonly XNamespace Office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    private static readonly XNamespace Table = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
    private static readonly XNamespace Text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
    private static readonly XNamespace Style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
    private static readonly XNamespace Number = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";
    private static readonly XNamespace Fo = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
    private static readonly XNamespace Config = "urn:oasis:names:tc:opendocument:xmlns:config:1.0";
    private static readonly XNamespace Ooo = "http://openoffice.org/2004/office";
    private static readonly XNamespace XLink = "http://www.w3.org/1999/xlink";
    private static readonly XNamespace Manifest = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
    private static readonly XNamespace Meta = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0";
    private static readonly XNamespace Of = "urn:oasis:names:tc:opendocument:xmlns:of:1.2";

    public async Task WriteAsync(
        Stream destination,
        OdsSpreadsheetDocument document,
        OdsWriteLimits? limits = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(destination);
        ArgumentNullException.ThrowIfNull(document);
        if (!destination.CanWrite)
            throw new ArgumentException("The ODS destination stream must be writable.", nameof(destination));

        limits ??= OdsWriteLimits.Default;
        Validate(document, limits);
        cancellationToken.ThrowIfCancellationRequested();

        using var buffer = new MemoryStream();
        using (var limited = new LimitedWriteStream(buffer, limits.MaximumOutputBytes, leaveOpen: true))
        using (var archive = new ZipArchive(limited, ZipArchiveMode.Create, leaveOpen: true, Encoding.UTF8))
        {
            AddTextEntry(archive, "mimetype", MimeType, CompressionLevel.NoCompression);
            AddXmlEntry(archive, "content.xml", CreateContent(document, cancellationToken));
            AddXmlEntry(archive, "styles.xml", CreateStyles());
            AddXmlEntry(archive, "settings.xml", CreateSettings(document));
            AddXmlEntry(archive, "meta.xml", CreateMeta(document.Generator));
            AddXmlEntry(archive, "META-INF/manifest.xml", CreateManifest());
        }

        cancellationToken.ThrowIfCancellationRequested();
        buffer.Position = 0;
        await buffer.CopyToAsync(destination, 81920, cancellationToken).ConfigureAwait(false);
    }

    private static void Validate(OdsSpreadsheetDocument document, OdsWriteLimits limits)
    {
        if (limits.MaximumWorksheets < 1 || limits.MaximumColumnsPerWorksheet < 1 || limits.MaximumRows < 1 ||
            limits.MaximumCells < 1 || limits.MaximumOutputBytes < 1)
            throw new ArgumentOutOfRangeException(nameof(limits), "All ODS output limits must be positive.");

        if (document.Worksheets.IsDefaultOrEmpty)
            throw new InvalidDataException("An ODS document requires at least one worksheet.");
        if (document.Worksheets.Length > limits.MaximumWorksheets)
            throw new InvalidDataException($"ODS worksheet limit exceeded ({limits.MaximumWorksheets}).");

        var names = new HashSet<string>(StringComparer.Ordinal);
        var addresses = new HashSet<StatementCellAddress>();
        var rowCount = 0;
        var cellCount = 0;
        foreach (var worksheet in document.Worksheets)
        {
            if (string.IsNullOrWhiteSpace(worksheet.Name) || !names.Add(worksheet.Name))
                throw new InvalidDataException($"ODS worksheet names must be non-empty and unique: '{worksheet.Name}'.");
            if (worksheet.Columns.IsDefaultOrEmpty)
                throw new InvalidDataException($"ODS worksheet '{worksheet.Name}' requires at least one column.");
            if (worksheet.Columns.Length > limits.MaximumColumnsPerWorksheet)
                throw new InvalidDataException($"ODS column limit exceeded ({limits.MaximumColumnsPerWorksheet}).");
            if (worksheet.Columns.Any(column => column.WidthCentimetres <= 0))
                throw new InvalidDataException("ODS column widths must be positive.");
            if (worksheet.FrozenColumns < 0 || worksheet.FrozenRows < 0)
                throw new InvalidDataException("Frozen row and column counts cannot be negative.");

            rowCount += worksheet.Rows.Length;
            foreach (var row in worksheet.Rows)
            {
                cellCount += row.Cells.Sum(cell => cell.ColumnSpan);
                if (row.Cells.Any(cell => cell.ColumnSpan < 1))
                    throw new InvalidDataException("ODS cell column spans must be positive.");
                if (row.Cells.Sum(cell => cell.ColumnSpan) > worksheet.Columns.Length)
                    throw new InvalidDataException($"ODS row exceeds the column allocation in '{worksheet.Name}'.");
                foreach (var cell in row.Cells)
                    if (cell.SemanticAddress is { } address && !addresses.Add(address))
                        throw new InvalidDataException($"Duplicate semantic ODS cell address: '{address}'.");
            }
        }

        if (rowCount > limits.MaximumRows)
            throw new InvalidDataException($"ODS row limit exceeded ({limits.MaximumRows}).");
        if (cellCount > limits.MaximumCells)
            throw new InvalidDataException($"ODS cell limit exceeded ({limits.MaximumCells}).");
    }

    private static XDocument CreateContent(OdsSpreadsheetDocument document, CancellationToken cancellationToken)
    {
        var addresses = new Dictionary<StatementCellAddress, OdsCellReference>();
        foreach (var worksheet in document.Worksheets)
        {
            for (var rowIndex = 0; rowIndex < worksheet.Rows.Length; rowIndex++)
            {
                var columnIndex = 1;
                foreach (var cell in worksheet.Rows[rowIndex].Cells)
                {
                    if (cell.SemanticAddress is { } address)
                        addresses.Add(address, new OdsCellReference(worksheet.Name, rowIndex + 1, columnIndex));
                    columnIndex += cell.ColumnSpan;
                }
            }
        }

        var translator = new CashStatementOdsFormulaTranslator(addresses);
        var cellStyles = BuildCellStyles(document);
        var spreadsheet = new XElement(Office + "spreadsheet",
            new XElement(Table + "calculation-settings",
                new XAttribute(Table + "automatic-find-labels", "false"),
                new XAttribute(Table + "case-sensitive", "false"),
                new XAttribute(Table + "precision-as-shown", "false")));

        for (var sheetIndex = 0; sheetIndex < document.Worksheets.Length; sheetIndex++)
        {
            var worksheet = document.Worksheets[sheetIndex];
            cancellationToken.ThrowIfCancellationRequested();
            var table = new XElement(Table + "table",
                new XAttribute(Table + "name", worksheet.Name));
            if (worksheet.Protected)
                table.Add(new XAttribute(Table + "protected", "true"));
            if (!worksheet.Print)
                table.Add(new XAttribute(Table + "print", "false"));
            for (var index = 0; index < worksheet.Columns.Length; index++)
            {
                var column = worksheet.Columns[index];
                var element = new XElement(Table + "table-column",
                    new XAttribute(Table + "style-name", ColumnStyleName(sheetIndex, index)));
                if (column.Hidden)
                    element.Add(new XAttribute(Table + "visibility", "collapse"));
                table.Add(element);
            }

            foreach (var row in worksheet.Rows)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var rowElement = new XElement(Table + "table-row");
                foreach (var cell in row.Cells)
                {
                    rowElement.Add(CreateCell(cell, translator, cellStyles[CellStyleKey(cell)]));
                    for (var covered = 1; covered < cell.ColumnSpan; covered++)
                        rowElement.Add(new XElement(Table + "covered-table-cell"));
                }
                table.Add(rowElement);
            }

            spreadsheet.Add(table);
        }

        var automaticStyles = new XElement(Office + "automatic-styles");
        for (var sheetIndex = 0; sheetIndex < document.Worksheets.Length; sheetIndex++)
        {
            var worksheet = document.Worksheets[sheetIndex];
            for (var columnIndex = 0; columnIndex < worksheet.Columns.Length; columnIndex++)
            {
                automaticStyles.Add(new XElement(Style + "style",
                    new XAttribute(Style + "name", ColumnStyleName(sheetIndex, columnIndex)),
                    new XAttribute(Style + "family", "table-column"),
                    new XElement(Style + "table-column-properties",
                        new XAttribute(Style + "column-width",
                            $"{worksheet.Columns[columnIndex].WidthCentimetres.ToString("0.###", CultureInfo.InvariantCulture)}cm"))));
            }
        }
        foreach (var pair in cellStyles)
            automaticStyles.Add(DynamicCellStyle(pair.Value, pair.Key));

        return Document(Office + "document-content",
            automaticStyles,
            new XElement(Office + "body", spreadsheet));
    }

    private static XElement CreateCell(
        OdsCell cell,
        CashStatementOdsFormulaTranslator translator,
        string styleName)
    {
        var element = new XElement(Table + "table-cell",
            new XAttribute(Table + "style-name", styleName));
        if (cell.ColumnSpan > 1)
            element.Add(new XAttribute(Table + "number-columns-spanned", cell.ColumnSpan));
        if (cell.Formula is not null)
            element.Add(new XAttribute(Table + "formula", translator.Translate(cell.Formula)));
        if (cell.Formula is not null &&
            cell.Value is { State: StatementValueState.Unsupported, Reason: "spreadsheet-engine-evaluation-required" })
            return element;

        switch (cell.Value)
        {
            case { State: StatementValueState.Missing }:
                return element;
            case { State: StatementValueState.Unsupported }:
                AddString(element, cell.Value.Reason ?? "Unsupported value");
                return element;
            case { Kind: StatementScalarKind.Decimal, DecimalValue: not null }:
                var number = cell.Value.DecimalValue.Value.ToString(CultureInfo.InvariantCulture);
                element.Add(new XAttribute(Office + "value-type", "float"));
                element.Add(new XAttribute(Office + "value", number));
                element.Add(new XElement(Text + "p", number));
                return element;
            case { Kind: StatementScalarKind.Text, TextValue: not null }:
                AddString(element, cell.Value.TextValue);
                return element;
            case { Kind: StatementScalarKind.Date, DateValue: not null }:
                var date = cell.Value.DateValue.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                element.Add(new XAttribute(Office + "value-type", "date"));
                element.Add(new XAttribute(Office + "date-value", date));
                element.Add(new XElement(Text + "p", date));
                return element;
            case { Kind: StatementScalarKind.Boolean, BooleanValue: not null }:
                var boolean = cell.Value.BooleanValue.Value ? "true" : "false";
                element.Add(new XAttribute(Office + "value-type", "boolean"));
                element.Add(new XAttribute(Office + "boolean-value", boolean));
                element.Add(new XElement(Text + "p", boolean.ToUpperInvariant()));
                return element;
            default:
                throw new InvalidDataException("Present ODS cell does not contain the declared scalar value.");
        }
    }

    private static void AddString(XElement element, string value)
    {
        element.Add(new XAttribute(Office + "value-type", "string"));
        element.Add(new XAttribute(Office + "string-value", value));
        element.Add(new XElement(Text + "p", value));
    }

    private static IReadOnlyDictionary<OdsCellStyleKey, string> BuildCellStyles(OdsSpreadsheetDocument document)
    {
        var styles = new Dictionary<OdsCellStyleKey, string>();
        foreach (var cell in document.Worksheets.SelectMany(sheet => sheet.Rows).SelectMany(row => row.Cells))
        {
            var key = CellStyleKey(cell);
            if (!styles.ContainsKey(key))
                styles.Add(key, $"ce{styles.Count + 1}");
        }
        return styles;
    }

    private static OdsCellStyleKey CellStyleKey(OdsCell cell)
    {
        var parent = cell.PresentationRole switch
        {
            StatementPresentationRole.DocumentTitle => "Title",
            StatementPresentationRole.Metadata => "Metadata",
            StatementPresentationRole.ColumnHeader => "ColumnHeader",
            StatementPresentationRole.SectionHeader => "SectionHeader",
            StatementPresentationRole.SummaryHeader => "SummaryHeader",
            StatementPresentationRole.CalculatedTotalHeader => "CalculatedHeader",
            StatementPresentationRole.AnalysisHeader => "AnalysisHeader",
            StatementPresentationRole.BankHeader => "BankHeader",
            StatementPresentationRole.VatHeader => "VatHeader",
            StatementPresentationRole.BalanceSheetHeader => "BalanceSheetHeader",
            StatementPresentationRole.ReconciliationHeader => "ReconciliationHeader",
            StatementPresentationRole.ReconciliationPass => "Pass",
            StatementPresentationRole.ReconciliationWarning => "Warning",
            StatementPresentationRole.ReconciliationFailure => "Failure",
            StatementPresentationRole.HiddenMarker => "HiddenMarker",
            StatementPresentationRole.ExpenseTotal => cell.NumericKind == StatementNumericKind.Money
                ? "MoneyExpenseTotal" : "ExpenseTotal",
            StatementPresentationRole.IncomeTotal => cell.NumericKind == StatementNumericKind.Money
                ? "MoneyIncomeTotal" : "IncomeTotal",
            StatementPresentationRole.NeutralTotal => cell.NumericKind == StatementNumericKind.Money
                ? "MoneyNeutralTotal" : "NeutralTotal",
            _ => cell.NumericKind switch
            {
                StatementNumericKind.Money => "Money",
                StatementNumericKind.Percentage => "Percent",
                StatementNumericKind.Ratio => "Ratio",
                _ => "Default"
            }
        };
        var background = cell.ActivePeriod && !HasSemanticBackground(cell.PresentationRole)
            ? "#ffff00"
            : null;
        return new OdsCellStyleKey(
            parent,
            DataStyleName(cell.NumberFormat),
            background,
            cell.Bold,
            cell.Locked,
            cell.TopBorder,
            cell.BottomBorder,
            cell.LeftBorder,
            cell.RightBorder,
            cell.FontColor,
            cell.HorizontalAlignment);
    }

    private static bool HasSemanticBackground(StatementPresentationRole role) => role is
        StatementPresentationRole.ExpenseTotal or
        StatementPresentationRole.IncomeTotal or
        StatementPresentationRole.NeutralTotal or
        StatementPresentationRole.SummaryHeader or
        StatementPresentationRole.CalculatedTotalHeader or
        StatementPresentationRole.AnalysisHeader or
        StatementPresentationRole.BankHeader or
        StatementPresentationRole.VatHeader or
        StatementPresentationRole.BalanceSheetHeader or
        StatementPresentationRole.ReconciliationPass or
        StatementPresentationRole.ReconciliationWarning or
        StatementPresentationRole.ReconciliationFailure;

    private static string? DataStyleName(string? format)
    {
        if (string.IsNullOrWhiteSpace(format))
            return null;
        var normalized = format.Trim().ToUpperInvariant();
        if (normalized.StartsWith("PCT", StringComparison.Ordinal) &&
            int.TryParse(normalized.AsSpan(3), out var percentageDecimals))
            return $"P{Math.Clamp(percentageDecimals, 0, 2)}";
        if ((normalized.StartsWith("NUM", StringComparison.Ordinal) ||
             normalized.StartsWith("CASH", StringComparison.Ordinal)) &&
            int.TryParse(normalized.AsSpan(normalized.StartsWith("NUM", StringComparison.Ordinal) ? 3 : 4), out var numberDecimals))
            return $"N{Math.Clamp(numberDecimals, 0, 2)}";
        if (normalized.Contains('%'))
            return $"P{DecimalPlaces(normalized)}";
        return $"N{DecimalPlaces(normalized)}";
    }

    private static int DecimalPlaces(string format)
    {
        var firstSection = format.Split(';')[0];
        var decimalPoint = firstSection.IndexOf('.');
        if (decimalPoint < 0)
            return 0;
        return Math.Clamp(firstSection[(decimalPoint + 1)..].Count(character => character is '0' or '#'), 0, 2);
    }

    private static XElement DynamicCellStyle(string name, OdsCellStyleKey key)
    {
        var style = new XElement(Style + "style",
            new XAttribute(Style + "name", name),
            new XAttribute(Style + "family", "table-cell"),
            new XAttribute(Style + "parent-style-name", key.ParentStyle));
        if (key.DataStyle is not null)
            style.Add(new XAttribute(Style + "data-style-name", key.DataStyle));

        var cellProperties = new XElement(Style + "table-cell-properties",
            new XAttribute(Style + "cell-protect", key.Locked ? "protected" : "none"),
            new XAttribute(Style + "vertical-align", "middle"),
            new XAttribute(Style + "wrap-option", "no-wrap"));
        if (key.Background is not null)
            cellProperties.Add(new XAttribute(Fo + "background-color", key.Background));
        AddBorder(cellProperties, "border-top", key.TopBorder);
        AddBorder(cellProperties, "border-bottom", key.BottomBorder);
        AddBorder(cellProperties, "border-left", key.LeftBorder);
        AddBorder(cellProperties, "border-right", key.RightBorder);
        style.Add(cellProperties);

        if (key.HorizontalAlignment != OdsHorizontalAlignment.Default)
            style.Add(new XElement(Style + "paragraph-properties",
                new XAttribute(Fo + "text-align", key.HorizontalAlignment switch
                {
                    OdsHorizontalAlignment.Left => "start",
                    OdsHorizontalAlignment.Centre => "center",
                    OdsHorizontalAlignment.Right => "end",
                    _ => "start"
                })));
        if (key.Bold || key.FontColor is not null)
        {
            var textProperties = new XElement(Style + "text-properties");
            if (key.Bold)
                textProperties.Add(new XAttribute(Fo + "font-weight", "bold"));
            if (key.FontColor is not null)
                textProperties.Add(new XAttribute(Fo + "color", key.FontColor));
            style.Add(textProperties);
        }
        return style;
    }

    private static void AddBorder(XElement properties, string name, OdsBorderWeight weight)
    {
        if (weight == OdsBorderWeight.None)
            return;
        var width = weight switch
        {
            OdsBorderWeight.Thin => "0.5pt",
            OdsBorderWeight.Medium => "1pt",
            OdsBorderWeight.Thick => "1.5pt",
            _ => "0.5pt"
        };
        properties.Add(new XAttribute(Fo + name, $"{width} solid #000000"));
    }

    private static string ColumnStyleName(int sheetIndex, int columnIndex) =>
        $"co{sheetIndex + 1}_{columnIndex + 1}";

    private static XDocument CreateStyles() => Document(Office + "document-styles",
        new XElement(Office + "styles",
            NumberStyles("N0", 0),
            NumberStyles("N1", 1),
            NumberStyles("N2", 2),
            PercentageStyles("P0", 0),
            PercentageStyles("P1", 1),
            PercentageStyles("P2", 2),
            StyleElement("Default"),
            StyleElement("Title", bold: true, fontSize: "12pt"),
            StyleElement("Metadata", bold: true, fontSize: "10pt"),
            StyleElement("Money", "N0", paragraphEnd: true),
            StyleElement("Percent", "P2", paragraphEnd: true),
            StyleElement("Ratio", "N2", paragraphEnd: true),
            StyleElement("ColumnHeader", bold: true, fontSize: "8pt"),
            StyleElement("SectionHeader", bold: true),
            StyleElement("SummaryHeader", background: "#90ee90", bold: true),
            StyleElement("CalculatedHeader", background: "#adff2f", bold: true),
            StyleElement("AnalysisHeader", background: "#32cd32", bold: true),
            StyleElement("BankHeader", background: "#b0c4de", bold: true),
            StyleElement("VatHeader", background: "#da70d6", bold: true),
            StyleElement("BalanceSheetHeader", background: "#ff0000", fontColor: "#ffffff", bold: true),
            StyleElement("ReconciliationHeader", bold: true),
            StyleElement("Pass", background: "#ccffcc", fontColor: "#006400", bold: true),
            StyleElement("Warning", background: "#fff2cc", fontColor: "#ff8c00", bold: true),
            StyleElement("Failure", background: "#ffe0e0", fontColor: "#ff0000", bold: true),
            StyleElement("HiddenMarker", fontColor: "#32cd32"),
            StyleElement("ExpenseTotal", background: "#ffa07a", bold: true),
            StyleElement("MoneyExpenseTotal", "N0", background: "#ffa07a", bold: true, paragraphEnd: true),
            StyleElement("IncomeTotal", background: "#6495ed", bold: true),
            StyleElement("MoneyIncomeTotal", "N0", background: "#6495ed", bold: true, paragraphEnd: true),
            StyleElement("NeutralTotal", background: "#d3d3d3", bold: true),
            StyleElement("MoneyNeutralTotal", "N0", background: "#d3d3d3", bold: true, paragraphEnd: true)),
        new XElement(Office + "automatic-styles"),
        new XElement(Office + "master-styles"));

    private static IEnumerable<XElement> NumberStyles(string name, int decimals)
    {
        var negativeName = $"{name}Neg";
        yield return new XElement(Number + "number-style",
            new XAttribute(Style + "name", name),
            NumberElement(decimals, grouping: true),
            new XElement(Style + "map",
                new XAttribute(Style + "condition", "value()<0"),
                new XAttribute(Style + "apply-style-name", negativeName)));
        yield return NegativeNumberStyle(negativeName, decimals, percentage: false);
    }

    private static IEnumerable<XElement> PercentageStyles(string name, int decimals)
    {
        var negativeName = $"{name}Neg";
        yield return new XElement(Number + "percentage-style",
            new XAttribute(Style + "name", name),
            NumberElement(decimals, grouping: false),
            new XElement(Number + "text", "%"),
            new XElement(Style + "map",
                new XAttribute(Style + "condition", "value()<0"),
                new XAttribute(Style + "apply-style-name", negativeName)));
        yield return NegativeNumberStyle(negativeName, decimals, percentage: true);
    }

    private static XElement NegativeNumberStyle(string name, int decimals, bool percentage) =>
        new(percentage ? Number + "percentage-style" : Number + "number-style",
            new XAttribute(Style + "name", name),
            new XElement(Number + "text", "("),
            NumberElement(decimals, grouping: !percentage),
            percentage ? new XElement(Number + "text", "%)") : new XElement(Number + "text", ")"),
            new XElement(Style + "text-properties", new XAttribute(Fo + "color", "#ff0000")));

    private static XElement NumberElement(int decimals, bool grouping) =>
        new(Number + "number",
            new XAttribute(Number + "decimal-places", decimals),
            new XAttribute(Number + "min-integer-digits", "1"),
            new XAttribute(Number + "grouping", grouping ? "true" : "false"));

    private static XElement StyleElement(
        string name,
        string? dataStyle = null,
        string? background = null,
        string? fontColor = null,
        string? fontSize = null,
        bool bold = false,
        bool paragraphEnd = false,
        bool bottomBorder = false,
        bool rightBorder = false)
    {
        var element = new XElement(Style + "style",
            new XAttribute(Style + "name", name),
            new XAttribute(Style + "family", "table-cell"));
        if (dataStyle is not null)
            element.Add(new XAttribute(Style + "data-style-name", dataStyle));
        if (background is not null || bottomBorder || rightBorder)
        {
            var properties = new XElement(Style + "table-cell-properties");
            if (background is not null)
                properties.Add(new XAttribute(Fo + "background-color", background));
            if (bottomBorder)
                properties.Add(new XAttribute(Fo + "border-bottom", "0.04cm solid #000000"));
            if (rightBorder)
                properties.Add(new XAttribute(Fo + "border-right", "0.06cm solid #000000"));
            element.Add(properties);
        }
        if (paragraphEnd)
            element.Add(new XElement(Style + "paragraph-properties", new XAttribute(Fo + "text-align", "end")));
        if (bold || fontColor is not null || fontSize is not null)
        {
            var textProperties = new XElement(Style + "text-properties");
            if (bold)
                textProperties.Add(new XAttribute(Fo + "font-weight", "bold"));
            if (fontColor is not null)
                textProperties.Add(new XAttribute(Fo + "color", fontColor));
            if (fontSize is not null)
                textProperties.Add(new XAttribute(Fo + "font-size", fontSize));
            element.Add(textProperties);
        }
        return element;
    }

    private static XDocument CreateSettings(OdsSpreadsheetDocument document)
    {
        var active = document.Worksheets[0];
        var tableEntries = new XElement(Config + "config-item-map-named",
            new XAttribute(Config + "name", "Tables"));
        foreach (var worksheet in document.Worksheets)
        {
            var activeSplitRange = (worksheet.FrozenColumns > 0 ? 1 : 0) +
                                   (worksheet.FrozenRows > 0 ? 2 : 0);
            tableEntries.Add(new XElement(Config + "config-item-map-entry",
                new XAttribute(Config + "name", worksheet.Name),
                ConfigItem("CursorPositionX", "int", worksheet.FrozenColumns),
                ConfigItem("CursorPositionY", "int", worksheet.FrozenRows),
                ConfigItem("HorizontalSplitMode", "short", worksheet.FrozenColumns > 0 ? 2 : 0),
                ConfigItem("VerticalSplitMode", "short", worksheet.FrozenRows > 0 ? 2 : 0),
                ConfigItem("HorizontalSplitPosition", "int", worksheet.FrozenColumns),
                ConfigItem("VerticalSplitPosition", "int", worksheet.FrozenRows),
                ConfigItem("ActiveSplitRange", "short", activeSplitRange),
                ConfigItem("PositionLeft", "int", 0),
                ConfigItem("PositionRight", "int", worksheet.FrozenColumns),
                ConfigItem("PositionTop", "int", 0),
                ConfigItem("PositionBottom", "int", worksheet.FrozenRows),
                ConfigItem("ZoomType", "short", 0),
                ConfigItem("ZoomValue", "int", 100),
                ConfigItem("PageViewZoomValue", "int", 60),
                ConfigItem("ShowGrid", "boolean", "true")));
        }

        var settings = Document(Office + "document-settings",
            new XElement(Office + "settings",
                new XElement(Config + "config-item-set",
                    new XAttribute(Config + "name", "ooo:view-settings"),
                    ConfigItem("VisibleAreaTop", "int", 0),
                    ConfigItem("VisibleAreaLeft", "int", 0),
                    ConfigItem("VisibleAreaWidth", "int", 18898),
                    ConfigItem("VisibleAreaHeight", "int", 4676),
                    new XElement(Config + "config-item-map-indexed",
                        new XAttribute(Config + "name", "Views"),
                        new XElement(Config + "config-item-map-entry",
                            ConfigItem("ViewId", "string", "view1"),
                            tableEntries,
                            ConfigItem("ActiveTable", "string", active.Name),
                            ConfigItem("HorizontalScrollbarWidth", "int", 1359),
                            ConfigItem("ZoomType", "short", 0),
                            ConfigItem("ZoomValue", "int", 100),
                            ConfigItem("PageViewZoomValue", "int", 60),
                            ConfigItem("ShowPageBreakPreview", "boolean", "false"),
                            ConfigItem("ShowZeroValues", "boolean", "true"),
                            ConfigItem("ShowNotes", "boolean", "true"),
                            ConfigItem("ShowNoteAuthor", "boolean", "true"),
                            ConfigItem("ShowFormulasMarks", "boolean", "false"),
                            ConfigItem("ShowGrid", "boolean", "true"),
                            ConfigItem("GridColor", "int", 12632256),
                            ConfigItem("ShowPageBreaks", "boolean", "true"),
                            ConfigItem("HasColumnRowHeaders", "boolean", "true"),
                            ConfigItem("HasSheetTabs", "boolean", "true"),
                            ConfigItem("IsOutlineSymbolsSet", "boolean", "true"),
                            ConfigItem("IsValueHighlightingEnabled", "boolean", "false"),
                            ConfigItem("IsSnapToRaster", "boolean", "false"),
                            ConfigItem("RasterIsVisible", "boolean", "false"),
                            ConfigItem("RasterResolutionX", "int", 1000),
                            ConfigItem("RasterResolutionY", "int", 1000),
                            ConfigItem("RasterSubdivisionX", "int", 1),
                            ConfigItem("RasterSubdivisionY", "int", 1),
                            ConfigItem("IsRasterAxisSynchronized", "boolean", "true"),
                            ConfigItem("FormulaBarHeight", "short", 1)))),
                new XElement(Config + "config-item-set",
                    new XAttribute(Config + "name", "ooo:configuration-settings"))));
        settings.Root!.Add(
            new XAttribute(XNamespace.Xmlns + "ooo", Ooo),
            new XAttribute(XNamespace.Xmlns + "xlink", XLink));
        return settings;
    }

    private static XElement ConfigItem(string name, string type, object value) =>
        new(Config + "config-item",
            new XAttribute(Config + "name", name),
            new XAttribute(Config + "type", type),
            Convert.ToString(value, CultureInfo.InvariantCulture));

    private static XDocument CreateMeta(string generator) => Document(Office + "document-meta",
        new XElement(Office + "meta", new XElement(Meta + "generator", generator)));

    private static XDocument CreateManifest() => new(
        new XDeclaration("1.0", "UTF-8", null),
        new XElement(Manifest + "manifest",
            new XAttribute(XNamespace.Xmlns + "manifest", Manifest),
            new XAttribute(Manifest + "version", "1.3"),
            ManifestEntry("/", MimeType, "1.3"),
            ManifestEntry("content.xml", "text/xml"),
            ManifestEntry("styles.xml", "text/xml"),
            ManifestEntry("settings.xml", "text/xml"),
            ManifestEntry("meta.xml", "text/xml")));

    private static XElement ManifestEntry(string path, string mediaType, string? version = null)
    {
        var entry = new XElement(Manifest + "file-entry",
            new XAttribute(Manifest + "full-path", path),
            new XAttribute(Manifest + "media-type", mediaType));
        if (version is not null)
            entry.Add(new XAttribute(Manifest + "version", version));
        return entry;
    }

    private static XDocument Document(XName rootName, params object[] content) => new(
        new XDeclaration("1.0", "UTF-8", null),
        new XElement(rootName,
            new XAttribute(XNamespace.Xmlns + "office", Office),
            new XAttribute(XNamespace.Xmlns + "table", Table),
            new XAttribute(XNamespace.Xmlns + "text", Text),
            new XAttribute(XNamespace.Xmlns + "style", Style),
            new XAttribute(XNamespace.Xmlns + "number", Number),
            new XAttribute(XNamespace.Xmlns + "fo", Fo),
            new XAttribute(XNamespace.Xmlns + "config", Config),
            new XAttribute(XNamespace.Xmlns + "of", Of),
            new XAttribute(Office + "version", "1.3"),
            content));

    private static void AddXmlEntry(ZipArchive archive, string name, XDocument document) =>
        AddTextEntry(archive, name, document.ToString(SaveOptions.DisableFormatting), CompressionLevel.Optimal);

    private static void AddTextEntry(ZipArchive archive, string name, string content, CompressionLevel compression)
    {
        var entry = archive.CreateEntry(name, compression);
        entry.LastWriteTime = StableTimestamp;
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(content);
    }

    private sealed record OdsCellStyleKey(
        string ParentStyle,
        string? DataStyle,
        string? Background,
        bool Bold,
        bool Locked,
        OdsBorderWeight TopBorder,
        OdsBorderWeight BottomBorder,
        OdsBorderWeight LeftBorder,
        OdsBorderWeight RightBorder,
        string? FontColor,
        OdsHorizontalAlignment HorizontalAlignment);

    private sealed class LimitedWriteStream(Stream inner, long maximumBytes, bool leaveOpen) : Stream
    {
        public override bool CanRead => false;
        public override bool CanSeek => inner.CanSeek;
        public override bool CanWrite => true;
        public override long Length => inner.Length;
        public override long Position { get => inner.Position; set => inner.Position = value; }
        public override void Flush() => inner.Flush();
        public override long Seek(long offset, SeekOrigin origin) => inner.Seek(offset, origin);
        public override void SetLength(long value)
        {
            if (value > maximumBytes)
                throw new InvalidDataException($"ODS output limit exceeded ({maximumBytes} bytes).");
            inner.SetLength(value);
        }
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count)
        {
            EnsureWithinLimit(count);
            inner.Write(buffer, offset, count);
        }
        public override void Write(ReadOnlySpan<byte> buffer)
        {
            EnsureWithinLimit(buffer.Length);
            inner.Write(buffer);
        }
        private void EnsureWithinLimit(int count)
        {
            if (inner.Position + count > maximumBytes)
                throw new InvalidDataException($"ODS output limit exceeded ({maximumBytes} bytes).");
        }
        protected override void Dispose(bool disposing)
        {
            if (disposing && !leaveOpen)
                inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
