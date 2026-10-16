using System.Collections.Immutable;
using TCExports.Generator.CashStatement.Model;

namespace TCExports.Generator.Ods;

public sealed record OdsSpreadsheetDocument(
    ImmutableArray<OdsWorksheet> Worksheets,
    string Generator = "Trade Control TCExports");

public sealed record OdsWorksheet(
    StatementSheetId Id,
    string Name,
    ImmutableArray<OdsColumn> Columns,
    ImmutableArray<OdsRow> Rows,
    int FrozenColumns = 0,
    int FrozenRows = 0,
    bool Protected = false,
    bool Print = true);

public sealed record OdsColumn(decimal WidthCentimetres, bool Hidden = false);

public sealed record OdsRow(ImmutableArray<OdsCell> Cells);

public enum OdsBorderWeight
{
    None,
    Thin,
    Medium,
    Thick
}

public enum OdsHorizontalAlignment
{
    Default,
    Left,
    Centre,
    Right
}

/// <summary>
/// Physical ODS cell bound to the renderer-neutral Cash Statement value and formula types.
/// Phase 6 owns row allocation; this type deliberately contains no financial calculation.
/// </summary>
public sealed record OdsCell(
    StatementValue Value,
    StatementNumericKind? NumericKind = null,
    StatementPresentationRole PresentationRole = StatementPresentationRole.Default,
    StatementFormula? Formula = null,
    StatementCellAddress? SemanticAddress = null,
    int ColumnSpan = 1,
    string? NumberFormat = null,
    bool Bold = false,
    bool Locked = false,
    bool ActivePeriod = false,
    OdsBorderWeight TopBorder = OdsBorderWeight.None,
    OdsBorderWeight BottomBorder = OdsBorderWeight.None,
    OdsBorderWeight LeftBorder = OdsBorderWeight.None,
    OdsBorderWeight RightBorder = OdsBorderWeight.None,
    string? FontColor = null,
    OdsHorizontalAlignment HorizontalAlignment = OdsHorizontalAlignment.Default);

public sealed record OdsWriteLimits(
    int MaximumWorksheets = 3,
    int MaximumColumnsPerWorksheet = 1_024,
    int MaximumRows = 10_000,
    int MaximumCells = 2_000_000,
    long MaximumOutputBytes = 64L * 1024 * 1024)
{
    public static OdsWriteLimits Default { get; } = new();
}

