using System.Collections.Immutable;

namespace TCExports.Generator.CashStatement.Model;

/// <summary>
/// Versioned business options understood by the Cash Statement builder.
/// Transport-specific parsing belongs at the request boundary, not in a renderer.
/// </summary>
public sealed record CashStatementRequestV1(
    bool IncludeActivePeriods = false,
    bool IncludeOrderBook = false,
    bool IncludeTaxAccruals = true,
    bool IncludeVatDetails = true,
    bool IncludeBankBalances = true,
    bool IncludeBalanceSheet = true,
    bool IncludeReconciliation = false,
    int CommandTimeoutSeconds = 30)
{
    public const int ContractVersion = 1;

    public int Version => ContractVersion;
}

public readonly record struct StatementSheetId(string Value)
{
    public override string ToString() => Value;
}

public readonly record struct StatementSectionId(string Value)
{
    public override string ToString() => Value;
}

public readonly record struct StatementRowId(string Value)
{
    public override string ToString() => Value;
}

public readonly record struct StatementColumnId(string Value)
{
    public override string ToString() => Value;
}

public readonly record struct StatementPeriodId(short YearNumber, byte? MonthNumber)
{
    public bool IsAnnual => MonthNumber is null;

    public static StatementPeriodId Month(short yearNumber, byte monthNumber) =>
        new(yearNumber, monthNumber);

    public static StatementPeriodId Annual(short yearNumber) => new(yearNumber, null);

    public override string ToString() =>
        IsAnnual ? $"{YearNumber}:annual" : $"{YearNumber}:{MonthNumber:00}";
}

public readonly record struct StatementCellAddress(
    StatementSheetId SheetId,
    StatementRowId RowId,
    StatementColumnId ColumnId);

public enum StatementSheetKind
{
    CashFlow,
    Settings,
    Reconciliation
}

public enum StatementSectionKind
{
    Metadata,
    FlowCategories,
    FlowSummary,
    CategoryTotals,
    Analysis,
    BankBalances,
    VatRecurrence,
    VatPeriods,
    BalanceSheet,
    Settings,
    Reconciliation
}

public enum StatementRowKind
{
    Metadata,
    Header,
    Detail,
    CategoryTotal,
    SummaryReference,
    PeriodTotal,
    CalculatedTotal,
    Expression,
    BankAccount,
    BankTotal,
    VatMetric,
    BalanceSheetAsset,
    Capital,
    Reconciliation
}

public enum StatementColumnKind
{
    Identifier,
    Label,
    HiddenMarker,
    Period,
    AnnualPeriod,
    Attribute,
    Result
}

/// <summary>
/// Format-neutral style intent. XLSX and ODS map these roles to native styles.
/// </summary>
public enum StatementPresentationRole
{
    Default,
    DocumentTitle,
    Metadata,
    ColumnHeader,
    CurrentPeriod,
    AnnualPeriod,
    SectionHeader,
    Detail,
    ExpenseTotal,
    IncomeTotal,
    NeutralTotal,
    SummaryHeader,
    SummaryTotal,
    CalculatedTotalHeader,
    AnalysisHeader,
    AnalysisValue,
    BankHeader,
    VatHeader,
    BalanceSheetHeader,
    CapitalTotal,
    ReconciliationHeader,
    ReconciliationPass,
    ReconciliationWarning,
    ReconciliationFailure,
    HiddenMarker
}

public enum StatementNumericKind
{
    Money,
    Percentage,
    Ratio,
    Count
}

public enum StatementValueState
{
    Present,
    Missing,
    Unsupported
}

public enum StatementScalarKind
{
    Decimal,
    Text,
    Date,
    Boolean
}

/// <summary>
/// A typed authoritative value. Present decimal zero is distinct from missing and unsupported.
/// </summary>
public sealed record StatementValue
{
    private StatementValue(
        StatementValueState state,
        StatementScalarKind kind,
        decimal? decimalValue = null,
        string? textValue = null,
        DateOnly? dateValue = null,
        bool? booleanValue = null,
        string? reason = null)
    {
        State = state;
        Kind = kind;
        DecimalValue = decimalValue;
        TextValue = textValue;
        DateValue = dateValue;
        BooleanValue = booleanValue;
        Reason = reason;
    }

    public StatementValueState State { get; }
    public StatementScalarKind Kind { get; }
    public decimal? DecimalValue { get; }
    public string? TextValue { get; }
    public DateOnly? DateValue { get; }
    public bool? BooleanValue { get; }
    public string? Reason { get; }

    public static StatementValue FromDecimal(decimal value) =>
        new(StatementValueState.Present, StatementScalarKind.Decimal, decimalValue: value);

    public static StatementValue FromText(string value) =>
        new(StatementValueState.Present, StatementScalarKind.Text, textValue: value);

    public static StatementValue FromDate(DateOnly value) =>
        new(StatementValueState.Present, StatementScalarKind.Date, dateValue: value);

    public static StatementValue FromBoolean(bool value) =>
        new(StatementValueState.Present, StatementScalarKind.Boolean, booleanValue: value);

    public static StatementValue Missing(StatementScalarKind kind, string? reason = null) =>
        new(StatementValueState.Missing, kind, reason: reason);

    public static StatementValue Unsupported(StatementScalarKind kind, string reason) =>
        new(StatementValueState.Unsupported, kind, reason: reason);
}

public enum StatementUnaryOperator
{
    Negate,
    Absolute
}

public enum StatementBinaryOperator
{
    Add,
    Subtract,
    Multiply,
    Divide,
    Equal,
    NotEqual,
    LessThan,
    LessThanOrEqual,
    GreaterThan,
    GreaterThanOrEqual
}

public enum StatementAggregateOperator
{
    Sum,
    Count
}

/// <summary>
/// Semantic formula tree. It is optional traceability metadata; StatementCell.Value remains authoritative.
/// </summary>
public abstract record StatementFormula;

public sealed record StatementFormulaReference(StatementCellAddress Address) : StatementFormula;

public sealed record StatementFormulaLiteral(StatementValue Value) : StatementFormula;

public sealed record StatementFormulaUnary(
    StatementUnaryOperator Operator,
    StatementFormula Operand) : StatementFormula;

public sealed record StatementFormulaBinary(
    StatementBinaryOperator Operator,
    StatementFormula Left,
    StatementFormula Right) : StatementFormula;

public sealed record StatementFormulaAggregate(
    StatementAggregateOperator Operator,
    ImmutableArray<StatementFormula> Operands) : StatementFormula;

public sealed record StatementFormulaConditional(
    StatementFormula Condition,
    StatementFormula WhenTrue,
    StatementFormula WhenFalse) : StatementFormula;

/// <summary>
/// An opaque management expression evaluated only by the target spreadsheet engine.
/// The builder binds its reserved category tokens to semantic cell addresses without
/// parsing or otherwise interpreting the expression language.
/// </summary>
public sealed record StatementFormulaTerminal(
    string Expression,
    ImmutableDictionary<string, StatementCellAddress> References) : StatementFormula;

public sealed record StatementCell(
    StatementValue Value,
    StatementNumericKind? NumericKind = null,
    StatementPresentationRole PresentationRole = StatementPresentationRole.Default,
    StatementFormula? Formula = null,
    string? NumberFormat = null);

public sealed record StatementColumn(
    StatementColumnId Id,
    string Label,
    StatementColumnKind Kind,
    int PresentationOrder,
    StatementPresentationRole PresentationRole = StatementPresentationRole.Default,
    StatementPeriodId? PeriodId = null);

public sealed record StatementRow(
    StatementRowId Id,
    StatementRowKind Kind,
    int PresentationOrder,
    string Label,
    string? BusinessCode,
    StatementPresentationRole PresentationRole,
    ImmutableDictionary<StatementColumnId, StatementCell> Cells);

public sealed record StatementSection(
    StatementSectionId Id,
    StatementSectionKind Kind,
    int PresentationOrder,
    string Title,
    StatementPresentationRole PresentationRole,
    ImmutableArray<StatementRow> Rows);

public sealed record StatementSheet(
    StatementSheetId Id,
    StatementSheetKind Kind,
    int PresentationOrder,
    string Name,
    ImmutableArray<StatementColumn> Columns,
    ImmutableArray<StatementSection> Sections);

public sealed record StatementPeriod(
    StatementPeriodId Id,
    int PresentationOrder,
    string MonthLabel,
    string YearLabel,
    string CashStatus,
    DateOnly? StartOn,
    bool IsActive);

public sealed record StatementPeriodGrid(ImmutableArray<StatementPeriod> Periods);

public sealed record CashStatementMetadata(
    string CompanyName,
    string Title,
    string RequestedBy,
    DateTimeOffset GeneratedAt,
    StatementPeriodId ActivePeriod,
    string DateLabel = "Date:");

/// <summary>
/// Immutable renderer-neutral Cash Statement. Sheets are already ordered for presentation;
/// formula evaluation order is derived independently from semantic dependencies.
/// </summary>
public sealed record CashStatementDocument(
    CashStatementMetadata Metadata,
    CashStatementRequestV1 Request,
    StatementPeriodGrid PeriodGrid,
    ImmutableArray<StatementSheet> Sheets);
