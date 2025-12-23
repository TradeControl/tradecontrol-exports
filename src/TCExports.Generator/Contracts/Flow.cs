namespace TCExports.Generator.Contracts;

/// <summary>
/// A cash flow category descriptor used to group cash codes in the statement.
/// </summary>
public sealed class FlowCategoryDto
{
    public string CategoryCode { get; init; } = "";
    public string Category { get; init; } = "";
    public short? CashPolarityCode { get; init; } // 0=Expense,1=Income,2=Neutral
    public short? DisplayOrder { get; init; }
}

/// <summary>
/// A cash code mapped to a flow category, with its display description.
/// </summary>
public sealed class FlowCategoryCashCodeDto
{
    public string CashCode { get; init; } = "";
    public string CashDescription { get; init; } = "";
}

/// <summary>
/// A category flagged as a total/summary for a given cash type.
/// </summary>
public sealed class FlowCategoriesByTypeDto
{
    public string CategoryCode { get; init; } = "";
    public string Category { get; init; } = "";
    public string CashType { get; init; } = "";
}

/// <summary>
/// Identifies a category that should be populated by summing other categories.
/// </summary>
public sealed class CategoryTotalDto
{
    public string CategoryCode { get; init; } = "";
}

/// <summary>
/// A computed expression row that references category totals to derive analysis values.
/// </summary>
public sealed class CategoryExpressionDto
{
    public short DisplayOrder { get; init; }
    public string Category { get; init; } = "";
    public string Expression { get; init; } = "";
    public string Format { get; init; } = "";
}

public sealed class BankAccountDto
{
    public string AccountCode { get; init; } = "";
    public string AccountName { get; init; } = "";
    public decimal OpeningBalance { get; init; }
}

public sealed class BankBalanceDto
{
    public string AccountCode { get; init; } = "";
    public short YearNumber { get; init; }
    public byte MonthNumber { get; init; }
    public decimal Balance { get; init; }
}