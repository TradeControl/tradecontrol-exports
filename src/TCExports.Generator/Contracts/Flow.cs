namespace TCExports.Generator.Contracts;

public sealed class FlowCategoryDto
{
    public string CategoryCode { get; init; } = "";
    public string Category { get; init; } = "";
    public short? CashPolarityCode { get; init; } // 0=Expense,1=Income,2=Neutral
    public short? DisplayOrder { get; init; }
}

public sealed class FlowCategoryCashCodeDto
{
    public string CashCode { get; init; } = "";
    public string CashDescription { get; init; } = "";
}

public sealed class FlowCategoriesByTypeDto
{
    public string CategoryCode { get; init; } = "";
    public string Category { get; init; } = "";
    public string CashType { get; init; } = "";
}