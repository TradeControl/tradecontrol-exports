namespace TCExports.Generator.Contracts;

public sealed class EquityReconciliationByYearDto
{
    public short YearNumber { get; init; }
    public string? Description { get; init; }

    public decimal OpeningCapital { get; init; }
    public decimal ClosingCapital { get; init; }

    public decimal Profit { get; init; }
    public decimal BusinessTax { get; init; }
    public decimal ProfitAfterTax { get; init; }

    public decimal CapitalMovement { get; init; }
    public decimal OpeningPosition { get; init; }
    public decimal OpeningAccountPosition { get; init; }

    public decimal CapitalDelta { get; init; }
    public decimal Variance { get; init; }
}
