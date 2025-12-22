namespace TCExports.Generator.Contracts;

public sealed class ActivePeriodDto
{
    public short YearNumber { get; init; }
    public short MonthNumber { get; init; }
    public DateTime StartOn { get; init; }
    public string MonthName { get; init; } = "";
    public string Description { get; init; } = "";
}

public sealed class ActiveYearDto
{
    public short YearNumber { get; init; }
    public string Description { get; init; } = "";
    public string CashStatus { get; init; } = "";
}

public sealed class MonthDto
{
    public short MonthNumber { get; init; }
    public string MonthName { get; init; } = "";    
    public decimal? StartOn { get; init; }  // Decimal sort key from vwMonths.StartOn (not a date)
}