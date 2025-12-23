namespace TCExports.Generator.Contracts;

/// <summary>
/// Represents VAT metrics for a recurrence period (e.g., quarter) across years.
/// Values are period totals, not point-in-time balances.
/// </summary>
public sealed class VatRecurrenceDto
{
    public short YearNumber { get; init; }
    public DateTime StartOn { get; init; }
    public decimal HomeSales { get; init; }
    public decimal HomePurchases { get; init; }
    public decimal ExportSales { get; init; }
    public decimal ExportPurchases { get; init; }
    public decimal HomeSalesVat { get; init; }
    public decimal HomePurchasesVat { get; init; }
    public decimal ExportSalesVat { get; init; }
    public decimal ExportPurchasesVat { get; init; }
    public decimal VatAdjustment { get; init; }
    public decimal VatDue { get; init; }
}

/// <summary>
/// Represents VAT accrual adjustments per recurrence period (e.g., quarter) across years.
/// Values may be null when no accrual was recorded for the period.
/// </summary>
public sealed class VatRecurrenceAccrualDto
{
    public short YearNumber { get; init; }
    public decimal? HomeSalesVat { get; init; }
    public decimal? HomePurchasesVat { get; init; }
    public decimal? ExportSalesVat { get; init; }
    public decimal? ExportPurchasesVat { get; init; }
    public decimal? VatDue { get; init; }
}

/// <summary>
/// Represents VAT totals per month across years.
/// Values are period totals used for monthly VAT reporting.
/// </summary>
public sealed class VatPeriodTotalDto
{
    public short YearNumber { get; init; }
    public DateTime StartOn { get; init; }
    public decimal HomeSales { get; init; }
    public decimal HomePurchases { get; init; }
    public decimal ExportSales { get; init; }
    public decimal ExportPurchases { get; init; }
    public decimal HomeSalesVat { get; init; }
    public decimal HomePurchasesVat { get; init; }
    public decimal ExportSalesVat { get; init; }
    public decimal ExportPurchasesVat { get; init; }
    public decimal VatDue { get; init; }
}

/// <summary>
/// Represents VAT accrual adjustments per month across years.
/// Values may be null when no accrual was recorded for the month.
/// </summary>
public sealed class VatPeriodAccrualDto
{
    public short YearNumber { get; init; }
    public decimal? HomeSalesVat { get; init; }
    public decimal? HomePurchasesVat { get; init; }
    public decimal? ExportSalesVat { get; init; }
    public decimal? ExportPurchasesVat { get; init; }
    public decimal? VatDue { get; init; }
}