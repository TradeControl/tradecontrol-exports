namespace TCExports.Generator.Contracts;

/// <summary>
/// Represents VAT metrics for a recurrence period (e.g., quarter) across years.
/// Values are period totals, not point-in-time balances.
/// </summary>
public sealed class VatRecurrenceDto
{
    public short YearNumber { get; init; }
    public DateTime StartOn { get; init; }
    public decimal VatDueSales { get; init; }
    public decimal VatDueAcquisitions { get; init; }
    public decimal TotalVatDue { get; init; }
    public decimal VatReclaimedCurrPeriod { get; init; }
    public decimal NetVatDue { get; init; }
    public decimal TotalValueSalesExVAT { get; init; }
    public decimal TotalValuePurchasesExVAT { get; init; }
    public decimal TotalValueGoodsSuppliedExVAT { get; init; }
    public decimal TotalValueGoodsReceivedExVAT { get; init; }
}

/// <summary>
/// Represents VAT accrual adjustments per recurrence period (e.g., quarter) across years.
/// Values may be null when no accrual was recorded for the period.
/// </summary>
public sealed class VatRecurrenceAccrualDto
{
    public short YearNumber { get; init; }
    public DateTime StartOn { get; init; }
    public decimal? VatDueSales { get; init; }
    public decimal? VatDueAcquisitions { get; init; }
    public decimal? TotalVatDue { get; init; }
    public decimal? VatReclaimedCurrPeriod { get; init; }
    public decimal? NetVatDue { get; init; }
    public decimal? TotalValueSalesExVAT { get; init; }
    public decimal? TotalValuePurchasesExVAT { get; init; }
    public decimal? TotalValueGoodsSuppliedExVAT { get; init; }
    public decimal? TotalValueGoodsReceivedExVAT { get; init; }
}

/// <summary>
/// Represents VAT totals per month across years.
/// Values are period totals used for monthly VAT reporting.
/// </summary>
public sealed class VatPeriodTotalDto
{
    public short YearNumber { get; init; }
    public DateTime StartOn { get; init; }
    public decimal VatDueSales { get; init; }
    public decimal VatDueAcquisitions { get; init; }
    public decimal TotalVatDue { get; init; }
    public decimal VatReclaimedCurrPeriod { get; init; }
    public decimal NetVatDue { get; init; }
    public decimal TotalValueSalesExVAT { get; init; }
    public decimal TotalValuePurchasesExVAT { get; init; }
    public decimal TotalValueGoodsSuppliedExVAT { get; init; }
    public decimal TotalValueGoodsReceivedExVAT { get; init; }
}

/// <summary>
/// Represents VAT accrual adjustments per month across years.
/// Values may be null when no accrual was recorded for the month.
/// </summary>
public sealed class VatPeriodAccrualDto
{
    public short YearNumber { get; init; }
    public DateTime StartOn { get; init; }
    public decimal? VatDueSales { get; init; }
    public decimal? VatDueAcquisitions { get; init; }
    public decimal? TotalVatDue { get; init; }
    public decimal? VatReclaimedCurrPeriod { get; init; }
    public decimal? NetVatDue { get; init; }
    public decimal? TotalValueSalesExVAT { get; init; }
    public decimal? TotalValuePurchasesExVAT { get; init; }
    public decimal? TotalValueGoodsSuppliedExVAT { get; init; }
    public decimal? TotalValueGoodsReceivedExVAT { get; init; }
}
