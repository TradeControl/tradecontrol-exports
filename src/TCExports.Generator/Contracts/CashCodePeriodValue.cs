namespace TCExports.Generator.Contracts;

public sealed class CashCodePeriodValue
{
    public DateTime StartOn { get; init; }
    public decimal InvoiceValue { get; init; }
    public decimal InvoiceTax { get; init; }
    public decimal ForecastValue { get; init; }
    public decimal ForecastTax { get; init; }
}