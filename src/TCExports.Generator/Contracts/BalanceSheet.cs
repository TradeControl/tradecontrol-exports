namespace TCExports.Generator.Contracts;

/// <summary>
/// Represents a single point-in-time balance for a balance sheet line item
/// (asset/liability/equity) for a given year and month.
/// </summary>
/// <remarks>
/// Values are point-in-time (not period sums). Month 12 is used to carry
/// the closing balance into the year-total column in the worksheet.
/// </remarks>
public sealed class BalanceSheetEntryDto
{
    /// <summary>
    /// The internal code identifying the balance sheet line item.
    /// </summary>
    public string AssetCode { get; init; } = "";

    /// <summary>
    /// The display name of the balance sheet line item.
    /// </summary>
    public string AssetName { get; init; } = "";

    /// <summary>
    /// The financial year number (e.g., 2024).
    /// </summary>
    public short YearNumber { get; init; }

    /// <summary>
    /// The month number within the year (1..12).
    /// </summary>
    public byte MonthNumber { get; init; }

    /// <summary>
    /// The closing balance for the line item at the end of the specified month.
    /// </summary>
    public decimal Balance { get; init; }
}