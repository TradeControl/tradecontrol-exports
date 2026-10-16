using System.Globalization;
using System.Resources;

namespace TCExports.Generator.CashStatement;

public enum CashStatementText
{
    StatementTitle,
    Date,
    CashFlow,
    Code,
    Name,
    Totals,
    Summary,
    PeriodTotal,
    TradeCashFlow,
    MoneyCashFlow,
    TaxCashFlow,
    TradeTotals,
    TaxTotals,
    Analysis,
    ClosingBalances,
    CompanyBalance,
    Vat,
    VatDueSales,
    VatDueAcquisitions,
    TotalVatDue,
    VatReclaimedCurrentPeriod,
    NetVatDue,
    TotalValueSalesExVat,
    TotalValuePurchasesExVat,
    TotalValueGoodsSuppliedExVat,
    TotalValueGoodsReceivedExVat,
    BalanceSheet,
    Capital,
    Reconciliation,
    ReconciliationChecks,
    Year,
    Description,
    OpeningCapital,
    ClosingCapital,
    Profit,
    BusinessTax,
    ProfitAfterTax,
    CapitalMovement,
    OpeningSubjectPosition,
    OpeningAccountPosition,
    CapitalDelta,
    BridgeTotal,
    Difference,
    Status,
    Tolerance
}

/// <summary>
/// Resolves report-owned text for one request. Database-owned descriptions are not localized.
/// </summary>
public interface ICashStatementLocalizer
{
    string Get(CashStatementText text);

    string Format(CashStatementText text, params object[] arguments);
}

/// <summary>
/// Uses the generator's satellite-resource convention without mutating thread or process culture.
/// </summary>
public sealed class ResourceCashStatementLocalizer : ICashStatementLocalizer
{
    private static readonly ResourceManager ResourceManager = new(
        "TCExports.Generator.Properties.Resources",
        typeof(ResourceCashStatementLocalizer).Assembly);

    private static readonly IReadOnlyDictionary<CashStatementText, string> ResourceNames =
        new Dictionary<CashStatementText, string>
        {
            [CashStatementText.StatementTitle] = "TextStatementTitle",
            [CashStatementText.Date] = "TextDate",
            [CashStatementText.CashFlow] = "TextCashFlow",
            [CashStatementText.Code] = "TextCode",
            [CashStatementText.Name] = "TextName",
            [CashStatementText.Totals] = "TextTotals",
            [CashStatementText.Summary] = "TextSummary",
            [CashStatementText.PeriodTotal] = "TextPeriodTotal",
            [CashStatementText.TradeCashFlow] = "TextTradeCashFlow",
            [CashStatementText.MoneyCashFlow] = "TextMoneyCashFlow",
            [CashStatementText.TaxCashFlow] = "TextTaxCashFlow",
            [CashStatementText.TradeTotals] = "TextTradeTotals",
            [CashStatementText.TaxTotals] = "TextTaxTotals",
            [CashStatementText.Analysis] = "TextAnalysis",
            [CashStatementText.ClosingBalances] = "TextClosingBalances",
            [CashStatementText.CompanyBalance] = "TextCompanyBalance",
            [CashStatementText.Vat] = "TextVatDueTitle",
            [CashStatementText.VatDueSales] = "TextVatDueSales",
            [CashStatementText.VatDueAcquisitions] = "TextVatDueAcquisitions",
            [CashStatementText.TotalVatDue] = "TextTotalVatDue",
            [CashStatementText.VatReclaimedCurrentPeriod] = "TextVatReclaimedCurrPeriod",
            [CashStatementText.NetVatDue] = "TextNetVatDue",
            [CashStatementText.TotalValueSalesExVat] = "TextTotalValueSalesExVAT",
            [CashStatementText.TotalValuePurchasesExVat] = "TextTotalValuePurchasesExVAT",
            [CashStatementText.TotalValueGoodsSuppliedExVat] = "TextTotalValueGoodsSuppliedExVAT",
            [CashStatementText.TotalValueGoodsReceivedExVat] = "TextTotalValueGoodsReceivedExVAT",
            [CashStatementText.BalanceSheet] = "TextBalanceSheet",
            [CashStatementText.Capital] = "TextCapital",
            [CashStatementText.Reconciliation] = "TextReconciliation",
            [CashStatementText.ReconciliationChecks] = "TextReconciliationChecks",
            [CashStatementText.Year] = "TextYear",
            [CashStatementText.Description] = "TextDescription",
            [CashStatementText.OpeningCapital] = "TextOpeningCapital",
            [CashStatementText.ClosingCapital] = "TextClosingCapital",
            [CashStatementText.Profit] = "TextProfit",
            [CashStatementText.BusinessTax] = "TextBusinessTax",
            [CashStatementText.ProfitAfterTax] = "TextProfitAfterTax",
            [CashStatementText.CapitalMovement] = "TextCapitalMovement",
            [CashStatementText.OpeningSubjectPosition] = "TextOpeningSubjectPosition",
            [CashStatementText.OpeningAccountPosition] = "TextOpeningAccountPosition",
            [CashStatementText.CapitalDelta] = "TextCapitalDelta",
            [CashStatementText.BridgeTotal] = "TextBridgeTotal",
            [CashStatementText.Difference] = "TextReconDifference",
            [CashStatementText.Status] = "TextStatus",
            [CashStatementText.Tolerance] = "TextTolerance"
        };

    public ResourceCashStatementLocalizer(CultureInfo? culture = null)
    {
        Culture = culture ?? CultureInfo.GetCultureInfo("en-GB");
    }

    public CultureInfo Culture { get; }

    public string Get(CashStatementText text)
    {
        var resourceName = ResourceNames[text];
        return ResourceManager.GetString(resourceName, Culture) ?? resourceName;
    }

    public string Format(CashStatementText text, params object[] arguments) =>
        string.Format(Culture, Get(text), arguments);
}
