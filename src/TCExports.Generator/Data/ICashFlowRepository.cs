using TCExports.Generator.Contracts;

namespace TCExports.Generator.Data;

public interface ICashFlowRepository
{
    Task<CategoryTreeVersionResolutionDto> ResolveCategoryTreeVersionAsync(
        string connectionString,
        string? categoryVersionId,
        int commandTimeoutSeconds = 30,
        CancellationToken ct = default);

    Task<IReadOnlyList<CategoryTreeVersionIntervalDto>> GetCategoryTreeVersionTimelineAsync(
        string connectionString,
        int commandTimeoutSeconds = 30,
        CancellationToken ct = default);

    Task<IReadOnlyList<CashCodePeriodValue>> GetCashCodeValuesAsync(
        string connectionString,
        string cashCode,
        short yearNumber,
        bool includeActivePeriods,
        bool includeOrderBook,
        bool includeTaxAccruals,
        int commandTimeoutSeconds = 30,
        CancellationToken ct = default);

    Task<string> GetCompanyNameAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default);
    Task<ActivePeriodDto> GetActivePeriodAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default);
    Task<IReadOnlyList<ActiveYearDto>> GetActiveYearsAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default);
    Task<IReadOnlyList<MonthDto>> GetMonthsAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default);
    Task<IReadOnlyList<StatementPeriodBoundaryDto>> GetStatementPeriodBoundariesAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default);

    Task<IReadOnlyList<FlowCategoryDto>> GetCategoriesAsync(string connectionString, string categoryVersionId, CashType cashType, int commandTimeoutSeconds = 30, CancellationToken ct = default);
    Task<IReadOnlyList<FlowCategoryCashCodeDto>> GetCashCodesAsync(string connectionString, string categoryVersionId, string categoryCode, int commandTimeoutSeconds = 30, CancellationToken ct = default);
    Task<IReadOnlyList<FlowCategoriesByTypeDto>> GetCategoriesByTypeAsync(string connectionString, string categoryVersionId, CashType cashType, CategoryType categoryType, int commandTimeoutSeconds = 30, CancellationToken ct = default);

    Task<IReadOnlyList<CategoryTotalDto>> GetCategoryTotalsAsync(string connectionString, string categoryVersionId, int commandTimeoutSeconds = 30, CancellationToken ct = default);
    Task<IReadOnlyList<string>> GetCategoryTotalCodesAsync(string connectionString, string categoryVersionId, string categoryCode, int commandTimeoutSeconds = 30, CancellationToken ct = default);

    Task<IReadOnlyList<CategoryExpressionDto>> GetCategoryExpressionsAsync(string connectionString, string categoryVersionId, int commandTimeoutSeconds = 30, CancellationToken ct = default);
    Task<string> GetCategoryCodeFromNameAsync(string connectionString, string categoryVersionId, string categoryName, int commandTimeoutSeconds = 30, CancellationToken ct = default);

    Task SetCategoryExpressionStatusAsync(
        string connectionString,
        string categoryCode,
        bool isError,
        string? errorMessage = null,
        int commandTimeoutSeconds = 30,
        CancellationToken ct = default
    );

    Task<IReadOnlyList<BankAccountDto>> GetBankAccountsAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default);
    Task<IReadOnlyList<BankBalanceDto>> GetBankBalancesAsync(string connectionString, string accountCode, int commandTimeoutSeconds = 30, CancellationToken ct = default);

    Task<string> GetVatRecurrenceTypeAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default);
    Task<IReadOnlyList<VatRecurrenceDto>> GetVatRecurrenceAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default);
    Task<IReadOnlyList<VatRecurrenceAccrualDto>> GetVatRecurrenceAccrualsAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default);
    Task<IReadOnlyList<VatPeriodTotalDto>> GetVatPeriodTotalsAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default);
    Task<IReadOnlyList<VatPeriodAccrualDto>> GetVatPeriodAccrualsAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default);

    Task<IReadOnlyList<BalanceSheetEntryDto>> GetBalanceSheetAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default);

    Task<IReadOnlyList<EquityReconciliationByYearDto>> GetEquityReconciliationByYearAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default);
    Task<IReadOnlyList<EquityReconciliationByYearDto>> GetEquityReconciliationByYearByVersionAsync(string connectionString, string categoryVersionId, int commandTimeoutSeconds = 30, CancellationToken ct = default);
}
