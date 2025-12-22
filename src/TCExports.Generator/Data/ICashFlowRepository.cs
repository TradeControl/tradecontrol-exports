using TCExports.Generator.Contracts;

namespace TCExports.Generator.Data;

public interface ICashFlowRepository
{
    Task<IReadOnlyList<CashCodePeriodValue>> GetCashCodeValuesAsync(
        string connectionString,
        string cashCode,
        short yearNumber,
        bool includeActivePeriods,
        bool includeOrderBook,
        bool includeTaxAccruals,
        int commandTimeoutSeconds = 30,
        CancellationToken ct = default);

    Task<ActivePeriodDto> GetActivePeriodAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default);
    Task<IReadOnlyList<ActiveYearDto>> GetActiveYearsAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default);
    Task<IReadOnlyList<MonthDto>> GetMonthsAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default);

    Task<IReadOnlyList<FlowCategoryDto>> GetCategoriesAsync(string connectionString, CashType cashType, int commandTimeoutSeconds = 30, CancellationToken ct = default);
    Task<IReadOnlyList<FlowCategoryCashCodeDto>> GetCashCodesAsync(string connectionString, string categoryCode, int commandTimeoutSeconds = 30, CancellationToken ct = default);
    Task<IReadOnlyList<FlowCategoriesByTypeDto>> GetCategoriesByTypeAsync(string connectionString, CashType cashType, CategoryType categoryType, int commandTimeoutSeconds = 30, CancellationToken ct = default);
}