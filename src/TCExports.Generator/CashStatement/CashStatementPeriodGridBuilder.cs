using System.Collections.Immutable;
using TCExports.Generator.CashStatement.Model;
using TCExports.Generator.Contracts;

namespace TCExports.Generator.CashStatement;

public static class CashStatementPeriodGridBuilder
{
    public static CashStatementContractResult<StatementPeriodGrid> Build(
        IReadOnlyList<ActiveYearDto> years,
        IReadOnlyList<MonthDto> months,
        ActivePeriodDto activePeriod,
        CashStatementLimits limits)
    {
        var violations = ImmutableArray.CreateBuilder<StatementContractViolation>();

        if (years.Count == 0)
            violations.Add(new("PERIOD_YEARS_EMPTY", "periodGrid", "At least one fiscal year is required."));
        if (years.Count > limits.MaxFiscalYears)
            violations.Add(new("LIMIT_FISCAL_YEARS", "periodGrid", $"Fiscal-year count {years.Count} exceeds {limits.MaxFiscalYears}."));
        if (months.Count == 0)
            violations.Add(new("PERIOD_MONTHS_EMPTY", "periodGrid", "At least one financial month is required."));

        var duplicateYears = years.GroupBy(y => y.YearNumber).Where(g => g.Count() > 1).Select(g => g.Key);
        foreach (var year in duplicateYears)
            violations.Add(new("PERIOD_YEAR_DUPLICATE", $"periodGrid.year:{year}", $"Fiscal year {year} occurs more than once."));

        var duplicateMonths = months.GroupBy(m => m.MonthNumber).Where(g => g.Count() > 1).Select(g => g.Key);
        foreach (var month in duplicateMonths)
            violations.Add(new("PERIOD_MONTH_DUPLICATE", $"periodGrid.month:{month}", $"Financial month {month} occurs more than once."));

        foreach (var month in months.Where(m => m.MonthNumber is < 1 or > 12))
            violations.Add(new("PERIOD_MONTH_INVALID", $"periodGrid.month:{month.MonthNumber}", "Financial month number must be from 1 through 12."));

        var periodColumnCount = (long)years.Count * (months.Count + 1L);
        if (periodColumnCount > limits.MaxPeriodColumns)
            violations.Add(new("LIMIT_PERIOD_COLUMNS", "periodGrid", $"Period-column count {periodColumnCount} exceeds {limits.MaxPeriodColumns}."));

        if (!years.Any(y => y.YearNumber == activePeriod.YearNumber) ||
            !months.Any(m => m.MonthNumber == activePeriod.MonthNumber))
        {
            violations.Add(new(
                "PERIOD_ACTIVE_NOT_FOUND",
                "periodGrid.activePeriod",
                $"Active period {activePeriod.YearNumber}:{activePeriod.MonthNumber:00} is outside the supplied financial calendar."));
        }

        if (violations.Count > 0)
            return new(null, violations.ToImmutable());

        var periods = ImmutableArray.CreateBuilder<StatementPeriod>((int)periodColumnCount);
        var presentationOrder = 0;

        foreach (var year in years)
        {
            foreach (var month in months)
            {
                var isActive = year.YearNumber == activePeriod.YearNumber &&
                    month.MonthNumber == activePeriod.MonthNumber;

                periods.Add(new StatementPeriod(
                    StatementPeriodId.Month(year.YearNumber, checked((byte)month.MonthNumber)),
                    presentationOrder++,
                    month.MonthName,
                    year.Description,
                    year.CashStatus,
                    isActive ? DateOnly.FromDateTime(activePeriod.StartOn) : null,
                    isActive));
            }

            periods.Add(new StatementPeriod(
                StatementPeriodId.Annual(year.YearNumber),
                presentationOrder++,
                string.Empty,
                year.Description,
                year.CashStatus,
                null,
                false));
        }

        return new(new StatementPeriodGrid(periods.ToImmutable()), ImmutableArray<StatementContractViolation>.Empty);
    }
}
