using System.Collections.Immutable;
using TCExports.Generator.CashStatement.Model;
using TCExports.Generator.Contracts;

namespace TCExports.Generator.CashStatement;

public static class CashStatementPeriodVersionAttributor
{
    public static CashStatementContractResult<StatementPeriodGrid> Attribute(
        StatementPeriodGrid grid,
        IReadOnlyList<StatementPeriodBoundaryDto> boundaries,
        IReadOnlyList<CategoryTreeVersionIntervalDto> versions,
        string selectedVersionId)
    {
        ArgumentNullException.ThrowIfNull(grid);
        ArgumentNullException.ThrowIfNull(boundaries);
        ArgumentNullException.ThrowIfNull(versions);
        ArgumentException.ThrowIfNullOrWhiteSpace(selectedVersionId);

        var violations = ImmutableArray.CreateBuilder<StatementContractViolation>();
        var orderedVersions = versions
            .OrderBy(version => version.EffectiveFromOn)
            .ThenBy(version => version.VersionId, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        if (orderedVersions.Length == 0)
        {
            violations.Add(new(
                "CATEGORY_VERSION_TIMELINE_EMPTY",
                "periodGrid",
                "No effective Category Tree version is available for period attribution."));
            return new(null, violations.ToImmutable());
        }

        ValidateVersions(orderedVersions, violations);

        var boundaryLookup = new Dictionary<(short YearNumber, short MonthNumber), StatementPeriodBoundaryDto>();
        foreach (var boundary in boundaries)
        {
            if (!boundaryLookup.TryAdd((boundary.YearNumber, boundary.MonthNumber), boundary))
            {
                violations.Add(new(
                    "PERIOD_BOUNDARY_DUPLICATE",
                    $"periodGrid.{boundary.YearNumber}:{boundary.MonthNumber:00}",
                    "The statement period has more than one date boundary."));
            }
            if (boundary.EndOn < boundary.StartOn)
            {
                violations.Add(new(
                    "PERIOD_BOUNDARY_INVALID",
                    $"periodGrid.{boundary.YearNumber}:{boundary.MonthNumber:00}",
                    "The statement period end must not be earlier than its start."));
            }
        }

        if (violations.Count > 0)
            return new(null, violations.ToImmutable());

        var attributed = ImmutableArray.CreateBuilder<StatementPeriod>(grid.Periods.Length);
        foreach (var period in grid.Periods.OrderBy(period => period.PresentationOrder))
        {
            StatementPeriodBoundaryDto[] periodBoundaries;
            if (period.Id.IsAnnual)
            {
                periodBoundaries = boundaries
                    .Where(boundary => boundary.YearNumber == period.Id.YearNumber)
                    .OrderBy(boundary => boundary.StartOn)
                    .ToArray();
            }
            else if (boundaryLookup.TryGetValue(
                (period.Id.YearNumber, checked((short)period.Id.MonthNumber!.Value)),
                out var boundary))
            {
                periodBoundaries = [boundary];
            }
            else
            {
                periodBoundaries = [];
            }

            if (periodBoundaries.Length == 0)
            {
                violations.Add(new(
                    "PERIOD_BOUNDARY_MISSING",
                    $"periodGrid.{period.Id}",
                    $"No date boundary is available for statement period '{period.Id}'."));
                continue;
            }

            var periodEnd = periodBoundaries.Max(item => item.EndOn);
            var applicable = Resolve(orderedVersions, periodEnd);
            if (applicable is null)
            {
                violations.Add(new(
                    "CATEGORY_VERSION_PERIOD_GAP",
                    $"periodGrid.{period.Id}",
                    $"No Category Tree version applies at the end of statement period '{period.Id}'."));
                continue;
            }

            attributed.Add(period with
            {
                StartOn = period.Id.IsAnnual
                    ? null
                    : periodBoundaries[0].StartOn,
                ApplicableCategoryVersionId = applicable.VersionId,
                DiffersFromSelectedVersion = !string.Equals(
                    applicable.VersionId,
                    selectedVersionId,
                    StringComparison.OrdinalIgnoreCase)
            });
        }

        return violations.Count == 0
            ? new(new StatementPeriodGrid(attributed.ToImmutable()), ImmutableArray<StatementContractViolation>.Empty)
            : new(null, violations.ToImmutable());
    }

    private static CategoryTreeVersionIntervalDto? Resolve(
        IReadOnlyList<CategoryTreeVersionIntervalDto> versions,
        DateOnly periodEnd)
    {
        var first = versions[0];
        if (periodEnd < first.EffectiveFromOn)
            return first;

        return versions
            .Where(version => version.EffectiveFromOn <= periodEnd
                && (version.EffectiveToOn is null || periodEnd <= version.EffectiveToOn))
            .OrderByDescending(version => version.EffectiveFromOn)
            .FirstOrDefault();
    }

    private static void ValidateVersions(
        IReadOnlyList<CategoryTreeVersionIntervalDto> versions,
        ImmutableArray<StatementContractViolation>.Builder violations)
    {
        var identifiers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var index = 0; index < versions.Count; index++)
        {
            var version = versions[index];
            if (string.IsNullOrWhiteSpace(version.VersionId) || !identifiers.Add(version.VersionId))
            {
                violations.Add(new(
                    "CATEGORY_VERSION_TIMELINE_DUPLICATE",
                    "periodGrid",
                    $"Category Tree version '{version.VersionId}' is duplicated in the effective timeline."));
            }
            if (version.EffectiveToOn is not null && version.EffectiveToOn < version.EffectiveFromOn)
            {
                violations.Add(new(
                    "CATEGORY_VERSION_INTERVAL_INVALID",
                    $"periodGrid.version:{version.VersionId}",
                    "The Category Tree effective interval must not end before it starts."));
            }

            if (index == 0)
                continue;
            var previous = versions[index - 1];
            if (previous.EffectiveToOn?.AddDays(1) != version.EffectiveFromOn)
            {
                violations.Add(new(
                    "CATEGORY_VERSION_TIMELINE_DISCONTINUOUS",
                    $"periodGrid.version:{version.VersionId}",
                    "Effective Category Tree versions must form one continuous timeline."));
            }
        }
    }
}
