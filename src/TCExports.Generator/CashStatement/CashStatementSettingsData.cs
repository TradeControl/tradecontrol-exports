using System.Collections.Immutable;
using System.Globalization;
using TCExports.Generator.CashStatement.Model;

namespace TCExports.Generator.CashStatement;

public sealed record CashStatementSetting(string Name, string Value);

public sealed record CashStatementPeriodComparison(
    string Period,
    string ApplicableCategoryVersionId);

public sealed record CashStatementSettingsData(
    ImmutableArray<CashStatementSetting> CategoryTree,
    ImmutableArray<CashStatementPeriodComparison> PeriodComparisons);

public static class CashStatementSettingsDataBuilder
{
    public static CashStatementSettingsData Build(CashStatementDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);
        if (document.CategoryTree is null)
        {
            return new(
                ImmutableArray<CashStatementSetting>.Empty,
                ImmutableArray<CashStatementPeriodComparison>.Empty);
        }

        var selected = document.CategoryTree.SelectedVersion;
        var active = document.CategoryTree.ActiveVersion;
        var settings = ImmutableArray.CreateBuilder<CashStatementSetting>();
        AddVersion(settings, "Selected", selected);
        AddVersion(settings, "Active", active);

        var comparisons = document.PeriodGrid.Periods
            .Where(period => period.DiffersFromSelectedVersion)
            .OrderBy(period => period.PresentationOrder)
            .Select(period => new CashStatementPeriodComparison(
                PeriodLabel(period),
                period.ApplicableCategoryVersionId))
            .ToImmutableArray();

        return new(settings.ToImmutable(), comparisons);
    }

    private static void AddVersion(
        ImmutableArray<CashStatementSetting>.Builder settings,
        string label,
        CashStatementCategoryTreeVersion version)
    {
        settings.Add(new($"{label} version", version.VersionId));
        settings.Add(new($"{label} description", version.Description));
        settings.Add(new($"{label} status", version.Status));
        settings.Add(new($"{label} effective interval", Interval(version)));

        var evidenceMessage = version.StatusCode == 1
            ? "Not validated since latest tree change"
            : "Canonical evidence unavailable";
        settings.Add(new($"{label} canonical SHA-256",
            string.IsNullOrWhiteSpace(version.CanonicalSha256)
                ? evidenceMessage
                : version.CanonicalSha256));
        settings.Add(new($"{label} canonical byte length",
            version.CanonicalByteLength?.ToString("N0", CultureInfo.InvariantCulture)
                ?? evidenceMessage));
    }

    private static string Interval(CashStatementCategoryTreeVersion version) =>
        $"{version.EffectiveFromOn:yyyy-MM-dd} – "
        + (version.EffectiveToOn is null ? "present" : $"{version.EffectiveToOn:yyyy-MM-dd}");

    private static string PeriodLabel(StatementPeriod period) => period.Id.IsAnnual
        ? $"{period.YearLabel} annual"
        : $"{period.YearLabel} {period.MonthLabel}";
}
