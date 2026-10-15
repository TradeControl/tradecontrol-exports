using System.Collections.Immutable;
using TCExports.Generator.CashStatement.Model;

namespace TCExports.Generator.CashStatement;

/// <summary>
/// Converts transport strings to the versioned internal request exactly once at the boundary.
/// </summary>
public static class CashStatementRequestParser
{
    private static readonly string[] BooleanNames =
    {
        "includeActivePeriods",
        "includeOrderBook",
        "includeTaxAccruals",
        "includeVatDetails",
        "includeBankBalances",
        "includeBalanceSheet",
        "includeReconciliation"
    };

    private static readonly HashSet<string> SupportedNames = new(
        BooleanNames.Append("commandTimeout"),
        StringComparer.OrdinalIgnoreCase);

    public static CashStatementContractResult<CashStatementRequestV1> Parse(
        IReadOnlyDictionary<string, string>? parameters)
    {
        var violations = ImmutableArray.CreateBuilder<StatementContractViolation>();
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        if (parameters is not null)
        {
            foreach (var pair in parameters)
            {
                if (!SupportedNames.Contains(pair.Key))
                {
                    violations.Add(new(
                        "OPTION_UNKNOWN",
                        $"params.{pair.Key}",
                        $"Unknown Cash Statement option '{pair.Key}'."));
                    continue;
                }

                if (!values.TryAdd(pair.Key, pair.Value))
                {
                    violations.Add(new(
                        "OPTION_DUPLICATE",
                        $"params.{pair.Key}",
                        $"Cash Statement option '{pair.Key}' was supplied more than once."));
                }
            }
        }

        var includeActivePeriods = ParseBoolean(values, "includeActivePeriods", false, violations);
        var includeOrderBook = ParseBoolean(values, "includeOrderBook", false, violations);
        var includeTaxAccruals = ParseBoolean(values, "includeTaxAccruals", true, violations);
        var includeVatDetails = ParseBoolean(values, "includeVatDetails", true, violations);
        var includeBankBalances = ParseBoolean(values, "includeBankBalances", true, violations);
        var includeBalanceSheet = ParseBoolean(values, "includeBalanceSheet", true, violations);
        var includeReconciliation = ParseBoolean(values, "includeReconciliation", false, violations);
        var commandTimeout = ParseTimeout(values, violations);

        if (violations.Count > 0)
            return new(null, violations.ToImmutable());

        return new(
            new CashStatementRequestV1(
                includeActivePeriods,
                includeOrderBook,
                includeTaxAccruals,
                includeVatDetails,
                includeBankBalances,
                includeBalanceSheet,
                includeReconciliation,
                commandTimeout),
            ImmutableArray<StatementContractViolation>.Empty);
    }

    private static bool ParseBoolean(
        IReadOnlyDictionary<string, string> values,
        string name,
        bool defaultValue,
        ImmutableArray<StatementContractViolation>.Builder violations)
    {
        if (!values.TryGetValue(name, out var raw) || string.IsNullOrWhiteSpace(raw))
            return defaultValue;

        var value = raw.Trim();
        if (bool.TryParse(value, out var parsed))
            return parsed;

        if (value is "1" || value.Equals("yes", StringComparison.OrdinalIgnoreCase) || value.Equals("y", StringComparison.OrdinalIgnoreCase))
            return true;

        if (value is "0" || value.Equals("no", StringComparison.OrdinalIgnoreCase) || value.Equals("n", StringComparison.OrdinalIgnoreCase))
            return false;

        violations.Add(new(
            "OPTION_BOOLEAN_INVALID",
            $"params.{name}",
            $"'{raw}' is not a valid Boolean value."));
        return defaultValue;
    }

    private static int ParseTimeout(
        IReadOnlyDictionary<string, string> values,
        ImmutableArray<StatementContractViolation>.Builder violations)
    {
        if (!values.TryGetValue("commandTimeout", out var raw) || string.IsNullOrWhiteSpace(raw))
            return 30;

        if (int.TryParse(raw, out var timeout) && timeout is >= 1 and <= 600)
            return timeout;

        violations.Add(new(
            "OPTION_TIMEOUT_INVALID",
            "params.commandTimeout",
            "Command timeout must be an integer from 1 through 600 seconds."));
        return 30;
    }
}
