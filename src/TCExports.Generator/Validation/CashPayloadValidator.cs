using TCExports.Generator.CashStatement;
using TCExports.Generator.Contracts;

namespace TCExports.Generator.Validation;

public sealed class CashPayloadValidator : IPayloadValidator
{
    private static readonly IReadOnlyDictionary<string, bool> BooleanDefaults =
        new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase)
        {
            ["includeActivePeriods"] = false,
            ["includeOrderBook"] = false,
            ["includeTaxAccruals"] = true,
            ["includeVatDetails"] = true,
            ["includeBankBalances"] = true,
            ["includeBalanceSheet"] = true,
            ["includeReconciliation"] = false
        };

    public bool CanValidate(ExportPayload payload) =>
        payload.DocumentType.Equals("cashflow", StringComparison.OrdinalIgnoreCase);

    public bool TryValidate(ExportPayload payload, out Dictionary<string, string[]> errors)
    {
        errors = new(StringComparer.OrdinalIgnoreCase);

        if (string.IsNullOrWhiteSpace(payload.SqlConnection))
            Add(errors, "sqlConnection", "required");

        if (payload.Format is null ||
            !(payload.Format.Equals("excel", StringComparison.OrdinalIgnoreCase) ||
              payload.Format.Equals("libre", StringComparison.OrdinalIgnoreCase)))
        {
            Add(errors, "format", "must be 'excel' or 'libre'");
        }

        var request = CashStatementRequestParser.Parse(payload.Params);
        foreach (var violation in request.Violations)
            Add(errors, violation.Path, $"{violation.Code}: {violation.Message}");

        return errors.Count == 0;
    }

    public void Normalize(ExportPayload payload)
    {
        var parameters = payload.Params ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var item in BooleanDefaults)
        {
            if (!parameters.TryGetValue(item.Key, out var raw) || string.IsNullOrWhiteSpace(raw))
                parameters[item.Key] = item.Value ? "true" : "false";
            else if (TryParseBoolean(raw, out var value))
                parameters[item.Key] = value ? "true" : "false";
        }

        if (!parameters.TryGetValue("commandTimeout", out var timeout) || string.IsNullOrWhiteSpace(timeout))
            parameters["commandTimeout"] = "30";

        payload.Params = parameters;
    }

    private static void Add(Dictionary<string, string[]> errors, string key, string message)
    {
        if (!errors.TryGetValue(key, out var existing))
            errors[key] = new[] { message };
        else
            errors[key] = existing.Append(message).ToArray();
    }

    private static bool TryParseBoolean(string value, out bool result)
    {
        if (bool.TryParse(value, out result))
            return true;

        switch (value.Trim().ToLowerInvariant())
        {
            case "1":
            case "yes":
            case "y":
                result = true;
                return true;
            case "0":
            case "no":
            case "n":
                result = false;
                return true;
            default:
                result = false;
                return false;
        }
    }
}
