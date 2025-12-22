using TCExports.Generator.Contracts;

namespace TCExports.Generator.Validation;

public sealed class CashPayloadValidator : IPayloadValidator
{
    private const string IncludeActivePeriods = "includeActivePeriods";
    private const string IncludeOrderBook = "includeOrderBook";
    private const string IncludeBankBalances = "includeBankBalances";
    private const string IncludeBankTypes = "includeBankTypes";
    private const string IncludeTaxAccruals = "includeTaxAccruals";
    private const string IncludeVatDetails = "includeVatDetails";
    private const string IncludeBalanceSheet = "includeBalanceSheet";
    private const string CommandTimeout = "commandTimeout";

    public bool CanValidate(ExportPayload payload) =>
        payload.DocumentType.Equals("cashflow", StringComparison.OrdinalIgnoreCase);

    // Implement interface: structural validation only
    public bool TryValidate(ExportPayload payload, out Dictionary<string, string[]> errors)
    {
        errors = new(StringComparer.OrdinalIgnoreCase);

        if (string.IsNullOrWhiteSpace(payload.SqlConnection))
            Add(errors, "sqlConnection", "required");

        if (!payload.Format.Equals("excel", StringComparison.OrdinalIgnoreCase))
            Add(errors, "format", "must be 'excel' for this slice");

        // For cashflow we expect spreadsheet; if FileType was removed, skip this check.
        // If you reintroduce it later, enforce: fileType == 'spreadsheet'.

        return errors.Count == 0;
    }

    // Implement interface: normalization only (apply defaults)
    public void Normalize(ExportPayload payload)
    {
        var p = payload.Params ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        NormalizeBool(p, IncludeActivePeriods, false);
        NormalizeBool(p, IncludeOrderBook, false);
        NormalizeBool(p, IncludeBankBalances, true);
        NormalizeBool(p, IncludeBankTypes, false);
        NormalizeBool(p, IncludeTaxAccruals, true);
        NormalizeBool(p, IncludeVatDetails, true);
        NormalizeBool(p, IncludeBalanceSheet, true);

        NormalizeInt(p, CommandTimeout, 30, minValue: 1, maxValue: 600);

        payload.Params = p;
    }

    // Convenience: combined flow for validators that support normalize-first
    public bool TryValidateAndNormalize(ExportPayload payload, out Dictionary<string, string[]> errors)
    {
        Normalize(payload);
        return TryValidate(payload, out errors);
    }

    private static void Add(Dictionary<string, string[]> dict, string key, string message)
    {
        if (!dict.TryGetValue(key, out var arr))
            dict[key] = new[] { message };
        else
            dict[key] = arr.Concat(new[] { message }).ToArray();
    }

    private static void NormalizeBool(Dictionary<string, string> p, string key, bool defaultValue)
    {
        if (!p.TryGetValue(key, out var raw) || string.IsNullOrWhiteSpace(raw))
        {
            p[key] = defaultValue ? "true" : "false";
            return;
        }
        if (!TryParseBool(raw, out var b))
        {
            p[key] = defaultValue ? "true" : "false";
            return;
        }
        p[key] = b ? "true" : "false";
    }

    private static void NormalizeInt(Dictionary<string, string> p, string key, int defaultValue, int minValue, int maxValue)
    {
        if (!p.TryGetValue(key, out var raw) || string.IsNullOrWhiteSpace(raw))
        {
            p[key] = defaultValue.ToString();
            return;
        }
        if (!int.TryParse(raw, out var v) || v < minValue || v > maxValue)
        {
            p[key] = defaultValue.ToString();
            return;
        }
        p[key] = v.ToString();
    }

    private static bool TryParseBool(string? value, out bool result)
    {
        if (bool.TryParse(value, out result)) return true;
        if (value is null) { result = false; return false; }
        var v = value.Trim().ToLowerInvariant();
        if (v is "1" or "yes" or "y" or "true") { result = true; return true; }
        if (v is "0" or "no" or "n" or "false") { result = false; return true; }
        result = false;
        return false;
    }
}