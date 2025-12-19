using System.Globalization;
using TCExports.Generator.Contracts;

namespace TCExports.Generator.Validation;

public static class ExportPayloadValidator
{
    private static readonly HashSet<string> AllowedDocumentTypes = new(StringComparer.OrdinalIgnoreCase)
    { "cashflow", "budget", "vat", "generic", "test" };

    private static readonly HashSet<string> AllowedFileTypes = new(StringComparer.OrdinalIgnoreCase)
    { "pdf", "spreadsheet" };

    private static readonly HashSet<string> AllowedFormats = new(StringComparer.OrdinalIgnoreCase)
    { "libre", "excel", "csv" };

    private static readonly Dictionary<string, HashSet<string>> AllowedParamsByDoc =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["cashflow"] = new(StringComparer.OrdinalIgnoreCase)
            {
                "includeActivePeriods","includeOrderBook","includeBankBalances","includeBankTypes",
                "includeTaxAccruals","includeVatDetails","includeBalanceSheet","greyscale"
            },
            ["budget"] = new(StringComparer.OrdinalIgnoreCase)
            {
                "includeActivePeriods","includeOrderBook","greyscale"
            },
            ["vat"] = new(StringComparer.OrdinalIgnoreCase)
            {
                "includeActivePeriods","includeTaxAccruals","greyscale"
            },
            ["generic"] = new(StringComparer.OrdinalIgnoreCase) { },
            ["test"] = new(StringComparer.OrdinalIgnoreCase) { "sqlQuery" }
        };

    public static bool TryValidate(ExportPayload payload, out Dictionary<string, string[]> errors)
    {
        errors = new(StringComparer.OrdinalIgnoreCase);

        if (string.IsNullOrWhiteSpace(payload.SqlConnection))
            Add(errors, "sqlConnection", "required");
        if (string.IsNullOrWhiteSpace(payload.UserName))
            Add(errors, "userName", "required");
        if (string.IsNullOrWhiteSpace(payload.DocumentType))
            Add(errors, "documentType", "required");
        if (string.IsNullOrWhiteSpace(payload.FileType))
            Add(errors, "fileType", "required");
        if (string.IsNullOrWhiteSpace(payload.Format))
            Add(errors, "format", "required");

        if (!string.IsNullOrWhiteSpace(payload.DocumentType) && !AllowedDocumentTypes.Contains(payload.DocumentType))
            Add(errors, "documentType", $"unsupported '{payload.DocumentType}'");
        if (!string.IsNullOrWhiteSpace(payload.FileType) && !AllowedFileTypes.Contains(payload.FileType))
            Add(errors, "fileType", $"unsupported '{payload.FileType}'");
        if (!string.IsNullOrWhiteSpace(payload.Format) && !AllowedFormats.Contains(payload.Format))
            Add(errors, "format", $"unsupported '{payload.Format}'");

        if (EqualsIgnoreCase(payload.FileType, "pdf") && EqualsIgnoreCase(payload.Format, "csv"))
            Add(errors, "format", "csv is not valid when fileType=pdf");

        if (!string.IsNullOrWhiteSpace(payload.DocumentType) &&
            AllowedParamsByDoc.TryGetValue(payload.DocumentType, out var allowedParams))
        {
            foreach (var kvp in payload.Params)
            {
                var key = kvp.Key ?? string.Empty;
                var value = kvp.Value;

                if (!allowedParams.Contains(key))
                {
                    Add(errors, $"params.{key}", $"not allowed for documentType '{payload.DocumentType}'");
                    continue;
                }

                // test: sqlQuery must be non-empty
                if (EqualsIgnoreCase(payload.DocumentType, "test") && key.Equals("sqlQuery", StringComparison.OrdinalIgnoreCase))
                {
                    if (string.IsNullOrWhiteSpace(value))
                        Add(errors, "params.sqlQuery", "required and must be non-empty");
                    continue;
                }

                // boolean-like params for other doc types
                if (payload.DocumentType is "cashflow" or "budget" or "vat")
                {
                    if (!TryParseBool(value, out _))
                        Add(errors, $"params.{key}", "must be boolean (true/false, 1/0, yes/no)");
                }
            }

            // Ensure required params for test
            if (EqualsIgnoreCase(payload.DocumentType, "test") && !payload.Params.ContainsKey("sqlQuery"))
                Add(errors, "params.sqlQuery", "required");
        }

        return errors.Count == 0;
    }

    private static void Add(Dictionary<string, string[]> dict, string key, string message)
    {
        if (!dict.TryGetValue(key, out var arr))
            dict[key] = new[] { message };
        else
            dict[key] = arr.Concat(new[] { message }).ToArray();
    }

    private static bool EqualsIgnoreCase(string? a, string? b) =>
        string.Equals(a, b, StringComparison.OrdinalIgnoreCase);

    private static bool TryParseBool(string? value, out bool result)
    {
        if (bool.TryParse(value, out result)) return true;
        if (string.Equals(value, "1", StringComparison.Ordinal)) { result = true; return true; }
        if (string.Equals(value, "0", StringComparison.Ordinal)) { result = false; return true; }
        if (value is not null)
        {
            var v = value.Trim().ToLowerInvariant();
            if (v is "yes" or "y") { result = true; return true; }
            if (v is "no" or "n") { result = false; return true; }
        }
        result = false;
        return false;
    }
}