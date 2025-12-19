using TCExports.Generator.Contracts;

namespace TCExports.Generator.Validation;

public static class GenericPayloadValidator
{
    private static readonly HashSet<string> AllowedFormats = new(StringComparer.OrdinalIgnoreCase)
    { "libre", "excel", "csv" };

    private static readonly HashSet<string> AllowedRootFields = new(StringComparer.OrdinalIgnoreCase)
    { "sqlConnection", "userName", "documentType", "format", "params" };

    public static bool TryValidate(ExportPayload payload, out Dictionary<string, string[]> errors)
    {
        errors = new(StringComparer.OrdinalIgnoreCase);

        // Required root fields
        if (string.IsNullOrWhiteSpace(payload.SqlConnection))
            Add(errors, "sqlConnection", "required");
        if (string.IsNullOrWhiteSpace(payload.UserName))
            Add(errors, "userName", "required");
        if (string.IsNullOrWhiteSpace(payload.DocumentType))
            Add(errors, "documentType", "required");
        if (string.IsNullOrWhiteSpace(payload.Format))
            Add(errors, "format", "required");

        // Format enumeration (document type itself is validated by per-document validator)
        if (!string.IsNullOrWhiteSpace(payload.Format) && !AllowedFormats.Contains(payload.Format))
            Add(errors, "format", $"unsupported '{payload.Format}'");

        // Unknown root fields
        if (payload.Extra is not null && payload.Extra.Count > 0)
        {
            foreach (var key in payload.Extra.Keys)
            {
                if (!AllowedRootFields.Contains(key))
                    Add(errors, key, "not allowed at root");
            }
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
}