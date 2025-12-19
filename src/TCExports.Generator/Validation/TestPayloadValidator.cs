using TCExports.Generator.Contracts;

namespace TCExports.Generator.Validation;

public sealed class TestPayloadValidator : IPayloadValidator
{
    private static readonly HashSet<string> AllowedParams = new(StringComparer.OrdinalIgnoreCase)
    { "sqlQuery" };

    public bool CanValidate(ExportPayload payload) =>
        payload.DocumentType.Equals("test", StringComparison.OrdinalIgnoreCase);

    public bool TryValidate(ExportPayload payload, out Dictionary<string, string[]> errors)
    {
        errors = new(StringComparer.OrdinalIgnoreCase);

        // Strict param whitelist
        foreach (var key in payload.Params.Keys)
        {
            if (!AllowedParams.Contains(key))
                Add(errors, $"params.{key}", "not allowed for documentType 'test'");
        }

        // Require sqlQuery and non-empty
        if (!payload.Params.TryGetValue("sqlQuery", out var sql) || string.IsNullOrWhiteSpace(sql))
            Add(errors, "params.sqlQuery", "required and must be non-empty");

        return errors.Count == 0;
    }

    public void Normalize(ExportPayload payload)
    {
        // No optional defaults for test at the moment.
    }

    private static void Add(Dictionary<string, string[]> dict, string key, string message)
    {
        if (!dict.TryGetValue(key, out var arr))
            dict[key] = new[] { message };
        else
            dict[key] = arr.Concat(new[] { message }).ToArray();
    }
}