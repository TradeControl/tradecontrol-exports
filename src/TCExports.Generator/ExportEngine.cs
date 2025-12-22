using System.Linq;
using TCExports.Generator.Contracts;
using TCExports.Generator.Validation;

namespace TCExports.Generator;

public sealed class ExportEngine : IExportEngine
{
    private readonly IEnumerable<IDocumentHandler> _handlers;
    private readonly IEnumerable<IPayloadValidator> _validators;

    public ExportEngine(IEnumerable<IDocumentHandler> handlers, IEnumerable<IPayloadValidator> validators)
    {
        _handlers = handlers;
        _validators = validators;
    }

    public async Task<ExportResult> GenerateAsync(ExportPayload payload, CancellationToken ct = default)
    {
        // Run all matching validators (normalize first, then validate)
        var aggregated = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

        foreach (var v in _validators)
        {
            if (!v.CanValidate(payload)) continue;

            v.Normalize(payload);

            if (!v.TryValidate(payload, out var errors)) continue;

            // accumulate any returned errors
            foreach (var kvp in errors)
            {
                if (!aggregated.TryGetValue(kvp.Key, out var list))
                    aggregated[kvp.Key] = new List<string>(kvp.Value);
                else
                    list.AddRange(kvp.Value);
            }
        }

        if (aggregated.Count > 0)
        {
            return new ExportResult
            {
                Status = "error",
                Code = "VALIDATION_ERROR",
                Message = "Payload validation failed.",
                Details = aggregated.ToDictionary(kvp => kvp.Key, kvp => kvp.Value.ToArray())
            };
        }

        // Dispatch to a handler
        var handler = _handlers.FirstOrDefault(h => h.CanHandle(payload));
        if (handler is null)
        {
            return new ExportResult
            {
                Status = "error",
                Code = "UNSUPPORTED_COMBINATION",
                Message = $"No handler for documentType='{payload.DocumentType}', format='{payload.Format}'."
            };
        }

        return await handler.HandleAsync(payload, ct);
    }
}