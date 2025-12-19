using System.Reflection.Metadata;
using TCExports.Generator.Contracts;

namespace TCExports.Generator;

public sealed class ExportEngine : IExportEngine
{
    private readonly IEnumerable<IDocumentHandler> _handlers;

    public ExportEngine(IEnumerable<IDocumentHandler> handlers)
    {
        _handlers = handlers;
    }

    public async Task<ExportResult> GenerateAsync(ExportPayload payload, CancellationToken ct = default)
    {
        foreach (var h in _handlers)
        {
            if (h.CanHandle(payload))
                return await h.HandleAsync(payload, ct);
        }

        return new ExportResult
        {
            Status = "error",
            Code = "UNSUPPORTED_COMBINATION",
            Message = $"No handler for documentType='{payload.DocumentType}', format='{payload.Format}'."
        };
    }
}