using TCExports.Generator.Contracts;

namespace TCExports.Generator;

public interface IDocumentHandler
{
    bool CanHandle(ExportPayload payload);
    Task<ExportResult> HandleAsync(ExportPayload payload, CancellationToken ct);
}