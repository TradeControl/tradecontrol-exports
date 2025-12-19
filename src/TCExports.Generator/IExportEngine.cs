using TCExports.Generator.Contracts;

namespace TCExports.Generator;

public interface IExportEngine
{
    Task<ExportResult> GenerateAsync(ExportPayload payload, CancellationToken ct = default);
}