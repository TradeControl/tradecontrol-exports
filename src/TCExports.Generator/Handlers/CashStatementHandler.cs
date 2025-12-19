using TCExports.Generator.Contracts;

namespace TCExports.Generator.Handlers;

public sealed class CashStatementHandler : IDocumentHandler
{
    public bool CanHandle(ExportPayload payload) =>
        payload.DocumentType.Equals("cashflow", StringComparison.OrdinalIgnoreCase);

    public Task<ExportResult> HandleAsync(ExportPayload payload, CancellationToken ct) =>
        Task.FromResult(new ExportResult
        {
            Status = "error",
            Code = "NOT_IMPLEMENTED",
            Message = "Cash Statement handler stub. Implementation to be added next."
        });
}