using TCExports.Generator.Contracts;

namespace TCExports.Generator.Handlers;

public sealed class CashStatementHandler : IDocumentHandler
{
    private readonly CashStatementExcelHandler _excel;
    private readonly CashStatementLibreHandler _libre;

    public CashStatementHandler(
        CashStatementExcelHandler excel,
        CashStatementLibreHandler libre)
    {
        _excel = excel;
        _libre = libre;
    }

    public bool CanHandle(ExportPayload payload) =>
        payload.DocumentType.Equals("cashflow", StringComparison.OrdinalIgnoreCase);

    public async Task<ExportResult> HandleAsync(ExportPayload payload, CancellationToken ct)
    {
        if (payload.Format.Equals("excel", StringComparison.OrdinalIgnoreCase))
            return await _excel.HandleAsync(payload, ct);

        if (payload.Format.Equals("libre", StringComparison.OrdinalIgnoreCase))
            return await _libre.HandleAsync(payload, ct);

        return new ExportResult
        {
            Status = "error",
            Code = "UNSUPPORTED_FORMAT",
            Message = $"Unsupported format '{payload.Format}' for documentType 'cashflow'."
        };
    }
}