using TCExports.Generator.Contracts;
using TCExports.Generator.Validation;

namespace TCExports.Generator.Handlers;

public sealed class TestHandler : IDocumentHandler
{
    private readonly IEnumerable<IPayloadValidator> _validators;
    private readonly TestCsvHandler _csv;
    private readonly TestExcelHandler _excel;
    private readonly TestLibreHandler _libre;

    public TestHandler(IEnumerable<IPayloadValidator> validators,
                       TestCsvHandler csv,
                       TestExcelHandler excel,
                       TestLibreHandler libre)
    {
        _validators = validators;
        _csv = csv;
        _excel = excel;
        _libre = libre;
    }

    public bool CanHandle(ExportPayload payload) =>
        payload.DocumentType.Equals("test", StringComparison.OrdinalIgnoreCase);

    public async Task<ExportResult> HandleAsync(ExportPayload payload, CancellationToken ct)
    {
        // Find and run document-specific validator
        var validator = _validators.FirstOrDefault(v => v.CanValidate(payload));
        if (validator is null)
        {
            return new ExportResult
            {
                Status = "error",
                Code = "VALIDATION_ERROR",
                Message = $"No validator registered for documentType '{payload.DocumentType}'."
            };
        }

        if (!validator.TryValidate(payload, out var errors))
        {
            return new ExportResult
            {
                Status = "error",
                Code = "VALIDATION_ERROR",
                Message = "Document-specific payload validation failed.",
                Details = errors
            };
        }

        validator.Normalize(payload);

        // Route by format
        if (payload.Format.Equals("csv", StringComparison.OrdinalIgnoreCase))
            return await _csv.HandleAsync(payload, ct);

        if (payload.Format.Equals("excel", StringComparison.OrdinalIgnoreCase))
            return await _excel.HandleAsync(payload, ct);

        if (payload.Format.Equals("libre", StringComparison.OrdinalIgnoreCase))
            return await _libre.HandleAsync(payload, ct);

        return new ExportResult
        {
            Status = "error",
            Code = "UNSUPPORTED_FORMAT",
            Message = $"Unsupported format '{payload.Format}' for documentType 'test'."
        };
    }
}