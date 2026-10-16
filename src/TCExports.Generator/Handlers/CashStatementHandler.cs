using TCExports.Generator.CashStatement;
using TCExports.Generator.CashStatement.Model;
using TCExports.Generator.Contracts;
using TCExports.Generator.Excel;
using TCExports.Generator.Ods;

namespace TCExports.Generator.Handlers;

/// <summary>
/// Generates either supported Cash Statement spreadsheet format from one shared,
/// renderer-neutral document model.
/// </summary>
public sealed class CashStatementHandler : IDocumentHandler
{
    private readonly CashStatementBuilder _builder;
    private readonly CashStatementXlsxRenderer _xlsxRenderer;
    private readonly CashStatementOdsRenderer _odsRenderer;

    public CashStatementHandler(
        CashStatementBuilder builder,
        CashStatementXlsxRenderer xlsxRenderer,
        CashStatementOdsRenderer odsRenderer)
    {
        _builder = builder;
        _xlsxRenderer = xlsxRenderer;
        _odsRenderer = odsRenderer;
    }

    public bool CanHandle(ExportPayload payload) =>
        payload.DocumentType.Equals("cashflow", StringComparison.OrdinalIgnoreCase);

    public async Task<ExportResult> HandleAsync(ExportPayload payload, CancellationToken ct)
    {
        var format = (payload.Format ?? string.Empty).Trim().ToLowerInvariant();
        if (format is not ("excel" or "libre"))
        {
            return new ExportResult
            {
                Status = "error",
                Code = "UNSUPPORTED_FORMAT",
                Message = $"Unsupported format '{payload.Format}' for documentType 'cashflow'."
            };
        }

        var request = CashStatementRequestParser.Parse(payload.Params);
        if (!request.IsSuccess)
            return ContractFailure("VALIDATION_ERROR", "Cash Statement options are invalid.", request.Violations);

        try
        {
            var build = await _builder.BuildAsync(
                payload.SqlConnection,
                request.Value!,
                payload.UserName,
                DateTimeOffset.UtcNow,
                ct).ConfigureAwait(false);

            if (!build.IsSuccess)
            {
                return ContractFailure(
                    "CASH_STATEMENT_INVALID",
                    "Cash Statement generation did not satisfy its document contract.",
                    build.Violations);
            }

            byte[] bytes;
            string extension;
            if (format == "excel")
            {
                bytes = _xlsxRenderer.Render(build.Value!, ct);
                extension = "xlsx";
            }
            else
            {
                bytes = await _odsRenderer.RenderAsync(build.Value!, ct).ConfigureAwait(false);
                extension = "ods";
            }

            var timestamp = DateTime.UtcNow.ToString("yyyyMMdd_HHmmss");
            var safeUser = SanitizeFilePart(payload.UserName);
            return new ExportResult
            {
                Status = "success",
                FileName = $"{safeUser}_cashflow_{timestamp}.{extension}",
                FileContent = Convert.ToBase64String(bytes)
            };
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            return new ExportResult
            {
                Status = "error",
                Code = "EXECUTION_ERROR",
                Message = "Cash Statement generation failed.",
                Details = new Dictionary<string, string[]>
                {
                    ["generation"] = new[] { ex.Message }
                }
            };
        }
    }

    private static ExportResult ContractFailure(
        string code,
        string message,
        IEnumerable<StatementContractViolation> violations) =>
        new()
        {
            Status = "error",
            Code = code,
            Message = message,
            Details = violations
                .GroupBy(violation => violation.Path, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(
                    group => group.Key,
                    group => group.Select(violation => $"{violation.Code}: {violation.Message}").ToArray(),
                    StringComparer.OrdinalIgnoreCase)
        };

    private static string SanitizeFilePart(string value)
    {
        var sanitized = System.Text.RegularExpressions.Regex.Replace(
            value,
            @"[^A-Za-z0-9_\-]+",
            "_");
        return string.IsNullOrWhiteSpace(sanitized) ? "cash-statement" : sanitized;
    }
}
