using Microsoft.Extensions.DependencyInjection;
using TCExports.Generator.Contracts;
using TCExports.Generator.Handlers;
using TCExports.Generator.Validation;

namespace TCExports.Generator;

public static class ExportRunner
{
    public static async Task<ExportResult> ExportDataAsync(ExportPayload payload)
    {
        if (!GenericPayloadValidator.TryValidate(payload, out var errors))
        {
            return new ExportResult
            {
                Status = "error",
                Code = "VALIDATION_ERROR",
                Message = "Payload validation failed.",
                Details = errors
            };
        }

        var services = new ServiceCollection()

            // format-specific handlers used internally by TestHandler
            .AddSingleton<TestCsvHandler>()
            .AddSingleton<TestExcelHandler>()
            .AddSingleton<TestLibreHandler>()

            // single meta handler for 'test'
            .AddSingleton<IDocumentHandler, TestHandler>()

            // document-specific validator(s)
            .AddSingleton<IPayloadValidator, TestPayloadValidator>()

            // single handler stub for future statements
            .AddSingleton<IDocumentHandler, CashStatementHandler>()

            .AddSingleton<IExportEngine, ExportEngine>()
            .BuildServiceProvider();

        var engine = services.GetRequiredService<IExportEngine>();
        return await engine.GenerateAsync(payload);
    }
}