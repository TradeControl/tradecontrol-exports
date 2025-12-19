using Microsoft.Extensions.DependencyInjection;
using TCExports.Generator.Contracts;
using TCExports.Generator.Handlers;
using TCExports.Generator.Validation;

namespace TCExports.Generator;

public static class ExportRunner
{
    public static async Task<ExportResult> ExportDataAsync(ExportPayload payload)
    {
        if (!ExportPayloadValidator.TryValidate(payload, out var errors))
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
            .AddSingleton<IDocumentHandler, TestCsvHandler>()
            .AddSingleton<IDocumentHandler, TestExcelHandler>() // new
            .AddSingleton<IExportEngine, ExportEngine>()
            .BuildServiceProvider();

        var engine = services.GetRequiredService<IExportEngine>();
        return await engine.GenerateAsync(payload);
    }
}