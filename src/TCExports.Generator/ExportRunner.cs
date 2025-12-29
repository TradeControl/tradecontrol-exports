using Microsoft.Extensions.DependencyInjection;
using TCExports.Generator.Contracts;
using TCExports.Generator.Data;
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

            #region test 
            .AddSingleton<IDocumentHandler, TestHandler>()
            .AddSingleton<TestCsvHandler>()
            .AddSingleton<TestExcelHandler>()
            .AddSingleton<TestLibreHandler>()            
            .AddSingleton<IPayloadValidator, TestPayloadValidator>()
            #endregion

            #region cash flow statements
            // Routing handler for cashflow (delegates to excel/libre handlers)
            .AddSingleton<IDocumentHandler, CashStatementHandler>()
            // Excel implementation + its repository
            .AddSingleton<CashStatementExcelHandler>()
            .AddSingleton<ICashFlowRepository, SqlServerCashFlowRepository>()
            // Libre implementation (shells out to Python)
            .AddSingleton<CashStatementLibreHandler>()
            .AddSingleton<IPayloadValidator, CashPayloadValidator>()
            #endregion

            .AddSingleton<IExportEngine, ExportEngine>()
            .BuildServiceProvider();

        var engine = services.GetRequiredService<IExportEngine>();
        return await engine.GenerateAsync(payload);
    }
}