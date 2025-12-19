using Microsoft.AspNetCore.Mvc;
using TCExports.Generator; // engine entry point
using TCExports.Generator.Contracts;

namespace TCExports.WebHarness.Controllers;

[ApiController]
[Route("api/export")]
public class ExportController : ControllerBase
{
    [HttpPost]
    public async Task<IActionResult> Export([FromBody] ExportPayload payload)
    {
        var result = await ExportRunner.ExportDataAsync(payload);

        if (result.Status != "success")
            return BadRequest(result);

        var bytes = Convert.FromBase64String(result.FileContent!);
        var contentType = GetContentType(result.FileName);
        return File(bytes, contentType, result.FileName);
    }

    private static string GetContentType(string? fileName)
    {
        var ext = Path.GetExtension(fileName ?? string.Empty).ToLowerInvariant();
        return ext switch
        {
            ".csv" => "text/csv",
            ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ".ods" => "application/vnd.oasis.opendocument.spreadsheet",
            ".pdf" => "application/pdf",
            _ => "application/octet-stream"
        };
    }
}