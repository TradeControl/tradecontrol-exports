using Microsoft.AspNetCore.Mvc;
using TCExports.Generator; // shared models + runner

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
        var contentType = "text/csv"; // or derive from payload.FileType
        return File(bytes, contentType, result.FileName);
    }
}
