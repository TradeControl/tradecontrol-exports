namespace TCExports.Generator;

public class ExportResult
{
    public string Status { get; set; } = "success";   // success|error
    public string? Message { get; set; }              // human-friendly message
    public string? Code { get; set; }                 // e.g., VALIDATION_ERROR|EXECUTION_ERROR|UNSUPPORTED_FORMAT
    public string? FileName { get; set; }             // safe filename
    public string? FileContent { get; set; }          // base64 encoded file bytes
    // Field-level validation details: key = field path (e.g., "format", "params.includeOrderBook")
    public Dictionary<string, string[]>? Details { get; set; }
}
