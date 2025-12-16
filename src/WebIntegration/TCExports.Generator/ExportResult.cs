namespace TCExports.Generator;

public class ExportResult
{
    public string Status { get; set; } = "success";   // "success" or "error"
    public string? Message { get; set; }              // error details if any
    public string? FileName { get; set; }             // safe filename
    public string? FileContent { get; set; }          // base64 encoded file bytes
}
