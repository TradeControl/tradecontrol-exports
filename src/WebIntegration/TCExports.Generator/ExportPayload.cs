namespace TCExports.Generator;

public class ExportPayload
{
    public string SqlConnection { get; set; } = string.Empty;
    public string UserName { get; set; } = string.Empty;
    public string DocumentType { get; set; } = string.Empty;
    public string FileType { get; set; } = "csv";
    public string Format { get; set; } = "Libre";
    public Dictionary<string, string> Params { get; set; } = new();
}
