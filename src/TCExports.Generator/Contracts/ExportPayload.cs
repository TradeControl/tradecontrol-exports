namespace TCExports.Generator.Contracts;

public class ExportPayload
{
    public string SqlConnection { get; set; } = string.Empty;
    public string UserName { get; set; } = string.Empty;

    // Enumerated in validation: cashflow|budget|vat|generic
    public string DocumentType { get; set; } = string.Empty;

    // Enumerated in validation: pdf|spreadsheet
    public string FileType { get; set; } = "spreadsheet";

    // Enumerated in validation: libre|xls|csv
    public string Format { get; set; } = "libre";

    // Document-type-specific flags (validated per document type)
    public Dictionary<string, string> Params { get; set; } = new();
}
