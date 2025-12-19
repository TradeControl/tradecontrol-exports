using TCExports.Generator.Contracts;

namespace TCExports.Generator.Validation;

public interface IPayloadValidator
{
    bool CanValidate(ExportPayload payload);
    bool TryValidate(ExportPayload payload, out Dictionary<string, string[]> errors);

    // Optional normalization (defaults for missing optional params)
    void Normalize(ExportPayload payload);
}