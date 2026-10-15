using System.Collections.Immutable;

namespace TCExports.Generator.CashStatement.Model;

public sealed record PhysicalCellAddress(string SheetName, int RowNumber, int ColumnNumber);

/// <summary>
/// Renderer-owned mapping from semantic addresses to physical worksheet coordinates.
/// </summary>
public sealed record StatementLayout(
    ImmutableDictionary<StatementCellAddress, PhysicalCellAddress> Cells);

public interface ICashStatementLayoutAllocator
{
    StatementLayout Allocate(CashStatementDocument document, CashStatementLimits limits);
}

public interface ICashStatementFormulaTranslator
{
    string Translate(StatementFormula formula, StatementLayout layout);
}

public sealed record StatementContractViolation(
    string Code,
    string Path,
    string Message);

public sealed record CashStatementContractResult<T>(
    T? Value,
    ImmutableArray<StatementContractViolation> Violations)
    where T : class
{
    public bool IsSuccess => Value is not null && Violations.IsEmpty;
}

public interface ICashStatementContractValidator
{
    ImmutableArray<StatementContractViolation> Validate(
        CashStatementDocument document,
        CashStatementLimits limits);
}
