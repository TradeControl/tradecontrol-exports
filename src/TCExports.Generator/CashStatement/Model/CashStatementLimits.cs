namespace TCExports.Generator.CashStatement.Model;

/// <summary>
/// Initial admission limits for model construction and rendering. They are deliberately
/// above the Phase 0 baselines but remain bounded for multi-tenant service operation.
/// </summary>
public sealed record CashStatementLimits(
    int MaxFiscalYears,
    int MaxPeriodColumns,
    int MaxSheets,
    int MaxSections,
    int MaxRows,
    int MaxCells,
    int MaxFormulaNodes,
    int MaxDependenciesPerCell,
    int MaxDependencyDepth,
    long MaxOutputBytes)
{
    public static CashStatementLimits Default { get; } = new(
        MaxFiscalYears: 20,
        MaxPeriodColumns: 260,
        MaxSheets: 3,
        MaxSections: 256,
        MaxRows: 10_000,
        MaxCells: 2_000_000,
        MaxFormulaNodes: 1_000_000,
        MaxDependenciesPerCell: 1_024,
        MaxDependencyDepth: 256,
        MaxOutputBytes: 64L * 1024L * 1024L);
}
