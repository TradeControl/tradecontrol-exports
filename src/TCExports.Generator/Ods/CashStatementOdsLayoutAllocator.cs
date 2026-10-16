using System.Collections.Immutable;
using TCExports.Generator.CashStatement.Model;

namespace TCExports.Generator.Ods;

/// <summary>
/// Allocates every semantic Cash Statement cell before the ODS package writer translates formulas.
/// </summary>
public sealed class CashStatementOdsLayoutAllocator : ICashStatementLayoutAllocator
{
    private static readonly IReadOnlyDictionary<string, int> ReconciliationColumns =
        new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["year"] = 1, ["description"] = 2, ["opening-capital"] = 3,
            ["closing-capital"] = 4, ["profit"] = 5, ["business-tax"] = 6,
            ["profit-after-tax"] = 7, ["capital-movement"] = 8,
            ["opening-subject-position"] = 9, ["opening-account-position"] = 10,
            ["capital-delta"] = 11, ["bridge-total"] = 13, ["difference"] = 14,
            ["status"] = 15
        };

    public StatementLayout Allocate(CashStatementDocument document, CashStatementLimits limits)
    {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(limits);
        var cells = ImmutableDictionary.CreateBuilder<StatementCellAddress, PhysicalCellAddress>();

        foreach (var sheet in document.Sheets.OrderBy(sheet => sheet.PresentationOrder))
        {
            if (sheet.Kind == StatementSheetKind.CashFlow)
                AllocateCashFlow(sheet, cells);
            else if (sheet.Kind == StatementSheetKind.Reconciliation)
                AllocateReconciliation(sheet, cells);
            else if (sheet.Kind == StatementSheetKind.Settings && sheet.Sections.Any(section => !section.Rows.IsEmpty))
                throw new InvalidOperationException("The ODS Settings sheet is renderer-owned.");
        }

        if (cells.Count > limits.MaxCells)
            throw new InvalidOperationException($"ODS layout contains {cells.Count} cells; limit is {limits.MaxCells}.");
        return new StatementLayout(cells.ToImmutable());
    }

    private static void AllocateCashFlow(
        StatementSheet sheet,
        ImmutableDictionary<StatementCellAddress, PhysicalCellAddress>.Builder cells)
    {
        var columns = sheet.Columns.ToDictionary(column => column.Id, column => column.PresentationOrder + 1);
        var currentRow = 4;
        foreach (var section in sheet.Sections.OrderBy(section => section.PresentationOrder))
        {
            if (section.Kind == StatementSectionKind.FlowCategories)
            {
                foreach (var row in section.Rows.OrderBy(row => row.PresentationOrder))
                {
                    if (row.Kind == StatementRowKind.Header)
                    {
                        currentRow += 2;
                        continue;
                    }
                    currentRow++;
                    AllocateRow(sheet, row, currentRow, columns, cells);
                }
                continue;
            }

            currentRow += 2;
            foreach (var row in section.Rows.OrderBy(row => row.PresentationOrder))
            {
                currentRow++;
                AllocateRow(sheet, row, currentRow, columns, cells);
            }
        }
    }

    private static void AllocateReconciliation(
        StatementSheet sheet,
        ImmutableDictionary<StatementCellAddress, PhysicalCellAddress>.Builder cells)
    {
        var dataRow = 5;
        foreach (var row in sheet.Sections.OrderBy(section => section.PresentationOrder)
                     .SelectMany(section => section.Rows.OrderBy(row => row.PresentationOrder)))
        {
            foreach (var pair in row.Cells)
            {
                var column = row.Kind == StatementRowKind.Metadata
                    ? pair.Key.Value switch
                    {
                        "tolerance" => 2,
                        "status" => 5,
                        _ => throw new InvalidOperationException($"Unsupported reconciliation summary column '{pair.Key.Value}'.")
                    }
                    : ReconciliationColumns.TryGetValue(pair.Key.Value, out var mapped)
                        ? mapped
                        : throw new InvalidOperationException($"Unsupported reconciliation column '{pair.Key.Value}'.");
                Add(sheet, row, pair.Key, row.Kind == StatementRowKind.Metadata ? 2 : dataRow, column, cells);
            }
            if (row.Kind != StatementRowKind.Metadata)
                dataRow++;
        }
    }

    private static void AllocateRow(
        StatementSheet sheet,
        StatementRow row,
        int rowNumber,
        IReadOnlyDictionary<StatementColumnId, int> columns,
        ImmutableDictionary<StatementCellAddress, PhysicalCellAddress>.Builder cells)
    {
        foreach (var pair in row.Cells)
        {
            if (!columns.TryGetValue(pair.Key, out var column))
                throw new InvalidOperationException($"Column '{pair.Key}' is absent from sheet '{sheet.Id}'.");
            Add(sheet, row, pair.Key, rowNumber, column, cells);
        }
    }

    private static void Add(
        StatementSheet sheet,
        StatementRow row,
        StatementColumnId column,
        int rowNumber,
        int columnNumber,
        ImmutableDictionary<StatementCellAddress, PhysicalCellAddress>.Builder cells)
    {
        var semantic = new StatementCellAddress(sheet.Id, row.Id, column);
        if (!cells.TryAdd(semantic, new PhysicalCellAddress(sheet.Name, rowNumber, columnNumber)))
            throw new InvalidOperationException($"Duplicate ODS layout address '{semantic}'.");
    }
}
