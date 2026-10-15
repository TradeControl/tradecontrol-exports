using System.Collections.Immutable;
using TCExports.Generator.CashStatement.Model;

namespace TCExports.Generator.CashStatement;

/// <summary>
/// Replaces formula-cell placeholders with authoritative values calculated from
/// semantic dependencies. The document must be contract-valid before evaluation.
/// </summary>
public sealed class CashStatementDocumentEvaluator
{
    private readonly CashStatementFormulaEvaluator _formulaEvaluator = new();

    public CashStatementDocument Evaluate(CashStatementDocument document)
    {
        var sourceCells = document.Sheets
            .SelectMany(sheet => sheet.Sections.SelectMany(section =>
                section.Rows.SelectMany(row => row.Cells.Select(pair => new
                {
                    Address = new StatementCellAddress(sheet.Id, row.Id, pair.Key),
                    Cell = pair.Value
                }))))
            .ToDictionary(item => item.Address, item => item.Cell);

        var values = new Dictionary<StatementCellAddress, StatementValue>();
        var visiting = new HashSet<StatementCellAddress>();

        StatementValue EvaluateCell(StatementCellAddress address)
        {
            if (values.TryGetValue(address, out var value))
                return value;

            if (!sourceCells.TryGetValue(address, out var cell))
                throw new InvalidOperationException($"Referenced cell '{Format(address)}' does not exist.");

            if (!visiting.Add(address))
                throw new InvalidOperationException($"Dependency cycle reached at '{Format(address)}'.");

            value = cell.Formula is null
                ? cell.Value
                : _formulaEvaluator.Evaluate(cell.Formula, EvaluateCell);

            visiting.Remove(address);
            values[address] = value;
            return value;
        }

        foreach (var address in sourceCells.Keys)
            EvaluateCell(address);

        return document with
        {
            Sheets = document.Sheets.Select(sheet => sheet with
            {
                Sections = sheet.Sections.Select(section => section with
                {
                    Rows = section.Rows.Select(row => row with
                    {
                        Cells = row.Cells.ToImmutableDictionary(
                            pair => pair.Key,
                            pair => pair.Value with
                            {
                                Value = values[new StatementCellAddress(sheet.Id, row.Id, pair.Key)]
                            })
                    }).ToImmutableArray()
                }).ToImmutableArray()
            }).ToImmutableArray()
        };
    }

    private static string Format(StatementCellAddress address) =>
        $"{address.SheetId.Value}/{address.RowId.Value}/{address.ColumnId.Value}";
}
