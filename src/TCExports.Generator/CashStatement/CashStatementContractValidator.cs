using System.Collections.Immutable;
using TCExports.Generator.CashStatement.Model;

namespace TCExports.Generator.CashStatement;

public sealed class CashStatementContractValidator : ICashStatementContractValidator
{
    public ImmutableArray<StatementContractViolation> Validate(
        CashStatementDocument document,
        CashStatementLimits limits)
    {
        var violations = ImmutableArray.CreateBuilder<StatementContractViolation>();
        var periodIds = ValidatePeriods(document.PeriodGrid, limits, violations);
        var cells = new Dictionary<StatementCellAddress, StatementCell>();
        var dependencyGraph = new Dictionary<StatementCellAddress, HashSet<StatementCellAddress>>();

        ValidateCount(document.Sheets.Length, limits.MaxSheets, "LIMIT_SHEETS", "sheets", violations);
        ValidateUniqueIds(document.Sheets.Select(s => s.Id.Value), "sheet", "sheets", violations);
        ValidateUniqueOrders(document.Sheets.Select(s => s.PresentationOrder), "SHEET_ORDER_DUPLICATE", "sheets", violations);

        var sectionIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var rowIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var sectionCount = 0;
        var rowCount = 0;
        var cellCount = 0;
        var formulaNodeCount = 0;

        foreach (var sheet in document.Sheets)
        {
            ValidateId(sheet.Id.Value, "sheet", $"sheets.{sheet.Id.Value}", violations);
            var sheetPath = $"sheets.{sheet.Id.Value}";

            ValidateUniqueIds(sheet.Columns.Select(c => c.Id.Value), "column", $"{sheetPath}.columns", violations);
            ValidateUniqueOrders(sheet.Columns.Select(c => c.PresentationOrder), "COLUMN_ORDER_DUPLICATE", $"{sheetPath}.columns", violations);
            ValidateUniqueOrders(sheet.Sections.Select(s => s.PresentationOrder), "SECTION_ORDER_DUPLICATE", $"{sheetPath}.sections", violations);

            var columns = new HashSet<StatementColumnId>();
            foreach (var column in sheet.Columns)
            {
                ValidateId(column.Id.Value, "column", $"{sheetPath}.columns.{column.Id.Value}", violations);
                columns.Add(column.Id);

                if (column.PeriodId is { } periodId && !periodIds.Contains(periodId))
                {
                    violations.Add(new(
                        "COLUMN_PERIOD_NOT_FOUND",
                        $"{sheetPath}.columns.{column.Id.Value}",
                        $"Column period '{periodId}' is not present in the period grid."));
                }
            }

            foreach (var section in sheet.Sections)
            {
                sectionCount++;
                var sectionPath = $"{sheetPath}.sections.{section.Id.Value}";
                ValidateId(section.Id.Value, "section", sectionPath, violations);
                if (!sectionIds.Add(section.Id.Value))
                    violations.Add(new("SECTION_ID_DUPLICATE", sectionPath, $"Section id '{section.Id.Value}' is not document-unique."));
                ValidateUniqueOrders(section.Rows.Select(r => r.PresentationOrder), "ROW_ORDER_DUPLICATE", $"{sectionPath}.rows", violations);

                foreach (var row in section.Rows)
                {
                    rowCount++;
                    var rowPath = $"{sectionPath}.rows.{row.Id.Value}";
                    ValidateId(row.Id.Value, "row", rowPath, violations);
                    if (!rowIds.Add(row.Id.Value))
                        violations.Add(new("ROW_ID_DUPLICATE", rowPath, $"Row id '{row.Id.Value}' is not document-unique."));

                    foreach (var pair in row.Cells)
                    {
                        cellCount++;
                        var address = new StatementCellAddress(sheet.Id, row.Id, pair.Key);
                        var cellPath = $"{rowPath}.cells.{pair.Key.Value}";

                        if (!columns.Contains(pair.Key))
                        {
                            violations.Add(new(
                                "CELL_COLUMN_NOT_FOUND",
                                cellPath,
                                $"Cell column '{pair.Key.Value}' is not declared on sheet '{sheet.Id.Value}'."));
                        }

                        if (!cells.TryAdd(address, pair.Value))
                            violations.Add(new("CELL_DUPLICATE", cellPath, "Semantic cell address occurs more than once."));

                        if (pair.Value.NumericKind is not null && pair.Value.Value.Kind != StatementScalarKind.Decimal)
                        {
                            violations.Add(new(
                                "CELL_NUMERIC_KIND_INVALID",
                                cellPath,
                                "A numeric kind may be assigned only to a decimal value."));
                        }

                        if (pair.Value.Formula is not null)
                        {
                            var dependencies = new HashSet<StatementCellAddress>();
                            formulaNodeCount += CollectFormula(
                                pair.Value.Formula,
                                dependencies,
                                cellPath,
                                violations);
                            dependencyGraph[address] = dependencies;

                            if (dependencies.Count > limits.MaxDependenciesPerCell)
                            {
                                violations.Add(new(
                                    "LIMIT_DEPENDENCIES_PER_CELL",
                                    cellPath,
                                    $"Dependency count {dependencies.Count} exceeds {limits.MaxDependenciesPerCell}."));
                            }
                        }
                    }
                }
            }
        }

        ValidateCount(sectionCount, limits.MaxSections, "LIMIT_SECTIONS", "sections", violations);
        ValidateCount(rowCount, limits.MaxRows, "LIMIT_ROWS", "rows", violations);
        ValidateCount(cellCount, limits.MaxCells, "LIMIT_CELLS", "cells", violations);
        ValidateCount(formulaNodeCount, limits.MaxFormulaNodes, "LIMIT_FORMULA_NODES", "formulas", violations);

        foreach (var pair in dependencyGraph)
        {
            var cellPath = Format(pair.Key);
            foreach (var dependency in pair.Value)
            {
                if (dependency == pair.Key)
                {
                    violations.Add(new("DEPENDENCY_SELF_REFERENCE", cellPath, $"Cell '{cellPath}' references itself."));
                }
                else if (!cells.ContainsKey(dependency))
                {
                    violations.Add(new(
                        "DEPENDENCY_NOT_FOUND",
                        cellPath,
                        $"Referenced cell '{Format(dependency)}' does not exist."));
                }
            }
        }

        ValidateDependencyCycles(dependencyGraph, cells, limits.MaxDependencyDepth, violations);
        return violations.ToImmutable();
    }

    private static HashSet<StatementPeriodId> ValidatePeriods(
        StatementPeriodGrid grid,
        CashStatementLimits limits,
        ImmutableArray<StatementContractViolation>.Builder violations)
    {
        ValidateCount(grid.Periods.Length, limits.MaxPeriodColumns, "LIMIT_PERIOD_COLUMNS", "periodGrid", violations);
        ValidateUniqueOrders(grid.Periods.Select(p => p.PresentationOrder), "PERIOD_ORDER_DUPLICATE", "periodGrid", violations);

        var ids = new HashSet<StatementPeriodId>();
        foreach (var period in grid.Periods)
        {
            if (!ids.Add(period.Id))
                violations.Add(new("PERIOD_ID_DUPLICATE", $"periodGrid.{period.Id}", $"Period '{period.Id}' occurs more than once."));
            if (period.Id.MonthNumber is < 1 or > 12)
                violations.Add(new("PERIOD_MONTH_INVALID", $"periodGrid.{period.Id}", "Monthly period must use month 1 through 12."));
        }

        return ids;
    }

    private static int CollectFormula(
        StatementFormula formula,
        HashSet<StatementCellAddress> dependencies,
        string path,
        ImmutableArray<StatementContractViolation>.Builder violations)
    {
        switch (formula)
        {
            case StatementFormulaReference reference:
                dependencies.Add(reference.Address);
                return 1;
            case StatementFormulaLiteral:
                return 1;
            case StatementFormulaUnary unary:
                return 1 + CollectFormula(unary.Operand, dependencies, path, violations);
            case StatementFormulaBinary binary:
                return 1 +
                    CollectFormula(binary.Left, dependencies, path, violations) +
                    CollectFormula(binary.Right, dependencies, path, violations);
            case StatementFormulaAggregate aggregate:
                if (aggregate.Operands.IsDefault)
                {
                    violations.Add(new("FORMULA_OPERANDS_DEFAULT", path, "Aggregate operands must be initialized."));
                    return 1;
                }
                return 1 + aggregate.Operands.Sum(o => CollectFormula(o, dependencies, path, violations));
            case StatementFormulaConditional conditional:
                return 1 +
                    CollectFormula(conditional.Condition, dependencies, path, violations) +
                    CollectFormula(conditional.WhenTrue, dependencies, path, violations) +
                    CollectFormula(conditional.WhenFalse, dependencies, path, violations);
            case StatementFormulaTerminal terminal:
                if (string.IsNullOrWhiteSpace(terminal.Expression))
                    violations.Add(new("TERMINAL_EXPRESSION_EMPTY", path, "Terminal expression must not be empty."));
                foreach (var reference in terminal.References.Values)
                    dependencies.Add(reference);
                return 1 + terminal.References.Count;
            default:
                violations.Add(new("FORMULA_NODE_UNKNOWN", path, $"Unknown formula node '{formula.GetType().Name}'."));
                return 1;
        }
    }

    private static void ValidateDependencyCycles(
        IReadOnlyDictionary<StatementCellAddress, HashSet<StatementCellAddress>> graph,
        IReadOnlyDictionary<StatementCellAddress, StatementCell> cells,
        int maxDepth,
        ImmutableArray<StatementContractViolation>.Builder violations)
    {
        var states = new Dictionary<StatementCellAddress, byte>();
        var stack = new List<StatementCellAddress>();
        var reportedCycles = new HashSet<string>(StringComparer.Ordinal);
        var reportedDepth = new HashSet<StatementCellAddress>();

        foreach (var node in graph.Keys)
            Visit(node, 1);

        void Visit(StatementCellAddress node, int depth)
        {
            if (states.TryGetValue(node, out var existing) && existing == 2)
                return;
            if (depth > maxDepth)
            {
                if (reportedDepth.Add(node))
                    violations.Add(new("LIMIT_DEPENDENCY_DEPTH", Format(node), $"Dependency depth exceeds {maxDepth}."));
                return;
            }

            states[node] = 1;
            stack.Add(node);

            if (graph.TryGetValue(node, out var dependencies))
            {
                foreach (var dependency in dependencies.Where(cells.ContainsKey))
                {
                    if (!states.TryGetValue(dependency, out var state))
                    {
                        Visit(dependency, depth + 1);
                    }
                    else if (state == 1)
                    {
                        var start = stack.FindIndex(item => item == dependency);
                        var cycle = stack.Skip(start).Append(dependency).Select(Format).ToArray();
                        var cycleText = string.Join(" -> ", cycle);
                        if (reportedCycles.Add(cycleText))
                            violations.Add(new("DEPENDENCY_CYCLE", Format(node), cycleText));
                    }
                }
            }

            stack.RemoveAt(stack.Count - 1);
            states[node] = 2;
        }
    }

    private static void ValidateId(
        string value,
        string kind,
        string path,
        ImmutableArray<StatementContractViolation>.Builder violations)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            violations.Add(new($"{kind.ToUpperInvariant()}_ID_EMPTY", path, $"{kind} id must not be empty."));
            return;
        }

        if (!string.Equals(value, value.ToLowerInvariant(), StringComparison.Ordinal) ||
            value.Any(c => !(char.IsAsciiLetterOrDigit(c) || c is ':' or '-' or '_' or '.')))
        {
            violations.Add(new(
                $"{kind.ToUpperInvariant()}_ID_NOT_CANONICAL",
                path,
                $"{kind} id '{value}' must contain only lowercase ASCII letters, digits, colon, hyphen, underscore, or dot."));
        }
    }

    private static void ValidateUniqueIds(
        IEnumerable<string> values,
        string kind,
        string path,
        ImmutableArray<StatementContractViolation>.Builder violations)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var value in values)
        {
            if (!seen.Add(value))
                violations.Add(new($"{kind.ToUpperInvariant()}_ID_DUPLICATE", path, $"Duplicate {kind} id '{value}'."));
        }
    }

    private static void ValidateUniqueOrders(
        IEnumerable<int> values,
        string code,
        string path,
        ImmutableArray<StatementContractViolation>.Builder violations)
    {
        var duplicate = values.GroupBy(v => v).FirstOrDefault(g => g.Count() > 1);
        if (duplicate is not null)
            violations.Add(new(code, path, $"Presentation order {duplicate.Key} occurs more than once."));
    }

    private static void ValidateCount(
        int count,
        int limit,
        string code,
        string path,
        ImmutableArray<StatementContractViolation>.Builder violations)
    {
        if (count > limit)
            violations.Add(new(code, path, $"Count {count} exceeds {limit}."));
    }

    private static string Format(StatementCellAddress address) =>
        $"{address.SheetId.Value}/{address.RowId.Value}/{address.ColumnId.Value}";
}
