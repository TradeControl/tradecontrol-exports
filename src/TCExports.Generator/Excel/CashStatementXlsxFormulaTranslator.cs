using System.Globalization;
using System.Text;
using TCExports.Generator.CashStatement.Model;

namespace TCExports.Generator.Excel;

/// <summary>
/// Translates validated semantic formulas to Excel A1 syntax through a completed layout.
/// </summary>
public sealed class CashStatementXlsxFormulaTranslator : ICashStatementFormulaTranslator
{
    public string Translate(StatementFormula formula, StatementLayout layout)
    {
        ArgumentNullException.ThrowIfNull(formula);
        ArgumentNullException.ThrowIfNull(layout);
        return $"={TranslateCore(formula, layout)}";
    }

    private static string TranslateCore(StatementFormula formula, StatementLayout layout) => formula switch
    {
        StatementFormulaReference reference => Reference(reference.Address, layout),
        StatementFormulaLiteral literal => Literal(literal.Value),
        StatementFormulaUnary unary => Unary(unary, layout),
        StatementFormulaBinary binary => Binary(binary, layout),
        StatementFormulaAggregate aggregate => Aggregate(aggregate, layout),
        StatementFormulaConditional conditional =>
            $"IF({TranslateCore(conditional.Condition, layout)},{TranslateCore(conditional.WhenTrue, layout)},{TranslateCore(conditional.WhenFalse, layout)})",
        StatementFormulaTerminal terminal => Terminal(terminal, layout),
        _ => throw new InvalidOperationException($"Unsupported formula node '{formula.GetType().Name}'.")
    };

    private static string Unary(StatementFormulaUnary formula, StatementLayout layout)
    {
        var operand = TranslateCore(formula.Operand, layout);
        return formula.Operator switch
        {
            StatementUnaryOperator.Negate when formula.Operand is StatementFormulaAggregate => $"{operand}*-1",
            StatementUnaryOperator.Negate => $"-({operand})",
            StatementUnaryOperator.Absolute => $"ABS({operand})",
            _ => throw new InvalidOperationException($"Unsupported unary operator '{formula.Operator}'.")
        };
    }

    private static string Binary(StatementFormulaBinary formula, StatementLayout layout)
    {
        var symbol = formula.Operator switch
        {
            StatementBinaryOperator.Add => "+",
            StatementBinaryOperator.Subtract => "-",
            StatementBinaryOperator.Multiply => "*",
            StatementBinaryOperator.Divide => "/",
            StatementBinaryOperator.Equal => "=",
            StatementBinaryOperator.NotEqual => "<>",
            StatementBinaryOperator.LessThan => "<",
            StatementBinaryOperator.LessThanOrEqual => "<=",
            StatementBinaryOperator.GreaterThan => ">",
            StatementBinaryOperator.GreaterThanOrEqual => ">=",
            _ => throw new InvalidOperationException($"Unsupported binary operator '{formula.Operator}'.")
        };
        return $"{TranslateCore(formula.Left, layout)}{symbol}{TranslateCore(formula.Right, layout)}";
    }

    private static string Aggregate(StatementFormulaAggregate formula, StatementLayout layout)
    {
        if (formula.Operands.IsEmpty)
            throw new InvalidOperationException("An empty aggregate has an authoritative value and must not be emitted as an XLSX formula.");

        var operands = formula.Operands.Select(operand => TranslateCore(operand, layout)).ToArray();
        return formula.Operator switch
        {
            StatementAggregateOperator.Sum => Sum(operands, formula.Operands, layout),
            StatementAggregateOperator.Count => $"COUNT({string.Join(',', operands)})",
            _ => throw new InvalidOperationException($"Unsupported aggregate operator '{formula.Operator}'.")
        };
    }

    private static string Sum(
        IReadOnlyList<string> operands,
        IReadOnlyList<StatementFormula> formulas,
        StatementLayout layout)
    {
        if (formulas.All(formula => formula is StatementFormulaReference))
        {
            var physical = formulas
                .Cast<StatementFormulaReference>()
                .Select(reference => Resolve(reference.Address, layout))
                .ToArray();
            if (IsContiguous(physical))
                return $"SUM({A1(physical[0])}:{A1(physical[^1])})";
        }

        return string.Join('+', operands);
    }

    private static bool IsContiguous(IReadOnlyList<PhysicalCellAddress> cells)
    {
        if (cells.Count == 0 || cells.Select(cell => cell.SheetName).Distinct(StringComparer.Ordinal).Count() != 1)
            return false;
        if (cells.All(cell => cell.RowNumber == cells[0].RowNumber))
            return cells.Select(cell => cell.ColumnNumber).SequenceEqual(
                Enumerable.Range(cells[0].ColumnNumber, cells.Count));
        if (cells.All(cell => cell.ColumnNumber == cells[0].ColumnNumber))
            return cells.Select(cell => cell.RowNumber).SequenceEqual(
                Enumerable.Range(cells[0].RowNumber, cells.Count));
        return false;
    }

    private static string Terminal(StatementFormulaTerminal formula, StatementLayout layout)
    {
        var translated = formula.Expression;
        foreach (var reference in formula.References.OrderByDescending(pair => pair.Key.Length))
        {
            translated = translated.Replace(
                $"[{reference.Key}]",
                Reference(reference.Value, layout),
                StringComparison.OrdinalIgnoreCase);
        }
        return translated.TrimStart().TrimStart('=');
    }

    private static string Reference(StatementCellAddress address, StatementLayout layout) => A1(Resolve(address, layout));

    private static PhysicalCellAddress Resolve(StatementCellAddress address, StatementLayout layout) =>
        layout.Cells.TryGetValue(address, out var physical)
            ? physical
            : throw new InvalidOperationException($"Semantic formula reference '{address}' is absent from the XLSX layout.");

    private static string A1(PhysicalCellAddress address) => $"{ColumnName(address.ColumnNumber)}{address.RowNumber}";

    private static string ColumnName(int columnNumber)
    {
        if (columnNumber < 1)
            throw new ArgumentOutOfRangeException(nameof(columnNumber));
        var result = new StringBuilder();
        for (var value = columnNumber; value > 0; value = (value - 1) / 26)
            result.Insert(0, (char)('A' + ((value - 1) % 26)));
        return result.ToString();
    }

    private static string Literal(StatementValue value)
    {
        if (value.State != StatementValueState.Present)
            throw new InvalidOperationException($"Cannot emit a {value.State} formula literal.");
        return value.Kind switch
        {
            StatementScalarKind.Decimal when value.DecimalValue is { } number =>
                number.ToString(CultureInfo.InvariantCulture),
            StatementScalarKind.Text => $"\"{(value.TextValue ?? string.Empty).Replace("\"", "\"\"")}\"",
            StatementScalarKind.Date when value.DateValue is { } date =>
                $"DATE({date.Year},{date.Month},{date.Day})",
            StatementScalarKind.Boolean when value.BooleanValue is { } boolean => boolean ? "TRUE" : "FALSE",
            _ => throw new InvalidOperationException($"Incomplete {value.Kind} formula literal.")
        };
    }
}
