using System.Globalization;
using TCExports.Generator.CashStatement.Model;

namespace TCExports.Generator.Ods;

public sealed class CashStatementOdsFormulaTranslator(
    IReadOnlyDictionary<StatementCellAddress, OdsCellReference> addresses)
{
    private readonly IReadOnlyDictionary<StatementCellAddress, OdsCellReference> _addresses = addresses;

    public string Translate(StatementFormula formula) => $"of:={TranslateNode(formula)}";

    private string TranslateNode(StatementFormula formula) => formula switch
    {
        StatementFormulaReference reference => Reference(reference.Address),
        StatementFormulaLiteral literal => Literal(literal.Value),
        StatementFormulaUnary unary => Unary(unary),
        StatementFormulaBinary binary => Binary(binary),
        StatementFormulaAggregate aggregate => Aggregate(aggregate),
        StatementFormulaConditional conditional =>
            $"IF({TranslateNode(conditional.Condition)};{TranslateNode(conditional.WhenTrue)};{TranslateNode(conditional.WhenFalse)})",
        StatementFormulaTerminal terminal => TranslateTerminal(terminal),
        _ => throw new NotSupportedException($"Unsupported statement formula type: {formula.GetType().Name}.")
    };

    private string Reference(StatementCellAddress address) =>
        _addresses.TryGetValue(address, out var reference)
            ? reference.ToOpenFormulaReference()
            : throw new InvalidDataException($"No ODS cell was allocated for semantic address '{address}'.");

    private static string Literal(StatementValue value) => value switch
    {
        { State: StatementValueState.Present, Kind: StatementScalarKind.Decimal, DecimalValue: not null } =>
            value.DecimalValue.Value.ToString(CultureInfo.InvariantCulture),
        { State: StatementValueState.Present, Kind: StatementScalarKind.Text } =>
            $"\"{(value.TextValue ?? string.Empty).Replace("\"", "\"\"")}\"",
        { State: StatementValueState.Present, Kind: StatementScalarKind.Date, DateValue: not null } =>
            $"DATE({value.DateValue.Value.Year};{value.DateValue.Value.Month};{value.DateValue.Value.Day})",
        { State: StatementValueState.Present, Kind: StatementScalarKind.Boolean, BooleanValue: not null } =>
            value.BooleanValue.Value ? "TRUE()" : "FALSE()",
        _ => throw new NotSupportedException($"A {value.State} {value.Kind} value is not a valid ODS formula literal.")
    };

    private string Unary(StatementFormulaUnary unary) => unary.Operator switch
    {
        StatementUnaryOperator.Negate => $"-({TranslateNode(unary.Operand)})",
        StatementUnaryOperator.Absolute => $"ABS({TranslateNode(unary.Operand)})",
        _ => throw new NotSupportedException($"Unsupported unary operator: {unary.Operator}.")
    };

    private string Binary(StatementFormulaBinary binary) =>
        $"({TranslateNode(binary.Left)}{Operator(binary.Operator)}{TranslateNode(binary.Right)})";

    private static string Operator(StatementBinaryOperator operation) => operation switch
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
        _ => throw new NotSupportedException($"Unsupported binary operator: {operation}.")
    };

    private string Aggregate(StatementFormulaAggregate aggregate)
    {
        var function = aggregate.Operator switch
        {
            StatementAggregateOperator.Sum => "SUM",
            StatementAggregateOperator.Count => "COUNT",
            _ => throw new NotSupportedException($"Unsupported aggregate operator: {aggregate.Operator}.")
        };

        if (aggregate.Operands.IsDefaultOrEmpty)
            return aggregate.Operator == StatementAggregateOperator.Sum ? "0" : $"{function}()";

        return $"{function}({string.Join(';', aggregate.Operands.Select(TranslateNode))})";
    }

    private string TranslateTerminal(StatementFormulaTerminal terminal)
    {
        var expression = terminal.Expression;
        foreach (var reference in terminal.References.OrderByDescending(item => item.Key.Length))
            expression = expression.Replace(
                $"[{reference.Key}]",
                Reference(reference.Value),
                StringComparison.Ordinal);

        return expression.Replace(',', ';');
    }
}

public readonly record struct OdsCellReference(string WorksheetName, int RowNumber, int ColumnNumber)
{
    public string ToOpenFormulaReference()
    {
        if (RowNumber < 1 || ColumnNumber < 1)
            throw new InvalidDataException("ODS cell references are one-based.");

        var column = ColumnName(ColumnNumber);
        return $"[.{column}{RowNumber}]";
    }

    private static string ColumnName(int number)
    {
        var result = string.Empty;
        while (number > 0)
        {
            number--;
            result = (char)('A' + number % 26) + result;
            number /= 26;
        }

        return result;
    }
}
