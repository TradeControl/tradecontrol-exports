using TCExports.Generator.CashStatement.Model;

namespace TCExports.Generator.CashStatement;

/// <summary>
/// Evaluates semantic formulas from authoritative values. It never reads or parses native spreadsheet formulas.
/// </summary>
public sealed class CashStatementFormulaEvaluator
{
    public StatementValue Evaluate(
        StatementFormula formula,
        Func<StatementCellAddress, StatementValue> resolveReference) =>
        formula switch
        {
            StatementFormulaLiteral literal => literal.Value,
            StatementFormulaReference reference => resolveReference(reference.Address),
            StatementFormulaUnary unary => EvaluateUnary(unary, resolveReference),
            StatementFormulaBinary binary => EvaluateBinary(binary, resolveReference),
            StatementFormulaAggregate aggregate => EvaluateAggregate(aggregate, resolveReference),
            StatementFormulaConditional conditional => EvaluateConditional(conditional, resolveReference),
            StatementFormulaTerminal => StatementValue.Unsupported(
                StatementScalarKind.Decimal,
                "spreadsheet-engine-evaluation-required"),
            _ => throw new InvalidOperationException($"Unsupported formula node '{formula.GetType().Name}'.")
        };

    private StatementValue EvaluateUnary(
        StatementFormulaUnary formula,
        Func<StatementCellAddress, StatementValue> resolveReference)
    {
        var operand = Evaluate(formula.Operand, resolveReference);
        if (operand.State != StatementValueState.Present)
            return operand;
        if (operand.Kind != StatementScalarKind.Decimal || operand.DecimalValue is null)
            return StatementValue.Unsupported(StatementScalarKind.Decimal, "unary-operand-not-decimal");

        return formula.Operator switch
        {
            StatementUnaryOperator.Negate => StatementValue.FromDecimal(-operand.DecimalValue.Value),
            StatementUnaryOperator.Absolute => StatementValue.FromDecimal(Math.Abs(operand.DecimalValue.Value)),
            _ => throw new InvalidOperationException($"Unsupported unary operator '{formula.Operator}'.")
        };
    }

    private StatementValue EvaluateBinary(
        StatementFormulaBinary formula,
        Func<StatementCellAddress, StatementValue> resolveReference)
    {
        var left = Evaluate(formula.Left, resolveReference);
        if (left.State != StatementValueState.Present)
            return left;

        var right = Evaluate(formula.Right, resolveReference);
        if (right.State != StatementValueState.Present)
            return right;

        if (formula.Operator is StatementBinaryOperator.Equal or StatementBinaryOperator.NotEqual)
        {
            var equal = ValuesEqual(left, right);
            return StatementValue.FromBoolean(
                formula.Operator == StatementBinaryOperator.Equal ? equal : !equal);
        }

        if (left.Kind != StatementScalarKind.Decimal || right.Kind != StatementScalarKind.Decimal ||
            left.DecimalValue is null || right.DecimalValue is null)
        {
            return StatementValue.Unsupported(StatementScalarKind.Decimal, "binary-operands-not-decimal");
        }

        var leftValue = left.DecimalValue.Value;
        var rightValue = right.DecimalValue.Value;

        return formula.Operator switch
        {
            StatementBinaryOperator.Add => StatementValue.FromDecimal(leftValue + rightValue),
            StatementBinaryOperator.Subtract => StatementValue.FromDecimal(leftValue - rightValue),
            StatementBinaryOperator.Multiply => StatementValue.FromDecimal(leftValue * rightValue),
            StatementBinaryOperator.Divide when rightValue == 0m =>
                StatementValue.Unsupported(StatementScalarKind.Decimal, "division-by-zero"),
            StatementBinaryOperator.Divide => StatementValue.FromDecimal(leftValue / rightValue),
            StatementBinaryOperator.LessThan => StatementValue.FromBoolean(leftValue < rightValue),
            StatementBinaryOperator.LessThanOrEqual => StatementValue.FromBoolean(leftValue <= rightValue),
            StatementBinaryOperator.GreaterThan => StatementValue.FromBoolean(leftValue > rightValue),
            StatementBinaryOperator.GreaterThanOrEqual => StatementValue.FromBoolean(leftValue >= rightValue),
            _ => throw new InvalidOperationException($"Unsupported binary operator '{formula.Operator}'.")
        };
    }

    private StatementValue EvaluateAggregate(
        StatementFormulaAggregate formula,
        Func<StatementCellAddress, StatementValue> resolveReference)
    {
        decimal sum = 0m;
        decimal count = 0m;

        foreach (var operandFormula in formula.Operands)
        {
            var operand = Evaluate(operandFormula, resolveReference);
            if (operand.State != StatementValueState.Present)
                return operand;

            count++;
            if (formula.Operator == StatementAggregateOperator.Sum)
            {
                if (operand.Kind != StatementScalarKind.Decimal || operand.DecimalValue is null)
                    return StatementValue.Unsupported(StatementScalarKind.Decimal, "sum-operand-not-decimal");
                sum += operand.DecimalValue.Value;
            }
        }

        return formula.Operator switch
        {
            StatementAggregateOperator.Sum => StatementValue.FromDecimal(sum),
            StatementAggregateOperator.Count => StatementValue.FromDecimal(count),
            _ => throw new InvalidOperationException($"Unsupported aggregate operator '{formula.Operator}'.")
        };
    }

    private StatementValue EvaluateConditional(
        StatementFormulaConditional formula,
        Func<StatementCellAddress, StatementValue> resolveReference)
    {
        var condition = Evaluate(formula.Condition, resolveReference);
        if (condition.State != StatementValueState.Present)
            return condition;
        if (condition.Kind != StatementScalarKind.Boolean || condition.BooleanValue is null)
            return StatementValue.Unsupported(StatementScalarKind.Boolean, "condition-not-boolean");

        // Deliberately lazy so an unselected division-by-zero branch cannot poison the result.
        return Evaluate(condition.BooleanValue.Value ? formula.WhenTrue : formula.WhenFalse, resolveReference);
    }

    private static bool ValuesEqual(StatementValue left, StatementValue right)
    {
        if (left.Kind != right.Kind)
            return false;

        return left.Kind switch
        {
            StatementScalarKind.Decimal => left.DecimalValue == right.DecimalValue,
            StatementScalarKind.Text => string.Equals(left.TextValue, right.TextValue, StringComparison.Ordinal),
            StatementScalarKind.Date => left.DateValue == right.DateValue,
            StatementScalarKind.Boolean => left.BooleanValue == right.BooleanValue,
            _ => false
        };
    }
}
