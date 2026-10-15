using System.Collections.Immutable;
using System.Text;

namespace TCExports.Generator.CashStatement;

public enum CashStatementReferenceKind
{
    Category,
    CashCode
}

public sealed record CashStatementCategoryReference(
    string CategoryCode,
    string Category,
    CashStatementReferenceKind Kind = CashStatementReferenceKind.Category);

public sealed record CashStatementExpressionDiagnostic(
    string Code,
    string Token,
    string Message);

public sealed record CashStatementExpressionTranslation(
    string Expression,
    ImmutableArray<string> CategoryCodes,
    ImmutableArray<CashStatementExpressionDiagnostic> Diagnostics)
{
    public bool IsValid => Diagnostics.IsDefaultOrEmpty;
}

/// <summary>
/// Translates only the Cash Statement's reserved [category] references. The
/// surrounding spreadsheet expression is deliberately treated as opaque text.
/// </summary>
public static class CashStatementExpressionTranslator
{
    public const int MaximumExpressionLength = 4000;

    public static CashStatementExpressionTranslation ForStorage(
        string expression,
        IEnumerable<CashStatementCategoryReference> categories) =>
        Translate(expression, categories, useDescriptions: false);

    public static CashStatementExpressionTranslation ForDisplay(
        string expression,
        IEnumerable<CashStatementCategoryReference> categories) =>
        Translate(expression, categories, useDescriptions: true);

    private static CashStatementExpressionTranslation Translate(
        string expression,
        IEnumerable<CashStatementCategoryReference> categories,
        bool useDescriptions)
    {
        ArgumentNullException.ThrowIfNull(expression);
        ArgumentNullException.ThrowIfNull(categories);

        var catalogue = categories
            .Where(category =>
                !string.IsNullOrWhiteSpace(category.CategoryCode) &&
                !string.IsNullOrWhiteSpace(category.Category))
            .ToImmutableArray();

        var categoriesByCode = Index(catalogue, CashStatementReferenceKind.Category, reference => reference.CategoryCode);
        var cashCodesByCode = Index(catalogue, CashStatementReferenceKind.CashCode, reference => reference.CategoryCode);
        var categoriesByDescription = Index(catalogue, CashStatementReferenceKind.Category, reference => reference.Category);
        var cashCodesByDescription = Index(catalogue, CashStatementReferenceKind.CashCode, reference => reference.Category);

        var output = new StringBuilder(expression.Length);
        var references = ImmutableArray.CreateBuilder<string>();
        var diagnostics = ImmutableArray.CreateBuilder<CashStatementExpressionDiagnostic>();
        var seenCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        var cursor = 0;
        while (cursor < expression.Length)
        {
            var open = expression.IndexOf('[', cursor);
            if (open < 0)
            {
                output.Append(expression, cursor, expression.Length - cursor);
                break;
            }

            output.Append(expression, cursor, open - cursor);
            var close = expression.IndexOf(']', open + 1);
            if (close < 0)
            {
                var remainder = expression[open..];
                output.Append(remainder);
                diagnostics.Add(new(
                    "EXPRESSION_TAG_UNCLOSED",
                    remainder,
                    $"Category reference beginning '{remainder}' has no closing bracket."));
                break;
            }

            var original = expression.Substring(open, close - open + 1);
            var token = expression.Substring(open + 1, close - open - 1).Trim();
            if (token.Length == 0)
            {
                output.Append(original);
                diagnostics.Add(new(
                    "EXPRESSION_TAG_EMPTY",
                    token,
                    "An empty category reference was found."));
                cursor = close + 1;
                continue;
            }

            CashStatementCategoryReference? match = null;
            if (TryUnique(categoriesByCode, token, out match) ||
                TryUnique(cashCodesByCode, token, out match))
            {
            }
            else if (categoriesByDescription.TryGetValue(token, out var categoryDescriptionMatches))
            {
                if (categoryDescriptionMatches.Length == 1)
                {
                    match = categoryDescriptionMatches[0];
                }
                else
                {
                    diagnostics.Add(new(
                        "EXPRESSION_TAG_AMBIGUOUS",
                        token,
                        $"Category reference [{token}] matches {categoryDescriptionMatches.Length} category descriptions."));
                }
            }
            else if (cashCodesByDescription.TryGetValue(token, out var cashDescriptionMatches))
            {
                if (cashDescriptionMatches.Length == 1)
                {
                    match = cashDescriptionMatches[0];
                }
                else
                {
                    diagnostics.Add(new(
                        "EXPRESSION_TAG_AMBIGUOUS",
                        token,
                        $"Category reference [{token}] matches {cashDescriptionMatches.Length} cash-code descriptions."));
                }
            }
            else
            {
                diagnostics.Add(new(
                    "EXPRESSION_TAG_NOT_FOUND",
                    token,
                    $"Category reference [{token}] does not match a category code or description."));
            }

            if (match is null)
            {
                output.Append(original);
            }
            else
            {
                var code = match.CategoryCode.Trim();
                output.Append('[');
                output.Append(useDescriptions ? match.Category.Trim() : code);
                output.Append(']');
                if (seenCodes.Add(code))
                    references.Add(code);
            }

            cursor = close + 1;
        }

        return new(output.ToString(), references.ToImmutable(), diagnostics.ToImmutable());
    }

    private static Dictionary<string, ImmutableArray<CashStatementCategoryReference>> Index(
        ImmutableArray<CashStatementCategoryReference> catalogue,
        CashStatementReferenceKind kind,
        Func<CashStatementCategoryReference, string> keySelector) =>
        catalogue
            .Where(reference => reference.Kind == kind)
            .GroupBy(reference => keySelector(reference).Trim(), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.ToImmutableArray(), StringComparer.OrdinalIgnoreCase);

    private static bool TryUnique(
        IReadOnlyDictionary<string, ImmutableArray<CashStatementCategoryReference>> index,
        string token,
        out CashStatementCategoryReference? match)
    {
        if (index.TryGetValue(token, out var matches) && matches.Length == 1)
        {
            match = matches[0];
            return true;
        }

        match = null;
        return false;
    }
}
