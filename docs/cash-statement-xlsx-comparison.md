# Cash Statement XLSX comparison contract

The Cash Statement migration uses a read-only workbook characterizer and comparator before any production Excel implementation is changed.

## Comparison profiles

- **Strict equivalence** requires identical inspected cells, formulas, effective styles, worksheet structure, defined names, normalized package parts, and non-volatile metadata.
- **Semantic definition equivalence** compares stable section, business-row, and fiscal-period identities. It includes input values, normalized formulas, and effective financial-cell styles but excludes formula cached results.
- **Semantic result equivalence** adds recalculated formula and displayed results. Use this profile when comparing recalculated workbooks.

Strict equivalence is the default old-versus-new raw XLSX gate. Semantic result equivalence is the financial gate after recalculation.

## Permitted normalizations

- ZIP entry timestamps are not workbook content.
- Core-property created time, modified time, and last-modified-by are volatile.
- The randomized OPC core-properties part name and its matching relationship/content-type references are volatile.
- `Cash Flow!B3` is the visible generation timestamp and may be normalized only for separately generated workbook comparisons.
- `SUM(D7:D7)` and `SUM(D7)` are treated as the same formula definition because LibreOffice 26.2.6.3 consistently performs that rewrite.
- Effective style comparison resolves number format, font, fill, borders, alignment, and protection. It does not compare raw style indices or equivalent XML spellings.

No other value, formula, style, worksheet property, defined name, relationship, calculation setting, package part, or metadata field is excluded without a reviewed change to this contract.

## Unknown properties

Every package part receives a normalized content hash. A change to a property without a higher-level inspection rule therefore produces a blocking package difference. New or removed parts also fail.

## Mutation proof

The characterizer's `self-test` command creates disposable package copies and proves detection of value, formula, style, structure, and non-volatile metadata changes. It also proves that semantic row comparison is independent of collection order.
