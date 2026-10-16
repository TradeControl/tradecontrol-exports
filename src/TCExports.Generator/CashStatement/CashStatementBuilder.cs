using System.Collections.Immutable;
using System.Text;
using TCExports.Generator.CashStatement.Model;
using TCExports.Generator.Contracts;
using TCExports.Generator.Data;

namespace TCExports.Generator.CashStatement;

/// <summary>
/// Builds the renderer-neutral Cash Statement beside the existing production handler.
/// Phase 3 is adding repository-backed sections incrementally while the production
/// spreadsheet handler remains unchanged.
/// </summary>
public sealed class CashStatementBuilder
{
    private static readonly StatementSheetId CashFlowSheetId = new("cash-flow");
    private static readonly ImmutableArray<VatMetricDefinition> VatMetrics = ImmutableArray.Create(
        new VatMetricDefinition("vat-due-sales", CashStatementText.VatDueSales, 0),
        new VatMetricDefinition("vat-due-acquisitions", CashStatementText.VatDueAcquisitions, 1),
        new VatMetricDefinition("total-vat-due", CashStatementText.TotalVatDue, 2),
        new VatMetricDefinition("vat-reclaimed-current-period", CashStatementText.VatReclaimedCurrentPeriod, 3),
        new VatMetricDefinition("total-value-sales-ex-vat", CashStatementText.TotalValueSalesExVat, 4),
        new VatMetricDefinition("total-value-purchases-ex-vat", CashStatementText.TotalValuePurchasesExVat, 5),
        new VatMetricDefinition("total-value-goods-supplied-ex-vat", CashStatementText.TotalValueGoodsSuppliedExVat, 6),
        new VatMetricDefinition("total-value-goods-received-ex-vat", CashStatementText.TotalValueGoodsReceivedExVat, 7),
        new VatMetricDefinition("net-vat-due", CashStatementText.NetVatDue, 8));

    private readonly ICashFlowRepository _repository;
    private readonly CashStatementLimits _limits;
    private readonly ICashStatementLocalizer _text;
    private readonly CashStatementContractValidator _validator = new();
    private readonly CashStatementDocumentEvaluator _documentEvaluator = new();

    public CashStatementBuilder(
        ICashFlowRepository repository,
        CashStatementLimits? limits = null,
        ICashStatementLocalizer? localizer = null)
    {
        _repository = repository;
        _limits = limits ?? CashStatementLimits.Default;
        _text = localizer ?? new ResourceCashStatementLocalizer();
    }

    public async Task<CashStatementContractResult<CashStatementDocument>> BuildAsync(
        string connectionString,
        CashStatementRequestV1 request,
        string requestedBy,
        DateTimeOffset generatedAt,
        CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(connectionString);
        ArgumentNullException.ThrowIfNull(request);
        ct.ThrowIfCancellationRequested();

        var timeout = request.CommandTimeoutSeconds;
        var active = await _repository.GetActivePeriodAsync(connectionString, timeout, ct);
        var years = await _repository.GetActiveYearsAsync(connectionString, timeout, ct);
        var months = await _repository.GetMonthsAsync(connectionString, timeout, ct);
        var companyName = await _repository.GetCompanyNameAsync(connectionString, timeout, ct);

        var periodResult = CashStatementPeriodGridBuilder.Build(years, months, active, _limits);
        if (!periodResult.IsSuccess)
            return new(null, periodResult.Violations);

        var sourceResult = await LoadFlowSourceAsync(connectionString, request, years, ct);
        if (!sourceResult.IsSuccess)
            return new(null, sourceResult.Violations);

        var periodGrid = periodResult.Value!;
        var buildViolations = ImmutableArray.CreateBuilder<StatementContractViolation>();
        var sections = BuildFlowSections(periodGrid, sourceResult.Value!, _text, buildViolations).ToBuilder();

        var expressionSection = await BuildExpressionSectionAsync(
            connectionString,
            periodGrid,
            sourceResult.Value!,
            sections,
            sections.Count,
            timeout,
            ct);
        if (expressionSection is not null)
            sections.Add(expressionSection);

        if (request.IncludeBankBalances)
        {
            var bankSection = await BuildBankBalancesSectionAsync(
                connectionString,
                periodGrid,
                sections.Count,
                timeout,
                buildViolations,
                ct);
            sections.Add(bankSection);
        }

        if (request.IncludeVatDetails)
        {
            var vatSections = await BuildVatSectionsAsync(
                connectionString,
                periodGrid,
                sections.Count,
                request,
                generatedAt.UtcDateTime,
                buildViolations,
                ct);
            sections.AddRange(vatSections);
        }

        if (request.IncludeBalanceSheet)
        {
            var balanceSheetSection = await BuildBalanceSheetSectionAsync(
                connectionString,
                periodGrid,
                sections.Count,
                timeout,
                buildViolations,
                ct);
            if (balanceSheetSection is not null)
                sections.Add(balanceSheetSection);
        }

        var sheets = ImmutableArray.CreateBuilder<StatementSheet>();
        sheets.Add(new StatementSheet(
            CashFlowSheetId,
            StatementSheetKind.CashFlow,
            0,
            _text.Get(CashStatementText.CashFlow),
            BuildColumns(periodGrid, _text),
            sections.ToImmutable()));
        if (request.IncludeReconciliation)
        {
            sheets.Add(await BuildReconciliationSheetAsync(
                connectionString,
                sheets.Count,
                timeout,
                ct));
        }

        var document = new CashStatementDocument(
            new CashStatementMetadata(
                companyName,
                _text.Format(CashStatementText.StatementTitle, active.MonthName, active.Description),
                requestedBy,
                generatedAt,
                StatementPeriodId.Month(active.YearNumber, checked((byte)active.MonthNumber)),
                _text.Get(CashStatementText.Date)),
            request,
            periodGrid,
            sheets.ToImmutable());

        buildViolations.AddRange(_validator.Validate(document, _limits));
        if (buildViolations.Count > 0)
            return new(null, buildViolations.ToImmutable());

        var evaluated = _documentEvaluator.Evaluate(document);
        var evaluatedViolations = _validator.Validate(evaluated, _limits);
        return evaluatedViolations.IsEmpty
            ? new(evaluated, ImmutableArray<StatementContractViolation>.Empty)
            : new(null, evaluatedViolations);
    }

    private async Task<CashStatementContractResult<FlowSource>> LoadFlowSourceAsync(
        string connectionString,
        CashStatementRequestV1 request,
        IReadOnlyList<ActiveYearDto> years,
        CancellationToken ct)
    {
        var violations = ImmutableArray.CreateBuilder<StatementContractViolation>();
        var categories = ImmutableArray.CreateBuilder<CategorySource>();
        var timeout = request.CommandTimeoutSeconds;

        foreach (var cashType in new[] { CashType.Trade, CashType.Money, CashType.Tax })
        {
            ct.ThrowIfCancellationRequested();
            var categoryDtos = await _repository.GetCategoriesAsync(connectionString, cashType, timeout, ct);
            foreach (var category in categoryDtos)
            {
                ct.ThrowIfCancellationRequested();
                var cashCodeDtos = await _repository.GetCashCodesAsync(
                    connectionString,
                    category.CategoryCode,
                    timeout,
                    ct);
                var cashCodes = ImmutableArray.CreateBuilder<CashCodeSource>();

                foreach (var cashCode in cashCodeDtos)
                {
                    var values = new Dictionary<StatementPeriodId, decimal>();
                    foreach (var year in years)
                    {
                        ct.ThrowIfCancellationRequested();
                        var yearValues = await _repository.GetCashCodeValuesAsync(
                            connectionString,
                            cashCode.CashCode,
                            year.YearNumber,
                            IncludeActivePeriods(cashType, request),
                            IncludeOrderBook(cashType, request),
                            IncludeTaxAccruals(cashType, request),
                            timeout,
                            ct);

                        foreach (var value in yearValues)
                        {
                            var periodId = StatementPeriodId.Month(
                                year.YearNumber,
                                checked((byte)value.StartOn.Month));
                            if (!values.TryAdd(periodId, value.InvoiceValue))
                            {
                                violations.Add(new(
                                    "FLOW_VALUE_DUPLICATE",
                                    $"flow.{Canonical(category.CategoryCode)}.{Canonical(cashCode.CashCode)}.{periodId}",
                                    $"Cash code '{cashCode.CashCode}' returned more than one value for period '{periodId}'."));
                            }
                        }
                    }

                    cashCodes.Add(new CashCodeSource(
                        cashCode.CashCode,
                        cashCode.CashDescription,
                        values.ToImmutableDictionary()));
                }

                categories.Add(new CategorySource(
                    cashType,
                    category.CategoryCode,
                    category.Category,
                    category.CashPolarityCode,
                    category.DisplayOrder,
                    cashCodes.ToImmutable()));
            }
        }

        var calculatedTotals = ImmutableArray.CreateBuilder<CalculatedTotalSource>();
        foreach (var cashType in new[] { CashType.Trade, CashType.Tax })
        {
            ct.ThrowIfCancellationRequested();
            var totalDtos = await _repository.GetCategoriesByTypeAsync(
                connectionString,
                cashType,
                CategoryType.Total,
                timeout,
                ct);
            calculatedTotals.AddRange(totalDtos.Select(total => new CalculatedTotalSource(
                cashType,
                total.CategoryCode,
                total.Category,
                total.CashType)));
        }

        var totalChildren = ImmutableDictionary.CreateBuilder<string, ImmutableArray<string>>(
            StringComparer.OrdinalIgnoreCase);
        var totals = await _repository.GetCategoryTotalsAsync(connectionString, timeout, ct);
        foreach (var total in totals)
        {
            ct.ThrowIfCancellationRequested();
            var children = await _repository.GetCategoryTotalCodesAsync(
                connectionString,
                total.CategoryCode,
                timeout,
                ct);
            if (!totalChildren.TryAdd(total.CategoryCode, children.ToImmutableArray()))
            {
                violations.Add(new(
                    "CATEGORY_TOTAL_DUPLICATE",
                    $"flow.{Canonical(total.CategoryCode)}",
                    $"Calculated category '{total.CategoryCode}' occurs more than once."));
            }
        }

        return violations.Count == 0
            ? new(new FlowSource(
                    categories.ToImmutable(),
                    calculatedTotals.ToImmutable(),
                    totalChildren.ToImmutable()),
                ImmutableArray<StatementContractViolation>.Empty)
            : new(null, violations.ToImmutable());
    }

    private static ImmutableArray<StatementColumn> BuildColumns(
        StatementPeriodGrid grid,
        ICashStatementLocalizer text)
    {
        var columns = ImmutableArray.CreateBuilder<StatementColumn>(grid.Periods.Length + 3);
        columns.Add(new(new StatementColumnId("code"), text.Get(CashStatementText.Code), StatementColumnKind.Identifier, 0));
        columns.Add(new(new StatementColumnId("label"), text.Get(CashStatementText.Name), StatementColumnKind.Label, 1));
        columns.Add(new(new StatementColumnId("marker"), string.Empty, StatementColumnKind.HiddenMarker, 2,
            StatementPresentationRole.HiddenMarker));

        foreach (var period in grid.Periods)
        {
            columns.Add(new StatementColumn(
                PeriodColumn(period.Id),
                period.Id.IsAnnual ? text.Get(CashStatementText.Totals) : period.MonthLabel,
                period.Id.IsAnnual ? StatementColumnKind.AnnualPeriod : StatementColumnKind.Period,
                period.PresentationOrder + 3,
                period.IsActive
                    ? StatementPresentationRole.CurrentPeriod
                    : period.Id.IsAnnual
                        ? StatementPresentationRole.AnnualPeriod
                        : StatementPresentationRole.Default,
                period.Id));
        }

        return columns.ToImmutable();
    }

    private static ImmutableArray<StatementSection> BuildFlowSections(
        StatementPeriodGrid grid,
        FlowSource source,
        ICashStatementLocalizer text,
        ImmutableArray<StatementContractViolation>.Builder violations)
    {
        var rowByCategory = new Dictionary<string, RowDraft>(StringComparer.OrdinalIgnoreCase);
        var flowRows = new Dictionary<CashType, List<RowDraft>>();
        var categoryOrder = new Dictionary<CashType, List<CategorySource>>();

        foreach (var cashType in new[] { CashType.Trade, CashType.Money, CashType.Tax })
        {
            var rows = new List<RowDraft>();
            var order = 0;
            var orderedCategories = source.Categories
                .Where(item => item.CashType == cashType)
                .OrderBy(item => item.DisplayOrder ?? short.MaxValue)
                .ThenBy(item => item.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();
            categoryOrder[cashType] = orderedCategories;

            foreach (var category in orderedCategories)
            {
                var categoryId = Canonical(category.Code);
                rows.Add(new RowDraft(new StatementRow(
                    new StatementRowId($"flow:{CashTypeId(cashType)}:{categoryId}:header"),
                    StatementRowKind.Header,
                    order++,
                    category.Name,
                    category.Code,
                    StatementPresentationRole.SectionHeader,
                    ImmutableDictionary<StatementColumnId, StatementCell>.Empty)));

                var detailRows = new List<RowDraft>();
                foreach (var cashCode in category.CashCodes)
                {
                    var rowId = new StatementRowId($"flow:{categoryId}:cash-code:{Canonical(cashCode.Code)}");
                    var cells = BuildDetailCells(rowId, grid, cashCode.Values);
                    var detail = new RowDraft(new StatementRow(
                        rowId,
                        StatementRowKind.Detail,
                        order++,
                        cashCode.Description,
                        cashCode.Code,
                        StatementPresentationRole.Detail,
                        cells));
                    rows.Add(detail);
                    detailRows.Add(detail);
                }

                var totalRowId = new StatementRowId($"flow:{categoryId}:total");
                var total = new RowDraft(new StatementRow(
                    totalRowId,
                    StatementRowKind.CategoryTotal,
                    order++,
                    text.Get(CashStatementText.Totals),
                    category.Code,
                    TotalRole(category.PolarityCode),
                    BuildCategoryTotalCells(totalRowId, grid, detailRows, category.PolarityCode)));
                rows.Add(total);

                if (!rowByCategory.TryAdd(category.Code, total))
                {
                    violations.Add(new(
                        "CATEGORY_CODE_DUPLICATE",
                        $"flow.{categoryId}",
                        $"Category code '{category.Code}' occurs more than once."));
                }
            }

            flowRows[cashType] = rows;
        }

        var calculatedRows = new Dictionary<CashType, List<RowDraft>>();
        foreach (var cashType in new[] { CashType.Trade, CashType.Tax })
        {
            var totals = source.CalculatedTotals.Where(total => total.CashType == cashType).ToList();
            var rows = new List<RowDraft>();
            if (totals.Count >= 2)
            {
                for (var order = 0; order < totals.Count; order++)
                {
                    var total = totals[order];
                    var row = new RowDraft(new StatementRow(
                        new StatementRowId($"calculated-total:{Canonical(total.Code)}"),
                        StatementRowKind.CalculatedTotal,
                        order,
                        total.Name,
                        total.Code,
                        StatementPresentationRole.CalculatedTotalHeader,
                        grid.Periods.ToImmutableDictionary(
                            period => PeriodColumn(period.Id),
                            _ => PendingMoneyCell(
                                new StatementFormulaAggregate(
                                    StatementAggregateOperator.Sum,
                                    ImmutableArray<StatementFormula>.Empty),
                                StatementPresentationRole.CalculatedTotalHeader))));
                    rows.Add(row);
                    if (!rowByCategory.TryAdd(total.Code, row))
                    {
                        violations.Add(new(
                            "CATEGORY_CODE_DUPLICATE",
                            $"calculated-total.{Canonical(total.Code)}",
                            $"Category code '{total.Code}' occurs more than once."));
                    }
                }
            }

            calculatedRows[cashType] = rows;
        }

        foreach (var total in source.TotalChildren)
        {
            if (!rowByCategory.TryGetValue(total.Key, out var target))
            {
                // The production workbook suppresses a calculated-total section when it
                // contains only one row (currently CT-VAT). Its repository dependency is
                // consequently non-materialized and must not become a dangling model row.
                if (source.CalculatedTotals.Any(candidate =>
                    candidate.Code.Equals(total.Key, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }
                violations.Add(new(
                    "CATEGORY_TOTAL_NOT_FOUND",
                    $"flow.{Canonical(total.Key)}",
                    $"Calculated category '{total.Key}' is not present in the flow categories."));
                continue;
            }

            if (total.Value.IsEmpty)
                continue;

            var cells = target.Row.Cells.ToBuilder();
            foreach (var period in grid.Periods)
            {
                var columnId = PeriodColumn(period.Id);
                var operands = total.Value.Select(childCode =>
                {
                    var childRowId = rowByCategory.TryGetValue(childCode, out var child)
                        ? child.Row.Id
                        : new StatementRowId($"flow:{Canonical(childCode)}:total");
                    return (StatementFormula)new StatementFormulaReference(new StatementCellAddress(
                        CashFlowSheetId,
                        childRowId,
                        columnId));
                }).ToImmutableArray();
                cells[columnId] = PendingMoneyCell(new StatementFormulaAggregate(
                    StatementAggregateOperator.Sum,
                    operands), target.Row.PresentationRole);
            }

            target.Row = target.Row with { Cells = cells.ToImmutable() };
        }

        var sections = ImmutableArray.CreateBuilder<StatementSection>();
        AddFlowAndSummary(CashType.Trade);
        AddFlowAndSummary(CashType.Money);
        AddCalculated(CashType.Trade);
        AddFlowAndSummary(CashType.Tax);
        AddCalculated(CashType.Tax);
        return sections.ToImmutable();

        void AddFlowAndSummary(CashType cashType)
        {
            var cashTypeId = CashTypeId(cashType);
            sections.Add(new StatementSection(
                new StatementSectionId($"flow:{cashTypeId}"),
                StatementSectionKind.FlowCategories,
                sections.Count,
                text.Get(FlowSectionText(cashType)),
                StatementPresentationRole.SectionHeader,
                flowRows[cashType].Select(row => row.Row).ToImmutableArray()));

            var categories = categoryOrder[cashType];
            if (categories.Count < 2)
                return;

            var summaryRows = ImmutableArray.CreateBuilder<StatementRow>();
            foreach (var category in categories)
            {
                if (!rowByCategory.TryGetValue(category.Code, out var categoryTotal))
                    continue;

                var rowId = new StatementRowId($"summary:{cashTypeId}:{Canonical(category.Code)}");
                summaryRows.Add(new StatementRow(
                    rowId,
                    StatementRowKind.SummaryReference,
                    summaryRows.Count,
                    category.Name,
                    category.Code,
                    StatementPresentationRole.SummaryTotal,
                    grid.Periods.ToImmutableDictionary(
                        period => PeriodColumn(period.Id),
                        period => PendingMoneyCell(
                            new StatementFormulaReference(new StatementCellAddress(
                                CashFlowSheetId,
                                categoryTotal.Row.Id,
                                PeriodColumn(period.Id))),
                            StatementPresentationRole.SummaryTotal))));
            }

            var periodTotalId = new StatementRowId($"summary:{cashTypeId}:period-total");
            summaryRows.Add(new StatementRow(
                periodTotalId,
                StatementRowKind.PeriodTotal,
                summaryRows.Count,
                text.Get(CashStatementText.PeriodTotal),
                null,
                StatementPresentationRole.SummaryTotal,
                grid.Periods.ToImmutableDictionary(
                    period => PeriodColumn(period.Id),
                    period => PendingMoneyCell(
                        new StatementFormulaAggregate(
                            StatementAggregateOperator.Sum,
                            summaryRows.Select(row =>
                                (StatementFormula)new StatementFormulaReference(new StatementCellAddress(
                                    CashFlowSheetId,
                                    row.Id,
                                    PeriodColumn(period.Id)))).ToImmutableArray()),
                        StatementPresentationRole.SummaryTotal))));

            sections.Add(new StatementSection(
                new StatementSectionId($"summary:{cashTypeId}"),
                StatementSectionKind.FlowSummary,
                sections.Count,
                text.Get(CashStatementText.Summary),
                StatementPresentationRole.SummaryHeader,
                summaryRows.ToImmutable()));
        }

        void AddCalculated(CashType cashType)
        {
            var rows = calculatedRows[cashType];
            if (rows.Count == 0)
                return;

            sections.Add(new StatementSection(
                new StatementSectionId($"calculated-total:{CashTypeId(cashType)}"),
                StatementSectionKind.CategoryTotals,
                sections.Count,
                text.Get(TotalsSectionText(cashType)),
                StatementPresentationRole.CalculatedTotalHeader,
                rows.Select(row => row.Row).ToImmutableArray()));
        }
    }

    private async Task<StatementSection> BuildBankBalancesSectionAsync(
        string connectionString,
        StatementPeriodGrid grid,
        int presentationOrder,
        int commandTimeoutSeconds,
        ImmutableArray<StatementContractViolation>.Builder violations,
        CancellationToken ct)
    {
        var accounts = await _repository.GetBankAccountsAsync(
            connectionString,
            commandTimeoutSeconds,
            ct);
        var rows = ImmutableArray.CreateBuilder<StatementRow>();
        var accountCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var monthlyPeriods = grid.Periods
            .Where(period => !period.Id.IsAnnual)
            .Select(period => period.Id)
            .ToHashSet();

        foreach (var account in accounts)
        {
            ct.ThrowIfCancellationRequested();
            var accountId = Canonical(account.AccountCode);
            if (string.IsNullOrEmpty(accountId))
            {
                violations.Add(new(
                    "BANK_ACCOUNT_CODE_REQUIRED",
                    "bank-balances.accounts",
                    "A bank account was returned without an account code."));
                continue;
            }
            if (!accountCodes.Add(account.AccountCode))
            {
                violations.Add(new(
                    "BANK_ACCOUNT_DUPLICATE",
                    $"bank-balances.{accountId}",
                    $"Bank account '{account.AccountCode}' occurs more than once."));
                continue;
            }

            var balances = await _repository.GetBankBalancesAsync(
                connectionString,
                account.AccountCode,
                commandTimeoutSeconds,
                ct);
            var values = new Dictionary<StatementPeriodId, decimal>();
            foreach (var balance in balances)
            {
                var periodId = StatementPeriodId.Month(balance.YearNumber, balance.MonthNumber);
                if (!monthlyPeriods.Contains(periodId))
                {
                    violations.Add(new(
                        "BANK_PERIOD_NOT_FOUND",
                        $"bank-balances.{accountId}.{periodId}",
                        $"Bank account '{account.AccountCode}' returned a balance outside the statement period grid."));
                    continue;
                }
                if (!values.TryAdd(periodId, balance.Balance))
                {
                    violations.Add(new(
                        "BANK_BALANCE_DUPLICATE",
                        $"bank-balances.{accountId}.{periodId}",
                        $"Bank account '{account.AccountCode}' returned more than one balance for period '{periodId}'."));
                }
            }

            var rowId = new StatementRowId($"bank:account:{accountId}");
            var cells = ImmutableDictionary.CreateBuilder<StatementColumnId, StatementCell>();
            foreach (var period in grid.Periods)
            {
                var columnId = PeriodColumn(period.Id);
                if (!period.Id.IsAnnual)
                {
                    cells[columnId] = new(
                        StatementValue.FromDecimal(values.GetValueOrDefault(period.Id)),
                        StatementNumericKind.Money,
                        StatementPresentationRole.Detail);
                    continue;
                }

                var month12 = StatementPeriodId.Month(period.Id.YearNumber, 12);
                cells[columnId] = monthlyPeriods.Contains(month12)
                    ? PendingMoneyCell(
                        new StatementFormulaReference(new StatementCellAddress(
                            CashFlowSheetId,
                            rowId,
                            PeriodColumn(month12))),
                        StatementPresentationRole.AnnualPeriod)
                    : new StatementCell(
                        StatementValue.Missing(StatementScalarKind.Decimal, "month-12-not-in-period-grid"),
                        StatementNumericKind.Money,
                        StatementPresentationRole.AnnualPeriod);
            }

            rows.Add(new StatementRow(
                rowId,
                StatementRowKind.BankAccount,
                rows.Count,
                account.AccountName,
                account.AccountCode,
                StatementPresentationRole.Detail,
                cells.ToImmutable()));
        }

        if (rows.Count > 0)
        {
            var totalRowId = new StatementRowId("bank:company-balance");
            rows.Add(new StatementRow(
                totalRowId,
                StatementRowKind.BankTotal,
                rows.Count,
                _text.Get(CashStatementText.CompanyBalance),
                null,
                StatementPresentationRole.NeutralTotal,
                grid.Periods.ToImmutableDictionary(
                    period => PeriodColumn(period.Id),
                    period => PendingMoneyCell(
                        new StatementFormulaAggregate(
                            StatementAggregateOperator.Sum,
                            rows.Select(row =>
                                (StatementFormula)new StatementFormulaReference(new StatementCellAddress(
                                    CashFlowSheetId,
                                    row.Id,
                                    PeriodColumn(period.Id)))).ToImmutableArray()),
                        StatementPresentationRole.NeutralTotal))));
        }

        return new StatementSection(
            new StatementSectionId("bank-balances"),
            StatementSectionKind.BankBalances,
            presentationOrder,
            _text.Get(CashStatementText.ClosingBalances),
            StatementPresentationRole.BankHeader,
            rows.ToImmutable());
    }

    private async Task<StatementSection?> BuildExpressionSectionAsync(
        string connectionString,
        StatementPeriodGrid grid,
        FlowSource source,
        IEnumerable<StatementSection> flowSections,
        int presentationOrder,
        int commandTimeoutSeconds,
        CancellationToken ct)
    {
        var expressions = await _repository.GetCategoryExpressionsAsync(
            connectionString,
            commandTimeoutSeconds,
            ct);
        if (expressions.Count == 0)
            return null;

        var catalogue = source.Categories
            .Select(category => new CashStatementCategoryReference(category.Code, category.Name))
            .Concat(source.Categories.SelectMany(category => category.CashCodes.Select(cashCode =>
                new CashStatementCategoryReference(
                    cashCode.Code,
                    cashCode.Description,
                    CashStatementReferenceKind.CashCode))))
            .Concat(source.CalculatedTotals.Select(total =>
                new CashStatementCategoryReference(total.Code, total.Name)))
            .ToImmutableArray();

        var rowsByCode = flowSections
            .SelectMany(section => section.Rows)
            .Where(row =>
                !string.IsNullOrWhiteSpace(row.BusinessCode) &&
                row.Kind is StatementRowKind.Detail or StatementRowKind.CategoryTotal or StatementRowKind.CalculatedTotal)
            .GroupBy(row => row.BusinessCode!, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.ToImmutableArray(), StringComparer.OrdinalIgnoreCase);

        var rows = ImmutableArray.CreateBuilder<StatementRow>();
        foreach (var expression in expressions
            .OrderBy(item => item.DisplayOrder)
            .ThenBy(item => item.Category, StringComparer.OrdinalIgnoreCase))
        {
            ct.ThrowIfCancellationRequested();
            var expressionCode = expression.CategoryCode;
            if (string.IsNullOrWhiteSpace(expressionCode))
            {
                expressionCode = await _repository.GetCategoryCodeFromNameAsync(
                    connectionString,
                    expression.Category,
                    commandTimeoutSeconds,
                    ct);
            }

            var translation = CashStatementExpressionTranslator.ForStorage(
                expression.Expression ?? string.Empty,
                catalogue);
            var errors = translation.Diagnostics.Select(diagnostic => diagnostic.Message).ToList();
            var referenceRows = new Dictionary<string, StatementRow>(StringComparer.OrdinalIgnoreCase);
            foreach (var code in translation.CategoryCodes)
            {
                if (!rowsByCode.TryGetValue(code, out var matches) || matches.Length != 1)
                {
                    errors.Add(matches.IsDefaultOrEmpty
                        ? $"Reference row not found: [{code}]"
                        : $"Reference row is ambiguous: [{code}]");
                    continue;
                }
                referenceRows.Add(code, matches[0]);
            }

            var expressionIdentity = string.IsNullOrWhiteSpace(expressionCode)
                ? expression.Category
                : expressionCode;
            var rowId = new StatementRowId($"expression:{Canonical(expressionIdentity)}");
            var cells = ImmutableDictionary.CreateBuilder<StatementColumnId, StatementCell>();
            foreach (var period in grid.Periods)
            {
                var columnId = PeriodColumn(period.Id);
                if (errors.Count > 0)
                {
                    cells[columnId] = new StatementCell(
                        StatementValue.Unsupported(
                            StatementScalarKind.Decimal,
                            string.Join("; ", errors.Distinct(StringComparer.Ordinal))),
                        ExpressionNumericKind(expression.Format),
                        StatementPresentationRole.AnalysisValue,
                        NumberFormat: expression.Format);
                    continue;
                }

                var references = referenceRows.ToImmutableDictionary(
                    pair => pair.Key,
                    pair => new StatementCellAddress(
                        CashFlowSheetId,
                        pair.Value.Id,
                        columnId),
                    StringComparer.OrdinalIgnoreCase);
                cells[columnId] = new StatementCell(
                    StatementValue.Unsupported(
                        StatementScalarKind.Decimal,
                        "spreadsheet-engine-evaluation-required"),
                    ExpressionNumericKind(expression.Format),
                    StatementPresentationRole.AnalysisValue,
                    new StatementFormulaTerminal(translation.Expression, references),
                    expression.Format);
            }

            rows.Add(new StatementRow(
                rowId,
                StatementRowKind.Expression,
                rows.Count,
                expression.Category,
                expressionCode,
                StatementPresentationRole.AnalysisValue,
                cells.ToImmutable()));
        }

        return new StatementSection(
            new StatementSectionId("analysis"),
            StatementSectionKind.Analysis,
            presentationOrder,
            _text.Get(CashStatementText.Analysis),
            StatementPresentationRole.AnalysisHeader,
            rows.ToImmutable());
    }

    private async Task<ImmutableArray<StatementSection>> BuildVatSectionsAsync(
        string connectionString,
        StatementPeriodGrid grid,
        int firstPresentationOrder,
        CashStatementRequestV1 request,
        DateTime generatedAtUtc,
        ImmutableArray<StatementContractViolation>.Builder violations,
        CancellationToken ct)
    {
        var timeout = request.CommandTimeoutSeconds;
        var recurrenceType = await _repository.GetVatRecurrenceTypeAsync(connectionString, timeout, ct);
        var recurrence = (await _repository.GetVatRecurrenceAsync(connectionString, timeout, ct))
            .Select(ToVatPeriodSource)
            .ToImmutableArray();
        var recurrenceAccruals = request.IncludeTaxAccruals
            ? (await _repository.GetVatRecurrenceAccrualsAsync(connectionString, timeout, ct))
                .Select(ToVatAccrualSource)
                .ToImmutableArray()
            : ImmutableArray<VatAccrualSource>.Empty;
        var periodTotals = (await _repository.GetVatPeriodTotalsAsync(connectionString, timeout, ct))
            .Select(ToVatPeriodSource)
            .ToImmutableArray();
        var periodAccruals = request.IncludeTaxAccruals
            ? (await _repository.GetVatPeriodAccrualsAsync(connectionString, timeout, ct))
                .Select(ToVatAccrualSource)
                .ToImmutableArray()
            : ImmutableArray<VatAccrualSource>.Empty;

        var vat = _text.Get(CashStatementText.Vat);
        var recurrenceTitle = string.IsNullOrWhiteSpace(recurrenceType)
            ? vat
            : $"{vat} {recurrenceType.ToUpperInvariant()}";
        var totalsTitle = $"{vat} {_text.Get(CashStatementText.Totals)}";

        return ImmutableArray.Create(
            BuildVatSection(
                new StatementSectionId("vat:recurrence"),
                StatementSectionKind.VatRecurrence,
                firstPresentationOrder,
                recurrenceTitle,
                grid,
                recurrence,
                recurrenceAccruals,
                request.IncludeActivePeriods,
                generatedAtUtc,
                violations),
            BuildVatSection(
                new StatementSectionId("vat:periods"),
                StatementSectionKind.VatPeriods,
                firstPresentationOrder + 1,
                totalsTitle,
                grid,
                periodTotals,
                periodAccruals,
                request.IncludeActivePeriods,
                generatedAtUtc,
                violations));
    }

    private StatementSection BuildVatSection(
        StatementSectionId sectionId,
        StatementSectionKind sectionKind,
        int presentationOrder,
        string title,
        StatementPeriodGrid grid,
        ImmutableArray<VatPeriodSource> sourcePeriods,
        ImmutableArray<VatAccrualSource> accruals,
        bool includeActivePeriods,
        DateTime generatedAtUtc,
        ImmutableArray<StatementContractViolation>.Builder violations)
    {
        var monthlyPeriods = grid.Periods
            .Where(period => !period.Id.IsAnnual)
            .Select(period => period.Id)
            .ToHashSet();
        var sourceByPeriod = new Dictionary<StatementPeriodId, VatPeriodSource>();
        var sourceByStart = new Dictionary<(short Year, DateTime StartOn), VatPeriodSource>();

        foreach (var source in sourcePeriods)
        {
            var periodId = StatementPeriodId.Month(source.YearNumber, checked((byte)source.StartOn.Month));
            if (!monthlyPeriods.Contains(periodId))
            {
                violations.Add(new(
                    "VAT_PERIOD_NOT_FOUND",
                    $"{sectionId.Value}.{periodId}",
                    $"VAT source period '{source.StartOn:yyyy-MM-dd}' is outside the statement period grid."));
                continue;
            }
            if (!sourceByPeriod.TryAdd(periodId, source))
            {
                violations.Add(new(
                    "VAT_PERIOD_DUPLICATE",
                    $"{sectionId.Value}.{periodId}",
                    $"VAT source returned more than one value for period '{periodId}'."));
                continue;
            }
            sourceByStart[(source.YearNumber, source.StartOn)] = source;
        }

        var accrualsByPeriod = new Dictionary<StatementPeriodId, decimal?[]>();
        foreach (var accrual in accruals)
        {
            if (!sourceByStart.TryGetValue((accrual.YearNumber, accrual.StartOn), out var source))
            {
                violations.Add(new(
                    "VAT_ACCRUAL_PERIOD_NOT_FOUND",
                    $"{sectionId.Value}.accruals.{accrual.StartOn:yyyy-MM-dd}",
                    $"VAT accrual period '{accrual.StartOn:yyyy-MM-dd}' has no matching source period."));
                continue;
            }

            var periodId = StatementPeriodId.Month(source.YearNumber, checked((byte)source.StartOn.Month));
            if (!accrualsByPeriod.TryGetValue(periodId, out var values))
            {
                values = new decimal?[VatMetrics.Length];
                accrualsByPeriod.Add(periodId, values);
            }
            for (var index = 0; index < values.Length; index++)
            {
                if (accrual.Values[index] is { } delta)
                    values[index] = values[index].GetValueOrDefault() + delta;
            }
        }

        var rows = ImmutableArray.CreateBuilder<StatementRow>(VatMetrics.Length);
        foreach (var metric in VatMetrics)
        {
            var rowId = new StatementRowId($"{sectionId.Value}:{metric.Code}");
            var cells = ImmutableDictionary.CreateBuilder<StatementColumnId, StatementCell>();
            foreach (var period in grid.Periods.Where(period => !period.Id.IsAnnual))
            {
                decimal value = 0m;
                if (sourceByPeriod.TryGetValue(period.Id, out var source) &&
                    (includeActivePeriods || source.StartOn <= generatedAtUtc))
                {
                    value = source.Values[metric.ValueIndex];
                }
                if (accrualsByPeriod.TryGetValue(period.Id, out var periodAccruals) &&
                    periodAccruals[metric.ValueIndex] is { } delta)
                {
                    value += delta;
                }
                cells[PeriodColumn(period.Id)] = new StatementCell(
                    StatementValue.FromDecimal(value),
                    StatementNumericKind.Money,
                    StatementPresentationRole.Detail);
            }

            foreach (var annual in grid.Periods.Where(period => period.Id.IsAnnual))
            {
                var operands = grid.Periods
                    .Where(period => !period.Id.IsAnnual && period.Id.YearNumber == annual.Id.YearNumber)
                    .Select(period => (StatementFormula)new StatementFormulaReference(new StatementCellAddress(
                        CashFlowSheetId,
                        rowId,
                        PeriodColumn(period.Id))))
                    .ToImmutableArray();
                cells[PeriodColumn(annual.Id)] = PendingMoneyCell(
                    new StatementFormulaAggregate(StatementAggregateOperator.Sum, operands),
                    StatementPresentationRole.AnnualPeriod);
            }

            var role = metric.Text == CashStatementText.NetVatDue
                ? StatementPresentationRole.NeutralTotal
                : StatementPresentationRole.Detail;
            rows.Add(new StatementRow(
                rowId,
                StatementRowKind.VatMetric,
                rows.Count,
                _text.Get(metric.Text),
                metric.Code,
                role,
                cells.ToImmutable()));
        }

        return new StatementSection(
            sectionId,
            sectionKind,
            presentationOrder,
            title,
            StatementPresentationRole.VatHeader,
            rows.ToImmutable());
    }

    private async Task<StatementSection?> BuildBalanceSheetSectionAsync(
        string connectionString,
        StatementPeriodGrid grid,
        int presentationOrder,
        int commandTimeoutSeconds,
        ImmutableArray<StatementContractViolation>.Builder violations,
        CancellationToken ct)
    {
        var entries = await _repository.GetBalanceSheetAsync(
            connectionString,
            commandTimeoutSeconds,
            ct);
        if (entries.Count == 0)
            return null;

        var monthlyPeriods = grid.Periods
            .Where(period => !period.Id.IsAnnual)
            .Select(period => period.Id)
            .ToHashSet();
        var groups = new Dictionary<string, BalanceSheetSource>(StringComparer.OrdinalIgnoreCase);
        var order = new List<string>();

        foreach (var entry in entries)
        {
            ct.ThrowIfCancellationRequested();
            var periodId = StatementPeriodId.Month(entry.YearNumber, entry.MonthNumber);
            if (!monthlyPeriods.Contains(periodId))
            {
                violations.Add(new(
                    "BALANCE_SHEET_PERIOD_NOT_FOUND",
                    $"balance-sheet.{Canonical(entry.AssetCode)}.{periodId}",
                    $"Balance-sheet item '{entry.AssetCode}' returned a balance outside the statement period grid."));
                continue;
            }

            var key = $"{entry.AssetCode}\u001f{entry.AssetName}";
            if (!groups.TryGetValue(key, out var source))
            {
                source = new BalanceSheetSource(
                    entry.AssetCode,
                    entry.AssetName,
                    new Dictionary<StatementPeriodId, decimal>());
                groups.Add(key, source);
                order.Add(key);
            }

            source.Values[periodId] = source.Values.GetValueOrDefault(periodId) + entry.Balance;
        }

        var rows = ImmutableArray.CreateBuilder<StatementRow>();
        foreach (var key in order)
        {
            var source = groups[key];
            var rowId = new StatementRowId(
                $"balance-sheet:asset:{Canonical(source.Code)}:{Canonical(source.Name)}");
            var cells = ImmutableDictionary.CreateBuilder<StatementColumnId, StatementCell>();

            foreach (var period in grid.Periods.Where(period => !period.Id.IsAnnual))
            {
                cells[PeriodColumn(period.Id)] = new StatementCell(
                    source.Values.TryGetValue(period.Id, out var value)
                        ? StatementValue.FromDecimal(value)
                        : StatementValue.Missing(StatementScalarKind.Decimal, "source-period-not-returned"),
                    StatementNumericKind.Money,
                    StatementPresentationRole.Detail);
            }

            foreach (var annual in grid.Periods.Where(period => period.Id.IsAnnual))
            {
                var lastReported = grid.Periods
                    .Where(period =>
                        !period.Id.IsAnnual &&
                        period.Id.YearNumber == annual.Id.YearNumber &&
                        source.Values.ContainsKey(period.Id))
                    .OrderBy(period => period.PresentationOrder)
                    .LastOrDefault();
                cells[PeriodColumn(annual.Id)] = lastReported is null
                    ? new StatementCell(
                        StatementValue.Missing(StatementScalarKind.Decimal, "source-year-not-returned"),
                        StatementNumericKind.Money,
                        StatementPresentationRole.AnnualPeriod)
                    : PendingMoneyCell(
                        new StatementFormulaReference(new StatementCellAddress(
                            CashFlowSheetId,
                            rowId,
                            PeriodColumn(lastReported.Id))),
                        StatementPresentationRole.AnnualPeriod);
            }

            rows.Add(new StatementRow(
                rowId,
                StatementRowKind.BalanceSheetAsset,
                rows.Count,
                source.Name,
                source.Code,
                StatementPresentationRole.Detail,
                cells.ToImmutable()));
        }

        var capitalRowId = new StatementRowId("balance-sheet:capital");
        rows.Add(new StatementRow(
            capitalRowId,
            StatementRowKind.Capital,
            rows.Count,
            _text.Get(CashStatementText.Capital),
            null,
            StatementPresentationRole.CapitalTotal,
            grid.Periods.ToImmutableDictionary(
                period => PeriodColumn(period.Id),
                period => PendingMoneyCell(
                    new StatementFormulaAggregate(
                        StatementAggregateOperator.Sum,
                        rows
                            .Where(row =>
                            {
                                var cell = row.Cells[PeriodColumn(period.Id)];
                                return cell.Value.State == StatementValueState.Present || cell.Formula is not null;
                            })
                            .Select(row =>
                                (StatementFormula)new StatementFormulaReference(new StatementCellAddress(
                                    CashFlowSheetId,
                                    row.Id,
                                    PeriodColumn(period.Id))))
                            .ToImmutableArray()),
                    StatementPresentationRole.CapitalTotal))));

        return new StatementSection(
            new StatementSectionId("balance-sheet"),
            StatementSectionKind.BalanceSheet,
            presentationOrder,
            _text.Get(CashStatementText.BalanceSheet),
            StatementPresentationRole.BalanceSheetHeader,
            rows.ToImmutable());
    }

    private async Task<StatementSheet> BuildReconciliationSheetAsync(
        string connectionString,
        int presentationOrder,
        int commandTimeoutSeconds,
        CancellationToken ct)
    {
        const decimal tolerance = 0.10m;
        var sheetId = new StatementSheetId("reconciliation");
        var sourceRows = await _repository.GetEquityReconciliationByYearAsync(
            connectionString,
            commandTimeoutSeconds,
            ct);
        var columns = ImmutableArray.Create(
            ReconciliationColumn("year", CashStatementText.Year, StatementColumnKind.Identifier, 0),
            ReconciliationColumn("description", CashStatementText.Description, StatementColumnKind.Label, 1),
            ReconciliationColumn("opening-capital", CashStatementText.OpeningCapital, StatementColumnKind.Attribute, 2),
            ReconciliationColumn("closing-capital", CashStatementText.ClosingCapital, StatementColumnKind.Attribute, 3),
            ReconciliationColumn("profit", CashStatementText.Profit, StatementColumnKind.Attribute, 4),
            ReconciliationColumn("business-tax", CashStatementText.BusinessTax, StatementColumnKind.Attribute, 5),
            ReconciliationColumn("profit-after-tax", CashStatementText.ProfitAfterTax, StatementColumnKind.Result, 6),
            ReconciliationColumn("capital-movement", CashStatementText.CapitalMovement, StatementColumnKind.Attribute, 7),
            ReconciliationColumn("opening-subject-position", CashStatementText.OpeningSubjectPosition, StatementColumnKind.Attribute, 8),
            ReconciliationColumn("opening-account-position", CashStatementText.OpeningAccountPosition, StatementColumnKind.Attribute, 9),
            ReconciliationColumn("capital-delta", CashStatementText.CapitalDelta, StatementColumnKind.Result, 10),
            ReconciliationColumn("bridge-total", CashStatementText.BridgeTotal, StatementColumnKind.Result, 11),
            ReconciliationColumn("difference", CashStatementText.Difference, StatementColumnKind.Result, 12),
            ReconciliationColumn("tolerance", CashStatementText.Tolerance, StatementColumnKind.Attribute, 13),
            ReconciliationColumn("status", CashStatementText.Status, StatementColumnKind.Result, 14));

        var rows = ImmutableArray.CreateBuilder<StatementRow>();
        var overallStatus = sourceRows.Select(row => ReconciliationStatus(row.Variance, tolerance))
            .Aggregate("PASS", WorstReconciliationStatus);
        var overallStatusFormula = OverallStatusFormula(sourceRows.Select(source =>
            new StatementCellAddress(
                sheetId,
                new StatementRowId($"reconciliation:{source.YearNumber}"),
                new StatementColumnId("status"))));
        rows.Add(new StatementRow(
            new StatementRowId("reconciliation:summary"),
            StatementRowKind.Metadata,
            rows.Count,
            _text.Get(CashStatementText.ReconciliationChecks),
            null,
            ReconciliationRole(overallStatus),
            ImmutableDictionary<StatementColumnId, StatementCell>.Empty
                .Add(new StatementColumnId("tolerance"), new StatementCell(
                    StatementValue.FromDecimal(tolerance),
                    StatementNumericKind.Money,
                    StatementPresentationRole.Metadata))
                .Add(new StatementColumnId("status"), new StatementCell(
                    StatementValue.FromText(overallStatus),
                    null,
                    ReconciliationRole(overallStatus),
                    overallStatusFormula))));

        foreach (var source in sourceRows)
        {
            ct.ThrowIfCancellationRequested();
            var rowId = new StatementRowId($"reconciliation:{source.YearNumber}");
            var cells = ImmutableDictionary.CreateBuilder<StatementColumnId, StatementCell>();
            AddDecimal("year", source.YearNumber, StatementNumericKind.Count);
            cells[new StatementColumnId("description")] = new StatementCell(
                StatementValue.FromText(source.Description ?? string.Empty));
            AddDecimal("opening-capital", source.OpeningCapital);
            AddDecimal("closing-capital", source.ClosingCapital);
            AddDecimal("profit", source.Profit);
            AddDecimal("business-tax", source.BusinessTax);
            AddDecimal("profit-after-tax", source.ProfitAfterTax);
            AddDecimal("capital-movement", source.CapitalMovement);
            AddDecimal("opening-subject-position", source.OpeningSubjectPosition);
            AddDecimal("opening-account-position", source.OpeningAccountPosition);
            AddDecimal("capital-delta", source.CapitalDelta);

            var bridgeAddress = ReconciliationAddress("bridge-total");
            var bridgeFormula = SumReferences(
                "profit-after-tax",
                "capital-movement",
                "opening-account-position");
            cells[bridgeAddress.ColumnId] = new StatementCell(
                StatementValue.FromDecimal(
                    source.ProfitAfterTax + source.CapitalMovement + source.OpeningAccountPosition),
                StatementNumericKind.Money,
                StatementPresentationRole.ReconciliationHeader,
                bridgeFormula);

            var differenceAddress = ReconciliationAddress("difference");
            var differenceFormula = new StatementFormulaBinary(
                StatementBinaryOperator.Subtract,
                Reference("capital-delta"),
                new StatementFormulaReference(bridgeAddress));
            cells[differenceAddress.ColumnId] = new StatementCell(
                StatementValue.FromDecimal(source.Variance),
                StatementNumericKind.Money,
                StatementPresentationRole.ReconciliationHeader,
                differenceFormula);

            var status = ReconciliationStatus(source.Variance, tolerance);
            var statusFormula = new StatementFormulaConditional(
                AtMostAbsolute(differenceAddress, tolerance),
                new StatementFormulaLiteral(StatementValue.FromText("PASS")),
                new StatementFormulaConditional(
                    AtMostAbsolute(differenceAddress, tolerance * 10m),
                    new StatementFormulaLiteral(StatementValue.FromText("WARN")),
                    new StatementFormulaLiteral(StatementValue.FromText("FAIL"))));
            cells[new StatementColumnId("status")] = new StatementCell(
                StatementValue.FromText(status),
                null,
                ReconciliationRole(status),
                statusFormula);

            rows.Add(new StatementRow(
                rowId,
                StatementRowKind.Reconciliation,
                rows.Count,
                source.Description ?? source.YearNumber.ToString(),
                source.YearNumber.ToString(),
                ReconciliationRole(status),
                cells.ToImmutable()));

            void AddDecimal(
                string column,
                decimal value,
                StatementNumericKind numericKind = StatementNumericKind.Money) =>
                cells[new StatementColumnId(column)] = new StatementCell(
                    StatementValue.FromDecimal(value),
                    numericKind,
                    StatementPresentationRole.Default);

            StatementCellAddress ReconciliationAddress(string column) =>
                new(sheetId, rowId, new StatementColumnId(column));

            StatementFormula Reference(string column) =>
                new StatementFormulaReference(ReconciliationAddress(column));

            StatementFormula SumReferences(params string[] columnIds) =>
                new StatementFormulaAggregate(
                    StatementAggregateOperator.Sum,
                    columnIds.Select(Reference).ToImmutableArray());
        }

        return new StatementSheet(
            sheetId,
            StatementSheetKind.Reconciliation,
            presentationOrder,
            _text.Get(CashStatementText.Reconciliation),
            columns,
            ImmutableArray.Create(new StatementSection(
                new StatementSectionId("reconciliation"),
                StatementSectionKind.Reconciliation,
                0,
                _text.Get(CashStatementText.Reconciliation),
                StatementPresentationRole.ReconciliationHeader,
                rows.ToImmutable())));

        StatementColumn ReconciliationColumn(
            string id,
            CashStatementText label,
            StatementColumnKind kind,
            int order) => new(new StatementColumnId(id), _text.Get(label), kind, order);

        static StatementFormula AtMostAbsolute(StatementCellAddress address, decimal limit) =>
            new StatementFormulaBinary(
                StatementBinaryOperator.LessThanOrEqual,
                new StatementFormulaUnary(
                    StatementUnaryOperator.Absolute,
                    new StatementFormulaReference(address)),
                new StatementFormulaLiteral(StatementValue.FromDecimal(limit)));

        static StatementFormula OverallStatusFormula(IEnumerable<StatementCellAddress> statusAddresses)
        {
            var addresses = statusAddresses.ToArray();
            StatementFormula result = new StatementFormulaLiteral(StatementValue.FromText("PASS"));
            foreach (var address in addresses.Reverse())
            {
                result = new StatementFormulaConditional(
                    new StatementFormulaBinary(
                        StatementBinaryOperator.Equal,
                        new StatementFormulaReference(address),
                        new StatementFormulaLiteral(StatementValue.FromText("WARN"))),
                    new StatementFormulaLiteral(StatementValue.FromText("WARN")),
                    result);
            }
            foreach (var address in addresses.Reverse())
            {
                result = new StatementFormulaConditional(
                    new StatementFormulaBinary(
                        StatementBinaryOperator.Equal,
                        new StatementFormulaReference(address),
                        new StatementFormulaLiteral(StatementValue.FromText("FAIL"))),
                    new StatementFormulaLiteral(StatementValue.FromText("FAIL")),
                    result);
            }
            return result;
        }
    }

    private static string ReconciliationStatus(decimal variance, decimal tolerance) =>
        Math.Abs(variance) <= tolerance
            ? "PASS"
            : Math.Abs(variance) <= tolerance * 10m
                ? "WARN"
                : "FAIL";

    private static string WorstReconciliationStatus(string left, string right)
    {
        static int Rank(string status) => status switch
        {
            "FAIL" => 2,
            "WARN" => 1,
            _ => 0
        };
        return Rank(right) > Rank(left) ? right : left;
    }

    private static StatementPresentationRole ReconciliationRole(string status) => status switch
    {
        "PASS" => StatementPresentationRole.ReconciliationPass,
        "WARN" => StatementPresentationRole.ReconciliationWarning,
        _ => StatementPresentationRole.ReconciliationFailure
    };

    private static ImmutableDictionary<StatementColumnId, StatementCell> BuildDetailCells(
        StatementRowId rowId,
        StatementPeriodGrid grid,
        ImmutableDictionary<StatementPeriodId, decimal> values)
    {
        var cells = ImmutableDictionary.CreateBuilder<StatementColumnId, StatementCell>();
        foreach (var period in grid.Periods.Where(period => !period.Id.IsAnnual))
        {
            var value = values.TryGetValue(period.Id, out var amount)
                ? StatementValue.FromDecimal(amount)
                : StatementValue.Missing(StatementScalarKind.Decimal, "source-period-not-returned");
            cells[PeriodColumn(period.Id)] = new(value, StatementNumericKind.Money, StatementPresentationRole.Detail);
        }

        foreach (var annual in grid.Periods.Where(period => period.Id.IsAnnual))
        {
            var monthReferences = grid.Periods
                .Where(period => period.Id.YearNumber == annual.Id.YearNumber && !period.Id.IsAnnual)
                .Select(period => (StatementFormula)new StatementFormulaReference(new StatementCellAddress(
                    CashFlowSheetId,
                    rowId,
                    PeriodColumn(period.Id))))
                .ToImmutableArray();
            cells[PeriodColumn(annual.Id)] = PendingMoneyCell(new StatementFormulaAggregate(
                StatementAggregateOperator.Sum,
                monthReferences), StatementPresentationRole.AnnualPeriod);
        }

        return cells.ToImmutable();
    }

    private static ImmutableDictionary<StatementColumnId, StatementCell> BuildCategoryTotalCells(
        StatementRowId rowId,
        StatementPeriodGrid grid,
        IReadOnlyList<RowDraft> detailRows,
        short? polarityCode)
    {
        var cells = ImmutableDictionary.CreateBuilder<StatementColumnId, StatementCell>();
        foreach (var period in grid.Periods)
        {
            var columnId = PeriodColumn(period.Id);
            StatementFormula formula = new StatementFormulaAggregate(
                StatementAggregateOperator.Sum,
                detailRows.Select(detail =>
                    (StatementFormula)new StatementFormulaReference(new StatementCellAddress(
                        CashFlowSheetId,
                        detail.Row.Id,
                        columnId))).ToImmutableArray());
            if (polarityCode == (short)CashPolarity.Expense)
                formula = new StatementFormulaUnary(StatementUnaryOperator.Negate, formula);

            cells[columnId] = PendingMoneyCell(formula, TotalRole(polarityCode));
        }

        return cells.ToImmutable();
    }

    private static StatementCell PendingMoneyCell(
        StatementFormula formula,
        StatementPresentationRole role) =>
        new(
            StatementValue.Missing(StatementScalarKind.Decimal, "pending-evaluation"),
            StatementNumericKind.Money,
            role,
            formula);

    private static StatementPresentationRole TotalRole(short? polarityCode) => polarityCode switch
    {
        (short)CashPolarity.Expense => StatementPresentationRole.ExpenseTotal,
        (short)CashPolarity.Income => StatementPresentationRole.IncomeTotal,
        _ => StatementPresentationRole.NeutralTotal
    };

    private static StatementNumericKind ExpressionNumericKind(string? format) =>
        format?.Contains('%', StringComparison.Ordinal) == true
            ? StatementNumericKind.Percentage
            : StatementNumericKind.Ratio;

    private static bool IncludeActivePeriods(CashType cashType, CashStatementRequestV1 request) =>
        cashType != CashType.Money && request.IncludeActivePeriods;

    private static bool IncludeOrderBook(CashType cashType, CashStatementRequestV1 request) =>
        cashType == CashType.Trade && request.IncludeOrderBook;

    private static bool IncludeTaxAccruals(CashType cashType, CashStatementRequestV1 request) =>
        cashType == CashType.Tax && request.IncludeTaxAccruals;

    private static StatementColumnId PeriodColumn(StatementPeriodId periodId) =>
        new($"period:{periodId}");

    private static string CashTypeId(CashType cashType) => cashType.ToString().ToLowerInvariant();

    private static CashStatementText FlowSectionText(CashType cashType) => cashType switch
    {
        CashType.Trade => CashStatementText.TradeCashFlow,
        CashType.Money => CashStatementText.MoneyCashFlow,
        CashType.Tax => CashStatementText.TaxCashFlow,
        _ => throw new ArgumentOutOfRangeException(nameof(cashType), cashType, null)
    };

    private static CashStatementText TotalsSectionText(CashType cashType) => cashType switch
    {
        CashType.Trade => CashStatementText.TradeTotals,
        CashType.Tax => CashStatementText.TaxTotals,
        _ => throw new ArgumentOutOfRangeException(nameof(cashType), cashType, null)
    };

    private static string Canonical(string value)
    {
        var result = new StringBuilder(value.Length);
        var previousWasSeparator = false;
        foreach (var character in value.Trim().ToLowerInvariant())
        {
            if (char.IsAsciiLetterOrDigit(character) || character is '_' or '.')
            {
                result.Append(character);
                previousWasSeparator = false;
            }
            else if (!previousWasSeparator)
            {
                result.Append('-');
                previousWasSeparator = true;
            }
        }

        return result.ToString().Trim('-');
    }

    private static VatPeriodSource ToVatPeriodSource(VatRecurrenceDto value) => new(
        value.YearNumber,
        value.StartOn,
        ImmutableArray.Create(
            value.VatDueSales,
            value.VatDueAcquisitions,
            value.TotalVatDue,
            value.VatReclaimedCurrPeriod,
            value.TotalValueSalesExVAT,
            value.TotalValuePurchasesExVAT,
            value.TotalValueGoodsSuppliedExVAT,
            value.TotalValueGoodsReceivedExVAT,
            value.NetVatDue));

    private static VatPeriodSource ToVatPeriodSource(VatPeriodTotalDto value) => new(
        value.YearNumber,
        value.StartOn,
        ImmutableArray.Create(
            value.VatDueSales,
            value.VatDueAcquisitions,
            value.TotalVatDue,
            value.VatReclaimedCurrPeriod,
            value.TotalValueSalesExVAT,
            value.TotalValuePurchasesExVAT,
            value.TotalValueGoodsSuppliedExVAT,
            value.TotalValueGoodsReceivedExVAT,
            value.NetVatDue));

    private static VatAccrualSource ToVatAccrualSource(VatRecurrenceAccrualDto value) => new(
        value.YearNumber,
        value.StartOn,
        ImmutableArray.Create(
            value.VatDueSales,
            value.VatDueAcquisitions,
            value.TotalVatDue,
            value.VatReclaimedCurrPeriod,
            value.TotalValueSalesExVAT,
            value.TotalValuePurchasesExVAT,
            value.TotalValueGoodsSuppliedExVAT,
            value.TotalValueGoodsReceivedExVAT,
            value.NetVatDue));

    private static VatAccrualSource ToVatAccrualSource(VatPeriodAccrualDto value) => new(
        value.YearNumber,
        value.StartOn,
        ImmutableArray.Create(
            value.VatDueSales,
            value.VatDueAcquisitions,
            value.TotalVatDue,
            value.VatReclaimedCurrPeriod,
            value.TotalValueSalesExVAT,
            value.TotalValuePurchasesExVAT,
            value.TotalValueGoodsSuppliedExVAT,
            value.TotalValueGoodsReceivedExVAT,
            value.NetVatDue));

    private sealed record CashCodeSource(
        string Code,
        string Description,
        ImmutableDictionary<StatementPeriodId, decimal> Values);

    private sealed record CategorySource(
        CashType CashType,
        string Code,
        string Name,
        short? PolarityCode,
        short? DisplayOrder,
        ImmutableArray<CashCodeSource> CashCodes);

    private sealed record CalculatedTotalSource(
        CashType CashType,
        string Code,
        string Name,
        string CashTypeLabel);

    private sealed record FlowSource(
        ImmutableArray<CategorySource> Categories,
        ImmutableArray<CalculatedTotalSource> CalculatedTotals,
        ImmutableDictionary<string, ImmutableArray<string>> TotalChildren);

    private sealed record VatMetricDefinition(
        string Code,
        CashStatementText Text,
        int ValueIndex);

    private sealed record VatPeriodSource(
        short YearNumber,
        DateTime StartOn,
        ImmutableArray<decimal> Values);

    private sealed record VatAccrualSource(
        short YearNumber,
        DateTime StartOn,
        ImmutableArray<decimal?> Values);

    private sealed record BalanceSheetSource(
        string Code,
        string Name,
        Dictionary<StatementPeriodId, decimal> Values);

    private sealed class RowDraft
    {
        public RowDraft(StatementRow row) => Row = row;
        public StatementRow Row { get; set; }
    }
}
