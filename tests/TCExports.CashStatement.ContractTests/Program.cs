using System.Collections.Immutable;
using System.Diagnostics;
using System.Reflection;
using System.Text.Json;
using TCExports.Generator;
using TCExports.Generator.CashStatement;
using TCExports.Generator.CashStatement.Model;
using TCExports.Generator.Contracts;
using TCExports.Generator.Data;
using TCExports.Generator.Excel;
using TCExports.Generator.Handlers;
using TCExports.Generator.Ods;
using TCExports.Generator.Validation;

var tests = new (string Name, Func<Task> Run)[]
{
    ("request defaults preserve the current boundary contract", Sync(RequestDefaults)),
    ("legacy includeBankTypes is absent", Sync(LegacyOptionIsAbsent)),
    ("zero, missing, and unsupported are distinct", Sync(ValueStatesAreDistinct)),
    ("period identities distinguish monthly and annual cells", Sync(PeriodIdentities)),
    ("formula references use semantic addresses", Sync(FormulaReferencesAreSemantic)),
    ("document collections are immutable", Sync(DocumentCollectionsAreImmutable)),
    ("model types do not depend on spreadsheet or database packages", Sync(ModelIsRendererNeutral)),
    ("default limits cover the accepted baselines while remaining bounded", Sync(LimitsCoverBaselines)),
    ("typed boundary parsing rejects invalid and unknown options", Sync(TypedBoundaryParsing)),
    ("cash export validation accepts both native spreadsheet formats", CashExportValidation),
    ("cash export handler routes both formats through the shared model", NativeFormatRouting),
    ("export runner honours caller cancellation", ExportRunnerCancellation),
    ("builder observes cancellation during source acquisition", BuilderSourceCancellation),
    ("native renderers observe cancellation during serialization", RendererCancellation),
    ("period grid preserves a non-January financial calendar", Sync(NonJanuaryPeriodGrid)),
    ("formula evaluator uses authoritative values and lazy conditionals", Sync(FormulaEvaluation)),
    ("money and ratio calculations preserve deterministic decimal precision", Sync(DecimalPrecisionSemantics)),
    ("forward dependencies are independent of presentation order", Sync(ForwardDependencyValidation)),
    ("missing references and complete cycle paths are reported", Sync(InvalidDependencyValidation)),
    ("expression references are stored by code and displayed by description", Sync(ExpressionReferenceTranslation)),
    ("expression reference failures are explicit and never coerced to zero", Sync(ExpressionReferenceDiagnostics)),
    ("flow builder evaluates forward totals and explicit empty-category zero", FlowBuilderEvaluation),
    ("flow builder rejects a missing category child", FlowBuilderMissingChild),
    ("builder binds opaque expressions and preserves invalid references", ExpressionBuilderBinding),
    ("report text is localized without changing business descriptions", StatementLocalization),
    ("optional sections do not query excluded sources", OptionalSectionsAreNotLoaded),
    ("bank balances preserve closing-balance semantics", BankBalanceBuilderEvaluation),
    ("VAT sections use the period grid, cutoff, and accrual semantics", VatBuilderEvaluation),
    ("balance sheet preserves missing values and closing balances", BalanceSheetBuilderEvaluation),
    ("reconciliation is optional and applies deterministic thresholds", ReconciliationBuilderEvaluation)
};

var failures = new List<string>();
var testResults = new List<ContractTestResult>();
foreach (var test in tests)
{
    var stopwatch = Stopwatch.StartNew();
    try
    {
        await test.Run();
        stopwatch.Stop();
        testResults.Add(new(test.Name, true, null, stopwatch.Elapsed.TotalMilliseconds));
        Console.WriteLine($"PASS {test.Name}");
    }
    catch (Exception ex)
    {
        stopwatch.Stop();
        testResults.Add(new(test.Name, false, ex.Message, stopwatch.Elapsed.TotalMilliseconds));
        failures.Add($"FAIL {test.Name}: {ex.Message}");
    }
}

WriteEvidenceIfRequested(args, testResults);

if (failures.Count > 0)
{
    foreach (var failure in failures)
        Console.Error.WriteLine(failure);

    return 1;
}

Console.WriteLine($"Cash Statement contract tests passed: {tests.Length}/{tests.Length}.");
return 0;

static Func<Task> Sync(Action action) => () =>
{
    action();
    return Task.CompletedTask;
};

static async Task CashExportValidation()
{
    var repository = CreateFlowRepository();
    var handler = new CashStatementHandler(
        new CashStatementBuilder(repository),
        new CashStatementXlsxRenderer(),
        new CashStatementOdsRenderer());
    var validator = new CashPayloadValidator();

    foreach (var format in new[] { "excel", "libre" })
    {
        var payload = new ExportPayload
        {
            SqlConnection = "Server=(test);Database=(test)",
            UserName = "contract-test",
            DocumentType = "cashflow",
            Format = format
        };
        validator.Normalize(payload);
        Require(validator.TryValidate(payload, out var errors),
            $"{format} was rejected: {string.Join(';', errors.SelectMany(error => error.Value))}");
        Require(payload.Params.ContainsKey("includeReconciliation"),
            "reconciliation default was not normalized");
    }

    var invalid = new ExportPayload
    {
        SqlConnection = "Server=(test);Database=(test)",
        UserName = "contract-test",
        DocumentType = "cashflow",
        Format = "csv",
        Params = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["commandTimeout"] = "0",
            ["unknown"] = "true"
        }
    };
    var engine = new ExportEngine(new[] { handler }, new IPayloadValidator[] { validator });
    var result = await engine.GenerateAsync(invalid);
    Equal("error", result.Status, "invalid export status");
    Equal("VALIDATION_ERROR", result.Code, "invalid export code");
    Require(result.Details?.ContainsKey("format") == true, "invalid format was not reported");
    Require(result.Details?.ContainsKey("params.commandTimeout") == true, "invalid timeout was not reported");
    Require(result.Details?.ContainsKey("params.unknown") == true, "unknown option was not reported");
}

static async Task ExportRunnerCancellation()
{
    using var cancellation = new CancellationTokenSource();
    cancellation.Cancel();
    try
    {
        await ExportRunner.ExportDataAsync(new ExportPayload(), cancellation.Token);
    }
    catch (OperationCanceledException)
    {
        return;
    }

    throw new InvalidOperationException("A cancelled caller reached export validation.");
}

static async Task BuilderSourceCancellation()
{
    using var cancellation = new CancellationTokenSource();
    var repository = new StubCashFlowRepository { CancelAfterFirstCall = cancellation };
    var builder = new CashStatementBuilder(repository);

    try
    {
        await builder.BuildAsync(
            "Server=(test);Database=(test)",
            new CashStatementRequestV1(),
            "contract-test",
            DateTimeOffset.UnixEpoch,
            cancellation.Token);
    }
    catch (OperationCanceledException)
    {
        Equal(1, repository.Calls.Count, "repository calls before cancellation");
        return;
    }

    throw new InvalidOperationException("Builder continued after source acquisition was cancelled.");
}

static async Task RendererCancellation()
{
    var build = await new CashStatementBuilder(CreateFlowRepository()).BuildAsync(
        "Server=(test);Database=(test)",
        new CashStatementRequestV1(),
        "contract-test",
        DateTimeOffset.UnixEpoch,
        CancellationToken.None);
    Require(build.IsSuccess, "Serialization cancellation fixture did not build.");
    var document = build.Value!;

    using (var cancellation = new CancellationTokenSource())
    {
        var renderer = new CashStatementXlsxRenderer(
            allocator: new CancellingLayoutAllocator(
                new CashStatementXlsxLayoutAllocator(),
                cancellation));
        try
        {
            renderer.Render(document, cancellation.Token);
        }
        catch (OperationCanceledException)
        {
            goto ods;
        }

        throw new InvalidOperationException("XLSX renderer continued after serialization was cancelled.");
    }

ods:
    using (var cancellation = new CancellationTokenSource())
    {
        var renderer = new CashStatementOdsRenderer(
            allocator: new CancellingLayoutAllocator(
                new CashStatementOdsLayoutAllocator(),
                cancellation));
        try
        {
            await renderer.RenderAsync(document, cancellation.Token);
        }
        catch (OperationCanceledException)
        {
            return;
        }

        throw new InvalidOperationException("ODS renderer continued after serialization was cancelled.");
    }
}

static async Task NativeFormatRouting()
{
    foreach (var format in new[] { "excel", "libre" })
    {
        var handler = new CashStatementHandler(
            new CashStatementBuilder(CreateFlowRepository()),
            new CashStatementXlsxRenderer(),
            new CashStatementOdsRenderer());
        var payload = new ExportPayload
        {
            SqlConnection = "Server=(test);Database=(test)",
            UserName = "contract-test",
            DocumentType = "cashflow",
            Format = format,
            Params = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["includeActivePeriods"] = "true",
                ["includeOrderBook"] = "true",
                ["includeTaxAccruals"] = "true",
                ["includeVatDetails"] = "false",
                ["includeBankBalances"] = "false",
                ["includeBalanceSheet"] = "false",
                ["includeReconciliation"] = "false",
                ["commandTimeout"] = "30"
            }
        };

        var result = await handler.HandleAsync(payload, CancellationToken.None);
        Equal("success", result.Status, $"{format} status");
        var extension = format == "excel" ? ".xlsx" : ".ods";
        Require(result.FileName?.EndsWith(extension, StringComparison.OrdinalIgnoreCase) == true,
            $"{format} file extension was incorrect");
        var bytes = Convert.FromBase64String(result.FileContent!);
        Require(bytes.Length > 1_000, $"{format} output was unexpectedly small");
        Equal((byte)'P', bytes[0], $"{format} ZIP signature byte 1");
        Equal((byte)'K', bytes[1], $"{format} ZIP signature byte 2");
    }
}

static void WriteEvidenceIfRequested(string[] arguments, IReadOnlyList<ContractTestResult> results)
{
    var optionIndex = Array.FindIndex(arguments, argument =>
        argument.Equals("--evidence", StringComparison.OrdinalIgnoreCase));
    if (optionIndex < 0)
        return;
    if (optionIndex == arguments.Length - 1 || string.IsNullOrWhiteSpace(arguments[optionIndex + 1]))
        throw new ArgumentException("--evidence requires an output path.");

    var path = Path.GetFullPath(arguments[optionIndex + 1]);
    Directory.CreateDirectory(Path.GetDirectoryName(path)!);
    var evidence = new ContractTestEvidence(
        DateTimeOffset.UtcNow,
        Environment.Version.ToString(),
        results.Count,
        results.Count(result => result.Passed),
        results.Count(result => !result.Passed),
        results);
    var json = JsonSerializer.Serialize(evidence, new JsonSerializerOptions
    {
        WriteIndented = true
    });
    File.WriteAllText(
        path,
        json.ReplaceLineEndings("\n"),
        new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
    Console.WriteLine($"Evidence written: {path}");
}

static void RequestDefaults()
{
    var request = new CashStatementRequestV1();
    Equal(false, request.IncludeActivePeriods, nameof(request.IncludeActivePeriods));
    Equal(false, request.IncludeOrderBook, nameof(request.IncludeOrderBook));
    Equal(true, request.IncludeTaxAccruals, nameof(request.IncludeTaxAccruals));
    Equal(true, request.IncludeVatDetails, nameof(request.IncludeVatDetails));
    Equal(true, request.IncludeBankBalances, nameof(request.IncludeBankBalances));
    Equal(true, request.IncludeBalanceSheet, nameof(request.IncludeBalanceSheet));
    Equal(false, request.IncludeReconciliation, nameof(request.IncludeReconciliation));
    Equal(30, request.CommandTimeoutSeconds, nameof(request.CommandTimeoutSeconds));
    Equal(1, CashStatementRequestV1.ContractVersion, nameof(CashStatementRequestV1.ContractVersion));
    Equal(1, request.Version, nameof(request.Version));
}

static void LegacyOptionIsAbsent()
{
    var optionNames = typeof(CashStatementRequestV1)
        .GetProperties(BindingFlags.Instance | BindingFlags.Public)
        .Select(p => p.Name)
        .ToHashSet(StringComparer.OrdinalIgnoreCase);

    Require(!optionNames.Contains("includeBankTypes"), "includeBankTypes leaked into request v1");
}

static void ValueStatesAreDistinct()
{
    var zero = StatementValue.FromDecimal(0m);
    var missing = StatementValue.Missing(StatementScalarKind.Decimal, "not supplied");
    var unsupported = StatementValue.Unsupported(StatementScalarKind.Decimal, "division-by-zero");

    Equal(StatementValueState.Present, zero.State, "zero state");
    Equal(0m, zero.DecimalValue, "zero value");
    Equal(StatementValueState.Missing, missing.State, "missing state");
    Equal(null, missing.DecimalValue, "missing numeric value");
    Equal(StatementValueState.Unsupported, unsupported.State, "unsupported state");
    Equal("division-by-zero", unsupported.Reason, "unsupported reason");
}

static void PeriodIdentities()
{
    var month = StatementPeriodId.Month(2026, 4);
    var annual = StatementPeriodId.Annual(2026);

    Require(!month.IsAnnual, "monthly period identified as annual");
    Require(annual.IsAnnual, "annual period identified as monthly");
    Require(month != annual, "monthly and annual identities collided");
    Equal("2026:04", month.ToString(), "monthly canonical identity");
    Equal("2026:annual", annual.ToString(), "annual canonical identity");
}

static void FormulaReferencesAreSemantic()
{
    var address = new StatementCellAddress(
        new StatementSheetId("cash-flow"),
        new StatementRowId("flow:fees:total"),
        new StatementColumnId("period:2026:04"));

    StatementFormula formula = new StatementFormulaReference(address);
    var reference = RequireType<StatementFormulaReference>(formula);
    Equal(address, reference.Address, "semantic formula address");
}

static void DocumentCollectionsAreImmutable()
{
    var sheetArray = ImmutableArray.Create(
        new StatementSheet(
            new StatementSheetId("cash-flow"),
            StatementSheetKind.CashFlow,
            0,
            "Cash Flow",
            ImmutableArray<StatementColumn>.Empty,
            ImmutableArray<StatementSection>.Empty));

    var document = new CashStatementDocument(
        new CashStatementMetadata(
            "Example",
            "Cash Statement",
            "owner",
            DateTimeOffset.UnixEpoch,
            StatementPeriodId.Month(2026, 1)),
        new CashStatementRequestV1(),
        new StatementPeriodGrid(ImmutableArray<StatementPeriod>.Empty),
        sheetArray);

    Equal(1, document.Sheets.Length, "sheet count");
    Require(document.Sheets.GetType().Name.StartsWith("ImmutableArray", StringComparison.Ordinal),
        "document sheets are not immutable");
}

static void ModelIsRendererNeutral()
{
    var modelAssembly = typeof(CashStatementDocument).Assembly;
    var forbiddenPrefixes = new[]
    {
        "ClosedXML.",
        "DocumentFormat.OpenXml.",
        "Microsoft.Data.SqlClient."
    };

    var modelTypes = modelAssembly.GetTypes()
        .Where(t => t.Namespace == typeof(CashStatementDocument).Namespace);

    foreach (var type in modelTypes)
    {
        var exposedTypes = type.GetProperties().Select(p => p.PropertyType)
            .Concat(type.GetMethods(BindingFlags.Instance | BindingFlags.Public | BindingFlags.DeclaredOnly)
                .Select(m => m.ReturnType));

        foreach (var exposedType in exposedTypes)
        {
            var name = exposedType.FullName ?? exposedType.Name;
            Require(!forbiddenPrefixes.Any(name.StartsWith),
                $"{type.Name} exposes renderer/database type {name}");
        }
    }
}

static void ExpressionReferenceTranslation()
{
    var categories = new[]
    {
        new CashStatementCategoryReference("CAT-SALES", "Sales"),
        new CashStatementCategoryReference("CT-GROSS", "Gross Profit"),
        new CashStatementCategoryReference("CC-SALES", "Sales", CashStatementReferenceKind.CashCode),
        new CashStatementCategoryReference("CC-LABOUR", "Sales - Labour", CashStatementReferenceKind.CashCode)
    };

    const string entered = "IF([Sales]=0,0,[Gross Profit]/[Sales])";
    var stored = CashStatementExpressionTranslator.ForStorage(entered, categories);
    Require(stored.IsValid, "description-based expression was not accepted");
    Equal("IF([CAT-SALES]=0,0,[CT-GROSS]/[CAT-SALES])", stored.Expression, "canonical stored expression");
    Equal(2, stored.CategoryCodes.Length, "distinct category references");

    var renamedCategories = new[]
    {
        new CashStatementCategoryReference("CAT-SALES", "Revenue"),
        new CashStatementCategoryReference("CT-GROSS", "Gross Margin")
    };
    var displayed = CashStatementExpressionTranslator.ForDisplay(stored.Expression, renamedCategories);
    Require(displayed.IsValid, "stored codes did not survive category renames");
    Equal("IF([Revenue]=0,0,[Gross Margin]/[Revenue])", displayed.Expression, "renamed display expression");

    var codeFirst = CashStatementExpressionTranslator.ForStorage("[CAT-SALES]+[Gross Margin]", renamedCategories);
    Require(codeFirst.IsValid, "mixed code and description references were not accepted");
    Equal("[CAT-SALES]+[CT-GROSS]", codeFirst.Expression, "code-first compatibility");

    var detail = CashStatementExpressionTranslator.ForStorage("[Sales]+[Sales - Labour]", categories);
    Require(detail.IsValid, "category and cash-code references were not accepted together");
    Equal("[CAT-SALES]+[CC-LABOUR]", detail.Expression, "category-before-cash-code precedence");
}

static void ExpressionReferenceDiagnostics()
{
    var categories = new[]
    {
        new CashStatementCategoryReference("CAT-A", "Duplicate"),
        new CashStatementCategoryReference("CAT-B", "Duplicate")
    };

    var result = CashStatementExpressionTranslator.ForStorage(
        "IF([Missing]=0,[],[Duplicate])+[CAT-A]+SUM(1,2)",
        categories);

    Require(!result.IsValid, "invalid references were accepted");
    Equal("EXPRESSION_TAG_NOT_FOUND", result.Diagnostics[0].Code, "missing reference diagnostic");
    Equal("EXPRESSION_TAG_EMPTY", result.Diagnostics[1].Code, "empty reference diagnostic");
    Equal("EXPRESSION_TAG_AMBIGUOUS", result.Diagnostics[2].Code, "ambiguous reference diagnostic");
    Require(result.Expression.Contains("[Missing]", StringComparison.Ordinal), "missing reference was rewritten");
    Require(!result.Expression.Contains("[0]", StringComparison.Ordinal), "missing reference was coerced to zero");
    Require(result.Expression.EndsWith("+SUM(1,2)", StringComparison.Ordinal), "opaque formula text was changed");
}

static void LimitsCoverBaselines()
{
    var limits = CashStatementLimits.Default;

    Require(limits.MaxFiscalYears >= 5, "fiscal-year limit is below baseline");
    Require(limits.MaxPeriodColumns >= 65, "period-column limit is below baseline");
    Require(limits.MaxSheets >= 3, "sheet limit is below baseline");
    Require(limits.MaxRows >= 220, "row limit is below baseline");
    Require(limits.MaxCells >= 11_278, "cell limit is below baseline");
    Require(limits.MaxFormulaNodes >= 4_468, "formula limit is below baseline");
    Require(limits.MaxOutputBytes >= 87_184, "output limit is below baseline");

    Require(limits.MaxFiscalYears < int.MaxValue, "fiscal-year limit is unbounded");
    Require(limits.MaxRows < int.MaxValue, "row limit is unbounded");
    Require(limits.MaxCells < int.MaxValue, "cell limit is unbounded");
    Require(limits.MaxOutputBytes < long.MaxValue, "output limit is unbounded");
}

static void TypedBoundaryParsing()
{
    var valid = CashStatementRequestParser.Parse(new Dictionary<string, string>
    {
        ["includeActivePeriods"] = "yes",
        ["includeOrderBook"] = "1",
        ["includeTaxAccruals"] = "false",
        ["includeVatDetails"] = "true",
        ["includeBankBalances"] = "0",
        ["includeBalanceSheet"] = "no",
        ["includeReconciliation"] = "y",
        ["commandTimeout"] = "75"
    });

    Require(valid.IsSuccess, "valid typed request did not parse");
    Equal(true, valid.Value!.IncludeActivePeriods, "active periods");
    Equal(true, valid.Value.IncludeOrderBook, "order book");
    Equal(false, valid.Value.IncludeTaxAccruals, "tax accruals");
    Equal(false, valid.Value.IncludeBankBalances, "bank balances");
    Equal(false, valid.Value.IncludeBalanceSheet, "balance sheet");
    Equal(true, valid.Value.IncludeReconciliation, "reconciliation");
    Equal(75, valid.Value.CommandTimeoutSeconds, "command timeout");

    var invalid = CashStatementRequestParser.Parse(new Dictionary<string, string>
    {
        ["includeActivePeriods"] = "perhaps",
        ["includeBankTypes"] = "true",
        ["commandTimeout"] = "601"
    });

    Require(!invalid.IsSuccess, "invalid request parsed successfully");
    Require(invalid.Violations.Any(v => v.Code == "OPTION_BOOLEAN_INVALID"), "invalid Boolean was not reported");
    Require(invalid.Violations.Any(v => v.Code == "OPTION_UNKNOWN"), "unknown option was not reported");
    Require(invalid.Violations.Any(v => v.Code == "OPTION_TIMEOUT_INVALID"), "invalid timeout was not reported");
}

static void NonJanuaryPeriodGrid()
{
    var monthNumbers = Enumerable.Range(4, 9).Concat(Enumerable.Range(1, 3)).ToArray();
    var months = monthNumbers.Select((month, index) => new MonthDto
    {
        MonthNumber = (short)month,
        MonthName = new DateTime(2000, month, 1).ToString("MMM").ToUpperInvariant(),
        StartOn = index + 1
    }).ToArray();

    var years = new[]
    {
        new ActiveYearDto { YearNumber = 2026, Description = "2026-27", CashStatus = "OPEN" },
        new ActiveYearDto { YearNumber = 2027, Description = "2027-28", CashStatus = "FORECAST" }
    };
    var active = new ActivePeriodDto
    {
        YearNumber = 2026,
        MonthNumber = 10,
        MonthName = "OCT",
        Description = "2026-27",
        StartOn = new DateTime(2026, 10, 1)
    };

    var result = CashStatementPeriodGridBuilder.Build(years, months, active, CashStatementLimits.Default);
    Require(result.IsSuccess, "non-January period grid failed");
    Equal(26, result.Value!.Periods.Length, "period count");
    Equal(StatementPeriodId.Month(2026, 4), result.Value.Periods[0].Id, "first financial month");
    Equal(StatementPeriodId.Month(2026, 3), result.Value.Periods[11].Id, "last financial month");
    Equal(StatementPeriodId.Annual(2026), result.Value.Periods[12].Id, "first annual period");
    Equal(StatementPeriodId.Month(2027, 4), result.Value.Periods[13].Id, "second-year first month");
    Equal(1, result.Value.Periods.Count(p => p.IsActive), "active-period count");
}

static void FormulaEvaluation()
{
    var evaluator = new CashStatementFormulaEvaluator();
    var divideByZero = new StatementFormulaBinary(
        StatementBinaryOperator.Divide,
        new StatementFormulaLiteral(StatementValue.FromDecimal(10m)),
        new StatementFormulaLiteral(StatementValue.FromDecimal(0m)));

    var unsupported = evaluator.Evaluate(divideByZero, _ => throw new InvalidOperationException());
    Equal(StatementValueState.Unsupported, unsupported.State, "division-by-zero state");
    Equal("division-by-zero", unsupported.Reason, "division-by-zero reason");

    var conditional = new StatementFormulaConditional(
        new StatementFormulaLiteral(StatementValue.FromBoolean(true)),
        new StatementFormulaLiteral(StatementValue.FromDecimal(42m)),
        divideByZero);
    var selected = evaluator.Evaluate(conditional, _ => throw new InvalidOperationException());
    Equal(42m, selected.DecimalValue, "lazy conditional result");

    var forwardAddress = Address("later");
    var reference = evaluator.Evaluate(
        new StatementFormulaReference(forwardAddress),
        address => address == forwardAddress ? StatementValue.FromDecimal(-12.5m) : throw new InvalidOperationException());
    Equal(-12.5m, reference.DecimalValue, "reference value");
}

static void ForwardDependencyValidation()
{
    var later = Address("later");
    var document = CreateDocument(
        Row("earlier", 0, new StatementFormulaReference(later), 9m),
        Row("later", 1, null, 9m));

    var violations = new CashStatementContractValidator().Validate(document, CashStatementLimits.Default);
    Require(violations.IsEmpty, string.Join("; ", violations.Select(v => $"{v.Code}: {v.Message}")));
}

static void DecimalPrecisionSemantics()
{
    var evaluator = new CashStatementFormulaEvaluator();
    var preciseMoney = evaluator.Evaluate(
        new StatementFormulaAggregate(
            StatementAggregateOperator.Sum,
            ImmutableArray.Create<StatementFormula>(
                new StatementFormulaLiteral(StatementValue.FromDecimal(1.005m)),
                new StatementFormulaLiteral(StatementValue.FromDecimal(2.005m)))),
        _ => throw new InvalidOperationException());
    Equal(3.010m, preciseMoney.DecimalValue, "money source precision");

    var ratio = evaluator.Evaluate(
        new StatementFormulaBinary(
            StatementBinaryOperator.Divide,
            new StatementFormulaLiteral(StatementValue.FromDecimal(1m)),
            new StatementFormulaLiteral(StatementValue.FromDecimal(8m))),
        _ => throw new InvalidOperationException());
    Equal(0.125m, ratio.DecimalValue, "dimensionless ratio precision");
    Equal(2.35m, decimal.Round(2.345m, 2, MidpointRounding.AwayFromZero), "positive explicit business rounding");
    Equal(-2.35m, decimal.Round(-2.345m, 2, MidpointRounding.AwayFromZero), "negative explicit business rounding");
}

static void InvalidDependencyValidation()
{
    var document = CreateDocument(
        Row("a", 0, new StatementFormulaReference(Address("b")), 0m),
        Row("b", 1, new StatementFormulaReference(Address("c")), 0m),
        Row("c", 2, new StatementFormulaReference(Address("a")), 0m),
        Row("broken", 3, new StatementFormulaReference(Address("missing")), 0m));

    var violations = new CashStatementContractValidator().Validate(document, CashStatementLimits.Default);
    var cycle = violations.SingleOrDefault(v => v.Code == "DEPENDENCY_CYCLE");
    Require(cycle is not null, "cycle was not reported");
    Require(cycle!.Message.Contains("/a/", StringComparison.Ordinal) &&
        cycle.Message.Contains("/b/", StringComparison.Ordinal) &&
        cycle.Message.Contains("/c/", StringComparison.Ordinal),
        $"cycle path was incomplete: {cycle.Message}");
    Require(violations.Any(v => v.Code == "DEPENDENCY_NOT_FOUND" && v.Message.Contains("missing", StringComparison.Ordinal)),
        "missing dependency was not reported");
}

static async Task FlowBuilderEvaluation()
{
    var repository = CreateFlowRepository();
    using var cts = new CancellationTokenSource();
    var request = new CashStatementRequestV1(
        IncludeActivePeriods: true,
        IncludeOrderBook: true,
        IncludeTaxAccruals: true,
        CommandTimeoutSeconds: 47);

    var result = await new CashStatementBuilder(repository).BuildAsync(
        "Server=(test);Database=(test)",
        request,
        "contract-test",
        DateTimeOffset.UnixEpoch,
        cts.Token);

    Require(result.IsSuccess, string.Join("; ", result.Violations.Select(v => $"{v.Code}: {v.Message}")));
    var document = result.Value!;
    Equal(7m, DecimalCell(document, "calculated-total:parent", "period:2026:04"), "forward total April");
    Equal(16m, DecimalCell(document, "calculated-total:parent", "period:2026:05"), "forward total May");
    Equal(23m, DecimalCell(document, "calculated-total:parent", "period:2026:annual"), "forward total annual");
    Equal(-7m, DecimalCell(document, "flow:costs:total", "period:2026:annual"), "expense polarity annual");
    Equal(30m, DecimalCell(document, "flow:income:total", "period:2026:annual"), "income annual");
    Equal(0m, DecimalCell(document, "flow:empty:total", "period:2026:04"), "empty category April");
    Equal(0m, DecimalCell(document, "flow:empty:total", "period:2026:annual"), "empty category annual");

    Equal(7m, DecimalCell(document, "summary:trade:period-total", "period:2026:04"), "trade summary April");
    var parent = FindCell(document, "calculated-total:parent", "period:2026:04");
    Require(parent.Formula is StatementFormulaAggregate, "calculated total did not retain semantic dependencies");
    Require(repository.Calls.Count > 0, "repository was not called");
    Require(repository.Calls.All(call => call.Timeout == 47), "command timeout was not propagated to every repository call");
    Require(repository.Calls.All(call => call.Token == cts.Token), "cancellation token was not propagated to every repository call");

    var tradeCalls = repository.ValueCalls.Where(call => call.CashCode is "SALES" or "COST").ToArray();
    Require(tradeCalls.All(call => call.IncludeActivePeriods), "trade active-period option was not propagated");
    Require(tradeCalls.All(call => call.IncludeOrderBook), "trade order-book option was not propagated");
    Require(tradeCalls.All(call => !call.IncludeTaxAccruals), "tax-accrual option leaked into trade values");
}

static async Task FlowBuilderMissingChild()
{
    var repository = CreateFlowRepository();
    repository.TotalChildren["PARENT"] = new[] { "DOES-NOT-EXIST" };

    var result = await new CashStatementBuilder(repository).BuildAsync(
        "Server=(test);Database=(test)",
        new CashStatementRequestV1(),
        "contract-test",
        DateTimeOffset.UnixEpoch);

    Require(!result.IsSuccess, "builder accepted a missing category child");
    Require(result.Violations.Any(v =>
            v.Code == "DEPENDENCY_NOT_FOUND" &&
            v.Message.Contains("does-not-exist", StringComparison.Ordinal)),
        "missing category child was not reported by semantic identity");
}

static async Task OptionalSectionsAreNotLoaded()
{
    var repository = CreateFlowRepository();
    var result = await new CashStatementBuilder(repository).BuildAsync(
        "Server=(test);Database=(test)",
        new CashStatementRequestV1(
            IncludeVatDetails: false,
            IncludeBankBalances: false,
            IncludeBalanceSheet: false),
        "contract-test",
        DateTimeOffset.UnixEpoch);

    Require(result.IsSuccess, string.Join("; ", result.Violations.Select(v => $"{v.Code}: {v.Message}")));
    Equal(0, repository.BankAccountCalls, "bank-account calls");
    Equal(0, repository.VatSourceCalls, "VAT source calls");
    Equal(0, repository.BalanceSheetCalls, "balance-sheet calls");
    Equal(0, repository.ReconciliationCalls, "reconciliation calls");
    Require(result.Value!.Sheets.Single().Sections.All(section =>
            section.Kind is not StatementSectionKind.BankBalances and not StatementSectionKind.BalanceSheet),
        "an excluded optional section was added to the document");
}

static async Task ExpressionBuilderBinding()
{
    var repository = CreateFlowRepository();
    repository.CategoryExpressions = new[]
    {
        new CategoryExpressionDto
        {
            DisplayOrder = 1,
            CategoryCode = "CE-GP",
            Category = "Gross profit ratio",
            Expression = "IF([INCOME]=0,0,[INCOME]/[SALES])",
            Format = "0.0%"
        },
        new CategoryExpressionDto
        {
            DisplayOrder = 2,
            CategoryCode = "CE-BAD",
            Category = "Broken ratio",
            Expression = "[Missing category]",
            Format = "0.00"
        }
    };

    var result = await new CashStatementBuilder(repository).BuildAsync(
        "Server=(test);Database=(test)",
        new CashStatementRequestV1(
            IncludeVatDetails: false,
            IncludeBankBalances: false,
            IncludeBalanceSheet: false),
        "contract-test",
        DateTimeOffset.UnixEpoch);

    Require(result.IsSuccess, string.Join("; ", result.Violations.Select(v => $"{v.Code}: {v.Message}")));
    var valid = FindCell(result.Value!, "expression:ce-gp", "period:2026:04");
    var terminal = RequireType<StatementFormulaTerminal>(valid.Formula!);
    Equal("IF([INCOME]=0,0,[INCOME]/[SALES])", terminal.Expression, "opaque expression text");
    Equal(new StatementRowId("flow:income:total"), terminal.References["INCOME"].RowId, "category binding");
    Equal(new StatementRowId("flow:income:cash-code:sales"), terminal.References["SALES"].RowId, "cash-code binding");
    Equal(StatementValueState.Unsupported, valid.Value.State, "terminal expression value state");
    Equal(StatementNumericKind.Percentage, valid.NumericKind, "expression numeric kind");
    Equal("0.0%", valid.NumberFormat, "expression number format");

    var invalid = FindCell(result.Value!, "expression:ce-bad", "period:2026:04");
    Equal(StatementValueState.Unsupported, invalid.Value.State, "invalid expression state");
    Require(invalid.Formula is null, "invalid expression retained an executable terminal formula");
    Require(invalid.Value.Reason!.Contains("does not match", StringComparison.Ordinal),
        "invalid expression did not retain its reference diagnostic");
    Require(invalid.Value.DecimalValue is null, "invalid expression was coerced to zero");
}

static async Task VatBuilderEvaluation()
{
    var repository = CreateFlowRepository();
    repository.VatRecurrenceType = "quarterly";
    repository.VatRecurrence = new[]
    {
        VatRecurrence(2026, 5, vatDueSales: 30m),
        VatRecurrence(2026, 4, vatDueSales: 10m)
    };
    repository.VatRecurrenceAccruals = new[]
    {
        VatRecurrenceAccrual(2026, 4, vatDueSales: 2m),
        VatRecurrenceAccrual(2026, 5, vatDueSales: 3m)
    };
    repository.VatPeriodTotals = new[]
    {
        VatPeriod(2026, 5, vatDueSales: 40m),
        VatPeriod(2026, 4, vatDueSales: 20m)
    };
    repository.VatPeriodAccruals = new[]
    {
        VatPeriodAccrual(2026, 4, vatDueSales: 1m),
        VatPeriodAccrual(2026, 5, vatDueSales: 4m)
    };

    var result = await new CashStatementBuilder(repository).BuildAsync(
        "Server=(test);Database=(test)",
        new CashStatementRequestV1(
            IncludeActivePeriods: false,
            IncludeTaxAccruals: true,
            IncludeBankBalances: false,
            IncludeBalanceSheet: false),
        "contract-test",
        new DateTimeOffset(2026, 4, 15, 0, 0, 0, TimeSpan.Zero));

    Require(result.IsSuccess, string.Join("; ", result.Violations.Select(v => $"{v.Code}: {v.Message}")));
    var document = result.Value!;
    Equal("VAT QUARTERLY",
        document.Sheets.Single().Sections.Single(section => section.Id.Value == "vat:recurrence").Title,
        "VAT recurrence title");
    Equal(12m, DecimalCell(document, "vat:recurrence:vat-due-sales", "period:2026:04"), "recurrence April with accrual");
    Equal(3m, DecimalCell(document, "vat:recurrence:vat-due-sales", "period:2026:05"), "future recurrence retains accrual only");
    Equal(15m, DecimalCell(document, "vat:recurrence:vat-due-sales", "period:2026:annual"), "recurrence annual");
    Equal(21m, DecimalCell(document, "vat:periods:vat-due-sales", "period:2026:04"), "VAT period April with accrual");
    Equal(4m, DecimalCell(document, "vat:periods:vat-due-sales", "period:2026:05"), "future VAT period retains accrual only");
    Equal(25m, DecimalCell(document, "vat:periods:vat-due-sales", "period:2026:annual"), "VAT period annual");
    Equal(5, repository.VatSourceCalls, "VAT repository calls");

    var withoutAccruals = CreateFlowRepository();
    var withoutAccrualsResult = await new CashStatementBuilder(withoutAccruals).BuildAsync(
        "Server=(test);Database=(test)",
        new CashStatementRequestV1(
            IncludeTaxAccruals: false,
            IncludeBankBalances: false,
            IncludeBalanceSheet: false),
        "contract-test",
        DateTimeOffset.UnixEpoch);
    Require(withoutAccrualsResult.IsSuccess,
        string.Join("; ", withoutAccrualsResult.Violations.Select(v => $"{v.Code}: {v.Message}")));
    Equal(3, withoutAccruals.VatSourceCalls, "VAT calls without accrual sources");
}

static VatRecurrenceDto VatRecurrence(short year, int month, decimal vatDueSales) => new()
{
    YearNumber = year,
    StartOn = new DateTime(year, month, 1),
    VatDueSales = vatDueSales
};

static VatRecurrenceAccrualDto VatRecurrenceAccrual(short year, int month, decimal vatDueSales) => new()
{
    YearNumber = year,
    StartOn = new DateTime(year, month, 1),
    VatDueSales = vatDueSales
};

static VatPeriodTotalDto VatPeriod(short year, int month, decimal vatDueSales) => new()
{
    YearNumber = year,
    StartOn = new DateTime(year, month, 1),
    VatDueSales = vatDueSales
};

static VatPeriodAccrualDto VatPeriodAccrual(short year, int month, decimal vatDueSales) => new()
{
    YearNumber = year,
    StartOn = new DateTime(year, month, 1),
    VatDueSales = vatDueSales
};

static async Task StatementLocalization()
{
    var repository = CreateFlowRepository();
    var result = await new CashStatementBuilder(
        repository,
        localizer: new StubCashStatementLocalizer()).BuildAsync(
        "Server=(test);Database=(test)",
        new CashStatementRequestV1(
            IncludeBankBalances: false,
            IncludeBalanceSheet: false),
        "contract-test",
        DateTimeOffset.UnixEpoch);

    Require(result.IsSuccess, string.Join("; ", result.Violations.Select(v => $"{v.Code}: {v.Message}")));
    var document = result.Value!;
    Equal("[[StatementTitle]]:APR|2026-27", document.Metadata.Title, "localized statement title");
    Equal("[[CashFlow]]", document.Sheets.Single().Name, "localized sheet name");
    Equal("[[Code]]", document.Sheets.Single().Columns[0].Label, "localized code column");
    Equal("[[TradeCashFlow]]", document.Sheets.Single().Sections[0].Title, "localized flow section");
    Equal("Costs", document.Sheets.Single().Sections[0].Rows[0].Label, "database category description");
}

static async Task BankBalanceBuilderEvaluation()
{
    var repository = CreateFlowRepository();
    repository.Months = new[]
    {
        new MonthDto { MonthNumber = 4, MonthName = "APR", StartOn = 1m },
        new MonthDto { MonthNumber = 12, MonthName = "DEC", StartOn = 2m }
    };
    repository.BankAccounts = new[]
    {
        new BankAccountDto { AccountCode = "CUR", AccountName = "Current Account" },
        new BankAccountDto { AccountCode = "SAV", AccountName = "Savings Account" }
    };
    repository.BankBalances["CUR"] = new[]
    {
        new BankBalanceDto { AccountCode = "CUR", YearNumber = 2026, MonthNumber = 4, Balance = 10m },
        new BankBalanceDto { AccountCode = "CUR", YearNumber = 2026, MonthNumber = 12, Balance = 20m }
    };
    repository.BankBalances["SAV"] = new[]
    {
        new BankBalanceDto { AccountCode = "SAV", YearNumber = 2026, MonthNumber = 12, Balance = -3m }
    };

    var result = await new CashStatementBuilder(repository).BuildAsync(
        "Server=(test);Database=(test)",
        new CashStatementRequestV1(IncludeBalanceSheet: false),
        "contract-test",
        DateTimeOffset.UnixEpoch);

    Require(result.IsSuccess, string.Join("; ", result.Violations.Select(v => $"{v.Code}: {v.Message}")));
    var document = result.Value!;
    Equal(10m, DecimalCell(document, "bank:account:cur", "period:2026:04"), "current account April");
    Equal(0m, DecimalCell(document, "bank:account:sav", "period:2026:04"), "missing bank month is explicit zero");
    Equal(20m, DecimalCell(document, "bank:account:cur", "period:2026:annual"), "current account annual closing balance");
    Equal(-3m, DecimalCell(document, "bank:account:sav", "period:2026:annual"), "savings annual closing balance");
    Equal(10m, DecimalCell(document, "bank:company-balance", "period:2026:04"), "company April balance");
    Equal(17m, DecimalCell(document, "bank:company-balance", "period:2026:annual"), "company annual closing balance");
    Equal(1, repository.BankAccountCalls, "bank-account calls");
    Equal(2, repository.BankBalanceCalls, "bank-balance calls");
}

static async Task BalanceSheetBuilderEvaluation()
{
    var repository = CreateFlowRepository();
    repository.BalanceSheetEntries = new[]
    {
        new BalanceSheetEntryDto { AssetCode = "BANK", AssetName = "Cash at bank", YearNumber = 2026, MonthNumber = 4, Balance = 10m },
        new BalanceSheetEntryDto { AssetCode = "BANK", AssetName = "Cash at bank", YearNumber = 2026, MonthNumber = 5, Balance = 12m },
        new BalanceSheetEntryDto { AssetCode = "BANK", AssetName = "Cash at bank", YearNumber = 2026, MonthNumber = 5, Balance = 3m },
        new BalanceSheetEntryDto { AssetCode = "LIAB", AssetName = "Liabilities", YearNumber = 2026, MonthNumber = 4, Balance = -4m }
    };

    var result = await new CashStatementBuilder(repository).BuildAsync(
        "Server=(test);Database=(test)",
        new CashStatementRequestV1(IncludeBankBalances: false),
        "contract-test",
        DateTimeOffset.UnixEpoch);

    Require(result.IsSuccess, string.Join("; ", result.Violations.Select(v => $"{v.Code}: {v.Message}")));
    var document = result.Value!;
    Equal(15m, DecimalCell(document, "balance-sheet:asset:bank:cash-at-bank", "period:2026:05"), "duplicate asset values are summed");
    Equal(15m, DecimalCell(document, "balance-sheet:asset:bank:cash-at-bank", "period:2026:annual"), "asset annual carries last reported month");
    Equal(-4m, DecimalCell(document, "balance-sheet:asset:liab:liabilities", "period:2026:annual"), "earlier asset balance carries into annual");
    Equal(StatementValueState.Missing,
        FindCell(document, "balance-sheet:asset:liab:liabilities", "period:2026:05").Value.State,
        "unreported balance-sheet month remains missing");
    Equal(6m, DecimalCell(document, "balance-sheet:capital", "period:2026:04"), "capital April");
    Equal(15m, DecimalCell(document, "balance-sheet:capital", "period:2026:05"), "capital ignores missing source cells like spreadsheet SUM");
    Equal(11m, DecimalCell(document, "balance-sheet:capital", "period:2026:annual"), "capital annual");
    Equal(1, repository.BalanceSheetCalls, "balance-sheet calls");
}

static async Task ReconciliationBuilderEvaluation()
{
    var repository = CreateFlowRepository();
    repository.ReconciliationRows = new[]
    {
        Reconciliation(2024, 10.10m, 0.10m),
        Reconciliation(2025, 10.11m, 0.11m),
        Reconciliation(2026, 11.01m, 1.01m)
    };

    var result = await new CashStatementBuilder(repository).BuildAsync(
        "Server=(test);Database=(test)",
        new CashStatementRequestV1(
            IncludeVatDetails: false,
            IncludeBankBalances: false,
            IncludeBalanceSheet: false,
            IncludeReconciliation: true),
        "contract-test",
        DateTimeOffset.UnixEpoch);

    Require(result.IsSuccess, string.Join("; ", result.Violations.Select(v => $"{v.Code}: {v.Message}")));
    var document = result.Value!;
    Equal(2, document.Sheets.Length, "sheet count with reconciliation");
    Equal(StatementSheetKind.Reconciliation, document.Sheets[1].Kind, "reconciliation sheet kind");
    Equal(10m, DecimalCell(document, "reconciliation:2024", "bridge-total"), "reconciliation bridge total");
    Equal(0.10m, DecimalCell(document, "reconciliation:2024", "difference"), "PASS boundary difference");
    Equal("PASS", TextCell(document, "reconciliation:2024", "status"), "PASS boundary status");
    Equal("WARN", TextCell(document, "reconciliation:2025", "status"), "WARN status");
    Equal("FAIL", TextCell(document, "reconciliation:2026", "status"), "FAIL status");
    Equal("FAIL", TextCell(document, "reconciliation:summary", "status"), "overall status");
    Require(FindCell(document, "reconciliation:summary", "status").Formula is StatementFormulaConditional,
        "overall reconciliation status is not a shared-model formula");
    Equal(1, repository.ReconciliationCalls, "reconciliation calls");
}

static EquityReconciliationByYearDto Reconciliation(short year, decimal capitalDelta, decimal variance) => new()
{
    YearNumber = year,
    Description = year.ToString(),
    ProfitAfterTax = 5m,
    CapitalMovement = 2m,
    OpeningAccountPosition = 3m,
    CapitalDelta = capitalDelta,
    Variance = variance
};

static StubCashFlowRepository CreateFlowRepository()
{
    var repository = new StubCashFlowRepository
    {
        ActivePeriod = new ActivePeriodDto
        {
            YearNumber = 2026,
            MonthNumber = 4,
            MonthName = "APR",
            Description = "2026-27",
            StartOn = new DateTime(2026, 4, 1)
        },
        Years = new[]
        {
            new ActiveYearDto { YearNumber = 2026, Description = "2026-27", CashStatus = "OPEN" }
        },
        Months = new[]
        {
            new MonthDto { MonthNumber = 4, MonthName = "APR", StartOn = 1m },
            new MonthDto { MonthNumber = 5, MonthName = "MAY", StartOn = 2m }
        }
    };

    repository.Categories[CashType.Trade] = new[]
    {
        new FlowCategoryDto { CategoryCode = "COSTS", Category = "Costs", CashPolarityCode = 0, DisplayOrder = 1 },
        new FlowCategoryDto { CategoryCode = "INCOME", Category = "Income", CashPolarityCode = 1, DisplayOrder = 2 },
        new FlowCategoryDto { CategoryCode = "EMPTY", Category = "Empty", CashPolarityCode = 2, DisplayOrder = 3 }
    };
    repository.CashCodes["COSTS"] = new[]
    {
        new FlowCategoryCashCodeDto { CashCode = "COST", CashDescription = "Costs" }
    };
    repository.CashCodes["INCOME"] = new[]
    {
        new FlowCategoryCashCodeDto { CashCode = "SALES", CashDescription = "Sales" }
    };
    repository.Values[("COST", (short)2026)] = new[]
    {
        new CashCodePeriodValue { StartOn = new DateTime(2026, 4, 1), InvoiceValue = 3m },
        new CashCodePeriodValue { StartOn = new DateTime(2026, 5, 1), InvoiceValue = 4m }
    };
    repository.Values[("SALES", (short)2026)] = new[]
    {
        new CashCodePeriodValue { StartOn = new DateTime(2026, 4, 1), InvoiceValue = 10m },
        new CashCodePeriodValue { StartOn = new DateTime(2026, 5, 1), InvoiceValue = 20m }
    };
    repository.TotalCategories = new[]
    {
        new CategoryTotalDto { CategoryCode = "PARENT" },
        new CategoryTotalDto { CategoryCode = "SECOND" }
    };
    repository.TotalChildren["PARENT"] = new[] { "SECOND" };
    repository.TotalChildren["SECOND"] = new[] { "INCOME", "COSTS" };
    repository.CalculatedTotals[(CashType.Trade, CategoryType.Total)] = new[]
    {
        new FlowCategoriesByTypeDto { CategoryCode = "PARENT", Category = "Parent", CashType = "TRADE" },
        new FlowCategoriesByTypeDto { CategoryCode = "SECOND", Category = "Second", CashType = "TRADE" }
    };
    return repository;
}

static decimal DecimalCell(CashStatementDocument document, string rowId, string columnId) =>
    FindCell(document, rowId, columnId).Value.DecimalValue
    ?? throw new InvalidOperationException($"{rowId}/{columnId} is not a present decimal.");

static string TextCell(CashStatementDocument document, string rowId, string columnId) =>
    FindCell(document, rowId, columnId).Value.TextValue
    ?? throw new InvalidOperationException($"{rowId}/{columnId} is not present text.");

static StatementCell FindCell(CashStatementDocument document, string rowId, string columnId) =>
    document.Sheets
        .SelectMany(sheet => sheet.Sections)
        .SelectMany(section => section.Rows)
        .Single(row => row.Id.Value == rowId)
        .Cells[new StatementColumnId(columnId)];

static StatementCellAddress Address(string rowId) => new(
    new StatementSheetId("cash-flow"),
    new StatementRowId(rowId),
    new StatementColumnId("value"));

static StatementRow Row(string id, int order, StatementFormula? formula, decimal value) => new(
    new StatementRowId(id),
    StatementRowKind.Detail,
    order,
    id,
    id,
    StatementPresentationRole.Detail,
    ImmutableDictionary<StatementColumnId, StatementCell>.Empty.Add(
        new StatementColumnId("value"),
        new StatementCell(
            StatementValue.FromDecimal(value),
            StatementNumericKind.Money,
            StatementPresentationRole.Detail,
            formula)));

static CashStatementDocument CreateDocument(params StatementRow[] rows) => new(
    new CashStatementMetadata(
        "Example",
        "Cash Statement",
        "owner",
        DateTimeOffset.UnixEpoch,
        StatementPeriodId.Month(2026, 1)),
    new CashStatementRequestV1(),
    new StatementPeriodGrid(ImmutableArray<StatementPeriod>.Empty),
    ImmutableArray.Create(new StatementSheet(
        new StatementSheetId("cash-flow"),
        StatementSheetKind.CashFlow,
        0,
        "Cash Flow",
        ImmutableArray.Create(new StatementColumn(
            new StatementColumnId("value"),
            "Value",
            StatementColumnKind.Result,
            0)),
        ImmutableArray.Create(new StatementSection(
            new StatementSectionId("flow:test"),
            StatementSectionKind.FlowCategories,
            0,
            "Test",
            StatementPresentationRole.SectionHeader,
            rows.ToImmutableArray())))));

static T RequireType<T>(object value)
{
    if (value is T typed)
        return typed;

    throw new InvalidOperationException($"Expected {typeof(T).Name}, found {value.GetType().Name}.");
}

static void Equal<T>(T expected, T actual, string name)
{
    if (!EqualityComparer<T>.Default.Equals(expected, actual))
        throw new InvalidOperationException($"{name}: expected '{expected}', found '{actual}'.");
}

static void Require(bool condition, string message)
{
    if (!condition)
        throw new InvalidOperationException(message);
}

sealed class StubCashFlowRepository : ICashFlowRepository
{
    public CancellationTokenSource? CancelAfterFirstCall { get; init; }
    public ActivePeriodDto ActivePeriod { get; init; } = new();
    public IReadOnlyList<ActiveYearDto> Years { get; init; } = Array.Empty<ActiveYearDto>();
    public IReadOnlyList<MonthDto> Months { get; set; } = Array.Empty<MonthDto>();
    public Dictionary<CashType, IReadOnlyList<FlowCategoryDto>> Categories { get; } = new();
    public Dictionary<string, IReadOnlyList<FlowCategoryCashCodeDto>> CashCodes { get; } =
        new(StringComparer.OrdinalIgnoreCase);
    public Dictionary<(CashType CashType, CategoryType CategoryType), IReadOnlyList<FlowCategoriesByTypeDto>> CalculatedTotals { get; } = new();
    public Dictionary<(string CashCode, short Year), IReadOnlyList<CashCodePeriodValue>> Values { get; } = new();
    public IReadOnlyList<CategoryTotalDto> TotalCategories { get; set; } = Array.Empty<CategoryTotalDto>();
    public Dictionary<string, IReadOnlyList<string>> TotalChildren { get; } =
        new(StringComparer.OrdinalIgnoreCase);
    public IReadOnlyList<BankAccountDto> BankAccounts { get; set; } = Array.Empty<BankAccountDto>();
    public Dictionary<string, IReadOnlyList<BankBalanceDto>> BankBalances { get; } =
        new(StringComparer.OrdinalIgnoreCase);
    public string VatRecurrenceType { get; set; } = string.Empty;
    public IReadOnlyList<VatRecurrenceDto> VatRecurrence { get; set; } = Array.Empty<VatRecurrenceDto>();
    public IReadOnlyList<VatRecurrenceAccrualDto> VatRecurrenceAccruals { get; set; } = Array.Empty<VatRecurrenceAccrualDto>();
    public IReadOnlyList<VatPeriodTotalDto> VatPeriodTotals { get; set; } = Array.Empty<VatPeriodTotalDto>();
    public IReadOnlyList<VatPeriodAccrualDto> VatPeriodAccruals { get; set; } = Array.Empty<VatPeriodAccrualDto>();
    public IReadOnlyList<BalanceSheetEntryDto> BalanceSheetEntries { get; set; } = Array.Empty<BalanceSheetEntryDto>();
    public IReadOnlyList<CategoryExpressionDto> CategoryExpressions { get; set; } = Array.Empty<CategoryExpressionDto>();
    public IReadOnlyList<EquityReconciliationByYearDto> ReconciliationRows { get; set; } = Array.Empty<EquityReconciliationByYearDto>();
    public int BankAccountCalls { get; private set; }
    public int BankBalanceCalls { get; private set; }
    public int VatSourceCalls { get; private set; }
    public int BalanceSheetCalls { get; private set; }
    public int ReconciliationCalls { get; private set; }
    public List<(int Timeout, CancellationToken Token)> Calls { get; } = new();
    public List<(string CashCode, bool IncludeActivePeriods, bool IncludeOrderBook, bool IncludeTaxAccruals)> ValueCalls { get; } = new();

    public Task<IReadOnlyList<CashCodePeriodValue>> GetCashCodeValuesAsync(
        string connectionString,
        string cashCode,
        short yearNumber,
        bool includeActivePeriods,
        bool includeOrderBook,
        bool includeTaxAccruals,
        int commandTimeoutSeconds = 30,
        CancellationToken ct = default)
    {
        Track(commandTimeoutSeconds, ct);
        ValueCalls.Add((cashCode, includeActivePeriods, includeOrderBook, includeTaxAccruals));
        Values.TryGetValue((cashCode, yearNumber), out var values);
        return Task.FromResult(values ?? (IReadOnlyList<CashCodePeriodValue>)Array.Empty<CashCodePeriodValue>());
    }

    public Task<string> GetCompanyNameAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        Track(commandTimeoutSeconds, ct);
        return Task.FromResult("Example Company");
    }

    public Task<ActivePeriodDto> GetActivePeriodAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        Track(commandTimeoutSeconds, ct);
        return Task.FromResult(ActivePeriod);
    }

    public Task<IReadOnlyList<ActiveYearDto>> GetActiveYearsAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        Track(commandTimeoutSeconds, ct);
        return Task.FromResult(Years);
    }

    public Task<IReadOnlyList<MonthDto>> GetMonthsAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        Track(commandTimeoutSeconds, ct);
        return Task.FromResult(Months);
    }

    public Task<IReadOnlyList<FlowCategoryDto>> GetCategoriesAsync(string connectionString, CashType cashType, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        Track(commandTimeoutSeconds, ct);
        Categories.TryGetValue(cashType, out var categories);
        return Task.FromResult(categories ?? (IReadOnlyList<FlowCategoryDto>)Array.Empty<FlowCategoryDto>());
    }

    public Task<IReadOnlyList<FlowCategoryCashCodeDto>> GetCashCodesAsync(string connectionString, string categoryCode, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        Track(commandTimeoutSeconds, ct);
        CashCodes.TryGetValue(categoryCode, out var cashCodes);
        return Task.FromResult(cashCodes ?? (IReadOnlyList<FlowCategoryCashCodeDto>)Array.Empty<FlowCategoryCashCodeDto>());
    }

    public Task<IReadOnlyList<FlowCategoriesByTypeDto>> GetCategoriesByTypeAsync(string connectionString, CashType cashType, CategoryType categoryType, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        Track(commandTimeoutSeconds, ct);
        CalculatedTotals.TryGetValue((cashType, categoryType), out var totals);
        return Task.FromResult(totals ?? (IReadOnlyList<FlowCategoriesByTypeDto>)Array.Empty<FlowCategoriesByTypeDto>());
    }

    public Task<IReadOnlyList<CategoryTotalDto>> GetCategoryTotalsAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        Track(commandTimeoutSeconds, ct);
        return Task.FromResult(TotalCategories);
    }

    public Task<IReadOnlyList<string>> GetCategoryTotalCodesAsync(string connectionString, string categoryCode, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        Track(commandTimeoutSeconds, ct);
        TotalChildren.TryGetValue(categoryCode, out var children);
        return Task.FromResult(children ?? (IReadOnlyList<string>)Array.Empty<string>());
    }

    public Task<IReadOnlyList<CategoryExpressionDto>> GetCategoryExpressionsAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        Track(commandTimeoutSeconds, ct);
        return Task.FromResult(CategoryExpressions);
    }

    public Task<string> GetCategoryCodeFromNameAsync(string connectionString, string categoryName, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        Track(commandTimeoutSeconds, ct);
        return Task.FromResult(string.Empty);
    }

    public Task SetCategoryExpressionStatusAsync(string connectionString, string categoryCode, bool isError, string? errorMessage = null, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        Track(commandTimeoutSeconds, ct);
        return Task.CompletedTask;
    }

    public Task<IReadOnlyList<BankAccountDto>> GetBankAccountsAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        Track(commandTimeoutSeconds, ct);
        BankAccountCalls++;
        return Task.FromResult(BankAccounts);
    }

    public Task<IReadOnlyList<BankBalanceDto>> GetBankBalancesAsync(string connectionString, string accountCode, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        Track(commandTimeoutSeconds, ct);
        BankBalanceCalls++;
        BankBalances.TryGetValue(accountCode, out var balances);
        return Task.FromResult(balances ?? (IReadOnlyList<BankBalanceDto>)Array.Empty<BankBalanceDto>());
    }

    public Task<string> GetVatRecurrenceTypeAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        Track(commandTimeoutSeconds, ct);
        VatSourceCalls++;
        return Task.FromResult(VatRecurrenceType);
    }

    public Task<IReadOnlyList<VatRecurrenceDto>> GetVatRecurrenceAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        Track(commandTimeoutSeconds, ct);
        VatSourceCalls++;
        return Task.FromResult(VatRecurrence);
    }

    public Task<IReadOnlyList<VatRecurrenceAccrualDto>> GetVatRecurrenceAccrualsAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        Track(commandTimeoutSeconds, ct);
        VatSourceCalls++;
        return Task.FromResult(VatRecurrenceAccruals);
    }

    public Task<IReadOnlyList<VatPeriodTotalDto>> GetVatPeriodTotalsAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        Track(commandTimeoutSeconds, ct);
        VatSourceCalls++;
        return Task.FromResult(VatPeriodTotals);
    }

    public Task<IReadOnlyList<VatPeriodAccrualDto>> GetVatPeriodAccrualsAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        Track(commandTimeoutSeconds, ct);
        VatSourceCalls++;
        return Task.FromResult(VatPeriodAccruals);
    }

    public Task<IReadOnlyList<BalanceSheetEntryDto>> GetBalanceSheetAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        Track(commandTimeoutSeconds, ct);
        BalanceSheetCalls++;
        return Task.FromResult(BalanceSheetEntries);
    }

    public Task<IReadOnlyList<EquityReconciliationByYearDto>> GetEquityReconciliationByYearAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        Track(commandTimeoutSeconds, ct);
        ReconciliationCalls++;
        return Task.FromResult(ReconciliationRows);
    }

    private Task<IReadOnlyList<T>> Empty<T>(int timeout, CancellationToken ct)
    {
        Track(timeout, ct);
        return Task.FromResult<IReadOnlyList<T>>(Array.Empty<T>());
    }

    private void Track(int timeout, CancellationToken ct)
    {
        ct.ThrowIfCancellationRequested();
        Calls.Add((timeout, ct));
        if (Calls.Count == 1)
            CancelAfterFirstCall?.Cancel();
    }
}

sealed class CancellingLayoutAllocator(
    ICashStatementLayoutAllocator inner,
    CancellationTokenSource cancellation) : ICashStatementLayoutAllocator
{
    public StatementLayout Allocate(CashStatementDocument document, CashStatementLimits limits)
    {
        var layout = inner.Allocate(document, limits);
        cancellation.Cancel();
        return layout;
    }
}

sealed record ContractTestResult(
    string Name,
    bool Passed,
    string? Error,
    double DurationMilliseconds);

sealed record ContractTestEvidence(
    DateTimeOffset GeneratedAtUtc,
    string RuntimeVersion,
    int Total,
    int Passed,
    int Failed,
    IReadOnlyList<ContractTestResult> Tests);

sealed class StubCashStatementLocalizer : ICashStatementLocalizer
{
    public string Get(CashStatementText text) => $"[[{text}]]";

    public string Format(CashStatementText text, params object[] arguments) =>
        $"{Get(text)}:{string.Join('|', arguments)}";
}
