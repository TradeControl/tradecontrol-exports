using Microsoft.Data.SqlClient;
using System.Data;
using TCExports.Generator.Contracts;

namespace TCExports.Generator.Data;

public sealed partial class SqlServerCashFlowRepository : ICashFlowRepository
{
    private static SqlCommand StoredProcedure(
        SqlConnection connection,
        string procedure,
        int commandTimeoutSeconds) => new(procedure, connection)
    {
        CommandType = CommandType.StoredProcedure,
        CommandTimeout = commandTimeoutSeconds
    };

    private static SqlCommand Query(
        SqlConnection connection,
        string sql,
        int commandTimeoutSeconds) => new(sql, connection)
    {
        CommandType = CommandType.Text,
        CommandTimeout = commandTimeoutSeconds
    };

    private static void AddVersionParameter(SqlCommand command, string categoryVersionId) =>
        command.Parameters.Add(new SqlParameter("@CategoryVersionId", SqlDbType.NVarChar, 10)
        {
            Value = categoryVersionId
        });


    public async Task<CategoryTreeVersionResolutionDto> ResolveCategoryTreeVersionAsync(
        string connectionString,
        string? categoryVersionId,
        int commandTimeoutSeconds = 30,
        CancellationToken ct = default)
    {
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var connection = new SqlConnection(ado);
        await connection.OpenAsync(ct);
        await using var command = Query(
            connection,
            "SELECT * FROM Cash.fnCategoryVersionResolve(@CategoryVersionId);",
            commandTimeoutSeconds);
        command.Parameters.Add(new SqlParameter("@CategoryVersionId", SqlDbType.NVarChar, 10)
        {
            Value = string.IsNullOrWhiteSpace(categoryVersionId) ? DBNull.Value : categoryVersionId.Trim()
        });
        await using var reader = await command.ExecuteReaderAsync(CommandBehavior.SingleRow, ct);
        if (!await reader.ReadAsync(ct))
            throw new InvalidOperationException("The active Category Tree version could not be resolved.");

        var active = ReadVersion(reader, "Active");
        var selected = reader.IsDBNull(reader.GetOrdinal("SelectedVersionId")) ? null : ReadVersion(reader, "Selected");
        return new(active, selected);

        static CategoryTreeVersionDto ReadVersion(SqlDataReader reader, string prefix) => new(
            reader.GetString(reader.GetOrdinal($"{prefix}VersionId")),
            reader.GetString(reader.GetOrdinal($"{prefix}Description")),
            reader.GetInt16(reader.GetOrdinal($"{prefix}StatusCode")),
            reader.GetString(reader.GetOrdinal($"{prefix}Status")),
            reader.IsDBNull(reader.GetOrdinal($"{prefix}EffectiveFromOn")) ? null : reader.GetDateTime(reader.GetOrdinal($"{prefix}EffectiveFromOn")),
            reader.IsDBNull(reader.GetOrdinal($"{prefix}EffectiveToOn")) ? null : reader.GetDateTime(reader.GetOrdinal($"{prefix}EffectiveToOn")),
            reader.GetDateTime(reader.GetOrdinal($"{prefix}ContentChangedOn")),
            reader.IsDBNull(reader.GetOrdinal($"{prefix}CanonicalSha256")) ? null : reader.GetString(reader.GetOrdinal($"{prefix}CanonicalSha256")),
            reader.IsDBNull(reader.GetOrdinal($"{prefix}CanonicalByteLength")) ? null : reader.GetInt64(reader.GetOrdinal($"{prefix}CanonicalByteLength")));
    }

    public async Task<IReadOnlyList<CategoryTreeVersionIntervalDto>> GetCategoryTreeVersionTimelineAsync(
        string connectionString,
        int commandTimeoutSeconds = 30,
        CancellationToken ct = default)
    {
        var list = new List<CategoryTreeVersionIntervalDto>();
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var connection = new SqlConnection(ado);
        await connection.OpenAsync(ct);
        await using var command = Query(
            connection,
            "SELECT VersionId, EffectiveFromOn, EffectiveToOn FROM Cash.fnCategoryVersionTimeline() ORDER BY EffectiveFromOn, VersionId;",
            commandTimeoutSeconds);
        await using var reader = await command.ExecuteReaderAsync(ct);
        while (await reader.ReadAsync(ct))
        {
            list.Add(new(
                reader.GetString(reader.GetOrdinal("VersionId")),
                DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("EffectiveFromOn"))),
                reader.IsDBNull(reader.GetOrdinal("EffectiveToOn"))
                    ? null
                    : DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("EffectiveToOn")))));
        }
        return list;
    }

    public async Task<IReadOnlyList<FlowCategoryDto>> GetCategoriesAsync(
        string connectionString,
        string categoryVersionId,
        CashType cashType,
        int commandTimeoutSeconds = 30,
        CancellationToken ct = default)
    {
        var list = new List<FlowCategoryDto>();
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var connection = new SqlConnection(ado);
        await connection.OpenAsync(ct);
        await using var command = Query(
            connection,
            "SELECT CategoryCode, Category, CashPolarityCode, DisplayOrder FROM Cash.fnFlowCategoryByVersion(@CategoryVersionId, @CashTypeCode) ORDER BY DisplayOrder, Category;",
            commandTimeoutSeconds);
        AddVersionParameter(command, categoryVersionId);
        command.Parameters.Add(new SqlParameter("@CashTypeCode", SqlDbType.SmallInt) { Value = (short)cashType });
        await using var reader = await command.ExecuteReaderAsync(ct);
        while (await reader.ReadAsync(ct))
        {
            list.Add(new FlowCategoryDto
            {
                CategoryCode = reader.GetString(reader.GetOrdinal("CategoryCode")),
                Category = reader.GetString(reader.GetOrdinal("Category")),
                CashPolarityCode = reader.IsDBNull(reader.GetOrdinal("CashPolarityCode")) ? null : reader.GetInt16(reader.GetOrdinal("CashPolarityCode")),
                DisplayOrder = reader.IsDBNull(reader.GetOrdinal("DisplayOrder")) ? null : reader.GetInt16(reader.GetOrdinal("DisplayOrder"))
            });
        }
        return list;
    }

    public async Task<IReadOnlyList<FlowCategoryCashCodeDto>> GetCashCodesAsync(
        string connectionString,
        string categoryVersionId,
        string categoryCode,
        int commandTimeoutSeconds = 30,
        CancellationToken ct = default)
    {
        var list = new List<FlowCategoryCashCodeDto>();
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var connection = new SqlConnection(ado);
        await connection.OpenAsync(ct);
        await using var command = Query(
            connection,
            "SELECT CashCode, CashDescription FROM Cash.fnFlowCategoryCashCodesByVersion(@CategoryVersionId, @CategoryCode) ORDER BY CashDescription;",
            commandTimeoutSeconds);
        AddVersionParameter(command, categoryVersionId);
        command.Parameters.Add(new SqlParameter("@CategoryCode", SqlDbType.NVarChar, 10) { Value = categoryCode });
        await using var reader = await command.ExecuteReaderAsync(ct);
        while (await reader.ReadAsync(ct))
            list.Add(new FlowCategoryCashCodeDto
            {
                CashCode = reader.GetString(reader.GetOrdinal("CashCode")),
                CashDescription = reader.GetString(reader.GetOrdinal("CashDescription"))
            });
        return list;
    }

    public async Task<IReadOnlyList<FlowCategoriesByTypeDto>> GetCategoriesByTypeAsync(
        string connectionString,
        string categoryVersionId,
        CashType cashType,
        CategoryType categoryType,
        int commandTimeoutSeconds = 30,
        CancellationToken ct = default)
    {
        var list = new List<FlowCategoriesByTypeDto>();
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var connection = new SqlConnection(ado);
        await connection.OpenAsync(ct);
        await using var command = Query(
            connection,
            "SELECT CategoryCode, Category, CashType FROM Cash.fnFlowCategoriesByTypeByVersion(@CategoryVersionId, @CashTypeCode, @CategoryTypeCode) ORDER BY DisplayOrder, Category;",
            commandTimeoutSeconds);
        AddVersionParameter(command, categoryVersionId);
        command.Parameters.Add(new SqlParameter("@CashTypeCode", SqlDbType.SmallInt) { Value = (short)cashType });
        command.Parameters.Add(new SqlParameter("@CategoryTypeCode", SqlDbType.SmallInt) { Value = (short)categoryType });
        await using var reader = await command.ExecuteReaderAsync(ct);
        while (await reader.ReadAsync(ct))
            list.Add(new FlowCategoriesByTypeDto
            {
                CategoryCode = reader.GetString(reader.GetOrdinal("CategoryCode")),
                Category = reader.GetString(reader.GetOrdinal("Category")),
                CashType = reader.IsDBNull(reader.GetOrdinal("CashType")) ? string.Empty : reader.GetString(reader.GetOrdinal("CashType"))
            });
        return list;
    }

    public async Task<IReadOnlyList<CategoryTotalDto>> GetCategoryTotalsAsync(
        string connectionString,
        string categoryVersionId,
        int commandTimeoutSeconds = 30,
        CancellationToken ct = default)
    {
        var list = new List<CategoryTotalDto>();
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var connection = new SqlConnection(ado);
        await connection.OpenAsync(ct);
        await using var command = Query(
            connection,
            "SELECT CategoryCode FROM Cash.fnFlowCategoryTotalsByVersion(@CategoryVersionId) ORDER BY DisplayOrder, Category;",
            commandTimeoutSeconds);
        AddVersionParameter(command, categoryVersionId);
        await using var reader = await command.ExecuteReaderAsync(ct);
        while (await reader.ReadAsync(ct))
            list.Add(new CategoryTotalDto { CategoryCode = reader.GetString(0) });
        return list;
    }

    public async Task<IReadOnlyList<string>> GetCategoryTotalCodesAsync(
        string connectionString,
        string categoryVersionId,
        string categoryCode,
        int commandTimeoutSeconds = 30,
        CancellationToken ct = default)
    {
        var list = new List<string>();
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var connection = new SqlConnection(ado);
        await connection.OpenAsync(ct);
        await using var command = Query(
            connection,
            "SELECT CategoryCode FROM Cash.fnFlowCategoryTotalCodesByVersion(@CategoryVersionId, @CategoryCode) ORDER BY DisplayOrder, CategoryCode;",
            commandTimeoutSeconds);
        AddVersionParameter(command, categoryVersionId);
        command.Parameters.Add(new SqlParameter("@CategoryCode", SqlDbType.NVarChar, 10) { Value = categoryCode });
        await using var reader = await command.ExecuteReaderAsync(ct);
        while (await reader.ReadAsync(ct))
            list.Add(reader.GetString(0));
        return list;
    }

    public async Task<IReadOnlyList<CategoryExpressionDto>> GetCategoryExpressionsAsync(
        string connectionString,
        string categoryVersionId,
        int commandTimeoutSeconds = 30,
        CancellationToken ct = default)
    {
        var list = new List<CategoryExpressionDto>();
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var connection = new SqlConnection(ado);
        await connection.OpenAsync(ct);
        await using var command = Query(
            connection,
            "SELECT DisplayOrder, CategoryCode, Category, Expression, Format FROM Cash.fnCategoryExpressionsByVersion(@CategoryVersionId) ORDER BY DisplayOrder, Category;",
            commandTimeoutSeconds);
        AddVersionParameter(command, categoryVersionId);
        await using var reader = await command.ExecuteReaderAsync(ct);
        while (await reader.ReadAsync(ct))
            list.Add(new CategoryExpressionDto
            {
                DisplayOrder = reader.GetInt16(reader.GetOrdinal("DisplayOrder")),
                CategoryCode = reader.GetString(reader.GetOrdinal("CategoryCode")),
                Category = reader.GetString(reader.GetOrdinal("Category")),
                Expression = reader.GetString(reader.GetOrdinal("Expression")),
                Format = reader.GetString(reader.GetOrdinal("Format"))
            });
        return list;
    }

    public async Task<string> GetCategoryCodeFromNameAsync(
        string connectionString,
        string categoryVersionId,
        string categoryName,
        int commandTimeoutSeconds = 30,
        CancellationToken ct = default)
    {
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var connection = new SqlConnection(ado);
        await connection.OpenAsync(ct);
        await using var command = Query(
            connection,
            "SELECT Cash.fnCategoryCodeResolveByVersion(@CategoryVersionId, @Token);",
            commandTimeoutSeconds);
        AddVersionParameter(command, categoryVersionId);
        command.Parameters.Add(new SqlParameter("@Token", SqlDbType.NVarChar, 100) { Value = categoryName.Trim() });
        return await command.ExecuteScalarAsync(ct) as string ?? string.Empty;
    }

    public async Task<string> GetCompanyNameAsync(
        string connectionString,
        int commandTimeoutSeconds = 30,
        CancellationToken ct = default)
    {
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var connection = new SqlConnection(ado);
        await connection.OpenAsync(ct);
        await using var command = Query(
            connection,
            "SELECT TOP (1) SubjectName FROM App.vwHomeAccount;",
            commandTimeoutSeconds);
        return await command.ExecuteScalarAsync(ct) as string ?? string.Empty;
    }

    public async Task SetCategoryExpressionStatusAsync(
        string connectionString,
        string categoryCode,
        bool isError,
        string? errorMessage = null,
        int commandTimeoutSeconds = 30,
        CancellationToken ct = default)
    {
        const string proc = "Cash.proc_CategoryExprStatusSet";
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var conn = new SqlConnection(ado);
        await conn.OpenAsync(ct);

        await using (var stateCommand = Query(
            conn,
            "SELECT CONVERT(nvarchar(20), DATABASEPROPERTYEX(DB_NAME(), 'Updateability'));",
            commandTimeoutSeconds))
        {
            var updateability = await stateCommand.ExecuteScalarAsync(ct) as string;
            if (string.Equals(updateability, "READ_ONLY", StringComparison.OrdinalIgnoreCase))
                return;
        }

        await using var cmd = new SqlCommand(proc, conn)
        {
            CommandType = CommandType.StoredProcedure,
            CommandTimeout = commandTimeoutSeconds
        };

        cmd.Parameters.Add(new SqlParameter("@CategoryCode", SqlDbType.NVarChar, 10) { Value = categoryCode });
        cmd.Parameters.Add(new SqlParameter("@IsError", SqlDbType.Bit) { Value = isError ? 1 : 0 });
        cmd.Parameters.Add(new SqlParameter("@ErrorMessage", SqlDbType.NVarChar, -1)
        {
            Value = errorMessage is null ? DBNull.Value : errorMessage
        });

        await cmd.ExecuteNonQueryAsync(ct);
    }
}
