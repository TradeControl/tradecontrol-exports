using Microsoft.Data.SqlClient;
using System.Data;
using TCExports.Generator.Contracts;

namespace TCExports.Generator.Data;

public sealed partial class SqlServerCashFlowRepository : ICashFlowRepository
{
    public async Task<IReadOnlyList<FlowCategoryDto>> GetCategoriesAsync(string connectionString, CashType cashType, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        const string sql = @"
            SELECT CategoryCode, Category, CashPolarityCode, DisplayOrder
            FROM Cash.fnFlowCategory(@CashTypeCode)
            ORDER BY DisplayOrder, Category;";
        var list = new List<FlowCategoryDto>();
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var conn = new SqlConnection(ado);
        await conn.OpenAsync(ct);
        await using var cmd = new SqlCommand(sql, conn) { CommandType = CommandType.Text, CommandTimeout = commandTimeoutSeconds };
        cmd.Parameters.Add(new SqlParameter("@CashTypeCode", SqlDbType.SmallInt) { Value = (short)cashType });
        await using var rdr = await cmd.ExecuteReaderAsync(ct);
        while (await rdr.ReadAsync(ct))
        {
            list.Add(new FlowCategoryDto
            {
                CategoryCode = rdr.GetString(rdr.GetOrdinal("CategoryCode")),
                Category = rdr.GetString(rdr.GetOrdinal("Category")),
                CashPolarityCode = rdr.IsDBNull(rdr.GetOrdinal("CashPolarityCode")) ? (short?)null : rdr.GetInt16(rdr.GetOrdinal("CashPolarityCode")),
                DisplayOrder = rdr.IsDBNull(rdr.GetOrdinal("DisplayOrder")) ? (short?)null : rdr.GetInt16(rdr.GetOrdinal("DisplayOrder"))
            });
        }
        return list;
    }

    public async Task<IReadOnlyList<FlowCategoryCashCodeDto>> GetCashCodesAsync(string connectionString, string categoryCode, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        const string sql = @"
            SELECT CashCode, CashDescription
            FROM Cash.fnFlowCategoryCashCodes(@CategoryCode)
            ORDER BY CashDescription;";
        var list = new List<FlowCategoryCashCodeDto>();
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var conn = new SqlConnection(ado);
        await conn.OpenAsync(ct);
        await using var cmd = new SqlCommand(sql, conn) { CommandType = CommandType.Text, CommandTimeout = commandTimeoutSeconds };
        cmd.Parameters.Add(new SqlParameter("@CategoryCode", SqlDbType.NVarChar, 10) { Value = categoryCode });
        await using var rdr = await cmd.ExecuteReaderAsync(ct);
        while (await rdr.ReadAsync(ct))
        {
            list.Add(new FlowCategoryCashCodeDto
            {
                CashCode = rdr.GetString(rdr.GetOrdinal("CashCode")),
                CashDescription = rdr.GetString(rdr.GetOrdinal("CashDescription"))
            });
        }
        return list;
    }

    public async Task<IReadOnlyList<FlowCategoriesByTypeDto>> GetCategoriesByTypeAsync(string connectionString, CashType cashType, CategoryType categoryType, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        const string sql = @"
            SELECT CategoryCode, Category, CashType
            FROM Cash.fnFlowCategoriesByType(@CashTypeCode, @CategoryTypeCode)
            ORDER BY DisplayOrder, Category;";
        var list = new List<FlowCategoriesByTypeDto>();
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var conn = new SqlConnection(ado);
        await conn.OpenAsync(ct);
        await using var cmd = new SqlCommand(sql, conn) { CommandType = CommandType.Text, CommandTimeout = commandTimeoutSeconds };
        cmd.Parameters.Add(new SqlParameter("@CashTypeCode", SqlDbType.SmallInt) { Value = (short)cashType });
        cmd.Parameters.Add(new SqlParameter("@CategoryTypeCode", SqlDbType.SmallInt) { Value = (short)categoryType });
        await using var rdr = await cmd.ExecuteReaderAsync(ct);
        while (await rdr.ReadAsync(ct))
        {
            list.Add(new FlowCategoriesByTypeDto
            {
                CategoryCode = rdr.GetString(rdr.GetOrdinal("CategoryCode")),
                Category = rdr.GetString(rdr.GetOrdinal("Category")),
                CashType = rdr.IsDBNull(rdr.GetOrdinal("CashType")) ? "" : rdr.GetString(rdr.GetOrdinal("CashType"))
            });
        }
        return list;
    }
}