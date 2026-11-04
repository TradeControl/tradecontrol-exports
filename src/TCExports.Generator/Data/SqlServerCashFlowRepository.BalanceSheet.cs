using Microsoft.Data.SqlClient;
using TCExports.Generator.Contracts;

namespace TCExports.Generator.Data;

public sealed partial class SqlServerCashFlowRepository
{
    public async Task<IReadOnlyList<BalanceSheetEntryDto>> GetBalanceSheetAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        const string sql = """
            SELECT AssetCode,
                   AssetName,
                   YearNumber,
                   CAST(MONTH(StartOn) AS tinyint) AS MonthNumber,
                   CAST(Balance AS decimal(18,5)) AS Balance
            FROM Cash.vwBalanceSheet
            ORDER BY EntryNumber;
            """;
        var list = new List<BalanceSheetEntryDto>();
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var conn = new SqlConnection(ado);
        await conn.OpenAsync(ct);
        await using var cmd = Query(conn, sql, commandTimeoutSeconds);
        await using var rdr = await cmd.ExecuteReaderAsync(ct);
        while (await rdr.ReadAsync(ct))
        {
            list.Add(new BalanceSheetEntryDto
            {
                AssetCode = rdr.GetString(rdr.GetOrdinal("AssetCode")),
                AssetName = rdr.GetString(rdr.GetOrdinal("AssetName")),
                YearNumber = rdr.GetInt16(rdr.GetOrdinal("YearNumber")),
                MonthNumber = rdr.GetByte(rdr.GetOrdinal("MonthNumber")),
                Balance = rdr.GetDecimal(rdr.GetOrdinal("Balance"))
            });
        }
        return list;
    }
}

