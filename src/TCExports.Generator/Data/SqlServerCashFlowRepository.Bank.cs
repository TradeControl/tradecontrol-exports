using Microsoft.Data.SqlClient;
using System.Data;
using TCExports.Generator.Contracts;

namespace TCExports.Generator.Data;

public sealed partial class SqlServerCashFlowRepository
{
    public async Task<IReadOnlyList<BankAccountDto>> GetBankAccountsAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        const string sql = """
            SELECT AccountCode, AccountName, OpeningBalance
            FROM Cash.vwBankAccounts
            ORDER BY DisplayOrder, AccountCode;
            """;
        var list = new List<BankAccountDto>();
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var conn = new SqlConnection(ado);
        await conn.OpenAsync(ct);
        await using var cmd = Query(conn, sql, commandTimeoutSeconds);
        await using var rdr = await cmd.ExecuteReaderAsync(ct);
        while (await rdr.ReadAsync(ct))
        {
            list.Add(new BankAccountDto
            {
                AccountCode = rdr.GetString(rdr.GetOrdinal("AccountCode")),
                AccountName = rdr.GetString(rdr.GetOrdinal("AccountName")),
                OpeningBalance = rdr.GetDecimal(rdr.GetOrdinal("OpeningBalance"))
            });
        }
        return list;
    }

    public async Task<IReadOnlyList<BankBalanceDto>> GetBankBalancesAsync(string connectionString, string accountCode, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        const string sql = """
            SELECT @AccountCode AS AccountCode,
                   YearNumber,
                   CAST(MONTH(StartOn) AS tinyint) AS MonthNumber,
                   CAST(Balance AS decimal(18,5)) AS Balance
            FROM Cash.fnFlowBankBalances(@AccountCode)
            ORDER BY YearNumber, StartOn;
            """;
        var list = new List<BankBalanceDto>();
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var conn = new SqlConnection(ado);
        await conn.OpenAsync(ct);
        await using var cmd = Query(conn, sql, commandTimeoutSeconds);
        cmd.Parameters.Add(new SqlParameter("@AccountCode", SqlDbType.NVarChar, 10) { Value = accountCode });
        await using var rdr = await cmd.ExecuteReaderAsync(ct);
        while (await rdr.ReadAsync(ct))
        {
            list.Add(new BankBalanceDto
            {
                AccountCode = rdr.GetString(rdr.GetOrdinal("AccountCode")),
                YearNumber = rdr.GetInt16(rdr.GetOrdinal("YearNumber")),
                MonthNumber = rdr.GetByte(rdr.GetOrdinal("MonthNumber")),
                Balance = rdr.GetDecimal(rdr.GetOrdinal("Balance"))
            });
        }
        return list;
    }
}

