using Microsoft.Data.SqlClient;
using TCExports.Generator.Contracts;

namespace TCExports.Generator.Data;

public sealed partial class SqlServerCashFlowRepository : ICashFlowRepository
{
    public async Task<ActivePeriodDto> GetActivePeriodAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        var adoConnString = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var conn = new SqlConnection(adoConnString);
        await conn.OpenAsync(ct);
        await using var cmd = Query(
            conn,
            "SELECT TOP (1) YearNumber, MonthNumber, StartOn, MonthName, Description FROM App.vwActivePeriod;",
            commandTimeoutSeconds);
        await using var rdr = await cmd.ExecuteReaderAsync(ct);

        if (!await rdr.ReadAsync(ct))
            throw new InvalidOperationException("Active period not found.");

        return new ActivePeriodDto
        {
            YearNumber = rdr.GetInt16(rdr.GetOrdinal("YearNumber")),
            MonthNumber = rdr.GetInt16(rdr.GetOrdinal("MonthNumber")),
            StartOn = rdr.GetDateTime(rdr.GetOrdinal("StartOn")),
            MonthName = rdr.GetString(rdr.GetOrdinal("MonthName")),
            Description = rdr.GetString(rdr.GetOrdinal("Description"))
        };
    }

    public async Task<IReadOnlyList<ActiveYearDto>> GetActiveYearsAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        var list = new List<ActiveYearDto>();
        var adoConnString = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var conn = new SqlConnection(adoConnString);
        await conn.OpenAsync(ct);
        await using var cmd = Query(
            conn,
            "SELECT YearNumber, Description, CashStatus FROM App.vwActiveYears ORDER BY YearNumber;",
            commandTimeoutSeconds);
        await using var rdr = await cmd.ExecuteReaderAsync(ct);

        while (await rdr.ReadAsync(ct))
        {
            list.Add(new ActiveYearDto
            {
                YearNumber = rdr.GetInt16(rdr.GetOrdinal("YearNumber")),
                Description = rdr.GetString(rdr.GetOrdinal("Description")),
                CashStatus = rdr.GetString(rdr.GetOrdinal("CashStatus"))
            });
        }
        return list;
    }

    public async Task<IReadOnlyList<MonthDto>> GetMonthsAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        var list = new List<MonthDto>();
        var adoConnString = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var conn = new SqlConnection(adoConnString);
        await conn.OpenAsync(ct);
        await using var cmd = Query(
            conn,
            "SELECT MonthNumber, MonthName, StartOn FROM App.vwMonths ORDER BY StartOn;",
            commandTimeoutSeconds);
        await using var rdr = await cmd.ExecuteReaderAsync(ct);

        while (await rdr.ReadAsync(ct))
        {
            list.Add(new MonthDto
            {
                MonthNumber = rdr.GetInt16(rdr.GetOrdinal("MonthNumber")),
                MonthName   = rdr.GetString(rdr.GetOrdinal("MonthName")),
                StartOn     = rdr.IsDBNull(rdr.GetOrdinal("StartOn")) ? (decimal?)null : rdr.GetDecimal(rdr.GetOrdinal("StartOn"))
            });
        }
        return list;
    }
}
