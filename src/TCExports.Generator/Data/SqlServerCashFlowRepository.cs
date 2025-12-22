using Microsoft.Data.SqlClient;
using System.Data;
using TCExports.Generator.Contracts;

namespace TCExports.Generator.Data;

public sealed partial class SqlServerCashFlowRepository : ICashFlowRepository
{
    public async Task<IReadOnlyList<CashCodePeriodValue>> GetCashCodeValuesAsync(
        string connectionString,
        string cashCode,
        short yearNumber,
        bool includeActivePeriods,
        bool includeOrderBook,
        bool includeTaxAccruals,
        int commandTimeoutSeconds = 30,
        CancellationToken ct = default)
    {
        var results = new List<CashCodePeriodValue>();

        var adoConnString = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var conn = new SqlConnection(adoConnString);
        await conn.OpenAsync(ct);

        await using var cmd = new SqlCommand("Cash.proc_FlowCashCodeValues", conn)
        {
            CommandType = CommandType.StoredProcedure,
            CommandTimeout = commandTimeoutSeconds
        };

        cmd.Parameters.Add(new SqlParameter("@CashCode", SqlDbType.NVarChar, 50) { Value = cashCode });
        cmd.Parameters.Add(new SqlParameter("@YearNumber", SqlDbType.SmallInt) { Value = yearNumber });
        cmd.Parameters.Add(new SqlParameter("@IncludeActivePeriods", SqlDbType.Bit) { Value = includeActivePeriods });
        cmd.Parameters.Add(new SqlParameter("@IncludeOrderBook", SqlDbType.Bit) { Value = includeOrderBook });
        cmd.Parameters.Add(new SqlParameter("@IncludeTaxAccruals", SqlDbType.Bit) { Value = includeTaxAccruals });

        await using var reader = await cmd.ExecuteReaderAsync(ct);
        while (await reader.ReadAsync(ct))
        {
            results.Add(new CashCodePeriodValue
            {
                StartOn = reader.GetDateTime(0),
                InvoiceValue = reader.GetDecimal(1),
                InvoiceTax = reader.GetDecimal(2),
                ForecastValue = reader.GetDecimal(3),
                ForecastTax = reader.GetDecimal(4)
            });
        }

        return results;
    }
}