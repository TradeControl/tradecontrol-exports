using Microsoft.Data.SqlClient;
using System.Data;
using TCExports.Generator.Contracts;

namespace TCExports.Generator.Data;

public sealed partial class SqlServerCashFlowRepository
{
    public async Task<IReadOnlyList<EquityReconciliationByYearDto>> GetEquityReconciliationByYearAsync(
        string connectionString,
        int commandTimeoutSeconds = 30,
        CancellationToken ct = default)
    {
        var results = new List<EquityReconciliationByYearDto>();

        var adoConnString = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var conn = new SqlConnection(adoConnString);
        await conn.OpenAsync(ct);

        const string sql = """
            SELECT
                YearNumber,
                [Description],
                OpeningCapital,
                ClosingCapital,
                Profit,
                BusinessTax,
                ProfitAfterTax,
                CapitalMovement,
                OpeningSubjectPosition,
                OpeningAccountPosition,
                CapitalDelta,
                Variance
            FROM Cash.vwEquityReconciliationByYear
            ORDER BY YearNumber;
            """;

        await using var cmd = new SqlCommand(sql, conn) {
            CommandType = CommandType.Text,
            CommandTimeout = commandTimeoutSeconds
        };

        await using var reader = await cmd.ExecuteReaderAsync(ct);
        while (await reader.ReadAsync(ct))
        {
            results.Add(new EquityReconciliationByYearDto {
                YearNumber = reader.GetInt16(0),
                Description = reader.IsDBNull(1) ? null : reader.GetString(1),
                OpeningCapital = reader.GetDecimal(2),
                ClosingCapital = reader.GetDecimal(3),
                Profit = reader.GetDecimal(4),
                BusinessTax = reader.GetDecimal(5),
                ProfitAfterTax = reader.GetDecimal(6),
                CapitalMovement = reader.GetDecimal(7),
                OpeningSubjectPosition = reader.GetDecimal(8),
                OpeningAccountPosition = reader.GetDecimal(9),
                CapitalDelta = reader.GetDecimal(10),
                Variance = reader.GetDecimal(11)
            });
        }

        return results;
    }
}
