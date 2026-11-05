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
        var adoConnString = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var conn = new SqlConnection(adoConnString);
        await conn.OpenAsync(ct);

        await using var cmd = Query(
            conn,
            """
            SELECT YearNumber, Description, OpeningCapital, ClosingCapital, Profit,
                   BusinessTax, ProfitAfterTax, CapitalMovement,
                   OpeningSubjectPosition, OpeningAccountPosition, CapitalDelta, Variance
            FROM Cash.vwEquityReconciliationByYear
            ORDER BY YearNumber;
            """,
            commandTimeoutSeconds);

        return await ReadEquityReconciliationAsync(cmd, ct);
    }

    public async Task<IReadOnlyList<EquityReconciliationByYearDto>> GetEquityReconciliationByYearByVersionAsync(
        string connectionString,
        string categoryVersionId,
        int commandTimeoutSeconds = 30,
        CancellationToken ct = default)
    {
        var adoConnString = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var conn = new SqlConnection(adoConnString);
        await conn.OpenAsync(ct);

        await using var cmd = Query(
            conn,
            """
            SELECT eb.YearNumber, eb.Description, eb.OpeningCapital, eb.ClosingCapital,
                   eb.Profit, eb.BusinessTax, eb.ProfitAfterTax, eb.CapitalMovement,
                   eb.OpeningSubjectPosition, eb.OpeningAccountPosition,
                   eb.CapitalDelta, eb.Variance
            FROM Cash.tbCategoryVersion AS cv
            CROSS APPLY Cash.fnEquityReconciliationByYearByVersion(cv.VersionId) AS eb
            WHERE cv.VersionId = @VersionId
            ORDER BY eb.YearNumber;
            """,
            commandTimeoutSeconds);
        cmd.Parameters.Add(new SqlParameter("@VersionId", SqlDbType.NVarChar, 10)
        {
            Value = categoryVersionId
        });

        return await ReadEquityReconciliationAsync(cmd, ct);
    }

    private static async Task<IReadOnlyList<EquityReconciliationByYearDto>> ReadEquityReconciliationAsync(
        SqlCommand cmd,
        CancellationToken ct)
    {
        var results = new List<EquityReconciliationByYearDto>();

        await using var reader = await cmd.ExecuteReaderAsync(ct);
        var yearNumberOrdinal = reader.GetOrdinal("YearNumber");
        var descriptionOrdinal = reader.GetOrdinal("Description");
        var openingCapitalOrdinal = reader.GetOrdinal("OpeningCapital");
        var closingCapitalOrdinal = reader.GetOrdinal("ClosingCapital");
        var profitOrdinal = reader.GetOrdinal("Profit");
        var businessTaxOrdinal = reader.GetOrdinal("BusinessTax");
        var profitAfterTaxOrdinal = reader.GetOrdinal("ProfitAfterTax");
        var capitalMovementOrdinal = reader.GetOrdinal("CapitalMovement");
        var openingSubjectPositionOrdinal = reader.GetOrdinal("OpeningSubjectPosition");
        var openingAccountPositionOrdinal = reader.GetOrdinal("OpeningAccountPosition");
        var capitalDeltaOrdinal = reader.GetOrdinal("CapitalDelta");
        var varianceOrdinal = reader.GetOrdinal("Variance");

        while (await reader.ReadAsync(ct))
        {
            results.Add(new EquityReconciliationByYearDto {
                YearNumber = reader.GetInt16(yearNumberOrdinal),
                Description = reader.IsDBNull(descriptionOrdinal) ? null : reader.GetString(descriptionOrdinal),
                OpeningCapital = reader.GetDecimal(openingCapitalOrdinal),
                ClosingCapital = reader.GetDecimal(closingCapitalOrdinal),
                Profit = reader.GetDecimal(profitOrdinal),
                BusinessTax = reader.GetDecimal(businessTaxOrdinal),
                ProfitAfterTax = reader.GetDecimal(profitAfterTaxOrdinal),
                CapitalMovement = reader.GetDecimal(capitalMovementOrdinal),
                OpeningSubjectPosition = reader.GetDecimal(openingSubjectPositionOrdinal),
                OpeningAccountPosition = reader.GetDecimal(openingAccountPositionOrdinal),
                CapitalDelta = reader.GetDecimal(capitalDeltaOrdinal),
                Variance = reader.GetDecimal(varianceOrdinal)
            });
        }

        return results;
    }
}
