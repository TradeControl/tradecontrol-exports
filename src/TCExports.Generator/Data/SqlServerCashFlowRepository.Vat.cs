using Microsoft.Data.SqlClient;
using TCExports.Generator.Contracts;

namespace TCExports.Generator.Data;

public sealed partial class SqlServerCashFlowRepository
{
    public async Task<string> GetVatRecurrenceTypeAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        // Matches legacy DataContext.VatRecurrenceType (TaxType.Vat == TaxType.Vat)
        const string sql = "SELECT TOP (1) UPPER(Recurrence) FROM Cash.vwFlowTaxType WHERE TaxTypeCode = 1;";
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var conn = new SqlConnection(ado);
        await conn.OpenAsync(ct);
        await using var cmd = Query(conn, sql, commandTimeoutSeconds);
        var result = await cmd.ExecuteScalarAsync(ct);
        return result as string ?? string.Empty;
    }

    public async Task<IReadOnlyList<VatRecurrenceDto>> GetVatRecurrenceAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        const string sql = """
            SELECT YearNumber, StartOn, vatDueSales, vatDueAcquisitions, totalVatDue,
                   vatReclaimedCurrPeriod, netVatDue, totalValueSalesExVAT,
                   totalValuePurchasesExVAT, totalValueGoodsSuppliedExVAT,
                   totalValueGoodsReceivedExVAT
            FROM Cash.vwFlowVatRecurrence
            ORDER BY YearNumber, StartOn;
            """;
        var list = new List<VatRecurrenceDto>();
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var conn = new SqlConnection(ado);
        await conn.OpenAsync(ct);
        await using var cmd = Query(conn, sql, commandTimeoutSeconds);
        await using var rdr = await cmd.ExecuteReaderAsync(ct);
        while (await rdr.ReadAsync(ct))
        {
            list.Add(new VatRecurrenceDto
            {
                YearNumber = rdr.GetInt16(rdr.GetOrdinal("YearNumber")),
                StartOn = rdr.GetDateTime(rdr.GetOrdinal("StartOn")),
                VatDueSales = rdr.GetDecimal(rdr.GetOrdinal("vatDueSales")),
                VatDueAcquisitions = rdr.GetDecimal(rdr.GetOrdinal("vatDueAcquisitions")),
                TotalVatDue = rdr.GetDecimal(rdr.GetOrdinal("totalVatDue")),
                VatReclaimedCurrPeriod = rdr.GetDecimal(rdr.GetOrdinal("vatReclaimedCurrPeriod")),
                NetVatDue = rdr.GetDecimal(rdr.GetOrdinal("netVatDue")),
                TotalValueSalesExVAT = rdr.GetDecimal(rdr.GetOrdinal("totalValueSalesExVAT")),
                TotalValuePurchasesExVAT = rdr.GetDecimal(rdr.GetOrdinal("totalValuePurchasesExVAT")),
                TotalValueGoodsSuppliedExVAT = rdr.GetDecimal(rdr.GetOrdinal("totalValueGoodsSuppliedExVAT")),
                TotalValueGoodsReceivedExVAT = rdr.GetDecimal(rdr.GetOrdinal("totalValueGoodsReceivedExVAT"))
            });
        }
        return list;
    }

    public async Task<IReadOnlyList<VatRecurrenceAccrualDto>> GetVatRecurrenceAccrualsAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        const string sql = """
            SELECT YearNumber, StartOn, vatDueSales, vatDueAcquisitions, totalVatDue,
                   vatReclaimedCurrPeriod, netVatDue, totalValueSalesExVAT,
                   totalValuePurchasesExVAT, totalValueGoodsSuppliedExVAT,
                   totalValueGoodsReceivedExVAT
            FROM Cash.vwFlowVatRecurrenceAccruals
            ORDER BY YearNumber, StartOn;
            """;
        var list = new List<VatRecurrenceAccrualDto>();
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var conn = new SqlConnection(ado);
        await conn.OpenAsync(ct);
        await using var cmd = Query(conn, sql, commandTimeoutSeconds);
        await using var rdr = await cmd.ExecuteReaderAsync(ct);
        while (await rdr.ReadAsync(ct))
        {
            list.Add(new VatRecurrenceAccrualDto
            {
                YearNumber = rdr.GetInt16(rdr.GetOrdinal("YearNumber")),
                StartOn = rdr.GetDateTime(rdr.GetOrdinal("StartOn")),
                VatDueSales = rdr.IsDBNull(rdr.GetOrdinal("vatDueSales")) ? (decimal?)null : rdr.GetDecimal(rdr.GetOrdinal("vatDueSales")),
                VatDueAcquisitions = rdr.IsDBNull(rdr.GetOrdinal("vatDueAcquisitions")) ? (decimal?)null : rdr.GetDecimal(rdr.GetOrdinal("vatDueAcquisitions")),
                TotalVatDue = rdr.IsDBNull(rdr.GetOrdinal("totalVatDue")) ? (decimal?)null : rdr.GetDecimal(rdr.GetOrdinal("totalVatDue")),
                VatReclaimedCurrPeriod = rdr.IsDBNull(rdr.GetOrdinal("vatReclaimedCurrPeriod")) ? (decimal?)null : rdr.GetDecimal(rdr.GetOrdinal("vatReclaimedCurrPeriod")),
                NetVatDue = rdr.IsDBNull(rdr.GetOrdinal("netVatDue")) ? (decimal?)null : rdr.GetDecimal(rdr.GetOrdinal("netVatDue")),
                TotalValueSalesExVAT = rdr.IsDBNull(rdr.GetOrdinal("totalValueSalesExVAT")) ? (decimal?)null : rdr.GetDecimal(rdr.GetOrdinal("totalValueSalesExVAT")),
                TotalValuePurchasesExVAT = rdr.IsDBNull(rdr.GetOrdinal("totalValuePurchasesExVAT")) ? (decimal?)null : rdr.GetDecimal(rdr.GetOrdinal("totalValuePurchasesExVAT")),
                TotalValueGoodsSuppliedExVAT = rdr.IsDBNull(rdr.GetOrdinal("totalValueGoodsSuppliedExVAT")) ? (decimal?)null : rdr.GetDecimal(rdr.GetOrdinal("totalValueGoodsSuppliedExVAT")),
                TotalValueGoodsReceivedExVAT = rdr.IsDBNull(rdr.GetOrdinal("totalValueGoodsReceivedExVAT")) ? (decimal?)null : rdr.GetDecimal(rdr.GetOrdinal("totalValueGoodsReceivedExVAT"))
            });
        }
        return list;
    }

    public async Task<IReadOnlyList<VatPeriodTotalDto>> GetVatPeriodTotalsAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        const string sql = """
            SELECT YearNumber, StartOn, vatDueSales, vatDueAcquisitions, totalVatDue,
                   vatReclaimedCurrPeriod, netVatDue, totalValueSalesExVAT,
                   totalValuePurchasesExVAT, totalValueGoodsSuppliedExVAT,
                   totalValueGoodsReceivedExVAT
            FROM Cash.vwFlowVatPeriodTotals
            ORDER BY YearNumber, StartOn;
            """;
        var list = new List<VatPeriodTotalDto>();
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var conn = new SqlConnection(ado);
        await conn.OpenAsync(ct);
        await using var cmd = Query(conn, sql, commandTimeoutSeconds);
        await using var rdr = await cmd.ExecuteReaderAsync(ct);
        while (await rdr.ReadAsync(ct))
        {
            list.Add(new VatPeriodTotalDto
            {
                YearNumber = rdr.GetInt16(rdr.GetOrdinal("YearNumber")),
                StartOn = rdr.GetDateTime(rdr.GetOrdinal("StartOn")),
                VatDueSales = rdr.GetDecimal(rdr.GetOrdinal("vatDueSales")),
                VatDueAcquisitions = rdr.GetDecimal(rdr.GetOrdinal("vatDueAcquisitions")),
                TotalVatDue = rdr.GetDecimal(rdr.GetOrdinal("totalVatDue")),
                VatReclaimedCurrPeriod = rdr.GetDecimal(rdr.GetOrdinal("vatReclaimedCurrPeriod")),
                NetVatDue = rdr.GetDecimal(rdr.GetOrdinal("netVatDue")),
                TotalValueSalesExVAT = rdr.GetDecimal(rdr.GetOrdinal("totalValueSalesExVAT")),
                TotalValuePurchasesExVAT = rdr.GetDecimal(rdr.GetOrdinal("totalValuePurchasesExVAT")),
                TotalValueGoodsSuppliedExVAT = rdr.GetDecimal(rdr.GetOrdinal("totalValueGoodsSuppliedExVAT")),
                TotalValueGoodsReceivedExVAT = rdr.GetDecimal(rdr.GetOrdinal("totalValueGoodsReceivedExVAT"))
            });
        }
        return list;
    }

    public async Task<IReadOnlyList<VatPeriodAccrualDto>> GetVatPeriodAccrualsAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        const string sql = """
            SELECT YearNumber, StartOn, vatDueSales, vatDueAcquisitions, totalVatDue,
                   vatReclaimedCurrPeriod, netVatDue, totalValueSalesExVAT,
                   totalValuePurchasesExVAT, totalValueGoodsSuppliedExVAT,
                   totalValueGoodsReceivedExVAT
            FROM Cash.vwFlowVatPeriodAccruals
            ORDER BY YearNumber, StartOn;
            """;
        var list = new List<VatPeriodAccrualDto>();
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var conn = new SqlConnection(ado);
        await conn.OpenAsync(ct);
        await using var cmd = Query(conn, sql, commandTimeoutSeconds);
        await using var rdr = await cmd.ExecuteReaderAsync(ct);
        while (await rdr.ReadAsync(ct))
        {
            list.Add(new VatPeriodAccrualDto
            {
                YearNumber = rdr.GetInt16(rdr.GetOrdinal("YearNumber")),
                StartOn = rdr.GetDateTime(rdr.GetOrdinal("StartOn")),
                VatDueSales = rdr.IsDBNull(rdr.GetOrdinal("vatDueSales")) ? (decimal?)null : rdr.GetDecimal(rdr.GetOrdinal("vatDueSales")),
                VatDueAcquisitions = rdr.IsDBNull(rdr.GetOrdinal("vatDueAcquisitions")) ? (decimal?)null : rdr.GetDecimal(rdr.GetOrdinal("vatDueAcquisitions")),
                TotalVatDue = rdr.IsDBNull(rdr.GetOrdinal("totalVatDue")) ? (decimal?)null : rdr.GetDecimal(rdr.GetOrdinal("totalVatDue")),
                VatReclaimedCurrPeriod = rdr.IsDBNull(rdr.GetOrdinal("vatReclaimedCurrPeriod")) ? (decimal?)null : rdr.GetDecimal(rdr.GetOrdinal("vatReclaimedCurrPeriod")),
                NetVatDue = rdr.IsDBNull(rdr.GetOrdinal("netVatDue")) ? (decimal?)null : rdr.GetDecimal(rdr.GetOrdinal("netVatDue")),
                TotalValueSalesExVAT = rdr.IsDBNull(rdr.GetOrdinal("totalValueSalesExVAT")) ? (decimal?)null : rdr.GetDecimal(rdr.GetOrdinal("totalValueSalesExVAT")),
                TotalValuePurchasesExVAT = rdr.IsDBNull(rdr.GetOrdinal("totalValuePurchasesExVAT")) ? (decimal?)null : rdr.GetDecimal(rdr.GetOrdinal("totalValuePurchasesExVAT")),
                TotalValueGoodsSuppliedExVAT = rdr.IsDBNull(rdr.GetOrdinal("totalValueGoodsSuppliedExVAT")) ? (decimal?)null : rdr.GetDecimal(rdr.GetOrdinal("totalValueGoodsSuppliedExVAT")),
                TotalValueGoodsReceivedExVAT = rdr.IsDBNull(rdr.GetOrdinal("totalValueGoodsReceivedExVAT")) ? (decimal?)null : rdr.GetDecimal(rdr.GetOrdinal("totalValueGoodsReceivedExVAT"))
            });
        }
        return list;
    }
}

