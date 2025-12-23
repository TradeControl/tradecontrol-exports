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

    public async Task<string> GetCompanyNameAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        const string sql = @"SELECT TOP (1) SubjectName FROM App.vwHomeAccount;";
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var conn = new SqlConnection(ado);
        await conn.OpenAsync(ct);
        await using var cmd = new SqlCommand(sql, conn) { CommandType = CommandType.Text, CommandTimeout = commandTimeoutSeconds };
        var result = await cmd.ExecuteScalarAsync(ct);
        return result as string ?? string.Empty;
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

    public async Task<IReadOnlyList<CategoryTotalDto>> GetCategoryTotalsAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        const string sql = @"SELECT CategoryCode FROM Cash.vwCategoryTotals ORDER BY DisplayOrder, Category;";
        var list = new List<CategoryTotalDto>();
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var conn = new SqlConnection(ado);
        await conn.OpenAsync(ct);
        await using var cmd = new SqlCommand(sql, conn) { CommandType = CommandType.Text, CommandTimeout = commandTimeoutSeconds };
        await using var rdr = await cmd.ExecuteReaderAsync(ct);
        while (await rdr.ReadAsync(ct))
            list.Add(new CategoryTotalDto { CategoryCode = rdr.GetString(0) });
        return list;
    }

    public async Task<IReadOnlyList<string>> GetCategoryTotalCodesAsync(string connectionString, string categoryCode, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        const string sql = @"SELECT CategoryCode FROM Cash.fnFlowCategoryTotalCodes(@CategoryCode) ORDER BY CategoryCode;";
        var list = new List<string>();
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var conn = new SqlConnection(ado);
        await conn.OpenAsync(ct);
        await using var cmd = new SqlCommand(sql, conn) { CommandType = CommandType.Text, CommandTimeout = commandTimeoutSeconds };
        cmd.Parameters.Add(new SqlParameter("@CategoryCode", SqlDbType.NVarChar, 10) { Value = categoryCode });
        await using var rdr = await cmd.ExecuteReaderAsync(ct);
        while (await rdr.ReadAsync(ct))
            list.Add(rdr.GetString(0));
        return list;
    }

    public async Task<IReadOnlyList<CategoryExpressionDto>> GetCategoryExpressionsAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        const string sql = @"
            SELECT DisplayOrder, Category, Expression, Format
            FROM Cash.vwCategoryExpressions
            WHERE SyntaxTypeCode IN (0, 2)
            ORDER BY DisplayOrder, Category;";
        var list = new List<CategoryExpressionDto>();
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var conn = new SqlConnection(ado);
        await conn.OpenAsync(ct);
        await using var cmd = new SqlCommand(sql, conn) { CommandType = CommandType.Text, CommandTimeout = commandTimeoutSeconds };
        await using var rdr = await cmd.ExecuteReaderAsync(ct);
        while (await rdr.ReadAsync(ct))
        {
            list.Add(new CategoryExpressionDto
            {
                DisplayOrder = rdr.GetInt16(rdr.GetOrdinal("DisplayOrder")),
                Category = rdr.GetString(rdr.GetOrdinal("Category")),
                Expression = rdr.GetString(rdr.GetOrdinal("Expression")),
                Format = rdr.GetString(rdr.GetOrdinal("Format"))
            });
        }
        return list;
    }

    public async Task<string> GetCategoryCodeFromNameAsync(string connectionString, string categoryName, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        // Matches DataContext.proc_FlowCategoryCodeFromName usage
        const string procName = "Cash.proc_FlowCategoryCodeFromName";
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var conn = new SqlConnection(ado);
        await conn.OpenAsync(ct);
        await using var cmd = new SqlCommand(procName, conn) { CommandType = CommandType.StoredProcedure, CommandTimeout = commandTimeoutSeconds };
        cmd.Parameters.Add(new SqlParameter("@Category", SqlDbType.NVarChar, 50) { Value = categoryName });
        var outParam = new SqlParameter("@CategoryCode", SqlDbType.NVarChar, 10) { Direction = ParameterDirection.Output };
        cmd.Parameters.Add(outParam);
        await cmd.ExecuteNonQueryAsync(ct);
        return outParam.Value as string ?? string.Empty;
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

    public async Task<IReadOnlyList<BankAccountDto>> GetBankAccountsAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        const string sql = @"
            SELECT AccountCode, AccountName, OpeningBalance
            FROM Cash.vwBankAccounts
            ORDER BY DisplayOrder, AccountCode;";
        var list = new List<BankAccountDto>();
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var conn = new SqlConnection(ado);
        await conn.OpenAsync(ct);
        await using var cmd = new SqlCommand(sql, conn) { CommandType = CommandType.Text, CommandTimeout = commandTimeoutSeconds };
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
        const string sql = @"
            SELECT 
                @AccountCode AS AccountCode,
                YearNumber,
                CAST(MONTH(StartOn) AS tinyint) AS MonthNumber,
                CAST(Balance AS decimal(18,5)) AS Balance
            FROM Cash.fnFlowBankBalances(@AccountCode)
            ORDER BY YearNumber, StartOn;";
        var list = new List<BankBalanceDto>();
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var conn = new SqlConnection(ado);
        await conn.OpenAsync(ct);
        await using var cmd = new SqlCommand(sql, conn) { CommandType = CommandType.Text, CommandTimeout = commandTimeoutSeconds };
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

    public async Task<string> GetVatRecurrenceTypeAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        // Matches legacy DataContext.VatRecurrenceType (TaxType.Vat == TaxType.Vat)
        const string sql = @"
            SELECT TOP (1) UPPER(Recurrence) 
            FROM Cash.vwFlowTaxType 
            WHERE TaxTypeCode = 1;";
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var conn = new SqlConnection(ado);
        await conn.OpenAsync(ct);
        await using var cmd = new SqlCommand(sql, conn) { CommandType = CommandType.Text, CommandTimeout = commandTimeoutSeconds };
        var result = await cmd.ExecuteScalarAsync(ct);
        return result as string ?? string.Empty;
    }

    public async Task<IReadOnlyList<VatRecurrenceDto>> GetVatRecurrenceAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        const string sql = @"
            SELECT YearNumber, StartOn, HomeSales, HomePurchases, ExportSales, ExportPurchases,
                   HomeSalesVat, HomePurchasesVat, ExportSalesVat, ExportPurchasesVat, VatAdjustment, VatDue
            FROM Cash.vwFlowVatRecurrence
            ORDER BY YearNumber, StartOn;";
        var list = new List<VatRecurrenceDto>();
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var conn = new SqlConnection(ado);
        await conn.OpenAsync(ct);
        await using var cmd = new SqlCommand(sql, conn) { CommandType = CommandType.Text, CommandTimeout = commandTimeoutSeconds };
        await using var rdr = await cmd.ExecuteReaderAsync(ct);
        while (await rdr.ReadAsync(ct))
        {
            list.Add(new VatRecurrenceDto
            {
                YearNumber = rdr.GetInt16(rdr.GetOrdinal("YearNumber")),
                StartOn = rdr.GetDateTime(rdr.GetOrdinal("StartOn")),
                HomeSales = rdr.GetDecimal(rdr.GetOrdinal("HomeSales")),
                HomePurchases = rdr.GetDecimal(rdr.GetOrdinal("HomePurchases")),
                ExportSales = rdr.GetDecimal(rdr.GetOrdinal("ExportSales")),
                ExportPurchases = rdr.GetDecimal(rdr.GetOrdinal("ExportPurchases")),
                HomeSalesVat = rdr.GetDecimal(rdr.GetOrdinal("HomeSalesVat")),
                HomePurchasesVat = rdr.GetDecimal(rdr.GetOrdinal("HomePurchasesVat")),
                ExportSalesVat = rdr.GetDecimal(rdr.GetOrdinal("ExportSalesVat")),
                ExportPurchasesVat = rdr.GetDecimal(rdr.GetOrdinal("ExportPurchasesVat")),
                VatAdjustment = rdr.GetDecimal(rdr.GetOrdinal("VatAdjustment")),
                VatDue = rdr.GetDecimal(rdr.GetOrdinal("VatDue"))
            });
        }
        return list;
    }

    public async Task<IReadOnlyList<VatRecurrenceAccrualDto>> GetVatRecurrenceAccrualsAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        const string sql = @"
            SELECT YearNumber, HomeSalesVat, HomePurchasesVat, ExportSalesVat, ExportPurchasesVat, VatDue
            FROM Cash.vwFlowVatRecurrenceAccruals
            ORDER BY YearNumber;";
        var list = new List<VatRecurrenceAccrualDto>();
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var conn = new SqlConnection(ado);
        await conn.OpenAsync(ct);
        await using var cmd = new SqlCommand(sql, conn) { CommandType = CommandType.Text, CommandTimeout = commandTimeoutSeconds };
        await using var rdr = await cmd.ExecuteReaderAsync(ct);
        while (await rdr.ReadAsync(ct))
        {
            list.Add(new VatRecurrenceAccrualDto
            {
                YearNumber = rdr.GetInt16(rdr.GetOrdinal("YearNumber")),
                HomeSalesVat = rdr.IsDBNull(rdr.GetOrdinal("HomeSalesVat")) ? (decimal?)null : rdr.GetDecimal(rdr.GetOrdinal("HomeSalesVat")),
                HomePurchasesVat = rdr.IsDBNull(rdr.GetOrdinal("HomePurchasesVat")) ? (decimal?)null : rdr.GetDecimal(rdr.GetOrdinal("HomePurchasesVat")),
                ExportSalesVat = rdr.IsDBNull(rdr.GetOrdinal("ExportSalesVat")) ? (decimal?)null : rdr.GetDecimal(rdr.GetOrdinal("ExportSalesVat")),
                ExportPurchasesVat = rdr.IsDBNull(rdr.GetOrdinal("ExportPurchasesVat")) ? (decimal?)null : rdr.GetDecimal(rdr.GetOrdinal("ExportPurchasesVat")),
                VatDue = rdr.IsDBNull(rdr.GetOrdinal("VatDue")) ? (decimal?)null : rdr.GetDecimal(rdr.GetOrdinal("VatDue"))
            });
        }
        return list;
    }

    public async Task<IReadOnlyList<VatPeriodTotalDto>> GetVatPeriodTotalsAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        const string sql = @"
            SELECT YearNumber, StartOn, HomeSales, HomePurchases, ExportSales, ExportPurchases,
                   HomeSalesVat, HomePurchasesVat, ExportSalesVat, ExportPurchasesVat, VatDue
            FROM Cash.vwFlowVatPeriodTotals
            ORDER BY YearNumber, StartOn;";
        var list = new List<VatPeriodTotalDto>();
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var conn = new SqlConnection(ado);
        await conn.OpenAsync(ct);
        await using var cmd = new SqlCommand(sql, conn) { CommandType = CommandType.Text, CommandTimeout = commandTimeoutSeconds };
        await using var rdr = await cmd.ExecuteReaderAsync(ct);
        while (await rdr.ReadAsync(ct))
        {
            list.Add(new VatPeriodTotalDto
            {
                YearNumber = rdr.GetInt16(rdr.GetOrdinal("YearNumber")),
                StartOn = rdr.GetDateTime(rdr.GetOrdinal("StartOn")),
                HomeSales = rdr.GetDecimal(rdr.GetOrdinal("HomeSales")),
                HomePurchases = rdr.GetDecimal(rdr.GetOrdinal("HomePurchases")),
                ExportSales = rdr.GetDecimal(rdr.GetOrdinal("ExportSales")),
                ExportPurchases = rdr.GetDecimal(rdr.GetOrdinal("ExportPurchases")),
                HomeSalesVat = rdr.GetDecimal(rdr.GetOrdinal("HomeSalesVat")),
                HomePurchasesVat = rdr.GetDecimal(rdr.GetOrdinal("HomePurchasesVat")),
                ExportSalesVat = rdr.GetDecimal(rdr.GetOrdinal("ExportSalesVat")),
                ExportPurchasesVat = rdr.GetDecimal(rdr.GetOrdinal("ExportPurchasesVat")),
                VatDue = rdr.GetDecimal(rdr.GetOrdinal("VatDue"))
            });
        }
        return list;
    }

    public async Task<IReadOnlyList<VatPeriodAccrualDto>> GetVatPeriodAccrualsAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        const string sql = @"
            SELECT YearNumber, HomeSalesVat, HomePurchasesVat, ExportSalesVat, ExportPurchasesVat, VatDue
            FROM Cash.vwFlowVatPeriodAccruals
            ORDER BY YearNumber;";
        var list = new List<VatPeriodAccrualDto>();
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var conn = new SqlConnection(ado);
        await conn.OpenAsync(ct);
        await using var cmd = new SqlCommand(sql, conn) { CommandType = CommandType.Text, CommandTimeout = commandTimeoutSeconds };
        await using var rdr = await cmd.ExecuteReaderAsync(ct);
        while (await rdr.ReadAsync(ct))
        {
            list.Add(new VatPeriodAccrualDto
            {
                YearNumber = rdr.GetInt16(rdr.GetOrdinal("YearNumber")),
                HomeSalesVat = rdr.IsDBNull(rdr.GetOrdinal("HomeSalesVat")) ? (decimal?)null : rdr.GetDecimal(rdr.GetOrdinal("HomeSalesVat")),
                HomePurchasesVat = rdr.IsDBNull(rdr.GetOrdinal("HomePurchasesVat")) ? (decimal?)null : rdr.GetDecimal(rdr.GetOrdinal("HomePurchasesVat")),
                ExportSalesVat = rdr.IsDBNull(rdr.GetOrdinal("ExportSalesVat")) ? (decimal?)null : rdr.GetDecimal(rdr.GetOrdinal("ExportSalesVat")),
                ExportPurchasesVat = rdr.IsDBNull(rdr.GetOrdinal("ExportPurchasesVat")) ? (decimal?)null : rdr.GetDecimal(rdr.GetOrdinal("ExportPurchasesVat")),
                VatDue = rdr.IsDBNull(rdr.GetOrdinal("VatDue")) ? (decimal?)null : rdr.GetDecimal(rdr.GetOrdinal("VatDue"))
            });
        }
        return list;
    }

    public async Task<IReadOnlyList<BalanceSheetEntryDto>> GetBalanceSheetAsync(string connectionString, int commandTimeoutSeconds = 30, CancellationToken ct = default)
    {
        const string sql = @"
            SELECT 
                AssetCode,
                AssetName,
                YearNumber,
                CAST(MONTH(StartOn) AS tinyint) AS MonthNumber,
                CAST(Balance AS decimal(18,5)) AS Balance
            FROM Cash.vwBalanceSheet
            ORDER BY EntryNumber;";
        var list = new List<BalanceSheetEntryDto>();
        var ado = ConnectionStringUtil.ToSqlClient(connectionString);
        await using var conn = new SqlConnection(ado);
        await conn.OpenAsync(ct);
        await using var cmd = new SqlCommand(sql, conn) { CommandType = CommandType.Text, CommandTimeout = commandTimeoutSeconds };
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