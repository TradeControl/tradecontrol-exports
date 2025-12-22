using Microsoft.Data.SqlClient;

namespace TCExports.Generator.Data;

internal static class ConnectionStringUtil
{
    public static string ToSqlClient(string connectionString)
    {
        if (string.IsNullOrWhiteSpace(connectionString)) return connectionString;

        if (connectionString.IndexOf("Driver=", StringComparison.OrdinalIgnoreCase) >= 0 ||
            connectionString.IndexOf("Trusted_Connection", StringComparison.OrdinalIgnoreCase) >= 0 ||
            connectionString.IndexOf("Trust_Server_Certificate", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var seg in connectionString.Split(';', StringSplitOptions.RemoveEmptyEntries))
            {
                var kv = seg.Split('=', 2);
                if (kv.Length == 2) dict[kv[0].Trim()] = kv[1].Trim();
            }

            var b = new SqlConnectionStringBuilder();

            if (dict.TryGetValue("Server", out var server) || dict.TryGetValue("Data Source", out server))
                b.DataSource = server;

            if (dict.TryGetValue("Database", out var db) || dict.TryGetValue("Initial Catalog", out db))
                b.InitialCatalog = db;

            if (dict.TryGetValue("Trusted_Connection", out var trusted) || dict.TryGetValue("Integrated Security", out trusted))
                b.IntegratedSecurity = IsTrue(trusted);

            if (dict.TryGetValue("TrustServerCertificate", out var tsc) || dict.TryGetValue("Trust_Server_Certificate", out tsc))
                b.TrustServerCertificate = IsTrue(tsc);

            if (!b.IntegratedSecurity)
            {
                if (dict.TryGetValue("Uid", out var uid) || dict.TryGetValue("User Id", out uid))
                    b.UserID = uid;
                if (dict.TryGetValue("Pwd", out var pwd) || dict.TryGetValue("Password", out pwd))
                    b.Password = pwd;
            }

            if (dict.TryGetValue("Encrypt", out var enc))
                b.Encrypt = IsTrue(enc);

            return b.ConnectionString;
        }

        return connectionString;
    }

    private static bool IsTrue(string s) =>
        s.Equals("yes", StringComparison.OrdinalIgnoreCase) ||
        s.Equals("true", StringComparison.OrdinalIgnoreCase) ||
        s.Equals("1", StringComparison.OrdinalIgnoreCase);
}