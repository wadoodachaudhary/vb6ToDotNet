using HomeFront.Infrastructure.Logging;
using HomeFront.Services;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using System.Data;
using System.Diagnostics;
using System.Dynamic;
using System.Text.RegularExpressions;

namespace HomeFront.Data;

public class DbWrapperSqlServer
{
    private readonly string? _connectionStringOverride;
    private readonly ISessionStateService? _sessionState;
    private readonly ILogger<DbWrapperSqlServer> _logger;

    /// <summary>
    /// Default connection string read from appsettings.json at startup.
    /// Set once in Program.cs via DbWrapperSqlServer.DefaultConnectionString = ...
    /// </summary>
    public static string DefaultConnectionString { get; set; } = "";

    private static ILoggerFactory LoggerFactory { get; set; } = NullLoggerFactory.Instance;
    public static DatabaseLoggingOptions LoggingOptions { get; private set; } = new();

    public static void ConfigureLogging(ILoggerFactory? loggerFactory, DatabaseLoggingOptions? loggingOptions = null)
    {
        LoggerFactory = loggerFactory ?? NullLoggerFactory.Instance;
        LoggingOptions = loggingOptions ?? new DatabaseLoggingOptions();
    }

    public DbWrapperSqlServer(ISessionStateService? sessionState = null)
    {
        _sessionState = sessionState;
        _connectionStringOverride = null;
        _logger = LoggerFactory.CreateLogger<DbWrapperSqlServer>();
    }

    public DbWrapperSqlServer(string connectionString, ISessionStateService? sessionState = null)
    {
        _sessionState = sessionState;
        _connectionStringOverride = connectionString ?? throw new ArgumentNullException(nameof(connectionString));
        _logger = LoggerFactory.CreateLogger<DbWrapperSqlServer>();
    }

    public DbWrapperSqlServer(IConfiguration configuration, ISessionStateService? sessionState = null)
    {
        _sessionState = sessionState;
        _connectionStringOverride = ComposeConnectionString(configuration);
        _logger = LoggerFactory.CreateLogger<DbWrapperSqlServer>();
    }

    /// <summary>
    /// Build the runtime connection string by merging the non-secret base
    /// from <c>ConnectionStrings:DefaultConnection</c> in appsettings.json
    /// with the SA password from <c>Database:Password</c> (which comes
    /// from user-secrets in dev or the <c>Database__Password</c> env var
    /// in prod). See <c>memory/sa_password_secret.md</c>.
    ///
    /// Single source of truth — both Program.cs (for the static
    /// <see cref="DefaultConnectionString"/>) and the DI-resolved
    /// <see cref="DbWrapperSqlServer(IConfiguration)"/> constructor call
    /// here so a misalignment between the two can't cause a partial
    /// "Login failed for user 'sa'" outage.
    /// </summary>
    public static string ComposeConnectionString(IConfiguration configuration)
    {
        var baseConn = configuration.GetConnectionString("DefaultConnection")
            ?? throw new InvalidOperationException(
                "ConnectionStrings:DefaultConnection not found in appsettings.json.");

        var builder = new SqlConnectionStringBuilder(baseConn);

        // Integrated Security (a.k.a. Trusted Connection / Windows
        // Authentication) means the app authenticates as the Windows
        // identity it runs under — no SQL login + password involved.
        // Skip the Database:Password lookup entirely; SQL Server will
        // use the process's Windows token. The DBA must have configured
        // that token's account as a SQL Server login with the right perms.
        if (builder.IntegratedSecurity)
        {
            return builder.ConnectionString;
        }

        // SQL Server authentication path — password required.
        var sqlPwd = configuration["Database:Password"]
            ?? throw new InvalidOperationException(
                "Database:Password not configured. In dev: " +
                "`dotnet user-secrets set \"Database:Password\" \"<sa-password>\"` " +
                "from the project folder. In production: set env var " +
                "`Database__Password`. Alternatively, switch the connection " +
                "string to use `Integrated Security=True` and remove the " +
                "`User Id=...` clause to authenticate as the Windows account " +
                "running the app.");
        builder.Password = sqlPwd;
        return builder.ConnectionString;
    }

    // Matches a single SQL identifier: a leading letter or underscore, then
    // letters / digits / underscores. No brackets, quotes, spaces, semicolons
    // or other punctuation — so a validated value can be safely interpolated as
    // a table or column name where parameterization isn't possible (SQL Server
    // doesn't allow parameter placeholders for identifiers).
    private static readonly Regex _identifierRegex =
        new("^[A-Za-z_][A-Za-z0-9_]*$", RegexOptions.Compiled);

    /// <summary>
    /// Guards a SQL identifier (table or column name) that must be concatenated
    /// into a SQL string because SQL Server does not allow parameter placeholders
    /// for identifiers. Throws <see cref="ArgumentException"/> unless the value is
    /// a simple identifier (<c>^[A-Za-z_][A-Za-z0-9_]*$</c>); when
    /// <paramref name="allowSchema"/> is true a single dotted <c>schema.table</c>
    /// form is also accepted (each part validated independently). Returns the
    /// identifier unchanged so callers can write
    /// <c>$"... {db.ValidateIdentifier(name)} ..."</c> inline.
    /// </summary>
    public static string ValidateIdentifier(string id, bool allowSchema = false)
    {
        if (string.IsNullOrWhiteSpace(id))
            throw new ArgumentException("SQL identifier cannot be null or empty.", nameof(id));

        if (allowSchema)
        {
            var parts = id.Split('.');
            if (parts.Length > 2)
                throw new ArgumentException(
                    $"Invalid SQL identifier '{id}': at most one schema-qualifying dot is allowed.",
                    nameof(id));
            foreach (var part in parts)
            {
                if (!_identifierRegex.IsMatch(part))
                    throw new ArgumentException(
                        $"Invalid SQL identifier '{id}': each part must match [A-Za-z_][A-Za-z0-9_]*.",
                        nameof(id));
            }
            return id;
        }

        if (!_identifierRegex.IsMatch(id))
            throw new ArgumentException(
                $"Invalid SQL identifier '{id}': must match [A-Za-z_][A-Za-z0-9_]*.",
                nameof(id));

        return id;
    }

    private SqlConnection GetConnection() => new(GetConnectionString());

    public string GetConnectionString()
    {
        if (_connectionStringOverride != null)
            return _connectionStringOverride;

        if (_sessionState != null && !string.IsNullOrEmpty(_sessionState.DatabaseName))
        {
            var builder = new SqlConnectionStringBuilder(DefaultConnectionString)
            {
                InitialCatalog = _sessionState.DatabaseName
            };
            return builder.ConnectionString;
        }

        return DefaultConnectionString;
    }

    public DataTable SqlExec(string sql, object? parameters = null)
    {
        var sw = Stopwatch.StartNew();
        try
        {
            using var connection = GetConnection();
            connection.Open();
            using var command = new SqlCommand(sql, connection);
            AddParameters(command, parameters);
            var dt = new DataTable();
            using var adapter = new SqlDataAdapter(command);
            adapter.Fill(dt);

            LogSqlSuccess(nameof(SqlExec), sql, sw.ElapsedMilliseconds, dt.Rows.Count);
            return dt;
        }
        catch (Exception ex)
        {
            LogSqlError(nameof(SqlExec), sql, sw.ElapsedMilliseconds, ex);
            throw;
        }
    }

    public async Task<List<Dictionary<string, object>>> QueryAsync(string sql, object? parameters = null)
    {
        var sw = Stopwatch.StartNew();
        try
        {
            var results = new List<Dictionary<string, object>>();
            await using var connection = GetConnection();
            await connection.OpenAsync();
            await using var command = new SqlCommand(sql, connection);
            AddParameters(command, parameters);
            await using var reader = await command.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                var row = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    var name = reader.GetName(i);
                    var value = reader.IsDBNull(i) ? null! : reader.GetValue(i);
                    row[name] = value;
                }
                results.Add(row);
            }

            LogSqlSuccess(nameof(QueryAsync), sql, sw.ElapsedMilliseconds, results.Count);
            return results;
        }
        catch (Exception ex)
        {
            LogSqlError(nameof(QueryAsync), sql, sw.ElapsedMilliseconds, ex);
            throw;
        }
    }

    public async Task<List<ExpandoObject>> QueryDynamicAsync(string sql, object? parameters = null)
    {
        var sw = Stopwatch.StartNew();
        try
        {
            var results = new List<ExpandoObject>();
            await using var connection = GetConnection();
            await connection.OpenAsync();
            await using var command = new SqlCommand(sql, connection);
            AddParameters(command, parameters);
            await using var reader = await command.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                var row = new ExpandoObject();
                var dict = (IDictionary<string, object>)row;
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    var columnName = reader.GetName(i);
                    var value = reader.IsDBNull(i) ? null! : reader.GetValue(i);
                    dict[columnName] = value;
                }
                results.Add(row);
            }

            LogSqlSuccess(nameof(QueryDynamicAsync), sql, sw.ElapsedMilliseconds, results.Count);
            return results;
        }
        catch (Exception ex)
        {
            LogSqlError(nameof(QueryDynamicAsync), sql, sw.ElapsedMilliseconds, ex);
            throw;
        }
    }

    public async Task<int> ExecuteAsync(string sql, object? parameters = null)
    {
        var sw = Stopwatch.StartNew();
        try
        {
            await using var connection = GetConnection();
            await connection.OpenAsync();
            await using var command = new SqlCommand(sql, connection);
            AddParameters(command, parameters);
            var affected = await command.ExecuteNonQueryAsync();
            LogSqlSuccess(nameof(ExecuteAsync), sql, sw.ElapsedMilliseconds, affected);
            return affected;
        }
        catch (Exception ex)
        {
            LogSqlError(nameof(ExecuteAsync), sql, sw.ElapsedMilliseconds, ex);
            throw;
        }
    }

    public long ExecuteNonQuery(string sql, object? parameters = null)
    {
        var sw = Stopwatch.StartNew();
        try
        {
            using var connection = GetConnection();
            connection.Open();
            using var command = new SqlCommand(sql, connection);
            AddParameters(command, parameters);
            var affected = command.ExecuteNonQuery();
            LogSqlSuccess(nameof(ExecuteNonQuery), sql, sw.ElapsedMilliseconds, affected);
            return affected;
        }
        catch (Exception ex)
        {
            LogSqlError(nameof(ExecuteNonQuery), sql, sw.ElapsedMilliseconds, ex);
            throw;
        }
    }

    public async Task<T?> ExecuteScalarAsync<T>(string sql, object? parameters = null)
    {
        var sw = Stopwatch.StartNew();
        try
        {
            await using var connection = GetConnection();
            await connection.OpenAsync();
            await using var command = new SqlCommand(sql, connection);
            AddParameters(command, parameters);
            var result = await command.ExecuteScalarAsync();
            if (result == null || result == DBNull.Value)
            {
                LogSqlSuccess(nameof(ExecuteScalarAsync), sql, sw.ElapsedMilliseconds, 0);
                return default;
            }

            LogSqlSuccess(nameof(ExecuteScalarAsync), sql, sw.ElapsedMilliseconds, 1);
            // Unwrap Nullable<T>: Convert.ChangeType can't target a Nullable type (it
            // isn't IConvertible), so ExecuteScalar<int?> threw "Invalid cast Int32 →
            // Nullable<Int32>". Convert to the underlying type; the (T) cast re-wraps it.
            return (T)Convert.ChangeType(result, Nullable.GetUnderlyingType(typeof(T)) ?? typeof(T));
        }
        catch (Exception ex)
        {
            LogSqlError(nameof(ExecuteScalarAsync), sql, sw.ElapsedMilliseconds, ex);
            throw;
        }
    }

    public T? ExecuteScalar<T>(string sql, object? parameters = null)
    {
        var sw = Stopwatch.StartNew();
        try
        {
            using var connection = GetConnection();
            connection.Open();
            using var command = new SqlCommand(sql, connection);
            AddParameters(command, parameters);
            var result = command.ExecuteScalar();
            if (result == null || result == DBNull.Value)
            {
                LogSqlSuccess(nameof(ExecuteScalar), sql, sw.ElapsedMilliseconds, 0);
                return default;
            }

            LogSqlSuccess(nameof(ExecuteScalar), sql, sw.ElapsedMilliseconds, 1);
            // Unwrap Nullable<T>: Convert.ChangeType can't target a Nullable type (it
            // isn't IConvertible), so ExecuteScalar<int?> threw "Invalid cast Int32 →
            // Nullable<Int32>". Convert to the underlying type; the (T) cast re-wraps it.
            return (T)Convert.ChangeType(result, Nullable.GetUnderlyingType(typeof(T)) ?? typeof(T));
        }
        catch (Exception ex)
        {
            LogSqlError(nameof(ExecuteScalar), sql, sw.ElapsedMilliseconds, ex);
            throw;
        }
    }

    /// <summary>Open one connection + transaction spanning a multi-statement
    /// save. Statements issued through the returned session all share the
    /// transaction: CommitAsync makes them permanent, and disposing the
    /// session without committing rolls every one of them back. Added for
    /// FDBGrid.SaveData (data incident 2026-08-01: a save that aborted
    /// mid-loop left the UPDATEs already executed committed — a partial
    /// save that had to be restored by hand).</summary>
    public async Task<DbTransactionSession> BeginTransactionAsync()
    {
        var connection = GetConnection();
        try
        {
            await connection.OpenAsync();
            var transaction = (SqlTransaction)await connection.BeginTransactionAsync();
            return new DbTransactionSession(this, connection, transaction);
        }
        catch
        {
            await connection.DisposeAsync();
            throw;
        }
    }

    /// <summary>A single open connection + transaction from
    /// <see cref="BeginTransactionAsync"/>. Nested so it can reuse the
    /// wrapper's parameter binding and logging.</summary>
    public sealed class DbTransactionSession : IAsyncDisposable
    {
        private readonly DbWrapperSqlServer _owner;
        private readonly SqlConnection _connection;
        private readonly SqlTransaction _transaction;
        private bool _committed;

        internal DbTransactionSession(DbWrapperSqlServer owner, SqlConnection connection, SqlTransaction transaction)
        {
            _owner = owner;
            _connection = connection;
            _transaction = transaction;
        }

        public async Task<int> ExecuteAsync(string sql, object? parameters = null)
        {
            var sw = Stopwatch.StartNew();
            try
            {
                await using var command = new SqlCommand(sql, _connection, _transaction);
                _owner.AddParameters(command, parameters);
                var affected = await command.ExecuteNonQueryAsync();
                _owner.LogSqlSuccess(nameof(ExecuteAsync), sql, sw.ElapsedMilliseconds, affected);
                return affected;
            }
            catch (Exception ex)
            {
                _owner.LogSqlError(nameof(ExecuteAsync), sql, sw.ElapsedMilliseconds, ex);
                throw;
            }
        }

        /// <summary>Stream a whole DataTable into <paramref name="destinationTable"/>
        /// with SqlBulkCopy, enlisted in THIS session's transaction so it rolls
        /// back with everything else. Sends no parameters at all, so neither the
        /// 2 100-parameter cap nor the 1 000-row VALUES limit applies — the
        /// import can go in one pass instead of hundreds of chunked INSERTs.
        /// Column mapping is by NAME: the DataTable's column names must match
        /// the destination's.</summary>
        public async Task<int> BulkCopyAsync(DataTable table, string destinationTable, int timeoutSeconds = 300)
        {
            if (table.Rows.Count == 0) return 0;

            var sw = Stopwatch.StartNew();
            var label = $"SqlBulkCopy -> {destinationTable} ({table.Rows.Count} rows)";
            try
            {
                using var bulk = new SqlBulkCopy(_connection, SqlBulkCopyOptions.Default, _transaction)
                {
                    DestinationTableName = destinationTable,
                    BulkCopyTimeout = timeoutSeconds,
                    BatchSize = 5000,
                };
                foreach (DataColumn c in table.Columns)
                    bulk.ColumnMappings.Add(c.ColumnName, c.ColumnName);

                await bulk.WriteToServerAsync(table);

                _owner.LogSqlSuccess(nameof(BulkCopyAsync), label, sw.ElapsedMilliseconds, table.Rows.Count);
                return table.Rows.Count;
            }
            catch (Exception ex)
            {
                _owner.LogSqlError(nameof(BulkCopyAsync), label, sw.ElapsedMilliseconds, ex);
                throw;
            }
        }

        public async Task<T?> ExecuteScalarAsync<T>(string sql, object? parameters = null)
        {
            var sw = Stopwatch.StartNew();
            try
            {
                await using var command = new SqlCommand(sql, _connection, _transaction);
                _owner.AddParameters(command, parameters);
                var result = await command.ExecuteScalarAsync();
                if (result == null || result == DBNull.Value)
                {
                    _owner.LogSqlSuccess(nameof(ExecuteScalarAsync), sql, sw.ElapsedMilliseconds, 0);
                    return default;
                }

                _owner.LogSqlSuccess(nameof(ExecuteScalarAsync), sql, sw.ElapsedMilliseconds, 1);
                // Unwrap Nullable<T>: Convert.ChangeType can't target a Nullable type (it
                // isn't IConvertible), so ExecuteScalar<int?> threw "Invalid cast Int32 →
                // Nullable<Int32>". Convert to the underlying type; the (T) cast re-wraps it.
                return (T)Convert.ChangeType(result, Nullable.GetUnderlyingType(typeof(T)) ?? typeof(T));
            }
            catch (Exception ex)
            {
                _owner.LogSqlError(nameof(ExecuteScalarAsync), sql, sw.ElapsedMilliseconds, ex);
                throw;
            }
        }

        public async Task CommitAsync()
        {
            await _transaction.CommitAsync();
            _committed = true;
        }

        public async ValueTask DisposeAsync()
        {
            if (!_committed)
            {
                try
                {
                    await _transaction.RollbackAsync();
                }
                catch
                {
                    // Connection already broken — SQL Server rolls the open
                    // transaction back server-side when the session drops.
                }
            }
            await _transaction.DisposeAsync();
            await _connection.DisposeAsync();
        }
    }

    private void AddParameters(SqlCommand command, object? parameters)
    {
        if (parameters == null)
            return;

        // Support IDictionary<string, object?> (used by ReportWriterControl / CrystalXmlReportLoader
        // which produce parameter bags from XML report definitions).
        if (parameters is System.Collections.Generic.IDictionary<string, object?> dict)
        {
            foreach (var kv in dict)
            {
                var value = kv.Value ?? DBNull.Value;
                var paramName = kv.Key.StartsWith("@", StringComparison.Ordinal) ? kv.Key : "@" + kv.Key;
                command.Parameters.AddWithValue(paramName, value);
            }
            return;
        }
        if (parameters is System.Collections.Generic.IDictionary<string, object> dict2)
        {
            foreach (var kv in dict2)
            {
                var value = kv.Value ?? DBNull.Value;
                var paramName = kv.Key.StartsWith("@", StringComparison.Ordinal) ? kv.Key : "@" + kv.Key;
                command.Parameters.AddWithValue(paramName, value);
            }
            return;
        }

        foreach (var prop in parameters.GetType().GetProperties())
        {
            var value = prop.GetValue(parameters) ?? DBNull.Value;
            var paramName = prop.Name.StartsWith("@", StringComparison.Ordinal) ? prop.Name : "@" + prop.Name;
            command.Parameters.AddWithValue(paramName, value);
        }
    }

    private void LogSqlSuccess(string operation, string sql, long elapsedMs, int? rowCount)
    {
        // Slow-query warnings only. Per-query Info-level success logging
        // was the main log-clog source — every successful SQL call wrote
        // an Info line with full SQL text. Errors continue to flow
        // through LogSqlError; routine successful commands stay silent.
        var slowThreshold = Math.Max(0, LoggingOptions.SlowQueryThresholdMs);
        if (slowThreshold <= 0 || elapsedMs < slowThreshold)
            return;

        if (LoggingOptions.LogSqlText)
        {
            _logger.LogWarning(
                "{Operation} completed slowly in {ElapsedMs}ms. RowCount={RowCount}. SQL={SqlText}",
                operation, elapsedMs, rowCount, SummarizeSql(sql));
        }
        else
        {
            _logger.LogWarning(
                "{Operation} completed slowly in {ElapsedMs}ms. RowCount={RowCount}.",
                operation, elapsedMs, rowCount);
        }
    }

    private void LogSqlError(string operation, string sql, long elapsedMs, Exception ex)
    {
        if (LoggingOptions.LogSqlText)
        {
            _logger.LogError(
                ex,
                "{Operation} failed after {ElapsedMs}ms. SQL={SqlText}",
                operation,
                elapsedMs,
                SummarizeSql(sql));
        }
        else
        {
            _logger.LogError(ex, "{Operation} failed after {ElapsedMs}ms.", operation, elapsedMs);
        }
    }

    private static string SummarizeSql(string sql)
    {
        if (string.IsNullOrWhiteSpace(sql))
            return string.Empty;

        var condensed = Regex.Replace(sql, "\\s+", " ").Trim();
        const int maxLen = 600;
        if (condensed.Length <= maxLen)
            return condensed;

        return condensed[..maxLen] + " ...";
    }
}

public class ExportRequestPayload
{
    public string Sql { get; set; } = "";
    public string FileName { get; set; } = "";
}
