using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using System.Data;
using System.Data.SqlTypes;
using System.Dynamic;

namespace HomeFront.Data;

public class DbWrapperSqlServer
{
    private readonly string _connectionString;
    private bool debug_mode=false;

    /// <summary>
    /// Default connection string read from appsettings.json at startup.
    /// Set once in Program.cs via DbWrapperSqlServer.DefaultConnectionString = ...
    /// </summary>
    public static string DefaultConnectionString { get; set; } = "";

    public DbWrapperSqlServer()
    {
        if (string.IsNullOrEmpty(DefaultConnectionString))
            throw new InvalidOperationException(
                "DbWrapperSqlServer.DefaultConnectionString has not been configured. " +
                "Set it from appsettings.json in Program.cs.");
        _connectionString = DefaultConnectionString;
    }

    public DbWrapperSqlServer(string connectionString)
    {
        _connectionString = connectionString ?? throw new ArgumentNullException(nameof(connectionString));
    }

    public DbWrapperSqlServer(IConfiguration configuration)
    {
        _connectionString = configuration.GetConnectionString("DefaultConnection")
                            ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
    }

    private SqlConnection GetConnection() => new(_connectionString);

    public string GetConnectionString() => _connectionString;

    public DataTable SqlExec(string sql, object? parameters = null)
    {
        if (debug_mode) Console.WriteLine($"SQL (Grid) = {sql}");
        using var connection = GetConnection();
        connection.Open();
        using var command = new SqlCommand(sql, connection);
        AddParameters(command, parameters);
        var dt = new DataTable();
        using var adapter = new SqlDataAdapter(command);
        adapter.Fill(dt);
        if (debug_mode) Console.WriteLine($"Row Count = {dt.Rows.Count}");
        return dt;
    }

    public async Task<List<Dictionary<string, object>>> QueryAsync(string sql, object? parameters = null)
    {
        if (debug_mode) Console.WriteLine($"QueryAsSync = {sql}");
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
                var value = reader.IsDBNull(i) ? null : reader.GetValue(i);
                row[name] = value;
            }
            results.Add(row);
        }

        if (debug_mode) Console.WriteLine($"Row Count = {results.Count}");
        return results;
    }

    public async Task<List<ExpandoObject>> QueryDynamicAsync(string sql, object? parameters = null)
    {
        if (debug_mode) Console.WriteLine($"QueryDynamicAsync SQL = {sql}");
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
                var value = reader.IsDBNull(i) ? null : reader.GetValue(i);
                dict[columnName] = value;
            }
            results.Add(row);
        }
        if (debug_mode) Console.WriteLine($"Row Count = {results.Count}");
        return results;
    }

    public async Task<int> ExecuteAsync(string sql, object? parameters = null)
    {
        await using var connection = GetConnection();
        await connection.OpenAsync();
        await using var command = new SqlCommand(sql, connection);
        AddParameters(command, parameters);
        return await command.ExecuteNonQueryAsync();
    }

    public long ExecuteNonQuery(string sql, object? parameters = null)
    {
        using var connection = GetConnection();
        connection.Open();
        using var command = new SqlCommand(sql, connection);
        AddParameters(command, parameters);
        return command.ExecuteNonQuery();
    }

    public async Task<T?> ExecuteScalarAsync<T>(string sql, object? parameters = null)
    {
        await using var connection = GetConnection();
        await connection.OpenAsync();
        await using var command = new SqlCommand(sql, connection);
        if (debug_mode) Console.WriteLine($"QueryAsSync = {sql}");
        AddParameters(command, parameters);
        var result = await command.ExecuteScalarAsync();
        if (result == null || result == DBNull.Value)
        {
            return default;
        }

        return (T)Convert.ChangeType(result, typeof(T));
    }

    public T? ExecuteScalar<T>(string sql, object? parameters = null)
    {
        using var connection = GetConnection();
        connection.Open();
        using var command = new SqlCommand(sql, connection);
        AddParameters(command, parameters);
        var result = command.ExecuteScalar();
        if (result == null || result == DBNull.Value)
        {
            return default;
        }

        return (T)Convert.ChangeType(result, typeof(T));
    }

    private void AddParameters(SqlCommand command, object? parameters)
    {
        if (parameters == null)
        {
            return;
        }

        foreach (var prop in parameters.GetType().GetProperties())
        {
            var value = prop.GetValue(parameters) ?? DBNull.Value;
            var paramName = prop.Name.StartsWith("@", StringComparison.Ordinal) ? prop.Name : "@" + prop.Name;
            if (debug_mode) Console.WriteLine($"{paramName} = {value}");
            command.Parameters.AddWithValue(paramName, value);
        }
    }
}
