namespace BlazorDemos.Pages.Grid;
using Microsoft.Extensions.Configuration;
using Microsoft.Data.SqlClient;
using System.Data;
using System.Dynamic;
public class DbWrapperSqlServer {
    private string _connectionString;

    public DbWrapperSqlServer(){
        _connectionString = "Server=localhost,1433;Database=HomeFrontDB;User Id=sa;Password=Nusr@t7860;TrustServerCertificate=true;";
    }
    
    // Inject via IConfiguration or set directly
    public DbWrapperSqlServer(string connectionString){
        _connectionString = connectionString ?? throw new ArgumentNullException(nameof(connectionString));
    }

    // Constructor for dependency injection (recommended)
    public DbWrapperSqlServer(IConfiguration configuration){
        _connectionString = configuration.GetConnectionString("DefaultConnection")
                            ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
    }
    private SqlConnection GetConnection() => new(_connectionString);

    // ─────────────────────────────────────────────────────────────
    // 1. Classic DataTable (perfect for Syncfusion Grid binding)
    // ─────────────────────────────────────────────────────────────
    public DataTable SqlExec(string sql, object? parameters = null){
        Console.WriteLine($"SQL (Grid) = {sql}");
        using var connection = GetConnection();
        connection.Open();
        using var command = new SqlCommand(sql, connection);
        AddParameters(command, parameters);
        var dt = new DataTable();
        using var adapter = new SqlDataAdapter(command);
        adapter.Fill(dt);
        Console.WriteLine($"Row Count = {dt.Rows.Count}");
        return dt;
    }

    // ─────────────────────────────────────────────────────────────
    // 2. Async List<Dictionary<string, object>> (flexible, dynamic)
    // ─────────────────────────────────────────────────────────────
    public async Task<List<Dictionary<string, object>>> QueryAsync(string sql, object? parameters = null){
        Console.WriteLine($"QueryAsSync = {sql}");
        var results = new List<Dictionary<string, object>>();
        await using var connection = GetConnection();
        await connection.OpenAsync();
        await using var command = new SqlCommand(sql, connection);
        AddParameters(command, parameters);
        await using var reader = await command.ExecuteReaderAsync();
        int RowCount = 0;
        while (await reader.ReadAsync()){
            var row = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < reader.FieldCount; i++){
                var name = reader.GetName(i);
                var value = reader.IsDBNull(i) ? null : reader.GetValue(i);
                row[name] = value;
            }
            RowCount++;
            results.Add(row);
        }
        
        Console.WriteLine($"Row Count = {results.Count}");
        return results;
    }

    // ─────────────────────────────────────────────────────────────
    // 3. Async ExpandoObject (great for dynamic JSON-like objects)
    // ─────────────────────────────────────────────────────────────
    public async Task<List<ExpandoObject>> QueryDynamicAsync(string sql, object? parameters = null){
        var results = new List<ExpandoObject>();
        await using var connection = GetConnection();
        await connection.OpenAsync();
        await using var command = new SqlCommand(sql, connection);
        AddParameters(command, parameters);
        await using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync()){
            var row = new ExpandoObject();
            var dict = (IDictionary<string, object>)row;
            for (int i = 0; i < reader.FieldCount; i++){
                var columnName = reader.GetName(i);
                var value = reader.IsDBNull(i) ? null : reader.GetValue(i);
                dict[columnName] = value;
            }
            results.Add(row);
        }
        return results;
    }

    // ─────────────────────────────────────────────────────────────
    // 4. ExecuteNonQuery (INSERT, UPDATE, DELETE)
    // ─────────────────────────────────────────────────────────────
    public async Task<int> ExecuteAsync(string sql, object? parameters = null){
        await using var connection = GetConnection();
        await connection.OpenAsync();
        await using var command = new SqlCommand(sql, connection);
        AddParameters(command, parameters);
        return await command.ExecuteNonQueryAsync();
    }

    // ─────────────────────────────────────────────────────────────
    // 5. Legacy sync ExecuteNonQuery (for old code)
    // ─────────────────────────────────────────────────────────────
    public long ExecuteNonQuery(string sql, object? parameters = null){
        using var connection = GetConnection();
        connection.Open();
        using var command = new SqlCommand(sql, connection);
        AddParameters(command, parameters);
        return command.ExecuteNonQuery();
    }

    // ─────────────────────────────────────────────────────────────
    // 6. ExecuteScalar (e.g., for your Purch_GetItemRate function)
    // ─────────────────────────────────────────────────────────────
    public async Task<T?> ExecuteScalarAsync<T>(string sql, object? parameters = null){
        await using var connection = GetConnection();
        await connection.OpenAsync();
        await using var command = new SqlCommand(sql, connection);
        Console.WriteLine($"QueryAsSync = {sql}");
        AddParameters(command, parameters);
        var result = await command.ExecuteScalarAsync();
        if (result == null || result == DBNull.Value)
            return default;
        return (T)Convert.ChangeType(result, typeof(T));
    }
    public T? ExecuteScalar<T>(string sql, object? parameters = null){
        using var connection = GetConnection();
        connection.Open();
        using var command = new SqlCommand(sql, connection);
        AddParameters(command, parameters);
        var result = command.ExecuteScalar();
        if (result == null || result == DBNull.Value)
            return default;
        return (T)Convert.ChangeType(result, typeof(T));
    }

    // ─────────────────────────────────────────────────────────────
    // Helper: Add parameters from anonymous object
    // ─────────────────────────────────────────────────────────────
    private void AddParameters(SqlCommand command, object? parameters){
        if (parameters == null) return;
        foreach (var prop in parameters.GetType().GetProperties()){
            var value = prop.GetValue(parameters) ?? DBNull.Value;
            var paramName = prop.Name.StartsWith("@") ? prop.Name : "@" + prop.Name;
            Console.WriteLine($"{paramName} = {value}");
            command.Parameters.AddWithValue(paramName, value);
        }
    }
}