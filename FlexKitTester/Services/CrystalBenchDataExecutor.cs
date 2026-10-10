using System.Data;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Fx.ControlKit.Reports;
using Microsoft.Data.SqlClient;
using Microsoft.SqlServer.TransactSql.ScriptDom;

namespace FlexKitTester.Services;

public enum CrystalBenchDataMode { None, LayoutOnly, Fixture, Database, SyntheticSqlite }

/// <summary>Circuit-local data choice. No database operation occurs until a report is explicitly run.</summary>
public sealed class CrystalBenchDataExecutor(IConfiguration configuration, CrystalBenchSamples? samples = null) : IReportDefinitionDataExecutor
{
    public const int MaxRows = 10000;
    public CrystalBenchDataMode Mode { get; set; }
    public string? SampleReportHash { get; set; }
    public string? FixtureName { get; private set; }
    public bool DatabaseAvailable => configuration.GetValue<bool>("CrystalReportsBench:EnableDatabase")
        && !string.IsNullOrWhiteSpace(configuration.GetConnectionString("CrystalReportsBench"));
    private List<FixtureResult> _fixtures = [];
    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true, WriteIndented = true };

    public sealed record FixtureColumn(string Name, string Type);
    public sealed record FixtureResult(string ReportId, string SqlSha256, Dictionary<string, JsonElement> Parameters,
        List<FixtureColumn> Columns, List<JsonElement[]> Rows);
    public sealed record FixtureFile(List<FixtureResult> Reports);

    public void LoadFixture(string name, string json)
    {
        if (Encoding.UTF8.GetByteCount(json) > 16 * 1024 * 1024) throw new InvalidDataException("Fixtures are limited to 16 MB.");
        var file = JsonSerializer.Deserialize<FixtureFile>(json, JsonOptions) ?? throw new InvalidDataException("Empty fixture.");
        if (file.Reports is null || file.Reports.Count is < 1 or > 256) throw new InvalidDataException("Provide 1 to 256 report result sets.");
        foreach (var result in file.Reports)
        {
            if (result is null || string.IsNullOrWhiteSpace(result.ReportId) || result.SqlSha256?.Length != 64 || result.Parameters is null
                || result.Columns is null || result.Rows is null || result.Columns.Count is < 1 or > 2048 || result.Rows.Count > MaxRows)
                throw new InvalidDataException("Invalid fixture report, query fingerprint, parameters, columns, or row limit.");
            if (result.Parameters.Keys.Select(k => k.TrimStart('@')).Distinct(StringComparer.OrdinalIgnoreCase).Count() != result.Parameters.Count)
                throw new InvalidDataException("Fixture parameter names must be unique.");
            foreach (var value in result.Parameters.Values) _ = Scalar(value);
            _ = Materialize(result);
        }
        _fixtures = file.Reports;
        FixtureName = Path.GetFileName(name);
    }

    public DataTable Execute(ReportDefinition definition, IDictionary<string, object>? parameters)
    {
        if (Mode == CrystalBenchDataMode.Database) definition.ValidateSqlParameters(parameters);
        if (Mode == CrystalBenchDataMode.SyntheticSqlite)
            return (samples ?? throw new InvalidOperationException("SQLite sample pack is not configured.")).Execute(
                SampleReportHash ?? throw new InvalidOperationException("Select a catalog report with sample data."), definition);
        if (Mode == CrystalBenchDataMode.LayoutOnly)
        {
            var empty = new DataTable(definition.ReportId);
            foreach (var column in RequiredColumns(definition)) empty.Columns.Add(column, typeof(object));
            return empty;
        }
        if (Mode == CrystalBenchDataMode.Database) return Execute(definition.Sql, parameters);
        if (Mode != CrystalBenchDataMode.Fixture) throw new InvalidOperationException("Select a report data source before running.");
        var matches = _fixtures.Where(f => string.Equals(f.ReportId, definition.ReportId, StringComparison.OrdinalIgnoreCase)
            && string.Equals(f.SqlSha256, QueryHash(definition.Sql), StringComparison.OrdinalIgnoreCase)
            && ParametersMatch(f.Parameters, parameters, definition)).ToArray();
        if (matches.Length != 1) throw new InvalidDataException($"Fixture requires exactly one matching query and parameter set for '{definition.ReportId}'; found {matches.Length}.");
        var table = Materialize(matches[0]);
        var missing = RequiredColumns(definition).Where(c => !table.Columns.Contains(c)).ToArray();
        if (missing.Length != 0) throw new InvalidDataException($"Fixture '{definition.ReportId}' is missing columns: {string.Join(", ", missing)}.");
        return table;
    }

    public DataTable Execute(string sql, IDictionary<string, object>? parameters)
    {
        if (Mode != CrystalBenchDataMode.Database || !DatabaseAvailable)
            throw new InvalidOperationException("The read-only report database is not enabled and configured.");
        ValidateReadOnlyQuery(sql);
        using var connection = new SqlConnection(configuration.GetConnectionString("CrystalReportsBench"));
        using var command = connection.CreateCommand();
        command.CommandText = sql;
        command.CommandTimeout = 30;
        foreach (var parameter in parameters ?? new Dictionary<string, object>())
            command.Parameters.AddWithValue("@" + parameter.Key.TrimStart('@'), parameter.Value ?? DBNull.Value);
        connection.Open();
        using var reader = command.ExecuteReader(CommandBehavior.SingleResult);
        var table = new DataTable();
        for (var i = 0; i < reader.FieldCount; i++) table.Columns.Add(reader.GetName(i), reader.GetFieldType(i));
        while (reader.Read())
        {
            if (table.Rows.Count >= MaxRows) throw new InvalidDataException($"Report query exceeds {MaxRows:N0} rows. Narrow the parameters.");
            var values = new object[reader.FieldCount];
            reader.GetValues(values);
            table.Rows.Add(values);
        }
        return table;
    }

    public static void ValidateReadOnlyQuery(string sql)
    {
        if (sql.Length > 1024 * 1024) throw new InvalidDataException("Report SQL exceeds 1 MB.");
        var parsed = new TSql160Parser(true).Parse(new StringReader(sql), out var errors);
        if (errors.Count != 0 || parsed is not TSqlScript script || script.Batches.Count != 1
            || script.Batches[0].Statements.Count != 1 || script.Batches[0].Statements[0] is not SelectStatement select
            || select.Into is not null)
            throw new InvalidDataException("The bench accepts one read-only SELECT statement only.");
        parsed.Accept(new ReadOnlyVisitor());
    }

    private sealed class ReadOnlyVisitor : TSqlFragmentVisitor
    {
        public override void Visit(TSqlFragment node)
        {
            // Reject mutation/remote-source nodes inside a SELECT too (including CTEs and OUTPUT tables).
            if (node is DataModificationTableReference or NextValueForExpression or OpenRowsetTableReference
                or BulkOpenRowset or OpenRowsetCosmos or OpenQueryTableReference or AdHocTableReference
                || node is SchemaObjectName name && (name.ServerIdentifier is not null || name.DatabaseIdentifier is not null))
                throw new InvalidDataException("Report SQL cannot mutate data or access external databases.");
            base.Visit(node);
        }
    }

    public static string QueryHash(string sql) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(sql)));
    public static IEnumerable<string> RequiredColumns(ReportDefinition definition) =>
        (definition.PositionedLayout?.Bindings.Values ?? definition.Columns.Select(c => c.Field))
        .Where(c => !string.IsNullOrWhiteSpace(c)).Distinct(StringComparer.OrdinalIgnoreCase);

    private static bool ParametersMatch(Dictionary<string, JsonElement> expected, IDictionary<string, object>? actual, ReportDefinition definition)
    {
        if (expected.Count != (actual?.Count ?? 0)) return false;
        return expected.All(e => actual!.Any(a => string.Equals(a.Key.TrimStart('@'), e.Key.TrimStart('@'), StringComparison.OrdinalIgnoreCase)
            && ParameterScalar(e.Value, e.Key, definition) == ParameterScalar(JsonSerializer.SerializeToElement(a.Value is DBNull ? null : a.Value), a.Key, definition)));
    }

    private static string ParameterScalar(JsonElement value, string name, ReportDefinition definition)
    {
        var text = Scalar(value);
        var type = definition.Parameters.FirstOrDefault(p => string.Equals(p.Name.TrimStart('@'), name.TrimStart('@'), StringComparison.OrdinalIgnoreCase))?.ParameterType;
        if (type == ReportParameterType.Date && DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var date))
            return date.ToString("yyyy-MM-ddTHH:mm:ss.fffffff", CultureInfo.InvariantCulture);
        if (type is ReportParameterType.Decimal or ReportParameterType.Integer && decimal.TryParse(text, NumberStyles.Number, CultureInfo.InvariantCulture, out var number))
            return number.ToString("G29", CultureInfo.InvariantCulture);
        if (type == ReportParameterType.Boolean && bool.TryParse(text, out var boolean)) return boolean ? "true" : "false";
        return text;
    }

    private static string Scalar(JsonElement value) => value.ValueKind switch
    {
        JsonValueKind.Number => value.GetDecimal().ToString("G29", CultureInfo.InvariantCulture),
        JsonValueKind.String => value.GetString()!,
        JsonValueKind.Null => "\0null",
        JsonValueKind.True => "true",
        JsonValueKind.False => "false",
        _ => throw new InvalidDataException("Fixture parameters must be scalar values.")
    };

    private static DataTable Materialize(FixtureResult fixture)
    {
        var table = new DataTable(fixture.ReportId) { Locale = CultureInfo.InvariantCulture };
        foreach (var column in fixture.Columns)
        {
            if (string.IsNullOrWhiteSpace(column.Name) || table.Columns.Contains(column.Name)) throw new InvalidDataException("Fixture column names must be unique and nonempty.");
            var type = column.Type?.ToLowerInvariant() switch
            {
                "string" => typeof(string), "decimal" => typeof(decimal), "integer" => typeof(long),
                "number" => typeof(double), "boolean" => typeof(bool), "datetime" => typeof(DateTime),
                _ => throw new InvalidDataException($"Unsupported fixture column type '{column.Type}'.")
            };
            table.Columns.Add(column.Name, type);
        }
        foreach (var row in fixture.Rows)
        {
            if (row is null || row.Length != fixture.Columns.Count) throw new InvalidDataException("Fixture row width does not match its columns.");
            var values = row.Select((value, i) => value.ValueKind == JsonValueKind.Null ? DBNull.Value
                : JsonSerializer.Deserialize(value.GetRawText(), table.Columns[i].DataType) ?? DBNull.Value).ToArray();
            table.Rows.Add(values);
        }
        return table;
    }
}

public sealed class CrystalBenchSessionContext : IReportSessionContext
{
    public object? Get(string parameterName) => null;
}

public sealed class CrystalBenchNoRuntimeExporter : IReportExporter
{
    public bool IsLoaded => false;
    private static NotSupportedException Unavailable() => new("This bench uses only the native C# converter and viewer exports.");
    public void LoadReport(string rptFilePath) => throw Unavailable();
    public void SetParameters(Dictionary<string, string> parameters) => throw Unavailable();
    public byte[] ExportToPdf() => throw Unavailable();
    public byte[] ExportToExcel() => throw Unavailable();
    public byte[] ExportToWord() => throw Unavailable();
    public byte[] ExportToRtf() => throw Unavailable();
    public byte[] ExportToCsv() => throw Unavailable();
}
