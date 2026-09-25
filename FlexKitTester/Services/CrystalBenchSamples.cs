using System.Data;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Fx.ControlKit.Reports;
using Microsoft.Data.Sqlite;

namespace FlexKitTester.Services;

public sealed record CrystalSampleIssue(string Category, string Message);
public sealed record CrystalSampleReport(string Hash, string Name, string RelativePath, string Collection,
    string Status, int Pages, int Rows, List<CrystalSampleIssue> Issues, int Id = 0);
public sealed record CrystalSampleColumn(string Name, string Type, string Reference);
public sealed record CrystalSampleDataset(string Key, List<CrystalSampleColumn> Columns,
    Dictionary<string, string> Parameters, List<JsonElement[]> Rows);
public sealed record CrystalSampleCapture(CrystalSampleReport Report, List<CrystalSampleDataset> Datasets);

/// <summary>Reads the offline sample pack. Report SQL is never sent to SQLite or an external database.</summary>
public sealed class CrystalBenchSamples
{
    public string DatabasePath { get; }
    public bool Available => File.Exists(DatabasePath);
    public CrystalBenchSamples(IConfiguration configuration, IWebHostEnvironment environment)
        : this(Path.GetFullPath(configuration["CrystalReportsBench:SampleDatabase"] ?? "Data/CrystalSamples.db", environment.ContentRootPath)) { }
    public CrystalBenchSamples(string path) => DatabasePath = Path.GetFullPath(path);

    public IReadOnlyList<CrystalSampleReport> ReadCatalog()
    {
        if (!Available) return [];
        using var connection = Open();
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT Metadata FROM ReportCatalog ORDER BY Name COLLATE NOCASE, Hash LIMIT 10000";
        using var reader = command.ExecuteReader();
        var reports = new List<CrystalSampleReport>();
        while (reader.Read()) reports.Add(JsonSerializer.Deserialize<CrystalSampleReport>(reader.GetString(0))!);
        if (reports.Any(r => r.Id <= 0) || reports.Select(r => r.Id).Distinct().Count() != reports.Count)
            throw new InvalidDataException("Sample catalog needs unique report IDs. Run the offline catalog command before browsing this pack.");
        return reports.OrderBy(r => r.Id).ToArray();
    }

    public Dictionary<string, string> Parameters(string hash, ReportDefinition definition) => ReadDataset(hash, definition).Parameters;

    public DataTable Execute(string hash, ReportDefinition definition)
    {
        var dataset = ReadDataset(hash, definition);
        using var connection = Open();
        using var command = connection.CreateCommand();
        command.CommandText = $"SELECT * FROM {TableName(hash, dataset.Key)} ORDER BY rowid LIMIT 10001";
        using var reader = command.ExecuteReader();
        var table = new DataTable(definition.ReportId) { Locale = CultureInfo.InvariantCulture };
        foreach (var column in dataset.Columns) table.Columns.Add(column.Name, ClrType(column.Type));
        while (reader.Read())
        {
            if (table.Rows.Count >= CrystalBenchDataExecutor.MaxRows) throw new InvalidDataException("Sample dataset exceeds the row limit.");
            var values = dataset.Columns.Select((c, i) => reader.IsDBNull(i) ? (object)DBNull.Value : Parse(reader.GetValue(i), c.Type)).ToArray();
            table.Rows.Add(values);
        }
        return table;
    }

    private CrystalSampleDataset ReadDataset(string hash, ReportDefinition definition)
    {
        ValidateHash(hash);
        using var connection = Open();
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT Metadata FROM SampleDatasets WHERE ReportHash=$hash AND DatasetKey=$key";
        command.Parameters.AddWithValue("$hash", hash);
        command.Parameters.AddWithValue("$key", Key(definition));
        var json = command.ExecuteScalar() as string ?? throw new InvalidDataException(
            "No matching SQLite sample schema. Generate a sample pack for this RPT or load a matching data fixture; no substitute data was used.");
        return JsonSerializer.Deserialize<CrystalSampleDataset>(json)!;
    }

    private SqliteConnection Open()
    {
        if (!Available) throw new FileNotFoundException("Generate the offline Crystal sample database first.", DatabasePath);
        var connection = new SqliteConnection(new SqliteConnectionStringBuilder { DataSource = DatabasePath, Mode = SqliteOpenMode.ReadOnly, Pooling = false }.ToString());
        try
        {
            connection.Open();
            using var command = connection.CreateCommand();
            command.CommandText = "PRAGMA query_only=ON; PRAGMA trusted_schema=OFF;";
            command.ExecuteNonQuery();
            return connection;
        }
        catch { connection.Dispose(); throw; }
    }

    public static string Key(ReportDefinition definition) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(
        (definition.PositionedLayout is { } layout ? JsonSerializer.Serialize(new { layout.Document.CustomSql,
            Tables = layout.Document.DataSources.Select(t => new { t.Name, t.Schema, t.SourceName, t.CommandText }),
            Links = layout.Document.Links.Select(l => new { l.LeftTable, l.LeftField, l.RightTable, l.RightField, l.JoinType }),
            Fields = layout.Document.Fields.Where(f => !f.IsFormula).Select(f => new { f.Reference, f.Type }) }) : definition.Sql)
        + "\n" + string.Join("\n", definition.PositionedLayout?.Bindings.OrderBy(x => x.Key, StringComparer.Ordinal)
            .Select(x => x.Key + "=" + x.Value) ?? CrystalBenchDataExecutor.RequiredColumns(definition)))));
    public static string TableName(string hash, string key) { ValidateHash(hash); ValidateHash(key); return "sample_" + hash + "_" + key; }
    private static void ValidateHash(string hash)
    {
        if (hash.Length != 64 || hash.Any(c => !char.IsAsciiHexDigit(c))) throw new InvalidDataException("Invalid sample fingerprint.");
    }
    public static Type ClrType(string type) => type switch
    {
        "integer" => typeof(long), "decimal" => typeof(decimal), "number" => typeof(double), "boolean" => typeof(bool),
        "datetime" => typeof(DateTime), "string" => typeof(string), _ => throw new InvalidDataException("Unknown sample column type: " + type)
    };
    public static object Parse(object value, string type) => type switch
    {
        "integer" => checked((long)decimal.Parse(Convert.ToString(value, CultureInfo.InvariantCulture)!, CultureInfo.InvariantCulture)),
        "boolean" => Convert.ToString(value, CultureInfo.InvariantCulture) is "1" or "0" ? Convert.ToString(value, CultureInfo.InvariantCulture) == "1" : bool.Parse(Convert.ToString(value, CultureInfo.InvariantCulture)!),
        "datetime" => DateTime.Parse(Convert.ToString(value, CultureInfo.InvariantCulture)!, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind),
        _ => Convert.ChangeType(value, ClrType(type), CultureInfo.InvariantCulture)
    };
}
