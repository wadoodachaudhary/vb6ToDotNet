using System.Diagnostics;
using Microsoft.Data.Sqlite;

namespace FlexKitTester.Services;

/// <summary>
/// Tester-only SQLite copy of the HomeFront EstimatingItems export. The file is
/// created lazily; visiting any other bench has no SQLite startup or memory cost.
/// </summary>
public sealed class SqliteProviderBenchDatabase : IAsyncDisposable
{
    private readonly BenchDataStore _source;
    private readonly ILogger<SqliteProviderBenchDatabase> _logger;
    private readonly SemaphoreSlim _initializeGate = new(1, 1);
    private readonly string _databasePath;
    private readonly string _connectionString;
    private volatile bool _initialized;
    private long _generation;
    private long _requestCount;
    private long _completedCount;
    private long _cancelledCount;
    private long _failedCount;
    private long _countQueryCount;
    private long _rangeQueryCount;
    private long _rowsReturned;
    private long _inFlight;
    private long _maxInFlight;
    private long _totalElapsedTicks;
    private long _lastElapsedTicks;
    private long _lastStartIndex;
    private long _lastRequestedCount;
    private long _lastReturnedCount;
    private long _initializationElapsedTicks;
    private int _rowCount;

    public SqliteProviderBenchDatabase(
        BenchDataStore source,
        ILogger<SqliteProviderBenchDatabase> logger)
    {
        _source = source;
        _logger = logger;

        var directory = Path.Combine(Path.GetTempPath(), "FlexKitTester", "sqlite-provider-bench");
        Directory.CreateDirectory(directory);
        _databasePath = Path.Combine(directory, $"homefront-estimating-{Environment.ProcessId}.db");
        _connectionString = new SqliteConnectionStringBuilder
        {
            DataSource = _databasePath,
            Mode = SqliteOpenMode.ReadWriteCreate,
            Cache = SqliteCacheMode.Shared,
            Pooling = true
        }.ToString();
    }

    public string DatabasePath => _databasePath;

    public async ValueTask<SqliteProviderDatabaseStatus> GetStatusAsync(
        CancellationToken cancellationToken = default)
    {
        await EnsureInitializedAsync(cancellationToken).ConfigureAwait(false);
        return new SqliteProviderDatabaseStatus(
            _databasePath,
            _rowCount,
            Interlocked.Read(ref _generation),
            GetDatabaseBytes(),
            TicksToMilliseconds(Interlocked.Read(ref _initializationElapsedTicks)));
    }

    public SqliteProviderMetricsSnapshot GetMetrics() => new(
        Interlocked.Read(ref _requestCount),
        Interlocked.Read(ref _completedCount),
        Interlocked.Read(ref _cancelledCount),
        Interlocked.Read(ref _failedCount),
        Interlocked.Read(ref _countQueryCount),
        Interlocked.Read(ref _rangeQueryCount),
        Interlocked.Read(ref _rowsReturned),
        Interlocked.Read(ref _inFlight),
        Interlocked.Read(ref _maxInFlight),
        TicksToMilliseconds(Interlocked.Read(ref _lastElapsedTicks)),
        TicksToMilliseconds(Interlocked.Read(ref _totalElapsedTicks)),
        checked((int)Interlocked.Read(ref _lastStartIndex)),
        checked((int)Interlocked.Read(ref _lastRequestedCount)),
        checked((int)Interlocked.Read(ref _lastReturnedCount)));

    public void ResetMetrics()
    {
        Interlocked.Exchange(ref _requestCount, 0);
        Interlocked.Exchange(ref _completedCount, 0);
        Interlocked.Exchange(ref _cancelledCount, 0);
        Interlocked.Exchange(ref _failedCount, 0);
        Interlocked.Exchange(ref _countQueryCount, 0);
        Interlocked.Exchange(ref _rangeQueryCount, 0);
        Interlocked.Exchange(ref _rowsReturned, 0);
        // An operator can reset while an async query is still running. Keep the
        // real in-flight counter so its finally block cannot drive it negative.
        Interlocked.Exchange(ref _maxInFlight, Interlocked.Read(ref _inFlight));
        Interlocked.Exchange(ref _totalElapsedTicks, 0);
        Interlocked.Exchange(ref _lastElapsedTicks, 0);
        Interlocked.Exchange(ref _lastStartIndex, 0);
        Interlocked.Exchange(ref _lastRequestedCount, 0);
        Interlocked.Exchange(ref _lastReturnedCount, 0);
    }

    public async ValueTask<SqliteProviderPage> QueryEstimatingItemsAsync(
        int startIndex,
        int count,
        int datasetRowLimit,
        string? searchText,
        bool includeTotalCount,
        int? knownTotalCount,
        int artificialLatencyMilliseconds,
        CancellationToken cancellationToken = default)
    {
        await EnsureInitializedAsync(cancellationToken).ConfigureAwait(false);

        startIndex = Math.Max(0, startIndex);
        count = Math.Clamp(count, 1, 2_000);
        datasetRowLimit = Math.Clamp(datasetRowLimit, 1, _rowCount);
        artificialLatencyMilliseconds = Math.Clamp(artificialLatencyMilliseconds, 0, 5_000);
        var search = (searchText ?? string.Empty).Trim();
        var stopwatch = Stopwatch.StartNew();

        Interlocked.Increment(ref _requestCount);
        var inFlight = Interlocked.Increment(ref _inFlight);
        UpdateMaximum(ref _maxInFlight, inFlight);

        try
        {
            if (artificialLatencyMilliseconds > 0)
                await Task.Delay(artificialLatencyMilliseconds, cancellationToken).ConfigureAwait(false);

            await using var connection = new SqliteConnection(_connectionString);
            await connection.OpenAsync(cancellationToken).ConfigureAwait(false);

            var totalCount = knownTotalCount;
            if (includeTotalCount || totalCount is null)
            {
                await using var countCommand = connection.CreateCommand();
                countCommand.CommandText = $$"""
                    SELECT COUNT(*)
                    FROM estimating_items
                    WHERE source_ordinal <= $row_limit
                      AND (
                          $search = ''
                          OR phase LIKE $pattern ESCAPE '\'
                          OR item LIKE $pattern ESCAPE '\'
                          OR item_description LIKE $pattern ESCAPE '\'
                          OR phase_description LIKE $pattern ESCAPE '\'
                          OR po_index LIKE $pattern ESCAPE '\'
                          OR cost_code LIKE $pattern ESCAPE '\'
                      );
                    """;
                AddFilterParameters(countCommand, datasetRowLimit, search);
                Interlocked.Increment(ref _countQueryCount);
                totalCount = Convert.ToInt32(
                    await countCommand.ExecuteScalarAsync(cancellationToken).ConfigureAwait(false));
            }

            await using var rangeCommand = connection.CreateCommand();
            rangeCommand.CommandText = $$"""
                SELECT row_id, division_id, phase, cost_type, item,
                       phase_description, item_description, notes,
                       po_index, po_index_description, cost_code,
                       cost_code_description, tax_group, waste_percent,
                       round_direction, round_to, price, is_quote,
                       category, category_description, order_uom,
                       conversion_factor, part_number, formula
                FROM estimating_items
                WHERE source_ordinal <= $row_limit
                  AND (
                      $search = ''
                      OR phase LIKE $pattern ESCAPE '\'
                      OR item LIKE $pattern ESCAPE '\'
                      OR item_description LIKE $pattern ESCAPE '\'
                      OR phase_description LIKE $pattern ESCAPE '\'
                      OR po_index LIKE $pattern ESCAPE '\'
                      OR cost_code LIKE $pattern ESCAPE '\'
                  )
                ORDER BY phase COLLATE NOCASE, item COLLATE NOCASE, row_id
                LIMIT $take OFFSET $skip;
                """;
            AddFilterParameters(rangeCommand, datasetRowLimit, search);
            rangeCommand.Parameters.AddWithValue("$take", count);
            rangeCommand.Parameters.AddWithValue("$skip", startIndex);

            Interlocked.Increment(ref _rangeQueryCount);
            var rows = new List<SqliteEstimatingItemRow>(count);
            await using var reader = await rangeCommand.ExecuteReaderAsync(cancellationToken).ConfigureAwait(false);
            while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false))
                rows.Add(ReadRow(reader));

            Interlocked.Increment(ref _completedCount);
            Interlocked.Add(ref _rowsReturned, rows.Count);
            Interlocked.Exchange(ref _lastStartIndex, startIndex);
            Interlocked.Exchange(ref _lastRequestedCount, count);
            Interlocked.Exchange(ref _lastReturnedCount, rows.Count);

            return new SqliteProviderPage(rows, totalCount.Value);
        }
        catch (OperationCanceledException)
        {
            Interlocked.Increment(ref _cancelledCount);
            throw;
        }
        catch
        {
            Interlocked.Increment(ref _failedCount);
            throw;
        }
        finally
        {
            stopwatch.Stop();
            Interlocked.Decrement(ref _inFlight);
            Interlocked.Exchange(ref _lastElapsedTicks, stopwatch.ElapsedTicks);
            Interlocked.Add(ref _totalElapsedTicks, stopwatch.ElapsedTicks);
        }
    }

    private async Task EnsureInitializedAsync(CancellationToken cancellationToken)
    {
        if (_initialized)
            return;

        await _initializeGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            if (_initialized)
                return;

            var stopwatch = Stopwatch.StartNew();
            _source.Load();
            if (_source.EstimatingError is not null)
                throw new InvalidOperationException(
                    $"HomeFront EstimatingItems export is unavailable: {_source.EstimatingError}");
            if (_source.EstimatingItems.Count == 0)
                throw new InvalidOperationException("HomeFront EstimatingItems export contains no rows.");

            SqliteConnection.ClearAllPools();
            DeleteDatabaseFiles();

            await using var connection = new SqliteConnection(_connectionString);
            await connection.OpenAsync(cancellationToken).ConfigureAwait(false);
            await using (var schemaCommand = connection.CreateCommand())
            {
                schemaCommand.CommandText = """
                    PRAGMA journal_mode = WAL;
                    PRAGMA synchronous = NORMAL;
                    PRAGMA temp_store = MEMORY;

                    CREATE TABLE estimating_items (
                        row_id INTEGER PRIMARY KEY,
                        source_ordinal INTEGER NOT NULL UNIQUE,
                        division_id INTEGER NOT NULL,
                        phase TEXT NOT NULL,
                        cost_type TEXT NOT NULL,
                        item TEXT NOT NULL,
                        phase_description TEXT NOT NULL,
                        item_description TEXT NOT NULL,
                        notes TEXT NOT NULL,
                        po_index TEXT NOT NULL,
                        po_index_description TEXT NOT NULL,
                        cost_code TEXT NOT NULL,
                        cost_code_description TEXT NOT NULL,
                        tax_group TEXT NOT NULL,
                        waste_percent INTEGER NOT NULL,
                        round_direction INTEGER NOT NULL,
                        round_to NUMERIC NOT NULL,
                        price NUMERIC NOT NULL,
                        is_quote INTEGER NOT NULL,
                        category TEXT NOT NULL,
                        category_description TEXT NOT NULL,
                        order_uom TEXT NOT NULL,
                        conversion_factor NUMERIC NOT NULL,
                        part_number TEXT NOT NULL,
                        formula TEXT NOT NULL
                    );

                    CREATE INDEX ix_estimating_items_phase_item
                        ON estimating_items(phase COLLATE NOCASE, item COLLATE NOCASE, row_id);
                    CREATE INDEX ix_estimating_items_division_phase_item
                        ON estimating_items(division_id, phase COLLATE NOCASE, item COLLATE NOCASE);
                    CREATE INDEX ix_estimating_items_po_index
                        ON estimating_items(po_index COLLATE NOCASE, row_id);
                    CREATE INDEX ix_estimating_items_cost_code
                        ON estimating_items(cost_code COLLATE NOCASE, row_id);

                    CREATE TABLE bench_metadata (
                        metadata_key TEXT PRIMARY KEY,
                        metadata_value TEXT NOT NULL
                    );
                    """;
                await schemaCommand.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false);
            }

            await using var transaction = await connection.BeginTransactionAsync(cancellationToken).ConfigureAwait(false);
            await using var insert = CreateInsertCommand(connection, (SqliteTransaction)transaction);
            var ordinal = 0;
            foreach (var sourceRow in _source.EstimatingItems)
            {
                cancellationToken.ThrowIfCancellationRequested();
                ordinal++;
                AssignInsertParameters(insert, ordinal, sourceRow);
                await insert.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false);
            }

            await InsertMetadataAsync(connection, (SqliteTransaction)transaction,
                "schema_version", "1", cancellationToken).ConfigureAwait(false);
            await InsertMetadataAsync(connection, (SqliteTransaction)transaction,
                "source", BenchDataStore.EstimatingFileName, cancellationToken).ConfigureAwait(false);
            await InsertMetadataAsync(connection, (SqliteTransaction)transaction,
                "source_row_count", ordinal.ToString(System.Globalization.CultureInfo.InvariantCulture),
                cancellationToken).ConfigureAwait(false);
            await transaction.CommitAsync(cancellationToken).ConfigureAwait(false);

            _rowCount = ordinal;
            Interlocked.Increment(ref _generation);
            _initialized = true;
            stopwatch.Stop();
            Interlocked.Exchange(ref _initializationElapsedTicks, stopwatch.ElapsedTicks);
            _logger.LogInformation(
                "SQLite provider bench imported {RowCount} HomeFront estimating rows into {DatabasePath} in {ElapsedMs:N1} ms.",
                ordinal, _databasePath, stopwatch.Elapsed.TotalMilliseconds);
        }
        finally
        {
            _initializeGate.Release();
        }
    }

    private static SqliteCommand CreateInsertCommand(SqliteConnection connection, SqliteTransaction transaction)
    {
        var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = """
            INSERT INTO estimating_items (
                row_id, source_ordinal, division_id, phase, cost_type, item,
                phase_description, item_description, notes, po_index,
                po_index_description, cost_code, cost_code_description,
                tax_group, waste_percent, round_direction, round_to, price,
                is_quote, category, category_description, order_uom,
                conversion_factor, part_number, formula)
            VALUES (
                $row_id, $source_ordinal, $division_id, $phase, $cost_type, $item,
                $phase_description, $item_description, $notes, $po_index,
                $po_index_description, $cost_code, $cost_code_description,
                $tax_group, $waste_percent, $round_direction, $round_to, $price,
                $is_quote, $category, $category_description, $order_uom,
                $conversion_factor, $part_number, $formula);
            """;

        foreach (var name in new[]
        {
            "$row_id", "$source_ordinal", "$division_id", "$phase", "$cost_type", "$item",
            "$phase_description", "$item_description", "$notes", "$po_index",
            "$po_index_description", "$cost_code", "$cost_code_description", "$tax_group",
            "$waste_percent", "$round_direction", "$round_to", "$price", "$is_quote",
            "$category", "$category_description", "$order_uom", "$conversion_factor",
            "$part_number", "$formula"
        })
            command.Parameters.Add(new SqliteParameter(name, null));

        command.Prepare();
        return command;
    }

    private static void AssignInsertParameters(
        SqliteCommand command,
        int ordinal,
        EstimatingItemRecord row)
    {
        command.Parameters["$row_id"].Value = ordinal;
        command.Parameters["$source_ordinal"].Value = ordinal;
        command.Parameters["$division_id"].Value = row.DivisionID;
        command.Parameters["$phase"].Value = row.Phase ?? string.Empty;
        command.Parameters["$cost_type"].Value = row.CostCategory ?? string.Empty;
        command.Parameters["$item"].Value = row.Item ?? string.Empty;
        command.Parameters["$phase_description"].Value = row.PhaseDesc ?? string.Empty;
        command.Parameters["$item_description"].Value = row.ItemDesc ?? string.Empty;
        command.Parameters["$notes"].Value = row.Notes ?? string.Empty;
        command.Parameters["$po_index"].Value = row.POIndex ?? string.Empty;
        command.Parameters["$po_index_description"].Value = row.POIndexDescription ?? string.Empty;
        command.Parameters["$cost_code"].Value = row.JCCostCode ?? string.Empty;
        command.Parameters["$cost_code_description"].Value = row.JCCostCodeDesc ?? string.Empty;
        command.Parameters["$tax_group"].Value = row.TaxGroup ?? string.Empty;
        command.Parameters["$waste_percent"].Value = row.WastePercent ?? 0;
        command.Parameters["$round_direction"].Value = row.RoundDir ?? 0;
        command.Parameters["$round_to"].Value = row.RoundTo ?? 0m;
        command.Parameters["$price"].Value = row.Price ?? 0m;
        command.Parameters["$is_quote"].Value = row.IsQuote == true ? 1 : 0;
        command.Parameters["$category"].Value = row.JCCategory ?? string.Empty;
        command.Parameters["$category_description"].Value = row.JCCategoryDesc ?? string.Empty;
        command.Parameters["$order_uom"].Value = row.OrderUOM ?? string.Empty;
        command.Parameters["$conversion_factor"].Value = row.ConversionFactor ?? 0m;
        command.Parameters["$part_number"].Value = row.PartNumber ?? string.Empty;
        command.Parameters["$formula"].Value = row.Formula ?? string.Empty;
    }

    private static async Task InsertMetadataAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        string key,
        string value,
        CancellationToken cancellationToken)
    {
        await using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText =
            "INSERT INTO bench_metadata(metadata_key, metadata_value) VALUES ($key, $value);";
        command.Parameters.AddWithValue("$key", key);
        command.Parameters.AddWithValue("$value", value);
        await command.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false);
    }

    private static void AddFilterParameters(SqliteCommand command, int rowLimit, string search)
    {
        command.Parameters.AddWithValue("$row_limit", rowLimit);
        command.Parameters.AddWithValue("$search", search);
        command.Parameters.AddWithValue("$pattern", $"%{EscapeLikePattern(search)}%");
    }

    private static string EscapeLikePattern(string value) =>
        value.Replace("\\", "\\\\", StringComparison.Ordinal)
            .Replace("%", "\\%", StringComparison.Ordinal)
            .Replace("_", "\\_", StringComparison.Ordinal);

    private static SqliteEstimatingItemRow ReadRow(SqliteDataReader reader) => new()
    {
        RowId = reader.GetInt64(0),
        DivisionId = reader.GetInt32(1),
        Phase = reader.GetString(2),
        CostType = reader.GetString(3),
        Item = reader.GetString(4),
        PhaseDescription = reader.GetString(5),
        Description = reader.GetString(6),
        Notes = reader.GetString(7),
        POIndex = reader.GetString(8),
        POIndexDescription = reader.GetString(9),
        CostCode = reader.GetString(10),
        CostCodeDescription = reader.GetString(11),
        TaxGroup = reader.GetString(12),
        WastePercent = reader.GetInt32(13),
        RoundDirection = reader.GetInt32(14),
        RoundTo = reader.GetDecimal(15),
        Price = reader.GetDecimal(16),
        IsQuote = reader.GetInt32(17) != 0,
        Category = reader.GetString(18),
        CategoryDescription = reader.GetString(19),
        OrderUOM = reader.GetString(20),
        ConversionFactor = reader.GetDecimal(21),
        PartNumber = reader.GetString(22),
        Formula = reader.GetString(23)
    };

    private void DeleteDatabaseFiles()
    {
        foreach (var path in new[] { _databasePath, _databasePath + "-wal", _databasePath + "-shm" })
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    private long GetDatabaseBytes()
    {
        long total = 0;
        foreach (var path in new[] { _databasePath, _databasePath + "-wal", _databasePath + "-shm" })
        {
            if (File.Exists(path))
                total += new FileInfo(path).Length;
        }
        return total;
    }

    private static void UpdateMaximum(ref long target, long candidate)
    {
        while (true)
        {
            var current = Interlocked.Read(ref target);
            if (candidate <= current || Interlocked.CompareExchange(ref target, candidate, current) == current)
                return;
        }
    }

    private static double TicksToMilliseconds(long stopwatchTicks) =>
        stopwatchTicks * 1000d / Stopwatch.Frequency;

    public ValueTask DisposeAsync()
    {
        SqliteConnection.ClearAllPools();
        try
        {
            DeleteDatabaseFiles();
        }
        catch (IOException)
        {
            // The OS will reclaim the process-scoped temp file.
        }
        _initializeGate.Dispose();
        return ValueTask.CompletedTask;
    }
}

public sealed class SqliteEstimatingItemRow
{
    public long RowId { get; init; }
    public int DivisionId { get; init; }
    public string Phase { get; init; } = string.Empty;
    public string CostType { get; init; } = string.Empty;
    public string Item { get; init; } = string.Empty;
    public string PhaseDescription { get; init; } = string.Empty;
    public string Description { get; init; } = string.Empty;
    public string Notes { get; init; } = string.Empty;
    public string POIndex { get; init; } = string.Empty;
    public string POIndexDescription { get; init; } = string.Empty;
    public string CostCode { get; init; } = string.Empty;
    public string CostCodeDescription { get; init; } = string.Empty;
    public string TaxGroup { get; init; } = string.Empty;
    public int WastePercent { get; init; }
    public int RoundDirection { get; init; }
    public decimal RoundTo { get; init; }
    public decimal Price { get; init; }
    public bool IsQuote { get; init; }
    public string Category { get; init; } = string.Empty;
    public string CategoryDescription { get; init; } = string.Empty;
    public string OrderUOM { get; init; } = string.Empty;
    public decimal ConversionFactor { get; init; }
    public string PartNumber { get; init; } = string.Empty;
    public string Formula { get; init; } = string.Empty;
}

public sealed record SqliteProviderPage(
    IReadOnlyList<SqliteEstimatingItemRow> Items,
    int TotalCount);

public sealed record SqliteProviderDatabaseStatus(
    string DatabasePath,
    int RowCount,
    long Generation,
    long DatabaseBytes,
    double InitializationMilliseconds);

public sealed record SqliteProviderMetricsSnapshot(
    long Requests,
    long Completed,
    long Cancelled,
    long Failed,
    long CountQueries,
    long RangeQueries,
    long RowsReturned,
    long InFlight,
    long MaxInFlight,
    double LastElapsedMilliseconds,
    double TotalElapsedMilliseconds,
    int LastStartIndex,
    int LastRequestedCount,
    int LastReturnedCount)
{
    public double AverageElapsedMilliseconds => Completed + Cancelled + Failed == 0
        ? 0d
        : TotalElapsedMilliseconds / (Completed + Cancelled + Failed);
}
