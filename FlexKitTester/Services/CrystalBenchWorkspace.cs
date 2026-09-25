using Fx.ControlKit.Reports;

namespace FlexKitTester.Services;

/// <summary>Owns only temporary copies. Source reports and reference XML are never modified or consulted.</summary>
public sealed class CrystalBenchWorkspace : IAsyncDisposable
{
    public const long MaxReportBytes = 64 * 1024 * 1024;
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "FlexKitTester", "reports", Guid.NewGuid().ToString("N"));
    private readonly SemaphoreSlim _gate = new(1);
    private bool _disposed;
    private readonly Queue<string> _runs = new();
    public string? ReportRoot { get; }

    public CrystalBenchWorkspace(string? reportRoot)
    {
        ReportRoot = string.IsNullOrWhiteSpace(reportRoot) ? null : Path.GetFullPath(reportRoot);
    }

    public static string? ResolveRoot(IConfiguration configuration, string contentRoot)
    {
        var configured = configuration["CrystalReportsBench:ReportRoot"];
        if (!string.IsNullOrWhiteSpace(configured)) return Path.GetFullPath(configured, contentRoot);
        var sibling = Path.GetFullPath("../../../JavaToCSharp/Reports", contentRoot);
        return Directory.Exists(sibling) ? sibling : null;
    }

    public string ValidateSource(string path)
    {
        if (ReportRoot is null) throw new InvalidOperationException("No server report folder is configured.");
        var full = Path.GetFullPath(path);
        var relative = Path.GetRelativePath(ReportRoot, full);
        if (Path.IsPathRooted(relative) || relative == ".." || relative.StartsWith(".." + Path.DirectorySeparatorChar))
            throw new InvalidDataException("The report must be inside the configured report folder.");
        var current = ReportRoot;
        foreach (var part in relative.Split(Path.DirectorySeparatorChar))
        {
            current = Path.Combine(current, part);
            if ((File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0)
                throw new InvalidDataException("Symbolic links are not allowed in the report browser.");
        }
        ValidateName(full);
        if (new FileInfo(full).Length > MaxReportBytes) throw new InvalidDataException("RPT files are limited to 64 MB.");
        return full;
    }

    public async Task<CrystalRptConversionResult> ConvertFileAsync(string path, Action<string>? progress, CancellationToken cancellationToken)
    {
        var full = ValidateSource(path);
        await using var stream = new FileStream(full, FileMode.Open, FileAccess.Read, FileShare.Read);
        return await ConvertStreamAsync(Path.GetFileName(full), stream, progress, cancellationToken);
    }

    public async Task<CrystalRptConversionResult> ConvertStreamAsync(string name, Stream stream, Action<string>? progress, CancellationToken cancellationToken)
    {
        ValidateName(name);
        await _gate.WaitAsync(cancellationToken);
        string? run = null;
        try
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            run = Path.Combine(_directory, Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(run);
            var source = Path.Combine(run, Path.GetFileName(name));
            await using (var output = new FileStream(source, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                var buffer = new byte[81920];
                long total = 0;
                int count;
                while ((count = await stream.ReadAsync(buffer, cancellationToken)) != 0)
                {
                    total += count;
                    if (total > MaxReportBytes) throw new InvalidDataException("RPT files are limited to 64 MB.");
                    await output.WriteAsync(buffer.AsMemory(0, count), cancellationToken);
                }
            }
            var result = await CrystalRptToXml.ConvertAsync(source, Path.ChangeExtension(source, ".xml"),
                new(ExtractSubreports: true, Progress: progress), cancellationToken);
            // The synchronous binary decoder cannot be interrupted mid-record. Do not publish a cancelled result.
            cancellationToken.ThrowIfCancellationRequested();
            _runs.Enqueue(run);
            while (_runs.Count > 4)
            {
                var expired = _runs.Dequeue();
                if (Directory.Exists(expired)) Directory.Delete(expired, true);
            }
            return result;
        }
        catch
        {
            if (run is not null && Directory.Exists(run)) Directory.Delete(run, true);
            throw;
        }
        finally { _gate.Release(); }
    }

    private static void ValidateName(string name)
    {
        if (!string.Equals(Path.GetExtension(name), ".rpt", StringComparison.OrdinalIgnoreCase)
            || string.IsNullOrWhiteSpace(Path.GetFileNameWithoutExtension(name)))
            throw new InvalidDataException("Select a .rpt file.");
    }

    public async ValueTask DisposeAsync()
    {
        await _gate.WaitAsync();
        try
        {
            if (_disposed) return;
            _disposed = true;
            if (Directory.Exists(_directory)) Directory.Delete(_directory, true);
        }
        finally { _gate.Release(); }
    }
}
