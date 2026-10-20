using System.Security.Cryptography;
using System.Text.Json;

namespace FlexKitTester.Services;

public sealed record CrystalMenuReport(int Number, string CorpusId, string Name, string Sha256, string RelativePath);

/// <summary>Exact binary selection for HomeFront phase 1; never reads the synthetic catalog.</summary>
public sealed class CrystalBenchMenuScope
{
    public const string ConfigurationKey = "CrystalReportsBench:MenuScope";
    public string Root { get; }
    public IReadOnlyList<CrystalMenuReport> Reports { get; }
    private readonly string _manifestPath;
    private readonly byte[] _manifestHash;

    private CrystalBenchMenuScope(string root, string manifestPath, byte[] manifestHash, List<CrystalMenuReport> reports)
    {
        Root = root;
        _manifestPath = manifestPath;
        _manifestHash = manifestHash;
        Reports = reports.AsReadOnly();
    }

    public static CrystalBenchMenuScope Load(string manifestPath)
    {
        var full = Path.GetFullPath(manifestPath);
        var corpus = Path.GetDirectoryName(full)!;
        var root = Directory.GetParent(corpus)?.FullName ?? throw new InvalidDataException("Menu scope requires a corpus directory.");
        var bytes = File.ReadAllBytes(full);
        var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
        var scope = JsonSerializer.Deserialize<ScopeDocument>(bytes, options);
        var catalog = JsonSerializer.Deserialize<CatalogReport[]>(File.ReadAllBytes(Path.Combine(corpus, "reports.json")), options);
        if (scope?.Scope != "homefront-reports-menu-phase-1" || scope.Reports?.Length != 19
            || scope.SourceDirectories?.Length != 2 || scope.SourceDirectories.Distinct(StringComparer.Ordinal).Count() != 2
            || scope.SourceDirectories.Any(string.IsNullOrWhiteSpace) || catalog is null)
            throw new InvalidDataException("Phase 1 requires exactly 19 approved HomeFront menu reports and both source directories.");

        var reports = new List<CrystalMenuReport>();
        var ids = new HashSet<string>(StringComparer.Ordinal);
        var hashes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in scope.Reports)
        {
            if (entry is null || entry.Number != reports.Count + 1 || string.IsNullOrWhiteSpace(entry.CorpusId)
                || !ids.Add(entry.CorpusId) || string.IsNullOrWhiteSpace(entry.Name)
                || Path.GetFileName(entry.Name) != entry.Name || entry.Sha256?.Length != 64
                || entry.Sha256.Any(c => !char.IsAsciiHexDigit(c)) || !hashes.Add(entry.Sha256))
                throw new InvalidDataException("Menu numbers, corpus IDs, names and hashes must be unique and ordered.");
            var matches = catalog.Where(r => r?.Id == entry.CorpusId).ToArray();
            if (matches.Length != 1) throw new InvalidDataException($"Missing or ambiguous menu corpus ID: {entry.CorpusId}.");
            var match = matches[0];
            var name = entry.Name + ".rpt";
            if (match.Sha256 != entry.Sha256 || match.Names is null || !match.Names.Contains(name, StringComparer.Ordinal)
                || match.Sources is null || !scope.SourceDirectories.All(dir => match.Sources.Contains(dir + "/" + name, StringComparer.Ordinal))
                || string.IsNullOrWhiteSpace(match.Path))
                throw new InvalidDataException($"Menu source, name or hash mismatch: {entry.CorpusId}.");
            var path = Path.GetFullPath(match.Path, root);
            EnsureContainedFile(Path.Combine(corpus, "rpt"), path);
            if (Path.GetFileName(path) != name) throw new InvalidDataException($"Menu file name mismatch: {entry.CorpusId}.");
            ValidateHash(path, entry.Sha256);
            reports.Add(new(entry.Number, entry.CorpusId, entry.Name, entry.Sha256, match.Path));
        }
        return new(root, full, SHA256.HashData(bytes), reports);
    }

    public CrystalMenuReport RequireSource(string path)
    {
        ValidateManifest();
        var full = Path.GetFullPath(path);
        var report = Reports.SingleOrDefault(r => Path.GetFullPath(r.RelativePath, Root) == full)
            ?? throw new InvalidDataException("Select one of the 19 approved HomeFront menu reports.");
        EnsureContainedFile(Path.Combine(Root, "corpus", "rpt"), full);
        ValidateHash(full, report.Sha256);
        return report;
    }

    /// <summary>Recheck the isolated copy immediately before native conversion, including stream inputs.</summary>
    public CrystalMenuReport RequireAllowedFile(string path)
    {
        ValidateManifest();
        using var stream = File.OpenRead(path);
        var hash = Convert.ToHexString(SHA256.HashData(stream));
        return Reports.SingleOrDefault(r => r.Sha256.Equals(hash, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidDataException("This RPT binary is outside the 19-report HomeFront menu scope.");
    }

    private void ValidateManifest()
    {
        if (!File.Exists(_manifestPath) || !SHA256.HashData(File.ReadAllBytes(_manifestPath)).SequenceEqual(_manifestHash))
            throw new InvalidDataException("The menu scope changed or is unavailable. Restart the bench after validating its configuration.");
    }

    private static void ValidateHash(string path, string expected)
    {
        if (new FileInfo(path).Length > CrystalBenchWorkspace.MaxReportBytes)
            throw new InvalidDataException("RPT files are limited to 64 MB.");
        using var stream = File.OpenRead(path);
        if (!Convert.ToHexString(SHA256.HashData(stream)).Equals(expected, StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("The report bytes do not match the approved HomeFront menu hash.");
    }

    private static void EnsureContainedFile(string root, string path)
    {
        var relative = Path.GetRelativePath(root, path);
        if (Path.IsPathRooted(relative) || relative == ".." || relative.StartsWith(".." + Path.DirectorySeparatorChar)
            || !string.Equals(Path.GetExtension(path), ".rpt", StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("Menu reports must be original RPT files within corpus/rpt.");
        var current = root;
        if ((File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0)
            throw new InvalidDataException("Symbolic links are not allowed in menu report paths.");
        foreach (var part in relative.Split(Path.DirectorySeparatorChar))
        {
            current = Path.Combine(current, part);
            if ((File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0)
                throw new InvalidDataException("Symbolic links are not allowed in menu report paths.");
        }
    }

    private sealed record ScopeDocument(string Scope, string[] SourceDirectories, ScopeReport[] Reports);
    private sealed record ScopeReport(int Number, string CorpusId, string Name, string Sha256);
    private sealed record CatalogReport(string Id, string Sha256, string Path, string[] Names, string[] Sources);
}
