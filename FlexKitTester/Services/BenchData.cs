using System.Diagnostics;
using System.Text.Json;

namespace FlexKitTester.Services;

/// <summary>
/// One row of the EstimatingItems export. FOR JSON PATH OMITS null properties,
/// so every member must be nullable and no member may be required.
/// </summary>
public sealed class EstimatingItemRecord
{
    public int DivisionID { get; set; }
    public string? Phase { get; set; }
    public string? Item { get; set; }
    public string? ItemDesc { get; set; }
    public string? PhaseDesc { get; set; }
    public string? CostCategory { get; set; }
    public string? Notes { get; set; }
    public string? POIndex { get; set; }
    public string? POIndexDescription { get; set; }
    public string? JCCostCode { get; set; }
    public string? JCCostCodeDesc { get; set; }
    public string? JCCategory { get; set; }
    public string? JCCategoryDesc { get; set; }
    public string? TaxGroup { get; set; }
    public int? WastePercent { get; set; }
    public int? RoundDir { get; set; }
    public decimal? RoundTo { get; set; }
    public decimal? Price { get; set; }
    public bool? IsQuote { get; set; }
    public string? OrderUOM { get; set; }
    public decimal? ConversionFactor { get; set; }
    public string? PartNumber { get; set; }
    public string? Formula { get; set; }
}

/// <summary>
/// One row of the tblDBAssemblyDetails export. Same null-omission rule.
/// </summary>
public sealed class AssemblyDetailRecord
{
    public int DivisionID { get; set; }
    public string? Community { get; set; }
    public string? Assembly { get; set; }
    public string? Model { get; set; }
    public string? OptionID { get; set; }
    public string? Phase { get; set; }
    public string? Item { get; set; }
    public string? ItemChart { get; set; }
    public int Sequence { get; set; }
    public decimal? TakeoffQty { get; set; }
    public decimal? OrderQty { get; set; }
}

/// <summary>
/// One row of the tblVendors export (vendors-200.json). Same null-omission
/// rule; export keys not modeled here are ignored by System.Text.Json.
/// </summary>
public sealed class VendorRecord
{
    public int DivisionID { get; set; }
    public string? Vendor_ID { get; set; }
    public string? Vendor_Name { get; set; }
    public string? Addr1 { get; set; }
    public string? Addr2 { get; set; }
    public string? City { get; set; }
    public string? State { get; set; }
    public string? Zip { get; set; }
    public string? Phone { get; set; }
    public string? Fax { get; set; }
    public string? WebUID { get; set; }
    public string? Password { get; set; }
    // Absent from this export — always null today.
    public string? WalletPayeeID { get; set; }
    public string? TaxID { get; set; }
    public string? VendorGroupID { get; set; }
    public string? TradeType { get; set; }
    public string? POFormat { get; set; }
    // json key "isTBD" (case-insensitive match).
    public bool? IsTBD { get; set; }
    public bool? BuildProEnabled { get; set; }
    public bool? InActive { get; set; }
    public double? HoldBackPercentage { get; set; }
    public bool? GLInsRequired { get; set; }
    public bool? WCInsRequired { get; set; }
    public bool? UmbInsRequired { get; set; }
    public bool? AutoInsRequired { get; set; }
    // *ExpDate keys are absent from this export — always null; declared so the
    // insurance projection compiles and lights up if ever re-exported.
    public DateTime? GLInsExpDate { get; set; }
    public DateTime? WCInsExpDate { get; set; }
    public DateTime? UmbInsExpDate { get; set; }
    public DateTime? AutoInsExpDate { get; set; }
    public string? GLInsCompany { get; set; }
    public string? GLInsPolicyNumber { get; set; }
    public string? WCInsCompany { get; set; }
    public string? WCInsPolicyNumber { get; set; }
    public string? UmbInsCompany { get; set; }
    public string? UmbInsPolicyNumber { get; set; }
    public string? AutoInsCompany { get; set; }
    public string? AutoInsPolicyNumber { get; set; }
    public string? LabourTaxGroup { get; set; }
    public string? MaterialTaxGroup { get; set; }
    public string? SubContractTaxGroup { get; set; }
    public string? EquipmentTaxGroup { get; set; }
    public string? OverheadTaxGroup { get; set; }
    public string? OtherTaxGroup { get; set; }
}

/// <summary>
/// One row of the contacts export (vendor-contacts.json, contacttypeid=99).
/// Keyed DivisionID + VendorCode. ContactID is INT in the export.
/// </summary>
public sealed class VendorContactRecord
{
    public int DivisionID { get; set; }
    public int? ContactID { get; set; }
    public int? CommModeID { get; set; }
    public int? ContactTypeID { get; set; }
    public int? VendorID { get; set; }
    public string? Role { get; set; }
    public string? FirstName { get; set; }
    public string? WorkPhone { get; set; }
    public string? CellPhone { get; set; }
    public string? Fax { get; set; }
    public string? Email { get; set; }
    public string? SMSAddress { get; set; }
    public string? VendorCode { get; set; }
    public string? DisplayName { get; set; }
    public bool? PONotice { get; set; }
    public bool? FPONotice { get; set; }
    public bool? SchedNotice { get; set; }
    public bool? ServiceNotice { get; set; }
    public bool? Active { get; set; }
}

/// <summary>
/// One row of the VendorPaymentPoints export (vendor-paypoints.json).
/// Keyed DivisionID + Vendor.
/// </summary>
public sealed class VendorPayPointRecord
{
    public int DivisionID { get; set; }
    public string? Vendor { get; set; }
    public string? POIndex { get; set; }
    public double? PayPoint1Percent { get; set; }
    public double? PayPoint2Percent { get; set; }
    public double? PayPoint3Percent { get; set; }
    public double? PayPoint4Percent { get; set; }
    public double? PayPoint5Percent { get; set; }
}

/// <summary>
/// One row of the vendortaxgroups export (vendor-taxgroups.json — empty by
/// honest export today; shape mirrors the columns FVendor selects).
/// </summary>
public sealed class VendorTaxGroupRecord
{
    public int DivisionID { get; set; }
    public string? Vendor { get; set; }
    public string? Community { get; set; }
    public string? LabourTaxGroup { get; set; }
    public string? MaterialTaxGroup { get; set; }
    public string? SubContractTaxGroup { get; set; }
    public string? EquipmentTaxGroup { get; set; }
    public string? OverheadTaxGroup { get; set; }
    public string? OtherTaxGroup { get; set; }
}

/// <summary>
/// Parses the two exports ONCE per process and hands out the parsed records.
/// Registered as a singleton and preloaded from Program.cs; the bench pages
/// only project the records into their row models.
/// </summary>
public sealed class BenchDataStore
{
    public const string EstimatingFileName = "estimating-items-full.json";
    public const string AssemblyFileName = "assembly-details-full.json";
    public const string VendorsFileName = "vendors-200.json";
    public const string VendorContactsFileName = "vendor-contacts.json";
    public const string VendorPayPointsFileName = "vendor-paypoints.json";
    public const string VendorTaxGroupsFileName = "vendor-taxgroups.json";

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    private readonly string[] _probeRoots;
    private readonly object _gate = new();
    private bool _loaded;

    public BenchDataStore(IWebHostEnvironment env)
    {
        // A published app takes ContentRootPath from the process working
        // directory, so the exe folder must be probed FIRST or a launch from a
        // foreign directory loses the exports.
        _probeRoots =
        [
            Path.Combine(AppContext.BaseDirectory, "Data"),
            Path.Combine(env.ContentRootPath, "Data")
        ];
    }

    public IReadOnlyList<EstimatingItemRecord> EstimatingItems { get; private set; } = [];

    public IReadOnlyList<AssemblyDetailRecord> AssemblyDetails { get; private set; } = [];

    /// <summary>
    /// (DivisionID, Phase, Item) is unique in the export and is the exact key
    /// VB6 FAssembly joins tblDBAssemblyDetails to tblPhaseItem on.
    /// </summary>
    public IReadOnlyDictionary<(int Division, string Phase, string Item), EstimatingItemRecord> ItemLookup
    { get; private set; } = new Dictionary<(int, string, string), EstimatingItemRecord>();

    /// <summary>
    /// Sorted (DivisionID, Vendor_ID trimmed, case-insensitive): the export
    /// spans divisions and Vendor_ID is unique only per division, so paging is
    /// by index over this composite order, never by Vendor_ID lookup.
    /// </summary>
    public IReadOnlyList<VendorRecord> Vendors { get; private set; } = [];

    public IReadOnlyList<VendorContactRecord> VendorContacts { get; private set; } = [];

    public IReadOnlyList<VendorPayPointRecord> VendorPayPoints { get; private set; } = [];

    public IReadOnlyList<VendorTaxGroupRecord> VendorTaxGroups { get; private set; } = [];

    /// <summary>Non-null when an export is missing or unreadable.</summary>
    public string? EstimatingError { get; private set; }

    public string? AssemblyError { get; private set; }

    public string? VendorsError { get; private set; }

    public string? VendorContactsError { get; private set; }

    public string? VendorPayPointsError { get; private set; }

    public string? VendorTaxGroupsError { get; private set; }

    public string? DataRoot { get; private set; }

    public string LoadReport { get; private set; } = "not loaded";

    public void Load()
    {
        if (_loaded) return;
        lock (_gate)
        {
            if (_loaded) return;
            var sw = Stopwatch.StartNew();

            DataRoot = _probeRoots.FirstOrDefault(Directory.Exists);

            EstimatingItems = ReadArray<EstimatingItemRecord>(EstimatingFileName, out var estError);
            EstimatingError = estError;

            var details = ReadArray<AssemblyDetailRecord>(AssemblyFileName, out var asmError);
            AssemblyError = asmError;

            var lookup = new Dictionary<(int, string, string), EstimatingItemRecord>(EstimatingItems.Count);
            foreach (var r in EstimatingItems)
            {
                if (string.IsNullOrEmpty(r.Item)) continue;
                lookup.TryAdd((r.DivisionID, r.Phase ?? "", r.Item), r);
            }
            ItemLookup = lookup;

            // VB6 FAssembly orders gItems by Phase,Item,ItemChart,Sequence
            // within one assembly; the assembly key leads because this export
            // is every assembly at once.
            AssemblyDetails = details
                .OrderBy(d => d.DivisionID)
                .ThenBy(d => d.Community ?? "", StringComparer.Ordinal)
                .ThenBy(d => d.Assembly ?? "", StringComparer.Ordinal)
                .ThenBy(d => d.Model ?? "", StringComparer.Ordinal)
                .ThenBy(d => d.OptionID ?? "", StringComparer.Ordinal)
                .ThenBy(d => d.Phase ?? "", StringComparer.Ordinal)
                .ThenBy(d => d.Item ?? "", StringComparer.Ordinal)
                .ThenBy(d => d.ItemChart ?? "", StringComparer.Ordinal)
                .ThenBy(d => d.Sequence)
                .ToList();

            Vendors = ReadArray<VendorRecord>(VendorsFileName, out var vendorsError)
                .OrderBy(v => v.DivisionID)
                .ThenBy(v => (v.Vendor_ID ?? "").Trim(), StringComparer.OrdinalIgnoreCase)
                .ToList();
            VendorsError = vendorsError;

            VendorContacts = ReadArray<VendorContactRecord>(VendorContactsFileName, out var vendorContactsError);
            VendorContactsError = vendorContactsError;

            VendorPayPoints = ReadArray<VendorPayPointRecord>(VendorPayPointsFileName, out var vendorPayPointsError);
            VendorPayPointsError = vendorPayPointsError;

            VendorTaxGroups = ReadArray<VendorTaxGroupRecord>(VendorTaxGroupsFileName, out var vendorTaxGroupsError);
            VendorTaxGroupsError = vendorTaxGroupsError;

            sw.Stop();
            LoadReport =
                $"data root: {DataRoot ?? "NOT FOUND"} | estimating: {EstimatingItems.Count} rows"
                + $"{(EstimatingError is null ? "" : " [" + EstimatingError + "]")}"
                + $" | assembly: {AssemblyDetails.Count} rows"
                + $"{(AssemblyError is null ? "" : " [" + AssemblyError + "]")}"
                + $" | {sw.ElapsedMilliseconds} ms"
                + $" | vendors: {Vendors.Count} rows"
                + $"{(VendorsError is null ? "" : " [" + VendorsError + "]")}"
                + $" | contacts: {VendorContacts.Count}"
                + $"{(VendorContactsError is null ? "" : " [" + VendorContactsError + "]")}"
                + $" | paypoints: {VendorPayPoints.Count}"
                + $"{(VendorPayPointsError is null ? "" : " [" + VendorPayPointsError + "]")}"
                + $" | taxgroups: {VendorTaxGroups.Count}"
                + $"{(VendorTaxGroupsError is null ? "" : " [" + VendorTaxGroupsError + "]")}";
            _loaded = true;
        }
    }

    private List<T> ReadArray<T>(string fileName, out string? error)
    {
        error = null;
        var path = _probeRoots.Select(root => Path.Combine(root, fileName)).FirstOrDefault(File.Exists);
        if (path is null)
        {
            error = $"{fileName} not found under {string.Join(" or ", _probeRoots)}";
            return [];
        }
        try
        {
            var bytes = File.ReadAllBytes(path);
            var length = StripChunkBreaks(bytes);
            return JsonSerializer.Deserialize<List<T>>(new ReadOnlySpan<byte>(bytes, 0, length), JsonOptions) ?? [];
        }
        catch (Exception ex)
        {
            error = $"{fileName} could not be parsed: {ex.Message}";
            return [];
        }
    }

    /// <summary>
    /// sqlcmd chunks FOR JSON output with a hard CR/LF every 2033 characters,
    /// which is not valid JSON. FOR JSON escapes real newlines as \n, so every
    /// raw CR/LF byte is a chunk break and dropping them is lossless.
    /// Compacts in place and returns the surviving length.
    /// </summary>
    private static int StripChunkBreaks(byte[] buffer)
    {
        var read = buffer.Length >= 3 && buffer[0] == 0xEF && buffer[1] == 0xBB && buffer[2] == 0xBF ? 3 : 0;
        var write = 0;
        for (; read < buffer.Length; read++)
        {
            var b = buffer[read];
            if (b is (byte)'\n' or (byte)'\r') continue;
            buffer[write++] = b;
        }
        return write;
    }
}
