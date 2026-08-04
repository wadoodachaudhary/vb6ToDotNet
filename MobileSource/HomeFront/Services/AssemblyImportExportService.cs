using System.Data;
using System.Text;
using ClosedXML.Excel;
using HomeFront.Data;

namespace HomeFront.Services;

/// <summary>
/// Non-UI engine behind FImportExportAssemblies (the unified Data Import/Export
/// Wizard — VB6 HFEst FAssemblyImport.frm / MAssemblyImport.bas / "MAssemblyExport",
/// plus the HHM-440 vendor modes): the workbook/delimited-text readers, the
/// HHM-616 RFC-4180 CSV writer, the export table/workbook builders, and the
/// ImportedAssemblies / ImportedAssemblyTakeoffs staging + Purch_* validate/commit
/// stored-proc wrappers. Extracted verbatim from the wizard page — every SQL
/// string, parse rule and message is unchanged. Stateless per call: the page
/// passes divisionId / session / bytes / ids / tables and maps the returned
/// rows / bytes / counts into its own wizard state (Stage, ErrorRows, Messages,
/// progress). Registered scoped in Program.cs of BOTH apps; this file is kept
/// byte-identical between HomeFrontPB and HomeFront (sync pair).
/// </summary>
public sealed class AssemblyImportExportService
{
    private const string SRCFILE = "AssemblyImportExportService::";

    private readonly DbWrapperSqlServer _db;
    private readonly ILogger<AssemblyImportExportService> _logger;
    private const int MaxSqlParams = 2000; // stay under SQL Server's 2100-parameter cap, with margin
    public AssemblyImportExportService(DbWrapperSqlServer db, ILogger<AssemblyImportExportService> logger)
    {
        _db = db;
        _logger = logger;
    }

    // ──────────────────────────────────────────────────────────────────────
    // Workbook / delimited-text readers
    // ──────────────────────────────────────────────────────────────────────

    // Convert a ReadExcel DataTable into a sheet where row 0 is the column headers
    // (column names) followed by the data rows.
    public List<List<string>> DataTableToSheet(DataTable dt)
    {
        var sheet = new List<List<string>>();
        var header = dt.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToList();
        sheet.Add(header);
        foreach (DataRow r in dt.Rows)
        {
            var row = new List<string>(header.Count);
            for (int c = 0; c < header.Count; c++)
                row.Add(r[c]?.ToString()?.Trim() ?? "");
            sheet.Add(row);
        }
        return sheet;
    }

    // ── HHM-557: legacy BIFF .xls support ────────────────────────────────────
    // VB6 "Export Assemblies" (MAssemblyImport.bas:161 -> FDataExport.frm:164-168)
    // writes a binary Excel 97-2003 (BIFF) workbook inside an OLE compound file;
    // ClosedXML reads only OPC (PK-zip) .xlsx. These helpers sniff the container
    // signature and read BIFF workbooks host-side via ExcelDataReader (FlexKit
    // stays .xlsx-only as documented in FilerControl.cs).
    public bool IsOleCompoundFile(byte[] bytes) =>
        bytes.Length >= 8 &&
        bytes[0] == 0xD0 && bytes[1] == 0xCF && bytes[2] == 0x11 && bytes[3] == 0xE0 &&
        bytes[4] == 0xA1 && bytes[5] == 0xB1 && bytes[6] == 0x1A && bytes[7] == 0xE1;

    private static bool IsZipPackage(byte[] bytes) =>
        bytes.Length >= 4 &&
        bytes[0] == 0x50 && bytes[1] == 0x4B && bytes[2] == 0x03 && bytes[3] == 0x04;

    // HHM-557 (3): build a fatal-error message that names the detected workbook
    // format instead of surfacing the raw parser message on its own.
    public string DescribeWorkbookReadFailure(byte[] bytes, Exception ex)
    {
        var fmt = IsOleCompoundFile(bytes)
            ? "Excel 97-2003 binary workbook (.xls, BIFF/OLE compound)"
            : IsZipPackage(bytes)
                ? "Excel Workbook (.xlsx, OPC zip)"
                : "unrecognized (neither a BIFF .xls nor an OPC .xlsx workbook)";
        return $"Could not read the Excel workbook — detected format: {fmt}. {ex.Message}";
    }

    // HHM-557: read sheet 1 of a legacy BIFF .xls into the SAME header-row +
    // data-rows string grid that FilerControl.ReadExcel + DataTableToSheet
    // produce for .xlsx (first used row = column headers, blank-header columns
    // dropped, duplicate headers suffixed _1/_2..., cells trimmed, fully-empty
    // data rows skipped). Reader API only — no ExcelDataReader.DataSet package.
    public List<List<string>> LegacyXlsToSheet(byte[] bytes)
    {
        // BIFF text records are legacy-codepage encoded; ExcelDataReader needs
        // the codepage table registered (no-op when already registered).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        var sheet = new List<List<string>>();
        using var stream = new MemoryStream(bytes);
        using var reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);

        List<int>? headerCols = null;   // source column indexes kept from the header row
        while (reader.Read())           // sheet 1 only — never NextResult()
        {
            var raw = new List<string>(reader.FieldCount);
            for (int c = 0; c < reader.FieldCount; c++)
                raw.Add(reader.GetValue(c)?.ToString()?.Trim() ?? "");

            if (headerCols == null)
            {
                if (raw.All(string.IsNullOrEmpty))
                    continue;           // leading blank rows (ClosedXML FirstRowUsed parity)
                headerCols = new List<int>();
                var header = new List<string>();
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < raw.Count; i++)
                {
                    if (string.IsNullOrEmpty(raw[i]))
                        continue;       // ReadExcel drops columns with blank headers
                    var unique = raw[i];
                    int n = 1;
                    while (!seen.Add(unique))
                        unique = $"{raw[i]}_{n++}";
                    headerCols.Add(i);
                    header.Add(unique);
                }
                sheet.Add(header);
                continue;
            }

            var row = headerCols.Select(i => i < raw.Count ? raw[i] : "").ToList();
            if (row.All(string.IsNullOrEmpty))
                continue;               // ReadExcel skips rows with no data
            sheet.Add(row);
        }
        return sheet;
    }

    // ── HHM-595/597/599/603: positional workbook readers for Import Assembly
    // Takeoffs. The takeoff file is a pivoted TEMPLATE, not a header-row table —
    // VB6 loads it with HasColHeadings=False so every row/column keeps its
    // spreadsheet position: sheet[r][c] == worksheet cell (r+1, c+1). The scan is
    // bounded by the last cell with CONTENT (the template formats rows out to
    // 55555 — style-only cells are ignored) and trailing all-empty rows are
    // trimmed. ──
    public List<List<string>> XlsxToPositionalSheet(byte[] bytes)
    {
        var sheet = new List<List<string>>();
        using var stream = new MemoryStream(bytes);
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheets.FirstOrDefault();   // sheet 1 only ("Assemblies")
        if (ws == null) return sheet;
        var last = ws.LastCellUsed(XLCellsUsedOptions.Contents);
        if (last == null) return sheet;

        int maxRow = last.Address.RowNumber;
        int maxCol = last.Address.ColumnNumber;
        for (int r = 1; r <= maxRow; r++)
        {
            var row = new List<string>(maxCol);
            for (int c = 1; c <= maxCol; c++)
                row.Add(ws.Cell(r, c).Value.ToString().Trim());
            sheet.Add(row);
        }
        TrimTrailingEmptyRows(sheet);
        return sheet;
    }

    // BIFF .xls variant (same OLE-compound sniff as LegacyXlsToSheet) — raw rows
    // in file order, no header semantics, cells padded to FieldCount.
    public List<List<string>> LegacyXlsToPositionalSheet(byte[] bytes)
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        var sheet = new List<List<string>>();
        using var stream = new MemoryStream(bytes);
        using var reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);
        while (reader.Read())       // sheet 1 only — never NextResult()
        {
            var row = new List<string>(reader.FieldCount);
            for (int c = 0; c < reader.FieldCount; c++)
                row.Add(reader.GetValue(c)?.ToString()?.Trim() ?? "");
            sheet.Add(row);
        }
        TrimTrailingEmptyRows(sheet);
        return sheet;
    }

    private static void TrimTrailingEmptyRows(List<List<string>> sheet)
    {
        while (sheet.Count > 0 && sheet[^1].All(string.IsNullOrEmpty))
            sheet.RemoveAt(sheet.Count - 1);
    }

    public List<List<string>> DelimitedToSheet(string content, char delimiter)
    {
        var sheet = new List<List<string>>();
        var lines = content.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
        foreach (var line in lines)
        {
            // keep blank lines as empty rows so pivot row indices stay aligned
            sheet.Add(ParseLine(line, delimiter));
        }
        // trim a single trailing empty line that Split commonly produces
        while (sheet.Count > 0 && sheet[^1].All(string.IsNullOrEmpty))
            sheet.RemoveAt(sheet.Count - 1);
        return sheet;
    }

    // Parse a single delimited line, honoring quoted fields ("" -> ")
    private static List<string> ParseLine(string line, char delimiter)
    {
        var result = new List<string>();
        var inQuotes = false;
        var field = new StringBuilder();
        for (int i = 0; i < line.Length; i++)
        {
            var c = line[i];
            if (c == '"')
            {
                if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                {
                    field.Append('"');
                    i++;
                }
                else inQuotes = !inQuotes;
            }
            else if (c == delimiter && !inQuotes)
            {
                result.Add(field.ToString());
                field.Clear();
            }
            else field.Append(c);
        }
        result.Add(field.ToString());
        return result;
    }

    // VB6 FImportVendors.frm ReadText 364-374 — IIF: drop leading rows until the
    // "!VEND" header and trailing rows after the "VEND" data block, then drop the
    // record-type tag column (first field: "!VEND" on the header, "VEND" on data).
    public void TrimIifVendorSheet(List<List<string>> sheet)
    {
        static string FirstCell(List<string> row) => row.Count > 0 ? (row[0] ?? "").Trim() : "";

        while (sheet.Count > 0 && !FirstCell(sheet[0]).StartsWith("!VEND", StringComparison.OrdinalIgnoreCase))
            sheet.RemoveAt(0);
        while (sheet.Count > 1 && !FirstCell(sheet[^1]).StartsWith("VEND", StringComparison.OrdinalIgnoreCase))
            sheet.RemoveAt(sheet.Count - 1);

        // the !VEND/VEND record-type indicator is not a data column
        foreach (var row in sheet)
            if (row.Count > 0)
                row.RemoveAt(0);
    }

    // ──────────────────────────────────────────────────────────────────────
    // CSV writer + mime constants (HHM-616)
    // ──────────────────────────────────────────────────────────────────────

    // HHM-616: CSV writer (FlexKit FilerControl has no CSV writer, so the bytes are
    // produced app-side). RFC-4180 shape matching VB6's flexFileCommaText output:
    // header row first, CRLF row endings, fields quoted only when they contain a
    // comma, quote, or line break (embedded quotes doubled), no BOM.
    public byte[] BuildCsvBytes(System.Data.DataTable table)
    {
        static string Field(object? v)
        {
            var s = v == null || v is DBNull
                ? ""
                : Convert.ToString(v, System.Globalization.CultureInfo.InvariantCulture) ?? "";
            return s.IndexOfAny(CsvQuoteTriggers) >= 0
                ? "\"" + s.Replace("\"", "\"\"") + "\""
                : s;
        }
        var sb = new StringBuilder();
        sb.Append(string.Join(",", table.Columns.Cast<System.Data.DataColumn>().Select(c => Field(c.ColumnName))));
        sb.Append("\r\n");
        foreach (System.Data.DataRow row in table.Rows)
        {
            sb.Append(string.Join(",", row.ItemArray.Select(Field)));
            sb.Append("\r\n");
        }
        return new UTF8Encoding(encoderShouldEmitUTF8Identifier: false).GetBytes(sb.ToString());
    }

    private static readonly char[] CsvQuoteTriggers = { ',', '"', '\r', '\n' };

    public const string XlsxMime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    public const string CsvMime = "text/csv";

    // ──────────────────────────────────────────────────────────────────────
    // Small shared helpers
    // ──────────────────────────────────────────────────────────────────────

    // VB6 Parse(s, 1, " - ") — text before the first " - " (whole string if absent).
    private static string ParseFirst(string s, string sep)
    {
        if (string.IsNullOrEmpty(s)) return s ?? "";
        var idx = s.IndexOf(sep, StringComparison.Ordinal);
        return idx < 0 ? s : s.Substring(0, idx);
    }

    // VB6 Parse(s, n, delim) — nth (1-based) delimited token; "" if out of range.
    private static string ParsePart(string s, int n, char delim)
    {
        if (string.IsNullOrEmpty(s)) return "";
        var parts = s.Split(delim);
        return (n >= 1 && n <= parts.Length) ? parts[n - 1] : "";
    }

    // VB6 FormatExcelColRef(i) — 1-based column index -> A, B, ... Z, AA, ...
    public string ExcelColRef(int col)
    {
        var sb = new StringBuilder();
        while (col > 0)
        {
            int rem = (col - 1) % 26;
            sb.Insert(0, (char)('A' + rem));
            col = (col - 1) / 26;
        }
        return sb.ToString();
    }

    // Quote an identifier as a bracketed name, doubling any ']' so a value can never break out
    // of the [...] (e.g. "Foo]bar" -> "[Foo]]bar]").
    public string BracketId(string ident)
        => "[" + (ident ?? "").Replace("]", "]]") + "]";

    // Parse a set of AssemblyID strings into a clean, de-duplicated, order-preserving list of
    // 64-bit integers. tbldbassemblymaster.AssemblyID is an int identity column, so any value
    // that does not parse to an integer is dropped (never quoted-and-passed-through). This is
    // the sole injection defense for the IN(...) and PIVOT [id] interpolations, both of which
    // cannot be parameterized.
    public List<long> ValidIntIds(IEnumerable<string> ids)
    {
        var seen = new HashSet<long>();
        var result = new List<long>();
        if (ids == null) return result;
        foreach (var raw in ids)
        {
            var s = (raw ?? "").Trim();
            if (s.Length == 0) continue;
            if (!long.TryParse(s, System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture, out var n)) continue;
            if (seen.Add(n)) result.Add(n);
        }
        return result;
    }

    // ──────────────────────────────────────────────────────────────────────
    // IMPORT ASSEMBLIES staging  (VB6 SaveAssemblyData)
    // ──────────────────────────────────────────────────────────────────────

    // Allowlist of real ImportedAssemblies columns, fetched once per (scoped) service
    // instance and reused. Used to filter user-supplied (file-header-derived) column
    // names before they are interpolated as SQL identifiers in StageImportedAssembliesAsync.
    private HashSet<string>? _importedAssembliesColumns;
    private async Task<HashSet<string>> GetImportedAssembliesColumnsAsync()
    {
        if (_importedAssembliesColumns != null) return _importedAssembliesColumns;
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var dt = await Task.Run(() => _db.SqlExec(
            "select name from sys.columns where object_id = object_id('ImportedAssemblies')"));
        foreach (DataRow r in dt.Rows)
        {
            var n = r["name"]?.ToString();
            if (!string.IsNullOrEmpty(n)) set.Add(n);
        }
        _importedAssembliesColumns = set;
        return set;
    }

    // VB6 SaveAssemblyData: insert one row per data row into ImportedAssemblies.
    // Columns = "session,row" + every non-blank ColKey.
    public async Task StageImportedAssembliesAsync(string session, List<string> columnKeys, List<ImportRow> rows)
    {
        // The column names come from the uploaded file's header row (columnKeys), so they are
        // user-controlled. They CANNOT be parameterized (they are SQL identifiers), so:
        //   (1) keep a key only if it is a real ImportedAssemblies column (allowlist), and
        //   (2) bracket-escape it ([ -> the value with ] doubled]) as defense in depth.
        // A header that is not a known staging column is dropped — the same column set VB6's
        // insert targets, minus any junk/injection header that wouldn't match a real column.
        var allowed = await GetImportedAssembliesColumnsAsync();

        var keptCols = new List<int>();
        for (int c = 0; c < columnKeys.Count; c++)
        {
            var key = columnKeys[c];
            if (string.IsNullOrEmpty(key)) continue;
            if (!allowed.Contains(key)) continue;   // allowlist against real staging columns
            keptCols.Add(c);
        }

        var columns = new List<string> { "session", "[row]" };
        columns.AddRange(keptCols.Select(c => BracketId(columnKeys[c])));

        // Insert in batches; parameterize values (VB6 used DbQuote string concat).
        const int batchSize = 100;
        var valuesClauses = new List<string>();
        var parms = new Dictionary<string, object>();
        int pIdx = 0;
        int rowNum = 1; // VB6 stores grid-row+1; we mirror sequential row numbers

        async Task FlushAsync()
        {
            if (valuesClauses.Count == 0) return;
            var sql = $"insert ImportedAssemblies({string.Join(",", columns)}) values\n{string.Join(",\n", valuesClauses)}";
            await _db.ExecuteAsync(sql, parms);
            valuesClauses.Clear();
            parms = new Dictionary<string, object>();
            pIdx = 0;
        }

        foreach (var row in rows)
        {
            rowNum++;
            var vals = new List<string>();
            // session
            var pSession = $"@p{pIdx++}"; parms[pSession] = session; vals.Add(pSession);
            // row
            var pRow = $"@p{pIdx++}"; parms[pRow] = rowNum; vals.Add(pRow);
            // kept columns
            foreach (var c in keptCols)
            {
                var p = $"@p{pIdx++}";
                parms[p] = row.GetValue($"Col{c}") ?? "";
                vals.Add(p);
            }
            valuesClauses.Add($"({string.Join(",", vals)})");

            if (valuesClauses.Count >= batchSize)
                await FlushAsync();
        }
        await FlushAsync();
    }

    // ──────────────────────────────────────────────────────────────────────
    // IMPORT TAKEOFFS staging  (VB6 SaveAssemblyTakeoffData)
    // Pivoted layout: assemblies are columns 6.. (0-based: index 5..),
    // header rows are grid rows 0..12, item/qty data starts at grid row 14.
    // ──────────────────────────────────────────────────────────────────────

    // Helper: cell text from the raw sheet at (rowIdx, colIdx) — both 0-based.
    private static string CellAt(List<List<string>> sheetRows, int rowIdx, int colIdx)
    {
        if (rowIdx < 0 || rowIdx >= sheetRows.Count) return "";
        var r = sheetRows[rowIdx];
        if (colIdx < 0 || colIdx >= r.Count) return "";
        return (r[colIdx] ?? "").Trim();
    }

    // VB6 SaveAssemblyTakeoffData.
    //   VB6 grid rows are 1-based with a fixed header row at index 0, so VB6
    //   TextMatrix(0,c)=Excel column letter, TextMatrix(1,c)=community line, ...
    //   In our 0-based sheet (no synthetic header inserted), the mapping is:
    //     VB6 TextMatrix(0,c)  -> Excel column letter for that column
    //     VB6 TextMatrix(k,c)  -> sheetRows[k-1][c]   (header rows k=1..12)
    //     VB6 data rows r=14.. -> sheetRows[r-1]      (item rows)
    //   Assemblies span VB6 cols 6..  -> 0-based colIdx 5..
    // Returns the counts the page shows on the ImportResult stage
    // ("{N} records imported" / "{N} item quantity row(s)").
    public async Task<(int HeadersSaved, int QtyRowsSaved)> StageAssemblyTakeoffsAsync(string session, List<List<string>> sheetRows)
    {
        int headersSaved = 0;
        int qtyRowsSaved = 0;
        int maxCols = sheetRows.Count == 0 ? 0 : sheetRows.Max(r => r.Count);

        // ── headers -> ImportedAssemblyTakeoffs ───────────────────────────
        var headerRows = new List<object[]>();

        for (int col = 5; col < maxCols; col++) // VB6 c = 6 To Cols-1
        {
            // VB6 skips blank assemblies: clip of rows 1..12 for this col, spaces/CR/LF removed
            var probe = new StringBuilder();
            for (int k = 1; k <= 12; k++) probe.Append(CellAt(sheetRows, k - 1, col));
            if (string.IsNullOrWhiteSpace(probe.ToString().Replace(" ", ""))) continue;

            // VB6 TextMatrix(0,c) = Excel column letter
            var colLetter = ExcelColRef(col + 1);

            // order MUST match: session,col,Community,Model,OptionID,Assembly,Description,
            // Comments,Series,ScheduleTemplate,EstimatorNotes,SubCat,ConstCutoff,UOM
            headerRows.Add(new object[]
            {
                session,                                            // session
                colLetter,                                          // col
                ParseFirst(CellAt(sheetRows, 0, col), " - "),       // Community  (row1, "community - desc")
                CellAt(sheetRows, 1, col),                          // Model      (row2)
                CellAt(sheetRows, 2, col),                          // OptionID   (row3)
                CellAt(sheetRows, 3, col),                          // Assembly   (row4)
                CellAt(sheetRows, 4, col),                          // Description (row5)
                CellAt(sheetRows, 5, col),                          // Comments   (row6)
                ParseFirst(CellAt(sheetRows, 6, col), " - "),       // Series     (row7, "series - desc")
                CellAt(sheetRows, 7, col),                          // ScheduleTemplate (row8)
                CellAt(sheetRows, 8, col),                          // EstimatorNotes   (row9, VB6 'notes)
                ParseFirst(CellAt(sheetRows, 9, col), " - "),       // SubCat     (row10, "category - desc")
                ParseFirst(CellAt(sheetRows, 10, col), " - "),      // ConstCutoff(row11)
                CellAt(sheetRows, 11, col),                         // UOM        (row12)
            });
        }

        await ExecuteBatchedInsertAsync(
            "ImportedAssemblyTakeoffs",
            "session,col,Community,Model,OptionID,Assembly,Description,Comments,Series,ScheduleTemplate,EstimatorNotes,SubCat,ConstCutoff,UOM",
            headerRows);
        headersSaved = headerRows.Count;

        // ── unformat poindex/phase/item for data rows (VB6 r = 14 To Rows-1) ──
        // Column A (VB6 col 1 -> 0-based 0): "POIndex - description"  -> POIndex
        // Column B (VB6 col 2 -> 0-based 1): "phase\item - description (uom)" -> phase + item
        // VB6 mutates the grid; we precompute per data-row values into dictionaries.
        var poIndexByRow = new Dictionary<int, string>();
        var phaseByRow = new Dictionary<int, string>();
        var itemByRow = new Dictionary<int, string>();

        for (int r = 14; r <= sheetRows.Count; r++) // VB6 r = 14 To Rows-1 (1-based, exclusive of header)
        {
            int sIdx = r - 1; // 0-based sheet index
            if (sIdx < 0 || sIdx >= sheetRows.Count) continue;

            poIndexByRow[r] = ParseFirst(CellAt(sheetRows, sIdx, 0), " - ");   // column A
            var s = ParseFirst(CellAt(sheetRows, sIdx, 1), " - ");             // column B before " - "
            phaseByRow[r] = ParsePart(s, 1, '\\');                              // phase
            itemByRow[r] = ParsePart(s, 2, '\\');                              // item
        }

        // ── quantities -> ImportedAssemblyTakeoffItems (batched insert per assembly col) ──
        for (int col = 5; col < maxCols; col++)
        {
            // VB6 skip blank assemblies (clip rows 1..12, tabs+spaces removed)
            var probe = new StringBuilder();
            for (int k = 1; k <= 12; k++) probe.Append(CellAt(sheetRows, k - 1, col));
            if (string.IsNullOrWhiteSpace(probe.ToString().Replace(" ", ""))) continue;

            var colLetter = ExcelColRef(col + 1);
            var qtyRows = new List<object[]>();

            for (int r = 14; r <= sheetRows.Count; r++)
            {
                int sIdx = r - 1;
                if (sIdx < 0 || sIdx >= sheetRows.Count) continue;

                var qty = CellAt(sheetRows, sIdx, col);
                if (string.IsNullOrWhiteSpace(qty)) continue; // VB6 skip blank quantities

                // order: session,row,col,Community,Model,OptionID,Assembly,
                //        POIndex,Phase,Item,Invertable,Location,Notes,Qty
                qtyRows.Add(new object[]
                {
                    session,                                        // session
                    r,                                              // row
                    colLetter,                                      // col
                    ParseFirst(CellAt(sheetRows, 0, col), " - "),   // Community
                    CellAt(sheetRows, 1, col),                      // Model
                    CellAt(sheetRows, 2, col),                      // OptionID
                    CellAt(sheetRows, 3, col),                      // Assembly
                    poIndexByRow.GetValueOrDefault(r, ""),          // POIndex (col A)
                    phaseByRow.GetValueOrDefault(r, ""),            // Phase   (col B)
                    itemByRow.GetValueOrDefault(r, ""),             // Item    (col B)
                    CellAt(sheetRows, sIdx, 2),                     // Invertable (col C / VB6 col 3)
                    CellAt(sheetRows, sIdx, 3),                     // Location   (col D / VB6 col 4)
                    CellAt(sheetRows, sIdx, 4),                     // Notes      (col E / VB6 col 5)
                    qty,                                            // Qty
                });
            }

            await ExecuteBatchedInsertAsync(
                "ImportedAssemblyTakeoffItems",
                "session,[row],col,Community,Model,OptionID,Assembly,POIndex,Phase,Item,Invertable,Location,Notes,Qty",
                qtyRows);
            qtyRowsSaved += qtyRows.Count;
        }

        return (headersSaved, qtyRowsSaved);
    }
    // ──────────────────────────────────────────────────────────────────────
    // Legacy Purch_* validate / commit / cleanup wrappers (VB6 ImportAssemblies /
    // ImportAssemblyTakeoffs). The page stages first, then validates; a commit
    // runs only when the validate proc returns no rows.
    // ──────────────────────────────────────────────────────────────────────

    public DataTable ValidateImportedAssemblies(int divisionId, string session, string fileType)
        => _db.SqlExec(
            "exec Purch_ImportedAssemblies_Validate @DivisionID, @session, @filetype",
            new { DivisionID = divisionId, session, filetype = fileType });

    public async Task CommitImportedAssembliesAsync(int divisionId, string session, string fileType)
        => await _db.ExecuteAsync(
            "exec Purch_ImportedAssemblies_Commit @DivisionID, @session, @filetype",
            new { DivisionID = divisionId, session, filetype = fileType });

    // VB6 deletes its staging rows after a commit / failed validate (eh path / rerun safety).
    public async Task DeleteImportedAssembliesStagingAsync(string session)
        => await _db.ExecuteAsync("delete ImportedAssemblies where session=@session", new { session });

    public DataTable ValidateImportedAssemblyTakeoffs(int divisionId, string session)
        => _db.SqlExec(
            "exec Purch_ImportedAssemblyTakeoffs_Validate @DivisionID, @session",
            new { DivisionID = divisionId, session });

    // VB6 commit path does NOT delete the takeoff staging tables; the commit proc
    // consumes them. We leave that to the proc, matching VB6.
    public async Task CommitImportedAssemblyTakeoffsAsync(int divisionId, string session)
        => await _db.ExecuteAsync(
            "exec Purch_ImportedAssemblyTakeoffs_Commit @DivisionID, @session",
            new { DivisionID = divisionId, session });

    public async Task CleanupTakeoffStagingAsync(string session)
    {
        try
        {
            await _db.ExecuteAsync("delete ImportedAssemblyTakeoffs where session=@session", new { session });
            await _db.ExecuteAsync("delete ImportedAssemblyTakeoffItems where session=@session", new { session });
        }
        catch { /* best effort cleanup */ }
    }

    // ──────────────────────────────────────────────────────────────────────
    // EXPORT — table builders
    // ──────────────────────────────────────────────────────────────────────

    // ══ HHM-440 ExportVendors — SELECTION-BASED as of the 2026-07-22 owner
    // redesign. The VB6 FVendor "ExcelExport" behavior (HFApp.RunTask(
    // "ExportData|select * from tblVendors where divisionid = N") — a whole-
    // division dump with no picklist) was deliberately REPLACED: the user now
    // multi-selects vendors in a Company/Vendor FPickList (modeled on FVendor's
    // vendor-list picker) and only those rows are exported. The old dump code
    // path was removed — picking nothing simply cancels back to ExportPick. ══
    //
    // Same column shape as the dump (SELECT * FROM tblVendors), scoped to the
    // division and the selected Vendor_IDs. Fully parameterized: @DivisionID +
    // one @v{n} parameter per selected id (the db wrapper's dictionary pattern).
    public DataTable BuildVendorExportTable(int divisionId, List<string> vendorIds)
    {
        var parms = new Dictionary<string, object> { ["@DivisionID"] = divisionId };
        var inNames = new List<string>();
        int i = 0;
        foreach (var id in vendorIds)
        {
            var p = $"@v{i++}";
            parms[p] = id;
            inNames.Add(p);
        }
        var sql = "select * from tblVendors where divisionid = @DivisionID" +
                  $" and Vendor_ID in ({string.Join(",", inNames)}) order by Vendor_Name";
        return _db.SqlExec(sql, parms);
    }

    // VB6 ExportAssemblies — three SELECT shapes keyed off FPickList.SelectedView
    // (models / options / globals). The view is chosen explicitly in the UI and the list is
    // filtered to that AssemblyType, so the column set is driven DIRECTLY off the chosen view
    // (no guessing from a mixed selection — the old ResolveExportAssemblyType path).
    // 2026-07-22 redesign: builds and RETURNS the DataTable only — the save happens
    // in the page's ExportSave stage (Download / Save As…).
    public DataTable BuildAssembliesExportTable(List<long> assemblyIds, ExportView view)
    {
        // assemblyIds are pre-validated integers (ValidIntIds, applied in OnExportPicked), so an
        // unquoted comma list is injection-safe here.
        var idList = string.Join(",", assemblyIds);

        // VB6 SelectedView 1/2/3 → output shape. Driven off the explicit view selector.
        var type = view switch
        {
            ExportView.ModelOptions => "options",
            ExportView.GlobalOptions => "globals",
            _ => "models"
        };

        string sql = type switch
        {
            "options" => $@"
select
 a.Community CommunityCode, a.Model, a.OptionID [Option], a.Assembly
,case when isnull(a.Inactive,0)=0 then 'n' else 'y' end Inactive
,a.Description
,a.GraphicPath Image
,a.Comments
,a.Notes EstimatorNotes
,a.Series
,a.Category SubCat
,a.JCExtra
,a.Qty
,a.AssemblyUOM UOM
,a.ConstCutoff
,a.Location
,case when isnull(a.IncludedOption,0)=0 then 'n' else 'y' end IncludedOption
,case when isnull(a.DesignCenterSalesOnly,0)=0 then 'n' else 'y' end DesignCenterSalesOnly
,case when isnull(a.SelectByRoom,0)=0 then 'n' else 'y' end SelectByRoom
,case when isnull(a.DisplayTotalOnly,0)=0 then 'n' else 'y' end DisplayTotalOnly
,case when isnull(a.TakeoffRequired,0)=0 then 'n' else 'y' end TakeoffRequired
,case when isnull(a.UseNormalSalesQtyFactors,0)=0 then 'n' else 'y' end UseNormalSalesQtyFactors
,clist.name ColorListName
,slist.name StyleListName
,flist.name FinishListName
,olist.name OtherListName
,a.Color
,a.StyleValue
,a.FinishValue
,a.OtherValue
from tbldbassemblymaster a
left join AttributeLists clist on a.ColorListID=cList.listid
left join AttributeLists slist on a.StyleListID=sList.listid
left join AttributeLists flist on a.FinishListID=fList.listid
left join AttributeLists olist on a.OtherListID=oList.listid
where a.AssemblyID in({idList})",

            "globals" => $@"
select
 a.Community CommunityCode, a.OptionID [Option], a.Assembly
,case when isnull(a.Inactive,0)=0 then 'n' else 'y' end Inactive
,a.Description
,a.GraphicPath Image
,a.Comments
,a.Notes EstimatorNotes
,a.Category SubCat
,a.JCExtra
,a.Qty
,a.AssemblyUOM UOM
,a.ConstCutoff
,a.Location
,case when isnull(a.IncludedOption,0)=0 then 'n' else 'y' end IncludedOption
,case when isnull(a.DesignCenterSalesOnly,0)=0 then 'n' else 'y' end DesignCenterSalesOnly
,case when isnull(a.SelectByRoom,0)=0 then 'n' else 'y' end SelectByRoom
,case when isnull(a.DisplayTotalOnly,0)=0 then 'n' else 'y' end DisplayTotalOnly
,case when isnull(a.TakeoffRequired,0)=0 then 'n' else 'y' end TakeoffRequired
,case when isnull(a.UseNormalSalesQtyFactors,0)=0 then 'n' else 'y' end UseNormalSalesQtyFactors
,clist.name ColorListName
,slist.name StyleListName
,flist.name FinishListName
,olist.name OtherListName
,a.Color
,a.StyleValue
,a.FinishValue
,a.OtherValue
from tbldbassemblymaster a
left join AttributeLists clist on a.ColorListID=cList.listid
left join AttributeLists slist on a.StyleListID=sList.listid
left join AttributeLists flist on a.FinishListID=fList.listid
left join AttributeLists olist on a.OtherListID=oList.listid
where a.AssemblyID in({idList})",

            _ /* models */ => $@"
select
 Community CommunityCode, Model, Assembly
,case when isnull(Inactive,0)=0 then 'n' else 'y' end Inactive
,Description
,GraphicPath Image
,Comments
,Series
,Style
,ScheduleTemplate
,Bedrooms
,Bathrooms
,FloorArea
,MaxWidth
,MaxLength
,SpecDocument
,case when isnull(TakeoffRequired,0)=0 then 'n' else 'y' end TakeoffRequired
,case when isnull(UseNormalSalesQtyFactors,0)=0 then 'n' else 'y' end UseNormalSalesQtyFactors
,MainSF
,UpperSF
,LowerSF
,GarageSF
from tbldbassemblymaster
where AssemblyID in({idList})"
        };

        return _db.SqlExec(sql);
    }

    // ──────────────────────────────────────────────────────────────────────
    // EXPORT TAKEOFFS — fixed-template workbook builder
    // ──────────────────────────────────────────────────────────────────────

    // VB6 ExportAssemblyTakeoffs (MAssemblyExport) — copies NewAssmImports.xlsx and writes
    // into two worksheets via Excel automation. We reproduce that template faithfully with
    // ClosedXML instead of copying the .xlsx (so the output is self-contained and matches the
    // exact sheet/validation/header layout the import side expects):
    //   • Sheet 1 "Assemblies"   : instructional + label cells, A13:E13/F13.. column headers,
    //                              one assembly-context column (rows 1-12) per selected assembly
    //                              starting at col F, and item/qty data rows from row 14.
    //   • Sheet 2 "DataValidation": 7 list columns (BPTemplate/Item/SubCat/JobStatus/Series/
    //                              Communities/POIndex), header in row 1, list-driven dropdowns
    //                              on the matching data cells of sheet 1, sheet protected ("admin").
    // assemblyIds are pre-validated integers.
    // 2026-07-22 redesign: builds and RETURNS the workbook bytes only — the save happens
    // in the page's ExportSave stage (the takeoff template's columns are structural,
    // so the preview shows a fixed-layout note instead of a checklist).
    // The result also carries the pivoted quantities relabelled for the ExportSave
    // preview (owner directive 2026-07-27) and the qty-row count for the result page.
    public async Task<TakeoffExportResult> BuildTakeoffExportAsync(int divisionId, List<long> assemblyIds, string templatePath)
    {
        var idList = string.Join(",", assemblyIds);
        // PIVOT column identifiers. assemblyIds are validated integers, so [id] is safe.
        var pivotCols = string.Join(",", assemblyIds.Select(id => $"[{id}]"));

        // ── 1) Header context per assembly (VB6 sheet-1 rows 1..12), order by 1,2,3,4 ──
        // top 500 mirrors VB6 (template grid caps assemblies at 500).
        var headerSql = $@"
select top 500
 l.area + ' - ' + isnull(l.description,'') Community
,a.Model
,a.OptionID
,a.Assembly
,a.AssemblyID
,a.Description
,a.Comments
,a.notes
,s.series + ' - ' + isnull(s.description,'') series
,a.scheduletemplate
,case when a.assemblytype=0 then '' else c.category + ' - ' + isnull(c.description,'') end category
,case when a.assemblytype=0 then '' else cast(j.job_status as varchar(10)) + ' - ' + isnull(j.description,'') end constcutoff
,case when a.assemblytype=0 then '' else a.assemblyuom end assemblyuom
from tbldbassemblymaster a
left join tblseries s on a.divisionid=s.divisionid and a.series=s.series
left join tblcategories c on a.category=c.category
left join tbljobstatus j on a.constcutoff=j.job_status
left join tbllocality l on a.community=l.area
where a.AssemblyID in({idList})
order by 1,2,3,4";
        var headers = _db.SqlExec(headerSql);

        // The PIVOT column order MUST match the header order (VB6 builds PivotColumns while
        // walking the same ordered header recordset). Re-derive the pivot column list from the
        // ordered AssemblyIDs returned above, keeping only validated integers.
        var orderedIds = ValidIntIds(headers.AsEnumerable().Select(r => r["AssemblyID"]?.ToString()));
        if (orderedIds.Count > 0)
            pivotCols = string.Join(",", orderedIds.Select(id => $"[{id}]"));

        // ── 2) Pivoted takeoff quantities (VB6 sheet-1 rows 14.., assemblies across cols) ──
        var sql = $@"
select *
from
(
    select
     p.poindex + ' - ' + isnull(p.description,'') poindex
    ,d.phase + '\' + d.item + ' - ' + isnull(i.description,'') + isnull(' (' + i.takeoffuom + ')','') item
    ,case when isnull(d.invertable,0)=1 then 'y' else '' end invertable
    ,d.location
    ,d.notes
    ,d.takeoffqty
    ,a.AssemblyID
    from tbldbassemblymaster a
    join tbldbassemblydetails d on a.assemblyid=d.assemblyid
    left join tblphaseitem i on a.divisionid=i.divisionid and d.phase=i.phase and d.item=i.item
    left join tblpoindex p on d.divisionid=p.divisionid and d.poindex=p.poindex
    where a.AssemblyID in({idList})
) p
pivot
(
    sum(takeoffqty)
    for assemblyid in({pivotCols})
) as pvt
order by 2,1,3,4,5";
        var dt = _db.SqlExec(sql);

        // Surface the pivoted takeoff quantities as the ExportSave preview table so the
        // step is NOT a blank screen for takeoffs. This is a RENAMED COPY — the raw pivot
        // headers are the AssemblyIDs, so each per-assembly quantity column is relabelled
        // to its "Model / Option / Assembly" identity (from the `headers` recordset) for
        // the preview only; the workbook build below still reads the id-keyed `dt`. The
        // columns are structural (fixed template), so the header checkboxes render disabled
        // — see the page's ExportSave stage. owner directive 2026-07-27.
        var previewTable = BuildTakeoffPreviewTable(dt, headers);

        // ── 3) Build the multi-sheet template workbook (downloaded later by the page) ──
        var bytes = await BuildTakeoffTemplateWorkbookAsync(divisionId, templatePath, headers, orderedIds, dt);
        var qtyRowCount = dt.AsEnumerable().Count(r =>
            !string.IsNullOrWhiteSpace(r.Table.Columns.Contains("item") ? r["item"]?.ToString() : ""));
        return new TakeoffExportResult(bytes, previewTable, qtyRowCount);
    }

    // Preview-only copy of the pivoted takeoff table (owner directive 2026-07-27). The raw
    // pivot names its per-assembly quantity columns by AssemblyID ("1512"), which is opaque
    // in the ExportSave preview. Relabel each to its "Model / Option / Assembly" identity
    // from the header recordset, and title-case the fixed leading columns. The workbook
    // build keeps using the ORIGINAL id-keyed pivot — this copy never feeds the export.
    private static DataTable BuildTakeoffPreviewTable(DataTable pivot, DataTable headers)
    {
        // AssemblyID → friendly label (non-empty of Model / Option / Assembly, id fallback).
        var labelById = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (DataRow hr in headers.Rows)
        {
            var id = headers.Columns.Contains("AssemblyID") ? hr["AssemblyID"]?.ToString() : null;
            if (string.IsNullOrWhiteSpace(id)) continue;
            var parts = new[] { "Model", "OptionID", "Assembly" }
                .Where(headers.Columns.Contains)
                .Select(c => hr[c]?.ToString()?.Trim())
                .Where(s => !string.IsNullOrEmpty(s));
            var label = string.Join(" / ", parts);
            labelById[id] = string.IsNullOrEmpty(label) ? id : label;
        }

        // Fixed leading columns → the export's row-13 header labels.
        var leadRename = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["poindex"] = "PO Index", ["item"] = "Item", ["invertable"] = "Invertable",
            ["location"] = "Location", ["notes"] = "Notes",
        };

        var copy = pivot.Copy();

        // Compute globally-unique target names (DataColumn names must be unique).
        var finalNames = new List<string>(copy.Columns.Count);
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (DataColumn c in copy.Columns)
        {
            var target = leadRename.TryGetValue(c.ColumnName, out var lead) ? lead
                       : labelById.TryGetValue(c.ColumnName, out var lbl) ? lbl
                       : c.ColumnName;
            if (string.IsNullOrWhiteSpace(target)) target = c.ColumnName;
            var unique = target;
            for (int n = 2; !used.Add(unique); n++) unique = $"{target} ({n})";
            finalNames.Add(unique);
        }

        // Two-phase rename (temp names first) so a new name equal to another column's OLD
        // name never triggers a transient DuplicateNameException.
        for (int i = 0; i < copy.Columns.Count; i++) copy.Columns[i].ColumnName = "__c" + i;
        for (int i = 0; i < copy.Columns.Count; i++) copy.Columns[i].ColumnName = finalNames[i];
        return copy;
    }

    // ──────────────────────────────────────────────────────────────────────
    // Multi-sheet .xlsx export template — faithful ClosedXML reproduction of
    // Estimating\Templates\NewAssmImports.xlsx (VB6 ExportAssemblyTakeoffs).
    //
    // Sheet 1 "Assemblies" fixed layout (from the shipped template):
    //   A1  : POIndex instructional note
    //   C1  : Invertable instructional note
    //   E1..E12 : the 12 row labels (Community, Model, Option, Assembly (version),
    //             Description, Comments, Series, BP Template (model only),
    //             Estimator Notes (option only), Option Subcategory (option only),
    //             Construction Cutoff (option only), Sales UOM (option only))
    //   A13:F13.. : column headers — POIndex, Item, Invertable, Location, Notes, then "Qty"
    //   Assembly context columns: F (col 6) onward, rows 1..12.
    //   Item/qty data: rows 14.. ; cols A..E = poindex,item,invertable,location,notes;
    //                  cols F.. = the per-assembly quantity.
    //
    // Sheet 2 "DataValidation" : 7 list columns (header row 1, values from row 2):
    //   A BPTemplate | B Item | C SubCat | D JobStatus | E Series | F Communities | G POIndex
    //   Sheet is protected with password "admin" (VB6 xlSH.Protect "admin",...).
    //
    // Data validation dropdowns on sheet 1 (matching the template's x14 list rules):
    //   Communities -> F1:.. (row 1, community header), src DataValidation!F
    //   Series      -> F7:.. (row 7),                   src DataValidation!E
    //   BPTemplate  -> F8:.. (row 8),                   src DataValidation!A
    //   SubCat      -> F10:..(row 10),                  src DataValidation!C
    //   JobStatus   -> F11:..(row 11),                  src DataValidation!D
    //   Item        -> B14:..(item data col),           src DataValidation!B
    //   POIndex     -> A14:..(poindex data col),        src DataValidation!G
    //   Invertable  -> C14:..(invertable data col),     inline list "y,n"
    // ──────────────────────────────────────────────────────────────────────
    private const string TEMPLATE_SHEET_DATA = "Assemblies";
    private const string TEMPLATE_SHEET_VALIDATION = "DataValidation";
    private const string TEMPLATE_PROTECT_PASSWORD = "admin";   // VB6 xlSH.Protect "admin"
    private const int TEMPLATE_FIRST_ASSEMBLY_COL = 6;          // col F (VB6 c = 5 + i)
    private const int TEMPLATE_DATA_START_ROW = 14;            // VB6 RowDataStart=13, then r=r+1
    private const int TEMPLATE_LIST_LAST_ROW = 55555;          // VB6 list ranges $2:$55555

    // Owner directive 2026-07-24: exports must use the SHIPPED templates under
    // wwwroot/resources/Estimating/Templates — the web mapping of VB6
    // HFApp.SystemFolder & "Estimating\Templates" (MAssemblyImport.bas:221
    // srcXLS = PathAppend(HFApp.SystemFolder, "Estimating\Templates\NewAssmImports.xlsx")).
    // The page resolves the path (it owns IWebHostEnvironment) and passes it in.
    private async Task<byte[]> BuildTakeoffTemplateWorkbookAsync(
        int divisionId, string templatePath, DataTable headers, List<long> orderedIds, DataTable pivot)
    {
        // Load the shipped NewAssmImports.xlsx template exactly like VB6 and fill
        // it in; the code-built reproduction below remains only as a fallback for
        // a missing/corrupt template file (logged, so a bad deploy is visible).
        ClosedXML.Excel.XLWorkbook wb;
        try
        {
            wb = File.Exists(templatePath)
                ? new ClosedXML.Excel.XLWorkbook(templatePath)
                : new ClosedXML.Excel.XLWorkbook();
            if (!File.Exists(templatePath))
                _logger.LogWarning("[{Src}] Takeoff export template not found at {Path} — using the code-built reproduction", SRCFILE, templatePath);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "[{Src}] Takeoff export template unreadable at {Path} — using the code-built reproduction", SRCFILE, templatePath);
            wb = new ClosedXML.Excel.XLWorkbook();
        }
        using var _ = wb;

        // Sheet order matches the template: Assemblies first, DataValidation second.
        // When the shipped template loaded, reuse ITS sheets (all their formatting,
        // print setup and validation chrome survive); the writes below only fill in
        // the dynamic data over the template's fixed labels.
        if (!wb.Worksheets.TryGetWorksheet(TEMPLATE_SHEET_DATA, out var ws))
            ws = wb.Worksheets.Add(TEMPLATE_SHEET_DATA);
        if (!wb.Worksheets.TryGetWorksheet(TEMPLATE_SHEET_VALIDATION, out var dv))
            dv = wb.Worksheets.Add(TEMPLATE_SHEET_VALIDATION);

        // The shipped template's validation sheet is protected ("admin") — writes
        // below would throw. Unprotect first; PopulateValidationSheetAsync re-protects.
        try { dv.Unprotect(TEMPLATE_PROTECT_PASSWORD); } catch { /* not protected */ }

        // ── Sheet 2: the seven validation lists (built first so sheet-1 dropdowns resolve) ──
        await PopulateValidationSheetAsync(dv, divisionId);

        // ── Sheet 1: instructional + label cells ──
        ws.Cell(1, 1).Value = "POIndex is optional. It should NOT be provided unless it is different than the default value in the item db. Do not remove this or any column from the spreadsheet.";
        ws.Cell(1, 3).Value = "When plan is inverted, is this item replaced with it's inverse?";
        string[] rowLabels =
        {
            "Community", "Model", "Option", "Assembly (version)", "Description", "Comments",
            "Series", "BP Template (model only)", "Estimator Notes (option only)",
            "Option Subcategory (option only)", "Construction Cutoff (option only)",
            "Sales UOM (option only)"
        };
        for (int i = 0; i < rowLabels.Length; i++)        // E1..E12
            ws.Cell(i + 1, 5).Value = rowLabels[i];

        // column headers (row 13): A POIndex, B Item, C Invertable, D Location, E Notes, F.. Qty
        ws.Cell(13, 1).Value = "POIndex";
        ws.Cell(13, 2).Value = "Item";
        ws.Cell(13, 3).Value = "Invertable";
        ws.Cell(13, 4).Value = "Location";
        ws.Cell(13, 5).Value = "Notes";
        int assemblyCount = orderedIds.Count;
        for (int i = 0; i < assemblyCount; i++)
            ws.Cell(13, TEMPLATE_FIRST_ASSEMBLY_COL + i).Value = "Qty";

        // ── Sheet 1: per-assembly context columns (rows 1..12, col F..) ──
        // header column order must match the pivot column order (orderedIds), so index headers
        // by AssemblyID. VB6 leaves a cell blank when the value is empty.
        var headerById = new Dictionary<long, DataRow>();
        foreach (DataRow hr in headers.Rows)
        {
            if (long.TryParse(hr["AssemblyID"]?.ToString(), out var hid))
                headerById[hid] = hr;
        }
        for (int i = 0; i < assemblyCount; i++)
        {
            int col = TEMPLATE_FIRST_ASSEMBLY_COL + i;
            if (!headerById.TryGetValue(orderedIds[i], out var hr)) continue;

            // VB6 row map: 1 Community,2 Model,3 Option,4 Assembly,5 Description,6 Comments,
            //              7 Series,8 ScheduleTemplate,9 Notes,10 Category,11 ConstCutoff,12 AssemblyUOM
            SetIfNotBlank(ws, 1, col, hr["Community"]);
            SetIfNotBlank(ws, 2, col, hr["Model"]);
            SetIfNotBlank(ws, 3, col, hr["OptionID"]);
            SetIfNotBlank(ws, 4, col, hr["Assembly"]);
            SetIfNotBlank(ws, 5, col, hr["Description"]);
            SetIfNotBlank(ws, 6, col, hr["Comments"]);
            SetIfNotBlank(ws, 7, col, hr["series"]);
            SetIfNotBlank(ws, 8, col, hr["scheduletemplate"]);
            SetIfNotBlank(ws, 9, col, hr["notes"]);
            SetIfNotBlank(ws, 10, col, hr["category"]);
            SetIfNotBlank(ws, 11, col, hr["constcutoff"]);
            SetIfNotBlank(ws, 12, col, hr["assemblyuom"]);
        }

        // ── Sheet 1: item/qty data rows (14..) ──
        // pivot columns: [0]=poindex, [1]=item, [2]=invertable, [3]=location, [4]=notes,
        // then one column per AssemblyID (named "<id>"). VB6 skips rows with a blank item.
        int r = TEMPLATE_DATA_START_ROW - 1;   // VB6: r = RowDataStart(13); r=r+1 before write
        foreach (DataRow pr in pivot.Rows)
        {
            var item = pr.Table.Columns.Contains("item") ? pr["item"]?.ToString() ?? "" : "";
            if (string.IsNullOrWhiteSpace(item)) continue;   // VB6: If "" & rs("item") <> ""
            r++;

            SetIfNotBlank(ws, r, 1, GetCol(pr, "poindex"));   // VB6 "'" & poindex → Text
            ws.Cell(r, 2).SetValue(item).Style.NumberFormat.Format = "@"; // VB6 "'" & item → Text (always written)
            SetIfNotBlank(ws, r, 3, GetCol(pr, "invertable"), forceText: false); // VB6 "" & invertable → not forced
            SetIfNotBlank(ws, r, 4, GetCol(pr, "location"));  // VB6 "'" & location → Text
            SetIfNotBlank(ws, r, 5, GetCol(pr, "notes"));     // VB6 "'" & notes → Text

            for (int i = 0; i < assemblyCount; i++)
            {
                var qcol = orderedIds[i].ToString();          // pivot column header is the id
                if (!pr.Table.Columns.Contains(qcol)) continue;
                var q = pr[qcol];
                if (q == null || q == DBNull.Value) continue;  // VB6 skips blank qty
                var qs = q.ToString();
                if (string.IsNullOrWhiteSpace(qs)) continue;
                int c = TEMPLATE_FIRST_ASSEMBLY_COL + i;
                if (decimal.TryParse(qs, out var qd)) ws.Cell(r, c).Value = qd;
                else ws.Cell(r, c).Value = qs;
            }
        }
        int lastDataRow = Math.Max(r, TEMPLATE_DATA_START_ROW);

        // ── Sheet 1: data-validation dropdowns (mirror the template's x14 list rules) ──
        int lastAssemblyCol = assemblyCount > 0
            ? TEMPLATE_FIRST_ASSEMBLY_COL + assemblyCount - 1
            : TEMPLATE_FIRST_ASSEMBLY_COL;

        // header-row dropdowns span the assembly columns F..lastAssemblyCol
        AddListValidation(ws, 1, TEMPLATE_FIRST_ASSEMBLY_COL, 1, lastAssemblyCol, "F"); // Communities
        AddListValidation(ws, 7, TEMPLATE_FIRST_ASSEMBLY_COL, 7, lastAssemblyCol, "E"); // Series
        AddListValidation(ws, 8, TEMPLATE_FIRST_ASSEMBLY_COL, 8, lastAssemblyCol, "A"); // BPTemplate
        AddListValidation(ws, 10, TEMPLATE_FIRST_ASSEMBLY_COL, 10, lastAssemblyCol, "C"); // SubCat
        AddListValidation(ws, 11, TEMPLATE_FIRST_ASSEMBLY_COL, 11, lastAssemblyCol, "D"); // JobStatus

        // data-column dropdowns span rows 14..lastDataRow
        AddListValidation(ws, TEMPLATE_DATA_START_ROW, 1, lastDataRow, 1, "G");  // POIndex (col A)
        AddListValidation(ws, TEMPLATE_DATA_START_ROW, 2, lastDataRow, 2, "B");  // Item    (col B)

        // Invertable (col C) — inline "y,n" list (template's standard dataValidation)
        var inv = ws.Range(TEMPLATE_DATA_START_ROW, 3, lastDataRow, 3).CreateDataValidation();
        inv.List("\"y,n\"", true);
        inv.IgnoreBlanks = true;

        // ── protect the validation sheet (VB6 xlSH.Protect "admin", ...) ──
        dv.Columns(1, 7).AdjustToContents();   // VB6 xlSH.columns("A:I").AutoFit
        dv.Protect(TEMPLATE_PROTECT_PASSWORD);

        // Owner 2026-07-24: do NOT auto-fit the Assemblies sheet — VB6's ONLY
        // AutoFit is the validation sheet's A:I (MAssemblyImport.bas:345). The
        // Assemblies sheet keeps the shipped template's column widths verbatim
        // (A=3.2, B=54, E=28, Qty cols=8…); the previous AdjustToContents here
        // blew them out (A→42, B→118) and broke template fidelity. Assembly
        // columns beyond the template's pre-formatted range get the template's
        // Qty width so a 30+-assembly export still looks right.
        const double TEMPLATE_QTY_COL_WIDTH = 8.0;
        for (int c = TEMPLATE_FIRST_ASSEMBLY_COL; c <= lastAssemblyCol; c++)
        {
            var cw = ws.Column(c).Width;
            if (cw <= 0 || cw > 60)             // untouched/default beyond template range
                ws.Column(c).Width = TEMPLATE_QTY_COL_WIDTH;
        }

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        return ms.ToArray();
    }

    // Fill the 7 list columns on the DataValidation sheet (VB6 WriteColumn calls).
    private async Task PopulateValidationSheetAsync(ClosedXML.Excel.IXLWorksheet dv, int divisionId)
    {
        // headers (row 1) — matches the shipped template
        dv.Cell(1, 1).Value = "BPTemplate";
        dv.Cell(1, 2).Value = "Item";
        dv.Cell(1, 3).Value = "SubCat";
        dv.Cell(1, 4).Value = "JobStatus";
        dv.Cell(1, 5).Value = "Series";
        dv.Cell(1, 6).Value = "Communities";
        dv.Cell(1, 7).Value = "POIndex";

        var div = new { DivisionID = divisionId };

        // col 1: BPTemplates — AppOptions ScheduleTemplates, pipe-delimited (VB6 Replace "|"->",")
        var schedTemplates = await _db.ExecuteScalarAsync<string>(
            "select OptionValue from AppOptions where OptionName='ScheduleTemplates' and DivisionID in (0,@DivisionID) order by DivisionID desc",
            div) ?? "";
        WriteListColumn(dv, 1, (schedTemplates ?? "")
            .Split('|', StringSplitOptions.RemoveEmptyEntries)
            .Select(x => x.Trim())
            .Where(x => x.Length > 0));

        // col 2: Item — "phase\item - description (orderuom)"
        var items = _db.SqlExec(
            "select phase + '\\' + item + ' - ' + isnull(description,'') + isnull(' (' + orderuom + ')','') v from tblphaseitem where divisionid=@DivisionID order by 1",
            div);
        WriteListColumn(dv, 2, items.AsEnumerable().Select(x => x[0]?.ToString() ?? ""));

        // col 3: SubCat — "category - description"
        var cats = _db.SqlExec(
            "select category + ' - ' + isnull(description,'') v from tblcategories order by 1");
        WriteListColumn(dv, 3, cats.AsEnumerable().Select(x => x[0]?.ToString() ?? ""));

        // col 4: JobStatus — "status - description"
        var statuses = _db.SqlExec(
            "select cast(job_status as varchar(10)) + ' - ' + isnull(description,'') v from tbljobstatus order by job_status");
        WriteListColumn(dv, 4, statuses.AsEnumerable().Select(x => x[0]?.ToString() ?? ""));

        // col 5: Series — "series - description"
        var series = _db.SqlExec(
            "select series + ' - ' + isnull(description,'') v from tblseries where divisionid=@DivisionID order by 1",
            div);
        WriteListColumn(dv, 5, series.AsEnumerable().Select(x => x[0]?.ToString() ?? ""));

        // col 6: Communities — "community - description" (joined to DivisionCommunities)
        var communities = _db.SqlExec(
            "select area + ' - ' + isnull(description,'') v from tbllocality l join DivisionCommunities d on l.area=d.community where d.divisionid=@DivisionID order by 1",
            div);
        WriteListColumn(dv, 6, communities.AsEnumerable().Select(x => x[0]?.ToString() ?? ""));

        // col 7: POIndex — "poindex - description"
        var poindexes = _db.SqlExec(
            "select poindex + ' - ' + isnull(description,'') v from tblpoindex where divisionid=@DivisionID order by 1",
            div);
        WriteListColumn(dv, 7, poindexes.AsEnumerable().Select(x => x[0]?.ToString() ?? ""));
    }

    // Write a validation-list column: header is already in row 1; values fill from row 2.
    // VB6 WriteColumn wrote each list value as "'" & rs(0) — force Excel Text so numeric
    // codes (POIndex/Series/JobStatus/etc.) stay verbatim and dropdown matches don't drift.
    private static void WriteListColumn(ClosedXML.Excel.IXLWorksheet dv, int col, IEnumerable<string> values)
    {
        int row = 1;
        foreach (var v in values)
        {
            row++;
            dv.Cell(row, col).SetValue(v ?? "").Style.NumberFormat.Format = "@";
        }
    }

    // Add a cross-sheet list dropdown on sheet 1 over the given rectangle, sourcing column
    // <srcCol> of the DataValidation sheet (rows 2..TEMPLATE_LIST_LAST_ROW).
    private static void AddListValidation(ClosedXML.Excel.IXLWorksheet ws, int r1, int c1, int r2, int c2, string srcCol)
    {
        var validation = ws.Range(r1, c1, r2, c2).CreateDataValidation();
        validation.List($"={TEMPLATE_SHEET_VALIDATION}!${srcCol}$2:${srcCol}${TEMPLATE_LIST_LAST_ROW}", true);
        validation.IgnoreBlanks = true;
    }

    // Null-safe column read: "" when the column is absent or the value is DBNull —
    // VB6 "" & rs(col) semantics for the pivot's optional columns.
    private static string GetCol(DataRow row, string col)
        => row.Table.Columns.Contains(col) ? row[col]?.ToString() ?? "" : "";

    // VB6 writes a cell only when the value is non-blank (leaves it empty otherwise).
    // forceText mirrors VB6's leading-apostrophe ("'" & rs(...)) on coded/context cells:
    // it pins the cell to Excel Text so a purely-numeric code (POIndex, Series, Option,
    // Community, etc.) is never coerced to a number/date by ClosedXML's type sniffing.
    // Pass forceText:false for cells VB6 wrote WITHOUT the apostrophe (e.g. Invertable).
    private static void SetIfNotBlank(ClosedXML.Excel.IXLWorksheet ws, int row, int col, object? value, bool forceText = true)
    {
        var s = value?.ToString();
        if (!string.IsNullOrEmpty(s))
        {
            var cell = ws.Cell(row, col);
            if (forceText)
                cell.SetValue(s).Style.NumberFormat.Format = "@";   // VB6 "'" & … → Text
            else
                cell.Value = s;
        }
    }



    // SQL Server caps a single batch at 2100 parameters. Each row here uses 14 params,
    // so any INSERT with more than ~150 rows in one statement would throw
    // "The incoming request has too many parameters." This helper chunks rows into
    // batches that stay under that cap, so the loop above can just keep adding rows
    // without worrying about how many the sheet ends up producing.

    private async Task ExecuteBatchedInsertAsync(string tableName, string columnsClause, List<object[]> rows)
    {
        if (rows.Count == 0) return;

        int paramsPerRow = rows[0].Length;
        int rowsPerBatch = Math.Max(1, MaxSqlParams / paramsPerRow);

        for (int i = 0; i < rows.Count; i += rowsPerBatch)
        {
            var batch = rows.Skip(i).Take(rowsPerBatch).ToList();
            var valueClauses = new List<string>(batch.Count);
            var parms = new Dictionary<string, object>();
            int p = 0;

            foreach (var row in batch)
            {
                var names = new string[row.Length];
                for (int c = 0; c < row.Length; c++)
                {
                    var pname = $"@p{p++}";
                    parms[pname] = row[c] ?? "";
                    names[c] = pname;
                }
                valueClauses.Add($"({string.Join(",", names)})");
            }

            var sql = $"insert {tableName}({columnsClause}) values\n{string.Join(",\n", valueClauses)}";
            await _db.ExecuteAsync(sql, parms);
        }
    }

}

// ── Export Assemblies output shape (VB6 ExportAssemblies / FPickList.SelectedView 1/2/3:
//    models / options / globals). The page sets it from the export FPickList's
//    SelectedViewIndex and passes it into BuildAssembliesExportTable. ──
public enum ExportView { Models, ModelOptions, GlobalOptions }

// Result of BuildTakeoffExportAsync: the finished template workbook, the
// relabelled pivot for the ExportSave preview, and the item-quantity row count
// the result page reports.
public sealed record TakeoffExportResult(byte[] Bytes, DataTable PreviewTable, int QtyRowCount);

// One parsed import row for the BASE modes' preview grid + staging (the page's
// GridControl binds Col0..Col39; StageImportedAssembliesAsync reads GetValue).
public class ImportRow
{
    public int RowIndex { get; set; }
    public Dictionary<string, string> Values { get; set; } = new();
    public string GetValue(string key) => Values.TryGetValue(key, out var v) ? v ?? "" : "";

    public string Col0 { get => GetValue("Col0"); set => Values["Col0"] = value; }
    public string Col1 { get => GetValue("Col1"); set => Values["Col1"] = value; }
    public string Col2 { get => GetValue("Col2"); set => Values["Col2"] = value; }
    public string Col3 { get => GetValue("Col3"); set => Values["Col3"] = value; }
    public string Col4 { get => GetValue("Col4"); set => Values["Col4"] = value; }
    public string Col5 { get => GetValue("Col5"); set => Values["Col5"] = value; }
    public string Col6 { get => GetValue("Col6"); set => Values["Col6"] = value; }
    public string Col7 { get => GetValue("Col7"); set => Values["Col7"] = value; }
    public string Col8 { get => GetValue("Col8"); set => Values["Col8"] = value; }
    public string Col9 { get => GetValue("Col9"); set => Values["Col9"] = value; }
    public string Col10 { get => GetValue("Col10"); set => Values["Col10"] = value; }
    public string Col11 { get => GetValue("Col11"); set => Values["Col11"] = value; }
    public string Col12 { get => GetValue("Col12"); set => Values["Col12"] = value; }
    public string Col13 { get => GetValue("Col13"); set => Values["Col13"] = value; }
    public string Col14 { get => GetValue("Col14"); set => Values["Col14"] = value; }
    public string Col15 { get => GetValue("Col15"); set => Values["Col15"] = value; }
    public string Col16 { get => GetValue("Col16"); set => Values["Col16"] = value; }
    public string Col17 { get => GetValue("Col17"); set => Values["Col17"] = value; }
    public string Col18 { get => GetValue("Col18"); set => Values["Col18"] = value; }
    public string Col19 { get => GetValue("Col19"); set => Values["Col19"] = value; }
    public string Col20 { get => GetValue("Col20"); set => Values["Col20"] = value; }
    public string Col21 { get => GetValue("Col21"); set => Values["Col21"] = value; }
    public string Col22 { get => GetValue("Col22"); set => Values["Col22"] = value; }
    public string Col23 { get => GetValue("Col23"); set => Values["Col23"] = value; }
    public string Col24 { get => GetValue("Col24"); set => Values["Col24"] = value; }
    public string Col25 { get => GetValue("Col25"); set => Values["Col25"] = value; }
    public string Col26 { get => GetValue("Col26"); set => Values["Col26"] = value; }
    public string Col27 { get => GetValue("Col27"); set => Values["Col27"] = value; }
    public string Col28 { get => GetValue("Col28"); set => Values["Col28"] = value; }
    public string Col29 { get => GetValue("Col29"); set => Values["Col29"] = value; }
    public string Col30 { get => GetValue("Col30"); set => Values["Col30"] = value; }
    public string Col31 { get => GetValue("Col31"); set => Values["Col31"] = value; }
    public string Col32 { get => GetValue("Col32"); set => Values["Col32"] = value; }
    public string Col33 { get => GetValue("Col33"); set => Values["Col33"] = value; }
    public string Col34 { get => GetValue("Col34"); set => Values["Col34"] = value; }
    public string Col35 { get => GetValue("Col35"); set => Values["Col35"] = value; }
    public string Col36 { get => GetValue("Col36"); set => Values["Col36"] = value; }
    public string Col37 { get => GetValue("Col37"); set => Values["Col37"] = value; }
    public string Col38 { get => GetValue("Col38"); set => Values["Col38"] = value; }
    public string Col39 { get => GetValue("Col39"); set => Values["Col39"] = value; }
}
