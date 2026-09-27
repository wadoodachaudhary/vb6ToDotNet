# Crystal Reports Bench

Open `/crystal-reports` in the running FlexKitTester app, or choose **Crystal Reports**
in its navigation. The HTTP launch profile uses `http://localhost:5043`.
The established default targets FlexCore; `-p:UseFlexKit=true` targets FlexKit.
The same bench works in either build. No Java, IKVM, or SAP runtime is registered.

1. Search the catalog, filter by collection/execution status, and select a report.
   The ID column and selected-report details use persistent numbers; search `#123`
   to find report 123. Filtering does not renumber reports.
   Previous/next advances through the filtered catalog. Server browse and upload
   remain available (maximum 64 MB).
2. The installed sample pack defaults to visibly labeled **Synthetic SQLite samples**.
   Other data modes remain explicit choices.
3. **Convert & Run** creates new XML and opens the positioned viewer's parameter
   dialog for non-synthetic data. SQLite mode seeds the report parameters; the
   viewer's Parameters command can change them. **Convert** creates XML without executing data.
4. The Designer tab opens the converted definition. Preview runs an isolated
   unsaved snapshot. The Viewer tab can rerun the already-converted XML.

Conversion reads the selected RPT binary. It never searches `xml_java`, sibling
reference XML, or a validation directory. All output, pictures and subreport
sidecars live in an owned temporary package. Repeated conversion gets a new
package; source files are never overwritten. Leaving the bench removes its
packages. Only the latest four packages are retained, so a session can traverse
the entire corpus without the previous 32-conversion limit. Cancellation waits for an
in-progress native decode to finish, then discards its result.

## Report Folder

`CrystalReportsBench:ReportRoot` configures the server browser's root. Without it,
the development checkout uses the sibling `JavaToCSharp/Reports` when present.
Catalog paths are relative to that corpus root, not its `rpt` subdirectory.
If no folder is available, uploading is still available. Browser selections cannot
escape the root or traverse descendant symbolic links.

## Data Sources

- **Synthetic SQLite samples** opens `Data/CrystalSamples.db` read-only (override
  with `CrystalReportsBench:SampleDatabase`). It contains a deduplicated catalog
  and typed per-report/per-subreport projection tables. No RPT-supplied SQL is
  executed in this mode. Records use deterministic sample names, amounts, dates,
  IDs, and simple selection/parameter constraints. Original selection and formula
  behavior stays enabled. A report can legitimately have no matching details.
  Binary and schema fingerprints reject unrelated data, including changed uploads
  or designer data-source edits. Uploads matching a catalog binary can reuse its data.

- **Layout only (no records)** supplies an empty result schema. It is visibly
  labeled, does not query a database, and is not a data/visual parity test.
- **Matching data fixture** uses uploaded JSON query results. No synthetic rows
  are silently substituted for missing data.
- **Read-only database** requires both `CrystalReportsBench:EnableDatabase=true`
  and the `CrystalReportsBench` connection string. Set these through user-secrets
  or environment variables; never commit credentials. Use a dedicated SQL Server
  principal with only the necessary SELECT permissions. ScriptDom rejects multiple
  statements, non-SELECT statements, SELECT INTO, sequence mutations and external
  data sources. This is an additional guard, not a substitute for database permissions.
  Queries are parameterized, time out after 30 seconds and stop at 10,000 rows.

Environment keys are `CrystalReportsBench__ReportRoot`,
`CrystalReportsBench__EnableDatabase`, and `ConnectionStrings__CrystalReportsBench`.
The bench does not issue DDL or install/upgrade a database.

## Offline Sample Pack

Report IDs are stored in the pack metadata and retained by binary SHA-256 in
`Reports/crystal-report-ids.json`. Keep this registry with the corpus when
regenerating packs: new binaries append IDs, and renaming an existing binary
does not change its ID. A pack without IDs is rejected rather than assigned
unstable display row numbers.

The 52 blocked binaries from the 2026-09-23 execution audit are collected under
`JavaToCSharp/Reports/blocked-52`. Filenames begin with their tester ID; README.md
lists each failure, and manifest.json contains original paths and verified hashes.
These are execution blockers (including synthetic-data problems), not 52 failed
native extractions. Copies are excluded from the seed corpus enumeration.

**2026-09-25:** after the blocked-52 engine fixes (FlexCore 344af31 and 0a473d8), the pack
was regenerated with `build` and `verify` against FlexCore and installed here. All 530 reports
now reach pagination: 220 Rendered, 310 Review, 0 Blocked. All 52 former blockers run on the
bench (14 Rendered, 38 Review). Review means it paginates with a caveat: text measurement
unavailable offline, chart/cross-tab placeholders, synthetic rows that miss a report's filters
(#12, #15, #33, #124), #281's uncompilable custom function, #248's clipped page-header subreport.
The previous pack is kept as `CrystalSamples-2026-09-23-backup.db` beside the regenerated one in
the regenerating session's `crystal-pack` folder.

Detailed diagnosis, prioritized triage and exact commands for another LLM are in
`JavaToCSharp/Reports/BLOCKED-REPORTS-HANDOFF.md`. Report #1 is not one of the 52
terminal failures: it previously completed with an unsupported cross-tab. Its
grid-field bindings, legacy group names and designer identity collision are now
fixed in both libraries. Its SQLite samples have been refreshed to include
shipping methods and order IDs satisfying the original selection. The remaining
style-defaults warning is intentional, not Crystal visual approval.

After a targeted importer fix, the offline tool can install one successful worker
capture into an existing pack without renumbering or replacing other reports:

```sh
dotnet run --project tools/CrystalSamples.Seed -p:UseFlexKit=true -- install-capture <copy-of-pack.db> <worker-capture.json>
```

Back up the pack, test a copy first, and rerun that report in SQLite mode before
installation. The operation is transactional and rejects blocked captures and
mismatched source paths. `CRYSTAL_TEST_SAMPLE_DB` overrides the pack used by the
`CrystalBench.Tests` harness; it is not a web-host configuration key.

To number an old pack and export an audit's blocked reports to a new folder:

```sh
dotnet run --project tools/CrystalSamples.Seed -p:UseFlexKit=true -- catalog Reports <pack.db> <audit-summary.json> <new-blocked-folder>
```

This offline command updates only catalog metadata in the database; it never
changes sample rows or source RPTs. Existing destination folders are rejected.

From `JavaToCSharp`, generate a new pack (existing output files are not overwritten):

```sh
dotnet run --project tools/CrystalSamples.Seed -p:UseFlexKit=true -- build Reports /tmp/CrystalSamples-new.db /tmp/crystal-sample-audit
dotnet run --project tools/CrystalSamples.Seed -p:UseFlexKit=true -- verify Reports /tmp/CrystalSamples-new.db /tmp/crystal-sqlite-audit
```

Install the generated database at `FlexKitTester/Data/CrystalSamples.db`, or configure
the sample database path. Schema creation is confined to this offline tool, never
web-app startup. The tool reads only `Reports/rpt` and `Reports/downloaded-samples`,
deduplicates SHA-256 binaries, converts every RPT afresh, and executes bounded
synthetic records through C# pagination in isolated 25-second workers. `verify`
actually reads the generated SQLite pack through the host adapter. It never uses
`xml_java` or a cached XML reference. Sample generation is heuristic, not a copy
of the original database or an implementation of its SQL joins/stored procedures.

The Compatibility tab lists import, formula, runtime, data-coverage, and blocked
findings. Catalog statuses are the offline audit results, not user visual approval.
Legacy OLE/OS-dependent graphics may be replaced after review; core data, charts,
cross-tabs, summaries, and formulas are not silently dropped or marked obsolete.

## Fixture Contract

The Conversion tab exposes the generated query, its SHA-256 fingerprint, parameter
definitions and projected aliases for the main report and inline subreports.
Supply **query-result rows**, not unjoined source-table rows. Column names must
match the projected SQL aliases. Maximum upload: 16 MB, 256 result sets,
2,048 columns per set and 10,000 rows per set.

```json
{
  "reports": [
    {
      "reportId": "Example",
      "sqlSha256": "<64-character SHA-256 of the exact generated SQL>",
      "parameters": { "JobID": 123 },
      "columns": [
        { "name": "Items_Amount", "type": "decimal" }
      ],
      "rows": [[125.50], [200.00]]
    }
  ]
}
```

Types: `string`, `integer`, `decimal`, `number`, `boolean`, `datetime`. JSON null
maps to database null. Datetime columns use ISO timestamps. Date parameters can
use ISO dates, which also match the parameter dialog's typed DateTime values.
Every request must match exactly one report ID + query hash + complete parameter
set, including fixed parameters and linked subreport values. Provide separate
result sets for different parameter combinations; omit duplicate identical sets.
Missing columns, missing cases, query changes and ambiguous cases are errors.
Multi-value parameters require the hash/aliases of the rewritten query.

The designer preserves the main report identity in preview snapshots. Changing
its query intentionally invalidates an old fixture rather than reusing unrelated
rows. Designer edits are not saved over generated conversion XML by this bench.

## Verification And Remaining Limits

Console/component harness (does not start a website):

```sh
dotnet run --project /Users/wadood/projects/JavaToCSharp/tools/CrystalBench.Tests -p:UseFlexKit=true
dotnet run --project /Users/wadood/projects/JavaToCSharp/tools/CrystalBench.Tests -p:UseFlexKit=false
node /Users/wadood/projects/JavaToCSharp/tools/CrystalBench.Tests/verify-browser.cjs
```

Run these from `JavaToCSharp` so the corpus paths resolve. Runtime limitations are
visible in the viewer. Fixed-height mutable page formulas are supported; growing
state-dependent bands and shared-state inline subreport scheduling still require
more work. Original Crystal PDFs/images with identical data, parameters, fonts
and page/printer settings are required for visual approval. Reference XML and
native screenshots do not constitute that approval.
