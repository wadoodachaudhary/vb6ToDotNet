# Crystal Reports conversion QA

Cloud batch of the FlexKitTester sample pack on 2026-09-26 (UTC), follow-up to the 2026-09-25 run.
Engine: native C# RPT to XML in FlexCore (`CrystalRptToXml`), then synthetic SQLite rows from `FlexKitTester/Data/CrystalSamples.db`, then C# pagination.
Sibling revisions used for this run: FlexCore `2587807` plus `flexcore-page-furniture.patch` (`0f3867b`), JavaToCSharp `a34ef9d`.
The patch could not be pushed to `wadoodachaudhary/FlexCore` from this run. Apply it on top of `2587807` before a re-run.
Source `.rpt` files were left unchanged. No SAP Crystal runtime was used.

## Counts

| Metric | 2026-09-25 | 2026-09-26 |
| --- | ---: | ---: |
| Reports in pack (attempted) | 530 | 530 |
| Conversion OK | 530 | 530 |
| Rendered clean (`Rendered`) | 249 | 296 |
| Rendered with diagnostics (`Review`) | 123 | 234 |
| Render OK (Rendered + Review) | 372 | 530 |
| Failed / blocked | 158 | 0 |
| Skipped | 0 | 0 |
| PNG screenshots | 530 | 530 |

Of the 158 reports that were blocked on 2026-09-25: **47 Rendered, 111 Review, 0 still blocked**. None of the 372 that already rendered changed status, page count, or diagnostics.

Primary corpus (`Reports/rpt`, 19 reports): 11 Review, 8 Rendered, 0 Blocked.
Downloaded samples: 288 Rendered, 223 Review, 0 Blocked.

Elapsed: 4 minutes 52 seconds, including one Chrome session that captured all 530 PNGs.

## Artifacts

- `screenshots/` — one PNG per report, `{id:0000}-{safeName}.png`. Each image is the QA page: status banner plus the positioned pages when pagination succeeded. PNGs for the 158 previously blocked reports were replaced. The other 372 PNGs are unchanged.
- `results.jsonl` — one JSON object per report: `Id`, `Name`, `Status`, `ConversionOk`, `RenderOk`, `Pages`, `Error`, `ScreenshotPath` (repo-relative), plus hash, relative RPT path, and up to 20 diagnostics.
- `gallery.html` — thumbnail index of every PNG. Open it in a browser from this folder.

Per-report HTML and worker JSON were produced by the harness and then left out of git. Each HTML file inlines the FlexCore stylesheet, so the full set is about 400 MB. Regenerate them with the command below if you need the HTML.

## How to open the gallery

From a checkout of this branch:

```sh
open FlexKitTester/crystal-qa-cloud/gallery.html
# or browse the PNGs directly
open FlexKitTester/crystal-qa-cloud/screenshots
```

## Sample schema changes

`FlexKitTester/Data/CrystalSamples.db` was refreshed for the 158 previously blocked binaries. Each report was converted again with the current FlexCore importer and reseeded by `CrystalSamples.Seed` `worker`, then installed with `install-capture`. Other catalog reports were not rewritten.

| Change | Reports |
| --- | ---: |
| Schema fingerprint replaced (stored dataset key did not match this conversion) | 153 |
| Column names or types also changed | 84 |
| Column signature unchanged (fingerprint only) | 74 |
| Gained a dataset the old pack did not have (usually a subreport) | 17 |
| Parameter values refreshed | 24 |

Column corrections are the current converter's field types. The old pack often stored dates as strings and amounts as datetimes, and it omitted columns the conversion now projects (for example `Orders_Ship Via`, `sales_order_sales_rep_id`). Rows are still the seed tool's deterministic synthetic projections (24 rows when the report has fields, 1 row for a static report).

Parameter fixes that unblocked runtime failures:

- #225 and #257 `EndDate`: `2005-01-30 0` → `2005-01-31 00:00:00`
- #521 `EndDate`: `2005-03-30 55920` → `2005-03-31 15:32:00` (55920 seconds after midnight is 15:32:00; the old pack stored Crystal's date and time parts as raw numbers)
- Linked subreport parameters that were the string `Sample` are now `1` when the linked column is numeric

#152 (76 Median of an Array) paginates with the refreshed rows. The previous pack's rows tripped the array subscript check.

## Runtime mitigations

Two FlexCore layout changes in `flexcore-page-furniture.patch` (`0f3867b`, parent `2587807`). Designed page headers and footers that already fill the page still raise Crystal's "page area too large" error.

- Page-header growth is clipped to the designed section height when an inline subreport would otherwise consume the page. #248 `crptDetails.rpt` and #249 `crptSubReport.rpt` are Review with that diagnostic. The subreport in the page header is clipped; the body still prints.
- Physical page replay no longer aborts when a while-printing formula throws. The fault is recorded as a field diagnostic, the same way `Format()` already did. #486 `TOCv8.rpt` is Review: `{@Index Display While Do}` indexes past the index array because the While loop stops on string length 250 and the short synthetic index never reaches that length. That formula behavior is inherent to this sample size.

## Review themes (234)

These reports paginated. The screenshot shows the pages. The status is `Review` because the run recorded at least one diagnostic. Common themes:

- CanGrow used approximate font metrics (batch pagination did not call the browser measurer)
- Cross-tab and chart objects rendered as unsupported-object placeholders
- Field or formula reference could not be resolved
- Subreport on-demand or definition unavailable
- Page header or footer content clipped to the designed section height (#248, #249)
- `{@Index Display While Do}` array subscript on #486

The CanGrow metric warning comes from `ReportLayoutSession.Paginate()` in the batch worker. The separate `--capture-native` path measures text in Chrome; this batch does not.

## How to re-run

Layout expected by the projects (same as a local Mac checkout):

```text
VBToCSharp/HomeFront/FlexKitTester/     this repo
VBToCSharp/FlexCore/                    github.com/wadoodachaudhary/FlexCore
JavaToCSharp/Reports/                   rpt + downloaded-samples
JavaToCSharp/tools/CrystalBench.Tests/
JavaToCSharp/tools/CrystalSamples.Seed/
```

`FlexKitTester.csproj` references `../../FlexCore`. `CrystalBench.Tests` and `CrystalSamples.Seed` reference `../../../VBToCSharp/HomeFront/FlexKitTester`. Run the harness from `JavaToCSharp` so `Reports/` resolves. Apply `FlexKitTester/crystal-qa-cloud/flexcore-page-furniture.patch` on FlexCore `2587807` so #248, #249, and #486 stay unblocked:

```sh
git -C /path/to/FlexCore apply /path/to/FlexKitTester/crystal-qa-cloud/flexcore-page-furniture.patch
```

```sh
dotnet build VBToCSharp/HomeFront/FlexKitTester/FlexKitTester.csproj
cd JavaToCSharp
# screenshot-html.cjs loads Playwright from PLAYWRIGHT_MODULE (Chrome channel)
export PLAYWRIGHT_MODULE=/path/to/node_modules/playwright
dotnet run --project tools/CrystalBench.Tests -- --batch-qa ../VBToCSharp/HomeFront/FlexKitTester/crystal-qa-cloud
```

Optional filters: `--limit N`, `--from-id N`, `--to-id N`. The pack path defaults to `../VBToCSharp/HomeFront/FlexKitTester/Data/CrystalSamples.db` (`CRYSTAL_TEST_SAMPLE_DB` overrides it). Each report runs in its own process with a 40-second limit.

To refresh one report's synthetic rows after a converter change (does not renumber the catalog):

```sh
dotnet run --project tools/CrystalSamples.Seed -- worker Reports <relative.rpt> <sha256> /tmp/capture.json
dotnet run --project tools/CrystalSamples.Seed -- install-capture <CrystalSamples.db> /tmp/capture.json
```

`install-capture` accepts a Review or Rendered capture only. Back up the pack first.

## What remains

The full pack of 530 reports was converted, paginated, and screenshotted. Nothing was skipped and nothing is blocked.
Screenshots show native conversion and synthetic-sample pagination. Matching original Crystal output still needs the same data, parameters, fonts, and page setup.
#486 still reports an array-subscript diagnostic inside `{@Index Display While Do}` with this short synthetic index. Charts and cross-tabs remain placeholders.
