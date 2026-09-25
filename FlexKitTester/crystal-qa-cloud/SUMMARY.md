# Crystal Reports conversion QA

Cloud batch of the FlexKitTester sample pack on 2026-09-25 (UTC).
Engine: native C# RPT to XML in FlexCore (`CrystalRptToXml`), then synthetic SQLite rows from `FlexKitTester/Data/CrystalSamples.db`, then C# pagination.
Sibling revisions used for this run: FlexCore `2587807`, JavaToCSharp `a34ef9d`.
Source `.rpt` files were left unchanged. No SAP Crystal runtime was used.

## Counts

| Metric | Count |
| --- | ---: |
| Reports in pack (attempted) | 530 |
| Conversion OK | 530 |
| Rendered clean (`Rendered`) | 249 |
| Rendered with diagnostics (`Review`) | 123 |
| Render OK (Rendered + Review) | 372 |
| Failed / blocked | 158 |
| Skipped | 0 |
| PNG screenshots | 530 |

Primary corpus (`Reports/rpt`, 19 reports): 9 Review, 6 Rendered, 4 Blocked.
Downloaded samples: 243 Rendered, 154 Blocked, 114 Review.

Elapsed: 4 minutes 29 seconds, including one Chrome session that captured all 530 PNGs.

## Artifacts

- `screenshots/` — one PNG per report, `{id:0000}-{safeName}.png`. Each image is the QA page: status banner plus the positioned pages when pagination succeeded.
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

## Top failure themes

Every report converted. Blocked rows failed during sample execution or pagination, and their screenshots are the status card.

### Blocked (158)

- 153 — SQLite sample schema fingerprint does not match this conversion
- 5 — Other runtime blocker

The schema mismatches mean the current native conversion fingerprint does not match the dataset stored in `CrystalSamples.db`. The bench refuses to substitute unrelated rows. Those 153 reports still converted.

Other runtime blockers:

- #152 76 Median of an Array.rpt: InvalidDataException: A subscript must be between 1 and the size of the array.
- #225 Combination Balance Sheet and Income Statement.RPT: FormatException: String '2005-01-30 0' was not recognized as a valid DateTime.
- #249 crptSubReport.rpt: InvalidDataException: Page headers plus footers leave no space for content (Crystal: page area too large).
- #486 TOCv8.rpt: InvalidDataException: A subscript must be between 1 and the size of the array.
- #521 Variance Analysis Report.rpt: InvalidDataException: Crystal function month: String '2005-03-30 55920' was not recognized as a valid DateTime.

### Review (123)

These reports paginated. The screenshot shows the pages. The status is `Review` because the run recorded at least one diagnostic. Primary theme (first diagnostic):

- 68 — CanGrow used approximate font metrics (batch pagination did not call the browser measurer)
- 26 — Cross-tab objects rendered as unsupported-object placeholders
- 10 — Field or formula reference could not be resolved
- 7 — Chart objects rendered as unsupported-object placeholders
- 4 — Unimplemented Crystal function or property
- 4 — Subreport on-demand or definition unavailable
- 2 — Other diagnostic: {@Title}: Document property 'filename' is not available to this report run.
- 1 — Other diagnostic: SectionAreaConditionFormulas/BackgroundColor: the XML contains only a line comment; reconvert the RPT to retai
- 1 — Other diagnostic: SectionAreaConditionFormulas/EnableSuppress: the XML contains only a line comment; reconvert the RPT to retain

The CanGrow metric warning comes from `ReportLayoutSession.Paginate()` in the batch worker. The separate `--capture-native` path measures text in Chrome; this batch does not, so many otherwise successful layouts are marked Review for approximate CanGrow pagination.

Charts and cross-tabs that the engine does not draw yet appear as labeled placeholders in the PNG, with the rest of the page still laid out.

## How to re-run

Layout expected by the projects (same as a local Mac checkout):

```text
VBToCSharp/HomeFront/FlexKitTester/     this repo
VBToCSharp/FlexCore/                    github.com/wadoodachaudhary/FlexCore
JavaToCSharp/Reports/                   rpt + downloaded-samples
JavaToCSharp/tools/CrystalBench.Tests/
```

`FlexKitTester.csproj` references `../../FlexCore`. `CrystalBench.Tests` references `../../../VBToCSharp/HomeFront/FlexKitTester`. Run the harness from `JavaToCSharp` so `Reports/` resolves.

```sh
dotnet build VBToCSharp/HomeFront/FlexKitTester/FlexKitTester.csproj
cd JavaToCSharp
# screenshot-html.cjs loads Playwright from PLAYWRIGHT_MODULE (Chrome channel)
export PLAYWRIGHT_MODULE=/path/to/node_modules/playwright
dotnet run --project tools/CrystalBench.Tests -- --batch-qa ../VBToCSharp/HomeFront/FlexKitTester/crystal-qa-cloud
```

Optional filters: `--limit N`, `--from-id N`, `--to-id N`. The pack path defaults to `../VBToCSharp/HomeFront/FlexKitTester/Data/CrystalSamples.db` (`CRYSTAL_TEST_SAMPLE_DB` overrides it). Each report runs in its own process with a 40-second limit.

## What remains

The full pack of 530 reports was converted, classified, and screenshotted. Nothing was skipped.
Screenshots show native conversion and synthetic-sample pagination. Matching original Crystal output still needs the same data, parameters, fonts, and page setup.
