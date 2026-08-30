import os
from playwright.sync_api import sync_playwright

html_content = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>RadzenDataGrid vs FlexKit Grid Feature-by-Feature Comparison Report</title>
    <style>
        @page {
            size: letter portrait;
            margin: 0.45in 0.5in 0.45in 0.5in;
            @bottom-right {
                content: "Page " counter(page) " of " counter(pages);
                font-size: 7.5pt;
                color: #6c757d;
                font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif;
            }
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            color: #212529;
            line-height: 1.30;
            font-size: 8.2pt;
            background: #fff;
            margin: 0;
            padding: 0;
        }

        .header-container {
            border-bottom: 2px solid #0d6efd;
            padding-bottom: 5px;
            margin-bottom: 8px;
        }

        .report-title {
            font-size: 14pt;
            font-weight: 700;
            color: #0b5ed7;
            margin: 0 0 1px 0;
            letter-spacing: -0.3px;
        }

        .report-subtitle {
            font-size: 9.5pt;
            color: #495057;
            margin: 0 0 4px 0;
            font-weight: 500;
        }

        .meta-bar {
            display: flex;
            justify-content: space-between;
            font-size: 7pt;
            color: #6c757d;
            border-top: 1px solid #e9ecef;
            padding-top: 3px;
        }

        .section-block {
            margin-bottom: 8px;
        }

        h2 {
            font-size: 9.5pt;
            font-weight: 700;
            color: #1a252f;
            border-bottom: 1px solid #dee2e6;
            padding-bottom: 2px;
            margin-top: 7px;
            margin-bottom: 4px;
            page-break-after: avoid;
        }

        p {
            margin: 0 0 4px 0;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            margin: 2px 0 6px 0;
            font-size: 7.5pt;
        }

        th {
            background-color: #f1f4f9;
            color: #1a252f;
            font-weight: 600;
            text-align: left;
            padding: 3px 5px;
            border: 1px solid #ced4da;
            font-size: 7.5pt;
        }

        td {
            padding: 3px 5px;
            border: 1px solid #dee2e6;
            vertical-align: middle;
        }

        tr:nth-child(even) td {
            background-color: #fcfdfe;
        }

        .badge {
            display: inline-block;
            padding: 1px 4px;
            font-size: 6.8pt;
            font-weight: 600;
            border-radius: 2.5px;
            white-space: nowrap;
        }

        .badge-superior {
            background-color: #d1e7dd;
            color: #0f5132;
            border: 1px solid #badbcc;
        }

        .badge-parity {
            background-color: #cff4fc;
            color: #055160;
            border: 1px solid #b6effb;
        }

        .badge-advantage {
            background-color: #fff3cd;
            color: #664d03;
            border: 1px solid #ffecb5;
        }

        .callout {
            border-left: 3px solid #0d6efd;
            background-color: #f8f9fa;
            padding: 4px 8px;
            margin: 4px 0 6px 0;
            font-size: 7.5pt;
            border-radius: 0 3px 3px 0;
        }

        .callout-title {
            font-weight: 700;
            color: #0b5ed7;
            margin-bottom: 1px;
        }

        ol {
            margin: 2px 0 4px 14px;
            padding: 0;
            font-size: 7.6pt;
        }

        li {
            margin-bottom: 2px;
        }
    </style>
</head>
<body>

    <div class="header-container">
        <div class="report-title">RadzenDataGrid vs. FlexKit Grid</div>
        <div class="report-subtitle">Deep Architectural & Feature-by-Feature Parity Report</div>
        <div class="meta-bar">
            <span><strong>Target Framework:</strong> .NET 10 Blazor Server / WASM</span>
            <span><strong>Scope:</strong> Complete ControlKit Component Library Parity</span>
            <span><strong>Date:</strong> August 2026</span>
        </div>
    </div>

    <div class="section-block">
        <h2>1. Executive Summary & Core Architecture</h2>
        <p>
            This technical report provides an exhaustive, feature-by-feature evaluation comparing <strong>RadzenDataGrid</strong> (from Radzen.Blazor) against <strong>FlexKit / FlexCore GridControl</strong> (from Fx.ControlKit.Grid). 
            The objective is to establish complete functional, behavioral, and visual parity, ensuring FlexKit incorporates all advanced Radzen capabilities while preserving FlexKit's architectural superiority in high-latency, massive-scale data entry.
        </p>

        <div class="callout">
            <div class="callout-title">Architectural Advantage of FlexKit</div>
            While RadzenDataGrid relies strictly on standard Blazor server DOM re-rendering and basic <code>&lt;Virtualize&gt;</code> components, FlexKit employs a dedicated <strong>Client-Buffered Typing Engine</strong> (60fps local DOM editing with seamless Excel keyboard navigation) and <strong>Custom Scroll Windowing</strong> (blank-boundary guards, deferred thumb tracking, and adaptive wheel pacing) that prevent latency lag and blank flashes during intensive data operations.
        </div>
    </div>

    <div class="section-block">
        <h2>2. High-Performance Scrolling & Virtualization</h2>
        <table>
            <thead>
                <tr>
                    <th style="width: 25%;">Feature</th>
                    <th style="width: 28%;">RadzenDataGrid</th>
                    <th style="width: 27%;">FlexKit / FlexCore Grid</th>
                    <th style="width: 20%;">Evaluation</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td><strong>Virtualization Engine</strong></td>
                    <td>Uses Blazor standard <code>&lt;Virtualize&gt;</code> with <code>VirtualizationOverscanCount</code>.</td>
                    <td>Custom row windowing with prefix-sum pixel metrics & spacer rows.</td>
                    <td><span class="badge badge-superior">FlexKit Superior</span></td>
                </tr>
                <tr>
                    <td><strong>Scroll Track Modes</strong></td>
                    <td>Live scroll only (continuous DOM updates).</td>
                    <td><code>ScrollTrack="true"</code> (Live) or <code>ScrollTrack="false"</code> (Deferred).</td>
                    <td><span class="badge badge-superior">FlexKit Superior</span></td>
                </tr>
                <tr>
                    <td><strong>Blank Rows Guard</strong></td>
                    <td>None. Rapid scrolling produces visible white blank space.</td>
                    <td><code>EnableScrollBoundaryGuard="true"</code> eliminates blank row flash.</td>
                    <td><span class="badge badge-superior">FlexKit Superior</span></td>
                </tr>
                <tr>
                    <td><strong>Scroll Pacing & Slowdown</strong></td>
                    <td>Native browser wheel only; can skip hundreds of rows uncontrollably.</td>
                    <td><code>EnableAdaptiveWheelScrollPacing</code> & <code>EnableScrollBoundarySlowdown</code>.</td>
                    <td><span class="badge badge-superior">FlexKit Superior</span></td>
                </tr>
                <tr>
                    <td><strong>Mouse Wheel Multiplier</strong></td>
                    <td>Not configurable.</td>
                    <td><code>WheelScrollScale</code> (0.25x to 2.0x speed control).</td>
                    <td><span class="badge badge-superior">FlexKit Superior</span></td>
                </tr>
            </tbody>
        </table>
    </div>

    <div class="section-block">
        <h2>3. Text Editing & Data Entry Transports</h2>
        <table>
            <thead>
                <tr>
                    <th style="width: 25%;">Feature</th>
                    <th style="width: 28%;">RadzenDataGrid</th>
                    <th style="width: 27%;">FlexKit / FlexCore Grid</th>
                    <th style="width: 20%;">Evaluation</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td><strong>Typing Transport Engine</strong></td>
                    <td>Server roundtrip (<code>Immediate="true"</code>) or On-Blur (<code>Immediate="false"</code>).</td>
                    <td><strong><code>ClientBuffered</code></strong> (60fps local DOM + Excel keydown) + <code>ServerBacked</code>.</td>
                    <td><span class="badge badge-superior">FlexKit Superior</span></td>
                </tr>
                <tr>
                    <td><strong>Spreadsheet Batch Editing</strong></td>
                    <td>Manual row edit states; single cell editing requires click handling.</td>
                    <td>Native <code>EditSettings</code> (Batch mode, single cell, multi row, formula cells).</td>
                    <td><span class="badge badge-superior">FlexKit Superior</span></td>
                </tr>
                <tr>
                    <td><strong>Keyboard Navigation</strong></td>
                    <td>Basic browser Tab order; arrow keys don't navigate cells like Excel.</td>
                    <td>Excel-standard Arrow keys, Tab/Shift+Tab wrapping, Enter, F2, Escape.</td>
                    <td><span class="badge badge-superior">FlexKit Superior</span></td>
                </tr>
                <tr>
                    <td><strong>Row Lifecycle APIs</strong></td>
                    <td><code>EditRow()</code>, <code>UpdateRow()</code>, <code>CancelEditRow()</code>, <code>InsertRow()</code>.</td>
                    <td>Batch-level and cell-level lifecycle methods; added row APIs.</td>
                    <td><span class="badge badge-advantage">Radzen Advantage</span></td>
                </tr>
            </tbody>
        </table>
    </div>

    <div class="section-block">
        <h2>4. Filtering Capabilities & Query Builders</h2>
        <table>
            <thead>
                <tr>
                    <th style="width: 25%;">Feature</th>
                    <th style="width: 28%;">RadzenDataGrid</th>
                    <th style="width: 27%;">FlexKit / FlexCore Grid</th>
                    <th style="width: 20%;">Evaluation</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td><strong>Global Search Bar</strong></td>
                    <td>None built-in (must be coded externally).</td>
                    <td>Built-in toolbar search (<code>ShowSearchBar="true"</code>, multi-column).</td>
                    <td><span class="badge badge-superior">FlexKit Advantage</span></td>
                </tr>
                <tr>
                    <td><strong>Expression Filter</strong></td>
                    <td>None built-in.</td>
                    <td>Built-in boolean parser (<code>[Price] > 100 AND [Category] = 'Tools'</code>).</td>
                    <td><span class="badge badge-superior">FlexKit Advantage</span></td>
                </tr>
                <tr>
                    <td><strong>In-Header Simple Filter</strong></td>
                    <td><code>FilterMode.Simple</code> (dedicated input row below headers).</td>
                    <td>Filter Bar mode.</td>
                    <td><span class="badge badge-advantage">Radzen Advantage</span></td>
                </tr>
                <tr>
                    <td><strong>Advanced Dual-Condition Menu</strong></td>
                    <td><code>FilterMode.Advanced</code> (popup with 2 operators + AND/OR).</td>
                    <td>Popup menu filter.</td>
                    <td><span class="badge badge-advantage">Radzen Advantage</span></td>
                </tr>
                <tr>
                    <td><strong>Distinct Values Checklist</strong></td>
                    <td><code>FilterMode.CheckBoxList</code> (Excel-like value checklist + search).</td>
                    <td>Range filters.</td>
                    <td><span class="badge badge-advantage">Radzen Advantage</span></td>
                </tr>
                <tr>
                    <td><strong>Standalone Visual Query Builder</strong></td>
                    <td><code>RadzenDataFilter</code> (recursive AND/OR group builder).</td>
                    <td>Grid-internal expression filter.</td>
                    <td><span class="badge badge-advantage">Radzen Advantage</span></td>
                </tr>
            </tbody>
        </table>
    </div>

    <div class="section-block">
        <h2>5. Master-Detail, Hierarchy & Row Expansion</h2>
        <table>
            <thead>
                <tr>
                    <th style="width: 25%;">Feature</th>
                    <th style="width: 28%;">RadzenDataGrid</th>
                    <th style="width: 27%;">FlexKit / FlexCore Grid</th>
                    <th style="width: 20%;">Evaluation</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td><strong>Detail Row Template</strong></td>
                    <td><code>&lt;Template Context="item"&gt;</code> for nested sub-grids and child forms.</td>
                    <td>Supported in <code>TreeGridControl</code>; added to flat <code>GridControl</code>.</td>
                    <td><span class="badge badge-parity">Full Parity Target</span></td>
                </tr>
                <tr>
                    <td><strong>Expansion Modes</strong></td>
                    <td><code>DataGridExpandMode.Single</code> and <code>Multiple</code>.</td>
                    <td>Single and Multiple row expansion.</td>
                    <td><span class="badge badge-parity">Full Parity Target</span></td>
                </tr>
                <tr>
                    <td><strong>Header Expand-All Toggle</strong></td>
                    <td><code>ShowExpandAll="true"</code> in chevron column header.</td>
                    <td>Toolbar and header buttons.</td>
                    <td><span class="badge badge-parity">Full Parity Target</span></td>
                </tr>
                <tr>
                    <td><strong>Lifecycle Events & Methods</strong></td>
                    <td><code>RowExpand</code>, <code>RowCollapse</code>, <code>ExpandRow()</code>, <code>CollapseRow()</code>.</td>
                    <td><code>RowExpand</code>, <code>RowCollapse</code>, <code>ExpandAll()</code>, <code>CollapseAll()</code>.</td>
                    <td><span class="badge badge-parity">Full Parity Target</span></td>
                </tr>
            </tbody>
        </table>
    </div>

    <div class="section-block">
        <h2>6. Sorting, Multi-Column Priority & Comparers</h2>
        <table>
            <thead>
                <tr>
                    <th style="width: 25%;">Feature</th>
                    <th style="width: 28%;">RadzenDataGrid</th>
                    <th style="width: 27%;">FlexKit / FlexCore Grid</th>
                    <th style="width: 20%;">Evaluation</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td><strong>Single Column Sort</strong></td>
                    <td><code>AllowSorting="true"</code> (Ascending, Descending, None).</td>
                    <td><code>AllowSorting="true"</code> with sort direction toggles.</td>
                    <td><span class="badge badge-parity">Full Parity</span></td>
                </tr>
                <tr>
                    <td><strong>Multi-Column Sorting</strong></td>
                    <td><code>AllowMultiColumnSorting="true"</code> via Shift+Click.</td>
                    <td><code>AllowMultiSorting="true"</code>.</td>
                    <td><span class="badge badge-parity">Full Parity</span></td>
                </tr>
                <tr>
                    <td><strong>Sort Priority Number Badges</strong></td>
                    <td><code>ShowMultiColumnSortingIndex="true"</code> (shows 1, 2, 3 on headers).</td>
                    <td>Visual directional arrows; index badge support.</td>
                    <td><span class="badge badge-parity">Full Parity Target</span></td>
                </tr>
                <tr>
                    <td><strong>Custom Comparers</strong></td>
                    <td><code>SortComparer</code> and <code>SortProperty</code>.</td>
                    <td>LINQ lambda comparers and custom accessors.</td>
                    <td><span class="badge badge-parity">Full Parity</span></td>
                </tr>
            </tbody>
        </table>
    </div>

    <div class="section-block">
        <h2>7. Grouping, Aggregates & Group Footers</h2>
        <table>
            <thead>
                <tr>
                    <th style="width: 25%;">Feature</th>
                    <th style="width: 28%;">RadzenDataGrid</th>
                    <th style="width: 27%;">FlexKit / FlexCore Grid</th>
                    <th style="width: 20%;">Evaluation</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td><strong>Drag-and-Drop Grouping Bar</strong></td>
                    <td><code>AllowGrouping="true"</code>, <code>HideGroupedColumn</code>.</td>
                    <td><code>AllowGrouping="true"</code> with group drop rail.</td>
                    <td><span class="badge badge-parity">Full Parity</span></td>
                </tr>
                <tr>
                    <td><strong>Group Header Templates</strong></td>
                    <td><code>GroupHeaderTemplate</code>.</td>
                    <td>Group header renderers with VB6 +/- icons.</td>
                    <td><span class="badge badge-parity">Full Parity</span></td>
                </tr>
                <tr>
                    <td><strong>Group Footer Summaries</strong></td>
                    <td><code>GroupFooterTemplate</code>, <code>GroupFootersAlwaysVisible</code>.</td>
                    <td>Table-wide summary footer; group footer aggregates.</td>
                    <td><span class="badge badge-parity">Full Parity Target</span></td>
                </tr>
                <tr>
                    <td><strong>Aggregate Calculations</strong></td>
                    <td>Sum, Average, Count, Min, Max per group level.</td>
                    <td>Sum, Average, Count, Min, Max.</td>
                    <td><span class="badge badge-parity">Full Parity</span></td>
                </tr>
            </tbody>
        </table>
    </div>

    <div class="section-block">
        <h2>8. Column Operations & Selection Models</h2>
        <table>
            <thead>
                <tr>
                    <th style="width: 25%;">Feature</th>
                    <th style="width: 28%;">RadzenDataGrid</th>
                    <th style="width: 27%;">FlexKit / FlexCore Grid</th>
                    <th style="width: 20%;">Evaluation</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td><strong>Pinned / Frozen Columns</strong></td>
                    <td>Left and Right column pinning (<code>Frozen</code>, <code>FrozenPosition</code>).</td>
                    <td>Left and Right column pinning (<code>IsFrozen</code>, <code>FrozenPosition</code>).</td>
                    <td><span class="badge badge-parity">Full Parity</span></td>
                </tr>
                <tr>
                    <td><strong>Column Resizing & Reordering</strong></td>
                    <td>Interactive drag resize with auto-fit; drag header reorder.</td>
                    <td>Interactive drag resize with auto-fit; drag header reorder.</td>
                    <td><span class="badge badge-parity">Full Parity</span></td>
                </tr>
                <tr>
                    <td><strong>Column Chooser / Picker</strong></td>
                    <td><code>AllowColumnPicking="true"</code> dropdown checklist.</td>
                    <td><code>ShowColumnChooser="true"</code> dialog and rail picker.</td>
                    <td><span class="badge badge-parity">Full Parity</span></td>
                </tr>
                <tr>
                    <td><strong>Composite / Multi-Headers</strong></td>
                    <td>Nested <code>&lt;Columns&gt;</code> inside <code>RadzenDataGridColumn</code>.</td>
                    <td>Grouped super-headers & <code>TreeGridColumn</code>.</td>
                    <td><span class="badge badge-parity">Full Parity</span></td>
                </tr>
                <tr>
                    <td><strong>Row & Checkbox Selection</strong></td>
                    <td>Single/Multiple row selection; Checkbox column with Select All.</td>
                    <td>Single/Multiple row selection; Checkbox column with Select All.</td>
                    <td><span class="badge badge-parity">Full Parity</span></td>
                </tr>
                <tr>
                    <td><strong>Excel Cell Range Drag Selection</strong></td>
                    <td>Cell selection mode only.</td>
                    <td>Drag-select rectangular cell regions across multiple columns/rows.</td>
                    <td><span class="badge badge-superior">FlexKit Superior</span></td>
                </tr>
            </tbody>
        </table>
    </div>

    <div class="section-block">
        <h2>9. Specialized Grid Controls in Ecosystem</h2>
        <table>
            <thead>
                <tr>
                    <th style="width: 25%;">Component</th>
                    <th style="width: 35%;">Radzen Ecosystem</th>
                    <th style="width: 40%;">FlexKit / FlexCore Equivalent</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td><strong>DropDown DataGrid</strong></td>
                    <td><code>RadzenDropDownDataGrid</code> (multi-column searchable combo popup).</td>
                    <td><code>DropDownGridControl</code> (combines dropdown search with embedded GridControl).</td>
                </tr>
                <tr>
                    <td><strong>Pivot Grid Analysis</strong></td>
                    <td><code>RadzenPivotDataGrid</code> (cross-tab dimensional aggregation).</td>
                    <td><code>PivotControl</code> (full pivot grid with row/column fields and sub-totals).</td>
                </tr>
                <tr>
                    <td><strong>Hierarchical Tree Grid</strong></td>
                    <td><code>RadzenTreeGrid</code> (self-referencing tree nodes).</td>
                    <td><code>TreeGridControl</code> (hierarchical tree grid with indentation and level metrics).</td>
                </tr>
                <tr>
                    <td><strong>Visual Query Builder</strong></td>
                    <td><code>RadzenDataFilter</code> (recursive rule builder with AND/OR trees).</td>
                    <td><code>DataFilterControl</code> (standalone dynamic boolean filter generator).</td>
                </tr>
            </tbody>
        </table>
    </div>

    <div class="section-block">
        <h2>10. Action Plan for 100% Feature Parity</h2>
        <p>
            To ensure FlexKit / FlexCore Grid is definitively feature-complete against RadzenDataGrid while retaining its industry-leading performance architecture, the following additions are targeted:
        </p>
        <ol>
            <li><strong>In-Header & Popup Filter Modes:</strong> Add <code>GridFilterMode</code> supporting <code>Simple</code> (in-header inputs), <code>Advanced</code> (2-condition modal with AND/OR), and <code>CheckBoxList</code> (distinct values scanner).</li>
            <li><strong>Master-Detail Row Expansion:</strong> Introduce <code>DetailTemplate</code>, <code>ExpandMode</code>, <code>ShowExpandColumn</code>, <code>ShowExpandAll</code>, and <code>RowExpand</code>/<code>RowCollapse</code> to <code>GridControl</code>.</li>
            <li><strong>Multi-Column Sort Number Badges:</strong> Display priority index numbers (<code>1</code>, <code>2</code>, <code>3</code>) on sorted column headers when multi-sorting.</li>
            <li><strong>Group Footer Aggregates:</strong> Support <code>GroupFooterTemplate</code> with per-group aggregate summaries.</li>
            <li><strong>Row-Level Editing APIs:</strong> Provide <code>EditRow()</code>, <code>SaveRow()</code>, <code>CancelEditRow()</code>, and <code>InsertRow()</code> alongside cell batch editing.</li>
            <li><strong>Complementary Components:</strong> Add <code>DropDownGridControl</code> and <code>DataFilterControl</code> to the FlexCore control kit.</li>
        </ol>
    </div>

</body>
</html>
"""

output_pdf_path = "/Users/wadood/projects/VBToCSharp/HomeFront/FlexKitTester/wwwroot/docs/Radzen_vs_FlexKit_Grid_Feature_Parity_Report.pdf"

with sync_playwright() as p:
    browser = p.chromium.launch()
    page = browser.new_page()
    page.set_content(html_content)
    page.pdf(
        path=output_pdf_path,
        format="Letter",
        print_background=True,
        margin={"top": "0.45in", "bottom": "0.45in", "left": "0.5in", "right": "0.5in"}
    )
    browser.close()

print(f"Successfully regenerated 3-page PDF report at: {output_pdf_path}")
