FlexKit grid harness - client repro package
===========================================

WHAT THIS IS
------------
A standalone copy of two HomeFront screens, rebuilt as a small test harness so
the "extreme slowness" can be reproduced and measured on the machine where it
actually happens.

NO EXTERNAL DATABASE IS NEEDED. NO .NET INSTALL IS NEEDED.
Everything required is inside this folder. The harness reads two exported data
files from the Data folder and nothing else. The packaged harness contains no
connection string, server name, host, or password. By default it makes no
external network calls and listens only on this machine's loopback address
(127.0.0.1).

The SQLite provider screen creates a process-local temporary database lazily
from the same exported HomeFront estimating rows. It does not contact HomeFront,
SQL Server, or UAT, and the temporary database is removed when the bench exits.

Each large-grid page has a collapsed "Latency diagnostics" panel. Its browser
and circuit button only calls back to this same bench process. SQL, LAN, and VPN
buttons stay disabled unless an operator explicitly supplies opt-in environment
configuration as described below. Nothing runs automatically.


HOW TO RUN IT
-------------
1. Unzip the whole package to a folder you can write to,
   for example  C:\Temp\HomeFront-Bench-Repro
   Keep the folder together - do not move individual files out of it.

2. Double-click       run-bench.cmd

3. A console window opens, the harness starts, and your default browser opens
   on the first screen. Leave the console window open while you test.

4. When you are finished, click the console window and press Ctrl+C, then
   press Enter to close it.

If your browser does not open by itself, type either address below into it.


THE MAIN SCREENS
----------------
   Edit Items DB             http://127.0.0.1:5399/edit-items
      The Edit Items grid. 11,930 real item rows, 29 columns.

   Edit Items SQLite         http://127.0.0.1:5399/edit-items-provider-sqlite
      A separate provider-backed grid using SQLite COUNT and LIMIT/OFFSET
      queries. The original Edit Items screen is unchanged.

   FlexCore Enterprise Grid  http://127.0.0.1:5399/flexcore-enterprise-grid
      FlexCore-only bench with a resizable feature hierarchy on the left and
      28 isolated test screens on the right. Only the selected grid is mounted;
      every screen has focused instructions and its own Reset, with controls and
      status shown only where that test needs them.
      The screens cover ordered local multi-sort, all editing modes, validation,
      separately testable typed/menu/checklist/template filters, header bands,
      auto-generated columns, column menus and two-edge locking, complete local
      state, individual template surfaces, adaptive layout,
      accessibility, horizontal column virtualization, provider-side query,
      grouping/aggregates/distinct-filter values, and local/provider CSV/XLSX
      export. Provider query, grouping, and export use separate screens so their
      request telemetry and state do not interfere with one another. The built-in
      Custom Sort dialog opens from a grid header menu or the sorting screen's
      button. Add/delete levels, select columns and order, move levels up/down,
      then Apply or Cancel. The multi-sort check compares every exported row key
      with an independent OrderBy/ThenBy result for the chosen levels. The bench is
      intentionally omitted when the harness is built with -p:UseFlexKit=true.

      Typed filter row uses FlexCore operator dropdowns. Turn off "Show operator
      dropdowns" to use a textbox-only row with Contains for every column type.
      All bench inputs, selectors, checkboxes, and buttons use FlexCore controls.

      Important scope: numbered pager tests use the local DataSource grid;
      ItemsProvider uses Skip/Take-style virtual ranges. Provider state cannot
      restore selected/active/detail row keys that are outside the reloaded
      window. A provider checklist only permits a partial selection after the
      provider confirms the distinct-value response is complete. Full-provider
      export pages the remote query but assembles the final result in memory.

   Edit Models and Options   http://127.0.0.1:5399/edit-model-options
      The Edit Models & Options grid. 21,317 real assembly-detail rows.

Both screens are real data taken from the estimating tables. A few columns have
no stored source and are therefore blank or zero on purpose - notably Unit Cost
and Total on the Edit Models and Options screen, which the real application
calculates at query time and does not store.


WHAT TO REPORT BACK
-------------------
The blue text at the top of each screen shows a build stamp and a selection
colour, like:   build 12490 - selection: Peach
The console window prints the same expected values when it starts. Please
include that line in any report, plus which screen was slow and what you were
doing (scrolling, typing in a cell, selecting rows, sorting a column).


IF SOMETHING BLOCKS IT
----------------------
"Windows protected your PC" / SmartScreen
   Click "More info", then "Run anyway". The launcher already tries to clear
   the downloaded-file flag on every file in the folder.

"running scripts is disabled on this system"
   Use run-bench.cmd (double-click it) rather than the .ps1 file. It bypasses
   the policy for that one run only and changes nothing on the machine.

"TCP port 5399 is already in use"
   The launcher stops before starting anything and tells you so. Start on a
   different port from a PowerShell window opened in this folder:
      .\run-bench.ps1 -Port 5400
   The addresses printed by the launcher follow the port you choose.

A red "DATA NOT LOADED" banner on a screen
   The Data folder did not come across intact. Unzip the package again.


OPTIONAL LATENCY PROBES (OPERATORS / DEVELOPERS ONLY)
----------------------------------------------------
The external probes are deliberately absent from the normal repro package. To
enable one for a controlled diagnostic run, configure it with user-secrets when
running from source or with environment variables when running a published app.
Never put a password or an internal host in appsettings.json.

Database probe:
   ConnectionStrings__BenchDatabase=<complete SQL Server connection string>

The database button performs new connection timings with pooling disabled and
read-only SELECT 1 timings on one open connection. It never reads an application
table and never displays or logs the server, database, login, or connection
string.

Fixed TCP probes (example key SHAPES; values are intentionally omitted):
   LatencyProbe__TcpTargets__0__Route=Lan
   LatencyProbe__TcpTargets__0__Host=<operator-supplied LAN target>
   LatencyProbe__TcpTargets__0__Port=<operator-supplied port>
   LatencyProbe__TcpTargets__1__Route=Vpn
   LatencyProbe__TcpTargets__1__Host=<operator-supplied VPN-side target>
   LatencyProbe__TcpTargets__1__Port=<operator-supplied port>

Only the fixed labels "LAN TCP handshake" and "VPN TCP handshake" appear in
the results. Target names and addresses remain server-side. A VPN figure measures
the configured endpoint through the route available to the bench host; the bench
does not try to infer whether a VPN is connected.

Editing bench: ordinary clicks replace a multi-row selection; Ctrl/Cmd-click,
Shift-click, drag, and the selection checkboxes build a range. Batch type-over
commits with Enter or Commit active edit even without a host bulk handler.
Generated inline/dialog fields and the popup calendar use FlexCore controls.

Grid context menu: right-click a header for the standard Multi-sort section and
Custom Sort dialog. Sortable grids allow multi-sort by default. Three-dot header
buttons are opt-in through ShowColumnMenuButton and are absent from this bench.

FlexCore dedicated control benches
---------------------------------
Open /flexcore-controls, or choose "FlexCore Control Benches (37)" in navigation.
Each of the 32 remaining dedicated controls has its own route and state: inputs,
overlays, layout, upload/drop zone, signature, chat/AI input, Spreadsheet, Scheduler,
and Map. How-to-test instructions appear at the top of every bench.

Upload data stays in bench memory. Chat and AI examples are labeled local demos.
The Map bench uses local vectors. Speech recognition needs browser support and a
microphone permission; errors stay visible. No production service is configured.

These benches compile for the default FlexCore target and are omitted with
-p:UseFlexKit=true. Library API details and feature boundaries are documented in
../../FlexCore/docs/remaining-controls.md.

EditorControl parity bench
--------------------------
Open /flexcore-controls/editor for find/replace, undo/redo, read-only mode,
formatted snapshots, pagination and two-editor selection isolation.

PDF Viewer and Spreadsheet parity benches
-----------------------------------------
/flexcore-controls/pdf has a three-page sample with repeated search terms,
font changes, a form field and nested bookmarks. Test next/previous results,
Match case, zoom/rotation and saving/reloading the PDF.
/flexcore-controls/spreadsheet now includes a 120-row filter/freeze sample.
Select B2 and freeze, scroll and page in both directions, apply Region North
and Item filters together, clear one column, undo and round-trip the XLSX.
Both pages expose snapshot and API buttons and use FlexCore controls.
See FlexCore/docs/pdf-spreadsheet-parity-progress.md for implemented scope.

TreeGrid editing and hierarchy bench
-----------------------------------
Open /flexcore-controls/tree-grid-operations for inline/dialog/batch drafts,
EditContext validation, atomic saves and cancellation, add/delete subtrees,
hierarchy checkboxes, lazy check inheritance, Before/Inside/After row moves,
indent/outdent, and left/right frozen columns with resizing and view state.
