FlexKit grid harness - client repro package
===========================================

WHAT THIS IS
------------
A standalone copy of two HomeFront screens, rebuilt as a small test harness so
the "extreme slowness" can be reproduced and measured on the machine where it
actually happens.

NO DATABASE IS NEEDED. NO .NET INSTALL IS NEEDED.
Everything required is inside this folder. The harness reads two exported data
files from the Data folder and nothing else. The packaged harness contains no
connection string, server name, host, or password. By default it makes no
external network calls and listens only on this machine's loopback address
(127.0.0.1).

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


THE TWO SCREENS
---------------
   Edit Items DB             http://127.0.0.1:5399/edit-items
      The Edit Items grid. 11,930 real item rows, 29 columns.

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
