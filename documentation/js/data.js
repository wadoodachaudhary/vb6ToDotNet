// Unified data storage for HomeFront Confluence documentation system
const DOCS_DATA = [
  {
    id: "overview",
    title: "Home & Architectural Overview",
    icon: "🏠",
    path: ["Home", "Overview"],
    lastUpdated: "June 15, 2026",
    content: `
      <h2>Welcome to HomeFront Suite Documentation</h2>
      <p>The <strong>HomeFront Suite</strong> is an enterprise-grade Residential Construction Management and ERP system. It bridges the gap between estimating, sales, scheduling, purchasing, and back-office accounting for home builders. Originally written in Visual Basic 6 (VB6), the system is undergoing a modernization process, migrating to a <strong>C# ASP.NET Core Blazor</strong> architecture.</p>
      
      <div class="callout note">
        <div class="callout-icon">ℹ️</div>
        <div class="callout-body">
          <strong>Modernization Note:</strong> Active synchronization exists between <code>HomeFront</code> (Mobile & Desktop Web client) and <code>HomeFrontPB</code> (Precision Builder modern core Web client). This documentation system highlights the VB6 origin modules and their migrated Blazor counterpart targets.
        </div>
      </div>

      <h3>Suite Hierarchy & ERP Workflow</h3>
      <p>The HomeFront suite operates in a sequential transactional workflow representing the lifecycle of residential construction projects. The following diagram illustrates how the core programs interact:</p>
      
      <pre><code>
[ Estimating Engine: HFEst ]
       │  (Calculates quantities, options, and assembly components)
       ▼
[ System Administration: HFSystem ]
       │  (Handles Job/Community setups, Security, and sends Purchase Orders)
       ├─────────────────────────────────┐
       ▼                                 ▼
[ AutoNotice / Workflow ]        [ Accounts Payable: HFPayables ]
       │  (Automates alerts,             │  (Tracks invoices, commitments,
       │   emails & scheduler tasks)     │   and imports credit cards)
       ▼                                 ▼
[ BuildPro API Integrations ] ➔  [ Accounting ERPs (Sage, QBO, Dynamics D365) ]
      </code></pre>

      <h3>Programs Overview</h3>
      <table>
        <thead>
          <tr>
            <th>Program / Component</th>
            <th>Original Tech</th>
            <th>Modern Tech</th>
            <th>Core Responsibilities</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td><strong>HFEst</strong></td>
            <td>VB6 (ActiveX EXE)</td>
            <td>Blazor Razor Pages</td>
            <td>Model & Option Library, Estimating, Formula evaluation, Takeoff Quantities, and Vendor Pricelists.</td>
          </tr>
          <tr>
            <td><strong>HFSystem</strong></td>
            <td>VB6 (Standard EXE)</td>
            <td>Blazor Razor Pages</td>
            <td>Job/Lot Management, Community configurations, Security Profiles, PO creation, and ERP mapping.</td>
          </tr>
          <tr>
            <td><strong>HFPayables</strong></td>
            <td>VB6 (Standard EXE)</td>
            <td>Blazor Razor Pages</td>
            <td>Accounts Payable ledger, Commitment tracking, Invoicing, Visa import, and Payment batches.</td>
          </tr>
          <tr>
            <td><strong>AutoNotice</strong></td>
            <td>C# Class Library</td>
            <td>C# Services</td>
            <td>Background workflow scheduler, condition builder, event traps, and automated email/fax dispatches.</td>
          </tr>
          <tr>
            <td><strong>HomeFrontPB</strong></td>
            <td>-</td>
            <td>ASP.NET Blazor Core</td>
            <td>The modernized, unified web portal hosting all components using FlexKit UI and C# SQL Services.</td>
          </tr>
        </tbody>
      </table>

      <h3>High-Level Architecture</h3>
      <ul>
        <li><strong>Database Layer:</strong> Shared SQL Server Database. Heavy reliance on stored procedures for heavy operations (e.g., pricing updates, imports validation, and batch commitment runs).</li>
        <li><strong>Business Services:</strong> Originally hosted inside ActiveX DLL/EXEs; modern Blazor app serves them via singleton/scoped C# helper files (like <code>Application.cs</code> and <code>DbWrapperSqlServer.cs</code>).</li>
        <li><strong>Integrations:</strong> Multi-adapter architecture supporting Sage 300 (Timberline), Sage Intacct, QuickBooks Online, and MS Dynamics 365 (ABN).</li>
      </ul>
    `
  },
  {
    id: "hfsystem",
    title: "HFSystem - Administration & Settings",
    icon: "⚙️",
    path: ["HomeFront VB6", "HFSystem"],
    lastUpdated: "June 15, 2026",
    content: `
      <h2>HFSystem: Core Setup & Job Costing Client</h2>
      <p><code>HFSystem.vbp</code> is the administrative hub of the HomeFront suite. It establishes the organizational boundaries (companies, divisions, and communities) and coordinates the purchasing cycle by managing vendors and producing Purchase Orders (POs).</p>

      <h3>Key Components & Forms</h3>
      
      <h4>1. System Configuration & Options (<code>FOptions.frm</code>)</h4>
      <p>Manages global system flags, directory links, SMTP server parameters, PDF generator templates, and accounting system bindings. It holds configuration profiles for integrations like Sage, QuickBooks, and BuildPro.</p>
      
      <h4>2. Security & User Management (<code>FUserPermissions.frm</code>)</h4>
      <p>Configures security profiles, mapping specific groups to application forms, reports, and administrative tasks. The system features a 9-tab permissions system, including a dedicated **Data Portal** access matrix.</p>

      <h4>3. Community & Job Setup (<code>FCommunities.frm</code>, <code>FJob.frm</code>)</h4>
      <ul>
        <li><strong>Communities (Subdivisions):</strong> Defines phases, layout styles, target pricing groups, and locality dimensions.</li>
        <li><strong>Jobs (Lots):</strong> Connects a lot to a division/community, references the schedule template, shell/unit templates, and tracks status. A key feature is the <em>Precon Jobs</em> pipeline which manages template stages prior to physical construction start.</li>
      </ul>

      <h4>4. Vendor Directory & Purchasing (<code>FVendor.frm</code>, <code>FPOIndex.frm</code>, <code>FPurchaseOrder.frm</code>)</h4>
      <ul>
        <li><strong>Vendor Profile:</strong> Houses address data, tax parameters, and default trades/cost code assignments.</li>
        <li><strong>PO Hub:</strong> Standard interface to search, audit, cancel, and email purchase orders. Cancelling a PO clears its BuildPro transmission timestamp (<code>DateSentToBuildPro = NULL</code>) and updates the record timestamp.</li>
      </ul>

      <div class="callout tip">
        <div class="callout-icon">💡</div>
        <div class="callout-body">
          <strong>Database Locking Mechanism:</strong> Admin functions make use of the <code>FTableLocks.frm</code> system. When performing critical tasks (like updating community-wide pricelists or schema upgrades), the application inserts rows in the lock directory to prevent concurrent edits.
        </div>
      </div>

      <h3>Program Hierarchy Map</h3>
      <pre><code>
HFSystem (Main Entry Point)
 ├── FLogin (Authentication & warning flags)
 ├── FMain (Top navigation toolbar and menu dispatching)
 │    ├── FCompany (Global details)
 │    ├── FOptions (Accounting system, DMS, & BuildPro setup)
 │    ├── FUserPermissions (Tabbed access rights registry)
 │    ├── FCommunities (Community phases & tax setups)
 │    │    └── FDefaultVendors (Default cost code assignments)
 │    ├── FJob (Lot status, templates, and Precon milestones)
 │    ├── FPOIndex (PO index, filter, and batch actions)
 │    │    └── FPurchaseOrder (PO editor and lineitems review)
 │    └── FDbUpgrade (Schema checking & script executor)
 └── MMain.bas (System startup routines & configuration loading)
      </code></pre>

      <h3>Schema Upgrade Scripts</h3>
      <p>HFSystem maintains legacy upgrade scripts (<code>MDBUpgrade2/3/4/5.bas</code>) that run DDL scripts (e.g. <code>ALTER TABLE</code>, <code>CREATE TABLE</code>) on startup. In the modern Blazor app, these schema updates are handled via separate migration scripts, with checks to ensure tables like <code>D365FinancialDimensionValues</code> are present.</p>
    `
  },
  {
    id: "hfest",
    title: "HFEst - Estimating & Takeoff Engine",
    icon: "📐",
    path: ["HomeFront VB6", "HFEst"],
    lastUpdated: "June 15, 2026",
    content: `
      <h2>HFEst: Model & Estimating Engine</h2>
      <p><code>HFEst.vbp</code> is the core estimating engine. It allows builders to create a standardized library of building options (assemblies) and run calculations to determine exact construction quantities, material takeoffs, and cost budgets for specific jobs.</p>

      <h3>Core Components</h3>

      <h4>1. The Assembly Library (<code>FAssembly.frm</code>)</h4>
      <p>An **Assembly** represents a construction component (e.g., "Standard Exterior Wall" or "3/4 inch Plumbing Rough-in"). Each assembly contains items (materials and labor) tied to cost codes, phases, and formula expressions.</p>
      
      <h4>2. Formula Parsing Engine (<code>MFormulas.bas</code>)</h4>
      <p>The system evaluates algebraic formulas using a custom interpreter in <code>MFormulas.bas</code>. Formulas reference dimensions (e.g., width, height, perimeter, wall-area) and compute exact quantities for materials.</p>
      
      <div class="callout note">
        <div class="callout-icon">📝</div>
        <div class="callout-body">
          <strong>Formula Variables:</strong> Formulas parse standard arithmetic, logic conditions, and custom variables defined in <code>FFormulaEditor.frm</code>. Variables reference parameters bound to the job lot profile or option selections.
        </div>
      </div>

      <h4>3. Takeoff Worksheets (<code>FTakeoff.frm</code>)</h4>
      <p>A **Takeoff** is a sheet where the estimator enters the dimensions of a job. The system processes the dimensions through the assembly formulas, automatically creating items, quantities, and starting costs to compile the budget.</p>

      <h4>4. Pricing & Adjustments (<code>FAddPricelist.frm</code>, <code>FPriceList.frm</code>)</h4>
      <p>Allows estimators to adjust vendor pricing globally or by division. The estimating client was updated to support **multi-assembly pricing sheets** (instead of single-assembly selection), enabling users to edit vendor pricing across multiple assemblies and models in a single grid.</p>

      <h3>Data Import Wizards</h3>
      <ul>
        <li><strong>BIM Import (<code>FBIMImport.frm</code>):</strong> Reads Building Information Modeling (BIM) export files and creates takeoff models automatically.</li>
        <li><strong>Assembly Import (<code>FAssemblyImport.frm</code>):</strong> Wizard that uploads external Excel/CSV spreadsheets and maps them into staging tables (<code>ImportedAssemblies</code>, <code>ImportedAssemblyTakeoffItems</code>) before validating and committing them to the live catalog.</li>
      </ul>

      <h3>Program Hierarchy Map</h3>
      <pre><code>
HFEst (Estimating Suite)
 ├── FMain (Estimating menu bar and dashboards)
 │    ├── FAssembly (Model & Option Library catalog)
 │    │    └── FAssemblyCosts (Cost breakdown and overhead margin grids)
 │    ├── FTakeoff (The estimating worksheet for jobs)
 │    │    └── FFormulaEditor (Formula and parameter script builder)
 │    ├── FAddPricelist (Multi-assembly vendor pricelist editor)
 │    ├── FAssemblyImport (Spreadsheet data import wizard)
 │    ├── FBIMImport (BIM schema parser)
 │    └── FAdjustPrices (Mass cost-adjustment tool)
 ├── MFormulas.bas (Equation parsing and variable evaluation engine)
 └── MAssemblyImport.bas (Staging tables validation and commit routines)
      </code></pre>
    `
  },
  {
    id: "hfpayables",
    title: "HFPayables - Accounts Payable Suite",
    icon: "💳",
    path: ["HomeFront VB6", "HFPayables"],
    lastUpdated: "June 15, 2026",
    content: `
      <h2>HFPayables: Invoicing & Disbursement Client</h2>
      <p><code>HFPayables.vbp</code> is the Accounts Payable ledger module. It acts as the gatekeeper for cash outflows, ensuring that vendor invoices match the original Purchase Orders (POs) and budgets before they are approved for payment.</p>

      <h3>Core Concepts</h3>
      
      <h4>1. Purchase Commitments (<code>FCommitments.frm</code>)</h4>
      <p>Tracks financial commitments. When a PO is sent from HFSystem, it is recorded as a commitment. HFPayables monitors the progression of each commitment as invoices are received and matched against the PO lineitems.</p>

      <h4>2. Select Invoices to Pay (<code>FSelectInvoicesToPay.frm</code>)</h4>
      <p>An interactive grid displaying outstanding invoices. Accounts Payable administrators select invoices, apply discounts, set hold codes (e.g., lien hold, quality hold), and build payment batches.</p>

      <h4>3. Visa Credit Card Importer (<code>FVisaImporter.frm</code>)</h4>
      <p>Imports credit card statements in Excel format, maps individual charges to cost codes, jobs, and divisions, and automatically creates corresponding invoice records in the system.</p>

      <h4>4. Wallet Batches (<code>FWalletBatches.frm</code>)</h4>
      <p>Aggregates approved payments into batches. These batches coordinate bank bank reconciliations and feed direct exports to accounting ERP systems like QuickBooks or Sage.</p>

      <div class="callout danger">
        <div class="callout-icon">🚨</div>
        <div class="callout-body">
          <strong>Backcharge Matching:</strong> The <code>FPickBackCharges.frm</code> module maps penalties or builder-repaired damages back to the vendor. Backcharges are deducted from payment balances during batch approval.
        </div>
      </div>

      <h3>Program Hierarchy Map</h3>
      <pre><code>
HFPayables (Disbursements Module)
 ├── FMain (AP menu control)
 │    ├── FCommitments (Purchase commitments ledger)
 │    │    └── FCommitment (Detail viewer and manual adjustments)
 │    ├── FImportInvoice (Batch OCR or EDI invoice imports)
 │    ├── FSelectInvoicesToPay (Payment selector, hold flagger, & terms editor)
 │    ├── FWalletBatches (Bank draft batches and checks generation)
 │    ├── FVisaImporter (Credit card charge allocations)
 │    └── FPickBackCharges (Debit notes mapping grid)
 ├── MAccounting.bas (Double-entry accounting calculations)
 └── MAccountingNEW.bas (Sage Intacct & modern JSON export extensions)
      </code></pre>
    `
  },
  {
    id: "workflow",
    title: "Workflow & AutoNotice Systems",
    icon: "🔄",
    path: ["HomeFront VB6", "Workflow & AutoNotice"],
    lastUpdated: "June 15, 2026",
    content: `
      <h2>Workflow & AutoNotice Systems</h2>
      <p>The <strong>AutoNotice</strong> framework is an automated messaging system written in C# (.NET). It listens for database triggers (like PO creation, job stage completion, or invoice holds) and dispatches alerts (emails, faxes, or files) using custom templates.</p>

      <h3>Components of the Workflow Suite</h3>

      <h4>1. AutoNotice Service (<code>AutoNotice.csproj</code>)</h4>
      <p>The background service that monitors events in the database. When an event fires, it evaluates user-defined rules and schedules notifications.</p>

      <h4>2. HTML Editor Control (<code>HTMLEditorControl.csproj</code>)</h4>
      <p>A Windows Forms/C# control wrapper for editing rich HTML email templates. It is embedded in the setup screens to allow editors to customize notices with variables (e.g., <code>[VendorName]</code>, <code>[JobAddress]</code>, <code>[POTotal]</code>).</p>

      <div class="callout note">
        <div class="callout-icon">ℹ️</div>
        <div class="callout-body">
          <strong>AutoNotice Schemas:</strong> Configuration is stored in <code>dsAutoEvents</code>, <code>dsAutoReports</code>, and <code>dsTriggers</code> database datasets. These map database event listeners to specific layouts.
        </div>
      </div>

      <h3>Workflow Lifecycle Diagram</h3>
      <pre><code>
[ Database Event ] (e.g., Job Phase 3 complete)
       │
       ▼
[ Trigger Trap ] (Matches event code against dsTriggers)
       │
       ▼
[ Rule Validator ] (Verifies criteria, e.g., "Vendor is ABN")
       │
       ▼
[ HTML Builder ] (Compiles template using HTMLEditorControl)
       │
       ▼
[ Dispatch Engine ] ➔ [ Email SMTP ] OR [ File Export ]
      </code></pre>

      <h3>Recent Modernization in Blazor</h3>
      <p>The **DocType-Aware Sending Wizard** (<code>FSendingWizard.razor</code>) represents the Blazor modernization of the PO/RFQ/NOI dispatch workflows:</p>
      <ul>
        <li><strong>Document Type Parameter:</strong> The page supports path routing (<code>/po-sending-wizard/{DocType}</code>), resolving labels and task configurations for **POs**, **RFQs**, and **NOIs (Notice of Intent)**.</li>
        <li><strong>NOI Processing:</strong> Collects recipients using PO distribution lists, and logs notification tracking records to coordinate with the external <code>HFSend.exe</code> utility.</li>
      </ul>
    `
  },
  {
    id: "blazor",
    title: "HomeFrontPB - Modern Blazor Architecture",
    icon: "🌐",
    path: ["HomeFrontPB", "Architecture"],
    lastUpdated: "June 15, 2026",
    content: `
      <h2>HomeFrontPB: Modern Blazor Web Portal</h2>
      <p><code>HomeFrontPB</code> is the modern web version of the suite, built using **C# Blazor Server**. It replaces the legacy VB6 client interfaces with a unified, browser-accessible ERP portal, maintaining synchronization with the desktop database.</p>

      <h3>Architecture Stack</h3>
      <ul>
        <li><strong>Frontend Framework:</strong> Blazor Server (Razor Component lifecycle).</li>
        <li><strong>Database Helper (<code>DbWrapperSqlServer.cs</code>):</strong> Custom wrapper around ADO.NET and Dapper, executing parameterized SQL, stored procedures, and handling transaction scopes.</li>
        <li><strong>UI Controls:</strong> Built on the <strong>FlexKit Control Library</strong>, ensuring uniform aesthetics and grids matching the legacy desktop features.</li>
      </ul>

      <h3>Shared Code Synchronizations</h3>
      <p>Because the modernized client is used both on desktop WebApp and Mobile platforms, page files are kept byte-identical across directories. The core Blazor page files include:</p>
      
      <table>
        <thead>
          <tr>
            <th>Blazor Component</th>
            <th>Original VB6 File</th>
            <th>Aesthetic & Structural Design</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td><code>FOptions.razor</code></td>
            <td><code>FOptions.frm</code></td>
            <td>Tabbed settings panel including masked Client Secret controls for ERP adapters.</td>
          </tr>
          <tr>
            <td><code>FJob.razor</code></td>
            <td><code>FJob.frm</code></td>
            <td>TreeView grid representing Lot construction templates and Precon schedulers.</td>
          </tr>
          <tr>
            <td><code>FPOIndex.razor</code></td>
            <td><code>FPOIndex.frm</code></td>
            <td>Interactive grid with QuickFilter toolbars and a direct bridge to the Sending Wizard.</td>
          </tr>
          <tr>
            <td><code>FCommunities.razor</code></td>
            <td><code>FCommunities.frm</code></td>
            <td>Subdivision phase editor with conditional columns (e.g. Dynamics ABN Dimensions).</td>
          </tr>
          <tr>
            <td><code>FAssemblyImport.razor</code></td>
            <td><code>FAssemblyImport.frm</code></td>
            <td>Data Import Wizard utilizing ClosedXML to parse and stage spreadsheet data.</td>
          </tr>
          <tr>
            <td><code>FAddPricelist.razor</code></td>
            <td><code>FAddPricelist.frm</code></td>
            <td>Estimating multi-assembly grid editor utilizing complex SQL Union queries.</td>
          </tr>
        </tbody>
      </table>

      <h3>FlexKit UI Compliance</h3>
      <p>Modernized pages replace HTML elements with reusable **FlexKit** controls. This ensures consistency and supports responsive behaviors:</p>
      <ul>
        <li><code>ButtonControl</code>: Custom styled, theme-aware buttons.</li>
        <li><code>TextBoxControl</code>: Text inputs with validation and auto-commit bindings.</li>
        <li><code>CheckBoxControl</code>: Uniform checkbox switches.</li>
        <li><code>DropDownListControl</code>: Standardized combobox dropdowns with nullable defaults.</li>
        <li><code>AppGridLayout</code>: Interactive data grid supporting sorting, cell edit commits, and row actions.</li>
      </ul>

      <div class="callout tip">
        <div class="callout-icon">💡</div>
        <div class="callout-body">
          <strong>Database Upgrades Whitelist:</strong> Low-level database modifications (e.g. adding columns via <code>ALTER TABLE</code> in <code>FAddProperty.razor</code>) are routed through a connection with the application name set to <code>App=HFDBUpgradeWiz</code> to pass database trigger security.
        </div>
      </div>
    `
  },
  {
    id: "database",
    title: "Database Schemas & External ERP Integrations",
    icon: "🗄️",
    path: ["Database", "Integrations"],
    lastUpdated: "June 15, 2026",
    content: `
      <h2>Database Schemas & ERP Adaptors</h2>
      <p>The HomeFront database runs on Microsoft SQL Server. Transactions are synchronized with accounting software packages using dedicated integration modules.</p>

      <div class="tabs">
        <button class="tab-btn active" onclick="switchTab(event, 'db-schema')">Database Schema</button>
        <button class="tab-btn" onclick="switchTab(event, 'erp-dynamics')">Dynamics 365 (ABN)</button>
        <button class="tab-btn" onclick="switchTab(event, 'erp-sage')">Sage & QBO Adapters</button>
        <button class="tab-btn" onclick="switchTab(event, 'bp-portal')">BuildPro Solutions</button>
      </div>

      <!-- Tab Content 1 -->
      <div id="db-schema" class="tab-pane active">
        <h3>Core Schema Tables</h3>
        <p>Key tables that store ERP and Estimating data:</p>
        <table>
          <thead>
            <tr>
              <th>Table Name</th>
              <th>Description</th>
              <th>Key Columns Introduced</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td><code>D365FinancialDimensionValues</code></td>
              <td>Stores financial dimension keys imported from MS Dynamics ERP.</td>
              <td><code>LegalEntity</code>, <code>Dimension</code>, <code>Value</code>, <code>Description</code></td>
            </tr>
            <tr>
              <td><code>ImportedAssemblies</code></td>
              <td>Staging table for Excel imports.</td>
              <td><code>SessionID</code>, <code>AssemblyCode</code>, <code>Description</code>, <code>ModelCode</code></td>
            </tr>
            <tr>
              <td><code>system_setup</code></td>
              <td>Global parameters and options settings.</td>
              <td><code>ABN_TenantID</code>, <code>ABN_ClientID</code>, <code>UseBPDocManagment</code></td>
            </tr>
            <tr>
              <td><code>communityphase</code></td>
              <td>Tracks subdivision development stages.</td>
              <td><code>ABN_D03CostCentre</code>, <code>ABN_D06Brand</code></td>
            </tr>
          </tbody>
        </table>
      </div>

      <!-- Tab Content 2 -->
      <div id="erp-dynamics" class="tab-pane">
        <h3>Microsoft Dynamics 365 Business Central -- ABN</h3>
        <p>Introduces a modern accounting integration mapping divisions to D365 Financial Dimensions (Legal Entities and Cost Codes). Configured in <code>MD365_ABN.bas</code> and Blazor options page:</p>
        
        <ul>
          <li><strong>Global Options (system_setup):</strong> Requires 6 credentials: <code>ABN_Resource</code>, <code>ABN_TenantID</code>, <code>ABN_ClientID</code>, <code>ABN_ClientSecret</code> (stored masked), <code>ABN_LegalEntity</code>, and <code>ABN_D01Division</code>.</li>
          <li><strong>Dimension Filters:</strong>
            <ul>
              <li><code>ABN_D02Function</code>: Configured per community (filters dimensions where <code>Dimension = 'D02_Function'</code>).</li>
              <li><code>ABN_D03CostCentre</code> & <code>ABN_D06Brand</code>: Configured per community phase.</li>
              <li><code>ABN_D04SpendCategory</code>, <code>ABN_ProjectCategory</code>, <code>ABN_ItemCode</code>: Configured per standard cost code.</li>
            </ul>
          </li>
          <li><strong>Integration Quote Rules:</strong> Values sent to Dynamics are formatted using double-quote rules, and GUID strings (like <code>WalletPartyID</code>) are lowercased and stripped of curly braces before submission.</li>
        </ul>
      </div>

      <!-- Tab Content 3 -->
      <div id="erp-sage" class="tab-pane">
        <h3>Sage 300 / Intacct and QuickBooks Online</h3>
        <ul>
          <li><strong>Sage 300 / Timberline (<code>MTimberline.bas</code>):</strong> Uses legacy COM bindings to write Job Cost (JC) and Accounts Payable (AP) ledger transactions.</li>
          <li><strong>Sage Intacct (<code>MIntacct.bas</code>):</strong> Dispatches batch invoices. Recent updates ensure cash-basis taxes are rounded (<code>Round(Amount, 2)</code>) to avoid balance discrepancies in Intacct.</li>
          <li><strong>QuickBooks Online (<code>MQuickBooksOnline.bas</code>):</strong> Exports vendor lists, budgets, and invoices. Integrations use parent-child ID relations to format job-specific accounts in QBO.</li>
        </ul>
      </div>

      <!-- Tab Content 4 -->
      <div id="bp-portal" class="tab-pane">
        <h3>Hyphen Solutions BuildPro Portal</h3>
        <p>Integrates HomeFront with Hyphen Solutions BuildPro portal. Builds a pipeline to send purchase orders and retrieve vendor responses.</p>
        <ul>
          <li><strong>BuildPro Settings:</strong> Option parameters are configured in <code>FOptions.frm</code> (e.g. <code>BuildProWarrantyDocType</code> and <code>UseBPDocManagment</code>).</li>
          <li><strong>PO Transmission:</strong> Sending POs triggers background tasks inside the dispatch manager. In Blazor, this runs through the PO Sending Wizard.</li>
          <li><strong>PO Cancellations:</strong> Resetting a PO sets <code>DateSentToBuildPro = NULL</code>, pulling the job order back from BuildPro.</li>
        </ul>
      </div>
    `
  },
  {
    id: "migration",
    title: "VB6 to Blazor Migration Registry",
    icon: "🔄",
    path: ["Database", "Migration Registry"],
    lastUpdated: "June 15, 2026",
    content: `
      <h2>VB6 to Blazor Migration Registry</h2>
      <p>This registry details the progress of the migration from the legacy VB6 codebase to modern ASP.NET Core Blazor pages, detailing file-level mapping and verification statuses.</p>

      <h3>Migration Progress Matrix</h3>
      <table>
        <thead>
          <tr>
            <th>VB6 Module / Form</th>
            <th>Modern Blazor Target</th>
            <th>Migration Status</th>
            <th>Recent Fixes Applied</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td><code>FOptions.frm</code></td>
            <td><code>FOptions.razor</code></td>
            <td>✅ Complete</td>
            <td>Added Dynamics D365 ABN fields, masked secrets, and verified AppOptions serialization.</td>
          </tr>
          <tr>
            <td><code>FJob.frm</code></td>
            <td><code>FJob.razor</code></td>
            <td>✅ Complete</td>
            <td>Removed Holdback/Retainage UI, renamed labels, and added Precon schedules.</td>
          </tr>
          <tr>
            <td><code>FPOIndex.frm</code></td>
            <td><code>FPOIndex.razor</code></td>
            <td>✅ Complete</td>
            <td>Added Warranty/Precon checkboxes and BuildPro company cost-code checks.</td>
          </tr>
          <tr>
            <td><code>FCommunities.frm</code></td>
            <td><code>FCommunities.razor</code></td>
            <td>✅ Complete</td>
            <td>Added ABN Function column and prefilled blank records to user selectors.</td>
          </tr>
          <tr>
            <td><code>FAssemblyImport.frm</code></td>
            <td><code>FAssemblyImport.razor</code></td>
            <td>✅ Complete</td>
            <td>Created new Blazor page with ClosedXML spreadsheet parser.</td>
          </tr>
          <tr>
            <td><code>FAddPricelist.frm</code></td>
            <td><code>FAddPricelist.razor</code></td>
            <td>✅ Complete</td>
            <td>Reworked single-assembly pricing page to multi-assembly data grid.</td>
          </tr>
          <tr>
            <td><code>FAssemblyReplicator.frm</code></td>
            <td><code>FAssemblyReplicator.razor</code></td>
            <td>✅ Complete</td>
            <td>Fixed bug by limiting the copy join to DivisionID, Community, Assembly, Model, and Option.</td>
          </tr>
          <tr>
            <td><code>FProgress.frm</code></td>
            <td><code>FProgress.razor</code></td>
            <td>✅ Complete</td>
            <td>Cleaned up status message when count equals zero.</td>
          </tr>
        </tbody>
      </table>

      <h3>Verification & Compilation Protocols</h3>
      <p>Any modifications to shared pages must build without errors across both main solutions:</p>
      <ol>
        <li><strong>HomeFront Client Solution:</strong> [HomeFront.csproj](file:///Users/wadood/projects/VBToCSharp/HomeFront/MobileSource/HomeFront/HomeFront.csproj)</li>
        <li><strong>Precision Builder Portal Solution:</strong> [HomeFrontPB.csproj](file:///Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontPB/HomeFrontPB.csproj)</li>
      </ol>
      
      <pre><code>
# Recommended CLI Verification Commands
cd /Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontPB
dotnet build

cd /Users/wadood/projects/VBToCSharp/HomeFront/MobileSource/HomeFront
dotnet build
      </code></pre>

      <div class="callout note">
        <div class="callout-icon">ℹ️</div>
        <div class="callout-body">
          <strong>Database Consistency:</strong> Verify database scripts against the development schema. Ensure staging tables (e.g. <code>ImportedAssemblies</code>) exist in your SQL Server database target before running spreadsheet import processes.
        </div>
      </div>
    `
  }
];
