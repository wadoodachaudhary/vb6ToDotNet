# VB6 Fidelity Audit — HomeFront forms vs VB6

Read-only audit comparing each migrated `Components/Pages/Migrated/F*.razor` to its
original VB6 `.frm` (NEW tree: `HomeFrontVB6/HFEst|HFSystem/Source/*.frm`, never
`/Original`). Looking for: **[HIGH]** missing/stubbed business logic + missing/altered
DB queries; **[MED]** grids that should be AppGridLayout-dynamic (FAssembly pattern)
but are hardcoded, and missing FPickList pickers; **[LOW]** raw HTML controls that
should be FlexKit.

Started 2026-06-11. Direction = VB6 → HF only. Fixes applied **only after user
green-light**. PB synced only for forms already in PB when touched; FlexKit↔FlexCore
byte-identical. Excluded (drifted intentionally): report pipeline + FItemChart.

---

## Progress

| Wave | Cluster | Forms | Status |
|------|---------|-------|--------|
| 1 | Quotes | FQuote, FCustomQuote, FInboxCustomQuote | ✅ audited |
| 1-fix | FQuote in-grid pickers + recalcs | FQuote | ✅ implemented (build clean) |
| 1-fix | FQuote ColorizeItems + context menu + RefreshCosts/ZeroRate + Attachments | FQuote | ✅ implemented (build clean) — deferred: Update-Pricelist ctx action (grid-model mismatch) |
| 1-fix | FCustomQuote full audit + gaps (pickers, recalcs, ColorizeItems, ctx menu, COMarkup seed) | FCustomQuote | ✅ implemented 948→1806 lines (build clean) — deferred: Preview (Crystal `CustomPreEstimate.rpt`), selection-totals tooltip |
| 1-fix | FInboxJobs audit+fix (parameterized SQL, removed bogus `Sales_SetEstimateIndex` call that threw on every save, verified proc param order) | FInboxJobs | ✅ implemented HF + mirrored to PB (both build clean) |
| 2 | **Setup group — biggest-gaps-first** | FVendor → FJob → FExportPricelists → rest | in progress |
| 2 | FVendor: ~17 raw→FlexKit, 3 IniGetGrid grids made editable + in-cell pickers (PO Index/Community/Tax Group), full SaveData port, fixed SMSAddress load bug | FVendor | ✅ HF + mirrored PB (build clean) — residual: 3 grid-cell `<input type=checkbox>` (stopPropagation) |
| 2 | FJob targeted: AR Customer picker (was dead-nav)+create-customer, Tarion category load+save, Intacct Department row (latent save bug). Deferred external accounting job-number lookups | FJob | ✅ HF + mirrored PB (build clean) |
| — | **Off-limits (user, 2026-06-11):** FItems, FAssembly (= "Edit Models and Options"), FTakeoff — near-final, reviewed separately; do NOT edit | — | — |
| 2 | FProjectManagers: hand-rolled `<div>/<table>` purchaser modal → FlexKit FPickList; cascade verified | FProjectManagers | ✅ HF+PB |
| 2 | FCommunities: **GST/PST column visibility was hardcoded** → now data-driven (`Show_CDN_GST`/`UsePst`); file-browse honest defer (desktop path) | FCommunities | ✅ HF+PB (.razor; .css pre-existing diff) |
| 2 | FIntersection: 6 raw controls + a raw `<div>` overlay modal → FlexKit (TextBoxControl + in-cell FPickList "PO Index" + FPickList multi-view "Assembly"); PO-Index fan-out to selected rows | FIntersection | ✅ HF+PB |
| 2 | FDefaultVendors: hardcoded grid → AppGridLayout-dynamic; added Vendor FPickList; **fixed DivisionID hardcoded to 1** + global-default column was un-editable (239 assignments unsettable); parameterized SQL | FDefaultVendors | ✅ HF+PB |
| 1 | **FVendorPriceList purged** — file deleted (both apps), all refs → FPriceList (canonical VB6 FPriceList.frm), workflow routes to FPriceList. FPriceListLegacy = user's untracked retired backup (flagged) | — | ✅ |
| 2 | FPOIndex (PO Index editor hosted by FPOFormats): 19 raw controls → FlexKit; **LoginID hardcoded "SYSTEM"** + QuickBooks combo-display + JS confirm()→MessageBox + mDirty ordering bug all fixed; public API preserved | FPOIndex | ✅ HF+PB |
| 2 | FAssemblyImport (4 modes: import/export assemblies+takeoffs): **PIVOT injection hardened** (AssemblyID int → validate/drop) + column-name allowlist; **multi-sheet `NewAssmImports.xlsx` template fully built** (ClosedXML: Assemblies + protected DataValidation sheet + 8 dropdowns) | FAssemblyImport | ✅ HF+PB |
| — | FPOFormats (PO Indexes) = done externally; confirmed FlexKit-clean, no changes | FPOFormats | ✅ |
| infra | Pinned `ClosedXML 0.105.0` directly in HomeFront.csproj + HomeFrontPB.csproj (was transitive via FlexKit). Kept ClosedXML (it wraps MS's DocumentFormat.OpenXml SDK) per user. | — | ✅ |

### 🎉 SETUP SIDEBAR GROUP COMPLETE
Project Managers, Community setup, Job setup, Vendor setup, PO Indexes (FPOFormats+FPOIndex), Export/Import Assemblies, Export/Import Assembly Takeoffs (FAssemblyImport), Option Intersections, Default Vendors — all ✅. ("Edit Models and Options"=FAssembly and "Edit item database"=FItems are the off-limits near-final pair, handled separately.)

| — | Audit continues: other menu groups (Tasks/Sales-Pricing/Inquiries/Reports) + pending FPriceList review | … | pending |
| — | Pending review: **FPriceList** ("looks good" per user — review when ready) | FPriceList | — |
| — | (remaining forms; FInboxCustomQuote still pending in Quotes cluster) | … | pending |

---

## 🚨 Systemic finding (2026-06-11): DB functions gained a trailing `@DivisionID`

The updated DB redefined two scalar functions with an extra final parameter:
- `dbo.Purch_GetItemRate(...)` now takes **13** args (added `@DivisionID`)
- `dbo.Purch_GetDefaultTaxGroup(...)` now takes **10** args (added `@DivisionID`)

Scalar UDFs can't omit trailing args, so any call using the **old VB6 arity throws at
runtime** ("insufficient number of arguments") — caught/logged, so the recalc silently
does nothing. Swept all call sites across HF+PB:
- **Already on the new signature (OK):** FAssembly, FAssemblyOpus, EditModelAssembly,
  ModelAssembly, FAddPricelist, FInboxTBDAssignments, FVendorChange, FPOPriceChangeWiz,
  FTakeoff, FTakeoffOneTime.
- **Fixed (were old-arity → would throw):** **FQuote** (4 calls: 2 RePrice + GetVendorCost
  + GetDefaultTaxGroup), **FEstimateItems** (2 calls in RePrice). Appended `@DivisionID`
  (`SessionState`/`Session.DivisionID`) + verified empirically against the DB.

## Wave-1 FIX applied — FQuote in-grid pickers + recalcs (HF-only)
Implemented (mirroring FItems' `_lookupPickerFields`/`_popupTextFields` pattern + VB6
`gItems_CellButtonClick`/`gItems_AfterEdit`): cell pickers for Cost Code, Category, UOM,
Vendor (2-view Approved/All), Tax Group, PO Index, Extra; popup text for ItemDesc/
ItemComments; recalcs for JCCategory→`Purch_GetDefaultTaxGroup`, Vendor→`Purch_GetItemRate`
(GetVendorCost), TaxGroup→`TaxGroups` JCRate/NJCRate. Picks routed separately from the
existing Open-Quote dialog. Build clean (0 errors). **Still TODO on FQuote** (from the
audit, not yet done): ColorizeItems via `RowCssClassSelector`, right-click context menu,
RePrice `ZeroRateOnRefreshCosts` branch + checked-row scope, Attachments. FCustomQuote &
FInboxCustomQuote fixes also pending.

---

## Wave 1 — Quote cluster

**Grids are fine (cat 3 clean):** all three build `gItems`/`gData` dynamically from
`GridLayoutSvc.GetGridLayoutAsync(...)` with the FAssembly helper pattern — hardcoded
column lists are fallback-only. No grid-dynamic findings.

### FQuote — ❌ significant gaps
- **[HIGH] In-grid cell pickers entirely missing** (VB6 `gItems_CellButtonClick`, FQuote.frm:1954-2076): Vendor (2-view Approved/All), Cost Code, Category, Tax Group, PO Index, Order UOM, Extra, plus FComments for ItemDesc/ItemComments. HF grid is flat text-edit; the lone `<FPickList>` serves only the Open-quote dialog. (Also cat-4.)
- **[HIGH] `GetVendorCost` not implemented** (FQuote.frm:2512, `Purch_GetItemRate`) — rates go stale on vendor change.
- **[HIGH] `JCCategory`→default tax group missing** (`gItems_AfterEdit` "JCCategory", .frm:1498, `Purch_GetDefaultTaxGroup`); no `BudgetTaxGroup`/`POTaxGroup` rate-lookup recalc (.frm:1722).
- **[HIGH] `gItems_ValidateEdit` dropped** (.frm:2170) — no per-cell DB validation; no JCExtra add-new flow.
- **[HIGH] WastePercent not loaded** in `LoadItems` → budget-qty waste silently lost.
- **[HIGH] ColorizeItems reduced to a "⚠" glyph** (.frm:2410): per-row fore/back/italic/bold + grey-locked from `HFApp.Options(Format_*)` not applied; `.row-*` CSS classes in FQuote.razor.css are dead (no `RowCssClassSelector` wired — FlexKit `GridControl` supports it).
- **[HIGH] Right-click context menu missing** (`mnuEstimateItemsGridSub_Click`, .frm:2543): Copy Items, Change Item, Save One-Time to DB, Compare Prices (FPriceComparison), Update Price List (FPriceListUpdate), Files (FAttachments).
- **[HIGH] RePrice wrong scope + dropped option**: ignores checked-row `WhereClause` (updates ALL `BudgetGenerated=0` rows), omits the `ZeroRateOnRefreshCosts` branch, no "N updated" report (.frm:728-793).
- **[HIGH] Attachments stubbed** (`OnAttachmentsClick`→Notifications placeholder).
- **[LOW] FlexKit:** header uses raw `<input>/<select>/<textarea>`; **save-confirm is a raw `<div position:fixed inset:0>` overlay (RED FLAG → DialogControl)**; JS confirm/prompt/alert instead of DialogControl/MessageBoxControl.
- **Top fix:** restore the `gItems` in-grid pickers + `Purch_GetItemRate`/`Purch_GetDefaultTaxGroup`/tax-rate recalcs + wire `RowCssClassSelector` for ColorizeItems.

### FCustomQuote — ❌ significant gaps
- **[HIGH] All in-grid cell pickers missing** (VB6 `gItems_CellButtonClick`, .frm:1066-1179): Vendor 2-view, Cost Code, Category, Order UOM, Tax Group, PO Index, FComments. (cat-4.)
- **[HIGH] `GetVendorCost` (Purch_GetItemRate) + `JCCategory`→default tax group + `BudgetTaxGroup` JCRate/NJCRate lookup** all missing in `RecalcRow`.
- **[HIGH] `TakeoffQty` recalc drops WastePercent + RoundTo/RoundDir** (.frm:909 vs .razor) → budget quantities differ.
- **[HIGH] `gItems_ValidateEdit` dropped** (.frm:1294).
- **[HIGH] ColorizeItems/warnings entirely absent** (.frm:1395) — no warning column/glyph at all.
- **[HIGH] Attachments stubbed** (Notifications.Info). Preview = DEFERRED (external Crystal).
- **[LOW] FlexKit:** info panel is a raw `<table>` (VB6 `gInfoPanel` VSFlexGrid); JS confirm on save. (Otherwise good FlexKit adoption: NumericTextBox/DropDownList/GridControl.)
- **Top fix:** same as FQuote — pickers + Purch_* recalcs + fix TakeoffQty waste/round + ColorizeItems.

### FInboxCustomQuote — ⚠️ minor gaps
Strong port (GridControl/DropDownList/DialogControl/MessageBoxControl/ButtonControl; per-Location UPDATE chain + Calc procs + view SQL faithful).
- **[HIGH] "Build Quote" loses modal return values** (VB6 `gData_KeyDown` Return, .frm:658-680): VB6 calls `FCustomQuote.BuildQuote(...)` modally and writes back Cost/Price/Qty/UOM + `RowData="Dirty"`. HF just `NavigateTo("/custom-quote?…")` — inbox row never updated/marked dirty. Also the URL **drops** Model/OptionID/Extra/Category/CommunityPhase args. → Fix: host FCustomQuote as a **modal child** (FlexKit "host F* as modal" pattern) and write results back.
- **[HIGH] Declined-status flow drift**: VB6 prompts reason synchronously and cancels edit if blank; HF commits cell + sets Price=0 before confirming, leaving Declined/price-0/no-reason if user cancels.
- **[LOW]** email via `mailto` JS (DEFERRED external mail, weak headless substitute); subject literal differs.

---

## Cross-cutting (Quote cluster)
Highest-value fix: restore the **`gItems` in-grid editing surface** — cell pickers
(Vendor/Cost Code/Category/Tax Group/PO Index/UOM/Comments), `Purch_GetItemRate` /
`Purch_GetDefaultTaxGroup` recalcs, and `RowCssClassSelector`-driven ColorizeItems.
Both FQuote/FCustomQuote render AppGridLayout columns correctly but degrade the grid
to flat text-edit, dropping the core estimating workflow. FlexKit already has the
needed primitives (`FPickList`, `RowCssClassSelector`, cell `EditTemplate`).

---

## No-defer closure campaign (2026-06-12) — COMPLETE
Per user directive "Whatever VB6 is doing HF has to do. Do not defer anything. Flag it
if difficult / you need tools, but do not defer it." Every remaining `DEFER`/`stub`/
`fallback`/`not implemented` marker across the touched Setup+Quote forms was closed
against VB6 behavior, or — for genuinely-unavailable externals — converted to a
VISIBLE blocked-flag (not a silent swallow). Excel work standardized on **ClosedXML**
(pinned in both csproj; no truly built-in .NET xlsx writer exists). Decisions locked:
external accounting integration = flag-as-blocked; file-browse = list the real wwwroot
folder; tiny JS = allowed only where no C#/CSS path (one helper per need, C# fallback).

| Form | Closed items | JS added | PB mirror |
|------|--------------|----------|-----------|
| FVendor | Export(ClosedXML)/Import nav/Save YesNoCancel/PO-Format dir-list; flagged Timberline/APM sync | — | ✓ |
| FQuote | pickers + Purch_* (@DivisionID) + ColorizeItems + ctx menu + RefreshCosts + Attachments + Preview + Update-Pricelist | — | HF-only |
| FCustomQuote | full pickers/recalcs/ColorizeItems + Preview + selection-totals | — | HF-only |
| FCommunities | Contract-Document browse (dir-list+FPickList) + WarrantyJob inline-validate + find-highlight | hfScrollSelectorIntoView | ✓ |
| FDefaultVendors | find-highlight + internal copy/paste cell-tiling (no OS clipboard → no JS); outline-bar N/A (flat matrix, grouped=0) | — | ✓ |
| FPOIndex | select-all-on-focus (VB6 GotFocus×8) via one container-level focusin helper | hfSelectOnFocus | ✓ |
| FJob | revived 3 dead reads (AccountingSystem/SalesSystem + ValueByName two-tier + MultiFamily bit); accounting browse = visible [BLOCKED] flag gated on AccountingSystem≠asNone | — | ✓ |

**FJob deep find:** the old code read `AccountingSystem`/`SalesSystem` + Precon/Tarion/
JobContactRoles/ScheduleTemplates from **`system_setup` columns that don't exist** →
every read threw into a silent `catch` → the whole accounting-conditional form was
dead. Now reads `AccountingSystem`(singular)/`SalesSystem` from **AppOptions** with the
VB6 `ValueByName` two-tier (system_setup col → AppOptions `IN(0,@Div) ORDER BY divisionid
DESC`); enum verified vs VB6 Application.cls (asPeachtree=6/asXero=7 gap matches). For the
active div 1 (AccountingSystem=asNone) the external browse correctly stays hidden.

Both HomeFront.sln and HomeFrontPB.sln build **0 errors** after every form. FItems/
FAssembly/FTakeoff left untouched (near-final, reviewed separately). FlexKit/FlexCore
untouched by this campaign (no host-specific code added to the library).

---

## Setup-menu sweep (2026-06-12) — COMPLETE
After the Setup *sidebar* group, swept every remaining Setup *menu-bar* form. 7 forms
audited vs VB6 (DB/logic/dynamic-grids/PickList/FlexKit), no-defer. ("Work Region Setup"
is just an alias for Community Setup → FCommunities, already done.)

| Form | VB6 | Verdict / key fixes | Grid | PB |
|------|-----|---------------------|------|----|
| FCustomer | FCustomer.frm | broken AppOptions query (Name/Value cols don't exist → fabricated role fallback) fixed; PickList → VB6 views (Picklist_(In)ActiveARCustomers); raw inputs → FlexKit; JS confirm → DialogControl Yes/No/Cancel; stopped writing Inactive (VB6 never does) | gContacts dynamic AppGridLayout | ✓ |
| FWBSCodes | FWBSCodes.frm | raw div-modal/inputs → DialogControl+TextBoxControl; JS history.back → ParentNavigate; 40-way WBSDesc UNION preserved | fixed 2-col (faithful) | HF-only |
| FDimensionCategories | FDimensionCategories.frm | 3 raw tables → GridControl; gCategories now dynamic; Shift+Del UOM wired; all SQL parameterized | gCategories dynamic | HF-only |
| FCommunityStandards | CommunityStandards.frm | hardcoded grid cols → dynamic AppGridLayout; PhaseSpecific option read (was hardcoded); hand-rolled picklist → FPickList; notes → FComments; JS confirm → MessageBox | gData dynamic | HF-only |
| FModelDimensions | FModelDimensions.frm | raw overlay → DialogControl; pickers → FPickList; JS confirm → MessageBox; IN-clause parameterized | gModels/gCategories dynamic | ✓ |
| FImportAssemblies | FImportAssemblies.frm | ClosePage infinite-recursion fix; JS file-picker → InputFile; hardcoded AppOptions → real reads; tax math corrected to CalcPreTax/CalcTax; WritePhases + FixGroupPhaseValue restored; validation parity; **Excel via ClosedXML (read-only — VB6 only reads)** | fixed (Excel-populated, faithful) | ✓ |
| FDBGrid | FDBGrid.frm | generic editor: added 7 missing entity captions (Categories/Tax Groups/Book of Accounts/WBS Codes/Room Locations/Series/Option Subcategories); post-load side-effects; IsPersistedColumn guard (fixes "Invalid column Sage300GLPrefixes"); nested Sage300 picker; save-prompt on close; allowlist table mapping | gData dynamic | ✓ |

**Incidental no-defer fixes (same bug class, closed):**
- **FContacts.razor** — identical broken `SELECT Value FROM AppOptions WHERE Name='JobContactRoles'` (wrong cols → always threw → fabricated "Buyer|Cobuyer|…" default). Fixed to OptionValue/OptionName + IN(0,@Div) ORDER BY divisionid DESC, parameterized; removed the fabricated fallback (VB6 has none). HF+PB.
- **FOptions.razor** — "Wallet Configuration" + "Tax groups..." links were empty stubs. Now host FDBGrid as a modal (VB6 FDBGrid.ShowForm vbModal, FOptions.frm:6655-6685): Book of Accounts via caption auto-resolve, Tax Groups via explicit ShowForm with hiddenColumns="divisionid"; reload BookOfAccountList / run grouprate update on close. HF+PB.

**FLAGGED (not fixed — need sign-off / off-limits):**
- **FAssembly.razor caption mismatch** (OFF-LIMITS): its "Cat"/"Sub Cat" links send "Option Groups"/"Option Categories" but VB6 FAssembly maps lblGroup→OptionCategories table, lblCategory→OptionSubcategories. To fix during the FAssembly review.
- **FUserPermissions**: `.lic` license-file subsystem (LicensedSeats/CLicense) + CRM Sales_DeactivateUser outbound API — not DI-wired, no license table; safe documented fallbacks in place.
- **FModelDimensions / FDBGrid context menus**: FlexKit has no context-menu primitive — using minimal CSS popups (documented exception). A `ContextMenuControl` could be added to FlexKit+FlexCore in a separate task.

Both HomeFront.sln and HomeFrontPB.sln build **0 errors** throughout. FItems/FAssembly/FTakeoff untouched.

---

## Setup-menu routing audit + FJob re-verification (2026-06-12)

**FJob re-verification (answering "did VB6 also reference non-functional columns?"):** YES.
VB6 `Options.ValueByName` (Options.cls:796-815) runs `select <OptionName> from system_setup`
as tier-1 under `On Error Resume Next`, so a non-existent column throws and is silently
swallowed → falls through to tier-2 `select optionvalue from appoptions … IN(0,@div) ORDER BY
divisionid DESC`. The `Refresh` loads `mOptions` from BOTH system_setup columns AND AppOptions
rows; save writes system-setup options as `UPDATE system_setup SET <Name>=` else division-specific
`INSERT/UPDATE AppOptions`. AccountingSystem/SalesSystem live in AppOptions (verified: not
system_setup columns), so VB6 reads them from the AppOptions-loaded cache. **The FJob fix's
two-tier `GetOptionValue` (system_setup col → AppOptions) is a faithful port of ValueByName;
the original migration bug was dropping the AppOptions tier.** Also confirmed none of the other
swept option names (CustomerContactRoles/JobContactRoles/BuilderType/RefreshOverridden/
ApplyVarianceCategory/ChangeWizCategory/BuildProSendPOsImmediately) are system_setup columns,
so their single-tier AppOptions reads match VB6's two-tier result.

**Setup MENU-BAR routing (VB6 `mnuSetup`, FMain.frm:2778-2810) — fixed 3 mis-routes:**
- "Job Cost Codes" → was `/wbs-codes` (FWBSCodes — the job-level WBS-description editor, opened
  from FEstimateItems, NOT Setup). Now `/db-grid/Cost Codes` (FDBGrid standardcostcodes) per VB6
  `EditJCCostCodes`.
- "Job Cost Categories" → was `/dimension-categories` (wrong table). Now `/db-grid/Categories`
  (FDBGrid standardcategories) per VB6 `EditJCCategories`.
- "Option Attributes" → was unwired (no case). Now `/attribute-lists` (FAttributeLists) per VB6
  `EditAttributeLists`.
- Menu-bar dropdown reconciled to match `mnuSetup` exactly (order + BuilderType/UseCommunityStandards
  gates): added Rooms/Dimension Categories/Model Dimensions/Option Attributes/{Community} Phases Setup.
- **NOTE — two distinct VB6 Setup surfaces:** the menu-bar `mnuSetup` (data-setup items: JC Codes/
  Categories, Rooms, Dim Cat, Model Dim, Option Attributes, …) vs the CommandBar/sidebar (FMain.frm:
  1886-1908: Import/Export Assemblies, Takeoffs, Option Intersections). HF's Workflow **sidebar**
  `setupSection` correctly carries the assemblies/intersections items; the menu-bar correctly does
  NOT (they're not in mnuSetup). FIntersection/FAssemblyImport/FImportAssemblies stay reachable
  via the sidebar — not orphaned.
- **FDBGrid entity SQL verified** for Cost Codes / Categories / Rooms / Community Phases against the
  VB6 CALLER SQL (Application.cls EditJC*), incl. the ABN_* aliases + accounting-system-conditional
  hidden columns. No drift.

**FAttributeLists ("Option Attributes")** — full FlexKit rewrite (was raw div/table/input): DialogControl
shell, two GridControls (lists + values), trailing-blank-row invariant, Other-List combo over all lists,
Ctrl+Del value delete, Yes/No/Cancel save prompt. Fixed design-time grid (no AppGridLayout row). HF+PB.

**PB:** FAttributeLists mirrored (was 778 lines diverged → identical, PB builds clean). **FMain routing
held for PB** — HF↔PB FMain have diverged ~922 lines (PB lacks FDimensionCategories etc.), so the
menu changes are HF-only this round; promote to PB FMain surgically in the monthly PB sync.

---

## FAssembly + FTakeoff gap-fix (2026-06-12) — off-limits lifted
User opened up the near-final forms ("Setup is done, we can take on FAssembly, FItems, PriceList").
Audited FAssembly + FTakeoff vs VB6, fixed all findings (HF; integrations out of scope).

**FItems/FAssembly menu routing:** "Edit item database" → canonical `/items-db` (was dead `/items-db-opus`); "Edit Models and Options" → `/edit-model-assembly` (FAssembly) already correct.

**Lookup mis-routes (FAssembly Cat/Sub Cat/Location):** all 3 sent wrong FDBGrid captions/tables. Fixed: Cat → "Option Categories"/`OptionCategories`; Sub Cat → "Option Subcategories"/`OptionSubcategories`; Location → "Room Locations"/`tblareas` (VB6 FAssembly.frm:2448/2458/2463).

**Caption collision (FDBGrid + FOptions, mirrored to PB):** VB6 reuses title "Option Categories" for two tables (FAssembly→OptionCategories, FOptions/FMain Setup→tblcategories), disambiguating by per-call SQL. HF resolves by caption, so the tblcategories caller now uses a DISTINCT resolver key **"Sales Option Categories"** (display title still "Option Categories"). FDBGrid: "Option Categories"→OptionCategories, "Sales Option Categories"→tblcategories. FOptions OnEditOptionCategories + FMain dead case repointed.

**FAssembly data-loss saves restored (HIGH):** WBS01..40 on item INSERT/UPDATE; `UpdateItemPrice`/`PriceLevelIndex` (tblVendorCost write-back at the right phase/community/global/corporate level); `mDataSource=1` sales write-back (tblModels/tblOptions/tblDCOptions/tblGlobalOptions); component-Qty UPDATE; Copy-of-Current clones components + RoomQtyByModel; master Save BillingCode/BillingCategory + Color/Style/Finish/Other + 4 ListIDs (FAssemblyAttributes raises OnAttributesSaved for the host to persist — now wired). Context menu Edit-Item-Chart + View-Files; ValidateData ReqCategory rule.

**FTakeoff (HIGH):** real Add-Pass **formula evaluator** ported (VB6 ResolveFormula → CalcFormula → clsEquation: token expansion, variable substitution, recursive-descent arithmetic + function library) — was collapsing every non-numeric formula to 1, producing wrong takeoff quantities. Two-pass lookup handling + p1..p5 history per VB6.

**Integrations OUT OF SCOPE (user):** Pipeline import (FAssembly OnImportModels/Globals), FTakeoff One-Time/PlanSwift/OnScreen → "not available in this version" MessageBox instead of silent stub/break.

**Deferred (flagged):** FAssembly + FTakeoff header raw `<input>/<select>/<textarea>` → FlexKit (cosmetic, separate careful pass on the reference forms). FItems content audit. PB promotion of FAssembly/FTakeoff/FMain (heavily diverged — surgical, monthly).

Both HomeFront.sln + HomeFrontPB.sln build **0 errors**. FDBGrid + FOptions mirrored to PB; FAssembly/FTakeoff/FMain held.

---

## Import/Export form merge (2026-06-12, HF only)
User request: combine the two assembly import/export forms into one clearly-named form.
- `FAssemblyImport` (4 modes: import/export × assemblies/takeoffs) + `FImportAssemblies` (legacy "Import models and options"/"Import items") → **`FImportExportAssemblies.razor`** (`/assembly-import-export/{Mode}`, ~3927 lines, 6 modes). Both old forms + routes DELETED in HF; all logic preserved verbatim (3 collision renames). FMain route map + InferParameterName + 6 Setup-menu cases repointed. Build 0 errors, zero dangling refs.
- VB6 note: these were genuinely separate programs (FAssemblyImport.frm import + MAssemblyImport.bas export subs + legacy FImportAssemblies.frm) — HF consolidation, not VB6 structure.
- **PB: held** — PB still has the old two forms; FMain ~922 lines diverged; promote the whole refactor as a unit in the monthly PB sync. (FDBGrid + FOptions caption-collision fix WAS mirrored to PB this round.)

---

## FDBGrid "Cost Codes" empty-grid + tab-title fix (2026-06-12, HF)
Reported: Job Cost Codes opened to an EMPTY grid (VB6 "Cost Codes" shows standardcostcodes) and the tab read "FDBGrid".
- **Empty grid (root cause):** the "Cost Codes" case was the ONLY FDBGrid resolver case that delegated to `Application.HFApp?.BuildCostCodesForm()`. `Application.HFApp` is a per-circuit `static Application?` (set on ctor, nulled on dispose) → null when FDBGrid renders standalone via `/db-grid/Cost Codes` → spec null → it set the caption and ran NO query → empty. Data was fine (standardcostcodes: 311 rows div 1, 166 div 0). **Fix:** inlined the SQL in the case like every other entity (SELECT with ABN_* aliased to D04_SpendCategory/ProjectCategory/ItemCode, keys divisionid,costcode, @DivisionID param, asD365_ABN(11)-conditional hidden columns via the existing `ParseAccountingSystem(GetAppOptionValueAsync("AccountingSystem"))` helper, columnAliases for the UPDATE write-back). No more HFApp dependency.
- **Tab title:** FDBGrid only set `<PageTitle>` (browser title); it had no `SetPageTitle` cascade, so FMain's MDI tab fell back to `pageType.Name` = "FDBGrid". **Fix:** added `[CascadingParameter(Name="SetPageTitle")] Action<string?>?` (matching FRptViewer) + push the resolved `Caption` to the tab at the end of OnParametersSetAsync, guarded by a `_titledCaption` field so a title-triggered re-render can't loop. Now the tab shows the entity caption ("Cost Codes", "Rooms", …) for ALL /db-grid entities.
- **PB: HELD.** Bundled with the Setup-menu routing (Job Cost Codes → /db-grid/Cost Codes) which is already held for PB (FMain ~922 lines diverged). The whole Job-Cost-Codes path + FDBGrid Cost-Codes case + title cascade promote to PB together in the monthly sync. PB still routes "Job Cost Codes"→FWBSCodes until then.

---

## Business-logic HIGH fixes (2026-06-12) — FAssembly multi-tenancy + FItems data-loss + FCommunities multi-select
Read-only audits of FItems/FAssembly/FTakeoff/FCommunities → fixed the HIGH set:

**FAssembly [HIGH multi-tenancy] — was silently operating on division 1 for everyone.**
- `AppState_DivisionID => "1"` (15 sites: item/vendor/component/substitute/picklist/combo queries + SaveData + UpdateItemPrice + sales write-back) → `SessionState.DivisionID`. Picklist-SQL sites parameterized or int-interpolated.
- `static Dictionary AppOptions` of hardcoded constants (BuilderType=Residential, EstimatingSystem=1, fabricated Series/ScheduleTemplates) → INSTANCE dict DB-loaded per division (`SELECT OptionName,OptionValue FROM AppOptions WHERE ISNULL(divisionid,0) IN(0,@Div) ORDER BY divisionid`, FCustomQuote pattern). DB had divergent per-division values (EstimatingSystem 0/2 not 1, BuilderType=Homebuilder) → HF was wrong for divs 14/15/16. Unset option → "" (VB6 ValueByName semantics). Also removes a Blazor static. **Flagged next-wave:** `mUseComponents` still hardcoded true (doesn't read UseComponents option; div 14 = False).

**FItems [5 HIGH data-loss on Save]:** WBS01-40 now in the UPDATE (loop-built params); inverse-item link persisted to tblPhaseItem (forward+reciprocal, was memory-only); price-group propagation (Price+PartNumber to all PriceLink members); `UpdateItems` cascade into CustomPreEstimateItems/EstimateItems ported; `UStmp` audit field written. All parameterized; VB6 FItems.frm refs 1449-1475/523-543/2010-2038.

**FCommunities [HIGH]:** Divisions picker restored to `multiSelect:true` + `string.Join` of all codes (was single-select → silently collapsed multi-division communities on the DivisionCommunities junction). Removed a `Console.WriteLine` that index-crashed on an empty division. **Mirrored to PB** (clean sync pair).

Both HomeFront.sln + HomeFrontPB.sln build 0 errors. **PB: FAssembly (1021 lines diverged) + FItems (161) HELD for monthly promotion**; FCommunities mirrored.

**Next-wave MED (audited, not yet fixed):** FAssembly OrderQty→TakeoffQty waste-when-conversion≠1 + Item-DB-price-editable + one-time-item validation + JS-confirm removals; FItems AltJC/OptionCategory/UOM pickers + inline value-clamping + 8 raw `<div>` modals; FCommunities sp_CanDeleteCommunity bit-compare + Ctrl+Del/Ctrl+S + navigate-away guard + column persistence; FTakeoff + FItems FlexKit raw-control conversions; integrations (Excel/Sage/Pipeline) out of scope.

---

## MED wave (2026-06-12) — FAssembly / FItems / FCommunities logic fixes
- **FAssembly:** `mUseComponents` now reads the `UseComponents` AppOption (was hardcoded true; div 14=False hides the components grid, VB6 frm:1617/3136). OrderQty→TakeoffQty back-solve always divides by waste regardless of conversion + rounds to 5 (frm:2006). Item-DB rows are no longer Price-editable (frm:2424 — markup gate + OnCellEdit cancel + commit guard + VB6 message via _messageBox). ValidateData one-time-item = `(Phase|| Item empty) && ItemChart empty` (frm:4104). Item-removal raw `confirm()` → MessageBox.
- **FItems:** added `…` pickers + validation for AltJCCostCode/AltJCCategory (StandardCostCodes/Categories IN(0,@Div)) + OptionCategory (tblCategories, no div filter, frm:586) + OrderUOM/TakeoffUOM DISTINCT pickers (frm:557-631); inline value clamping on commit (ConversionFactor≤0→1, WastePercent Abs/clamp99, RoundTo Abs, Price≥0; frm:737-757); success toast moved off the Error channel.
- **FCommunities:** `sp_CanDeleteCommunity` read as `bool` (was fragile bit-as-string); Ctrl+Delete keyboard path + immediate row-hide (`.row-hidden`); `<NavigationLock>` unsaved-changes guard + Ctrl+S (no JS); GST/PST raw `<input>` → `NumericTextBoxControl`. **Mirrored to PB.**

Both .sln build 0 errors. PB: FCommunities mirrored; FAssembly/FItems HELD (diverged) for monthly.

**Still deferred (separate passes, flagged):** FItems 8 raw `<div>` modals → DialogControl + Excel import/export (feature); FCommunities toolbar/find raw controls + column persistence; FAssembly/FTakeoff header FlexKit raw-control conversion (the known careful pass). Integrations (PlanSwift/Sage/Pipeline) out of scope.

---

## FlexKit raw-control conversion (2026-06-12)
Converted remaining raw `<input>/<select>/<textarea>` + hand-rolled `<div>` overlay modals → FlexKit:
- **FItems** (8 input + 2 textarea + 8 overlay modals → DialogControl), **FTakeoff** (inputs/selects/textarea + dialog buttons), **FCommunities** (toolbar→ToolbarButtonControl, Find→TextBoxControl; mirrored to PB), **FQuote** (13 input→TextBoxControl, selects→DropDownListControl w/ ValueChanged adapters, textarea→TextAreaControl, 2 overlay modals→DialogControl).
- **FAssembly** already 100% FlexKit (external pass); its grep "overlay" divs are a drag-resize overlay + 2 context-menu click-catchers (NOT modals) — correctly left.
- **FCustomQuote** already clean.
- Render-gotcha respected: no lowercase `style=` on non-TextBox FlexKit controls (all on raw div/span); Width/CssClass/ButtonControl.Style used. Build 0 errors (--no-incremental) both apps. **User must click-verify converted screens** (render throws aren't build-caught).
- **REMAINING backlog (~25+ forms still have raw controls):** FEstimateItems(30), FMassChange(23), FCommunityAccounting(19), FSendingWizard(18), FPurchaseOrder(18), FRfis(17), FItemChart(15), FExportPricelists(12), FContacts(12), FCompany(11), FTBDAssignmentWizard(11), FEditItem(10) + a tail. Cosmetic (forms function); convert per priority.

---

## FlexKit raw-control FULL SWEEP — workflow run (2026-06-13, partial)
Ran `flexkit-sweep` multi-agent workflow over 74 remaining forms (excluded FQuote/FTakeoff already-done + FMain shell). **Hit the Anthropic session limit (reset 2:40am ET) ~10 min in**, which killed ~41 convert agents + the build-verify agent.
- **Converted: ~33 forms, 416 raw controls + 24 hand-rolled modals → FlexKit** (the high-control forms ran first: FMassChange 23→1, FCommunityAccounting 19→1, FRfis 17→2, FItemChart 15→2, FEstimateItems, FSendingWizard, FCompany, FEditItem, FStdAddons, FSplitItems, FCreateInvoice, FrmUser, FVendorChange, FItemPriceUpdateWizard, FAssemblyAttributes, FAddons, FTakeoffOneTime, FPOPriceChangeWiz, FPOMassCancel, FFormulaEditor, FEstimateItemsRefreshCosts, FDuplicateItems, FCreateHBInvoice, FCostForecast, FPriceListUpdate, FLogin, FImportPricelists, FReportRunner, …). Residual 1-2 counts on these = intentional grid-cell/glyph/file-input leftovers.
- **2 forms' agents died mid-edit** (FExportBudgets: missing OnTaskChoiceChanged adapter for RadioControl :after; FPOStatusByJobReport: 3 ValueChanged handlers had nullable int?/TreeNavigationMode? signatures vs the controls' non-nullable delegates) → **manually fixed (6 compile errors)**.
- **Build CLEAN (--no-incremental, 0 errors).** Render-gotcha heuristic: no `style=` on any non-TextBox FlexKit control.
- **REMAINING (~41 forms, blocked by session limit):** FContacts(12), FrmOptionListMultiSelect(5), FPurchaseOrder(partial,8), + the untouched failed-agent forms (FReportParams, FExportAREstimates, FDirectCosts, FCreateContract, FBIMImport, FAddContractItem, FVendor, FPricingWorkSheet, FPOGenerationOptions, FExportEstimates, FEstimateItemsFormatting, FDocuments, FAddon, FTakeoffSettings, FRFPWizard, FPhase, FJobCorrespondence, FDbUpgrade, FDataExport, FChangeTaxes, FAssemblyReplicator, FPriceListLegacy, FPOItems, FDefaultVendors, FList, FCrystalReports, FItemPicklist, FConfirmationMsgBox, DCalendar, FFieldPOs, FAssemblyPickList, FExportPOs, FMail, FItem, FFormulaTest, FFind, FPOStatusByVendorReport, FVariable). **RESUME** after limit reset: `Workflow({scriptPath: flexkit-sweep-wf_365fd126-f15.js, resumeFromRunId: "wf_5f2ec21f-0e9"})` (cached 33 return instantly; re-runs the failed). PB mirror of the whole sweep = held (monthly).
