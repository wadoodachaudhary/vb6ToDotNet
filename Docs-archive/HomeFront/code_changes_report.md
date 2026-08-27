# Detailed HomeFront VB6 Code Difference Report

This report provides a detailed breakdown of all changes made to the `.frm` and `.bas` files in the `HFEst` and `HFSystem` components between the Original VB6 codebase and the New VB6 codebase. For other files, a summary list of size and attribute changes is provided.

## Executive Summary

- **Total FRM/BAS files added:** 6
- **Total FRM/BAS files deleted:** 3
- **Total FRM/BAS files modified:** 47
- **Of the modified files, only ~18 carry real behavioral/SQL changes** — the rest are cosmetic noise (OCX-filename re-casing `vsflex8`→`vsFlex8`/`MSCOMCTL.OCX`, VB identifier casing `Button.Key`→`Button.key`/`flags`→`Flags`, form geometry/`.frx` offsets, dead `Stop:Resume` removals). `MDBUpgrade4/5.bas` and `FDbUpgrade.frm` are **DB schema-upgrade scripts — not migrated to app code** (the apps never run schema upgrades; they reveal the new schema the new features depend on).
- **Sections "Migration Impact Analysis" and "Prioritized Migration Update Plan" below (added 2026-06-08)** map each real VB6 change to the Blazor `F*.razor` that must be updated.

---

## Migration Impact Analysis (added 2026-06-08)

The real changes cluster into **7 themes**. Blazor paths are under `Components/Pages/Migrated/` (HomeFront) and `Components/Pages/` (HomeFrontPB) unless noted. "Status" = state of the Blazor port today.

### Theme 1 — New **D365 / ABN** accounting integration (large, net-new feature)
A new accounting system *"D365 -- ABN"* (`AccountingSystems.asD365_ABN = 11`) with per-division credentials and Microsoft-Dynamics financial-dimension mapping. Backed by `MDBUpgrade5.bas` schema (do **not** migrate the script; ensure the target DB has it applied).

| VB6 change | Blazor target | Status | Action |
|---|---|---|---|
| `Application.cls` `asD365_ABN = 11` enum | `Components/Pages/Migrated/Application.cs` + `Application.razor` (enum ~line 44/47) | enum stops at `asQuickbooksOnline=10` | Add `asD365_ABN = 11` to **both** enum copies (HF + PB) |
| `FOptions.frm` D365-ABN options frame (6 fields `ABN_Resource/TenantID/ClientID/ClientSecret(masked)/LegalEntity/D01Division`) + dropdown entry | `FOptions.razor` | missing | Add D365-ABN accounting branch + 6 option fields (load/save to `appoptions`); add "D365 -- ABN" to the accounting dropdown |
| `FCommunities.frm` `ABN_D02Function` col (hidden unless AccountingSystem=D365_ABN; picks `D365FinancialDimensionValues` Dimension='D02_Function') | `FCommunities.razor` | missing | Add hidden-unless-D365 column to grid + load SELECT + UPDATE + the dimension picklist |
| `FDBGrid.frm` D365 pickers `D03_CostCentre`/`D06_Brand`/`D04_SpendCategory` + `WalletBankAccount` picker; `D365_DimensionConfig.Value` auto-quote; lowercase brace-stripped `WalletPartyID` GUID | `FDBGrid.razor` | missing | Add the 3 dimension pickers + WalletBankAccount picker (divisions→datasources→accountingap join) + GUID/quoting rules |
| `Application.cls` `EditJCCostCodes` SELECT gains `ABN_D04SpendCategory/ABN_ProjectCategory/ABN_ItemCode` + conditional hidden cols | `Application.cs` `EditJCCostCodes` (empty stub) | stub | Implement when that grid is built; `FDBGrid.razor.ShowForm` already has the 9-param signature |
| `FMain.frm` (HFEst) `asD365_ABN` gates (hide PO-Price-Update wizard; Community-Phases grid gains `ABN_D03CostCentre`/`ABN_D06Brand`) | `FMain.razor` | missing | Add `asD365_ABN` gates + the 2 community-phase columns |

New schema (from `MDBUpgrade5.bas`, for reference): table `D365FinancialDimensionValues(LegalEntity,Dimension,Value,Description)`; columns `system_setup.ABN_*` (×6), `standardcostcodes.ABN_D04SpendCategory/ABN_ProjectCategory/ABN_ItemCode`, `tbllocality.ABN_D02Function`, `communityphase.ABN_D03CostCentre/ABN_D06Brand`.

### Theme 2 — **BuildPro retirement** from the desktop client
| VB6 change | Blazor target | Status | Action |
|---|---|---|---|
| `Application.cls` removed 8 `SendBuildPro*` RunTask cases | `Application.cs` (~739-755) | **still present** | Remove the 8 cases (+ orphaned `SendBuildProXxx` methods) |
| `FMain.frm` removed BuildPro right-click submenu + dispatch | `FMain.razor` | check | Remove any BuildPro submenu remnants |
| `MMain.bas` PO cancel: set `POMaster.DateSentToBuildPro=NULL`, `TStmp=GETDATE()`, drop `CancelBuildProPO` call; vendor-change stamps `DateVendorChanged=getdate()` | PO path in `Application.cs` / PO routines | partial | Mirror both SQL changes; verify both columns exist |
| `MBuildPro.bas` `GetBuildProPOs` → `exec dbo.BuildPro_GetPOs @Div,@Job` | `MBuildPro.razor` | low priority | Only if BuildPro PO-fetch is migrated |
| `FOptions.frm` removed `BuildProSendPOsImmediately`; URL → `integration2.hyphensolutions.com`; trigger toggle `POMaster_CreateInvoice` on save | `FOptions.razor` | has old `BuildProSendPOsImmediately` ×4 | Remove option, bump URL, emit trigger-toggle SQL |

### Theme 3 — New **Precon Jobs** thread (FOptions → FJob → FPOIndex → FCommunities)
| VB6 change | Blazor target | Status | Action |
|---|---|---|---|
| `FOptions.frm` new flags `EnablePreconJobs`, `AutoApproveTBDVendorAssignments`, `AllowJobReassignInBudgetsAndPOs`, `BuildProWarrantyDocType` | `FOptions.razor` | missing | Add the 4 option controls (wire `EnablePreconJobs` first — others read it) |
| `FJob.frm` `EnablePreconJobs`-gated rows "Precon Template" (ScheduleTemplates picklist) + "Precon Start Date" → cols `preconScheduleTemplate`/`PreconStartDate`; **removed Holdback/Retainage** (`HoldbackRate`); label renames *Schedule Template→Const Template*, *Construction Start Date→Const Start Date*, *Shell/Unit Schedule Template→Shell/Unit Template*; medium-date format | `FJob.razor` | has old Holdback + old labels | Add Precon rows (gated), remove Retainage UI + `HoldbackRate` from load/UPDATE/INSERT, rename property labels, add the 2 columns to the job SELECT. **NB: the user's VB6 screenshot already shows the new "Const Template"/"Const Start Date" labels** |
| `FPOIndex.frm` checkboxes "Warranty PO"/"Precon PO" → cols `WarrantyPO`/`PreconPO`; rename "BuildPro"→"Send to BuildPro"; require Cost Code+Category when `BuildProCompanyCode<>''` and Send-to-BuildPro checked | `FPOIndex.razor` | missing | Add 2 checkboxes + bit columns + label + validation |
| `FCommunities.frm` `PreconScheduleTemplate` col (combo from `ScheduleTemplates` option); person-pickers prepend blank `'' ID,'' Name` row | `FCommunities.razor` | missing | Add column to grid/SELECT/UPDATE; prepend blank row to the 4 person-picker queries |

### Theme 4 — **Assembly Import/Export** (new screens) + FAddPricelist redesign
| VB6 change | Blazor target | Status | Action |
|---|---|---|---|
| **NEW** `FAssemblyImport.frm` ("Data Import Wizard") + `MAssemblyImport.bas` (=`MAssemblyExport`): SP-based import — stages into `ImportedAssemblies`/`ImportedAssemblyTakeoffs`/`ImportedAssemblyTakeoffItems`, runs `Purch_ImportedAssemblies_Validate/_Commit` + `Purch_ImportedAssemblyTakeoffs_Validate/_Commit`; export builds xlsx from `tbldbassemblymaster` | `FAssemblyImport.razor` | **MISSING** | **Build new page** (use `FBIMImport.razor` CSV-stage pattern as template) + repoint FMain menu cases |
| `FMain.frm` new menu items "Export/Import Assemblies", "Export/Import Assembly Takeoffs" → new subs; legacy `FImportAssemblies` gated behind `UseOLDAssemblyImportForms`; new **"Send NOIs"** (`SendingWizard("NOI")` when `SendNOIsFromPrecisionBuilder`); "Send POs" also on `SendPOsFromPrecisionBuilder`; removed dead `WriteToWebAll/Models` | `FMain.razor` | **partly ahead** (has Export/Takeoff items; still has `WriteToWeb` cases; routes Import to old forms) | Repoint Import-Assemblies/Takeoffs → new `FAssemblyImport`; add "Send NOIs" + broaden "Send POs" gate; remove dead `WriteToWeb` cases |
| `FAddPricelist.frm` **single-assembly → multi-assembly** redesign: new `gAssemblies` grid, multi-select `FPickList`, gItems load = per-assembly `UNION ALL` over `tblPhaseItem`, save reads Assembly/Model per row; `ShowForm` drops the `Assembly` arg | `FAddPricelist.razor` (+ `FVendorPriceList.razor` host) | implements OLD single-assembly | Rework to multi-assembly grid; drop `Assembly`/`DefaultAssembly` arg in `ShowForm` + the `FVendorPriceList` modal call |

### Theme 5 — **Document Management** removal
| VB6 change | Blazor target | Status | Action |
|---|---|---|---|
| `FUserPermissions.frm` removed the Document-Management tab (and all `securitygroupdocumentclasses`/`dms_documentclasses` logic); Data Portal → tab 9 | `FUserPermissions.razor` | **already matches** | **Verify only** — Blazor already has Data Portal, no DMS code |
| `FOptions.frm` DMS dropped `customervisible`; added `UseBPDocManagment` checkbox | `FOptions.razor` | has `customervisible` ×7 | Remove `customervisible` from grid/INSERT/UPDATE; add `UseBPDocManagment` |

### Theme 6 — **CRM (HyphenSys) removal** — confirm-absent only
`MCrmDataService.bas` (HFSystem + HFEst) CRM push gutted/commented; `FInboxJobs`/`FInboxCustomQuote` removed `Sales_SetEstimateIndex`/`Sales_SetCustomOptionQuote`. These COM-based hooks were **never migrated** → just confirm no Blazor equivalent exists (it doesn't). No work.

### Theme 7 — **Standalone real fixes** (small, high value)
| VB6 change | Blazor target | Status | Action | Effort |
|---|---|---|---|---|
| `FAssemblyReplicator.frm` **bug fix** (04/23/2025): qualify `tblDBAssemblyMaster dstM` join by `DivisionID/Community/Assembly/Model/OptionID` to stop component double/tripling on multi-target copies | `FAssemblyReplicator.razor` (~line 435) | **has OLD buggy join** | Add the 5 join predicates | small |
| `FPOPriceChangeWiz.frm` `pomaster.ChangeOrderAudit` double-reprice guard (join+filter); category display `category + ' - ' + description` | `FPOPriceChangeWiz.razor` (~340/655/683) | old | Add the guard + display expr | small |
| `FMassChange.frm` apply-dirty-down the vertical selection range (mass single-cell edit) | `FMassChange.razor` | per-row only | Mark all rows in selection range dirty + fan value down | small–med |
| `FAssemblyAttributes.frm` add `isnull(inactive,0)=0` to `attributelistvalues` query | `FAssemblyAttributes.razor` (~227) | old | Add the filter (one-liner) | trivial |
| `FProgress.frm` blank the label when `Count=0` (was "row 0 of 0") | `FProgress.razor` (~11) | old | One-line `@if (Count==0)` | trivial |
| `FAssembly.frm` skip deleting source detail rows during a **copy** (`If mCopiedAssembly=""`) | `FAssembly.razor` (~2786-2790) | unconditional delete | Wrap delete loop in `if (string.IsNullOrEmpty(mCopiedAssembly))` | small |
| `FAddProperty.frm` runs `ALTER TABLE` on a connection tagged `;App=HFDBUpgradeWiz` (schema-trigger whitelist) | `FAddProperty.razor` (~288-289) | plain `DbWrapperSqlServer` | Route the 2 ALTERs through an `ApplicationName=HFDBUpgradeWiz` connection (needs a wrapper capability) | small |
| `FImportVendors.frm` `.xlsx` accept + scoped dup-key guard | `FImportVendors.razor` | mostly done (xlsx still errors server-side) | Optional: real xlsx parse; dup-key already equivalent | trivial |
| `FExportPOs.frm` Unit-Cost=0 when `linetotal=0`; `MIntacct.bas` tax `Round(...,2)` + friendly errors; `MQuickBooksOnline.bas` job-style parent-id | the accounting-export writers | **stubs** in Blazor | Defer — carry these in when the exporters are actually built | — |

### No action needed (cosmetic / superseded)
`FAttachments`, `FDocuments`, `FVendor`, `FItems`, `CommunityStandards`, `FImportAssemblies`, `FExportBudgets`, `FPOMassCancel`, `FInboxTBDAssignments`, `MCmnDlg`, `DCalendar` — casing/geometry/dead-code only. `FImport.frm`→`FDataImport.frm` is a **rename** (already covered by `FDataImport.razor` — just verify the new `ImportFile(…,Session,…)` overload + `session` column). `FDBGrid.frm`/`FDBGrid1.frm` deleted — covered by `FDBGrid.razor`. `FAttachments - Copy.frm` is a stray backup. `FTakeoff.frm` DivisionID-parameterize is already house-style in Blazor.

---

## Prioritized Migration Update Plan (added 2026-06-08)

Ordered by **value ÷ effort** and dependency. Each item is HF + HomeFrontPB (keep the sync pair byte-identical) and must build both `.sln` to 0 errors.

**Phase 0 — Quick fidelity fixes — ✅ DONE 2026-06-09 (both apps build 0 errors)**
1. ✅ `FAssemblyReplicator.razor` — added the 5 `dstM` join predicates (DivisionID/Community/Assembly/Model/OptionID) — the dated double/triple-component bug. *(HF only — PB lacks this page.)*
2. ✅ `FProgress.razor` — blank label when `Count==0`. *(HF only — PB lacks this page.)*
3. ✅ `FAssemblyAttributes.razor` — `ISNULL(inactive,0)=0` on `attributelistvalues`. *(HF + PB.)*
4. ✅ `FAssembly.razor` — detail-delete loop guarded with `if (string.IsNullOrEmpty(mCopiedAssembly))`. *(HF + PB.)*
5. ✅/↪ `Application.cs` + `Application.razor` + PB `Application.cs` — added `asD365_ABN = 11` to all 3 enum copies. **The `SendBuildPro*` case removal was RE-SCOPED OUT of Phase 0** — investigation found `MBuildPro.razor` still dispatches 4 of them (`RunTask("SendBuildProPOIndexes")` etc.), so they're NOT dead; removing them belongs with the full BuildPro retirement (Theme 2 / Phase 2) alongside MBuildPro + the menu + MMain PO-cancel. `CancelBuildProPO` (no caller) goes with that batch too.
6. ✅ `FMain.razor` — removed the dead `WriteToWebAll/WriteToWebModels` cases (confirmed no dispatcher). *(HF + PB.)*

**Phase 1 — Precon Jobs thread — ✅ DONE 2026-06-09 (both apps build 0 errors)**
7. ✅ `FOptions.razor` — added `EnablePreconJobs`/`AutoApproveTBDVendorAssignments` (stored lowercase "true"/"false" per VB6) + `AllowJobReassignInBudgetsAndPOs` ("True"/"False") + `BuildProWarrantyDocType`; removed `BuildProSendPOsImmediately`. **URL bump SKIPPED** — `BuildProURL`/`BuildProIntegrationURL` aren't in the Blazor app (BuildPro service not migrated). *(HF+PB identical.)*
8. ✅ `FJob.razor` — removed Holdback/Retainage entirely; renamed labels Const Template / Const Start Date / Shell Template / Unit Template; added `EnablePreconJobs`-gated Precon Template + Precon Start Date rows (`preconScheduleTemplate`/`PreconStartDate`) in load SELECT + tree + both save paths; template rows now use the ScheduleTemplates picklist. *(HF+PB identical.)*
9. ✅ `FPOIndex.razor` — Warranty PO / Precon PO FlexKit checkboxes + `WarrantyPO`/`PreconPO` bit columns; "BuildPro"→"Send to BuildPro"; BuildPro cost-code/category validation. *(HF+PB identical.)*
10. ✅ `FCommunities.razor` — `PreconScheduleTemplate` column (AppGridLayout already had it; dropdown from `ScheduleTemplates` option) + blank `'' ID,'' Name` row on the 4 person-pickers. ABN_D02Function deferred to Phase 4. *(HF+PB identical.)*

**Phase 2 — PO lifecycle + DMS cleanup — ✅ DONE 2026-06-09**
11. ✅/↪ PO path — `FPurchaseOrder.razor` cancel now `SET Cancelled=1, TStmp=GETDATE(), DateSentToBuildPro=NULL`. `CancelBuildProPO` was already an empty Blazor stub (the "drop the call" is already effective) and its RunTask case is removed (see #15). **`DateVendorChanged` SKIPPED** — no migrated PO vendor-change routine exists in Blazor yet (flag for when that path is built). *(HF+PB — files differ in baseline; identical logic edit.)*
12. ✅ `FPOPriceChangeWiz.razor` — ChangeOrderAudit double-reprice guard (JOIN pomaster pm + filters) + category display `category + ' - ' + description`. **HF only — PB has no FPOPriceChangeWiz page yet** (flag: bake in when ported).
13. ✅ `FMassChange.razor` — apply-down-selection mass edit (FlexKit `OnTypeAheadCommit` fan-out, FTakeoff/FAddPricelist pattern); also fixed a pre-existing bug (IsDirty never set → saved 0 rows). *(HF+PB — different baselines, identical logic.)*
14. ✅ `FOptions.razor` — DMS: dropped `customervisible` (grid col + load + save + model), added `UseBPDocManagment`. *(HF+PB identical.)*
15. ✅ `FAddProperty.razor` — the 2 `ALTER TABLE` now run on a one-off `SqlConnection` with `ApplicationName="HFDBUpgradeWiz"` (via `_db.GetConnectionString()`); shared `DbWrapperSqlServer` untouched. *(HF+PB identical.)*
16. ✅ `FUserPermissions.razor` — verify-only vs new VB6: already matches (no Document-Management tab / `securitygroupdocumentclasses`; Data Portal is the tab). **Plus (2026-06-09):** converted the page's remaining raw HTML controls to FlexKit — 8 `<button>`→ButtonControl, 2 `<select>`→DropDownListControl, 14 `<input>`→TextBoxControl, checkboxes→CheckBoxControl (logic/SQL/tabs unchanged) — closing the standing FlexKit-compliance gap. *(HF+PB byte-identical.)*

**BuildPro retirement (Theme 2) — ✅ DONE 2026-06-09 (per "make non-navigable, don't delete")**
- Removed the 8 `SendBuildPro*` + `CancelBuildProPO` RunTask cases from `Application.cs` + `Application.razor` (HF) + PB `Application.cs`. The methods remain (no longer dispatched).
- `MBuildPro.razor` is **not deleted** — it has no `@page` and no host instantiates it (already orphaned/non-navigable); its `RunTask("SendBuildPro*")` calls now no-op.

**Phase 3 — New screens / large reworks — ✅ DONE 2026-06-09 (both apps build 0 errors)**
17. ✅ **`FAssemblyImport.razor` built new** — SP-based "Data Import Wizard": uploads xlsx/csv → grid → stages into `ImportedAssemblies` / `ImportedAssemblyTakeoffs` / `ImportedAssemblyTakeoffItems` (GUID session) → runs `Purch_ImportedAssemblies_Validate/_Commit` + `Purch_ImportedAssemblyTakeoffs_Validate/_Commit` (≤25 errors shown, stop-on-error) → cleanup. Export is REAL `.xlsx` (app has ClosedXML). FMain route `/assembly-import` registered + both "Import Assemblies"/"Import Assembly Takeoffs" menu cases repointed to it (legacy "Import models and options" kept). *(HF+PB.)* **Send NOIs — ✅ DONE 2026-06-09:** `FSendingWizard` is now **DocType-aware** (`[Parameter] DocType` + route `/po-sending-wizard/{DocType}`; PO/RFQ/NOI captions + mode-appropriate task radios; RFQ mode now actually applies, fixing the previously-ignored `?mode=RFQ`). FMain adds a "Send NOIs" menu item → `/po-sending-wizard/NOI` (+ `_alwaysFreshRoutes`/`InferParameterName` arm). **BOUNDED:** the actual NOI *document + delivery* live in the external **HFSend.exe** (not in the migration source — no NOI report/table/SQL anywhere), so `SendNOIs()` collects recipients faithfully (shared PO "sendpos" flow) then surfaces an honest "delivery pending HFSend NOI spec" notice rather than fabricating a send. **PB FSendingWizard — ✅ PORTED 2026-06-09:** `FSendingWizard.razor`(+css) copied HF→PB and PB FMain wired (route `/po-sending-wizard`→`FSendingWizard`, `DocType` InferParameterName arm, `_alwaysFreshRoutes`) — merged alongside a parallel external edit to PB FMain that had already set Send-RFQ/NOI dispatch to the `/RFQ`/`/NOI` routes. PB's Send-PO/RFQ/NOI now open the real DocType-aware wizard. (Export-Assemblies menu still uses the pre-existing `OpenExportPickListAsync` in both apps; the new page's `.xlsx` export is reachable from the Import wizard.)
18. ✅ **`FAddPricelist.razor` reworked** single→multi-assembly (`GAssemblyRow` grid + trailing blank row, multi-select FPickList, per-assembly `UNION ALL` over tblPhaseItem carrying AssemblyDescription/Model/OptionID/Assembly/RowKey, save reads Assembly/Model per gItems row, delete-assembly removes its rows); `FVendorPriceList.razor` host dropped the `Assembly`/`DefaultAssembly` arg. *(HF+PB.)*

**Phase 4 — D365/ABN feature — ✅ CORE DONE 2026-06-09**
19. ✅ `FOptions.razor` — D365-ABN accounting frame (6 `ABN_*` fields, ClientSecret masked) + "D365 -- ABN" dropdown entry (index 11); load/save as AppOptions. *(HF+PB.)*
20. ✅ `FCommunities.razor` `ABN_D02Function` (dropdown from `D365FinancialDimensionValues` Dim='D02_Function', hidden unless asD365_ABN); `FDBGrid.razor` D365 dimension pickers (`D03_CostCentre`/`D06_Brand`/`D04_SpendCategory`) + `WalletBankAccount` picker + `D365_DimensionConfig.Value` quote rule + lowercase braceless `WalletPartyID` GUID; `Application.cs` `EditJCCostCodes` implemented (ABN spend-category SELECT + conditional hidden columns + addRecords/deleteRecords). *(HF+PB.)* **asD365 menu gating — ✅ DONE 2026-06-09:** FMain now loads the accounting system (`AppOptions.AccountingSystem`) in `OnInitializedAsync` and **hides the PO Price Update Wizard menu item when accounting = Timberline(1)/Intacct(9)/asD365_ABN(11)** (`@if (_accountingSystem is not (1 or 9 or 11))`). *(HF+PB.)* **Community-Phases D365 columns — ✅ DONE 2026-06-09:** the "Community Phases" FDBGrid SELECT now includes `ABN_D03CostCentre`/`ABN_D06Brand` (aliased to the `D03_CostCentre`/`D06_Brand` picker keys), shown only when accounting = asD365_ABN (hidden otherwise via a `BuildCommunityPhaseHiddenColumnsAsync` gate mirroring the VB6 `IIf`); `ShowForm` gained a `columnAliases` map so the UPDATE writes the real columns. *(HF+PB.)* (`D365FinancialDimensionValues` has 0 rows for those dimensions today → pickers empty until that data lands. **"Cost Codes" alias-save — ✅ FIXED 2026-06-09:** the same alias-save pattern (`ABN_D04SpendCategory D04_SpendCategory`, `ABN_ProjectCategory ProjectCategory`, `ABN_ItemCode ItemCode`) was failing on save ("Invalid column name 'D04_SpendCategory'"); `BuildCostCodesForm` now passes a `ColumnAliases` map (display→`ABN_*`) through `DbGridFormSpec` → `ShowForm`, so the UPDATE writes the real columns. *(HF+PB.)*)

**Schema upgrade (`MDBUpgrade4/5.bas`) is applied SEPARATELY (not by the app) — the Phase 3/4 code references the resulting tables/procs/columns, all verified present in the dev DB.**

**`FPOFormats.razor` is owned by an external process (do not edit/sync from here) per user 2026-06-09.**

**Not migrating:** `MDBUpgrade4/5.bas`, `FDbUpgrade.frm` (DB upgrade scripts); CRM `HyphenSys.SalesWrapper` and `MCmnDlg` Win32 dialogs (COM/native, no Blazor analog); `MMisc` base64 P/Invoke (use `System.Convert`). The accounting-export writers (`MIntacct`/QBO/Sage `linetotal=0`) stay deferred until those exporters are built.

---

## 1. Directory Structure Changes
> [!NOTE]
> In the new codebase, the source files for `HFEst` were consolidated into a `HFEst/Source/` subdirectory, whereas in the original codebase they resided directly in the `HFEst/` root. This report has aligned the paths to perform a content-based diff.

## 2. Other Files Details (Non-FRM/BAS)

| File Path | Status | Old Size | New Size | Old Attributes | New Attributes |
| --- | --- | --- | --- | --- | --- |
| `HFEst/Graphics/16/Thumbs.db` | **Added** | 0 | 6167 | `-` | `read|write` |
| `HFEst/Graphics/24/Thumbs.db` | **Added** | 0 | 7696 | `-` | `read|write` |
| `HFEst/Graphics/32/Thumbs.db` | **Added** | 0 | 208516 | `-` | `read|write` |
| `HFEst/Graphics/Thumbs.db` | **Added** | 0 | 20047 | `-` | `read|write` |
| `HFEst/Graphics/Workflow/Thumbs.db` | **Added** | 0 | 100796 | `-` | `read|write` |
| `HFEst/Graphics/Workflow/hf1.gif` | **Modified** | 985878 | 985878 | `read|write` | `read|write` |
| `HFEst/HFEst.BCE.exe` | **Added** | 0 | 14536704 | `-` | `read|write|exec` |
| `HFEst/HFEst.vbp` | **Added** | 0 | 7353 | `-` | `read|write` |
| `HFEst/HFEst.vbw` | **Added** | 0 | 4770 | `-` | `read|write` |
| `HFEst/HFSend.exe` | **Added** | 0 | 282624 | `-` | `read|write|exec` |
| `HFEst/S3W/CrystalDecisions.CrystalReports.Engine.dll` | **Deleted** | 380928 | 0 | `read|write` | `-` |
| `HFEst/S3W/CrystalDecisions.ReportAppServer.ClientDoc.dll` | **Deleted** | 65536 | 0 | `read|write` | `-` |
| `HFEst/S3W/CrystalDecisions.ReportAppServer.CommLayer.dll` | **Deleted** | 49152 | 0 | `read|write` | `-` |
| `HFEst/S3W/CrystalDecisions.ReportAppServer.CommonControls.dll` | **Deleted** | 135168 | 0 | `read|write` | `-` |
| `HFEst/S3W/CrystalDecisions.ReportAppServer.CommonObjectModel.dll` | **Deleted** | 36864 | 0 | `read|write` | `-` |
| `HFEst/S3W/CrystalDecisions.ReportAppServer.Controllers.dll` | **Deleted** | 176128 | 0 | `read|write` | `-` |
| `HFEst/S3W/CrystalDecisions.ReportAppServer.CubeDefModel.dll` | **Deleted** | 36864 | 0 | `read|write` | `-` |
| `HFEst/S3W/CrystalDecisions.ReportAppServer.DataDefModel.dll` | **Deleted** | 262144 | 0 | `read|write` | `-` |
| `HFEst/S3W/CrystalDecisions.ReportAppServer.DataSetConversion.dll` | **Deleted** | 57344 | 0 | `read|write` | `-` |
| `HFEst/S3W/CrystalDecisions.ReportAppServer.ObjectFactory.dll` | **Deleted** | 5120 | 0 | `read|write` | `-` |
| `HFEst/S3W/CrystalDecisions.ReportAppServer.Prompting.dll` | **Deleted** | 155648 | 0 | `read|write` | `-` |
| `HFEst/S3W/CrystalDecisions.ReportAppServer.ReportDefModel.dll` | **Deleted** | 372736 | 0 | `read|write` | `-` |
| `HFEst/S3W/CrystalDecisions.ReportAppServer.XmlSerialize.dll` | **Deleted** | 15872 | 0 | `read|write` | `-` |
| `HFEst/S3W/CrystalDecisions.ReportSource.dll` | **Deleted** | 86016 | 0 | `read|write` | `-` |
| `HFEst/S3W/CrystalDecisions.Shared.dll` | **Deleted** | 884736 | 0 | `read|write` | `-` |
| `HFEst/S3W/CrystalDecisions.Windows.Forms.dll` | **Deleted** | 552960 | 0 | `read|write` | `-` |
| `HFEst/S3W/FlashControlV71.dll` | **Deleted** | 28672 | 0 | `read|write` | `-` |
| `HFEst/S3W/Sage.AOF.dll` | **Deleted** | 143360 | 0 | `read|write` | `-` |
| `HFEst/S3W/Sage.Data.AOF.dll` | **Deleted** | 69632 | 0 | `read|write` | `-` |
| `HFEst/S3W/Sage.Data.CRE.Connect.dll` | **Deleted** | 67368 | 0 | `read|write` | `-` |
| `HFEst/S3W/Sage.Data.CRE.Interfaces.dll` | **Deleted** | 59688 | 0 | `read|write` | `-` |
| `HFEst/S3W/Sage.Data.CRE.Provider.dll` | **Deleted** | 321832 | 0 | `read|write` | `-` |
| `HFEst/S3W/Sage.Data.CRE.dll` | **Deleted** | 84264 | 0 | `read|write` | `-` |
| `HFEst/S3W/Sage.LS1.Core.dll` | **Deleted** | 330536 | 0 | `read|write` | `-` |
| `HFEst/S3W/Sage.LS1.Messaging.Interfaces.dll` | **Deleted** | 43304 | 0 | `read|write` | `-` |
| `HFEst/S3W/Sage.LS1.Messaging.dll` | **Deleted** | 108840 | 0 | `read|write` | `-` |
| `HFEst/S3W/Sage.STO.TransactionService.dll` | **Deleted** | 71976 | 0 | `read|write` | `-` |
| `HFEst/S3W/Sage.STO.TransactionServiceJCAddin.dll` | **Deleted** | 67880 | 0 | `read|write` | `-` |
| `HFEst/S3W/Sage300Wrapper.exe` | **Deleted** | 7680 | 0 | `read|write` | `-` |
| `HFEst/S3W/ShockwaveFlashObjects.dll` | **Deleted** | 32768 | 0 | `read|write` | `-` |
| `HFEst/S3W/stdole.dll` | **Deleted** | 25464 | 0 | `read|write` | `-` |
| `HFEst/Source/FAddPricelist.frx` | **Modified** | 3387 | 3383 | `read|write` | `read|write` |
| `HFEst/Source/FAssembly.log` | **Deleted** | 18175 | 0 | `read|write` | `-` |
| `HFEst/Source/FAssemblyImport.frx` | **Added** | 0 | 225 | `-` | `read|write` |
| `HFEst/Source/FDBGrid.frx` | **Deleted** | 86914 | 0 | `read|write` | `-` |
| `HFEst/Source/FDBGrid1.frx` | **Deleted** | 89152 | 0 | `read|write` | `-` |
| `HFEst/Source/FDataImport.frx` | **Added** | 0 | 2547 | `-` | `read|write` |
| `HFEst/Source/FDataImport.log` | **Added** | 0 | 137 | `-` | `read|write` |
| `HFEst/Source/FEstimateItems.log` | **Modified** | 41132 | 111 | `read|write` | `read|write` |
| `HFEst/Source/FEstimateItemsFormatting.log` | **Deleted** | 4833 | 0 | `read|write` | `-` |
| `HFEst/Source/FHome2.frx` | **Modified** | 5588435 | 5588435 | `read|write` | `read|write` |
| `HFEst/Source/FHome2.log` | **Deleted** | 4971 | 0 | `read|write` | `-` |
| `HFEst/Source/FHome3.frx` | **Modified** | 5595835 | 5595835 | `read|write` | `read|write` |
| `HFEst/Source/FHome3.log` | **Deleted** | 5575 | 0 | `read|write` | `-` |
| `HFEst/Source/FImport.frx` | **Deleted** | 2547 | 0 | `read|write` | `-` |
| `HFEst/Source/FMain.frx` | **Modified** | 382620 | 382620 | `read|write` | `read|write` |
| `HFEst/Source/FMain.log` | **Modified** | 7346 | 85 | `read|write` | `read|write` |
| `HFEst/Source/FPricingWorksheet.frx` | **Modified** | 9391 | 9399 | `read|write` | `read|write` |
| `HFEst/Source/FWebExport.log` | **Added** | 0 | 116 | `-` | `read|write` |
| `HFEst/Source/HFEst.BCE.exe` | **Deleted** | 14536704 | 0 | `read|write` | `-` |
| `HFEst/Source/HFEst.vbp` | **Deleted** | 7791 | 0 | `read|write` | `-` |
| `HFEst/Source/HFSend.exe` | **Deleted** | 282624 | 0 | `read|write` | `-` |
| `HFEst/Source/MSSCCPRJ.SCC` | **Deleted** | 190 | 0 | `read|write` | `-` |
| `HFEst/Source/Source2.zip` | **Deleted** | 1826072 | 0 | `read|write` | `-` |
| `HFEst/Source/enable snapshots.sql` | **Deleted** | 841 | 0 | `read|write` | `-` |
| `HFEst/Source/error.txt` | **Deleted** | 124447 | 0 | `read|write` | `-` |
| `HFEst/Source/expose cmd line.exe` | **Deleted** | 278528 | 0 | `read|write` | `-` |
| `HFEst/Source/filecopy.AVI` | **Deleted** | 22141 | 0 | `read|write` | `-` |
| `HFEst/System/install.dat` | **Added** | 0 | 0 | `-` | `read|write` |
| `HFEst/enable snapshots.sql` | **Added** | 0 | 841 | `-` | `read|write` |
| `HFEst/error.txt` | **Added** | 0 | 144319 | `-` | `read|write` |
| `HFEst/expose cmd line.exe` | **Added** | 0 | 278528 | `-` | `read|write|exec` |
| `HFEst/filecopy.AVI` | **Added** | 0 | 22141 | `-` | `read|write` |
| `HFSystem/Graphics/Thumbs.db` | **Added** | 0 | 10809 | `-` | `read|write` |
| `HFSystem/HFSystem.vbp` | **Modified** | 6185 | 5925 | `read|write` | `read|write` |
| `HFSystem/HFSystem.vbw` | **Modified** | 2932 | 2899 | `read|write` | `read|write` |
| `HFSystem/MailKit.pdb` | **Modified** | 281916 | 283642 | `read|write` | `read|write` |
| `HFSystem/MimeKit.pdb` | **Modified** | 341536 | 344151 | `read|write` | `read|write` |
| `HFSystem/Source.zip` | **Deleted** | 2766446 | 0 | `read|write` | `-` |
| `HFSystem/Source/Application.cls` | **Modified** | 107735 | 107553 | `read|write` | `read|write` |
| `HFSystem/Source/FAttachments.log` | **Modified** | 181 | 2112 | `read|write` | `read|write` |
| `HFSystem/Source/FCommunities.frx` | **Modified** | 88622 | 88727 | `read|write` | `read|write` |
| `HFSystem/Source/FCommunities.log` | **Deleted** | 181 | 0 | `read|write` | `-` |
| `HFSystem/Source/FCustomer.log` | **Deleted** | 180 | 0 | `read|write` | `-` |
| `HFSystem/Source/FDBGrid.log` | **Deleted** | 181 | 0 | `read|write` | `-` |
| `HFSystem/Source/FDbUpgrade.log` | **Deleted** | 43 | 0 | `read|write` | `-` |
| `HFSystem/Source/FDirectCosts.log` | **Deleted** | 274 | 0 | `read|write` | `-` |
| `HFSystem/Source/FDocuments.log` | **Modified** | 181 | 2146 | `read|write` | `read|write` |
| `HFSystem/Source/FJob.log` | **Deleted** | 272 | 0 | `read|write` | `-` |
| `HFSystem/Source/FJobCorrespondence.log` | **Deleted** | 272 | 0 | `read|write` | `-` |
| `HFSystem/Source/FMail.log` | **Deleted** | 369 | 0 | `read|write` | `-` |
| `HFSystem/Source/FOptions.frx` | **Modified** | 44949 | 47146 | `read|write` | `read|write` |
| `HFSystem/Source/FOptions.log` | **Deleted** | 93 | 0 | `read|write` | `-` |
| `HFSystem/Source/FPOFormats.log` | **Deleted** | 181 | 0 | `read|write` | `-` |
| `HFSystem/Source/FPOItems.log` | **Deleted** | 93 | 0 | `read|write` | `-` |
| `HFSystem/Source/FProjectManagers.log` | **Deleted** | 181 | 0 | `read|write` | `-` |
| `HFSystem/Source/FRfis.log` | **Deleted** | 272 | 0 | `read|write` | `-` |
| `HFSystem/Source/FUserPermissions.frx` | **Modified** | 32243 | 31774 | `read|write` | `read|write` |
| `HFSystem/Source/FVendor.log` | **Deleted** | 181 | 0 | `read|write` | `-` |
| `HFSystem/Workflow/.vs/AutoNotice/FileContentIndex/ce0039f5-18df-40b7-b357-c6d5cf94f877.vsidx` | **Deleted** | 976941 | 0 | `read|write` | `-` |
| `HFSystem/Workflow/.vs/AutoNotice/v17/DocumentLayout.json` | **Deleted** | 1718 | 0 | `read|write` | `-` |
| `HFSystem/Workflow/HtmlEditorControl/images/Thumbs.db` | **Added** | 0 | 37103 | `-` | `read|write` |
| `HFSystem/Workflow/HtmlEditorControl/obj/Debug/DesignTimeResolveAssemblyReferencesInput.cache` | **Modified** | 4947 | 9239 | `read|write` | `read|write` |
| `HFSystem/Workflow/HtmlEditorControl/obj/Debug/HTMLEditorControl.csproj.AssemblyReference.cache` | **Modified** | 4594 | 637 | `read|write` | `read|write` |
| `HFSystem/Workflow/HtmlEditorControl/obj/Debug/HTMLEditorControl.csproj.ResolveComReference.cache` | **Modified** | 248 | 775 | `read|write` | `read|write` |
| `HFSystem/Workflow/HtmlEditorControl/obj/Debug/Interop.SHDocVw.dll` | **Modified** | 156160 | 156160 | `read|write` | `read|write` |
| `HFSystem/Workflow/HtmlEditorControl/obj/Release/HTMLEditorControl.csproj.ResolveComReference.cache` | **Modified** | 772 | 775 | `read|write` | `read|write` |
| `HFSystem/Workflow/HtmlEditorControl/obj/x86/Debug/HTMLEditorControl.csproj.ResolveComReference.cache` | **Modified** | 772 | 775 | `read|write` | `read|write` |
| `HFSystem/Workflow/HtmlEditorControl/obj/x86/Release/HTMLEditorControl.csproj.ResolveComReference.cache` | **Modified** | 772 | 775 | `read|write` | `read|write` |
| `HFSystem/Workflow/Workflowbuilder/Backup/HTMLEditorControl/obj/Debug/HTMLEditorControl.csproj.AssemblyReference.cache` | **Modified** | 2398 | 2414 | `read|write` | `read|write` |
| `HFSystem/Workflow/Workflowbuilder/Backup/HTMLEditorControl/obj/Debug/HTMLEditorControl.csproj.ResolveComReference.cache` | **Modified** | 772 | 775 | `read|write` | `read|write` |
| `HFSystem/Workflow/Workflowbuilder/Backup/HTMLEditorControl/obj/Release/HTMLEditorControl.csproj.ResolveComReference.cache` | **Modified** | 772 | 775 | `read|write` | `read|write` |
| `HFSystem/Workflow/Workflowbuilder/Backup/HTMLEditorControl/obj/x86/Debug/HTMLEditorControl.csproj.AssemblyReference.cache` | **Modified** | 2398 | 2414 | `read|write` | `read|write` |
| `HFSystem/Workflow/Workflowbuilder/Backup/HTMLEditorControl/obj/x86/Debug/HTMLEditorControl.csproj.ResolveComReference.cache` | **Modified** | 772 | 775 | `read|write` | `read|write` |
| `HFSystem/Workflow/Workflowbuilder/Backup/HTMLEditorControl/obj/x86/Release/HTMLEditorControl.csproj.ResolveComReference.cache` | **Modified** | 772 | 775 | `read|write` | `read|write` |
| `HFSystem/Workflow/Workflowbuilder/Backup/obj/Debug/AutoNotice.csproj.AssemblyReference.cache` | **Modified** | 3094 | 3104 | `read|write` | `read|write` |
| `HFSystem/Workflow/Workflowbuilder/Backup/obj/Debug/AutoNotice.csproj.ResolveComReference.cache` | **Modified** | 845 | 849 | `read|write` | `read|write` |
| `HFSystem/Workflow/Workflowbuilder/Backup/obj/Release/AutoNotice.csproj.ResolveComReference.cache` | **Modified** | 845 | 849 | `read|write` | `read|write` |
| `HFSystem/Workflow/Workflowbuilder/Backup/obj/x86/Debug/AutoNotice.csproj.ResolveComReference.cache` | **Modified** | 845 | 849 | `read|write` | `read|write` |
| `HFSystem/Workflow/Workflowbuilder/Backup/obj/x86/Release/AutoNotice.csproj.ResolveComReference.cache` | **Modified** | 845 | 849 | `read|write` | `read|write` |
| `HFSystem/Workflow/Workflowbuilder/Graphics/Thumbs.db` | **Added** | 0 | 31350 | `-` | `read|write` |
| `HFSystem/Workflow/Workflowbuilder/Resources/Thumbs.db` | **Added** | 0 | 11328 | `-` | `read|write` |
| `HFSystem/Workflow/Workflowbuilder/obj/Debug/AutoNotice.csproj.AssemblyReference.cache` | **Modified** | 14193 | 9144 | `read|write` | `read|write` |
| `HFSystem/Workflow/Workflowbuilder/obj/Debug/AutoNotice.csproj.ResolveComReference.cache` | **Modified** | 373 | 849 | `read|write` | `read|write` |
| `HFSystem/Workflow/Workflowbuilder/obj/Debug/DesignTimeResolveAssemblyReferencesInput.cache` | **Modified** | 5278 | 9831 | `read|write` | `read|write` |
| `HFSystem/Workflow/Workflowbuilder/obj/Debug/Interop.ADODB.dll` | **Modified** | 89600 | 89600 | `read|write` | `read|write` |
| `HFSystem/Workflow/Workflowbuilder/obj/Debug/Interop.HFSystem.dll` | **Modified** | 30208 | 40960 | `read|write` | `read|write` |
| `HFSystem/Workflow/Workflowbuilder/obj/Debug/TempPE/dsAutoEvents.Designer.cs.dll` | **Deleted** | 43008 | 0 | `read|write` | `-` |
| `HFSystem/Workflow/Workflowbuilder/obj/Debug/TempPE/dsAutoReports.Designer.cs.dll` | **Deleted** | 33280 | 0 | `read|write` | `-` |
| `HFSystem/Workflow/Workflowbuilder/obj/Debug/TempPE/dsTriggerParams.Designer.cs.dll` | **Deleted** | 16896 | 0 | `read|write` | `-` |
| `HFSystem/Workflow/Workflowbuilder/obj/Debug/TempPE/dsTriggers.Designer.cs.dll` | **Deleted** | 16896 | 0 | `read|write` | `-` |
| `HFSystem/Workflow/Workflowbuilder/obj/Release/AutoNotice.csproj.AssemblyReference.cache` | **Modified** | 974 | 979 | `read|write` | `read|write` |
| `HFSystem/Workflow/Workflowbuilder/obj/Release/AutoNotice.csproj.ResolveComReference.cache` | **Modified** | 845 | 849 | `read|write` | `read|write` |
| `HFSystem/Workflow/Workflowbuilder/obj/x86/Debug/AutoNotice.csproj.AssemblyReference.cache` | **Modified** | 2392 | 2403 | `read|write` | `read|write` |
| `HFSystem/Workflow/Workflowbuilder/obj/x86/Debug/AutoNotice.csproj.ResolveComReference.cache` | **Modified** | 845 | 849 | `read|write` | `read|write` |
| `HFSystem/Workflow/Workflowbuilder/obj/x86/Release/AutoNotice.csproj.AssemblyReference.cache` | **Modified** | 4854 | 4876 | `read|write` | `read|write` |
| `HFSystem/Workflow/Workflowbuilder/obj/x86/Release/AutoNotice.csproj.ResolveComReference.cache` | **Modified** | 845 | 849 | `read|write` | `read|write` |
| `HFSystem/error.txt` | **Modified** | 108605 | 117852 | `read|write` | `read|write` |

## 3. Details of Code Changes in FRM & BAS Files

### `HFEst/Source/CommunityStandards.frm` (Modified)
- **Old Size:** 27147 bytes | **New Size:** 27148 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +21 additions, -21 deletions.
- **Query/SQL updates detected.**
- **Change Hunks count:** 4
#### Diff Preview:
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFEst/Source/CommunityStandards.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFEst/Source/CommunityStandards.frm
@@ -1,6 +1,6 @@
 VERSION 5.00

-Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"

-Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"

+Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"

+Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"

 Begin VB.Form FCommunityStandards 

    Caption         =   "Community Standards"

    ClientHeight    =   3180

@@ -111,16 +111,16 @@
    End

    Begin MSComctlLib.Toolbar Toolbar 

       Align           =   1  'Align Top

-      Height          =   570

+      Height          =   600

       Left            =   0

       Negotiate       =   -1  'True

       TabIndex        =   1

       Top             =   0

       Width           =   8685

       _ExtentX        =   15319

-      _ExtentY        =   1005

+      _ExtentY        =   1058

       ButtonWidth     =   820

-      ButtonHeight    =   953

+      ButtonHeight    =   1005

       AllowCustomize  =   0   'False

       Wrappable       =   0   'False

       Appearance      =   1

@@ -450,7 +450,7 @@
 End Sub

 

 Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)

-    Select Case UCase(Trim(Button.Key))

+    Select Case UCase(Trim(Button.key))

         Case "SAVE"

             Call SaveData(False)

     End Select

@@ -581,28 +581,28 @@
             If .RowOutlineLevel(i) = 1 And .RowData(i) <> "" Then

                 s = ""

                 s = s & "INSERT INTO CommunityStandards(Community,CommunityPhase,StdPhase,StdItem,Phase,Item,Notes)" & vbCrLf

-                s = s & "VALUES(" & DbQuote(str, .TextMatrix(i, .ColIndex("Community"))) & vbCrLf

-                s = s & "      ," & DbQuote(str, .TextMatrix(i, .ColIndex("CommunityPhase"))) & vbCrLf

-                s = s & "      ," & DbQuote(str, .TextMatrix(i, .ColIndex("StdPhase"))) & vbCrLf

-                s = s & "      ," & DbQuote(str, .TextMatrix(i, .ColIndex("StdItem"))) & vbCrLf

-                s = s & "      ," & DbQuote(str, .TextMatrix(i, .ColIndex("Phase"))) & vbCrLf

-                s = s & "      ," & DbQuote(str, .TextMatrix(i, .ColIndex("Item"))) & vbCrLf

-                s = s & "      ," & DbQuote(str, .TextMatrix(i, .ColIndex("Notes"))) & ")"

+                s = s & "VALUES(" & DbQuote(Str, .TextMatrix(i, .ColIndex("Community"))) & vbCrLf

+                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("CommunityPhase"))) & vbCrLf

...
```

---

### `HFEst/Source/DCalendar.frm` (Modified)
- **Old Size:** 4550 bytes | **New Size:** 4522 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +3 additions, -3 deletions.
- **Change Hunks count:** 2
#### Diff Preview:
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFEst/Source/DCalendar.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFEst/Source/DCalendar.frm
@@ -27,7 +27,7 @@
       BackColor       =   -2147483633

       BorderStyle     =   1

       Appearance      =   0

-      StartOfWeek     =   170393601

+      StartOfWeek     =   171245569

       CurrentDate     =   37995

    End

 End

@@ -120,7 +120,7 @@
 End Property

 Private Sub UpdateBuddy()

 On Error Resume Next

-    mBuddy = format(mValue, "medium date") 'HFApp.Options(DateFormat))

+    mBuddy = format(mValue, "medium date")

 End Sub

 

 

```

---

### `HFEst/Source/FAddPricelist.frm` (Modified)
- **Old Size:** 37132 bytes | **New Size:** 47190 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +288 additions, -147 deletions.
- **Added methods/subroutines:** Sub ShowForm, Sub gAssemblies_CellButtonClick, Sub gAssemblies_KeyDown
- **Removed methods/subroutines:** Sub ShowForm, Sub cmdAssembly_Click, Sub txtAssembly_Change, Sub txtAssembly_GotFocus, Sub txtAssembly_KeyDown
- **Query/SQL updates detected.**
- **Change Hunks count:** 29
#### Detailed Diff (Truncated Preview):
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFEst/Source/FAddPricelist.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFEst/Source/FAddPricelist.frm
@@ -2,7 +2,7 @@
 Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"

 Begin VB.Form FAddPricelist 

    Caption         =   "Add Vendor Pricing"

-   ClientHeight    =   7230

+   ClientHeight    =   8745

    ClientLeft      =   1905

    ClientTop       =   1635

    ClientWidth     =   14115

@@ -10,12 +10,12 @@
    Icon            =   "FAddPricelist.frx":0000

    LinkTopic       =   "Form1"

    MaxButton       =   0   'False

-   ScaleHeight     =   7230

+   ScaleHeight     =   8745

    ScaleWidth      =   14115

    Begin HFEst.Slider Slider 

       Height          =   3930

-      Left            =   2580

-      Top             =   2565

+      Left            =   2565

+      Top             =   2910

       Width           =   60

       _ExtentX        =   106

       _ExtentY        =   6932

...
                 s = s & "   ,PartNumber=" & DbQuote(Str, .TextMatrix(i, .ColIndex("SKU"))) & vbCrLf

                 s = s & "WHERE Community=" & DbQuote(Str, community) & vbCrLf

                 s = s & "  AND CommunityPhase=" & DbQuote(Str, CommunityPhase) & vbCrLf

-                s = s & "  AND Assembly=" & DbQuote(Str, mAssembly) & vbCrLf

-                s = s & "  AND Model=" & DbQuote(Str, mModel) & vbCrLf

+                s = s & "  AND Assembly=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Assembly"))) & vbCrLf

+                s = s & "  AND Model=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Model"))) & vbCrLf

                 s = s & "  AND Phase=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Phase"))) & vbCrLf

                 s = s & "  AND Item=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Item"))) & vbCrLf

                 s = s & "  AND Vendor=" & DbQuote(Str, Vendor) & vbCrLf

```

---

### `HFEst/Source/FAssembly.frm` (Modified)
- **Old Size:** 215230 bytes | **New Size:** 215583 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +31 additions, -22 deletions.
- **Query/SQL updates detected.**
- **Change Hunks count:** 14
#### Detailed Diff (Truncated Preview):
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFEst/Source/FAssembly.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFEst/Source/FAssembly.frm
@@ -5,8 +5,8 @@
 Begin VB.Form FAssembly 

    Caption         =   "Model & Option Library"

    ClientHeight    =   9000

-   ClientLeft      =   4905

-   ClientTop       =   5070

+   ClientLeft      =   1890

+   ClientTop       =   2145

    ClientWidth     =   20220

    Icon            =   "FAssembly.frx":0000

    KeyPreview      =   -1  'True

@@ -1738,7 +1738,7 @@
 End Sub

 

 Private Sub LoadCategories()

-    Call LoadComboBox(cboCategory, HFApp.Databases(dbHomefront), "SELECT isnull(description,category),category,0 FROM tblcategories order by 1")

+    Call LoadComboBox(cboCategory, HFApp.Databases(dbHomefront), "SELECT isnull(description,category),category,0 FROM tblcategories where isnull(inactive,0)=0 order by 1")

 End Sub

 Public Function ShowForm(SaveAs As Boolean, Optional community As String, Optional Model As String, Optional OptionID As String, Optional Assembly As String, Optional AssemblyDesc As String, Optional AssemblyType As AssemblyTypes) As Boolean

     

@@ -2388,6 +2388,7 @@
                     

     End With

 End Sub

+

 Private Sub gItems_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
...

                 mDirty = True

                 NewRow = Max(.Row, .RowSel) + 1

@@ -4151,6 +4159,7 @@
             Next

                    

         Case mcITEM_SUBSTITUE

+            If .Row <= 0 Then Exit Sub

             s = ""

             s = s & "SELECT isnull(phase,'')+char(1)+isnull(item,'') phaseitem " & vbCrLf

             s = s & "      ,i.Phase" & vbCrLf

```

---

### `HFEst/Source/FAssemblyAttributes.frm` (Modified)
- **Old Size:** 10342 bytes | **New Size:** 10368 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

#### Minor changes:
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFEst/Source/FAssemblyAttributes.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFEst/Source/FAssemblyAttributes.frm
@@ -272,7 +272,7 @@
     On Error GoTo 0

     

     'cboValue(Index).AutoCompleteListItemsOnly = b

-    Call LoadComboBox(cboValue(Index), HFApp.Databases(dbHomefront), "select value,'',0 from attributelistvalues where listid=" & GetComboBoxListID(cboList(Index)) & " order by sortorder")

+    Call LoadComboBox(cboValue(Index), HFApp.Databases(dbHomefront), "select value,'',0 from attributelistvalues where isnull(inactive,0)=0  and listid=" & GetComboBoxListID(cboList(Index)) & " order by sortorder")

 End Sub

 

 Private Sub cboList_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

```

---

### `HFEst/Source/FAssemblyImport.frm` (Added)
- **Old Size:** 0 bytes | **New Size:** 20092 bytes
- **Old Attributes:** `-` | **New Attributes:** `read|write`

- **New file added** with 0 lines.
#### File Preview (First 30 lines):
```vb
VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form FAssemblyImport 
   Caption         =   "Data Import Wizard"
   ClientHeight    =   3480
   ClientLeft      =   7740
   ClientTop       =   2220
   ClientWidth     =   5610
   ClipControls    =   0   'False
   Icon            =   "FAssemblyImport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   5610
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   1575
      Left            =   1125
      TabIndex        =   0
      Top             =   915
      Width           =   3210
      _cx             =   5662
      _cy             =   2778
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
```

---

### `HFEst/Source/FAssemblyReplicator.frm` (Modified)
- **Old Size:** 29759 bytes | **New Size:** 30221 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

#### Minor changes:
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFEst/Source/FAssemblyReplicator.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFEst/Source/FAssemblyReplicator.frm
@@ -555,7 +555,11 @@
     s = s & "select " & DbQuote(Num, HFApp.DivisionID) & " DivisionID, " & DbQuote(Str, DstCommunity) & " Community, dstM.AssemblyID, src.Assembly, src.OptionID, src.Sequence, src.Phase, src.Item, src.Vendor, src.Quantity, src.Rate, src.Cost, src.Notes, src.UStmp, src.TStmp, src.TakeoffQty, src.OrderQty, src.Model, src.Formula, src.ItemChart, src.Location, src.WBS01, src.WBS02, src.WBS03, src.WBS04, src.WBS05, src.WBS06, src.WBS07, src.WBS08, src.WBS09, src.WBS10, src.WBS11, src.WBS12, src.WBS13, src.WBS14, src.WBS15, src.WBS16, src.WBS17, src.WBS18, src.WBS19, " & vbCrLf

     s = s & "       src.WBS20, src.WBS21, src.WBS22, src.WBS23, src.WBS24, src.WBS25, src.WBS26, src.WBS27, src.WBS28, src.WBS29, src.WBS30, src.WBS31, src.WBS32, src.WBS33, src.WBS34, src.WBS35, src.WBS36, src.WBS37, src.WBS38, src.WBS39, src.WBS40, src.POIndex, src.UseModelCost, src.ConversionFactor, src.Invertable" & vbCrLf

     s = s & "from tblDBAssemblyDetails src" & vbCrLf

-    s = s & "join tblDBAssemblyMaster dstM on src.AssemblyID=dstM.SourceAssemblyID" & vbCrLf

+    

+' brian reported HQ bug where items double and triple when you copy an assembly multiple times to multiple communities or divisions or.. fixed 04/23/2025

+'    s = s & "join tblDBAssemblyMaster dstM on src.AssemblyID=dstM.SourceAssemblyID" & vbCrLf

+    s = s & "join tblDBAssemblyMaster dstM on src.AssemblyID=dstM.SourceAssemblyID and dstm.Divisionid=" & DbQuote(Num, HFApp.DivisionID) & " and dstm.community=" & DbQuote(Str, DstCommunity) & " and dstm.Assembly=src.Assembly and dstm.model=Src.Model and dstm.OptionID=src.OptionID" & vbCrLf

+    
...
```

---

### `HFEst/Source/FDBGrid.frm` (Deleted)
- **Old Size:** 39349 bytes | **New Size:** 0 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `-`

- **File deleted** (previously contained 0 lines).

---

### `HFEst/Source/FDBGrid1.frm` (Deleted)
- **Old Size:** 27400 bytes | **New Size:** 0 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `-`

- **File deleted** (previously contained 0 lines).

---

### `HFEst/Source/FDataImport.frm` (Added)
- **Old Size:** 0 bytes | **New Size:** 21291 bytes
- **Old Attributes:** `-` | **New Attributes:** `read|write`

- **New file added** with 0 lines.
#### File Preview (First 30 lines):
```vb
VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Begin VB.Form FDataImport 
   Caption         =   "Data Import Wizard"
   ClientHeight    =   6165
   ClientLeft      =   1980
   ClientTop       =   855
   ClientWidth     =   7305
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FDataImport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   7305
   Begin VB.PictureBox WizFoot 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   7305
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5580
      Width           =   7305
      Begin VB.CommandButton cmdNav 
```

---

### `HFEst/Source/FEstimateItems.frm` (Modified)
- **Old Size:** 658193 bytes | **New Size:** 664864 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +8005 additions, -39 deletions.
- **Added methods/subroutines:** Function CheckVendorInsurance, Function CleanJob, Function CombineImgs, Function GetComboList, Function GetCorrectingEntryRowNum, Function GetJob, Function GetOriginalEntryRowNum, Function GetReversingEntryRowNum, Function GetRowFromPoint, Function GetVendorCost, Function PrettyName, Function SaveBids, Function SaveContract, Function SaveData, Function SaveItems, Function ValidateJobNumber, Function ViewIndex, Property Get Dirty, Property Get Job, Property Get MultiStateIcon, Property Get ReadOnly, Property Let Dirty, Property Let ReadOnly, Sub BudgetsAreLocked, Sub Clear, Sub ClearGroups, Sub ClearItems, Sub ColorizeItems, Sub DeleteQuote, Sub DeleteRFQ, Sub EnableAssemblies, Sub ExportBidSheet, Sub ExportBidSheets, Sub GroupGrid, Sub ImportBidSheet, Sub LoadAssemblies, Sub LoadContract, Sub LoadCustomDescriptions, Sub LoadEstItems, Sub LoadInvoiceHistory, Sub LoadItems, Sub LoadJob, Sub LoadRFQ, Sub LoadViews, Sub LoadViews_CommunityProjectBudget, Sub LoadViews_CommunityProjectPO, Sub LoadViews_JobBudget, Sub LoadViews_JobPO, Sub LoadViews_PhaseProjectBudget, Sub LoadViews_PhaseProjectPO, Sub LoadWBSDescriptions, Sub LockMeAndMyChildren, Sub PostInvoice, Sub PostInvoicesToQB, Sub PostInvoicesToSimply, Sub RefreshTree, Sub ReloadAssemblyTotals, Sub SaveAs, Sub SaveDBAssembly, Sub SetPOStatuses, Sub ShowCancelledItems, Sub ShowPOStatus, Sub ShowSnapshotsMenu, Sub ShowTip, Sub SplitItems, Sub UpdateBidTotals, Sub WriteException, Sub cboModel_Click, Sub cmdApprovePO_Click, Sub cmdRepostPO_Click, Sub gAssemblies_AfterEdit, Sub gAssemblies_AfterUserResize, Sub gAssemblies_BeforeCollapse, Sub gAssemblies_BeforeEdit, Sub gAssemblies_BeforeMouseDown, Sub gAssemblies_DblClick, Sub gAssemblies_KeyDown, Sub gAssemblies_MouseDown, Sub gAssemblies_MouseMove, Sub gAssemblies_MouseUp, Sub gAssemblies_RowColChange, Sub gContract_RowColChange, Sub gInvoices_MouseDown, Sub gItems_AfterEdit, Sub gItems_AfterSelChange, Sub gItems_BeforeEdit, Sub gItems_BeforeMouseDown, Sub gItems_BeforeMoveColumn, Sub gItems_BeforeUserResize, Sub gItems_CellButtonClick, Sub gItems_KeyDown, Sub gItems_MouseMove, Sub gItems_RowColChange, Sub gItems_SelChange, Sub gItems_ValidateEdit, Sub gProperties_BeforeEdit, Sub gProperties_CellButtonClick, Sub gProperties_ComboCloseUp, Sub gProperties_KeyDown, Sub gProperties_RowColChange, Sub gProperties_ValidateEdit, Sub lblMoveLine_MouseMove, Sub lblMoveLine_MouseUp, Sub mnuDeleteSub_Click, Sub mnuEstimateBidItemsGridSub_Click, Sub mnuEstimateBidsGridSub_Click, Sub mnuEstimateItemsAssemblySub_Click, Sub mnuEstimateItemsCustomerSub_Click, Sub mnuEstimateItemsGridSub_Click, Sub mnuEstimateItemsPOSub_Click, Sub mnuEstimateItemsVendorsSub_Click, Sub mnuEstimateRFPItemsGridSub_Click, Sub mnuEstimateRFPsSub_Click, Sub mnuInvoicesSub_Click, Sub mnuSnapShotsSub_Click, Sub txtDeliveryAddress_Change, Sub txtDeliveryRecipient_Change, Sub txtFOB_Change, Sub txtHFComments_Change, Sub txtHFDescription_Change, Sub txtHFOption_GotFocus, Sub txtJCExtra_Change, Sub txtJCExtra_GotFocus, Sub txtJob_Change, Sub txtJob_GotFocus, Sub txtLotPlan_Change, Sub txtLotPlan_GotFocus, Sub txtLot_Change, Sub txtLot_GotFocus, Sub txtNotes_Change, Sub txtNotes_GotFocus, Sub txtOrderedBy_Change, Sub txtOrderedBy_GotFocus, Sub txtOther_Change, Sub txtOther_GotFocus, Sub txtPODescription_Change, Sub txtPODescription_GotFocus, Sub txtPOIndex_GotFocus, Sub txtPONumber_GotFocus, Sub txtPONumber_KeyDown, Sub txtPrice_Change, Sub txtPrice_Validate, Sub txtQuantity_Change, Sub txtQuantity_GotFocus, Sub txtRFPComments_Change, Sub txtRFPComments_GotFocus, Sub txtRFPDescription_Change, Sub txtRFPDescription_GotFocus, Sub txtRetainagePercent_Change, Sub txtRetainagePercent_GotFocus, Sub txtRetainagePercent_Validate, Sub txtShipVia_Change, Sub txtShipVia_GotFocus, Sub txtStandardText_Change, Sub txtStandardText_GotFocus, Sub txtStandardText_KeyDown, Sub txtStyle_Change, Sub txtStyle_GotFocus, Sub txtTerms_Change, Sub txtTerms_GotFocus, Sub txtTerms_KeyDown, Sub txtVendor_GotFocus
- **Query/SQL updates detected.**
- **Change Hunks count:** 22
#### Detailed Diff (Truncated Preview):
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFEst/Source/FEstimateItems.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFEst/Source/FEstimateItems.frm
@@ -930,9 +930,9 @@
       BorderStyle     =   0  'None

       Caption         =   "frmPurchasing"

       Height          =   8385

-      Left            =   60

+      Left            =   90

       TabIndex        =   65

-      Top             =   2610

+      Top             =   2625

       Width           =   21495

       Begin HFEst.Slider Slider 

          Height          =   5355

@@ -961,6 +961,15 @@
             Top             =   2430

             Visible         =   0   'False

             Width           =   12585

+            Begin VB.CommandButton cmdRepostPO 

+               Caption         =   "Resend"

+               Height          =   285

+               Left            =   7710

+               TabIndex        =   143

+               Top             =   915

+               Visible         =   0   'False

+               Width           =   1050

+            End

...
+    With b

+        .Render Picture1.hDC, ScaleX(.Width, vbHimetric, vbPixels), 0&, ScaleX(.Width, vbHimetric, vbPixels), ScaleY(.Height, vbHimetric, vbPixels), 0&, .Height, .Width, -.Height, ByVal 0&

+    End With

+    Set CombineImgs = Picture1.Image

+End Function

+

+Private Sub gProperties_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)

+    FinishEdit = True

+End Sub

+

```

---

### `HFEst/Source/FExportBudgets.frm` (Modified)
- **Old Size:** 59829 bytes | **New Size:** 59815 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

#### Minor changes:
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFEst/Source/FExportBudgets.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFEst/Source/FExportBudgets.frm
@@ -1170,7 +1170,6 @@
         Case 70:   MsgBox Err.Description & vbCrLf & debugStr, vbInformation

         Case Else: Call errHandler(SRCFILE & "SendBatchToTL")

     

-Stop: Resume

     End Select

     'clear batch if it didnt work.

     Call HFApp.SqlExec("UPDATE EstimateItems SET BudgetPostingBatch=0 WHERE BudgetPostingBatch=" & DbQuote(Num, Batch), dbHomefront)

```

---

### `HFEst/Source/FExportPOs.frm` (Modified)
- **Old Size:** 85051 bytes | **New Size:** 85124 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +23 additions, -19 deletions.
- **Change Hunks count:** 8
#### Detailed Diff (Truncated Preview):
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFEst/Source/FExportPOs.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFEst/Source/FExportPOs.frm
@@ -832,16 +832,20 @@
         s = s & "," & Quote("" & rs("lineCategory"))           'Category

         s = s & "," & Quote("" & rs("lineTaxGroup"))           'tax Group

         s = s & "," & Round(Val("" & rs("lineTax")), 2)        'tax

+        

         s = s & "," & Val("" & rs("linecommittedquantity"))    'Units

         

-        If Val("" & rs("linecommittedunitprice")) > 999999 Then

+        If Val("" & rs("linecommittedunitprice")) > 999999 Or Val("" & rs("linetotal")) = 0 Then

             s = s & ",0"                                           'Unit Cost

         Else

             s = s & "," & Quote("" & rs("linecommittedunitprice")) 'Unit Cost

         End If

         

         s = s & "," & Quote("" & rs("linecommitteduom"), , 6)  'Unit Description

+        

+        

         s = s & "," & Round(Val("" & rs("linetotal")), 2)      'amount

+        

         Print #i, s

         rs.MoveNext

     Wend

@@ -1112,7 +1116,7 @@
     Dim ThisPO As String

     Dim IsUSVersion As Boolean

...

                 qty = Abs(qty)

@@ -1679,7 +1683,7 @@
                 X = X & HFApp.XmlMBAdd(c, 50, "Desc", IIf("" & rs("ItemDesc") = "", "untitled", "" & rs("ItemDesc")))

                 X = X & HFApp.XmlMBAdd(n, 15.3, "CostCodeRef", "" & rs("JCCostCode"))

                 X = X & HFApp.XmlMBAdd(n, 2, "CostTypeRef", "" & rs("JCCategory"))

-                X = X & HFApp.XmlMBAdd(n, 12.2, "Amount", amount)

+                X = X & HFApp.XmlMBAdd(n, 12.2, "Amount", Amount)

                 If isCanadian Then

                     X = X & HFApp.XmlMBAdd(c, 1, "SubjectToGST", "" & rs("GST"))

                     X = X & HFApp.XmlMBAdd(c, 1, "SubjectToPST", "" & rs("PST"))

```

---

### `HFEst/Source/FImport.frm` (Deleted)
- **Old Size:** 19493 bytes | **New Size:** 0 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `-`

- **File deleted** (previously contained 0 lines).

---

### `HFEst/Source/FImportAssemblies.frm` (Modified)
- **Old Size:** 95153 bytes | **New Size:** 95139 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +3 additions, -4 deletions.
- **Change Hunks count:** 2
#### Diff Preview:
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFEst/Source/FImportAssemblies.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFEst/Source/FImportAssemblies.frm
@@ -842,12 +842,12 @@
     

     'load datafile

 stage = "EXCEL"

-    If VBGetOpenFileName(mXlsFileName, , , , , , "Excel files|*.xls*|All files|*.*", , , , , Me.hWnd) Then

+    If VBGetOpenFileName(mXlsFileName, , , , , , "Excel files|*.xls*|All files|*.*", , , , , Me.hwnd) Then

         Set mXlBook = GetObject(mXlsFileName)

 stage = ""

         

         TabStrip.Tabs(2).Selected = True

-        lblFileName = mXlsFileName

+        lblFilename = mXlsFileName

         

         Call LoadAssemblies

         Call LoadPhases

@@ -1651,7 +1651,6 @@
     End Select

     Exit Sub

 eh: Call errHandler(SRCFILE & "Validate Data")

-Stop: Resume

 End Sub

 Private Function CellAddress(Row As Long, Col As Long) As String

     CellAddress = gAssemblies.TextMatrix(0, Col) & gAssemblies.TextMatrix(Row, 0)

```

---

### `HFEst/Source/FInboxCustomQuote.frm` (Modified)
- **Old Size:** 28309 bytes | **New Size:** 28124 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

#### Minor changes:
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFEst/Source/FInboxCustomQuote.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFEst/Source/FInboxCustomQuote.frm
@@ -420,8 +420,6 @@
             s = "error in call to SendSaleRepEmail()"

             Call SendSaleRepEmail(.TextMatrix(i, .ColIndex("Location")), .TextMatrix(i, .ColIndex("Seq")))

         

-            s = "error in call to Sales_SetCustomOptionQuote()"

-            'Call Sales_SetCustomOptionQuote(.TextMatrix(i, .ColIndex("Location")), .ValueMatrix(i, .ColIndex("Seq")))

         

             .RowData(i) = ""

         End If

```

---

### `HFEst/Source/FInboxJobs.frm` (Modified)
- **Old Size:** 25481 bytes | **New Size:** 25401 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

#### Minor changes:
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFEst/Source/FInboxJobs.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFEst/Source/FInboxJobs.frm
@@ -316,9 +316,6 @@
     SaveData = firstJob

 

 

-    'set estimateindex in crm

-    Call Sales_SetEstimateIndex(InboxBatchID)

-

 

 

 Exit Function
...
```

---

### `HFEst/Source/FInboxTBDAssignments.frm` (Modified)
- **Old Size:** 16557 bytes | **New Size:** 16518 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +2 additions, -3 deletions.
- **Change Hunks count:** 2
#### Diff Preview:
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFEst/Source/FInboxTBDAssignments.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFEst/Source/FInboxTBDAssignments.frm
@@ -1,5 +1,5 @@
 VERSION 5.00

-Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"

+Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"

 Begin VB.Form FInboxTBDAssignments 

    Caption         =   "TBD Purchase Orders"

    ClientHeight    =   6135

@@ -247,7 +247,6 @@
         s = s & "and p.ponumber in(" & POs & ")" & vbCrLf

         Call HFApp.SqlExec(s)

         

-        Call SendBuildProPOs("", POs)

         

         For r = .Rows - 1 To 1 Step -1

             If .Cell(flexcpChecked, r, .ColIndex("PONumber")) = flexChecked Then

```

---

### `HFEst/Source/FItems.frm` (Modified)
- **Old Size:** 102198 bytes | **New Size:** 102198 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +14 additions, -14 deletions.
- **Query/SQL updates detected.**
- **Change Hunks count:** 7
#### Diff Preview:
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFEst/Source/FItems.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFEst/Source/FItems.frm
@@ -1,11 +1,11 @@
 VERSION 5.00

 Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"

-Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"

+Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"

 Begin VB.Form FItems 

    Caption         =   "Item Database"

    ClientHeight    =   5280

-   ClientLeft      =   7845

-   ClientTop       =   1530

+   ClientLeft      =   5295

+   ClientTop       =   2280

    ClientWidth     =   12960

    Icon            =   "FItems.frx":0000

    LinkTopic       =   "Form2"

@@ -714,7 +714,7 @@
             

             

     End Select

-EXITSUB: Exit Sub

+ExitSub: Exit Sub

 End Sub

 

 Private Sub gItems_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

@@ -821,14 +821,14 @@
     

     Select Case True

             

-        Case Button.Key = "Save"

+        Case Button.key = "Save"

             Call SaveData(False)

         

-        Case Button.Key = "View"

+        Case Button.key = "View"

             Call Toolbar_ButtonDropDown(Button)

 

 

-        Case Button.Key = "ExcelImport"

+        Case Button.key = "ExcelImport"

             If HFApp.ImportData("tblPhaseItem", "Phase,ItemNumber,CostCategory", "DB Items", "Description,TakeoffUOM,ConversionFactor,OrderUOM", "Phase_code=phase,Item=,Item_number=ItemNumber,phase_code=Phase") Then

                 s = ""

                 s = s & "update tblphaseitem" & vbCrLf

@@ -849,20 +849,20 @@
             

             End If

         

-        Case Button.Key = "ExcelExport"

+        Case Button.key = "ExcelExport"

             Call FDataExport.ExportData("SELECT Phase,ItemNumber,CostCategory,PriceLink,Description,Notes,POIndex,JCCostCode,JCCategory,TaxGroup,TakeoffUOM,OrderUOM,ConversionFactor,Price,WastePercent,RoundDir,Roundto,PartNumber,IsQuote FROM tblPhaseItem where DivisionID = " & HFApp.DivisionID & " ORDER BY Phase,ItemNumber,CostCategory", "Export Items to")

 

             

...
```

---

### `HFEst/Source/FMain.frm` (Modified)
- **Old Size:** 162208 bytes | **New Size:** 161009 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +65 additions, -65 deletions.
- **Query/SQL updates detected.**
- **Change Hunks count:** 14
#### Detailed Diff (Truncated Preview):
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFEst/Source/FMain.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFEst/Source/FMain.frm
@@ -1819,12 +1819,12 @@
             InboxRow = CommandBarAddItem("Inbox", "job")

             CustomRequestsRow = CommandBarAddItem("Custom Requests", "job")

         End If

-        

-        If HFApp.Options.ValueByName("BuildProCompanyCode") <> "" Then

+

+        If HFApp.Options.ValueByName("BuildProCompanyCode") <> "" And HFApp.Options.ValueByName("AutoApproveTBDVendorAssignments") <> "true" Then

             POVendorAssignmentRow = CommandBarAddItem("TBD Assignments", "job")

         End If

-        

-        

+

+

         If HFApp.Options(SalesSystem) = SalesSystems.asBuilder1440 Then

             Call CommandBarAddItem("Retrieve Jobs from Sales Center", "sendreceive")

         End If

@@ -1850,7 +1850,7 @@
                 Call CommandBarAddItem("Post budgets", "Intacct")

             End If

         End If

-        

+

         If HFApp.UserPermission("PostCommitments") Then

             If TL Then

...

                 Dim divID As String

@@ -3230,7 +3230,7 @@
 End Sub

 

 Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)

-    Select Case Button.Key

+    Select Case Button.key

         Case "close"

             CommandPanel.Visible = False

         Case "shrink"

```

---

### `HFEst/Source/FMassChange.frm` (Modified)
- **Old Size:** 57594 bytes | **New Size:** 57899 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +13 additions, -4 deletions.
- **Added methods/subroutines:** Sub gItems_AfterSelChange
- **Query/SQL updates detected.**
- **Change Hunks count:** 4
#### Diff Preview:
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFEst/Source/FMassChange.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFEst/Source/FMassChange.frm
@@ -72,7 +72,7 @@
          SheetBorder     =   -2147483643

          FocusRect       =   1

          HighLight       =   2

-         AllowSelection  =   0   'False

+         AllowSelection  =   -1  'True

          AllowBigSelection=   0   'False

          AllowUserResizing=   1

          SelectionMode   =   0

@@ -105,7 +105,7 @@
          Ellipsis        =   0

          ExplorerBar     =   7

          PicturesOver    =   0   'False

-         FillStyle       =   0

+         FillStyle       =   1

          RightToLeft     =   0   'False

          PictureType     =   0

          TabBehavior     =   0

@@ -1213,6 +1213,10 @@
     End If

 End Sub

 

+Private Sub gItems_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)

+    gItems.ColSel = gItems.Col

+End Sub

+

 Private Sub gItems_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

     With gItems

     Select Case .ColKey(Col)

@@ -1230,7 +1234,12 @@
 End Sub

 

 Private Sub gItems_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

-    gItems.RowData(Row) = "DIRTY"

+    Dim r As Long

+    With gItems

+        For r = Min(.Row, .RowSel) To Max(.Row, .RowSel)

+            gItems.RowData(r) = "DIRTY"

+        Next

+    End With

 End Sub

 

 Private Sub mnuPopupSub_Click(Index As Integer)

```

---

### `HFEst/Source/FPOMassCancel.frm` (Modified)
- **Old Size:** 21581 bytes | **New Size:** 21581 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

#### Minor changes:
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFEst/Source/FPOMassCancel.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFEst/Source/FPOMassCancel.frm
@@ -1,5 +1,5 @@
 VERSION 5.00

-Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"

+Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"

 Begin VB.Form FPOMassCancel 

    Caption         =   "Purchase Order Cancel Wizard"

    ClientHeight    =   10770

```

---

### `HFEst/Source/FPOPriceChangeWiz.frm` (Modified)
- **Old Size:** 46598 bytes | **New Size:** 46825 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +6 additions, -7 deletions.
- **Query/SQL updates detected.**
- **Change Hunks count:** 5
#### Diff Preview:
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFEst/Source/FPOPriceChangeWiz.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFEst/Source/FPOPriceChangeWiz.frm
@@ -781,7 +781,7 @@
 Private Sub LoadCategories()

     Dim s As String

     s = ""

-    s = s & "SELECT isnull(description,category),category,0,isvariance FROM standardcategories " & vbCrLf

+    s = s & "SELECT isnull(category + ' - ' + description,category),category,0,isvariance FROM standardcategories " & vbCrLf

     s = s & "where divisionid=" & HFApp.DivisionID & vbCrLf

     s = s & "order by 4 desc,1" & vbCrLf

     Call LoadComboBox(cboCategory, HFApp.Databases(dbHomefront), s)

@@ -915,6 +915,7 @@
                 s = s & "" & vbCrLf

                 s = s & "where ii.invoice is null" & vbCrLf

                 s = s & "and p.Cancelled=0" & vbCrLf

+                s = s & "and isnull(pm.ChangeOrderAudit,'')=''" & vbCrLf

                 s = s & "and i.DivisionID = " & HFApp.DivisionID & vbCrLf

                 s = s & WhereClause("v")

                 s = s & "and v.povendor=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Vendor"))) & vbCrLf

@@ -964,7 +965,7 @@
                 s = s & ",WBS31, WBS32, WBS33, WBS34, WBS35, WBS36, WBS37, WBS38, WBS39, WBS40)" & vbCrLf

                 s = s & "" & vbCrLf '-------------------------------------------------------------------------------------------------

                 s = s & "select " & vbCrLf

-                s = s & DbQuote(Str, SessionKey) & vbCrLf

+                s = s & " pm.ChangeOrderAudit" & vbCrLf

                 s = s & ",i.EstItemID ChangeOrderOrigEstItemID" & vbCrLf

                 s = s & ",n.NextChangeNumber ChangeOrder" & vbCrLf

                 

@@ -1014,11 +1015,13 @@
                 s = s & "join EstimatedItems v ON(i.EstItemID=v.EstItemID)" & vbCrLf

                 s = s & "join NextPOChangeOrders n on(i.divisionid=n.divisionid and i.ponumber=n.ponumber)" & vbCrLf

                 s = s & "join PurchaseOrders p on i.divisionid=p.divisionid and i.ponumber=p.ponumber" & vbCrLf

+                s = s & "join pomaster pm on i.divisionid=pm.divisionid and i.ponumber=pm.ponumber" & vbCrLf

                 s = s & "left join invoiceitems ii on p.divisionid=ii.divisionid and p.ponumber=ii.commitment" & vbCrLf

                 s = s & "" & vbCrLf

                 s = s & "where ii.invoice is null" & vbCrLf

                 s = s & "and isnull(i.pochangeordernumber,'')=''" & vbCrLf

                 s = s & "and p.Cancelled=0" & vbCrLf

+                s = s & "and isnull(pm.ChangeOrderAudit,'')= " & DbQuote(Str, SessionKey) & vbCrLf

                 s = s & "and i.DivisionID = " & HFApp.DivisionID & vbCrLf

                 s = s & "and isnull(i.RowType,'') not in('Original - Cancelled','Original - Reversal')" & vbCrLf

                 s = s & "and isnull(i.ItemReversed,0)=0" & vbCrLf

@@ -1047,16 +1050,12 @@
             rs.MoveNext

         Wend

         POs = Mid(POs, 2)

-        'send to BP

-        Call SendBuildProPOs("", POs)

     End If

     

     MsgBox "POs have been updated", vbInformation, App.ProductName

     

 Exit Sub

 eh: Call errHandler(SRCFILE & "SaveData")
...
```

---

### `HFEst/Source/FPriceList.frm` (Modified)
- **Old Size:** 60587 bytes | **New Size:** 60577 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +8 additions, -8 deletions.
- **Query/SQL updates detected.**
- **Change Hunks count:** 4
#### Diff Preview:
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFEst/Source/FPriceList.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFEst/Source/FPriceList.frm
@@ -1,11 +1,11 @@
 VERSION 5.00

-Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"

-Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"

+Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"

+Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"

 Begin VB.Form FPriceList 

    Caption         =   "Vendor Item Prices"

    ClientHeight    =   5295

-   ClientLeft      =   2820

-   ClientTop       =   2685

+   ClientLeft      =   2850

+   ClientTop       =   2730

    ClientWidth     =   11100

    Icon            =   "FPriceList.frx":0000

    LinkTopic       =   "Form2"

@@ -652,7 +652,7 @@
     Dim CommunityPhase As String

     Dim Assembly  As String

     

-    Select Case Button.Key

+    Select Case Button.key

         Case "DecreaseDecimals": Call SetDecimalPlaces(mDecimals - 1)

         Case "IncreaseDecimals": Call SetDecimalPlaces(mDecimals + 1)

     

@@ -674,7 +674,7 @@
                 Assembly = Parse(Parse(gAssemblies.RowData(gAssemblies.Row), 2, "Assembly = "), 2, "'")

             End If

             

-            If SaveData(True) Then Call FAddPricelist.ShowForm(Vendor, community, CommunityPhase, Assembly)

+            If SaveData(True) Then Call FAddPricelist.ShowForm(Vendor, community, CommunityPhase)

             

             

         Case "Save"

@@ -726,7 +726,7 @@
     Dim ParentMenu As Long

     Dim i As Long

     

-    Select Case Button.Key

+    Select Case Button.key

         Case "View"

             With FMain.PopMenu

                 ParentMenu = .MenuIndex("mnuPriceListViews")

```

---

### `HFEst/Source/FPricingWorksheet.frm` (Modified)
- **Old Size:** 247362 bytes | **New Size:** 247721 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +5 additions, -3 deletions.
- **Query/SQL updates detected.**
- **Change Hunks count:** 3
#### Diff Preview:
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFEst/Source/FPricingWorksheet.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFEst/Source/FPricingWorksheet.frm
@@ -1342,6 +1342,8 @@
         

         Case KeyCode = vbKeyDelete And Shift = vbCtrlMask And Not ReadOnly

             If mWorksheetView = "Estimating" Then

+                If vbCancel = MsgBox("Removing models and options from a worksheet does not remove them from the sales catalog. To remove them from the sales catalog, mark them as Inactive and then publish the worksheet." & vbCrLf & vbCrLf & "Do you want to remove these from the worksheet?", vbQuestion + vbOKCancel, App.ProductName) Then Exit Sub

+                

                 mDirty = True

                 For r = Max(.Row, .RowSel) To Max(2, Min(.Row, .RowSel)) Step -1

                 If Not gData.RowHidden(r) Then

@@ -1956,7 +1958,7 @@
     Dim bNeedPhase     As Boolean

     

 

-    Select Case Button.Key

+    Select Case Button.key

     

         Case "AssemblyCosts"

             With gData

@@ -1966,7 +1968,7 @@
             End With

         

         Case "Publish"

-            If vbYes = MsgBox("Publishing this worksheet will change the cost and selling" & vbCrLf & "prices in your Profit Builder database." & vbCrLf & "Are you sure this is what you want to do?", vbYesNo + vbExclamation, App.ProductName) Then

+            If vbYes = MsgBox("Publishing this worksheet will change the cost and selling" & vbCrLf & "prices in your sales database." & vbCrLf & "Are you sure this is what you want to do?", vbYesNo + vbExclamation, App.ProductName) Then

                 If SaveData(False) Then

                     Call PublishPricingWorksheet(mWorksheet)

                     If HFApp.Options(LockPostedSalesWorksheets) Then ReadOnly = True

```

---

### `HFEst/Source/FProgress.frm` (Modified)
- **Old Size:** 2955 bytes | **New Size:** 3061 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +11 additions, -8 deletions.
- **Change Hunks count:** 1
#### Diff Preview:
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFEst/Source/FProgress.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFEst/Source/FProgress.frm
@@ -67,24 +67,27 @@
 Public Sub Progress(Optional Title As String, _

                     Optional Caption As String, _

                     Optional Row As Long, _

-                    Optional count As Long, _

+                    Optional Count As Long, _

                     Optional Parent As Form)

 On Error Resume Next

 

-

-    Me.Show , IIf(Parent Is Nothing, FMain, Parent)

+    If Not FProgress.Visible Then Me.Show , IIf(Parent Is Nothing, FMain, Parent)

     

     DoEvents

     

     If Title <> "" Then Me.Caption = Title

     If Caption <> "" Then Label1.Caption = Caption

-    Label2.Caption = "row " & Row & " of " & count

+    If Count = 0 Then

+        Label2.Caption = ""

+    Else

+        Label2.Caption = "row " & Row & " of " & Count

+    End If

     Label1.Refresh

     Label2.Refresh

-    If count = 0 Then

-        ProgressBar.value = count

+    If Count = 0 Then

+        ProgressBar.value = Count

     Else

-        ProgressBar.value = Row / count * 100

+        ProgressBar.value = Row / Count * 100

     End If

     

     

```

---

### `HFEst/Source/FTakeoff.frm` (Modified)
- **Old Size:** 186397 bytes | **New Size:** 186425 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +15 additions, -16 deletions.
- **Added methods/subroutines:** Sub AddMBTakeoff, Sub AddModelOrOption, Sub AddSageEstimate
- **Removed methods/subroutines:** Sub AddMBTakeoff, Sub AddModelOrOption, Sub AddSageEstimate
- **Query/SQL updates detected.**
- **Change Hunks count:** 12
#### Detailed Diff (Truncated Preview):
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFEst/Source/FTakeoff.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFEst/Source/FTakeoff.frm
@@ -1273,14 +1273,14 @@
                     Call .AutoSize(0, .Cols - 1)

                     Select Case mViewIndex

                         Case Assembly_Models, Assembly_ModelOptions, Assembly_GlobalOptions, Assembly_DCOptions

-                            Call AddModelOrOption(.GetNode().Key)

+                            Call AddModelOrOption(.GetNode().key)

                             SetCtrlFocus gHFVariables

                         Case Job_Quote, Job_Jobs, Job_CustomOptions

                             Call AddQuote(Level, "", "", Val(.Cell(flexcpText, .Row, 1)))

                         Case Estimate_Sage100

-                            Call AddMBTakeoff(.GetNode().Key)

+                            Call AddMBTakeoff(.GetNode().key)

                         Case Estimate_SageEstimating

-                            Call AddSageEstimate(.GetNode().Key)

+                            Call AddSageEstimate(.GetNode().key)

                         Case Item_Group, Item_PO, Item_CostCode, Item_Description

                             Call AddItem(Parse(.Cell(flexcpText, .Row, 1), 1, "/"), Parse(.Cell(flexcpText, .Row, 1), 2, "/"))

                     End Select

@@ -1459,8 +1459,8 @@
                     Comments = .TextMatrix(r, .ColIndex("Notes"))

                     Formula = .TextMatrix(r, .ColIndex("RawFormula"))

                     Location = .TextMatrix(r, .ColIndex("Location"))

-                    RoundTo = .TextMatrix(r, .ColIndex("RoundTo"))

-                    RoundDir = .TextMatrix(r, .ColIndex("RoundDir"))

+                    RoundTo = Val(.TextMatrix(r, .ColIndex("RoundTo")))

...

     Dim r As Long

@@ -4089,7 +4088,7 @@
     s = s & "join Secured.Item i on e.EstimateId=i.EstimateId" & vbCrLf

     s = s & "left join Secured.Assembly a on i.AssemblyId=a.AssemblyId" & vbCrLf

     s = s & "where 1=1" & vbCrLf

-    s = s & Key & vbCrLf

+    s = s & key & vbCrLf

     s = s & "group by" & vbCrLf

     s = s & " i.PhaseCode       " & vbCrLf

     s = s & ",i.ItemCode        " & vbCrLf

```

---

### `HFEst/Source/MAssemblyImport.bas` (Added)
- **Old Size:** 0 bytes | **New Size:** 22493 bytes
- **Old Attributes:** `-` | **New Attributes:** `read|write`

- **New file added** with 0 lines.
#### File Preview (First 30 lines):
```vb
Attribute VB_Name = "MAssemblyExport"
Option Explicit
Option Compare Text
Const SRCFILE = "MAssemblyExport::"


Public Sub ExportAssemblies()
On Error GoTo eh

    Dim s As String
    Dim rs As Recordset
    
    Dim i As Long
    Dim Models As String
    Dim AssemblyIDs As String
    Dim AssemblyType As String
    
    'select upto 400 assemblies
    s = ""
    s = s & "Models" & Chr(1) & "select a.AssemblyID, c.area Community,c.description CommunityDesc,a.Model,a.Assembly,a.Description,a.Series,a.Style from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) where a.DivisionID = " & HFApp.DivisionID & " and isnull(a.inactive,0)=0 and a.assemblytype=0 order by 1,2,3,4,5,6,7" & Chr(0)
    s = s & "Model Specific Options" & Chr(1)
        s = s & "select distinct m.Model,m.Description" & vbCrLf
        s = s & "from tbldbassemblymaster o" & vbCrLf
        s = s & "join DistinctModelsByDivision m on o.divisionid=m.divisionid and m.model=o.model" & vbCrLf
        s = s & "where o.assemblytype=2" & vbCrLf
        s = s & "and isnull(o.inactive,0)=0 and o.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf & Chr(0)
    s = s & "Global Options" & Chr(1) & "select a.AssemblyID, ct.Description SubCat,c.area Community,c.description CommunityDesc,a.OptionID [Option],a.Assembly,a.Description from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) left outer join tblcategories ct on(a.category=ct.category) where a.DivisionID = " & HFApp.DivisionID & " and isnull(a.inactive,0)=0 and a.assemblytype=3 order by 2,3,4,5,6,7" & Chr(0)
    If Not FPickList.Choose(HFApp.Databases(dbHomefront), "Assemblies", s, , , , , "assemblyid", True) Then Exit Sub
    
    Select Case FPickList.SelectedView
```

---

### `HFEst/Source/MCmnDlg.bas` (Modified)
- **Old Size:** 46049 bytes | **New Size:** 46049 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +36 additions, -36 deletions.
- **Added methods/subroutines:** Function VBChooseColor, Function VBChooseFont, Function VBGetOpenFileName, Function VBGetSaveFileName, Function VBPageSetupDlg
- **Removed methods/subroutines:** Function VBChooseColor, Function VBChooseFont, Function VBGetOpenFileName, Function VBGetSaveFileName, Function VBPageSetupDlg
- **Query/SQL updates detected.**
- **Change Hunks count:** 25
#### Detailed Diff (Truncated Preview):
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFEst/Source/MCmnDlg.bas
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFEst/Source/MCmnDlg.bas
@@ -33,7 +33,7 @@
     nMaxFileTitle As Long        ' Handled internally

     lpstrInitialDir As String    ' Tied to InitDir

     lpstrTitle As String         ' Tied to DlgTitle

-    flags As Long                ' Tied to Flags

+    Flags As Long                ' Tied to Flags

     nFileOffset As Integer       ' Ignored (exercise for reader)

     nFileExtension As Integer    ' Ignored (exercise for reader)

     lpstrDefExt As String        ' Tied to DefaultExt

@@ -77,7 +77,7 @@
     hInstance As Long

     rgbResult As Long

     lpCustColors As Long

-    flags As Long

+    Flags As Long

     lCustData As Long

     lpfnHook As Long

     lpTemplateName As Long

@@ -106,7 +106,7 @@
     hDC As Long                 ' Printer DC/IC or NULL

     lpLogFont As Long           ' Pointer to LOGFONT

     iPointSize As Long          ' 10 * size in points of font

-    flags As Long               ' Type flags

+    Flags As Long               ' Type flags

     rgbColors As Long           ' Returned text color

     lCustData As Long           ' Data passed to hook function
...

     m_lApiReturn = 0

@@ -1225,7 +1225,7 @@
     psd.rtMinMargin.Left = MinLeftMargin * lUnits

     psd.rtMinMargin.Bottom = MinBottomMargin * lUnits

     psd.rtMinMargin.Right = MinRightMargin * lUnits

-    psd.flags = afFlags

+    psd.Flags = afFlags

     

     ' Show Print dialog

     If PageSetupDlg(psd) Then

```

---

### `HFEst/Source/MCrmDataService.bas` (Modified)
- **Old Size:** 3764 bytes | **New Size:** 213 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +1 additions, -120 deletions.
- **Removed methods/subroutines:** Function Integrated, Function SalesWrapper
- **Query/SQL updates detected.**
- **Change Hunks count:** 1
#### Detailed Diff (Truncated Preview):
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFEst/Source/MCrmDataService.bas
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFEst/Source/MCrmDataService.bas
@@ -1,129 +1,10 @@
 Attribute VB_Name = "MCrmDataService"

 Option Explicit

-Private Loaded As Boolean

-Private CrmApiKey As String

-Private CrmClientID As String

-Private CrmEnvironment As String

-

 

 Public Sub Sales_SetCustomOptionQuote(Location As String, seq As Long)

-    Dim s As String

-    Dim rs As Recordset

-    Dim xml As String

-    

-    If Not Integrated Then Exit Sub

-    

-    s = ""

-    s = s & "select " & vbCrLf

-    s = s & " CRMID CRMID" & vbCrLf

-    s = s & ",Bldr_Declined Declined" & vbCrLf

-    s = s & ",Bldr_Declined_Reason DeclineReason" & vbCrLf

-    s = s & ",Category" & vbCrLf

-    s = s & ",Description" & vbCrLf

-    s = s & ",Comments" & vbCrLf

-    s = s & ",EstimatorNotes" & vbCrLf
...
-End Function

 

-Private Function SalesWrapper(xml As String) As String

-    Dim crm As New HyphenSys.SalesWrapper

-    Dim results As String

-    SalesWrapper = crm.PostXml(CrmClientID, CrmApiKey, CrmEnvironment = "Production", "1", xml)

-End Function

-

-

-

```

---

### `HFEst/Source/MIntacct.bas` (Modified)
- **Old Size:** 28306 bytes | **New Size:** 31089 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +40 additions, -6 deletions.
- **Query/SQL updates detected.**
- **Change Hunks count:** 4
#### Diff Preview:
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFEst/Source/MIntacct.bas
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFEst/Source/MIntacct.bas
@@ -266,14 +266,14 @@
             "" & rs("Location"), "" & rs("job"), IIf(PostExtra, "" & rs("jcextra"), ""), "" & rs("IsNewCostCode") = "1", "" & rs("jccostcode"), "" & rs("IsNewCategory") = "1", "" & rs("jccategory"), "" & rs("Department"), _

             Form1099, Box1099, "" & rs("arcustomer"))

                                   

-        TaxNJCAmount = TaxNJCAmount + Val("" & rs("njctax"))

+        TaxNJCAmount = TaxNJCAmount + Round(Val("" & rs("njctax")), 2)

         TaxLocation = "" & rs("Location")

         TaxLineID = "" & rs("OrigTaxLineID")

         

         rs.MoveNext

     Wend

     If ThisPO <> "" Then

-        If TaxNJCAmount <> 0 Then Call Intacct.AddCOTranItem(ThisPO, TaxLineID, GstItemID, "tax (non-costed)", 1, GstItemUOM, TaxNJCAmount, TaxLocation, "", "", False, "", False, "", "", Form1099, Box1099, "" & rs("arcustomer"))

+        If TaxNJCAmount <> 0 Then Call Intacct.AddCOTranItem(ThisPO, TaxLineID, GstItemID, "tax (non-costed)", 1, GstItemUOM, TaxNJCAmount, TaxLocation, "", "", False, "", False, "", "", Form1099, Box1099, "")

         Call Intacct.CloseCOTran

         Call Intacct.CloseMessage

         req = Intacct.xml()

@@ -298,7 +298,24 @@
     

 Exit Sub

 eh:

-    Call errHandler(SRCFILE & "SendCOsToIntacct", FileName)

+    Dim ejob As String

+    Dim ecostcode As String

+    Dim ecategory As String

+    Select Case True

+    Case Err.Description Like "*The task you selected is not associated with the project you selected. When you select a task dimension, that task must be associated with the project dimension that you selected at the line level. The task you selected is *. Go back and select another task.*"

+        ecostcode = Parse(Parse(Err.Description, 3, "The task you selected is "), 1, ". Go back and select another task.")

+        MsgBox "Intacct rejected this PO because cost code " & ecostcode & " has not been added to the job.", vbCritical, App.ProductName & " (version " & App.Major & "." & format(App.Minor, "00") & "." & format(App.Revision, "0000") & ")"

+

+    Case Err.Description Like "*Cost type 'LAB--D60.10.30--ML390' specified is not valid.*"

+        ejob = Parse(Parse(Err.Description, 3, "--"), 1, "'")

+        ecostcode = Parse(Parse(Err.Description, 2, "--"), 1, "--")

+        ecategory = Parse(Parse(Err.Description, 1, "--"), 2, "'")

+        MsgBox "Intacct rejected this PO because cost type " & ecategory & " has not been added to cost code " & ecostcode & " for job " & ejob & ".", vbCritical, App.ProductName & " (version " & App.Major & "." & format(App.Minor, "00") & "." & format(App.Revision, "0000") & ")"

+    

+    Case Else

+        Call errHandler(SRCFILE & "SendCOsToIntacct", FileName)

+        

+    End Select

     Resume Next 'this is required. do not remove

 End Sub

 

@@ -403,7 +420,7 @@
                                   "" & rs("GL_Prefix"), "" & rs("linejob"), IIf(PostExtra, "" & rs("LineExtra"), ""), "" & rs("IsNewCostCode") = "1", "" & rs("linecostcode"), "" & rs("IsNewCategory") = "1", "" & rs("linecategory"), "" & rs("IntacctDepartment"), _

                                   Form1099, Box1099, "" & rs("arcustomer"))

                                   

-        TaxNJCAmount = TaxNJCAmount + Val("" & rs("linenjctax"))

+        TaxNJCAmount = TaxNJCAmount + Round(Val("" & rs("linenjctax")), 2)

         TaxLocation = "" & rs("GL_Prefix")

         
...
```

---

### `HFEst/Source/MMain.bas` (Modified)
- **Old Size:** 187594 bytes | **New Size:** 187875 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +87 additions, -62 deletions.
- **Added methods/subroutines:** Function DeQuote, Function ImageIndex, Function RegGetKey, Function RegSaveKey, Function SimpleEncrypt, Function vbSingleQuote, Sub SetComboBoxListIndex
- **Removed methods/subroutines:** Function ImageIndex, Function RegGetKey, Function RegSaveKey, Function SimpleEncrypt, Sub SetComboBoxListIndex
- **Query/SQL updates detected.**
- **Change Hunks count:** 29
#### Detailed Diff (Truncated Preview):
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFEst/Source/MMain.bas
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFEst/Source/MMain.bas
@@ -708,8 +708,8 @@
                 s = s & "UPDATE tblModels" & vbCrLf

                 s = s & "SET LastSalesWorksheet=SalesWorksheet" & vbCrLf

                 s = s & "   ,SalesWorksheet=" & DbQuote(Num, Worksheet) & vbCrLf

-                s = s & "   ,Assembly=" & DbQuote(Str, "" & rs("Assembly"), , , 20) & vbCrLf

-                s = s & "   ,Description=" & DbQuote(Str, "" & rs("Description"), , , 50) & vbCrLf

+                s = s & "   ,Assembly=" & DbQuote(Str, "" & rs("Assembly")) & vbCrLf

+                s = s & "   ,Description=" & DbQuote(Str, "" & rs("Description")) & vbCrLf

                 s = s & "   ,model_picture=" & DbQuote(Str, "" & rs("GraphicPath")) & vbCrLf

                 s = s & "   ,Spec_Document=" & DbQuote(Str, "" & rs("SpecDocument")) & vbCrLf

                 s = s & "   ,Max_Width=" & DbQuote(Num, "" & rs("MaxWidth")) & vbCrLf

@@ -717,8 +717,8 @@
                 s = s & "   ,Inactive=" & DbQuote(Bit, "" & rs("Inactive")) & vbCrLf

                 s = s & "   ,Comments=" & DbQuote(Str, "" & rs("Comments")) & vbCrLf

                 s = s & "   ,Style=" & DbQuote(Str, "" & rs("Style")) & vbCrLf

-                s = s & "   ,NoOfBedrooms=" & DbQuote(Str, "" & rs("Bedrooms"), , , 30) & vbCrLf

-                s = s & "   ,NoOfBathrooms=" & DbQuote(Str, "" & rs("Bathrooms"), , , 30) & vbCrLf

+                s = s & "   ,NoOfBedrooms=" & DbQuote(Str, "" & rs("Bedrooms")) & vbCrLf

+                s = s & "   ,NoOfBathrooms=" & DbQuote(Str, "" & rs("Bathrooms")) & vbCrLf

                 s = s & "   ,ModelSize=" & DbQuote(Num, "" & rs("FloorArea")) & vbCrLf

                 s = s & "   ,IncentiveCost=" & DbQuote(Cur, "" & rs("IncentiveCost")) & vbCrLf

                 s = s & "   ,IncentiveRetail=" & DbQuote(Cur, "" & rs("IncentiveRetail")) & vbCrLf

@@ -737,9 +737,9 @@
                 End If

                 s = s & "WHERE ISNULL(Area,'')=" & DbQuote(Str, "" & rs("Community")) & vbCrLf

                 s = s & "  AND ISNULL(CommunityPhase,'')=" & DbQuote(Str, "" & rs("CommunityPhase")) & vbCrLf
...
+        End If

+    End If

+    DeQuote = s

+

+

+End Function

+

 

 Public Function Quote(s As String, Optional RemoveFormatting As Boolean = True, Optional Length As Long, Optional RemoveSpecial As Boolean = True) As String

     

```

---

### `HFSystem/Source/DCalendar.frm` (Modified)
- **Old Size:** 5753 bytes | **New Size:** 5742 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +3 additions, -3 deletions.
- **Change Hunks count:** 2
#### Diff Preview:
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFSystem/Source/DCalendar.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFSystem/Source/DCalendar.frm
@@ -57,7 +57,7 @@
       BackColor       =   -2147483633

       BorderStyle     =   1

       Appearance      =   0

-      StartOfWeek     =   55640065

+      StartOfWeek     =   171245569

       CurrentDate     =   37995

    End

 End

@@ -153,7 +153,7 @@
 End Property

 Private Sub UpdateBuddy()

 On Error Resume Next

-    mBuddy = Format(mValue, HFApp.Options(DateFormat))

+    mBuddy = Format(mValue, "medium date")

 End Sub

 

 

```

---

### `HFSystem/Source/FAddProperty.frm` (Modified)
- **Old Size:** 11375 bytes | **New Size:** 11582 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +16 additions, -9 deletions.
- **Query/SQL updates detected.**
- **Change Hunks count:** 4
#### Diff Preview:
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFSystem/Source/FAddProperty.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFSystem/Source/FAddProperty.frm
@@ -209,9 +209,10 @@
 On Error GoTo eh

     Dim s As String

     Dim d As String

+    Dim c As Connection

     

     Dim i As Long

-    Dim size As Long

+    Dim SIZE As Long

     Dim list As String

     

     Select Case True

@@ -219,9 +220,9 @@
             

             If optType(5) Then

                 list = ""

-                size = Val(txtLength.Text)

+                SIZE = Val(txtLength.Text)

                 For i = 1 To Parse(txtPickList.Text, , vbCrLf)

-                    s = left(Trim(Parse(txtPickList.Text, i, vbCrLf)), size)

+                    s = left(Trim(Parse(txtPickList.Text, i, vbCrLf)), SIZE)

                     If s <> "" Then

                         list = list & "|" & s

                     End If

@@ -250,11 +251,11 @@
             

             If optType(5) Then

                 list = ""

-                size = -1

+                SIZE = -1

                 For i = 1 To Parse(txtPickList.Text, , vbCrLf)

                     s = Trim(Parse(txtPickList.Text, i, vbCrLf))

                     If s <> "" Then

-                        size = Max(size, Len(s))

+                        SIZE = Max(SIZE, Len(s))

                         list = list & "|" & s

                     End If

                 Next

@@ -266,10 +267,16 @@
             If optType(2) Then s = " datetime"

             If optType(3) Then s = " bit"

             If optType(4) Then s = " varchar(" & Val(txtLength.Text) & ")"

-            If optType(5) Then s = " varchar(" & IIf(size = -1, 25, size) & ")"

+            If optType(5) Then s = " varchar(" & IIf(SIZE = -1, 25, SIZE) & ")"

             If optType(6) Then s = " money"

-            Call HFApp.SqlExec("ALTER TABLE dbo.JobCustomFields ADD " & vbQuote & Trim(txtFieldName.Text) & vbQuote & s)

-            Call HFApp.SqlExec("ALTER TABLE dbo.WorkticketCustomFlds ADD " & vbQuote & Trim(txtFieldName.Text) & vbQuote & s)

+            

+            Set c = New Connection

+            c.Open HFApp.ConnectionString(dbHomefront) & ";App=HFDBUpgradeWiz"

+            Call c.Execute("ALTER TABLE dbo.JobCustomFields ADD " & vbQuote & Trim(txtFieldName.Text) & vbQuote & s)

...
```

---

### `HFSystem/Source/FAttachments - Copy.frm` (Added)
- **Old Size:** 0 bytes | **New Size:** 53300 bytes
- **Old Attributes:** `-` | **New Attributes:** `read|write`

- **New file added** with 0 lines.
#### File Preview (First 30 lines):
```vb
VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FAttachments 
   Caption         =   "Attachments"
   ClientHeight    =   5370
   ClientLeft      =   11520
   ClientTop       =   3450
   ClientWidth     =   7275
   FillColor       =   &H00FF0000&
   Icon            =   "FAttachments.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   7275
   Begin VB.PictureBox picEmbedded 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1380
      Picture         =   "FAttachments.frx":058A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   4590
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picFolder 
```

---

### `HFSystem/Source/FAttachments.frm` (Modified)
- **Old Size:** 55318 bytes | **New Size:** 52065 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +6 additions, -104 deletions.
- **Added methods/subroutines:** Sub gData_MouseDown, Sub gData_OLEDragDrop, Sub gData_OLEDragOver
- **Removed methods/subroutines:** Sub AquireImage, Sub gData_MouseDown, Sub gData_OLEDragDrop, Sub gData_OLEDragOver
- **Query/SQL updates detected.**
- **Change Hunks count:** 13
#### Detailed Diff (Truncated Preview):
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFSystem/Source/FAttachments.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFSystem/Source/FAttachments.frm
@@ -1,7 +1,6 @@
 VERSION 5.00

 Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"

-Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"

-Object = "{C46A8909-8107-11D8-8671-00C1261173F0}#2.3#0"; "TwainControlX.ocx"

+Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"

 Begin VB.Form FAttachments 

    Caption         =   "Attachments"

    ClientHeight    =   5370

@@ -14,48 +13,6 @@
    LinkTopic       =   "Form1"

    ScaleHeight     =   5370

    ScaleWidth      =   7275

-   Begin TwainControlX.Twain Twain1 

-      Height          =   480

-      Left            =   1800

-      TabIndex        =   4

-      Top             =   4410

-      Visible         =   0   'False

-      Width           =   480

-      CurrentDevice   =   -1

-      UseInterface    =   -1  'True

-      WaitForAcquire  =   -1  'True

-      DoubleBuffered  =   0   'False

-      Enabled         =   -1  'True

...
-        Call FAttachments.AddFile(gData.TextMatrix(gData.Row, gData.ColIndex("ObjectID")), gData.TextMatrix(gData.Row, gData.ColIndex("Folder")), True, s)

-        Kill s

-    End If

-    

-Exit Sub

-eh: Call errHandler(SRCFILE & "AquireImage")

-End Sub

 

 Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

     Select Case True

```

---

### `HFSystem/Source/FCommunities.frm` (Modified)
- **Old Size:** 37319 bytes | **New Size:** 38656 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +92 additions, -67 deletions.
- **Query/SQL updates detected.**
- **Change Hunks count:** 12
#### Detailed Diff (Truncated Preview):
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFSystem/Source/FCommunities.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFSystem/Source/FCommunities.frm
@@ -1,11 +1,11 @@
 VERSION 5.00

 Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"

-Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"

+Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"

 Begin VB.Form FCommunities 

    Caption         =   "Communities"

    ClientHeight    =   5520

-   ClientLeft      =   870

-   ClientTop       =   2205

+   ClientLeft      =   1290

+   ClientTop       =   2625

    ClientWidth     =   12750

    Icon            =   "FCommunities.frx":0000

    KeyPreview      =   -1  'True

@@ -56,7 +56,7 @@
       GridLinesFixed  =   2

       GridLineWidth   =   1

       Rows            =   4

-      Cols            =   45

+      Cols            =   47

       FixedRows       =   1

       FixedCols       =   0

       RowHeightMin    =   0

@@ -155,159 +155,159 @@
          BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
...
                 

-                

         End Select

         

         If s <> "" Then

-            If FPickList.Choose(HFApp.Databases(dbHomefront), .TextMatrix(0, Col), s, .TextMatrix(Row, Col), , False, , "id") Then

+            If FPickList.Choose(HFApp.Databases(dbHomefront), .TextMatrix(0, Col), s, .TextMatrix(Row, Col), , False, , hidecols) Then

                 .TextMatrix(Row, Col) = FPickList.SelectedItem("id")

             End If

         End If

```

---

### `HFSystem/Source/FDBGrid.frm` (Modified)
- **Old Size:** 41049 bytes | **New Size:** 43852 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +70 additions, -14 deletions.
- **Query/SQL updates detected.**
- **Change Hunks count:** 11
#### Detailed Diff (Truncated Preview):
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFSystem/Source/FDBGrid.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFSystem/Source/FDBGrid.frm
@@ -456,17 +456,21 @@
 Private Sub gData_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

 On Error Resume Next

     With gData

-        If .ColKey(.MouseCol) = "Sage300GLPrefixLength" Then

-            .ToolTipText = "Number of digits of the gl account that make up the company number"

+        If .MouseCol = -1 Then

+            .ToolTipText = ""

         Else

-            .ToolTipText = ""

+            If .ColKey(.MouseCol) = "Sage300GLPrefixLength" Then

+                .ToolTipText = "Number of digits of the gl account that make up the company number"

+            Else

+                .ToolTipText = ""

+            End If

         End If

     End With

 End Sub

 

 Private Sub gData_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

     Dim r As Long

-    

+    Dim s As String

     With gData

     
...

             'preserve key column values incase the user changes them.

@@ -950,7 +1006,7 @@
         End If

         Next

         s = Mid(ValuesClause, 1, Len(ValuesClause) - 1) & ")"

-        Call HFApp.SqlExec(s, dbHomeFront)

+        Call HFApp.SqlExec(s, dbHomefront)

         .RowData(r) = ""

         

         'get new identity value

```

---

### `HFSystem/Source/FDbUpgrade.frm` (Modified)
- **Old Size:** 19659 bytes | **New Size:** 25898 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +163 additions, -44 deletions.
- **Added methods/subroutines:** Sub RUNPROC_FixSchema, Sub RemoveZybDefaults, Sub sql
- **Removed methods/subroutines:** Sub DisableAudit, Sub EnableAudit, Sub Form_Unload, Sub sql
- **Query/SQL updates detected.**
- **Change Hunks count:** 11
#### Detailed Diff (Truncated Preview):
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFSystem/Source/FDbUpgrade.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFSystem/Source/FDbUpgrade.frm
@@ -282,6 +282,7 @@
     Description  As String  'displayed in listbox

     SQLStatement As String  'query to run

     IgnoreErrors As Boolean

+    DontLog      As Boolean 'only used for schema protection statements

 End Type

 Private mStatements() As StatementType

 

@@ -313,8 +314,8 @@
     'open connection

     Set mConnection = New Connection

     Set mConnection2 = New Connection

-    mConnection.Open ConnectionString

-    mConnection2.Open ConnectionString

+    mConnection.Open ConnectionString & ";App=HFDBUpgradeWiz"

+    mConnection2.Open ConnectionString & ";App=HFDBUpgradeWiz"

     mConnection.CommandTimeout = 3000

     mConnection2.CommandTimeout = 3000

     

@@ -339,9 +340,9 @@
     If mDbRevision > UBound(mStatements) Then

     

         s = ""

-        s = s & "A new version has been installed on the server. You should" & vbCrLf

-        s = s & "upgrade your workstation software as soon as possible. If you" & vbCrLf

-        s = s & "choose not to upgrade your software you may experience" & vbCrLf
...
-    mConnection.Execute "disable trigger audit_view_change on database"

-    mConnection.Execute "disable trigger audit_table_change on database"

-    mConnection.Execute "disable trigger audit_trigger_change on database"

-

-End Sub

-

-

 

 

 Private Sub RUNPROC_EncryptPswds()

```

---

### `HFSystem/Source/FDocuments.frm` (Modified)
- **Old Size:** 35962 bytes | **New Size:** 33665 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +8 additions, -70 deletions.
- **Query/SQL updates detected.**
- **Change Hunks count:** 13
#### Detailed Diff (Truncated Preview):
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFSystem/Source/FDocuments.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFSystem/Source/FDocuments.frm
@@ -1,7 +1,6 @@
 VERSION 5.00

-Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"

-Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"

-Object = "{C46A8909-8107-11D8-8671-00C1261173F0}#2.3#0"; "TwainControlX.ocx"

+Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"

+Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"

 Begin VB.Form FDocuments 

    Caption         =   "Documents"

    ClientHeight    =   5460

@@ -22,7 +21,7 @@
       Left            =   495

       ScaleHeight     =   240

       ScaleWidth      =   240

-      TabIndex        =   8

+      TabIndex        =   6

       Top             =   4650

       Visible         =   0   'False

       Width           =   240

@@ -36,7 +35,7 @@
       Picture         =   "FDocuments.frx":058A

       ScaleHeight     =   240

       ScaleWidth      =   240

-      TabIndex        =   7

+      TabIndex        =   5

       Top             =   4635
...

 On Error Resume Next: Call Kill(PathAppend(TempPath, "*.*"))

 End Sub

@@ -752,7 +691,6 @@
     With gData

     Select Case Index

         Case CLASS_ADDFILE:   Call AddDocument(gData.Row, "")

-        Case CLASS_SCANFILE

     End Select

     End With

 End Sub

```

---

### `HFSystem/Source/FImport.frm` (Modified)
- **Old Size:** 20982 bytes | **New Size:** 20982 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +7 additions, -7 deletions.
- **Added methods/subroutines:** Function SaveToCSV, Sub ReadText
- **Removed methods/subroutines:** Function SaveToCSV, Sub ReadText
- **Query/SQL updates detected.**
- **Change Hunks count:** 3
#### Diff Preview:
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFSystem/Source/FImport.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFSystem/Source/FImport.frm
@@ -376,11 +376,11 @@
     End If

 End Property

 

-Private Function SaveToCSV(FileName As String) As String

+Private Function SaveToCSV(Filename As String) As String

 On Error Resume Next

     Dim s As String

     Dim xlSheet As Object 'Excel.Worksheet

-    Set xlSheet = GetObject(FileName).Sheets(1)

+    Set xlSheet = GetObject(Filename).Sheets(1)

     s = TempFile("csv")

     Kill s

     Call xlSheet.SaveAs(s, 6, , , , , False)

@@ -388,7 +388,7 @@
     SaveToCSV = s

 End Function

 

-Private Sub ReadText(FileName As String)

+Private Sub ReadText(Filename As String)

 On Error GoTo eh

     Dim i As Long

     Dim c As Long

@@ -401,9 +401,9 @@
     Dim missingkey As Boolean

     

     

-    Select Case FileExt(FileName)

-        Case "csv": Call gData.LoadGrid(FileName, flexFileCommaText)

-        Case Else:  Call gData.LoadGrid(FileName, flexFileTabText)

+    Select Case FileExt(Filename)

+        Case "csv": Call gData.LoadGrid(Filename, flexFileCommaText)

+        Case Else:  Call gData.LoadGrid(Filename, flexFileTabText)

     End Select

     

     

```

---

### `HFSystem/Source/FImportVendors.frm` (Modified)
- **Old Size:** 20106 bytes | **New Size:** 20263 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +9 additions, -5 deletions.
- **Query/SQL updates detected.**
- **Change Hunks count:** 5
#### Diff Preview:
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFSystem/Source/FImportVendors.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFSystem/Source/FImportVendors.frm
@@ -1,5 +1,5 @@
 VERSION 5.00

-Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"

+Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"

 Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"

 Begin VB.Form FImportVendors 

    Caption         =   "Pricelist Import Wizard"

@@ -317,7 +317,7 @@
                 lblFileDate.Caption = "Last Modified: " & Format(fileinfo.ModifyTime, "long date")

                 

                 Select Case FileExt(mFilename)

-                    Case "xls"

+                    Case "xls", "xlsx"

                         mFilename = SaveToCSV(mFilename)

                         Call ReadText(mFilename)

                         On Error Resume Next

@@ -468,7 +468,7 @@
 Private Function SaveData() As Boolean

 On Error GoTo eh

     Dim tabdef As Recordset

-    

+    Dim bIgnoreDupKey As Boolean

     Dim r As Long

     Dim c As Long

     Dim X As Long

@@ -490,7 +490,9 @@
         

         

         s = "INSERT INTO tblVendors(DivisionID,webupdated,nodatachanged,Vendor_ID) VALUES (" & HFApp.DivisionID & ",0,0," & DbQuote(Str, .TextMatrix(r, .ColIndex("Vendor_ID")), , True) & ")"

+        bIgnoreDupKey = True

         Call HFApp.SqlExec(s, dbHomefront)

+        bIgnoreDupKey = False

         

         s = ""

         For c = 1 To .Cols - 1

@@ -551,10 +553,12 @@
     

     SaveData = True

 Exit Function

-eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then

+eh:

+    If bIgnoreDupKey And InStr(1, Err.Description, "duplicate", vbTextCompare) Then

         Resume Next

     Else

         Call errHandler(SRCFILE & "SaveData", s)

+        ProgressBar.Visible = False

     End If

 End Function

 

```

---

### `HFSystem/Source/FJob.frm` (Modified)
- **Old Size:** 147425 bytes | **New Size:** 148198 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +109 additions, -107 deletions.
- **Added methods/subroutines:** Sub gProperties_ComboCloseUp
- **Removed methods/subroutines:** Sub txtHoldback_Change, Sub txtHoldback_GotFocus, Sub txtHoldback_Validate
- **Query/SQL updates detected.**
- **Change Hunks count:** 56
#### Detailed Diff (Truncated Preview):
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFSystem/Source/FJob.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFSystem/Source/FJob.frm
@@ -4,23 +4,23 @@
 Begin VB.Form FJob 

    Caption         =   "Job Setup"

    ClientHeight    =   11070

-   ClientLeft      =   3765

-   ClientTop       =   1605

-   ClientWidth     =   24000

+   ClientLeft      =   6525

+   ClientTop       =   2055

+   ClientWidth     =   20430

    Icon            =   "FJob.frx":0000

    KeyPreview      =   -1  'True

    LinkTopic       =   "Form1"

    ScaleHeight     =   11070

-   ScaleWidth      =   24000

+   ScaleWidth      =   20430

    Begin MSComctlLib.Toolbar Toolbar 

       Align           =   1  'Align Top

       Height          =   600

       Left            =   0

       Negotiate       =   -1  'True

-      TabIndex        =   25

+      TabIndex        =   24

       Top             =   0

-      Width           =   24000
...

         .RowOutlineLevel(r) = 1

@@ -3392,7 +3394,7 @@
             Case Bit:

             Case DateTime:

                 If IsDate(.EditText) Or .EditText = "" Then

-                    .EditText = Format(.EditText, HFApp.Options(DateFormat))

+                    .EditText = Format(.EditText, "medium date")

                 Else

                     Cancel = True

                 End If

```

---

### `HFSystem/Source/FLogin.frm` (Modified)
- **Old Size:** 54123 bytes | **New Size:** 55402 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +82 additions, -67 deletions.
- **Query/SQL updates detected.**
- **Change Hunks count:** 9
#### Detailed Diff (Truncated Preview):
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFSystem/Source/FLogin.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFSystem/Source/FLogin.frm
@@ -457,7 +457,7 @@
         If section <> "" Then

             pwd = IniGet(file, section, "pwd")

             If pwd <> "" Then

-                Call IniPut(file, section, "pwe", HFApp.Encrypt(pwd))

+                Call IniPut(file, section, "pwe", Base64Encode(StrConv(HFApp.Encrypt(pwd), vbFromUnicode)))

                 Call IniRemove(file, section, "pwd")

             End If

         End If

@@ -471,8 +471,8 @@
         If section <> "" Then

             pwd = IniGet(file, section, "pwd")

             If pwd <> "" Then

-                Call IniPut(file, section, "pwe", HFApp.Encrypt(pwd))

-                Call IniRemove(file, section, "pwd")

+                Call IniPut(file, section, "pwe", Base64Encode(StrConv(HFApp.Encrypt(pwd), vbFromUnicode)))

+'                Call IniRemove(file, section, "pwd")

             End If

         End If

     Next

@@ -515,18 +515,20 @@
             .Clear

             sections = IniGetSectionNames(sIniFile)

             For i = 1 To Parse(sections)

-            

                 CompanyName = Parse(sections, i)
...
+'    s = s & ",isnull(v.city,'')       " & vbCrLf

+'    s = s & ",isnull(v.state,'')      " & vbCrLf

+'    s = s & ",isnull(v.zip,'')        " & vbCrLf

+'    s = s & ",isnull(v.Phone,'')      " & vbCrLf

+'    s = s & ",isnull(v.email,'')      " & vbCrLf

+'    s = s & ",isnull(v.PurchEmail,'') " & vbCrLf

+'    s = s & ",isnull(v.SchedEmail,'')" & vbCrLf

     

     s = s & "order by [Divisions!2!Code!element], [Parent], [Stats!3!Year!element], [Stats!3!Month!element]" & vbCrLf

     s = s & "for xml explicit" & vbCrLf

```

---

### `HFSystem/Source/FOptions.frm` (Modified)
- **Old Size:** 339550 bytes | **New Size:** 350688 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +1181 additions, -913 deletions.
- **Query/SQL updates detected.**
- **Change Hunks count:** 391
#### Detailed Diff (Truncated Preview):
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFSystem/Source/FOptions.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFSystem/Source/FOptions.frm
@@ -25,18 +25,152 @@
       Caption         =   "Accounting Integration"

       Height          =   12075

       Index           =   4

-      Left            =   4770

-      TabIndex        =   172

+      Left            =   2400

+      TabIndex        =   182

       Tag             =   "Security"

-      Top             =   1110

+      Top             =   270

       Visible         =   0   'False

       Width           =   23145

+      Begin VB.Frame AccountingFrame 

+         Caption         =   "D365 -- ABN"

+         Height          =   2325

+         Index           =   8

+         Left            =   30

+         TabIndex        =   410

+         Top             =   930

+         Width           =   6375

+         Begin VB.TextBox txtABND01Division 

+            BorderStyle     =   0  'None

+            Height          =   240

+            Left            =   1605
...

                 Call HFApp.SqlExec(s)

                 .Cell(flexcpData, r, .ColIndex("documentclass")) = .TextMatrix(r, .ColIndex("documentclass"))

@@ -8862,7 +9131,6 @@
                 s = s & "update dms_documentclasses set" & vbCrLf

                 s = s & " documentclass=" & DbQuote(Str, .TextMatrix(r, .ColIndex("documentclass"))) & vbCrLf

                 s = s & ",vendorvisible=" & DbQuote(Bit, .TextMatrix(r, .ColIndex("vendorvisible"))) & vbCrLf

-                s = s & ",customervisible=" & DbQuote(Bit, .TextMatrix(r, .ColIndex("customervisible"))) & vbCrLf

                 s = s & "where documentclass=" & DbQuote(Str, .Cell(flexcpData, r, .ColIndex("documentclass"))) & vbCrLf

                 Call HFApp.SqlExec(s)

                 .Cell(flexcpData, r, .ColIndex("documentclass")) = .TextMatrix(r, .ColIndex("documentclass"))

```

---

### `HFSystem/Source/FPOIndex.frm` (Modified)
- **Old Size:** 26327 bytes | **New Size:** 27731 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +93 additions, -59 deletions.
- **Added methods/subroutines:** Sub chkPreConPO_Click, Sub chkWarrantyPO_Click
- **Change Hunks count:** 24
#### Detailed Diff (Truncated Preview):
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFSystem/Source/FPOIndex.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFSystem/Source/FPOIndex.frm
@@ -3,24 +3,41 @@
    BorderStyle     =   3  'Fixed Dialog

    Caption         =   "Purchase Order"

    ClientHeight    =   7575

-   ClientLeft      =   3855

-   ClientTop       =   2745

+   ClientLeft      =   2025

+   ClientTop       =   3435

    ClientWidth     =   8670

    Icon            =   "FPOIndex.frx":0000

    LinkTopic       =   "Form1"

-   LockControls    =   -1  'True

    MaxButton       =   0   'False

    MinButton       =   0   'False

    ScaleHeight     =   7575

    ScaleWidth      =   8670

    ShowInTaskbar   =   0   'False

+   Begin VB.CheckBox chkPreConPO 

+      Alignment       =   1  'Right Justify

+      Caption         =   "Precon PO"

+      Height          =   225

+      Left            =   6060

+      TabIndex        =   13

+      Top             =   2610

+      Width           =   1920
...
-    

+    If Trim(txtPOIndex.Text) = "" Then s = s & "Purchase order is required" & vbCrLf

+    

+    If HFApp.Options.ValueByName("BuildProCompanyCode") <> "" And chkBuildPro.Value = vbChecked Then

+        If Trim(cboJCCostCode.Text) = "" Then s = s & "Cost Code is required" & vbCrLf

+        If Trim(cboJCCategory.Text) = "" Then s = s & "Category is required" & vbCrLf

+    End If

     

     If s <> "" Then

         ValidateData = False

```

---

### `HFSystem/Source/FUserPermissions.frm` (Modified)
- **Old Size:** 121250 bytes | **New Size:** 115115 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +496 additions, -658 deletions.
- **Query/SQL updates detected.**
- **Change Hunks count:** 38
#### Detailed Diff (Truncated Preview):
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFSystem/Source/FUserPermissions.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFSystem/Source/FUserPermissions.frm
@@ -11,8 +11,116 @@
    ClientWidth     =   14370

    Icon            =   "FUserPermissions.frx":0000

    LinkTopic       =   "Form1"

-   ScaleHeight     =   15255

-   ScaleWidth      =   28800

+   ScaleHeight     =   7770

+   ScaleWidth      =   14370

+   Begin VB.Frame TabFrame 

+      Caption         =   "Data Portal"

+      Height          =   4995

+      Index           =   9

+      Left            =   6972

+      TabIndex        =   59

+      Top             =   3264

+      Width           =   7005

+      Begin VSFlex8Ctl.VSFlexGrid gPermissions 

+         Height          =   5655

+         Index           =   9

+         Left            =   150

+         TabIndex        =   60

+         TabStop         =   0   'False

+         Top             =   240

+         Width           =   8295

+         _cx             =   1976842535
...
-            s = "delete securitygroupdocumentclasses where SecGroupID=" & DbQuote(Num, mSecGroupID)

-        Else

-            s = "delete securitygroupdocumentclasses where SecGroupID=" & DbQuote(Num, mSecGroupID) & vbCrLf & _

-                "insert securitygroupdocumentclasses(SecGroupID,documentclass) values" & vbCrLf & Mid(s, 2)

-        End If

-        Call HFApp.SqlExec(s)

-    End With

     

     

     

```

---

### `HFSystem/Source/FVendor.frm` (Modified)
- **Old Size:** 89236 bytes | **New Size:** 88929 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +2 additions, -11 deletions.
- **Query/SQL updates detected.**
- **Change Hunks count:** 2
#### Diff Preview:
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFSystem/Source/FVendor.frm
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFSystem/Source/FVendor.frm
@@ -14,15 +14,6 @@
    ScaleWidth      =   16395

    Begin VB.TextBox txtTaxID 

       BorderStyle     =   0  'None

-      BeginProperty Font 

-         Name            =   "MS Sans Serif"

-         Size            =   8.25

-         Charset         =   0

-         Weight          =   400

-         Underline       =   0   'False

-         Italic          =   0   'False

-         Strikethrough   =   0   'False

-      EndProperty

       Height          =   270

       Left            =   1455

       Locked          =   -1  'True

@@ -2122,7 +2113,7 @@
     s = s & "select *" & vbCrLf

     s = s & "  from contacts" & vbCrLf

     s = s & " where DivisionID = " & HFApp.DivisionID & " and contacttypeid=99" & vbCrLf

-    s = s & "   and vendorcode=" & DbQuote(Str, mVendor)

+    s = s & "   and vendorcode=" & DbQuote(Str, mVendor) & vbCrLf

     s = s & "order by firstname" & vbCrLf

     Set rs = HFApp.SqlExec(s)

     With gContacts

```

---

### `HFSystem/Source/MBuildPro.bas` (Modified)
- **Old Size:** 23851 bytes | **New Size:** 23826 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +126 additions, -113 deletions.
- **Query/SQL updates detected.**
- **Change Hunks count:** 28
#### Detailed Diff (Truncated Preview):
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFSystem/Source/MBuildPro.bas
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFSystem/Source/MBuildPro.bas
@@ -9,19 +9,19 @@
 Private Function GetImpDate(Name As String) As Date

 On Error Resume Next

     Dim dt As Date

-    

+

     dt = "1970-01-01"

     dt = CDate(HFApp.Options.ValueByName(Name))

     If dt < DateValue("1970-01-01") Then dt = DateValue("1970-01-01")

     GetImpDate = dt

-    

+

 End Function

 

 Private Sub PutImpDate(Name As String)

 On Error Resume Next

-    

+

     HFApp.Options.ValueByName(Name) = Now()

-    

+

 End Sub

 

 Private Sub LogXML(Name As String, XmlMessage As String)

@@ -42,15 +42,15 @@
...
     Screen.MousePointer = vbNormal

-    

-    

-        

+

+

+

 Exit Function

 eh: Call errHandler(SRCFILE & "SendBuildProPOs")

 End Function

```

---

### `HFSystem/Source/MCrmDataService.bas` (Modified)
- **Old Size:** 1747 bytes | **New Size:** 1750 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

#### Minor changes:
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFSystem/Source/MCrmDataService.bas
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFSystem/Source/MCrmDataService.bas
@@ -55,9 +55,9 @@
 End Function

 

 Private Function SalesWrapper(xml As String) As String

-    Dim crm As New HyphenSys.SalesWrapper

-    Dim results As String

-    SalesWrapper = crm.PostXml(CrmClientID, CrmApiKey, CrmEnvironment = "Production", "1", xml)

+'    Dim crm As New HyphenSys.SalesWrapper

+'    Dim results As String

+'    SalesWrapper = crm.PostXml(CrmClientID, CrmApiKey, CrmEnvironment = "Production", "1", xml)
...
```

---

### `HFSystem/Source/MD365_ABN.bas` (Added)
- **Old Size:** 0 bytes | **New Size:** 893 bytes
- **Old Attributes:** `-` | **New Attributes:** `read|write`

- **New file added** with 0 lines.
#### File Preview (First 30 lines):
```vb
Attribute VB_Name = "MD365_ABN"
'-------------------------------------------------------------
' Financial Dimension Mapping
'-------------------------------------------------------------
'VALUE              SEND WITH       MAPPED TO
'legal entity       job & PO        System setting, one per HF division.
'D01-Division       job & PO        System setting, one per HF division.
'D02-Function       job             New field on tblLocality
'D03-CostCenter     job             New field on communityphase
'D06-Brand          job             New field on communityphase
'D04-SpendCat       PO              Debitacct from poindex/costcode
'
'Group costcode description will contain a coded string in the format "ProjectCategory / ItemCode". Expected values are
'  "JobConsumables_Item / Proj_IntC"
'  "JobSubcontractor_Item / Proj_IntS"
'  "JobFulfillment_Item / Proj_IntF"
'

```

---

### `HFSystem/Source/MDBUpgrade4.bas` (Modified)
- **Old Size:** 2192697 bytes | **New Size:** 4425184 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +33950 additions, -198 deletions.
- **Added methods/subroutines:** Sub Revisions_2025_2_h, Sub Revisions_2025_2_i, Sub Revisions_2025_2_j, Sub Revisions_2025_3, Sub Revisions_2025_3_1, Sub Revisions_2025_3_a, Sub Revisions_2025_3_b, Sub Revisions_2025_4, Sub Revisions_2025_4_a, Sub Revisions_2025_4_b, Sub Revisions_2025_4_c, Sub Revisions_2025_4_d, Sub Revisions_2025_4_e, Sub Revisions_2025_4_f, Sub Revisions_2025_4_g
- **Query/SQL updates detected.**
- **Change Hunks count:** 61
#### Detailed Diff (Truncated Preview):
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFSystem/Source/MDBUpgrade4.bas
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFSystem/Source/MDBUpgrade4.bas
@@ -16766,64 +16766,8 @@
     s = s & "END" & vbCrLf

     sql c, s

     

-    s = ""

-    s = s & "alter view dbo.BuildPro_POs as" & vbCrLf

-    s = s & "select " & vbCrLf

-    s = s & " p.DivisionID " & vbCrLf

-    s = s & ",p.PONumber" & vbCrLf

-    s = s & ",convert(varchar,p.PODate,101) PODate" & vbCrLf

-    s = s & ",p.tstmp ModifiedDate" & vbCrLf

-    s = s & ",isnull(nullif(EPODocType,''),'PO') PODocType" & vbCrLf

-    s = s & ",'' PODocSuffix" & vbCrLf

-    s = s & ",p.EPOId" & vbCrLf

-    s = s & ",p.EPODocType" & vbCrLf

-    s = s & ",p.EPODocSuffix" & vbCrLf

-    s = s & ",dbo.BuildProSOAP_BPVendor(p.divisionid,p.vendor) Vendor" & vbCrLf

-    s = s & ",p.POIndex" & vbCrLf

-    s = s & ",i.phase+'/'+i.item sku" & vbCrLf

-    s = s & ",case when p.CancelledDate is null then 0 else 1 end POCancelled" & vbCrLf

-    s = s & ",p.Job HFJob" & vbCrLf

-    s = s & ",pj.BPJob Job" & vbCrLf

-    s = s & ",j.Community " & vbCrLf

-    s = s & ",0 CommunityPhase --- not used" & vbCrLf

-    s = s & ",poi.LineNumber" & vbCrLf
...
+    s = s & "alter table tblcustomers enable trigger tiu_tblcustomers " & vbCrLf

+    s = s & "alter table tbljobs enable trigger u_tbljobs" & vbCrLf

+    sql c, s

+

+    Call Revisions_2025_4_h

+

+End Sub

+

+

+

```

---

### `HFSystem/Source/MDBUpgrade5.bas` (Added)
- **Old Size:** 0 bytes | **New Size:** 2092638 bytes
- **Old Attributes:** `-` | **New Attributes:** `read|write`

- **New file added** with 0 lines.
#### File Preview (First 30 lines):
```vb
Attribute VB_Name = "MDBUpgrade5"
Option Explicit
Option Compare Text


Private Sub sql(Description As String, SQLStatement As String, Optional IgnoreErrors As Boolean, Optional DontLog As Boolean)
    FDbUpgrade.sql Description, SQLStatement, IgnoreErrors, DontLog
End Sub


Public Sub Revisions_2025_4_h()

Dim s As String
Dim c As String

c = "BuildPro MsgQ procs"

    s = ""
    s = s & "ALTER proc [dbo].[MsgQ_BuildPro_PutCommunities_Pre] @ID integer, @CutOffDate datetime as " & vbCrLf
    s = s & " -- CutOffDate argument is the last date this proc was run. Use it to filter for changed records. " & vbCrLf
    s = s & " -- save xml results to the response column. put errors in the status field. " & vbCrLf
    s = s & " -- use syntax like this to update the audit log if there is anything useful to be said. " & vbCrLf
    s = s & " --   update MsgQ set Audit=isnull(Audit,'') + 'your log entry here' + char(13) + char(10) where ID=@ID " & vbCrLf
    s = s & "begin try " & vbCrLf
    s = s & "    declare @DivisionID int  " & vbCrLf
    s = s & "    select @DivisionID=divisionid from MsgQ where id=@ID " & vbCrLf
    s = s & " " & vbCrLf
    s = s & "    declare @xml xml =( " & vbCrLf
    s = s & "    select  " & vbCrLf
    s = s & "     community + '.0' CommunityNumber" & vbCrLf
```

---

### `HFSystem/Source/MIntacct.bas` (Modified)
- **Old Size:** 10638 bytes | **New Size:** 10657 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

#### Minor changes:
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFSystem/Source/MIntacct.bas
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFSystem/Source/MIntacct.bas
@@ -216,8 +216,9 @@
     Wend

     X.CloseMessage

     Call WriteLogFile("intacct.writecustomers.req.xml", X.xml())

+    

     s = X.PostMessage(False)

-    Call WriteLogFile("intacct.writecustomers.res.xml", s)

+    Call WriteLogFile("intacct.writecustomers.res.xml", X.lastResponse)

         

         
...
```

---

### `HFSystem/Source/MMisc.bas` (Modified)
- **Old Size:** 135877 bytes | **New Size:** 137448 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

- **Lines changed:** +42 additions, -1 deletions.
- **Added methods/subroutines:** Function Base64Decode, Function Base64Encode
- **Change Hunks count:** 2
#### Diff Preview:
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFSystem/Source/MMisc.bas
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFSystem/Source/MMisc.bas
@@ -804,6 +804,12 @@
 Private Declare Function PtrToStr Lib "kernel32" Alias "lstrcpyW" (RetVal As Byte, ByVal Ptr As Long) As Long

 Private Declare Function StrLen Lib "kernel32" Alias "lstrlenW" (ByVal Ptr As Long) As Long

 

+Private Declare Function CryptBinaryToStringA Lib "crypt32.dll" (ByVal pbBinary As Long, ByVal cbBinary As Long, ByVal dwFlags As Long, ByVal pszString As Long, ByRef pcchString As Long) As Long

+

+Private Declare Function CryptStringToBinaryA Lib "crypt32.dll" (ByVal pszString As Long, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, ByRef pcbBinary As Long, ByVal pdwSkip As Long, ByVal pdwFlags As Long) As Long

+Private Const CRYPT_STRING_BASE64 = &H1

+Private Const CRYPT_STRING_BASE64_ANY = &H6

+

 

 Public Function GetLocalizedPath(sPath As String) As String

 'Dim d As String

@@ -3684,3 +3690,38 @@
 End Function

 

 

+

+

+

+     

+

+Function Base64Decode(ByVal data As String) As String

+  Dim s As Long

+  s = Len(data)

+  Dim b() As Byte

+

+  b = StrConv(data, vbFromUnicode)

+  '??VB???Unicode??,??????Ansi

+  Dim ret() As Byte

+  Dim retlen As Long

+  Call CryptStringToBinaryA(VarPtr(b(0)), s, CRYPT_STRING_BASE64_ANY, StrPtr(ret), retlen, 0, 0)

+  If retlen = 0 Then Base64Decode = "": Exit Function

+  ReDim ret(retlen - 1)

+  Call CryptStringToBinaryA(VarPtr(b(0)), s, CRYPT_STRING_BASE64_ANY, VarPtr(ret(0)), retlen, 0, 0)

+  Base64Decode = StrConv(LeftB(ret, retlen), vbUnicode)

+End Function

+Function Base64Encode(data() As Byte) As String

+  Dim s As Long

+  s = UBound(data) + 1

+

+  Dim ret() As Byte

+  Dim retlen As Long

+  Call CryptBinaryToStringA(VarPtr(data(0)), s, 1073741825, StrPtr(ret), retlen)

+  If retlen = 0 Then Exit Function

+

+  'MsgBox retlen

+  ReDim ret(retlen - 1)

+  Call CryptBinaryToStringA(VarPtr(data(0)), s, 1073741825, VarPtr(ret(0)), retlen)

+  Base64Encode = StrConv(LeftB(ret, retlen), vbUnicode)

...
```

---

### `HFSystem/Source/MQuickBooksOnline.bas` (Modified)
- **Old Size:** 7303 bytes | **New Size:** 7588 bytes
- **Old Attributes:** `read|write` | **New Attributes:** `read|write`

#### Minor changes:
```diff
--- Original//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/Original/HFSystem/Source/MQuickBooksOnline.bas
+++ New//Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6/HFSystem/Source/MQuickBooksOnline.bas
@@ -135,6 +135,11 @@
             s = s & "   and job_no=" & DbQuote(Str, "" & rs("Job_No")) & vbCrLf

             Call HFApp.SqlExec(s, dbHomefront)

             

+            

+            'if style = simple then write listid to customer and to job

+            'if style = heirarchy then write parentid to customer and listid to job

+            If HFApp.Options.ValueByName("QuickBooksJobStyle") <> "Simple" Then listid = parentid

+            

             'add customer to HF
...
```

---

