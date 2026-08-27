# HomeFront — Code-Quality / Improvement Audit

**Date:** 2026-06-15 · **Scope:** HomeFront (HF) app + the FlexKit controls it builds against.
**Method:** full `--no-incremental` build (warning census) + 3 parallel read-only code surveys
(duplication, dead code, JS reduction) + targeted verification of the top findings.
**Nothing was changed by this audit — it is a findings + prioritized plan document.**

---

## Executive summary

| Lever | Size | Effort | Notes |
|---|---|---|---|
| **1. Kill BL0005 at the source** | **~679 of 983 warnings (69%)** | **~19 one-line edits** | FlexKit settings DTOs carry vestigial `[Parameter]`. Verified safe. **Do this first.** |
| 2. Dedupe `using`s in `HomeFront.cs` | 42 warnings | 1 file | Trivial. |
| 3. Delete 5 orphaned forms | ~3,742 LOC + some warnings | verify+delete | No route, no nav. Needs your OK to delete. |
| 4. Remove unused vars/fields | ~70 warnings | mechanical | CS0168/0219/0169/0414/0649 across ~37 files. |
| 5. Migrate obsolete static calls → DI | ~57 warnings | medium | CS0618 = real multi-user-leak risk (`MMain.DivisionID/db/...`). |
| 6. `IDialogService` → kill confirm/alert/prompt JS | ~152 JS sites | high | Building blocks already exist in FlexKit. |
| 7. Extract grid-layout helper cluster | ~5,000 LOC, 37 forms | high | Biggest dedup; also prevents drift bugs. |
| 8. Fix FQuote column-width drift bug | 1 line | trivial | Confirmed copy-paste drift. |
| 9. Hardcoded SA connection string | 1 site (security) | small | `FrmOptionListMultiSelect.razor:488`. |

After **#1 + #2 + #4** (all low-risk/mechanical), the warning count drops from **983 → ~190**.
After **#3 + #5** it drops to **well under 100**, mostly package-level (Crystal) noise.

---

## 1. Compiler warnings — 983 unique (1,966 raw, each emitted twice)

| Code | Unique | Meaning | Verdict / fix |
|---|---|---|---|
| **BL0005** | **~679** | "Component parameter X set outside its component" | **ROOT CAUSE found — see below. 19 edits clears all.** |
| CS0618 | 57 | Use of `[Obsolete]` member | Migrate to DI (`ISessionStateService`/`DbWrapperSqlServer`/`IGridLayoutService`). Real leak risk. |
| CS0414 | 52 | Private field assigned, never used | Delete the fields. |
| CS0105 | 42 | Duplicate `using` directive | **41 are in one file: `HomeFront.cs`.** Dedupe. |
| NU1701 | 28 | Pkg restored via .NET-FW fallback | Crystal Reports pkg — expected, package-level, not code. |
| CA1416 | 26 | Windows-only API on cross-plat | Crystal/registry paths — guard with `[SupportedOSPlatform]` or suppress. |
| CS8618 | 25 | Non-nullable field uninitialized | Nullability — annotate `?` / `= null!` / init. Low priority. |
| CS8601/8619/8604/8625/8620/8714 | ~46 | Nullability mismatches | Annotate / null-check. Low priority. |
| CS0219 / CS0168 / CS0169 / CS0649 | ~17 | Local/field declared-or-assigned, never used | **The literal "variables declared not used."** Delete. |
| NU1902 / NU1901 | 10 | NuGet pkg vulnerability advisory | Review pkg versions (security). |
| CS4014 | 1 | Async call not awaited | **Inspect — possible bug.** |
| CS7022 | 1 | Entry point ignored | Stray `Program`/top-level conflict — inspect. |
| CS0162 | 1 | Unreachable code | Delete. |

### 1a. BL0005 root cause (THE headline finding)

`FlexKit/Grid/{EditSettings,SelectionSettings,FilterSettings,PageSettings}.cs` each:
- inherit `: ComponentBase`, **and**
- decorate every property with `[Parameter]` (19 total: Edit 7, Selection 6, Filter 3, Page 3).

But these types are **never rendered as markup** (`<EditSettings>` appears nowhere). Hosts build
them as plain data objects and set the properties in C#:
```csharp
private EditSettings editSettings = new() { AllowAdding = true, AllowEditing = true, ... };
```
…then pass the *object* to `GridControl` via `[Parameter] public EditSettings? EditSettingsRef`.
So every C# assignment to an `[Parameter]` property trips the analyzer → **~679 BL0005 across 37 forms**
(FAssembly 52, FVendor 46, FAttributeLists 46, FAddPricelist 38, DataGrid 32, FDBGrid 32, …).

**Fix (FlexKit, mirror to FlexCore byte-identically):** drop `[Parameter]` from the 19 properties
(optionally also drop `: ComponentBase` so they become plain POCO DTOs — what they already are in
practice). **Verified zero behavior change** — Blazor never bound these (the class is never a
component instance); GridControl just reads the object's properties. **~679 warnings → 0 from 19
one-line deletions.**

---

## 2. Dead / orphaned code (~3,742 LOC) + 1 security finding

**Orphaned forms** — `@page` route present, but NOT in `FMain._routeMap` and zero `NavigateTo`/`typeof`/`@ref`
references (VERY HIGH confidence):

| File | Route | LOC |
|---|---|---|
| `ItemDBCodex.razor` | /items-codex, /items-db-codex | 1,047 |
| `ModelAssembly.razor` | /model-assembly | 906 |
| `FHome2.razor` | /home2 | 651 |
| `VendorOpus.razor` | /vendor-opus | 614 |
| `DefaultVendorsOpus.razor` | /default-vendor-opus | 524 |

These are legacy/superseded duplicates (the canonical forms — FItems, FVendor, FDefaultVendors, FHome — are
the live ones). **Recommend deletion** after a final `git grep` per file. `FPickListOrig.razor` is a
**maybe** (referenced via `@ref` in IndexPicklist — review before touching).

**Security — hardcoded SA connection string:** `FrmOptionListMultiSelect.razor:488`
```csharp
private string ConnectionString = "Server=localhost,1433;Database=HomeFrontDB;User Id=sa;Password=YourPassword;...";
```
The form is **live** (route `/option-list-multi-select` is in `_routeMap`). It bypasses the injected
`DbWrapperSqlServer` + the whole SA-password-secret hardening. **Fix:** delete the literal, inject
`DbWrapperSqlServer` like every other form. (Password is a placeholder today, but the pattern ships.)

No large commented-out code blocks found (clean). No dead `_routeMap` entries (clean).

---

## 3. JavaScript reduction (~152 interop sites removable)

Custom JS footprint is already small (4 files, 373 LOC). The problem is **interop call sites that
bypass FlexKit**:

| Interop | Sites | Class | Replacement |
|---|---|---|---|
| `confirm` | 75 | A | FlexKit modal — `await Dialogs.ConfirmAsync(msg)` |
| `alert` | 43 | A/B | modal (`AlertAsync`) or existing `NotificationService` |
| `prompt` | 29 | A | FlexKit modal — `await Dialogs.PromptAsync(msg)` |
| `eval` | 24 | mixed | 5 → dialog; ~9 → `FocusAsync`/`InputFile`; ~4 stay (clipboard/print/drag) but as typed exports |

**Generalization — `IDialogService` (FlexKit + FlexCore mirror).** The dialog UI **already exists**:
`Fx.ControlKit.Dialogs.MessageBoxControl` (`ShowAsync(msg,title,buttons,icon)` with Ok/OkCancel/YesNo/
YesNoCancel + VB6 look) and `InputDialogControl` (`ShowAsync(prompt,title,default)`). FLogin already uses
this pattern per-page. Lift it into one app-wide service:
```csharp
public interface IDialogService {
    Task<bool>             ConfirmAsync(string msg, string title = "Confirm");
    Task<MessageBoxResult> ConfirmCancelAsync(string msg, string title = "Confirm"); // 3-state
    Task                   AlertAsync(string msg, string title = "");
    Task<string?>          PromptAsync(string msg, string title = "", string def = "");
}
```
Host one `<DialogHostControl/>` in `MainLayout.razor`, register `AddScoped<IDialogService,DialogService>()`
in `Program.cs` (exactly mirroring `NotificationService`). Call sites change from
`await JSRuntime.InvokeAsync<bool>("confirm", m)` → `await Dialogs.ConfirmAsync(m)` — already `await`ed,
so ergonomics barely change. **Eliminates all 147 confirm/alert/prompt sites + 5 eval-wrapped dialogs.**

**`eval` cleanup:** replace `document…select()/focus()` with `ElementReference.FocusAsync()`; replace the
hidden-file-input `eval(...).click()` hack with a direct/visible `<InputFile>` (no JS) or one typed export;
keep clipboard/print/drag as **named typed exports** (no raw `eval` strings survive). Remove the inline
`<script>` in `FAdjustPrices.razor` (hfFocus → FocusAsync; hfSelectAll → shared typed export).

**Keep:** `fx-report.js`, `hfgrid-export.js`, `reconnect-modal.js` (+ App.razor reconnect script) — genuine
browser-only capabilities (Crystal viewer, file export, circuit-down reconnect UI). `fx-mousetrack.js` is a
drop candidate if FApprovals drag-resize is its only consumer.

---

## 4. Duplication / generalization (~5,000–6,000 LOC removable)

`Components/Pages/Migrated/` has **0 base classes** — every form is self-contained, so sharing today
is literal copy-paste (the house-style even prescribes "mirror FAssembly's helpers").

| # | Pattern | Forms | Recommendation | LOC saved |
|---|---|---|---|---|
| 1 | **Grid-layout helper cluster** (`Load*GridLayoutAsync`/`GetOrdered*GridLayout`/`Resolve*GridField`/`ConvertLayoutWidthToPx`/`NormalizeGridKey` + reflection map) | **37** | Injected `IGridLayoutPresenter<TRow>` wrapping `IGridLayoutService`; grid-name **strings stay per-form** (VB6 fidelity) | **~4,200** |
| 2 | `Dictionary<string,PropertyInfo>` maps + `NormalizeGridKey` | 27 | Fold into #1 (static per-`TRow` cache) | ~400 |
| 3 | `ConvertLayoutWidthToPx` twips→px | 36 | One helper (`FlexKit GridMetrics.TwipsToPx`) | ~350 + **fixes FQuote bug** |
| 4 | `ComboBoxItem` re-declared | 7 | Shared `Fx.ControlKit.ComboBoxItem` | ~120 |
| 5 | `mDirty`/`IsLoading`/save try-catch boilerplate | 35–67 | Optional `MigratedFormBase` (also centralizes the identical `@inject` block) | ~600–800 |
| 6 | FPickList wiring | 27 | Mostly fine — low priority | ~150 |
| 7 | DivisionID SQL | 96 | **Do NOT extract** — SQL is per-form/VB6-faithful | 0 |

**Confirmed drift bug (item 3):** `FQuote.ConvertLayoutWidthToPx` = `Math.Max(30, twips/15)` — **missing the
`+16` and the upper clamp** every other form has → FQuote grid columns render ~16px too narrow. A shared
helper would have prevented it. (1-line fix regardless.)

**Largest forms** (decomposition candidates): FAssembly (6,186 LOC, hosts 8 child forms + 2 grids),
FImportExportAssemblies (4,109, cleanly splittable by its 6 modes), FEstimateItems (3,769, mostly cohesive).
All three are "3–5k-line C# classes wearing a `.razor`." Lower priority than the shared-helper extraction.

---

## 5. Recommended order of execution

**Phase A — safe & mechanical (clears ~800 warnings, ~3.7k dead LOC):**
1. BL0005: drop `[Parameter]` from the 4 FlexKit settings classes (+FlexCore mirror). *(~679 gone)*
2. Dedupe `using`s in `HomeFront.cs`. *(42 gone)*
3. Remove unused vars/fields (CS0168/0219/0169/0414/0649). *(~70 gone)*
4. Fix FQuote width-drift (1 line) + the hardcoded-SA connection string.
5. Delete the 5 orphaned forms (with your sign-off).

**Phase B — structural (higher value, higher risk, build both .sln + FlexCore each step):**
6. `IDialogService` in FlexKit; migrate confirm/alert/prompt + eval-dialogs off JS.
7. `IGridLayoutPresenter<TRow>` extraction; migrate the 37 forms (incrementally, VB6-fidelity preserved).
8. CS0618: migrate obsolete static `MMain.*` calls to DI.

**Phase C — polish:** nullability annotations, CA1416 platform guards, NU190x package bumps, large-form
decomposition (FImportExportAssemblies by mode first).

> All FlexKit edits must mirror to FlexCore (byte-identical) and rebuild **both** HomeFront.sln **and**
> HomeFrontPB.sln. Phase B/C items that touch migrated forms must also be mirrored to HomeFrontPB per the
> HF↔PB sync rule (PB's FMain/Workflow are NOT to be overwritten).

---

## EXECUTION LOG (2026-06-15)

**Track 1 — warning cleanup ✅** HF warnings **983 → 184**, 0 errors.
- BL0005 (~679) → **0**: dropped `[Parameter]` (+ `: ComponentBase`) from 5 FlexKit config classes
  (`EditSettings`, `SelectionSettings`, `FilterSettings`, `PageSettings`, `GridControlEvents` — 61 attrs),
  mirrored byte-identical to FlexCore. These are data objects passed via `*Ref`, never rendered.
- CS0105 (42) → **0**: `_Imports.razor` already provided `Fx.ControlKit.Notifications`; stripped the redundant
  per-form `@using` from 40 forms + 1 intra-file dup.
- CS0414/0168/0169/0219/0649 (~70) → **~2**: removed unused fields/locals across ~34 forms (the residual
  `Application._jobSimplicity` is read-but-never-assigned — left intentionally).
- FQuote `ConvertLayoutWidthToPx` drift fixed (now matches the canonical `+16`/clamp formula).
- Hardcoded `User Id=sa;Password=…` connection string removed (it was a dead `ConnectionString` field).

**Track 2 — dead code ✅** Deleted **9 orphaned forms (~6,600 LOC)**: VendorOpus, DefaultVendorsOpus,
ItemDBCodex, ModelAssembly, FHome2, FPriceListLegacy, the `Application.razor` dup, plus `Communities.razor`
+ `CommunitiesCodex.razor` (the latter was even `@namespace BlazorDemos.Pages`). All unrouted, 0 references.

**Track 3 — IDialogService ✅** New `Fx.ControlKit.Dialogs.{IDialogService, DialogService, DialogHostControl}`
(wraps the existing MessageBoxControl/InputDialogControl), mirrored to FlexCore, hosted once in MainLayout,
registered in Program.cs. Migrated **131 → 0** native `confirm`/`alert`/`prompt` interop sites across 46 forms
to `ConfirmAsync`/`AlertAsync`/`PromptAsync`; the 2 `eval`-wrapped Yes/No/Cancel dialogs (FFloorplans,
FIntersection) regained real 3-state semantics via `ConfirmCancelAsync`.

All three solutions (HomeFront, HomeFrontPB, FlexCore.Showcase) build **0 errors** after each track.

**Pending follow-ups:** (a) mirror the Track-1 form cleanups + Track-3 dialog migration into HomeFrontPB
(its forms still call native confirm/alert/prompt; the FlexKit infra already flows in); (b) **Track 4**
(grid-layout presenter extraction) — not yet started.
