# HomeFront — Operating Rules

Canonical conventions for the VB6 → Blazor (.NET 10) migration. Applies to **HomeFront**,
**HomeFrontPB**, **FlexKit**, and **FlexCore**. These are hard rules, not suggestions —
follow them unless a specific instruction overrides one for a single task.

---

## 1. VB6 fidelity — the prime directive
- **100% fidelity to VB6** in look, feel, and behavior. Read the matching `.frm` / `.bas`
  (under `HomeFront/HomeFrontVB6`) **before** building or debugging a migrated `F*` page —
  most bugs are translation errors.
- **Never hardcode what VB6 derives at runtime** — App-Options, `AppGridLayout`, DB-driven
  values. (Design-time *control definitions* ARE faithfully hardcoded — that is different.)
- **Grid columns come from `AppGridLayout`** (VB6 `IniGetGrid` → `GetGridLayoutAsync`), rendered
  **dynamically**. Never hardcode a `<GridColumn>` list.
- **Keep the exact VB6 form + control names** as lookup strings — `AppGridLayout` is keyed
  `gridname="FormName.GridControlName"` (e.g. `FAddPricelist.gItems`). The C# `@ref` may differ,
  but the lookup strings must equal the VB6 names. Same for App-Options / colkeys / SQL.
- **No silent fallbacks** — surface them visibly, e.g. the `[gItems FALLBACK]` marker in the title.

## 2. FlexKit is the only UI control source
- **All UI must be FlexKit** (`Fx.ControlKit.*`, `fx-*` CSS). Any raw `<input>` / `<select>` /
  `<dialog>` / `<div>` modal or other hand-rolled host control is a **red flag**.
- **No one-off controls.** If FlexKit lacks a primitive, **create it in FlexKit (and mirror to
  FlexCore)** — never hand-roll it in the host app.
- **Grep FlexKit first** before creating any control — check `Dialogs/`, `Grid/`, `Reports/`,
  `Drawing/` subfolders (e.g. a `ColorPickerControl` already exists in `Fx.ControlKit.Dialogs`).
- **Modals:** use `DialogControl`, or host a whole `F*` form as a modal. Never a raw `<div>` overlay.
- **Use the shared `GridLayoutPresenter` service** (HF `Services/`) for grid-layout work — don't
  re-copy the per-form helper cluster.
- **Render-time gotcha:** only `TextBoxControl` captures unmatched attributes. Passing `style=` or
  other arbitrary attrs to `NumericTextBoxControl` / `DropDownListControl` / `RadioControl` /
  `CheckBoxControl` / `ColorPickerControl` compiles green but **throws at render** — use the
  control's own `Width` / `CssClass` params instead.

## 3. Minimize JavaScript
- **Little-to-no JS anywhere** (HomeFront / PB / FlexKit / FlexCore). Prefer Blazor-native APIs
  (`FocusAsync`, `@ref`, `@bind`, `EventCallback`) and **CSS** (`position: sticky`, `:focus-within`).
- JS **only** when the DOM capability has no C#/CSS equivalent — then one tiny lazily-imported
  export with a C# fallback. **Never a new `.js` file or inline `<script>` without flagging it.**

## 4. Keep projects in sync (and build both sides)
- **FlexKit ↔ FlexCore:** NOT auto-mirrored (owner directive 2026-07-25). They may diverge;
  sync between them ONLY when explicitly requested (perf work like row virtualization is developed
  in FlexCore via the FlexKitTester bench while FlexKit stays stable for the live apps).
  - Verify: `diff -rq FlexCore FlexKit | grep -v "/bin\|/obj\|/.git\|.DS_Store\|\.csproj$"` → empty.
- **HomeFront ↔ HomeFrontPB:** mirror every migrated-form fix. HF keeps forms under
  `Components/Pages/Migrated/F*.razor`; PB keeps them flat in `Components/Pages/F*.razor`.
- **Two-tier pairing:** FlexCore.Showcase ↔ FlexCore; HomeFront / PB / GhostWriter /
  GhostWriterLLM ↔ FlexKit (live `ProjectReference`).
- **After any FlexKit edit, build BOTH** `HomeFront.sln` **and** `HomeFrontPB.sln`. After any
  FlexCore edit, rebuild `FlexCore.Showcase.sln`.
- Use `--no-incremental` after FlexKit edits (incremental can mask errors in other FlexKit files).
- **Both `CLAUDE.md` files are kept current.**

## 5. Active vs. archived apps
- **HomeFrontPB is the active app.** **HomeFrontPOC is archived — never build, edit, deploy, or
  sync it.**
- Namespace stays `Fx.ControlKit.*` and CSS prefix stays `fx-*` in **both** libraries — only the
  project / assembly / package name differs.

## 6. Security / SQL / DB
- **Never interpolate user-derived values into SQL** — always parameterize.
- **No schema changes from app code** — no `ALTER TABLE`, etc.
- **Report XMLs are read-only** — they regenerate from `.rpt`. Fix at the host-service / loader
  level instead.
- **No per-user static state** in Blazor Server — statics are shared across all circuits
  (multi-user leak). Use scoped services (`ISessionStateService`).

## 7. Operational
- **The user runs HomeFront / PB themselves.** Never auto-start `dotnet run`. Building is fine;
  live execution needs an explicit ask. Kill any run session started for verification as soon as
  it is no longer needed.
- **Don't generate the modified-files list or the deployment zip proactively** — only when
  explicitly asked (they have specific required formats).
- **FlexKit / FlexCore stay general-purpose** — library code must never name or depend on
  HomeFront. Host-specific needs go through interfaces (`IReportDataExecutor`,
  `IReportPickListProvider`, …) + host-side adapters.

---

## Build commands

```bash
# Libraries
dotnet build /Users/wadood/projects/VBToCSharp/FlexKit/FlexKit.csproj
dotnet build /Users/wadood/projects/VBToCSharp/FlexCore/FlexCore.csproj

# Showcase (builds FlexCore alongside)
dotnet build /Users/wadood/projects/VBToCSharp/FlexCore.Showcase/FlexCore.Showcase.sln

# HomeFront app (builds FlexKit alongside)
dotnet build /Users/wadood/projects/VBToCSharp/HomeFront/MobileSource/HomeFront/HomeFront.sln

# HomeFrontPB (builds FlexKit alongside via ProjectReference)
dotnet build /Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontPB/HomeFrontPB.sln
```

After any **FlexKit** change, rebuild **both** `HomeFront.sln` and `HomeFrontPB.sln`.
After any **FlexCore** change, rebuild `FlexCore.Showcase.sln`.
