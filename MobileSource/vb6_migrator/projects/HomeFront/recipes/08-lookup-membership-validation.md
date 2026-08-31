# Recipe 08 — Typed lookup-cell membership validation (VB6 ValidateField)

VB6 grids pair a combo-button cell with a membership check on TYPED entry:

```vb
' BeforeEdit: the cell is an editable combo with a "…" button
Case "POIndex": .ComboList = "|..."
' ValidateEdit: the typed value must exist in the lookup table
Case "POIndex"
    Cancel = Not ValidateField(gItems, s, "PO Index not found", _
        "SELECT POIndex, Description FROM tblPOIndex WHERE DivisionID=" & div & _
        " and POIndex=" & DbQuote(Str, s), "POIndexDescription")
```

`MMain.ValidateField` contract (MMain.bas:998): **empty is always valid** and
blanks every companion column passed in the ParamArray; a hit **canonicalizes**
`EditText = rs(0)` (stored casing/format) and fills companion *i* from
`rs(i+1)` across `.Row..RowSel`; a miss shows the warning MsgBox and cancels.

## Blazor mapping — grid-level check + page-level companions

The membership check is **FlexKit-native** — three attributes on the column:

```razor
<GridColumn Field="@fieldName" ... AllowEditing="@allowEdit" ShowEditButton="true"
            RequireEditValueInList="@RequiresEditValueInList(fieldName)"
            EditValueList="@GetEditValueList(fieldName)"
            EditValueListProvider="@GetEditValueListProvider(fieldName)"
            EditValueNotFoundMessage="@GetEditValueNotFoundMessage(fieldName)" />
```

The library then supplies the whole VB6 envelope: empty-always-valid,
case-insensitive match canonicalized to the stored value, the red validation
pill (put the refused entry in the message with a `{0}` placeholder —
`"PO Index '{0}' not found"`), Enter on a miss flashes the text red ~2s then
discards and returns control (the MsgBox analog), the cursor stays on the
invalid cell, Escape always cancels, and the "…" button stays visible while
the editor is open (VB6 `ComboList "|..."`). Helpers return false/null for
non-validated fields, so one shared column branch stays safe.

**Companion fills are the PAGE's job**, in its OnCellSave handler — by the time
the save fires the code has already passed membership (or is blank), so a
preloaded map lookup is authoritative:

```csharp
case "POIndex":   // VB6 VF ParamArray "POIndexDescription"
    item.POIndexDescription = LookupMembershipDesc(_poIndexMembership, item.POIndex);
    break;
```

with `code → Description` dictionaries (OrdinalIgnoreCase) loaded once at
form load. Blank code ⇒ blank companions (the VF empty path). Mass-edit
type-ahead commits route through the same OnCellSave, so fills fan out too.

## Rules learned wiring FPricingWorkSheet + FItems (2026-08-30)

- **The membership list must mirror the "…" PICKER's query verbatim** — same
  table, same division filter. FItems' dropdown lists used `DivisionID = div`
  while its pickers used `IN (0, div)` (shared/global rows): membership built
  from the wrong one rejects values the picker itself offers.
- **Row-dependent lists use `EditValueListProvider`** (a sync
  `Func<object, IEnumerable<string>>` over a preloaded map — no awaits).
  FPricingWorkSheet CommunityPhase: valid phases depend on the ROW's current
  Community; a blank community accepts no phase (VB6 SQL parity).
- **Companion CASCADE on the key column**: a typed valid Community fills
  CommunityDesc **and blanks CommunityPhase** (VB6 VF SQL selects a literal
  `'' CommunityPhase`) — mirror whatever the PICKER-apply path writes so
  typed and picked entries end identically.
- **Wire only columns VB6 lets you TYPE.** A `ValidateField` call whose column
  is blocked in BeforeEdit (`Case Else: Cancel = True` — e.g. FPricingWorksheet
  Series has no BeforeEdit case) is dead code reachable only via VB6's paste
  path; wiring it is harmless but proves nothing.
- Messages come from the VB6 warning text plus the `{0}` entry:
  `"Community '{0}' not found"`, `"Cost Code '{0}' not found"`, …

Reference implementations: `FPricingWorkSheet.razor` (Community with
companion fill + row-aware CommunityPhase provider; Location/Series lists),
`FItems.razor` (POIndex/JCCostCode/JCCategory/TaxGroup/OptionCategory with
`*Desc` fills). The VB6 source of truth for each column's SQL, warning, and
companions is the form's `ValidateEdit`/`MYValidateEdit` handler.
