# RECIPE: Tree grids (references: FPriceList.razor gAssemblies + FInboxJobs.razor; library = FlexKit TreeGridControl)

Use for every VB6 VSFlexGrid with outline rows (OutlineBar/IsCollapsed/GetNodeRow) → `TreeGridControl<TRow>`.

## Wiring
- Flat row list with `NodeID` / `ParentID` ints; bind `IdMapping="NodeID" ParentIdMapping="ParentID" TreeColumnIndex="0"`, `NodeExpandIconStyle="GroupExpandIconStyle.PlusMinus"`, `CssClass="fx-treegrid-compact fx-treegrid-dotted-lines"`, `EnableCollapseAll="true"` (arms the header ToggleExpandCollapse toolbar item — VB6 "Expand All").
- **Reference-equality trap:** TreeGridControl rebuilds its tree ONLY when the DataSource REFERENCE changes. After in-place Clear/Add/Insert, rebind with `Rows = new List<TRow>(Rows);` or the grid keeps showing the old build ("No records to display" bug).

## Lazy loading (VB6 LoadAssemblies dummy-child pattern)
- Build only the top level; give every expandable node ONE placeholder child (`IsPlaceholder = true`, empty text) so the [+] shows.
- On `Expanded`: if `node.IsPlaceholder || node.ChildrenLoaded` return; query the next level using the node's ACCUMULATED WhereClause (`"{parent.Where} AND {keyField} = '{value}'"`, single-quote-escaped); remove the placeholder; insert real children right after the parent (each with its own placeholder if a deeper level exists); set `ChildrenLoaded = true`; rebind with a fresh list reference.
- **Never filter empty-key rows** — an empty CommunityPhase/Assembly is the "- any -" ROLLUP node carrying the full child set; `WHERE Field = ''` is a valid filter.
- Ignore clicks/activation on placeholder rows.

## Selection vs activation (VB6-exact)
- `RowSelected` only moves the cursor/tracks state. Loading data + dirty-save prompts belong on **`RowActivated` (Enter) + `RowDoubleClicked`** — wiring the load to RowSelected fires a query (and a save prompt when dirty) on EVERY arrow-key move.

## Keyboard (built into TreeGridControl — do not re-implement)
Up/Down, Home/End, Enter (activate/toggle), ArrowRight (expand or first child), ArrowLeft (collapse or parent), **Space (toggle)**. Form-specific keys go through the `OnHostKeyDown` parameter (fires before built-ins) — e.g. FInboxJobs Delete clears the selected detail row's ChangeOrderNo (VB6 col 2 = c_Change_Order_No).

## Tri-state checkbox trees (FInboxJobs)
- Column 0 cell template hosts `CheckBoxControl` with `Checked`/`Indeterminate` + `StopClickPropagation`/`StopMouseDownPropagation`/`StopKeyDownPropagation` (so the row selection doesn't eat the click); a header-template checkbox drives check-all.
- Cascade: on change, set all descendants; walk ancestors setting Checked (all children checked) / Indeterminate (some) / Unchecked (none). Recursive helpers over the flat list by ParentNodeId.

## Node icons
VB6 imagelist bitmaps already live in `wwwroot/images/16/*.ico` — render with a plain `<img>` in the tree-column template, chosen by node type (Areas/Job/customer/Option) and warning quality (question/Warn/TakeoffCustom).

## MDI navigation
Every migrated page uses the `ParentNavigate` cascading parameter for navigation (`[CascadingParameter(Name = "ParentNavigate")] Action<string>?`), falling back to NavigationManager only when hosted standalone — direct NavigationManager calls break out of the FMain MDI shell and trigger the leave-page prompt.
