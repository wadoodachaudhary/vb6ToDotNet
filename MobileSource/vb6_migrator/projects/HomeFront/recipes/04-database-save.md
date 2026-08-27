# RECIPE: Database saving (references: FDBGrid.razor = the atomic gold standard; FAssembly.razor / FItems.razor = the sequential shape; FPriceList.razor = the applied example)

All data access goes through the injected `DbWrapperSqlServer db` with parameterized SQL (anonymous-object params: `new { item.Phase, item.Item }`). NEVER interpolate user data into SQL; identifiers (table/column names) can't be parameterized — pass them through `DbWrapperSqlServer.ValidateIdentifier`. No schema changes from app code, ever.

## Dirty tracking + Save gating
- Per-row state string (`RowState`): `"Insert"`/`"new"`, `"Update"`/`"DIRTY"`, `"Delete"`/`"Deleted"`, cleared to `""`/`"CLEAN"` per row as it lands. Form-level `mDirty`/`IsDirty` bool.
- **Save starts disabled, enabled only when dirty** (`disabled="@(!CanSave)"`, `CanSave => mDirty && !mReadOnly`) — the standing rule. (FItems' always-enabled Save is a documented VB6-parity exception, safe only because SaveDataAsync no-ops when clean.)
- Dirty-gate every exit path (view switch, tree activation, New/Export/Import, close) with the VB6 Yes/No/Cancel prompt: Cancel aborts the action, No proceeds without saving, Yes saves first and aborts on failure.

## Deletes: rows leave the grid IMMEDIATELY (HHM-259)
On delete (context menu, Ctrl+Delete): move the rows out of the grid's DataSource into a pending list (`List<TRow> _deletedRows`), rebind a fresh list reference, set dirty. The DELETE statements run on Save from the pending list. Marking RowState alone while the row stays visible reads as "delete does nothing".

## SaveData shape (sequential — FAssembly/FItems/FPriceList)
1. `if (!IsDirty && _deletedRows.Count == 0) return true;`
2. Optional prompt (above).
3. Validate + dup-key pre-check (`SELECT COUNT(*)` on the natural key) BEFORE any DML.
4. **Deletes first** (from the pending list), then inserts (`SELECT SCOPE_IDENTITY()` for new identity keys, write it back to the row), then updates.
5. Clear each row's state as it lands; `IsDirty = false` only after everything succeeded. On exception: log + `Notifications.Error`, return false — dirty stays true so retry re-saves.
6. Natural keys use VB6's NULL-tolerant match: `WHERE ISNULL(Community,'') = @Community AND ... AND DivisionID = {resolved}` — and respect scope rules (e.g. FPriceList: PriceLevel "Corporate (any Division)" saves against DivisionID 0).

## Atomic shape (FDBGrid — use for multi-statement saves that must not half-persist)
```csharp
// 1. PRE-VALIDATE every Insert row's key columns BEFORE ANY DML (a failed
//    mid-save otherwise half-persists — the 2026-08-01 data incident).
// 2. One transaction for the whole save; defer in-memory bookkeeping to post-commit:
await using var tx = await db.BeginTransactionAsync();
var postCommit = new List<Action>();
// deletes → inserts (IDENT_CURRENT for identity) → updates, all via tx.ExecuteAsync(sql, params)
await tx.CommitAsync();
foreach (var apply in postCommit) apply();   // clear RowStates, remove deleted rows
mDirty = false;
```
`DbTransactionSession` API: `tx.ExecuteAsync`, `tx.ExecuteScalarAsync<T>`, `tx.BulkCopyAsync(DataTable, table)`, `tx.CommitAsync()`; `await using` disposal ROLLS BACK when not committed — any exception skips commit and nothing persisted. Catch maps dup-key (SqlException 2601/2627) and NOT-NULL errors to friendly VB6-style messages.

## Bulk operations
Generate + validate ALL keys before the first INSERT (FItems bulk-copy: an overflowing later key must not leave a partially duplicated selection). Immediate-INSERT flows (add-item dialogs) save pending edits first (`SaveDataAsync(false)`), then INSERT, catching duplicates for the VB6 prompt.

## Options / config reads
- AppOptions: `WHERE DivisionID IN (0, @DivisionID) ORDER BY DivisionID` into a dict so the division row overrides global. Column names are `OptionName`/`OptionValue`.
- Per-user state (VB6 per-user INI) → `IUserPreferencesService`, NEVER AppOptions (cross-user leak) and never Blazor statics (shared across circuits).
- Field max lengths come from the DB (`CHARACTER_MAXIMUM_LENGTH` at load) — not hardcoded.
