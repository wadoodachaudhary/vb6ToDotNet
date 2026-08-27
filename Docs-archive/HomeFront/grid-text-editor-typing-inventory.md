# Grid Text Editor Typing Inventory

Inventory date: 2026-08-20

Scope: all 71 `GridControl` instances in the 36 Razor pages under
`HomeFrontPB/Components/Pages`. `TreeGridControl` and standalone
`TextBoxControl`/`TextAreaControl` instances are outside this inventory.

## How the setting works

- FlexKit defaults `GridControl.TextEditorTypingBehavior` to `ServerBacked`.
- The setting applies to the built-in `TextBoxControl` batch editor. That editor
  is used for text, password, and number columns.
- Checkbox, dropdown (`EditOptions`/`EditOptionsProvider`), date, custom
  `EditTemplate`, and popup-only cells do not use this setting.
- No HomeFront grid explicitly declares `ServerBacked`; 61 grids get it by
  omitting `TextEditorTypingBehavior`.

## ClientBuffered grids (10)

| Page / grid | Built-in cells that actually use ClientBuffered | Cells not affected by the setting |
|---|---|---|
| `FAttributeLists.gData` | `Name` | None |
| `FAttributeLists.gValues` | `Value` (Description), `UpCharge` | `OtherListName` is a dropdown |
| `FDBGrid.DataGrid` | Every runtime column for which `IsFieldEditable` is true and whose resolved type is text/number/password, excluding picker and combo columns | Dynamic schema; checkbox/date/combo/picker columns are excluded |
| `FItems.gItems` | Lookup text cells `POIndex`, `JCCostCode`, `JCCategory`, `TaxGroup`, `AltJCCostCode`, `AltJCCategory`, `OptionCategory`, `OrderUOM`, `TakeoffUOM`; direct text cells `PartNumber`, `OptionID`, `Color`, `Location`, `WBS01`-`WBS40`; numeric cells `ConversionFactor`, `Price`, `WastePercent`, `RoundTo`, `RetailPretax` when editable | `ItemDesc`, `Notes`, `Formula` use popup editors; `RoundDir` is a dropdown; booleans and read-only columns are excluded |
| `FModelDimensions.gModels` | None: the grid is read-only, so the declaration is currently inert | `Model`, `Description` are read-only |
| `FModelDimensions.gCategories` | `Qty` | Room/category/UOM columns are read-only |
| `FProjectManagers.ManagersGrid` | `PM`, `Password`, `PMName`, `Phone`, `Cell`, `Fax`, `Email`, `POLimit`, `PrimaryUserID`, `SecondaryUserID` | `PrjMgr`, `Estimator`, `Purchaser`, `Active` are checkboxes |
| `FVendor.gContacts` | Every rendered non-checkbox/non-dropdown field; the fallback set is `Role`, `Name`, `Phone`, `Cell`, `Email`, `Fax` | `PONotice`, `FPONotice`, `SchedNotice`, `ServiceNotice` are checkboxes; `SendViaText`, `SmsAddress` are dropdowns |
| `FVendor.gPayPoints` | `PayPoint1Percent`-`PayPoint5Percent` | `POIndex` is a read-only picker cell |
| `FVendor.gInsurance` | `Company`, `PolicyNumber` | `Name` is read-only; `Req` is a checkbox; `ExpiryDate` is a date editor |

### FVendor Tax Groups anomaly

`FVendor.TaxGroupsGrid` contains a line reading
`TextEditorTypingBehavior="TextBoxTypingBehavior.ClientBuffered"` after the
opening `GridControl` tag has already closed. It is rendered as stray content,
not passed as a component parameter. The grid is intentionally non-editable
(`_taxGroupsEdit.AllowEditing = false`), so no Tax Groups cell currently uses
either typing transport.

## ServerBacked grids with built-in editors (28)

These grids omit the parameter and therefore use FlexKit's default
`ServerBacked` transport for the listed text/number/password cells.

| Page / grid | ServerBacked built-in editable cells |
|---|---|
| `FAddPricelist.ItemsGrid` | `Price`, `SKU` |
| `FAssembly.gComponents` | `Qty` |
| `FAssembly.gItems` | `AssemblyConversionFactor`, `OrderQty`, `TakeoffQty`, conditional `Price`, `Location`, `WBS01`-`WBS40` |
| `FCommunities.gData` | `Area`, `Description`, `Abr`, `Address1`, `Address2`, `City`, `Province`, `Postal`, `Country`, `County`, `Phone`, `Fax`, `Email`, `TaxNumber`, `IncomePrefix`, `BalanceSheet`, `Mortgage_Credit`, `BankAccount`, `TarionBuilderNumber`, `LotInventoryCredit`, `LabourTaxGroup`, `MaterialTaxGroup`, `SubContractTaxGroup`, `EquipmentTaxGroup`, `OverheadTaxGroup`, `OtherTaxGroup`, `SalesManagerEmail`, `IntacctEntity`, `IntacctParentJob`, `IntacctDepartment`, `GSTRate`, `PSTRate` |
| `FContacts.JobContactsGrid` | `Phone`, `Cell`, `Fax`, `Email` |
| `FCostForecast.ForecastGrid` | Runtime `Forecast1`-`ForecastN` numeric columns |
| `FCustomer.ContactsGrid` | `Name`, `Phone`, `Cell`, `Email`, `Fax` |
| `FInboxCustomQuote.DataGrid` | `Qty`, `UOM`, `Cost`, `Price`, `Status` when `CanEdit` |
| `FIntersection.IntersectionsGrid` | `Description` |
| `FIntersection.ItemsGrid` | `TakeoffQty`, `ConversionFactor`, `OrderQty`, `Notes` |
| `FItemChart.gData` | All runtime layout columns marked editable; fallback fields are `D1`, `D2`, `D3` |
| `FJob.JobContactsGrid` | `Name`, `Phone`, `Cell`, `Fax`, `Email`, `SendVia`, `SmsAddress` |
| `FJobCorrespondence.MessagesGrid` | `Subject`, `Body` when present in the runtime layout |
| `FMassChange.ItemsGrid` | `Quantity` |
| `FOptions` Divisions grid | `Code`, `Company` |
| `FOptions._epoGrid` | Runtime editable text fields; fallback fields are `ReasonCode`, `Category` |
| `FOptions._deptGrid` | `Name` |
| `FOptions` Approvers grid | `InvoiceLimit` |
| `FOptions._deptGLGrid` | `Account` |
| `FOptions._docClassGrid` | `DocumentClass` |
| `FPOFormats.DataGrid` | `Description`, `PayPoint1Percent`-`PayPoint4Percent` |
| `FPriceList.gItems` | Runtime fields whose `allowEdit` is true; fallback fields are `SKU`, `Current` |
| `FPricingWorkSheet.gData` | Runtime fields for which `allowEdit && !IsReadOnly` |
| `FPurchaseOrder._gridRef` | `Description`, `Comments`, `OrderQty`, `OrderUOM`, `Rate`, `TaxGroup` while unlocked |
| `FTakeoff.HFVariablesGrid` | `Value`, `UOM` |
| `FTakeoff.TLVariablesGrid` | `Value` |
| `FTakeoff.ItemsGrid` | `TakeoffQty`, `OrderQty`, `ConversionFactor`, `OverridePrice`, `Location` |
| `FUserPermissions._groupsGrid` | `Description` |

## Grids with no built-in text/number/password editor (33)

These grids do not currently exercise ClientBuffered or ServerBacked typing.
They are read-only, or their only editable cells are checkbox, dropdown, date,
or custom popup/template controls.

| Page | Grids / reason |
|---|---|
| `FAddPricelist` | `AssembliesGrid` read-only; `POIndexGrid` checkbox-only |
| `FAttachments` | `_grid`: only `DocumentClass` is editable, as a dropdown |
| `FCognito` | Groups and Users grids are read-only |
| `FContacts` | `ContactsGrid` (Personnel) has no inline-editable fields; `PickGrid` is read-only |
| `FDataExport` | `ExportGrid` is read-only |
| `FDataImport` | `DataGrid` is read-only |
| `FDefaultVendors` | `_grid` is read-only |
| `FDimensionCategories` | `gCategories` and `gRooms` are read-only; `gUnits.UOM` is a dropdown |
| `FImportAssemblies` | Assembly, Phase, and Category grids are read-only |
| `FImportExportAssemblies` | Assembly, Phase, Category, and Import Data grids are read-only |
| `FIntersection` | `AssembliesGrid` is read-only |
| `FJob` | Personnel `ContactsGrid` and `POsGrid` are read-only |
| `FMassChange` | `AssembliesGrid` is read-only |
| `FOptions` | Standard Items and Custom Fields grids are read-only |
| `FPickList` | `_grid` is read-only |
| `FSuperUser` | Client Databases grid is read-only |
| `FUserPermissions` | Users grid is read-only; Divisions and Sales Communities are checkbox-only |
| `FVendor` | `TaxGroupsGrid` is non-editable and uses picker buttons |
| `WorkflowEstimating` | `gJobs` is read-only |

## Totals

- 71 total `GridControl` instances.
- 10 declare `ClientBuffered`.
- 9 actually have at least one built-in editor that uses ClientBuffered.
- 61 inherit `ServerBacked`.
- 28 of those 61 have at least one built-in editor that uses ServerBacked.
- 33 have no built-in text/number/password editor, so the transport setting is
  not exercised.

Standalone note: `FJob` also opts three non-grid `TextAreaControl` instances
into ClientBuffered. They are deliberately not included in the 71-grid totals.
