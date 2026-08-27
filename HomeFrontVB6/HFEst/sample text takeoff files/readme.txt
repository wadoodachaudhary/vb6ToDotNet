Text file takeoffs support csv or tab delimited files.

Column labels are optional. 
Unlabelled files with 3 columns are assumed to be Description, Qty, UOM.
All other unlabelled files are assumed to be Phase, Item, Description, Qty, UOM. Additional columns are permitted but ignored by HomeFront.

To be valid the file must include columns named:
"description" or "desc"
"qty" or "quantity" or "units"
"uom" or "unit"

All other columns are optional. If given the phase/item is validated against HomeFronts item database. The phase/item must exist and the uom must match either the order or takeoff uom. Quantities will be adjusted using the conversion factor as needed.

If phase/item is given all other columns (jcextra, jccostcode, jccategory, comments, rate) are retrieved from HomeFronts item database. Otherwise these values will be read from the import file (if they are available).