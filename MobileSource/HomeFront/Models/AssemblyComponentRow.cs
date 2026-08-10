namespace HomeFront.Models;

public class AssemblyComponentRow
{
    public long AssemblyId { get; set; }
    public string Community { get; set; } = "";
    public string Model { get; set; } = "";
    public string Option { get; set; } = "";
    public string Assembly { get; set; } = "";
    public string Description { get; set; } = "";
    public double Qty { get; set; }
    // Nullable: VB6 fills the Cost cell only when a pricing community is
    // selected (FAssembly.frm:3182) — otherwise the cell is BLANK, not $0.00.
    public double? Cost { get; set; }
    // HHM-581: true for rows picked while the parent assembly is still
    // unsaved (VB6 gComponents RowData="new") — flushed to
    // tblDBAssemblyComponents by FAssembly.SaveData once the master INSERT
    // has produced an AssemblyID, then cleared.
    public bool IsNew { get; set; }
}
