namespace HomeFront.Models;

/// <summary>
/// Represents a single column's layout/visibility state in a persisted grid.
/// Stored in the AppGridLayout database table.
/// </summary>
public sealed class GridLayoutColumn
{
    public string Field { get; set; } = "";
    public string Caption { get; set; } = "";
    public bool Hidden { get; set; }
    public double? Width { get; set; }
    public int? ColIndex { get; set; }
    public bool Grouped { get; set; }
    public bool IsEditable { get; set; }
}
