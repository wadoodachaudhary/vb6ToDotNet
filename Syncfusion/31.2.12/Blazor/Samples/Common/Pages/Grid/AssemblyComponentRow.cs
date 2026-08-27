namespace BlazorDemos.Pages.Grid;
public class AssemblyComponentRow
{
    public long AssemblyId { get; set; }
    public string Community { get; set; } = "";
    public string Model { get; set; } = "";
    public string Option { get; set; } = "";
    public string Assembly { get; set; } = "";
    public string Description { get; set; } = "";
    public double Qty { get; set; }
    public double Cost { get; set; }
}