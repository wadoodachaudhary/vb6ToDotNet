namespace HomeFront.Models;

public class AssemblyItemRow
{
    public int Sequence { get; set; }
    public string Phase { get; set; } = "";
    public string Item { get; set; } = "";
    public string ItemNumber { get; set; } = "";
    public string ItemChart { get; set; } = "";
    public string Description { get; set; } = "";
    public string TakeoffUOM { get; set; } = "";
    public string OrderUOM { get; set; } = "";
    public string Formula { get; set; } = "";
    public double TakeoffQty { get; set; }
    public double OrderQty { get; set; }
    public double WastePercent { get; set; }
    public string Notes { get; set; } = "";
    public string Location { get; set; } = "";
    public bool IsQuote { get; set; }
    public string ItemType { get; set; } = "Unit Price";
    public string PriceLevel { get; set; } = "";
    public double Price { get; set; }
    public double TaxRate { get; set; }
    public double TaxAmount { get; set; }
    public double PretaxAmount { get; set; }
    public double ExtendedAmount { get; set; }
    public string Vendor { get; set; } = "";
    public string VendorDesc { get; set; } = "";
    public string OldPriceLevel { get; set; } = "";
    public bool Invertable { get; set; }
    public bool UseModelCost { get; set; }
    public string POIndex { get; set; } = "";
    public string POIndexDescription { get; set; } = "";
    public string JCCostCode { get; set; } = "";
    public string JCCategory { get; set; } = "";
    public Dictionary<int, string> WBS { get; set; } = Enumerable.Range(1, 40).ToDictionary(i => i, i => "");
    public double AssemblyConversionFactor { get; set; }
    public double ItemConversionFactor { get; set; }
}
