using System.Data;
namespace BlazorDemos.Pages.Grid;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;
using BlazorDemos;
public class PhaseTreeData {
    public static List<PhaseTreeData> tree = new List<PhaseTreeData>();
    [Key] public int? NodeID{ get; set; }
    public string? GrpPhase{ get; set; }
    public string? GrpDesc{ get; set; }
    public string? GrpSortOrder{ get; set; }
    public string? Phase{ get; set; }
    public string? PhaseDesc{ get; set; }
    public string? PhaseSortOrder{ get; set; }
    public string? Node{ get; set; }
    public int? ParentID{ get; set; }
    public bool? IsParent{ get; set; }
    public int? ParentItem{ get; set; }
    public bool HasChildren { get; set; }
    public bool Expanded{ get; set; } = false;
    public int Level{ get; set; } = 0;

    public PhaseTreeData(){
    }

    public static List<PhaseTreeData> GetTree(){
        tree.Clear();
        DbWrapperSqlServer db = new();
        var itemsQuery = @"
                Select GrpPhase,GrpDesc,GrpSortOrder,Phase,PhaseDesc,PhaseSortOrder,Count(*) 
                from EstimatingItems i LEFT OUTER JOIN PriceGroups ON(PriceLink=PriceGroup and i.DivisionID = PriceGroups.divisionID)
                Group by GrpPhase,GrpDesc,GrpSortOrder,Phase,PhaseDesc,PhaseSortOrder
                Order by GrpSortOrder,PhaseSortOrder";
        Console.WriteLine($" ItemsQuery {itemsQuery}");
        //itemsQuery = itemsQuery.Replace("Select ", "Select top 50 ");
        var dt = db.SqlExec(itemsQuery, new {
            DivisionID = 1
        });
        int root = -1;
        int TreeID = 0;
        int ChildCount = -1;
        int SubTaskCount = -1;
        int t = 0;
        int rootItem = root + 1;
        var prevGrpPhase = "";
        var prevGrpDesc = "";
        var prevGrpSortOrder = "";
        var prevPhase = "";
        var prevPhaseDesc = "";
        var prevPhaseSortOrder = "";
        foreach (DataRow row in dt.Rows){
            // if (t > 10) break;
            var GrpPhase = row["GrpPhase"]?.ToString();
            var GrpDesc = row["GrpDesc"]?.ToString();
            var GrpSortOrder = row["GrpSortOrder"]?.ToString();
            var Phase = row["Phase"]?.ToString();
            var PhaseDesc = row["PhaseDesc"]?.ToString();
            var PhaseSortOrder = row["PhaseSortOrder"]?.ToString();
            ChildCount++;
            TreeID++;
            if (prevGrpPhase != GrpPhase){
                rootItem = TreeID;
                prevGrpPhase = GrpPhase;
                tree.Add(new PhaseTreeData() {
                    NodeID = rootItem, Node = GrpDesc, Phase=Phase,IsParent = true, ParentID = null, HasChildren = true});
                TreeID++;
            }
            tree.Add(new PhaseTreeData() {
                NodeID = TreeID, Node = PhaseDesc, Phase=Phase,IsParent = false, ParentID = rootItem, HasChildren = false});
        }
        return tree;
    }
}