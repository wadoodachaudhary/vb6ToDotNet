using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using HomeFront.Data;
using Microsoft.AspNetCore.Components;
using Microsoft.Win32;
using Microsoft.JSInterop;

namespace HomeFront.Components.Pages;

public partial class MMain
{
    [Inject] public IJSRuntime JSRuntime { get; set; } = default!;
    [Inject] public NavigationManager NavigationManager { get; set; } = default!;
    // ═══════════════════════════════════════════════════════════════════════
    // Constants
    // ═══════════════════════════════════════════════════════════════════════
    public const int MARKUPPRECISION = 2;
    public const int mcFILE_SAVE = 1;
    public const int mcFILE_CHANGEDIV = 3;
    public const int mcFILE_CLOSE = 4;
    public const int mcVIEW_TASKPANEL = 0;
    public const int mcVIEW_WORKFLOW = 1;
    public const int mcGRID_GROUP = 0;
    public const int mcGRID_EXPANDALL = 1;
    public const int mcGRID_COLLAPSEALL = 2;
    public const int mcGRID_HIDE = 4;
    public const int mcGRID_INSERT = 5;
    public const int mcGRID_RENAME = 6;
    public const int mcGRID_PRINT = 8;
    public const int mcGRID_SAVEAS = 9;
    public const int mcGRID_ATTACHMENTS = 10;
    public const int mcPRICE_REFRESHCOSTS = 1;
    public const int mcPRICE_ADJUSTPRICES = 2;
    public const int mcPRICE_REMOVEASSEMBLY = 4;
    public const int BUFFER_SIZE = 32767;

    // ═══════════════════════════════════════════════════════════════════════
    // Enums
    // ═══════════════════════════════════════════════════════════════════════
    public enum EstimatingSystems
    {
        esNone = 0,
        esPipeline = 1,
        esTimberline = 2
    }

    public enum TakeoffSystems
    {
        tsNone = 0,
        tsPlanSwift = 1
    }

    public enum SortAlgorithms
    {
        InsertSort,
        MergeSort,
        QuickSort,
        SelectionSort
    }

    public enum OptionLocations
    {
        olModel,
        olScheduleB,
        olChangeOrder,
        olDesignCenter,
        olDeletedItems
    }

    public enum AssemblyStatuses
    {
        asOpen,
        asPriceOnly,
        asClosed
    }

    public enum AssemblyTypes
    {
        atModel = 0,
        atElevation = 1,
        atOption = 2,
        atGlobal = 3,
        atDesignCenter = 4,
        atCustomOption = -1
    }

    public enum OptionTypes
    {
        otModel = 0,
        otOption = 1,
        otGlobal = 2,
        otDesignCenter = 3
    }

    public enum DeliveryTypes
    {
        dtPrint,
        dtEmail,
        dtFax
    }

    public enum CostBasisTypes
    {
        cbCurrent,
        cbNext1,
        cbNext2,
        cbLast1,
        cbLast2,
        cbLast3,
        cbForecast1,
        cbForecast2,
        cbForecast3,
        cbForecast4,
        cbForecast5,
        cbForecast6,
        cbForecast7,
        cbForecast8,
        cbForecast9,
        cbForecast10,
        cbForecast11,
        cbForecast12
    }

    public enum ZeroQtyTakeoffModes
    {
        ztPrompt = 0,
        ztIgnore = 1,
        ztAccept = 2
    }

    public enum MultiStateEnum
    {
        msNone = 0,
        msSome = 1,
        msAll = 2
    }

    public enum PrinterStatusCodes
    {
        psReady = 0x0,
        psPaused = 0x1,
        psError = 0x2,
        psPendingDeletion = 0x4,
        psPaperJam = 0x8,
        psPaperOut = 0x10,
        psManualFeed = 0x20,
        psPaperProblem = 0x40,
        psOffline = 0x80,
        psIoActive = 0x100,
        psBusy = 0x200,
        psPrinting = 0x400,
        psOutputBinFull = 0x800,
        psNotAvailable = 0x1000,
        psWaiting = 0x2000,
        psProcessing = 0x4000,
        psInitializing = 0x8000,
        psWarmingUp = 0x10000,
        psTonerLow = 0x20000,
        psNoToner = 0x40000,
        psPagePrint = 0x80000,
        psUserIntervention = 0x100000,
        psOutOfMemory = 0x200000,
        psDoorOpen = 0x400000,
        psServerUnknown = 0x800000
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Public State (formerly VB6 module-level variables)
    // ═══════════════════════════════════════════════════════════════════════
    [Obsolete("Use ISessionStateService.ProjectBased via DI instead. Static state is shared across all users.")]
    public static bool ProjectBased { get; set; }
    [Obsolete("Use ISessionStateService.ProjectPhaseBased via DI instead. Static state is shared across all users.")]
    public static bool ProjectPhaseBased { get; set; }
    [Obsolete("Use ISessionStateService via DI instead. Static state is shared across all users.")]
    public static TakeoffSystems TakeoffSystem { get; set; }
    [Obsolete("Use ISessionStateService.IsMultiFamily via DI instead. Static state is shared across all users.")]
    public static bool IsMultiFamily { get; set; }
    [Obsolete("Use ISessionStateService via DI instead. Static state is shared across all users.")]
    public static bool PostPOQtyToAccounting { get; set; }
    [Obsolete("Static state is shared across all users — migrate to scoped service.")]
    public static string LoadedForm { get; set; } = "";
    [Obsolete("Use ISessionStateService.MaxPOAmount via DI instead. Static state is shared across all users.")]
    public static double MaxPOAmount { get; set; }

    // ═══════════════════════════════════════════════════════════════════════
    // Database
    // ═══════════════════════════════════════════════════════════════════════
    [Obsolete("Use @inject DbWrapperSqlServer via DI instead. Static instance is shared across all users.")]
    private static DbWrapperSqlServer db = new();
    [Obsolete("Use ISessionStateService.DivisionID via DI instead. Static state is shared across all users.")]
    private static int DivisionID = 1;
    [Obsolete("Use ISessionStateService.LoginID via DI instead. Static state is shared across all users.")]
    private static string LoginID = "";

    private static readonly object _dirLock = new();
    private static List<string> _dirList = new();
    private static int _dirIndex = -1;
    private static bool _dirSorted;

    [DllImport("mpr.dll", CharSet = CharSet.Unicode)]
    private static extern int WNetGetConnection(string localName, StringBuilder remoteName, ref int length);

    // The process-wide static `Application.HFApp` was REMOVED (it held per-user
    // DivisionID/LoginID in a static → shared across all circuits in Blazor Server,
    // a cross-tenant leak). These remaining consumers are all [Obsolete] static MMain
    // methods superseded by scoped DI services (IGridLayoutService, ITaskService); they
    // all null-guard (`if (app == null) return ...`, `HFApp?.x ?? fallback`), so this
    // alias now resolves to null and they degrade to their safe fallbacks instead of
    // leaking the last-logged-in user's state. Do not reintroduce the static.
    private static Application? HFApp => null;

    // ═══════════════════════════════════════════════════════════════════════
    // Lifecycle
    // ═══════════════════════════════════════════════════════════════════════
    public static async Task InitializeAsync()
    {
        await ReadUserPermissions();
    }

    // ═══════════════════════════════════════════════════════════════════════
    // ReadUserPermissions
    // ═══════════════════════════════════════════════════════════════════════
    public static async Task ReadUserPermissions()
    {
        try
        {
            var result = await db.ExecuteScalarAsync<double>(
                "SELECT MaxPOAmount FROM user_manager WHERE user_id=@LoginID",
                new { LoginID });
            MaxPOAmount = result;
        }
        catch { /* ignore */ }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // FormatExcelColRef
    // ═══════════════════════════════════════════════════════════════════════
    public static string FormatExcelColRef(long index)
    {
        const string abc = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        var s = "";
        if (index > 26)
        {
            var i = (int)((index - 1) / 26);
            if (i > 0) s += abc[i - 1];
        }
        var remainder = (int)(index % 26);
        s += remainder == 0 ? "Z" : abc[remainder - 1].ToString();
        return s;
    }

    // ═══════════════════════════════════════════════════════════════════════
    // PublishPricingWorksheet
    // ═══════════════════════════════════════════════════════════════════════
    public static async Task PublishPricingWorksheet(long worksheet)
    {
        try
        {
            var sql = new StringBuilder();
            sql.AppendLine("SELECT sd.assemblytype, sd.community, sd.CommunityPhase, sd.OptionID, sd.model");
            sql.AppendLine("      ,sd.JCExtra, sd.elevation, sd.series, sd.description, sd.AssemblyUOM");
            sql.AppendLine("      ,sd.floorarea, sd.bedrooms, sd.bathrooms, sd.style, sd.cost");
            sql.AppendLine("      ,sd.pretax, sd.copretax, sd.IncludeTax, sd.IncentiveCost, sd.IncentiveRetail");
            sql.AppendLine("      ,sd.tax, sd.cotax, sd.optionid, sd.constcutoff, sd.assembly, sd.category");
            sql.AppendLine("      ,sd.GraphicPath, sd.SpecDocument, sd.MaxWidth, sd.MaxLength, sd.Inactive");
            sql.AppendLine("      ,sd.color, sd.location, sd.qty, sd.includedoption");
            sql.AppendLine("      ,oc.group_code, sd.notes, sd.comments, sd.Markup, sd.Margin");
            for (int i = 2; i <= 10; i++)
            {
                sql.AppendLine($"      ,sd.pretax{i}, sd.tax{i}, sd.Margin{i}, sd.Markup{i}");
            }
            sql.AppendLine("      ,sd.ColorListID, sd.StyleListID, sd.FinishListID, sd.OtherListID");
            sql.AppendLine("      ,sd.StyleValue, sd.FinishValue, sd.OtherValue, m.Worksheet");
            sql.AppendLine("      ,sd.DesignCenterSalesOnly, sd.SelectByRoom, sd.DisplayTotalOnly");
            sql.AppendLine("  FROM tblSalesSheetDetails sd");
            sql.AppendLine("  JOIN tblSalesSheetMaster m ON m.Worksheet=sd.Worksheet");
            sql.AppendLine("  LEFT OUTER JOIN tblCategories oc ON sd.Category=oc.Category");
            sql.AppendLine($" WHERE sd.Worksheet=@Worksheet");

            var rows = await db.QueryAsync(sql.ToString(), new { Worksheet = worksheet });

            foreach (var rs in rows)
            {
                var assemblyType = (AssemblyTypes)Convert.ToInt32(rs["AssemblyType"] ?? 0);
                var community = rs["Community"]?.ToString() ?? "";
                var communityPhase = rs["CommunityPhase"]?.ToString() ?? "";
                var model = rs["Model"]?.ToString() ?? "";
                var series = rs["Series"]?.ToString() ?? "";
                var elevation = rs["Elevation"]?.ToString() ?? "";
                var optionID = rs["OptionID"]?.ToString() ?? "";
                var assembly = rs["Assembly"]?.ToString() ?? "";
                var description = rs["Description"]?.ToString() ?? "";
                var cost = ToDecimal(rs["Cost"]);
                var pretax = ToDecimal(rs["Pretax"]);
                var tax = ToDecimal(rs["Tax"]);
                var includeTax = ToBool(rs["IncludeTax"]);
                var margin = ToDecimal(rs["Margin"]);
                var markup = ToDecimal(rs["Markup"]);

                try
                {
                    switch (assemblyType)
                    {
                        case AssemblyTypes.atModel:
                            await PublishModel(rs, worksheet, community, communityPhase, model, series, elevation,
                                assembly, description, cost, pretax, tax, includeTax, margin, markup);
                            break;

                        case AssemblyTypes.atOption:
                            await PublishOption(rs, worksheet, community, communityPhase, model, series, elevation,
                                optionID, assembly, description, cost, pretax, tax, includeTax, margin, markup);
                            break;

                        case AssemblyTypes.atDesignCenter:
                            await PublishDesignCenter(rs, worksheet, community, communityPhase, optionID,
                                assembly, description, cost, pretax, tax, margin, markup);
                            break;

                        case AssemblyTypes.atGlobal:
                            await PublishGlobal(rs, worksheet, community, communityPhase, optionID,
                                assembly, description, cost, pretax, tax, includeTax, margin, markup);
                            break;
                    }
                }
                catch (Exception ex)
                {
                    if (!ex.Message.Contains("duplicate", StringComparison.OrdinalIgnoreCase))
                        throw;
                }
            }

            await db.ExecuteAsync(
                "UPDATE tblSalesSheetMaster SET SalesEffectiveDate=GETDATE() WHERE Worksheet=@Worksheet",
                new { Worksheet = worksheet });
            await db.ExecuteAsync(
                "UPDATE tblSalesSheetDetails SET ReadyToPublish=0 WHERE Worksheet=@Worksheet",
                new { Worksheet = worksheet });
        }
        catch (Exception ex)
        {
            Console.WriteLine($"PublishPricingWorksheet error: {ex.Message}");
        }
    }

    private static async Task PublishModel(Dictionary<string, object> rs, long worksheet,
        string community, string communityPhase, string model, string series, string elevation,
        string assembly, string description, decimal cost, decimal pretax, decimal tax,
        bool includeTax, decimal margin, decimal markup)
    {
        // Insert if not exists
        await db.ExecuteAsync(@"
            INSERT INTO tblModels(DivisionID,Area,CommunityPhase,Model,Series,Elevation,Inactive)
            SELECT @DivisionID,@Community,@CommunityPhase,@Model,@Series,@Elevation,0
            WHERE NOT EXISTS (
                SELECT 1 FROM tblModels WHERE DivisionID=@DivisionID
                AND ISNULL(Area,'')=@Community AND ISNULL(CommunityPhase,'')=@CommunityPhase
                AND ISNULL(Model,'')=@Model AND ISNULL(Series,'')=@Series AND ISNULL(Elevation,'')=@Elevation
            )",
            new { DivisionID, Community = community, CommunityPhase = communityPhase, Model = model, Series = series, Elevation = elevation });

        var total = includeTax ? tax + pretax : pretax;
        await db.ExecuteAsync(@"
            UPDATE tblModels SET
                LastSalesWorksheet=SalesWorksheet, SalesWorksheet=@Worksheet,
                Assembly=@Assembly, Description=@Description,
                model_picture=@GraphicPath, Spec_Document=@SpecDocument,
                Max_Width=@MaxWidth, Max_Length=@MaxLength, Inactive=@Inactive,
                Comments=@Comments, Style=@Style,
                NoOfBedrooms=@Bedrooms, NoOfBathrooms=@Bathrooms,
                ModelSize=@FloorArea, IncentiveCost=@IncentiveCost, IncentiveRetail=@IncentiveRetail,
                Cost_amount=@Cost, base_house=@Pretax, Margin=@Margin, Markup=@Markup,
                UseTax=@UseTax, NetTax=@NetTax, TotalAmount=@Total
            WHERE ISNULL(Area,'')=@Community AND ISNULL(CommunityPhase,'')=@CommunityPhase
                AND ISNULL(Model,'')=@Model AND ISNULL(Series,'')=@Series
                AND ISNULL(Elevation,'')=@Elevation AND DivisionID=@DivisionID",
            new
            {
                DivisionID, Worksheet = worksheet,
                Assembly = Truncate(assembly, 20), Description = Truncate(description, 50),
                GraphicPath = rs["GraphicPath"]?.ToString() ?? "", SpecDocument = rs["SpecDocument"]?.ToString() ?? "",
                MaxWidth = ToDecimal(rs["MaxWidth"]), MaxLength = ToDecimal(rs["MaxLength"]),
                Inactive = ToBool(rs["Inactive"]), Comments = rs["Comments"]?.ToString() ?? "",
                Style = rs["Style"]?.ToString() ?? "",
                Bedrooms = Truncate(rs["Bedrooms"]?.ToString() ?? "", 30),
                Bathrooms = Truncate(rs["Bathrooms"]?.ToString() ?? "", 30),
                FloorArea = ToDecimal(rs["FloorArea"]),
                IncentiveCost = ToDecimal(rs["IncentiveCost"]), IncentiveRetail = ToDecimal(rs["IncentiveRetail"]),
                Cost = cost, Pretax = pretax, Margin = margin, Markup = markup,
                UseTax = includeTax ? 1 : 0, NetTax = includeTax ? tax : 0m, Total = total,
                Community = community, CommunityPhase = communityPhase,
                Model = Truncate(model, 20), Series = Truncate(series, 30), Elevation = Truncate(elevation, 30)
            });
    }

    private static async Task PublishOption(Dictionary<string, object> rs, long worksheet,
        string community, string communityPhase, string model, string series, string elevation,
        string optionID, string assembly, string description, decimal cost, decimal pretax, decimal tax,
        bool includeTax, decimal margin, decimal markup)
    {
        await db.ExecuteAsync(@"
            INSERT INTO tblOptions(DivisionID,Area,CommunityPhase,Model,Series,Elevation,Opt,Option_Type,Option_Type_Desc,Inactive)
            SELECT @DivisionID,@Community,@CommunityPhase,@Model,@Series,@Elevation,@OptionID,1,'Sales Center',0
            WHERE NOT EXISTS (
                SELECT 1 FROM tblOptions WHERE DivisionID=@DivisionID
                AND ISNULL(Area,'')=@Community AND ISNULL(CommunityPhase,'')=@CommunityPhase
                AND Model=@Model AND Series=@Series AND Elevation=@Elevation AND Opt=@OptionID
            )",
            new { DivisionID, Community = community, CommunityPhase = communityPhase, Model = Truncate(model, 20), Series = Truncate(series, 30), Elevation = Truncate(elevation, 30), OptionID = Truncate(optionID, 20) });

        var coPretax = ToDecimal(rs["COPretax"]);
        var coTax = ToDecimal(rs["COTax"]);
        var coPrice = Math.Max(pretax, coPretax);

        await db.ExecuteAsync(@"
            UPDATE tblOptions SET
                LastSalesWorksheet=SalesWorksheet, SalesWorksheet=@Worksheet,
                Assembly=@Assembly, UOM=@UOM, Description=@Description,
                Color=@Color, ColorListID=@ColorListID, StyleListID=@StyleListID,
                FinishListID=@FinishListID, OtherListID=@OtherListID,
                Style=@StyleValue, Finish=@FinishValue, Other=@OtherValue,
                location=@Location, qty=@Qty, includedoption=@IncludedOption,
                Construction_Cut_Off=@ConstCutoff, Graphic_Path=@GraphicPath,
                Inactive=@Inactive, DesignCenterSalesOnly=@DCSalesOnly,
                SelectByRoom=@SelectByRoom, DisplayTotalOnly=@DisplayTotalOnly,
                Cost_Amount=@Cost, Category=@Category, Major_Group=@GroupCode,
                EstimatorNotes=@Notes, comments=@Comments, TL_Extra=@JCExtra,
                Price=@Pretax, CO_Price=@COPrice, Margin=@Margin, Markup=@Markup,
                UseTax=@UseTax, NetTax=@NetTax, TotalAmount=@Total,
                NetCOTax=@NetCOTax, TotalCOAmount=@TotalCOAmount
            WHERE ISNULL(Area,'')=@Community AND ISNULL(CommunityPhase,'')=@CommunityPhase
                AND Model=@Model AND Elevation=@Elevation AND Series=@Series
                AND Opt=@OptionID AND DivisionID=@DivisionID",
            new
            {
                DivisionID, Worksheet = worksheet,
                Assembly = Truncate(assembly, 20), UOM = Truncate(rs["AssemblyUOM"]?.ToString() ?? "", 10),
                Description = Truncate(description, 200),
                Color = Truncate(rs["color"]?.ToString() ?? "", 30),
                ColorListID = ToDecimal(rs["ColorListID"]), StyleListID = ToDecimal(rs["StyleListID"]),
                FinishListID = ToDecimal(rs["FinishListID"]), OtherListID = ToDecimal(rs["OtherListID"]),
                StyleValue = rs["StyleValue"]?.ToString() ?? "", FinishValue = rs["FinishValue"]?.ToString() ?? "",
                OtherValue = rs["OtherValue"]?.ToString() ?? "",
                Location = Truncate(rs["location"]?.ToString() ?? "", 50),
                Qty = ToDecimal(rs["qty"]), IncludedOption = ToBool(rs["IncludedOption"]),
                ConstCutoff = ToDecimal(rs["ConstCutoff"]),
                GraphicPath = rs["GraphicPath"]?.ToString() ?? "", Inactive = ToBool(rs["Inactive"]),
                DCSalesOnly = ToBool(rs["DesignCenterSalesOnly"]),
                SelectByRoom = ToBool(rs["SelectByRoom"]), DisplayTotalOnly = ToBool(rs["DisplayTotalOnly"]),
                Cost = cost, Category = Truncate(rs["Category"]?.ToString() ?? "", 15),
                GroupCode = Truncate(rs["Group_Code"]?.ToString() ?? "", 10),
                Notes = Truncate(rs["Notes"]?.ToString() ?? "", 4000),
                Comments = Truncate(rs["Comments"]?.ToString() ?? "", 4000),
                JCExtra = Truncate(rs["JCExtra"]?.ToString() ?? "", 10),
                Pretax = pretax, COPrice = coPrice, Margin = margin, Markup = markup,
                UseTax = includeTax ? 1 : 0, NetTax = includeTax ? tax : 0m,
                Total = includeTax ? tax + pretax : pretax,
                NetCOTax = includeTax ? coTax : 0m,
                TotalCOAmount = includeTax ? coTax + coPretax : coPretax,
                Community = community, CommunityPhase = communityPhase,
                Model = Truncate(model, 20), Elevation = Truncate(elevation, 30),
                Series = Truncate(series, 30), OptionID = Truncate(optionID, 20)
            });
    }

    private static async Task PublishDesignCenter(Dictionary<string, object> rs, long worksheet,
        string community, string communityPhase, string optionID,
        string assembly, string description, decimal cost, decimal pretax, decimal tax,
        decimal margin, decimal markup)
    {
        await db.ExecuteAsync(@"
            INSERT INTO tblDCOptions(DivisionID,Option_Type,Community,CommunityPhase,Opt,Inactive)
            SELECT @DivisionID,3,@Community,@CommunityPhase,@OptionID,0
            WHERE NOT EXISTS (
                SELECT 1 FROM tblDCOptions WHERE DivisionID=@DivisionID
                AND ISNULL(Community,'')=@Community AND ISNULL(CommunityPhase,'')=@CommunityPhase AND Opt=@OptionID
            )",
            new { DivisionID, Community = community, CommunityPhase = communityPhase, OptionID = Truncate(optionID, 20) });

        var updateSql = new StringBuilder();
        updateSql.AppendLine("UPDATE tblDCOptions SET");
        updateSql.AppendLine("    SalesWorksheet=@Worksheet, Assembly=@Assembly, TL_Extra=@JCExtra,");
        updateSql.AppendLine("    Description=@Description, UOM=@UOM, EstimatorNotes=@Notes, comments=@Comments,");
        updateSql.AppendLine("    Item1=@Pretax, Item1_Cost=@Cost, Color=@Color,");
        updateSql.AppendLine("    ColorListID=@ColorListID, StyleListID=@StyleListID,");
        updateSql.AppendLine("    FinishListID=@FinishListID, OtherListID=@OtherListID,");
        updateSql.AppendLine("    DesignCenterSalesOnly=@DCSalesOnly, SelectByRoom=@SelectByRoom,");
        updateSql.AppendLine("    DisplayTotalOnly=@DisplayTotalOnly,");
        updateSql.AppendLine("    Style=@StyleValue, Finish=@FinishValue, Graphic_Path=@GraphicPath,");
        updateSql.AppendLine("    Inactive=@Inactive, Other=@OtherValue,");
        updateSql.AppendLine("    location=@Location, qty=@Qty, includedoption=@IncludedOption,");
        updateSql.AppendLine("    Nettax1=@Tax, TotalAmount1=@TotalAmount1,");
        updateSql.AppendLine("    Margin1=@Margin, Markup1=@Markup, Category=@Category");
        for (int i = 2; i <= 10; i++)
        {
            updateSql.AppendLine($"    ,Margin{i}=@Margin{i}, Markup{i}=@Markup{i}");
            updateSql.AppendLine($"    ,Item{i}=@Pretax{i}, Item{i}_Cost=@Cost");
            updateSql.AppendLine($"    ,Nettax{i}=@Tax{i}, TotalAmount{i}=@TotalAmount{i}");
        }
        updateSql.AppendLine("WHERE ISNULL(Community,'')=@Community AND ISNULL(CommunityPhase,'')=@CommunityPhase");
        updateSql.AppendLine("  AND Opt=@OptionID AND DivisionID=@DivisionID");

        var parameters = new Dictionary<string, object?>
        {
            ["DivisionID"] = DivisionID, ["Worksheet"] = worksheet,
            ["Assembly"] = Truncate(assembly, 20), ["JCExtra"] = Truncate(rs["JCExtra"]?.ToString() ?? "", 10),
            ["Description"] = Truncate(description, 200), ["UOM"] = Truncate(rs["AssemblyUOM"]?.ToString() ?? "", 10),
            ["Notes"] = Truncate(rs["Notes"]?.ToString() ?? "", 4000),
            ["Comments"] = Truncate(rs["Comments"]?.ToString() ?? "", 4000),
            ["Pretax"] = pretax, ["Cost"] = cost,
            ["Color"] = Truncate(rs["color"]?.ToString() ?? "", 30),
            ["ColorListID"] = ToDecimal(rs["ColorListID"]), ["StyleListID"] = ToDecimal(rs["StyleListID"]),
            ["FinishListID"] = ToDecimal(rs["FinishListID"]), ["OtherListID"] = ToDecimal(rs["OtherListID"]),
            ["DCSalesOnly"] = ToBool(rs["DesignCenterSalesOnly"]),
            ["SelectByRoom"] = ToBool(rs["SelectByRoom"]), ["DisplayTotalOnly"] = ToBool(rs["DisplayTotalOnly"]),
            ["StyleValue"] = rs["StyleValue"]?.ToString() ?? "", ["FinishValue"] = rs["FinishValue"]?.ToString() ?? "",
            ["GraphicPath"] = rs["GraphicPath"]?.ToString() ?? "", ["Inactive"] = ToBool(rs["Inactive"]),
            ["OtherValue"] = rs["OtherValue"]?.ToString() ?? "",
            ["Location"] = Truncate(rs["location"]?.ToString() ?? "", 50),
            ["Qty"] = ToDecimal(rs["qty"]), ["IncludedOption"] = ToBool(rs["IncludedOption"]),
            ["Tax"] = tax, ["TotalAmount1"] = pretax + tax,
            ["Margin"] = margin, ["Markup"] = markup,
            ["Category"] = Truncate(rs["Category"]?.ToString() ?? "", 15),
            ["Community"] = community, ["CommunityPhase"] = communityPhase,
            ["OptionID"] = Truncate(optionID, 20)
        };
        for (int i = 2; i <= 10; i++)
        {
            var ptx = ToDecimal(rs[$"pretax{i}"]);
            var tx = ToDecimal(rs[$"tax{i}"]);
            parameters[$"Margin{i}"] = ToDecimal(rs[$"Margin{i}"]);
            parameters[$"Markup{i}"] = ToDecimal(rs[$"Markup{i}"]);
            parameters[$"Pretax{i}"] = ptx;
            parameters[$"Tax{i}"] = tx;
            parameters[$"TotalAmount{i}"] = ptx + tx;
        }

        await db.ExecuteAsync(updateSql.ToString(), parameters);
    }

    private static async Task PublishGlobal(Dictionary<string, object> rs, long worksheet,
        string community, string communityPhase, string optionID,
        string assembly, string description, decimal cost, decimal pretax, decimal tax,
        bool includeTax, decimal margin, decimal markup)
    {
        var category = rs["Category"]?.ToString() ?? "";
        await db.ExecuteAsync(@"
            INSERT INTO tblGlobalOptions(DivisionID,Option_Type,Community,CommunityPhase,Category,Opt,Inactive)
            SELECT @DivisionID,2,@Community,@CommunityPhase,@Category,@OptionID,0
            WHERE NOT EXISTS (
                SELECT 1 FROM tblGlobalOptions WHERE DivisionID=@DivisionID
                AND ISNULL(Community,'')=@Community AND ISNULL(CommunityPhase,'')=@CommunityPhase AND Opt=@OptionID
            )",
            new { DivisionID, Community = community, CommunityPhase = communityPhase, Category = Truncate(category, 15), OptionID = Truncate(optionID, 20) });

        var coPretax = ToDecimal(rs["COPretax"]);
        var coTax = ToDecimal(rs["COTax"]);
        var coPrice = Math.Max(pretax, coPretax);

        await db.ExecuteAsync(@"
            UPDATE tblGlobalOptions SET
                LastSalesWorksheet=SalesWorksheet, SalesWorksheet=@Worksheet,
                Assembly=@Assembly, UOM=@UOM, Description=@Description,
                Construction_Cut_Off=@ConstCutoff, Graphic_Path=@GraphicPath, Inactive=@Inactive,
                Cost_Amount=@Cost, DesignCenterSalesOnly=@DCSalesOnly,
                SelectByRoom=@SelectByRoom, DisplayTotalOnly=@DisplayTotalOnly,
                Category=@Category, TL_Extra=@JCExtra, Major_Group=@GroupCode,
                Color=@Color, ColorListID=@ColorListID, StyleListID=@StyleListID,
                FinishListID=@FinishListID, OtherListID=@OtherListID,
                Style=@StyleValue, Finish=@FinishValue, Other=@OtherValue,
                location=@Location, qty=@Qty, includedoption=@IncludedOption,
                EstimatorNotes=@Notes, comments=@Comments,
                Price=@Pretax, CO_Price=@COPrice, Margin=@Margin, Markup=@Markup,
                UseTax=@UseTax, NetTax=@NetTax, TotalAmount=@Total,
                NetCOTax=@NetCOTax, TotalCOAmount=@TotalCOAmount
            WHERE ISNULL(Community,'')=@Community AND ISNULL(CommunityPhase,'')=@CommunityPhase
                AND Opt=@OptionID AND DivisionID=@DivisionID",
            new
            {
                DivisionID, Worksheet = worksheet,
                Assembly = assembly, UOM = rs["AssemblyUOM"]?.ToString() ?? "",
                Description = description, ConstCutoff = ToDecimal(rs["ConstCutoff"]),
                GraphicPath = rs["GraphicPath"]?.ToString() ?? "", Inactive = ToBool(rs["Inactive"]),
                Cost = cost, DCSalesOnly = ToBool(rs["DesignCenterSalesOnly"]),
                SelectByRoom = ToBool(rs["SelectByRoom"]), DisplayTotalOnly = ToBool(rs["DisplayTotalOnly"]),
                Category = category, JCExtra = rs["JCExtra"]?.ToString() ?? "",
                GroupCode = rs["Group_Code"]?.ToString() ?? "",
                Color = rs["color"]?.ToString() ?? "",
                ColorListID = ToDecimal(rs["ColorListID"]), StyleListID = ToDecimal(rs["StyleListID"]),
                FinishListID = ToDecimal(rs["FinishListID"]), OtherListID = ToDecimal(rs["OtherListID"]),
                StyleValue = rs["StyleValue"]?.ToString() ?? "", FinishValue = rs["FinishValue"]?.ToString() ?? "",
                OtherValue = rs["OtherValue"]?.ToString() ?? "",
                Location = rs["location"]?.ToString() ?? "", Qty = ToDecimal(rs["qty"]),
                IncludedOption = ToBool(rs["IncludedOption"]),
                Notes = rs["Notes"]?.ToString() ?? "", Comments = rs["Comments"]?.ToString() ?? "",
                Pretax = pretax, COPrice = coPrice, Margin = margin, Markup = markup,
                UseTax = includeTax ? 1 : 0, NetTax = includeTax ? tax : 0m,
                Total = includeTax ? tax + pretax : pretax,
                NetCOTax = includeTax ? coTax : 0m,
                TotalCOAmount = includeTax ? coTax + coPretax : coPretax,
                Community = community, CommunityPhase = communityPhase, OptionID = Truncate(optionID, 20)
            });
    }

    // ═══════════════════════════════════════════════════════════════════════
    // LoadCostTypes
    // ═══════════════════════════════════════════════════════════════════════
    public static async Task<List<string>> LoadCostTypes(bool forecastsOnly = false)
    {
        var items = new List<string>();
        if (!forecastsOnly)
        {
            items.AddRange(new[] { "Current", "Next 1", "Next 2", "Last 1", "Last 2", "Last 3" });
        }

        var sql = new StringBuilder();
        for (int i = 1; i <= 12; i++)
        {
            if (i > 1) sql.AppendLine("UNION ALL");
            sql.AppendLine($"SELECT Item, MAX(Custom_Description), 'Forecast {i}' FROM CustomDescriptions WHERE Item='Forecast{i}' GROUP BY Item");
        }

        var rows = await db.QueryAsync(sql.ToString());
        foreach (var rs in rows)
        {
            var customDesc = rs.Values.ElementAtOrDefault(1)?.ToString() ?? "";
            var defaultDesc = rs.Values.ElementAtOrDefault(2)?.ToString() ?? "";
            items.Add(string.IsNullOrEmpty(customDesc) ? defaultDesc : customDesc);
        }
        return items;
    }

    // ═══════════════════════════════════════════════════════════════════════
    // RollPrices
    // ═══════════════════════════════════════════════════════════════════════
    public async Task RollPrices(bool prompt)
    {
        var sql = @"
            SELECT ISNULL(SUM(CASE WHEN next_effective1<=CONVERT(datetime,getdate()) AND ISNULL(next_cost1,0)=0 THEN 1 ELSE 0 END),0) ZeroRows1
                  ,ISNULL(SUM(CASE WHEN next_effective2<=CONVERT(datetime,getdate()) AND ISNULL(next_cost2,0)=0 THEN 1 ELSE 0 END),0) ZeroRows2
                  ,COUNT(*) AllRows
            FROM tblvendorcost
            WHERE (next_effective1<=CONVERT(datetime,getdate()) OR next_effective2<=CONVERT(datetime,getdate()))
              AND (divisionid=0 OR DivisionID=@DivisionID)";

        var rows = await db.QueryAsync(sql, new { DivisionID });
        if (!rows.Any()) return;

        var rs = rows.First();
        var zeroRows = Math.Max(Convert.ToInt64(rs["ZeroRows1"] ?? 0), Convert.ToInt64(rs["ZeroRows2"] ?? 0));
        var allRows = Convert.ToInt64(rs["AllRows"] ?? 0);

        if (allRows == 0)
        {
            if (!prompt)
                await JSRuntime.InvokeVoidAsync("alert", "No prices have expired. Your pricing database is current.");
            return;
        }

        bool includeZeros;
        if (zeroRows > 0)
        {
            includeZeros = await JSRuntime.InvokeAsync<bool>("confirm",
                $"Pricing has expired for {allRows} items. For {zeroRows} items the new price is $0.00. Set cost to $0.00?");
        }
        else
        {
            if (prompt)
            {
                var proceed = await JSRuntime.InvokeAsync<bool>("confirm",
                    $"Pricing has expired for {allRows} items. Update now?");
                if (!proceed) return;
            }
            includeZeros = true;
        }

        var updateSql = @"
            UPDATE tblVendorCost SET
                last_cost3=ISNULL(last_cost2,0), last3_expiry=last2_expiry,
                last_cost2=ISNULL(last_cost1,0), last2_expiry=last1_expiry,
                last_cost1=ISNULL(current_cost,0),
                last1_expiry=CAST(FLOOR(CAST(GETDATE() AS FLOAT)) AS DATETIME),
                Current_Cost=ISNULL(next_Cost1,0),
                next_cost1=ISNULL(next_cost2,0), next_effective1=next_effective2,
                next_cost2=0, next_effective2=NULL
            WHERE (next_effective1<=CAST(FLOOR(CAST(GETDATE() AS FLOAT)) AS DATETIME)
                OR next_effective2<=CAST(FLOOR(CAST(GETDATE() AS FLOAT)) AS DATETIME))
              AND (divisionid=0 OR DivisionID=@DivisionID)";

        if (!includeZeros)
            updateSql += " AND Next_cost1<>0";

        await db.ExecuteAsync(updateSql, new { DivisionID });
        await db.ExecuteAsync(updateSql, new { DivisionID }); // Run twice per VB6 logic

        var updatedCount = includeZeros ? allRows : allRows - zeroRows;
        await JSRuntime.InvokeVoidAsync("alert",
            $"{updatedCount} items have been updated. Your pricing database is current.");
    }

    // ═══════════════════════════════════════════════════════════════════════
    // FixGroupPhaseValue
    // ═══════════════════════════════════════════════════════════════════════
    public static async Task FixGroupPhaseValue()
    {
        var sql = @"
            UPDATE pp SET GroupPhaseValue=gg.phase
            FROM (
                SELECT p.divisionid, MAX(g.sortorder) groupsort, p.phase
                FROM tblestphases p
                JOIN estgroups g ON g.divisionid=p.divisionid AND g.sortorder<p.sortorder
                WHERE p.groupphase=0
                GROUP BY p.divisionid, p.sortorder, p.phase
            ) x
            JOIN tblestphases pp ON pp.divisionid=x.divisionid AND pp.phase=x.phase
            JOIN estgroups gg ON gg.divisionid=x.divisionid AND gg.sortorder=x.groupsort
            WHERE x.DivisionID=@DivisionID

            UPDATE tblestphases SET GroupPhaseValue=Phase
            WHERE DivisionID=@DivisionID AND GroupPhase=1

            UPDATE tblestphases SET groupphasevalue=(
                SELECT MIN(phase) FROM tblestphases WHERE divisionid=@DivisionID AND groupphase=1
            )
            WHERE groupphasevalue IS NULL";

        await db.ExecuteAsync(sql, new { DivisionID });
    }

    // ═══════════════════════════════════════════════════════════════════════
    // CancelPO
    // ═══════════════════════════════════════════════════════════════════════
    public async Task<bool> CancelPO(string poNumber, bool promptForCancelReason = true,
        int emailPrompt = 0, string cancelReason = "")
    {
        var reason = cancelReason;

        // Check if invoiced
        var amtInvoiced = await db.ExecuteScalarAsync<decimal>(
            "SELECT ISNULL(SUM(PreTax),0) FROM dbo.POInvoicedAmounts WHERE DivisionID=@DivisionID AND PONumber=@PONumber",
            new { DivisionID, PONumber = poNumber });

        if (amtInvoiced != 0)
        {
            await JSRuntime.InvokeVoidAsync("alert", $"Unable to cancel {poNumber}. It has invoices applied to it.");
            return false;
        }

        if (promptForCancelReason)
        {
            reason = await JSRuntime.InvokeAsync<string>("prompt", "Enter cancellation reason:");
            if (string.IsNullOrEmpty(reason)) return false;
        }

        // Get PO info
        var poRows = await db.QueryAsync(@"
            SELECT p.Vendor, ISNULL(NULLIF(p.DeliveryAddress,''),NULLIF(v.purchemail,'')) PurchEmail,
                   v.Vendor_Name, p.PostingBatch, p.Job
            FROM POMaster p
            LEFT OUTER JOIN tblVendors v ON p.Vendor=v.Vendor_ID AND p.DivisionID=v.DivisionID
            WHERE p.DivisionID=@DivisionID AND p.PONumber=@PONumber",
            new { DivisionID, PONumber = poNumber });

        if (!poRows.Any()) return false;

        var poInfo = poRows.First();
        var batch = Convert.ToInt64(poInfo["PostingBatch"] ?? 0);

        bool cancelled;
        if (batch != 0)
        {
            // Would delegate to accounting-system-specific cancellation
            cancelled = true; // Simplified
        }
        else
        {
            cancelled = true;
        }

        if (cancelled)
        {
            await db.ExecuteAsync(@"
                UPDATE POMaster SET Cancelled=1, CancelledBy=@LoginID,
                    CancelledNotes=@Reason, CancelledDate=GETDATE()
                WHERE DivisionID=@DivisionID AND PONumber=@PONumber",
                new { DivisionID, PONumber = poNumber, LoginID, Reason = reason });

            await db.ExecuteAsync(@"
                UPDATE EstimateItems SET POGenBatch=0, PONumber=''
                WHERE divisionid=@DivisionID AND PONumber=@PONumber",
                new { DivisionID, PONumber = poNumber });

            // Email notification
            bool sendEmail;
            switch (emailPrompt)
            {
                case 0:
                    sendEmail = await JSRuntime.InvokeAsync<bool>("confirm",
                        $"PO \"{poNumber}\" cancelled. Send email notification?");
                    break;
                case 1: sendEmail = true; break;
                default: sendEmail = false; break;
            }

            if (sendEmail)
            {
                // Email sending would be handled by a service
                Console.WriteLine($"Would send cancellation email for PO {poNumber}");
            }
        }

        return cancelled;
    }

    // ═══════════════════════════════════════════════════════════════════════
    // UpdateScheduleVendor
    // ═══════════════════════════════════════════════════════════════════════
    public static async Task UpdateScheduleVendor(string job, string oldVendorCsv, string newVendor)
    {
        try
        {
            var count = await db.ExecuteScalarAsync<int>(@"
                SELECT COUNT(*)
                FROM dbo.Schedule s
                JOIN dbo.ScheduleTasks t ON s.scheduleid=t.scheduleid
                JOIN dbo.RESERVATIONS r ON t.scheduletaskid=r.scheduletaskid
                WHERE s.job_no=@Job AND resourceid IN (@OldVendor)",
                new { Job = job, OldVendor = oldVendorCsv });

            if (count > 0)
            {
                // In a real implementation, would show a pick list dialog
                Console.WriteLine($"Would update {count} schedule tasks from vendor {oldVendorCsv} to {newVendor}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"UpdateScheduleVendor error: {ex.Message}");
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // CalcPreTax
    // ═══════════════════════════════════════════════════════════════════════
    public static async Task<double> CalcPreTax(AssemblyTypes assemblyType, double afterTaxPrice,
        string gstRate = "0", string pstRate = "0", string pstProvince = "", bool usePst = false, bool warn = false)
    {
        if (assemblyType == AssemblyTypes.atModel)
        {
            try
            {
                var pst = usePst ? Convert.ToDouble(pstRate) : 0;
                var result = await db.ExecuteScalarAsync<double>(@"
                    SELECT dbo.CalculatePretax(@AfterTaxPrice, '', @PstProvince, @GstRate, @PstRate)",
                    new { AfterTaxPrice = afterTaxPrice, PstProvince = pstProvince, GstRate = Convert.ToDouble(gstRate), PstRate = pst });
                return Math.Round(result, 2);
            }
            catch
            {
                return afterTaxPrice / (100 + Convert.ToDouble(gstRate) + Convert.ToDouble(pstRate)) * 100;
            }
        }
        return afterTaxPrice / (100 + Convert.ToDouble(gstRate) + Convert.ToDouble(pstRate)) * 100;
    }

    // ═══════════════════════════════════════════════════════════════════════
    // CalcTaxRebate
    // ═══════════════════════════════════════════════════════════════════════
    public static async Task<double> CalcTaxRebate(double pretaxPrice, string country = "")
    {
        if (country == "AU") return 0;
        try
        {
            var result = await db.ExecuteScalarAsync<double>(
                "SELECT dbo.CalculateFederalRebate(@Price) + dbo.CalculateProvincialRebate(@Price)",
                new { Price = pretaxPrice });
            return Math.Round(result, 2);
        }
        catch { return 0; }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // CalcTax
    // ═══════════════════════════════════════════════════════════════════════
    public static async Task<double> CalcTax(AssemblyTypes assemblyType, double pretaxPrice,
        double gstRate, double pstRate, bool usePst, string country = "")
    {
        var effectivePst = usePst ? pstRate : 0;
        var baseTax = Math.Round(pretaxPrice * (gstRate + effectivePst) / 100, 2);
        if (assemblyType == AssemblyTypes.atModel)
            return baseTax - await CalcTaxRebate(pretaxPrice, country);
        return baseTax;
    }

    // ═══════════════════════════════════════════════════════════════════════
    // String / Path Utility Functions
    // ═══════════════════════════════════════════════════════════════════════
    public static string StripFormatting(string s)
    {
        const string formatChars = "+=_-)(*&^%$#@!~`[]{}\\/|/?'\"<>;:,";
        foreach (var c in formatChars)
            s = s.Replace(c.ToString(), "");
        return s;
    }

    public static string Parse(string delimitedString, int index = 0, string delimiter = ",")
    {
        if (string.IsNullOrEmpty(delimitedString)) return index == 0 ? "0" : "";
        var parts = delimitedString.Split(new[] { delimiter }, StringSplitOptions.None);
        if (index == 0) return parts.Length.ToString();
        if (index < 0 || index > parts.Length) return "";
        return parts[index - 1];
    }

    public static int ParseCount(string delimitedString, string delimiter = ",")
    {
        if (string.IsNullOrEmpty(delimitedString)) return 0;
        return delimitedString.Split(new[] { delimiter }, StringSplitOptions.None).Length;
    }

    public static bool IsBetween(double value, double minimum, double maximum, bool inclusive = true)
    {
        return inclusive ? value >= minimum && value <= maximum : value > minimum && value < maximum;
    }

    public static double Between(double minimum, double maximum, double value)
    {
        return Math.Max(minimum, Math.Min(maximum, value));
    }

    public static T Min<T>(params T[] items) where T : IComparable<T>
    {
        return items.Min()!;
    }

    public static T Max<T>(params T[] items) where T : IComparable<T>
    {
        return items.Max()!;
    }

    public static bool ItemInList(string item, string list, string delimiter = ",")
    {
        return (delimiter + list + delimiter).Contains(delimiter + item + delimiter, StringComparison.OrdinalIgnoreCase);
    }

    public static string ParseDistinctList(string list, string delimiter = ",")
    {
        var parts = list.Split(new[] { delimiter }, StringSplitOptions.None);
        return string.Join(delimiter, parts.Distinct(StringComparer.OrdinalIgnoreCase));
    }

    public static bool IsIn(string expression, params string[] items)
    {
        return items.Any(i => string.Equals(expression, i, StringComparison.OrdinalIgnoreCase));
    }

    public static string FileTitle(string fullPath)
    {
        return Path.GetFileName(fullPath);
    }

    public static string CleanFileName(string fileName)
    {
        foreach (var c in new[] { '/', '\\', '*', '?', ':', '<', '>', '|' })
            fileName = fileName.Replace(c.ToString(), "");
        return fileName;
    }

    public static string FileName(string fullPath)
    {
        return Path.GetFileNameWithoutExtension(fullPath);
    }

    public static string FilePath(string fullPath)
    {
        return Path.GetDirectoryName(fullPath) ?? "";
    }

    public static string FileExt(string fullPath)
    {
        return Path.GetExtension(fullPath).TrimStart('.').ToLower();
    }

    public static string ForceExt(string fileName, string extension)
    {
        if (string.IsNullOrEmpty(extension)) return fileName;
        return Path.ChangeExtension(fileName, extension);
    }

    public static string StripExtension(string fileName)
    {
        return Path.GetFileNameWithoutExtension(fileName);
    }

    public static bool FileExists(string fileName)
    {
        return File.Exists(fileName);
    }

    public static bool PathExists(string path)
    {
        return Directory.Exists(path) || File.Exists(path);
    }

    public static void CreatePath(string pathDescription, string path)
    {
        try { Directory.CreateDirectory(path); } catch { }
    }

    public static string PathAppend(params string[] paths)
    {
        return Path.Combine(paths.Select(p => p.Trim().TrimEnd('\\').TrimStart('\\')).ToArray());
    }

    public static string TempPath()
    {
        var path = Path.Combine(Path.GetTempPath(), "HomeFront");
        Directory.CreateDirectory(path);
        return path;
    }

    public static string TempFile(string extension = "")
    {
        var fileName = Guid.NewGuid().ToString("N");
        if (!string.IsNullOrEmpty(extension))
            fileName = Path.ChangeExtension(fileName, extension);
        return Path.Combine(TempPath(), fileName);
    }

    public static string CreateGUID()
    {
        return Guid.NewGuid().ToString("B").ToUpper();
    }

    public static string MachineName()
    {
        return Environment.MachineName;
    }

    public static string NetworkLoginName()
    {
        return Environment.UserName;
    }

    public static string Quote(string s, bool removeFormatting = true, int length = 0, bool removeSpecial = true)
    {
        if (removeFormatting)
        {
            s = s.Replace("\t", " ").Replace("\n", " ").Replace("\r", " ");
        }
        if (removeSpecial)
            s = RemoveSpecialChar(s);
        s = s.Replace("\"", "''");
        if (length > 0) s = Truncate(s, length);
        return "\"" + s + "\"";
    }

    public static string AddQuotes(string s) => "\"" + s + "\"";

    private static string RemoveSpecialChar(string text)
    {
        var sb = new StringBuilder();
        foreach (var c in text)
        {
            if (c <= 0x7F) sb.Append(c);
        }
        return sb.ToString();
    }

    public static string SpaceCase(string s)
    {
        if (string.IsNullOrEmpty(s)) return s;
        s = s.Replace("_", "").Replace(" ", "");
        var result = new StringBuilder();
        for (int i = 0; i < s.Length; i++)
        {
            var cThis = s[i];
            var cPrev = i > 0 ? s[i - 1] : '\0';
            var cNext = i < s.Length - 1 ? s[i + 1] : 'Z';
            if (char.IsUpper(cThis) && (char.IsLower(cPrev) || char.IsLower(cNext)))
                result.Append(' ');
            result.Append(cThis);
        }
        return result.ToString().Trim();
    }

    public static string FormatPhone(string number, string countryCode = "US")
    {
        if (string.IsNullOrWhiteSpace(number)) return "";
        var digits = new string(number.Where(char.IsDigit).ToArray());
        if (!digits.All(char.IsDigit) || string.IsNullOrEmpty(digits)) return number;

        switch (countryCode.ToUpper())
        {
            case "AU":
                if (digits.Length == 11) digits = "+" + digits;
                if (digits.Length == 9) digits = "0" + digits;
                var isIntl = number.StartsWith("+");
                var isMobile = isIntl
                    ? IsIn(digits.Length > 3 ? digits.Substring(3, 1) : "", "4", "5")
                    : digits.Length >= 2 && IsIn(digits.Substring(0, 2), "04", "05");
                return (isMobile, isIntl) switch
                {
                    (true, true) => FormatWithMask(digits, "+### ### ### ###"),
                    (true, false) => FormatWithMask(digits, "#### ### ###"),
                    (false, true) => FormatWithMask(digits, "+### # #### ####"),
                    _ => FormatWithMask(digits, "## #### ####")
                };

            default:
                return digits.Length switch
                {
                    10 => $"({digits[..3]}) {digits[3..6]}-{digits[6..]}",
                    11 => $"{digits[0]} ({digits[1..4]}) {digits[4..7]}-{digits[7..]}",
                    7 => $"{digits[..3]}-{digits[3..]}",
                    _ => number
                };
        }
    }

    private static string FormatWithMask(string digits, string mask)
    {
        var result = new StringBuilder();
        int di = 0;
        foreach (var c in mask)
        {
            if (c == '#' || c == '@')
            {
                if (di < digits.Length) result.Append(digits[di++]);
            }
            else
            {
                result.Append(c);
            }
        }
        return result.ToString().Trim();
    }

    public static string CleanXML(string s)
    {
        return s; // Simplified — original had an early exit too
    }

    public static string Csv(params object[] values)
    {
        return string.Join(",", values);
    }

    public static string Update(string inputString, int col, string value, string delimiter = ",")
    {
        var parts = inputString.Split(new[] { delimiter }, StringSplitOptions.None).ToList();
        while (parts.Count < col) parts.Add("");
        parts[col - 1] = value;
        return string.Join(delimiter, parts);
    }

    public static string GetCustomDesc(string itemName, bool plural = false)
    {
        try
        {
            var result = db.ExecuteScalarAsync<string>(
                "SELECT ISNULL(NULLIF(custom_description,''),item) FROM customdescriptions WHERE item=@ItemName",
                new { ItemName = itemName }).Result;
            var s = result ?? itemName;
            if (plural)
            {
                if (s.EndsWith("y", StringComparison.OrdinalIgnoreCase))
                    s = s[..^1] + "ies";
                else if (!s.EndsWith("s", StringComparison.OrdinalIgnoreCase))
                    s += "s";
            }
            return s;
        }
        catch { return itemName; }
    }

    public static string SimpleEncrypt(string secret, string key)
    {
        var result = secret.ToCharArray();
        for (int i = 0; i < result.Length; i++)
        {
            var keyIndex = i % key.Length;
            result[i] = (char)(result[i] ^ key[keyIndex]);
        }
        return new string(result);
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Sort Utilities
    // ═══════════════════════════════════════════════════════════════════════
    public static void Sort(string[] strings, bool ascending = true, bool caseSensitive = false,
        SortAlgorithms algorithm = SortAlgorithms.QuickSort)
    {
        var comparer = caseSensitive ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase;
        Array.Sort(strings, comparer);
        if (!ascending) Array.Reverse(strings);
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Grid Layout Persistence (DB-based)
    // ═══════════════════════════════════════════════════════════════════════
    public static async Task<bool> DBGetGrid(string formName, string gridName, string? instanceKey = null)
    {
        var fullGridName = $"{formName}.{gridName}";
        if (!string.IsNullOrEmpty(instanceKey)) fullGridName += $".{instanceKey}";

        var rows = await db.QueryAsync(@"
            SELECT * FROM AppGridLayout
            WHERE uid=@LoginID AND gridname=@GridName",
            new { LoginID, GridName = fullGridName });

        return rows.Any();
    }

    [Obsolete("Use IGridLayoutService.SaveGridLayoutAsync() via DI instead.")]
    public static async Task DBPutGrid(string formName, string gridName, string? instanceKey,
        List<GridLayoutColumn> columns)
    {
        var fullGridName = $"{formName}.{gridName}";
        if (!string.IsNullOrEmpty(instanceKey)) fullGridName += $".{instanceKey}";

        await db.ExecuteAsync(@"
            DELETE AppGridLayout WHERE uid=@LoginID AND gridname=@GridName",
            new { LoginID, GridName = fullGridName });

        foreach (var col in columns)
        {
            await db.ExecuteAsync(@"
                INSERT INTO AppGridLayout(uid,gridname,split,colindex,colkey,caption,width,hidden,grouped)
                VALUES(@LoginID,@GridName,0,@ColIndex,@ColKey,@Caption,@Width,@Hidden,@Grouped)",
                new
                {
                    LoginID, GridName = fullGridName,
                    ColIndex = col.ColIndex,
                    ColKey = col.Field,
                    Caption = col.Caption,
                    Width = col.Width,
                    Hidden = col.Hidden,
                    Grouped = col.Grouped
                });
        }
    }


    // ═══════════════════════════════════════════════════════════════════════
    // SendingWizard / SyncCmd
    // ═══════════════════════════════════════════════════════════════════════
    public static string SyncCmd(string cmd, string parameters = "", bool skipAccounting = false, bool hfSend = false)
    {
        var exeName = hfSend ? "HFSend.exe" : "HFSync.exe";
        return $"{exeName} {cmd} {parameters}";
    }

    // ═══════════════════════════════════════════════════
    // Helpers
    // ═══════════════════════════════════════════════════
    private static decimal ToDecimal(object? value)
    {
        if (value == null || value == DBNull.Value) return 0m;
        if (value is decimal dec) return dec;
        if (value is double dbl) return Convert.ToDecimal(dbl);
        if (value is float flt) return Convert.ToDecimal(flt);
        if (value is int i32) return i32;
        if (value is long i64) return i64;
        var s = value.ToString();
        if (string.IsNullOrWhiteSpace(s)) return 0m;
        if (decimal.TryParse(s, out var parsed)) return parsed;
        return 0m;
    }

    private static bool ToBool(object? value)
    {
        if (value == null || value == DBNull.Value) return false;
        if (value is bool b) return b;
        if (value is int i) return i != 0;
        if (value is long l) return l != 0;
        if (value is short s16) return s16 != 0;
        if (value is byte b8) return b8 != 0;
        if (value is decimal dec) return dec != 0m;
        if (value is double dbl) return Math.Abs(dbl) > double.Epsilon;
        var s = value.ToString()?.Trim() ?? "";
        if (string.IsNullOrEmpty(s)) return false;
        if (string.Equals(s, "1", StringComparison.OrdinalIgnoreCase)) return true;
        if (string.Equals(s, "-1", StringComparison.OrdinalIgnoreCase)) return true;
        if (string.Equals(s, "true", StringComparison.OrdinalIgnoreCase)) return true;
        if (string.Equals(s, "yes", StringComparison.OrdinalIgnoreCase)) return true;
        if (string.Equals(s, "y", StringComparison.OrdinalIgnoreCase)) return true;
        if (string.Equals(s, "0", StringComparison.OrdinalIgnoreCase)) return false;
        if (string.Equals(s, "false", StringComparison.OrdinalIgnoreCase)) return false;
        if (string.Equals(s, "no", StringComparison.OrdinalIgnoreCase)) return false;
        if (string.Equals(s, "n", StringComparison.OrdinalIgnoreCase)) return false;
        if (bool.TryParse(s, out var parsedBool)) return parsedBool;
        if (int.TryParse(s, out var parsedInt)) return parsedInt != 0;
        return false;
    }

    private static string Truncate(string? value, int max)
    {
        if (string.IsNullOrEmpty(value)) return "";
        return value.Length <= max ? value : value.Substring(0, max);
    }
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

    [Obsolete("Use IGridLayoutService.GetGridLayoutAsync() via DI instead.")]
    public static async Task<Dictionary<string, GridLayoutColumn>> IniGetGridAsync(
        string formName,
        string gridName,
        int gridIndex = 0,
        string? instanceKey = null,
        bool grouped = false,
        bool staticNames = false,
        bool staticPositions = false,
        string? loginId = null)
    {
        var db = new DbWrapperSqlServer();
        return await DbGetGridAsync(db, formName, gridName, gridIndex, instanceKey, grouped, staticNames, staticPositions, loginId);
    }

    [Obsolete("Use IGridLayoutService.GetGridLayoutAsync() via DI instead.")]
    public static async Task<Dictionary<string, GridLayoutColumn>> DbGetGridAsync(
        DbWrapperSqlServer db,
        string formName,
        string gridName,
        int gridIndex,
        string? instanceKey = null,
        bool grouped = false,
        bool staticNames = false,
        bool staticPositions = false,
        string? loginId = null)
    {
        // loginId is passed by callers (the scoped IGridLayoutService passes _session.LoginID).
        // HFApp is now null (static removed), so this falls back to "admin" rather than the
        // last-logged-in user's LoginID — no cross-user leak.
        var uid = loginId ?? HFApp?.LoginID ?? "admin";
        var gridTag = gridIndex == 0 ? $"{formName}.{gridName}" : $"{formName}.{gridName}.({gridIndex})";

        if (!string.IsNullOrWhiteSpace(instanceKey))
            gridTag = $"{gridTag}.{instanceKey}";

        var sql = @"
            SELECT colkey, caption, width, hidden, colindex, grouped
            FROM AppGridLayout
            WHERE uid = @Uid AND gridname = @GridName
            ORDER BY colindex";

        var rows = await db.QueryAsync(sql, new { Uid = uid, GridName = gridTag });
        if (rows.Count == 0 && !string.Equals(uid, "", StringComparison.Ordinal))
        {
            var fallbackUid = "";
            rows = await db.QueryAsync(sql, new { Uid = fallbackUid, GridName = gridTag });
        }
        if (rows.Count == 0)
            return new Dictionary<string, GridLayoutColumn>(StringComparer.OrdinalIgnoreCase);

        var result = new Dictionary<string, GridLayoutColumn>(StringComparer.OrdinalIgnoreCase);
        var allHidden = true;

        foreach (var row in rows)
        {
            var field = row["colkey"]?.ToString() ?? "";
            if (string.IsNullOrEmpty(field))
                continue;

            var caption = row["caption"]?.ToString() ?? "";
            var hidden = ToBool(row["hidden"]);
            allHidden &= hidden;

            double? width = null;
            if (double.TryParse(row["width"]?.ToString(), out var widthVal))
                width = widthVal;

            int? colIndex = null;
            if (int.TryParse(row["colindex"]?.ToString(), out var idx))
                colIndex = idx;

            var groupedFlag = row.ContainsKey("grouped") && ToBool(row["grouped"]);

            result[field] = new GridLayoutColumn
            {
                Field = field,
                Caption = staticNames ? "" : caption,
                Hidden = hidden,
                Width = width,
                ColIndex = staticPositions ? null : colIndex,
                Grouped = groupedFlag
            };
        }

        if (allHidden)
            return new Dictionary<string, GridLayoutColumn>(StringComparer.OrdinalIgnoreCase);

        return result;
    }

    
    private enum DbQuoteType
    {
        Str,
        Num,
        Date
    }

    private static string DbQuote(DbQuoteType type, object? value, int length = 0)
    {
        switch (type)
        {
            case DbQuoteType.Str:
                var s = Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";
                if (length > 0 && s.Length > length) s = s[..length];
                return $"'{s.Replace("'", "''")}'";
            case DbQuoteType.Num:
                if (value == null || value == DBNull.Value) return "0";
                if (double.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Any, CultureInfo.InvariantCulture, out var d))
                    return d.ToString(CultureInfo.InvariantCulture);
                return "0";
            case DbQuoteType.Date:
                if (value is DateTime dt)
                    return $"'{dt:yyyy-MM-dd HH:mm:ss}'";
                if (DateTime.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), out var parsed))
                    return $"'{parsed:yyyy-MM-dd HH:mm:ss}'";
                return "NULL";
            default:
                return "NULL";
        }
    }

    private static string OptionValue(string name)
    {
        try { return HFApp?.Options?.ValueByName(name) ?? ""; } catch { return ""; }
    }

    private static int OptionInt(string name)
    {
        var s = OptionValue(name);
        return int.TryParse(s, out var v) ? v : 0;
    }

    private static object? TryInvokeResult(object? target, string methodName, params object?[] parameters)
    {
        if (target == null) return null;
        var type = target.GetType();
        var method = type.GetMethod(methodName);
        if (method == null) return null;
        try { return method.Invoke(target, parameters); } catch { return null; }
    }

    private static object? UnwrapTaskResult(object? maybeTask)
    {
        if (maybeTask is Task task)
        {
            try { task.GetAwaiter().GetResult(); } catch { }
            var resultProp = task.GetType().GetProperty("Result");
            return resultProp?.GetValue(task);
        }
        return maybeTask;
    }

    private static bool TryConvertDecimal(object? value, out decimal result)
    {
        result = 0m;
        if (value == null || value == DBNull.Value) return false;
        if (value is decimal d) { result = d; return true; }
        if (value is double dbl) { result = Convert.ToDecimal(dbl); return true; }
        if (value is float flt) { result = Convert.ToDecimal(flt); return true; }
        if (value is int i32) { result = i32; return true; }
        if (value is long i64) { result = i64; return true; }
        var s = value.ToString();
        return decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out result);
    }

    private static int CompareObjects(object? a, object? b)
    {
        if (a == null && b == null) return 0;
        if (a == null) return -1;
        if (b == null) return 1;

        if (TryConvertDecimal(a, out var da) && TryConvertDecimal(b, out var db))
            return da.CompareTo(db);

        if (a is IComparable cmp)
        {
            try { return cmp.CompareTo(b); } catch { }
        }

        var sa = a.ToString() ?? "";
        var sb = b.ToString() ?? "";
        return string.Compare(sa, sb, StringComparison.OrdinalIgnoreCase);
    }

    private static RegistryKey? OpenRegistryKey(object? hiveObj, string subKey, bool writable)
    {
        RegistryKey? root = null;
        if (hiveObj is RegistryHive hive)
        {
            root = RegistryKey.OpenBaseKey(hive, RegistryView.Default);
        }
        else if (hiveObj is int i)
        {
            root = i switch
            {
                unchecked((int)0x80000000) => Registry.ClassesRoot,
                unchecked((int)0x80000001) => Registry.CurrentUser,
                unchecked((int)0x80000002) => Registry.LocalMachine,
                unchecked((int)0x80000003) => Registry.Users,
                unchecked((int)0x80000005) => Registry.CurrentConfig,
                _ => Registry.LocalMachine
            };
        }
        else if (hiveObj is string s)
        {
            var name = s.ToUpperInvariant();
            root = name switch
            {
                "HKEY_CLASSES_ROOT" => Registry.ClassesRoot,
                "HKEY_CURRENT_USER" => Registry.CurrentUser,
                "HKEY_LOCAL_MACHINE" => Registry.LocalMachine,
                "HKEY_USERS" => Registry.Users,
                "HKEY_CURRENT_CONFIG" => Registry.CurrentConfig,
                _ => Registry.LocalMachine
            };
        }
        else
        {
            root = Registry.LocalMachine;
        }

        return writable ? root?.CreateSubKey(subKey) : root?.OpenSubKey(subKey);
    }

    private static IEnumerable<object> EnumerateGridColumns(object? grid)
    {
        if (grid == null) return Enumerable.Empty<object>();
        var prop = grid.GetType().GetProperty("Columns");
        if (prop?.GetValue(grid) is System.Collections.IEnumerable cols)
            return cols.Cast<object>();
        return Enumerable.Empty<object>();
    }

    private static IEnumerable<object> EnumerateGridRows(object? grid)
    {
        if (grid == null) return Enumerable.Empty<object>();
        var prop = grid.GetType().GetProperty("DataSource");
        if (prop?.GetValue(grid) is System.Collections.IEnumerable rows)
            return rows.Cast<object>();
        return Enumerable.Empty<object>();
    }

    private static bool ColumnVisible(object col)
    {
        var prop = col.GetType().GetProperty("Visible");
        if (prop?.GetValue(col) is bool b) return b;
        return true;
    }

    private static double? ColumnWidth(object col)
    {
        var runtimeProp = col.GetType().GetProperty("RuntimeWidth");
        if (runtimeProp?.GetValue(col) is double runtime && runtime > 0) return runtime;
        var widthProp = col.GetType().GetProperty("Width");
        var width = widthProp?.GetValue(col)?.ToString();
        if (string.IsNullOrWhiteSpace(width)) return null;
        var match = Regex.Match(width, @"-?\d+(\.\d+)?");
        if (match.Success && double.TryParse(match.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out var v))
            return v;
        return null;
    }

    private static int GridRowCount(object? grid)
    {
        return EnumerateGridRows(grid).Count();
    }

    private static int GridRowHeight(object? grid)
    {
        if (grid == null) return 24;
        var prop = grid.GetType().GetProperty("RowHeight");
        if (prop?.GetValue(grid) is int h && h > 0) return h;
        return 24;
    }

    private static bool TryGetCellTuple(object? tuple, out int row, out int col)
    {
        row = -1;
        col = -1;
        if (tuple == null) return false;
        var type = tuple.GetType();
        var rowProp = type.GetProperty("RowIndex") ?? type.GetProperty("Item1");
        var colProp = type.GetProperty("CellIndex") ?? type.GetProperty("Item2");
        if (rowProp == null || colProp == null) return false;
        try
        {
            row = Convert.ToInt32(rowProp.GetValue(tuple));
            col = Convert.ToInt32(colProp.GetValue(tuple));
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static object? GetArg(object?[] args, int index) => index < args.Length ? args[index] : null;

    private static string GetArgString(object?[] args, int index, string defaultValue = "")
    {
        var v = GetArg(args, index);
        return v?.ToString() ?? defaultValue;
    }

    private static bool GetArgBool(object?[] args, int index, bool defaultValue = false)
    {
        var v = GetArg(args, index);
        return v == null ? defaultValue : ToBool(v);
    }

    private static int? GetArgInt(object?[] args, int index)
    {
        var v = GetArg(args, index);
        if (v == null) return null;
        if (int.TryParse(v.ToString(), out var i)) return i;
        return null;
    }

    private static void TryInvoke(object? target, string methodName, params object?[] parameters)
    {
        if (target == null) return;
        var type = target.GetType();
        var method = type.GetMethod(methodName);
        if (method != null)
        {
            try { method.Invoke(target, parameters); } catch { }
        }
    }

    private static object? TryGetProperty(object? target, string propertyName)
    {
        if (target == null) return null;
        var prop = target.GetType().GetProperty(propertyName);
        if (prop == null) return null;
        try { return prop.GetValue(target); } catch { return null; }
    }

    private static void TrySetProperty(object? target, string propertyName, object? value)
    {
        if (target == null) return;
        var prop = target.GetType().GetProperty(propertyName);
        if (prop == null || !prop.CanWrite) return;
        try { prop.SetValue(target, value); } catch { }
    }

    private static string ResolveIniFile(object? fileArg)
    {
        var file = fileArg?.ToString();
        if (!string.IsNullOrWhiteSpace(file)) return file;
        return Path.Combine(AppContext.BaseDirectory, "HomeFront.ini");
    }

    private static Dictionary<string, Dictionary<string, string>> LoadIniFile(string path)
    {
        var data = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
        if (!File.Exists(path)) return data;
        string? current = null;
        foreach (var raw in File.ReadAllLines(path))
        {
            var line = raw.Trim();
            if (line.Length == 0 || line.StartsWith(";") || line.StartsWith("'"))
                continue;
            if (line.StartsWith("[") && line.EndsWith("]"))
            {
                current = line[1..^1].Trim();
                if (!data.ContainsKey(current))
                    data[current] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                continue;
            }
            var idx = line.IndexOf('=');
            if (idx <= 0) continue;
            var key = line[..idx].Trim();
            var value = line[(idx + 1)..].Trim();
            if (current == null)
                current = "";
            if (!data.ContainsKey(current))
                data[current] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            data[current][key] = value;
        }
        return data;
    }

    private static void SaveIniFile(string path, Dictionary<string, Dictionary<string, string>> data)
    {
        var lines = new List<string>();
        foreach (var section in data)
        {
            if (!string.IsNullOrEmpty(section.Key))
                lines.Add($"[{section.Key}]");
            foreach (var kv in section.Value)
            {
                lines.Add($"{kv.Key}={kv.Value}");
            }
            lines.Add("");
        }
        Directory.CreateDirectory(Path.GetDirectoryName(path) ?? ".");
        File.WriteAllLines(path, lines);
    }

    private static string BuildSectionName(object? formObj, object? instanceKey)
    {
        var name = formObj?.GetType().Name ?? "Form";
        var key = instanceKey?.ToString();
        return string.IsNullOrWhiteSpace(key) ? name : $"{name}.{key}";
    }

    private static void SortStringsInPlace(object? target)
    {
        if (target == null) return;
        if (target is string[] arr)
        {
            Array.Sort(arr, StringComparer.OrdinalIgnoreCase);
            return;
        }
        if (target is List<string> list)
        {
            list.Sort(StringComparer.OrdinalIgnoreCase);
            return;
        }
    }

// -------------------------------------------------------------------
    // VB6 parity stubs (ported)
    // -------------------------------------------------------------------
    public static object? CancelPO_MB(params object?[] args)
    {
        var poNumber = GetArgString(args, 0);
        var reason = GetArgString(args, 1);
        if (string.IsNullOrWhiteSpace(poNumber)) return false;
        var app = HFApp;
        if (app == null) return false;

        try
        {
            var divisionId = int.TryParse(app.DivisionID, out var div) ? div : DivisionID;
            var s = new StringBuilder();
            s.AppendLine("select p.PONumber,p.POIndex");
            s.AppendLine("  from POMaster p");
            s.AppendLine($" Where p.DivisionID = {divisionId}");
            s.AppendLine($"   and p.ponumber = {DbQuote(DbQuoteType.Str, poNumber)}");
            var dt = app.SqlExec(s.ToString(), Application.Connections.dbHomefront);
            if (dt.Rows.Count == 0)
            {
                Console.WriteLine($"CancelPO_MB: PO number \"{poNumber}\" not found.");
                return false;
            }

            s.Clear();
            s.AppendLine("select recnum,'po' doctype,rcvdte amtpd");
            s.AppendLine("from pchord");
            s.AppendLine($"where ordnum = {DbQuote(DbQuoteType.Str, poNumber)}");
            s.AppendLine("union");
            s.AppendLine("select recnum,'contract' doctype,invttl amtpd");
            s.AppendLine("from subcon");
            s.AppendLine($"where ctcnum = {DbQuote(DbQuoteType.Str, poNumber)}");
            var acct = app.SqlExec(s.ToString(), Application.Connections.dbAccounting);
            if (acct.Rows.Count == 0)
            {
                Console.WriteLine($"CancelPO_MB: PO number \"{poNumber}\" not found in accounting.");
                return false;
            }

            var row = acct.Rows[0];
            var mbObjectId = Convert.ToInt64(row["recnum"] == DBNull.Value ? 0 : row["recnum"]);
            var docType = (row["doctype"]?.ToString() ?? "").Trim();
            var amtPd = Convert.ToDouble(row["amtpd"] == DBNull.Value ? 0 : row["amtpd"], CultureInfo.InvariantCulture);
            if (Math.Abs(amtPd) > 0.0001)
            {
                Console.WriteLine($"CancelPO_MB: PO {poNumber} has invoices; cannot cancel.");
                return false;
            }

            var xml = app.XmlMbStart(OptionValue("MasterBuilderCompany"), OptionValue("MasterBuilderUID"));
            if (string.Equals(docType, "po", StringComparison.OrdinalIgnoreCase))
            {
                xml += "<PurchaseOrderDelRq requestID=\"1\">" + Environment.NewLine;
                xml += "<ObjectRef>" + Environment.NewLine;
                xml += $"<ObjectID>{mbObjectId}</ObjectID>" + Environment.NewLine;
                xml += "</ObjectRef>" + Environment.NewLine;
                xml += "</PurchaseOrderDelRq>" + Environment.NewLine;
            }
            else
            {
                xml += "<SubcontractDelRq requestID=\"1\">" + Environment.NewLine;
                xml += "<ObjectRef>" + Environment.NewLine;
                xml += $"<ObjectID>{mbObjectId}</ObjectID>" + Environment.NewLine;
                xml += "</ObjectRef>" + Environment.NewLine;
                xml += "</SubcontractDelRq>" + Environment.NewLine;
            }

            xml += app.XmlMBEnd();
            app.XmlMbSubmit(xml, OptionValue("MasterBuilderPWD"));
            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"CancelPO_MB error: {ex.Message}. Reason={reason}");
            return false;
        }
    }

    public static object? CancelPO_QB(params object?[] args)
    {
        var poNumber = GetArgString(args, 0);
        if (string.IsNullOrWhiteSpace(poNumber)) return false;
        var app = HFApp;
        if (app == null) return false;

        try
        {
            var divisionId = int.TryParse(app.DivisionID, out var div) ? div : DivisionID;
            var s = $"select * from pomaster where DivisionID = {divisionId} and ponumber={DbQuote(DbQuoteType.Str, poNumber)}";
            var dt = app.SqlExec(s, Application.Connections.dbHomefront);
            if (dt.Rows.Count == 0)
            {
                Console.WriteLine($"CancelPO_QB: PO number \"{poNumber}\" not found.");
                return false;
            }

            var poDate = dt.Rows[0]["PODate"];
            var xml = app.XmlQBStart();
            xml += "<GeneralDetailReportQueryRq>" + Environment.NewLine;
            xml += "<GeneralDetailReportType>OpenPOs</GeneralDetailReportType>" + Environment.NewLine;
            xml += "<ReportPeriod>" + Environment.NewLine;
            xml += app.XmlQBAdd(Application.XMLFieldTypes.d, 0, "FromReportDate", poDate);
            xml += app.XmlQBAdd(Application.XMLFieldTypes.d, 0, "ToReportDate", poDate);
            xml += "</ReportPeriod>" + Environment.NewLine;
            xml += "<IncludeColumn>RefNumber</IncludeColumn>" + Environment.NewLine;
            xml += "<IncludeColumn>TxnID</IncludeColumn>" + Environment.NewLine;
            xml += "<IncludeColumn>Amount</IncludeColumn>" + Environment.NewLine;
            xml += "<IncludeColumn>Name</IncludeColumn>" + Environment.NewLine;
            xml += "<IncludeColumn>Date</IncludeColumn>" + Environment.NewLine;
            xml += "</GeneralDetailReportQueryRq>" + Environment.NewLine;
            xml += app.XmlQBEnd();
            var response = app.XmlQBSubmit(xml);

            var q = "exec qb_POtxnid " + DbQuote(DbQuoteType.Str, response) + ", " + DbQuote(DbQuoteType.Str, poNumber);
            var txn = app.SqlExec(q, Application.Connections.dbHomefront);
            if (txn.Rows.Count == 0)
            {
                Console.WriteLine($"CancelPO_QB: PO {poNumber} not found in QuickBooks.");
                return false;
            }

            var txnId = txn.Rows[0][0]?.ToString() ?? "";
            if (string.IsNullOrWhiteSpace(txnId))
                return false;

            xml = app.XmlQBStart();
            xml += "<PurchaseOrderQueryRq>" + Environment.NewLine;
            xml += "<TxnID>" + "</TxnID>" + Environment.NewLine;
            xml += app.XmlQBAdd(Application.XMLFieldTypes.st, 40, "TxnID", txnId);
            xml += "<IncludeLineItems>true</IncludeLineItems>" + Environment.NewLine;
            xml += "</PurchaseOrderQueryRq>" + Environment.NewLine;
            xml += app.XmlQBEnd();
            var detail = app.XmlQBSubmit(xml);
            var editSeq = Parse(Parse(detail, 2, "EditSequence>"), 1, "<");

            xml = app.XmlQBStart();
            xml += "<PurchaseOrderModRq>" + Environment.NewLine;
            xml += "   <PurchaseOrderMod>" + Environment.NewLine;
            xml += app.XmlQBAdd(Application.XMLFieldTypes.st, 40, "TxnID", txnId);
            xml += app.XmlQBAdd(Application.XMLFieldTypes.st, 40, "EditSequence", editSeq);
            xml += "      <IsManuallyClosed> 1 </IsManuallyClosed>" + Environment.NewLine;
            xml += app.XmlQBAdd(Application.XMLFieldTypes.st, 500, "Memo", $"closed by {app.LoginID}, {DateTime.Now}");
            xml += "   </PurchaseOrderMod>" + Environment.NewLine;
            xml += "</PurchaseOrderModRq>" + Environment.NewLine;
            xml += app.XmlQBEnd();
            app.XmlQBSubmit(xml);
            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"CancelPO_QB error: {ex.Message}");
            return false;
        }
    }

    public static object? CancelPO_TL(params object?[] args)
    {
        var poNumber = GetArgString(args, 0);
        if (string.IsNullOrWhiteSpace(poNumber)) return false;
        var app = HFApp;
        if (app == null) return false;

        try
        {
            var divisionId = int.TryParse(app.DivisionID, out var div) ? div : DivisionID;
            app.WriteAuditLog("Edit PO", $"Cancel posted PO \"{poNumber}\"");

            decimal tlco = 0m;
            decimal hfco = 0m;
            var dt = app.SqlExec(
                "select max(commitment_co) from jcm_master__commitment_co where commitment_co<'A' and Commitment=@PoNumber",
                new { PoNumber = poNumber },
                Application.Connections.dbAccounting);
            if (dt.Rows.Count > 0) tlco = ToDecimal(dt.Rows[0][0]);
            dt = app.SqlExec(
                "select max(ChangeOrder) from pochangeorders where changeorder<'A' and ponumber=@PoNumber",
                new { PoNumber = poNumber },
                Application.Connections.dbHomefront);
            if (dt.Rows.Count > 0) hfco = ToDecimal(dt.Rows[0][0]);
            var co = (int)(tlco > hfco ? tlco : hfco) + 1;

            var s = new StringBuilder();
            s.AppendLine("INSERT INTO POChangeOrders(PONumber,DivisionID,ChangeOrder,Description,CODate,UStmp,TStmp,PostingBatch)");
            s.AppendLine("SELECT PONumber");
            s.AppendLine($"      ,{divisionId}");
            s.AppendLine($"      ,{DbQuote(DbQuoteType.Str, co)}");
            s.AppendLine("      ,'po cancelled'");
            s.AppendLine("      ,GETDATE()");
            s.AppendLine($"      ,{DbQuote(DbQuoteType.Str, app.LoginID)}");
            s.AppendLine("      ,GETDATE()");
            s.AppendLine("      ,0");
            s.AppendLine("  FROM POMaster");
            s.AppendLine($" WHERE DivisionID = {divisionId} and PONumber={DbQuote(DbQuoteType.Str, poNumber)}");
            app.SqlExec(s.ToString(), Application.Connections.dbHomefront);

            s.Clear();
            s.AppendLine("SELECT");
            s.AppendLine(" description");
            s.AppendLine(",units + approved_commitment_co_units - units_invoiced as Qty");
            s.AppendLine(",unit_description UOM");
            s.AppendLine(",unit_cost Rate");
            s.AppendLine(",amount-amount_invoiced-tax-approved_commitment_co_tax_amount Pretax");
            s.AppendLine(",tax_group TaxGroup");
            s.AppendLine(",tax+approved_commitment_co_tax_amount jctax");
            s.AppendLine(",Job");
            s.AppendLine(",extra JCExtra");
            s.AppendLine(",cost_code JCCostCode");
            s.AppendLine(",category JCCategory");
            s.AppendLine(",item_number LineNumber");
            s.AppendLine("FROM JCM_Master__Commitment_Item");
            s.AppendLine($"WHERE Commitment={DbQuote(DbQuoteType.Str, poNumber)}");
            s.AppendLine("ORDER BY item_number");
            var items = app.SqlExec(s.ToString(), Application.Connections.dbAccounting);
            foreach (DataRow row in items.Rows)
            {
                var insert = new StringBuilder();
                insert.AppendLine("insert into pochangeorderitems(DivisionID, PONumber, ChangeOrder, Description, Qty, UOM, Rate, Pretax, TaxGroup, JCTax, Job, JCExtra, JCCostCode, JCCategory, LineNumber)");
                insert.AppendLine("values(" + DbQuote(DbQuoteType.Num, divisionId));
                insert.AppendLine("      ," + DbQuote(DbQuoteType.Str, poNumber));
                insert.AppendLine("      ," + DbQuote(DbQuoteType.Str, co));
                insert.AppendLine("      ," + DbQuote(DbQuoteType.Str, row["Description"]));
                insert.AppendLine("      ," + DbQuote(DbQuoteType.Num, -1 * ToDecimal(row["Qty"])));
                insert.AppendLine("      ," + DbQuote(DbQuoteType.Str, row["UOM"]));
                insert.AppendLine("      ," + DbQuote(DbQuoteType.Num, ToDecimal(row["Rate"])));
                insert.AppendLine("      ," + DbQuote(DbQuoteType.Num, -1 * ToDecimal(row["Pretax"])));
                insert.AppendLine("      ," + DbQuote(DbQuoteType.Str, row["TaxGroup"]));
                insert.AppendLine("      ," + DbQuote(DbQuoteType.Num, -1 * ToDecimal(row["jctax"])));
                insert.AppendLine("      ," + DbQuote(DbQuoteType.Str, StripFormatting(row["Job"]?.ToString() ?? "")));
                insert.AppendLine("      ," + DbQuote(DbQuoteType.Str, row["JCExtra"]));
                insert.AppendLine("      ," + DbQuote(DbQuoteType.Str, row["JCCostCode"]));
                insert.AppendLine("      ," + DbQuote(DbQuoteType.Str, row["JCCategory"]));
                insert.AppendLine("      ," + DbQuote(DbQuoteType.Num, row["LineNumber"]));
                insert.AppendLine("      )");
                app.SqlExec(insert.ToString(), Application.Connections.dbHomefront);
            }

            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"CancelPO_TL error: {ex.Message}");
            return false;
        }
    }

    public static object? ChangePOVendor(params object?[] args)
    {
        var poNumber = GetArgString(args, 0);
        var newVendor = GetArgString(args, 1);
        var updateInvoices = GetArgBool(args, 2);
        if (string.IsNullOrWhiteSpace(poNumber)) return false;
        var app = HFApp;
        if (app == null) return false;

        try
        {
            var divisionId = int.TryParse(app.DivisionID, out var div) ? div : DivisionID;
            var acctSystem = (Application.AccountingSystems)OptionInt("AccountingSystem");

            if (acctSystem != Application.AccountingSystems.asTimberline &&
                acctSystem != Application.AccountingSystems.asQuickBooks)
            {
                var s = new StringBuilder();
                s.AppendLine("select v.istbd");
                s.AppendLine("from pomaster p");
                s.AppendLine("join tblvendors v on p.vendor=v.vendor_id and p.divisionid=v.divisionid");
                s.AppendLine($"where p.ponumber={DbQuote(DbQuoteType.Str, poNumber)}");
                s.AppendLine($"and p.divisionid={DbQuote(DbQuoteType.Num, divisionId)}");
                var check = app.SqlExec(s.ToString(), Application.Connections.dbHomefront);
                var isTbd = check.Rows.Count > 0 && ToBool(check.Rows[0][0]);
                if (!isTbd)
                {
                    Console.WriteLine("ChangePOVendor: accounting system does not support change vendor.");
                    return false;
                }
            }

            var posted = false;
            if (acctSystem == Application.AccountingSystems.asTimberline)
            {
                var rs = app.SqlExec(
                    "SELECT Amount_Invoiced FROM JCM_Master__Commitment WHERE Commitment=@PoNumber",
                    new { PoNumber = poNumber },
                    Application.Connections.dbAccounting);
                if (rs.Rows.Count == 0)
                {
                    posted = false;
                    if (!updateInvoices)
                    {
                        var inv = app.SqlExec(
                            "Select isnull(sum(PreTax),0) from dbo.POInvoicedAmounts where DivisionID=@DivisionId and PONumber=@PoNumber",
                            new { DivisionId = divisionId, PoNumber = poNumber },
                            Application.Connections.dbHomefront);
                        var amtInvoiced = inv.Rows.Count > 0 ? ToDecimal(inv.Rows[0][0]) : 0m;
                        if (amtInvoiced != 0)
                        {
                            Console.WriteLine("ChangePOVendor: PO has invoices; cannot change vendor.");
                            return false;
                        }
                    }
                }
                else
                {
                    posted = true;
                    if (ToDecimal(rs.Rows[0][0]) != 0)
                    {
                        Console.WriteLine("ChangePOVendor: commitment has invoice; cannot change vendor.");
                        return false;
                    }
                    var cos = app.SqlExec(
                        "SELECT Commitment_CO FROM JCM_Master__Commitment_CO WHERE Commitment=@PoNumber",
                        new { PoNumber = poNumber },
                        Application.Connections.dbAccounting);
                    if (cos.Rows.Count > 0)
                    {
                        Console.WriteLine("ChangePOVendor: commitment has change order; cannot change vendor.");
                        return false;
                    }
                }
            }
            else if (!updateInvoices)
            {
                var inv = app.SqlExec(
                    "Select isnull(sum(PreTax),0) from dbo.POInvoicedAmounts where DivisionID=@DivisionId and PONumber=@PoNumber",
                    new { DivisionId = divisionId, PoNumber = poNumber },
                    Application.Connections.dbHomefront);
                var amtInvoiced = inv.Rows.Count > 0 ? ToDecimal(inv.Rows[0][0]) : 0m;
                if (amtInvoiced != 0)
                {
                    Console.WriteLine("ChangePOVendor: PO has invoices; cannot change vendor.");
                    return false;
                }
            }

            if (!string.IsNullOrWhiteSpace(OptionValue("BuildProCompanyCode")))
            {
                var sent = app.SqlExec(
                    "select * from pomaster where datesenttobuildpro is not null and divisionid=@DivisionId and PONumber=@PoNumber",
                    new { DivisionId = divisionId, PoNumber = poNumber },
                    Application.Connections.dbHomefront);
                if (sent.Rows.Count > 0)
                {
                    Console.WriteLine("ChangePOVendor: PO sent to BuildPro; cannot change vendor.");
                    return false;
                }
            }

            var poInfo = app.SqlExec(
                "SELECT Vendor,Email,Vendor_Name,job FROM POMaster LEFT OUTER JOIN tblVendors ON(POMaster.DivisionID = tblVendors.DivisionID and Vendor=Vendor_ID) WHERE pomaster.DivisionID=@DivisionId and PONumber=@PoNumber",
                new { DivisionId = divisionId, PoNumber = poNumber },
                Application.Connections.dbHomefront);
            if (poInfo.Rows.Count == 0) return false;
            var oldVendor = poInfo.Rows[0]["vendor"]?.ToString() ?? "";
            var emailAddr = poInfo.Rows[0]["email"]?.ToString() ?? "";
            var oldVendorName = poInfo.Rows[0]["vendor_name"]?.ToString() ?? "";
            var job = poInfo.Rows[0]["job"]?.ToString() ?? "";

            if (string.IsNullOrWhiteSpace(newVendor))
            {
                Console.WriteLine("ChangePOVendor: New vendor is required in Blazor.");
                return false;
            }

            if (string.Equals(oldVendor.Trim(), newVendor.Trim(), StringComparison.OrdinalIgnoreCase))
                return false;

            if (acctSystem == Application.AccountingSystems.asTimberline)
            {
                var check = app.SqlExec(
                    "SELECT Vendor FROM APM_Master__Vendor WHERE Vendor=@NewVendor",
                    new { NewVendor = newVendor },
                    Application.Connections.dbAccounting);
                if (check.Rows.Count == 0)
                {
                    Console.WriteLine("ChangePOVendor: vendor not found in accounting.");
                    return false;
                }
            }

            var vendInfo = app.SqlExec(
                "SELECT Vendor_ID Vendor, Vendor_Name Company,City,Phone,PurchContact Contact,purchdelmethod DeliveryMethod,case purchdelmethod when 1 then case isnull(Purchemail,'') when '' then Email else purchEmail end else '' end Email FROM tblVendors where DivisionID=@DivisionId and vendor_id=@NewVendor",
                new { DivisionId = divisionId, NewVendor = newVendor },
                Application.Connections.dbHomefront);
            var newVendorName = vendInfo.Rows.Count > 0 ? vendInfo.Rows[0]["Company"]?.ToString() ?? newVendor : newVendor;
            var newEmailAddr = vendInfo.Rows.Count > 0 ? vendInfo.Rows[0]["Email"]?.ToString() ?? "" : "";
            var newDeliveryMethod = vendInfo.Rows.Count > 0 ? Convert.ToInt32(vendInfo.Rows[0]["DeliveryMethod"] == DBNull.Value ? 0 : vendInfo.Rows[0]["DeliveryMethod"]) : 0;
            var newContact = vendInfo.Rows.Count > 0 ? vendInfo.Rows[0]["Contact"]?.ToString() ?? "" : "";

            app.WriteAuditLog("Edit PO", $"Change vendor on \"{poNumber}\" from \"{oldVendor}\" to \"{newVendor}\"");

            if (acctSystem == Application.AccountingSystems.asTimberline && posted)
            {
                Console.WriteLine("ChangePOVendor: posted in Timberline; external update required.");
            }
            else if (acctSystem == Application.AccountingSystems.asQuickBooks)
            {
                var ok = (bool)(ChangeQBPOVendor(poNumber, newVendor) ?? false);
                if (!ok) return false;
            }

            app.SqlExec(
                "UPDATE POMaster SET Vendor=@NewVendor,DeliveryMethod=@DeliveryMethod,DeliveryAddress=@DeliveryAddress,DeliveryRecipient=@DeliveryRecipient WHERE DivisionID=@DivisionId and PONumber=@PoNumber",
                new
                {
                    NewVendor = newVendor,
                    DeliveryMethod = newDeliveryMethod,
                    DeliveryAddress = newEmailAddr,
                    DeliveryRecipient = newContact,
                    DivisionId = divisionId,
                    PoNumber = poNumber,
                },
                Application.Connections.dbHomefront);
            app.SqlExec(
                "UPDATE EstimateItems SET POVendor=@NewVendor WHERE DivisionID=@DivisionId and PONumber=@PoNumber",
                new { NewVendor = newVendor, DivisionId = divisionId, PoNumber = poNumber },
                Application.Connections.dbHomefront);

            if (updateInvoices)
            {
                var invoiceParams = new { NewVendor = newVendor, DivisionId = divisionId, PoNumber = poNumber };
                app.SqlExec(
                    "update invoices set vendor=@NewVendor " +
                    "where invoiceid in(select invoiceid from invoiceitems " +
                    "where divisionid=@DivisionId and commitment=@PoNumber)",
                    invoiceParams,
                    Application.Connections.dbHomefront);

                app.SqlExec(
                    "update invoiceitems set vendor=@NewVendor " +
                    "where invoiceid in(select invoiceid from invoiceitems " +
                    "where divisionid=@DivisionId and commitment=@PoNumber)",
                    invoiceParams,
                    Application.Connections.dbHomefront);

                app.SqlExec(
                    "update invoiceitems set commitmentvendor=@NewVendor " +
                    "where divisionid=@DivisionId and commitment=@PoNumber",
                    invoiceParams,
                    Application.Connections.dbHomefront);
            }

            UpdateScheduleVendor(job, DbQuote(DbQuoteType.Str, oldVendor), DbQuote(DbQuoteType.Str, newVendor))
                .GetAwaiter().GetResult();

            if (string.IsNullOrWhiteSpace(oldVendorName)) oldVendorName = oldVendor;
            if (string.IsNullOrWhiteSpace(emailAddr)) emailAddr = oldVendorName;
            if (!string.IsNullOrWhiteSpace(emailAddr))
            {
                Console.WriteLine($"ChangePOVendor: notify vendor {oldVendorName} at {emailAddr}.");
            }

            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"ChangePOVendor error: {ex.Message}");
            return false;
        }
    }

    public static object? ChangeQBPOVendor(params object?[] args)
    {
        var poNumber = GetArgString(args, 0);
        var newVendor = GetArgString(args, 1);
        if (string.IsNullOrWhiteSpace(poNumber) || string.IsNullOrWhiteSpace(newVendor)) return false;
        var app = HFApp;
        if (app == null) return false;

        try
        {
            var divisionId = int.TryParse(app.DivisionID, out var div) ? div : DivisionID;
            var s = "select * from pomaster where divisionid=" + DbQuote(DbQuoteType.Num, divisionId) +
                    " and ponumber=" + DbQuote(DbQuoteType.Str, poNumber);
            var rs = app.SqlExec(s, Application.Connections.dbHomefront);
            if (rs.Rows.Count == 0) return false;

            var postingBatch = rs.Rows[0]["postingbatch"] == DBNull.Value ? 0 : Convert.ToInt32(rs.Rows[0]["postingbatch"]);
            if (postingBatch == 0)
                return true;

            var poDate = rs.Rows[0]["PODate"];
            var xml = app.XmlQBStart();
            xml += "<GeneralDetailReportQueryRq>" + Environment.NewLine;
            xml += "<GeneralDetailReportType>OpenPOs</GeneralDetailReportType>" + Environment.NewLine;
            xml += "<ReportPeriod>" + Environment.NewLine;
            xml += app.XmlQBAdd(Application.XMLFieldTypes.d, 0, "FromReportDate", poDate);
            xml += app.XmlQBAdd(Application.XMLFieldTypes.d, 0, "ToReportDate", poDate);
            xml += "</ReportPeriod>" + Environment.NewLine;
            xml += "<IncludeColumn>RefNumber</IncludeColumn>" + Environment.NewLine;
            xml += "<IncludeColumn>TxnID</IncludeColumn>" + Environment.NewLine;
            xml += "<IncludeColumn>Amount</IncludeColumn>" + Environment.NewLine;
            xml += "<IncludeColumn>Name</IncludeColumn>" + Environment.NewLine;
            xml += "<IncludeColumn>Date</IncludeColumn>" + Environment.NewLine;
            xml += "</GeneralDetailReportQueryRq>" + Environment.NewLine;
            xml += app.XmlQBEnd();
            var response = app.XmlQBSubmit(xml);

            var txn = app.SqlExec(
                "exec qb_POtxnid @Response, @PoNumber",
                new { Response = response, PoNumber = poNumber },
                Application.Connections.dbHomefront);
            if (txn.Rows.Count == 0) return false;
            var txnId = txn.Rows[0][0]?.ToString() ?? "";
            if (string.IsNullOrWhiteSpace(txnId)) return false;

            xml = app.XmlQBStart();
            xml += "<PurchaseOrderQueryRq>" + Environment.NewLine;
            xml += "<TxnID>" + "</TxnID>" + Environment.NewLine;
            xml += app.XmlQBAdd(Application.XMLFieldTypes.st, 40, "TxnID", txnId);
            xml += "<IncludeLineItems>true</IncludeLineItems>" + Environment.NewLine;
            xml += "</PurchaseOrderQueryRq>" + Environment.NewLine;
            xml += app.XmlQBEnd();
            var detail = app.XmlQBSubmit(xml);
            var editSeq = Parse(Parse(detail, 2, "EditSequence>"), 1, "<");
            var rcvdQty = Parse(Parse(detail, 2, "ReceivedQuantity>"), 1, "<");

            if (!double.TryParse(rcvdQty, NumberStyles.Any, CultureInfo.InvariantCulture, out var rcvd) || rcvd == 0)
            {
                xml = app.XmlQBStart();
                xml += "<PurchaseOrderModRq>" + Environment.NewLine;
                xml += "   <PurchaseOrderMod>" + Environment.NewLine;
                xml += app.XmlQBAdd(Application.XMLFieldTypes.st, 40, "TxnID", txnId);
                xml += app.XmlQBAdd(Application.XMLFieldTypes.st, 40, "EditSequence", editSeq);
                xml += "<VendorRef>" + app.XmlQBAdd(Application.XMLFieldTypes.st, 40, "ListID", newVendor) + "</VendorRef>" + Environment.NewLine;
                xml += "   </PurchaseOrderMod>" + Environment.NewLine;
                xml += "</PurchaseOrderModRq>" + Environment.NewLine;
                xml += app.XmlQBEnd();
                app.XmlQBSubmit(xml);
                return true;
            }

            Console.WriteLine("ChangeQBPOVendor: PO has received qty; cannot change vendor.");
            return false;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"ChangeQBPOVendor error: {ex.Message}");
            return false;
        }
    }

    public static object? DefaultPrinterName(params object?[] args)
    {
        try
        {
            var type = Type.GetType("System.Drawing.Printing.PrinterSettings, System.Drawing.Common");
            if (type == null) return "";
            var instance = Activator.CreateInstance(type);
            var name = type.GetProperty("PrinterName")?.GetValue(instance)?.ToString();
            return name ?? "";
        }
        catch
        {
            return "";
        }
    }

    public static object? Dir(params object?[] args)
    {
        var pathArg = GetArg(args, 0);
        var sorted = GetArgBool(args, 2);
        lock (_dirLock)
        {
            if (pathArg == null)
            {
                if (_dirIndex + 1 < _dirList.Count)
                {
                    _dirIndex++;
                    return _dirList[_dirIndex];
                }
                return "";
            }

            var path = pathArg.ToString() ?? "";
            if (string.IsNullOrWhiteSpace(path)) return "";
            var dir = Path.GetDirectoryName(path);
            var pattern = Path.GetFileName(path);
            if (string.IsNullOrWhiteSpace(dir)) dir = ".";
            if (string.IsNullOrWhiteSpace(pattern)) pattern = "*";

            var list = new List<string>();
            try
            {
                foreach (var entry in Directory.GetFileSystemEntries(dir, pattern))
                    list.Add(Path.GetFileName(entry));
            }
            catch
            {
                list.Clear();
            }

            if (sorted)
                list.Sort(StringComparer.OrdinalIgnoreCase);

            _dirList = list;
            _dirIndex = 0;
            _dirSorted = sorted;
            return _dirList.Count > 0 ? _dirList[0] : "";
        }
    }

    public static object? DNull(params object?[] args)
    {
        var s = GetArgString(args, 0);
        var idx = s.IndexOf('\0');
        return idx >= 0 ? s[..idx] : s;
    }

    public static void EnableCtrl(params object?[] args)
    {
        var ctrl = GetArg(args, 0);
        var enabled = GetArgBool(args, 1, true);
        TrySetProperty(ctrl, "Enabled", enabled);
        TrySetProperty(ctrl, "IsEnabled", enabled);
        TrySetProperty(ctrl, "Disabled", !enabled);
    }

    public static object? FileDrive(params object?[] args)
    {
        var fullpath = GetArgString(args, 0).Trim();
        if (fullpath.StartsWith(@"\\"))
        {
            var parts = fullpath.TrimStart('\\').Split('\\');
            return parts.Length > 0 ? @"\\" + parts[0] : @"\\";
        }
        return fullpath.Length >= 2 ? fullpath.Substring(0, 2) : fullpath;
    }

    public static object? GetComboBoxListID(params object?[] args)
    {
        var combo = GetArg(args, 0);
        var indexObj = TryGetProperty(combo, "ListIndex") ?? TryGetProperty(combo, "SelectedIndex");
        var index = indexObj is int i ? i : -1;
        if (index < 0) return -1;

        var itemData = TryGetProperty(combo, "ItemData");
        if (itemData is System.Collections.IList list && index < list.Count)
            return list[index];

        var items = TryGetProperty(combo, "Items") as System.Collections.IList;
        if (items != null && index < items.Count)
        {
            var item = items[index];
            var id = TryGetProperty(item, "Id") ?? TryGetProperty(item, "ID") ?? TryGetProperty(item, "Value");
            if (id != null) return id;
        }

        var selectedValue = TryGetProperty(combo, "SelectedValue") ?? TryGetProperty(combo, "Value");
        return selectedValue ?? -1;
    }

    public static object? GetComboBoxListIndex(params object?[] args)
    {
        var combo = GetArg(args, 0);
        var indexObj = TryGetProperty(combo, "ListIndex") ?? TryGetProperty(combo, "SelectedIndex");
        if (indexObj is int i) return i;
        return -1;
    }

    public static object? GetComboBoxListKey(params object?[] args)
    {
        var combo = GetArg(args, 0);
        var indexObj = TryGetProperty(combo, "ListIndex") ?? TryGetProperty(combo, "SelectedIndex");
        var index = indexObj is int i ? i : -1;
        if (index < 0) return "";
        var tag = TryGetProperty(combo, "Tag")?.ToString() ?? "";
        return Parse(tag, index + 1, "\u0001");
    }

    public static object? GetLocalizedPath(params object?[] args)
    {
        var sPath = GetArgString(args, 0);
        if (string.IsNullOrWhiteSpace(sPath)) return sPath;
        if (!ToBool(IsPathNetPath(sPath))) return sPath;

        var s = sPath.StartsWith(@"\\", StringComparison.Ordinal) ? sPath : (Mapped2UNC(sPath)?.ToString() ?? sPath);
        var server = "";
        if (s.StartsWith(@"\\"))
        {
            var parts = s.TrimStart('\\').Split('\\');
            if (parts.Length > 0) server = parts[0];
        }

        if (string.Equals(MachineName(), server, StringComparison.OrdinalIgnoreCase))
            return UNC2Local(s) ?? sPath;
        return sPath;
    }

    public static void GetPrinterInfo(params object?[] args)
    {
        // Not supported in Blazor; leave output args unchanged.
    }

    public static object? GetStrFromPtrA(params object?[] args)
    {
        if (args.Length == 0 || args[0] == null) return "";
        var ptr = new IntPtr(Convert.ToInt64(args[0], CultureInfo.InvariantCulture));
        return Marshal.PtrToStringAnsi(ptr) ?? "";
    }

    public static void GridAutoSizeRows(params object?[] args)
    {
        var grid = GetArg(args, 0);
        TryInvokeResult(grid, "AutoFitColumnsAsync");
        TryInvokeResult(grid, "AutoSizeRows");
        TryInvokeResult(grid, "AutoSize");
    }

    public static object? GridCellSelected(params object?[] args)
    {
        var grid = GetArg(args, 0);
        var row = GetArgInt(args, 1) ?? 0;
        var col = GetArgInt(args, 2) ?? 0;
        var result = TryInvokeResult(grid, "GetSelectedRowCellIndexesAsync");
        var listObj = UnwrapTaskResult(result) as System.Collections.IEnumerable;
        if (listObj == null) return false;
        foreach (var item in listObj)
        {
            if (TryGetCellTuple(item, out var r, out var c))
            {
                if (r == row && c == col) return true;
            }
        }
        return false;
    }

    public static void GridExpandALL(params object?[] args)
    {
        var grid = GetArg(args, 0);
        TryInvokeResult(grid, "ExpandAllGroupAsync");
        TryInvokeResult(grid, "ExpandAllGroups");
        TryInvokeResult(grid, "ExpandAllRows");
    }

    public static void GridGotFocus(params object?[] args)
    {
        var grid = GetArg(args, 0);
        TryInvoke(grid, "Focus");
    }

    public static object? GridHeight(params object?[] args)
    {
        var grid = GetArg(args, 0);
        var heightStr = TryGetProperty(grid, "Height")?.ToString();
        if (!string.IsNullOrWhiteSpace(heightStr))
        {
            var m = Regex.Match(heightStr, @"-?\d+(\.\d+)?");
            if (m.Success && double.TryParse(m.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out var hv))
                return (long)hv;
        }

        var rows = GridRowCount(grid);
        var rowHeight = GridRowHeight(grid);
        var header = rowHeight;
        return (long)(rows * rowHeight + header);
    }

    public static object? GridLastVisibleColumn(params object?[] args)
    {
        var grid = GetArg(args, 0);
        var cols = EnumerateGridColumns(grid).ToList();
        for (var i = cols.Count - 1; i >= 0; i--)
        {
            if (ColumnVisible(cols[i])) return i;
        }
        return -1;
    }

    public static object? GridLastVisibleRow(params object?[] args)
    {
        var grid = GetArg(args, 0);
        var count = GridRowCount(grid);
        return count > 0 ? count - 1 : -1;
    }

    public static object? GridNextVisibleColumn(params object?[] args)
    {
        var grid = GetArg(args, 0);
        var start = GetArgInt(args, 1) ?? 0;
        var cols = EnumerateGridColumns(grid).ToList();
        for (var i = Math.Max(0, start); i < cols.Count; i++)
        {
            if (ColumnVisible(cols[i])) return i;
        }
        return -1;
    }

    public static object? GridNextVisibleRow(params object?[] args)
    {
        var grid = GetArg(args, 0);
        var start = GetArgInt(args, 1) ?? 0;
        var count = GridRowCount(grid);
        return start >= 0 && start < count ? start : -1;
    }

    public static object? GridVisibleColumns(params object?[] args)
    {
        var grid = GetArg(args, 0);
        var cols = EnumerateGridColumns(grid);
        return cols.Count(ColumnVisible);
    }

    public static object? GridVisibleRowIndex(params object?[] args)
    {
        var grid = GetArg(args, 0);
        var nthRow = GetArgInt(args, 1) ?? 0;
        var startAt = GetArgInt(args, 2) ?? 0;
        if (nthRow <= 0) return -1;
        var count = GridRowCount(grid);
        var idx = startAt + (nthRow - 1);
        return idx >= 0 && idx < count ? idx : -1;
    }

    public static object? GridVisibleRows(params object?[] args)
    {
        var grid = GetArg(args, 0);
        return GridRowCount(grid);
    }

    public static object? GridWidth(params object?[] args)
    {
        var grid = GetArg(args, 0);
        double total = 0;
        foreach (var col in EnumerateGridColumns(grid))
        {
            if (!ColumnVisible(col)) continue;
            var width = ColumnWidth(col) ?? 100;
            total += width;
        }
        return (long)total;
    }

    public static object? ImageIndex(params object?[] args)
    {
        var imageList = GetArg(args, 0);
        var key = GetArgString(args, 1);
        if (imageList == null || string.IsNullOrWhiteSpace(key)) return -1;

        if (imageList is System.Collections.IList list)
        {
            for (var i = 0; i < list.Count; i++)
            {
                var item = list[i];
                var itemKey = TryGetProperty(item, "Key")?.ToString() ?? TryGetProperty(item, "Name")?.ToString();
                if (string.Equals(itemKey, key, StringComparison.OrdinalIgnoreCase)) return i;
            }
        }

        if (imageList is System.Collections.IDictionary dict)
        {
            var idx = 0;
            foreach (var k in dict.Keys)
            {
                if (string.Equals(k?.ToString(), key, StringComparison.OrdinalIgnoreCase)) return idx;
                idx++;
            }
        }

        var listImages = TryGetProperty(imageList, "ListImages");
        if (listImages is System.Collections.IEnumerable enumImages)
        {
            var idx = 0;
            foreach (var item in enumImages)
            {
                var itemKey = TryGetProperty(item, "Key")?.ToString() ?? TryGetProperty(item, "Name")?.ToString();
                if (string.Equals(itemKey, key, StringComparison.OrdinalIgnoreCase)) return idx;
                idx++;
            }
        }
        return -1;
    }

    public static object? InIde(params object?[] args)
    {
        return Debugger.IsAttached;
    }

    [Obsolete("Use IUserPreferencesService.Get() via DI instead.")]
    public static object? IniGet(params object?[] args)
    {
        var file = ResolveIniFile(GetArg(args, 0));
        var section = GetArgString(args, 1);
        var key = GetArgString(args, 2);
        var defaultVal = GetArg(args, 3) ?? "";
        var data = LoadIniFile(file);
        if (data.TryGetValue(section, out var sec) && sec.TryGetValue(key, out var val))
            return val;
        return defaultVal;
    }

    public static void IniGetForm(params object?[] args)
    {
        var form = GetArg(args, 0);
        var iniFile = ResolveIniFile(GetArg(args, 1));
        var instanceKey = GetArg(args, 2);
        var section = BuildSectionName(form, instanceKey);
        var data = LoadIniFile(iniFile);
        if (!data.TryGetValue(section, out var sec)) return;

        if (sec.TryGetValue("Left", out var left) && int.TryParse(left, out var l)) TrySetProperty(form, "Left", l);
        if (sec.TryGetValue("Top", out var top) && int.TryParse(top, out var t)) TrySetProperty(form, "Top", t);
        if (sec.TryGetValue("Width", out var width) && int.TryParse(width, out var w)) TrySetProperty(form, "Width", w);
        if (sec.TryGetValue("Height", out var height) && int.TryParse(height, out var h)) TrySetProperty(form, "Height", h);
    }

    public static void IniGetGrid(params object?[] args)
    {
        var form = GetArg(args, 0);
        var grid = GetArg(args, 1);
        var staticPositions = GetArgBool(args, 3);
        var instanceKey = GetArg(args, 4)?.ToString();
        var grouped = GetArgBool(args, 5);
        var staticNames = GetArgBool(args, 6);

        if (grid == null) return;
        var formName = TryGetProperty(form, "Name")?.ToString() ?? form?.GetType().Name ?? "Form";
        var gridName = TryGetProperty(grid, "Name")?.ToString() ?? grid.GetType().Name ?? "Grid";
        var gridIndex = 0;
        var idxObj = TryGetProperty(grid, "Index");
        if (idxObj is int gi) gridIndex = gi;

        var layout = IniGetGridAsync(formName, gridName, gridIndex, instanceKey, grouped, staticNames, staticPositions, HFApp?.LoginID)
            .GetAwaiter().GetResult();
        if (layout.Count == 0) return;

        var allHidden = layout.Values.All(l => l.Hidden);
        var groupFields = new List<string>();

        foreach (var col in EnumerateGridColumns(grid))
        {
            var field = TryGetProperty(col, "Field")?.ToString();
            if (string.IsNullOrWhiteSpace(field)) continue;
            if (!layout.TryGetValue(field, out var layoutCol)) continue;

            if (!staticNames && !string.IsNullOrWhiteSpace(layoutCol.Caption))
                TrySetProperty(col, "HeaderText", layoutCol.Caption);

            if (layoutCol.Width.HasValue)
            {
                TrySetProperty(col, "Width", $"{layoutCol.Width.Value}px");
                TrySetProperty(col, "RuntimeWidth", layoutCol.Width.Value);
            }

            if (!allHidden)
                TrySetProperty(col, "Visible", !layoutCol.Hidden);

            if (layoutCol.Grouped)
                groupFields.Add(field);
        }

        if (!staticPositions)
        {
            try
            {
                var container = grid.GetType().GetField("_columnsContainer", BindingFlags.NonPublic | BindingFlags.Instance)?.GetValue(grid);
                var listField = container?.GetType().GetField("_columns", BindingFlags.NonPublic | BindingFlags.Instance);
                if (listField?.GetValue(container) is List<Fx.ControlKit.Grid.GridColumn> list)
                {
                    list.Sort((a, b) =>
                    {
                        var ai = layout.TryGetValue(a.Field, out var la) ? la.ColIndex ?? int.MaxValue : int.MaxValue;
                        var bi = layout.TryGetValue(b.Field, out var lb) ? lb.ColIndex ?? int.MaxValue : int.MaxValue;
                        return ai.CompareTo(bi);
                    });
                }
            }
            catch
            {
                // ignore ordering issues
            }
        }

        if (grouped && groupFields.Count > 0)
        {
            TryInvokeResult(grid, "ClearGroupingAsync");
            foreach (var field in groupFields)
                TryInvokeResult(grid, "GroupByColumnAsync", field);
        }
    }

    public static object? IniGetSectionNames(params object?[] args)
    {
        var file = ResolveIniFile(GetArg(args, 0));
        var data = LoadIniFile(file);
        return string.Join(",", data.Keys.Where(k => !string.IsNullOrEmpty(k)));
    }

    public static object? IniGetVariableNames(params object?[] args)
    {
        var file = ResolveIniFile(GetArg(args, 0));
        var section = GetArgString(args, 1);
        var data = LoadIniFile(file);
        return data.TryGetValue(section, out var sec) ? string.Join(",", sec.Keys) : "";
    }

    [Obsolete("Use IUserPreferencesService.Put() via DI instead.")]
    public static void IniPut(params object?[] args)
    {
        var file = ResolveIniFile(GetArg(args, 0));
        var section = GetArgString(args, 1);
        var key = GetArgString(args, 2);
        var value = GetArg(args, 3)?.ToString() ?? "";
        var data = LoadIniFile(file);
        if (!data.ContainsKey(section))
            data[section] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        data[section][key] = value;
        SaveIniFile(file, data);
    }

    public static void IniPutForm(params object?[] args)
    {
        var form = GetArg(args, 0);
        var iniFile = ResolveIniFile(GetArg(args, 1));
        var instanceKey = GetArg(args, 2);
        var section = BuildSectionName(form, instanceKey);
        var data = LoadIniFile(iniFile);
        if (!data.ContainsKey(section))
            data[section] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        void set(string key, object? value)
        {
            if (value == null) return;
            data[section][key] = value.ToString() ?? "";
        }

        set("Left", TryGetProperty(form, "Left"));
        set("Top", TryGetProperty(form, "Top"));
        set("Width", TryGetProperty(form, "Width"));
        set("Height", TryGetProperty(form, "Height"));
        SaveIniFile(iniFile, data);
    }

    public static void IniPutGrid(params object?[] args)
    {
        var form = GetArg(args, 0);
        var grid = GetArg(args, 1);
        var instanceKey = GetArg(args, 3)?.ToString();

        if (grid == null) return;
        var formName = TryGetProperty(form, "Name")?.ToString() ?? form?.GetType().Name ?? "Form";
        var gridName = TryGetProperty(grid, "Name")?.ToString() ?? grid.GetType().Name ?? "Grid";
        var gridIndex = 0;
        var idxObj = TryGetProperty(grid, "Index");
        if (idxObj is int gi) gridIndex = gi;
        if (gridIndex > 0)
            gridName = $"{gridName}.({gridIndex})";

        var groupedFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            var gdField = grid.GetType().GetField("_groupDescriptors", BindingFlags.NonPublic | BindingFlags.Instance);
            if (gdField?.GetValue(grid) is System.Collections.IEnumerable gds)
            {
                foreach (var gd in gds)
                {
                    var field = TryGetProperty(gd, "Field")?.ToString();
                    if (!string.IsNullOrWhiteSpace(field)) groupedFields.Add(field);
                }
            }
        }
        catch
        {
            // ignore
        }

        var cols = new List<GridLayoutColumn>();
        var colIndex = 0;
        foreach (var col in EnumerateGridColumns(grid))
        {
            var field = TryGetProperty(col, "Field")?.ToString();
            if (string.IsNullOrWhiteSpace(field)) continue;
            var caption = TryGetProperty(col, "HeaderText")?.ToString() ?? "";
            var width = ColumnWidth(col);
            var hidden = !ColumnVisible(col);
            cols.Add(new GridLayoutColumn
            {
                Field = field,
                Caption = caption,
                Width = width,
                Hidden = hidden,
                ColIndex = colIndex,
                Grouped = groupedFields.Contains(field)
            });
            colIndex++;
        }

        DBPutGrid(formName, gridName, instanceKey, cols).GetAwaiter().GetResult();
    }

    public static void IniRemove(params object?[] args)
    {
        var file = ResolveIniFile(GetArg(args, 0));
        var section = GetArgString(args, 1);
        var key = GetArgString(args, 2);
        var data = LoadIniFile(file);
        if (string.IsNullOrEmpty(section)) return;
        if (string.IsNullOrEmpty(key))
        {
            data.Remove(section);
        }
        else if (data.TryGetValue(section, out var sec))
        {
            sec.Remove(key);
        }
        SaveIniFile(file, data);
    }

    public static void InsertSortStrings(params object?[] args)
    {
        SortStringsInPlace(GetArg(args, 0));
    }

    public static void InsertSortStringsStart(params object?[] args)
    {
        SortStringsInPlace(GetArg(args, 0));
    }

    public static object? ISDEBUG(params object?[] args)
    {
        try
        {
            var name = Process.GetCurrentProcess().ProcessName ?? "";
            return name.Contains("debug", StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    public static object? IsFormLoaded(params object?[] args)
    {
        var formName = GetArgString(args, 0);
        if (string.IsNullOrWhiteSpace(formName)) return false;
        return string.Equals(LoadedForm, formName, StringComparison.OrdinalIgnoreCase);
    }

    public static object? IsPathNetPath(params object?[] args)
    {
        var path = GetArgString(args, 0);
        if (string.IsNullOrWhiteSpace(path)) return false;
        if (path.StartsWith(@"\\", StringComparison.Ordinal)) return true;
        try
        {
            var root = Path.GetPathRoot(path);
            if (string.IsNullOrWhiteSpace(root)) return false;
            var drive = new DriveInfo(root);
            return drive.DriveType == DriveType.Network;
        }
        catch
        {
            return false;
        }
    }

    public static object? IsUNCPathValid(params object?[] args)
    {
        var path = GetArgString(args, 0);
        return path.StartsWith(@"\\", StringComparison.Ordinal);
    }

    public static void LoadComboBox(params object?[] args)
    {
        var combo = GetArg(args, 0);
        var db = GetArg(args, 1) as DbWrapperSqlServer;
        var sql = GetArgString(args, 2);
        if (combo is System.Collections.IList list)
        {
            list.Clear();
            if (db != null && !string.IsNullOrWhiteSpace(sql))
            {
                var rows = db.QueryAsync(sql).Result;
                foreach (var row in rows)
                {
                    if (row.Count == 0) continue;
                    var val = row.Values.FirstOrDefault()?.ToString() ?? "";
                    list.Add(val);
                }
            }
        }
    }

    public static object? LoadExcelSheet(params object?[] args)
    {
        // Excel automation not supported in Blazor. Return null.
        return null;
    }

    public static void Main(params object?[] args)
    {
        // Entry point not used in Blazor.
    }

    public static object? Mapped2UNC(params object?[] args)
    {
        var mapped = GetArgString(args, 0);
        if (string.IsNullOrWhiteSpace(mapped)) return "";
        var drive = FileDrive(mapped)?.ToString() ?? "";
        if (string.IsNullOrWhiteSpace(drive)) return "";

        try
        {
            var sb = new StringBuilder(260);
            var len = sb.Capacity;
            var result = WNetGetConnection(drive, sb, ref len);
            if (result != 0) return "";
            var root = sb.ToString();
            var remainder = StripRootFromPath(mapped)?.ToString() ?? "";
            return PathAppend(root, remainder);
        }
        catch
        {
            return "";
        }
    }

    public static object? Max(params object?[] args)
    {
        if (args.Length == 0) return null;
        object? max = args[0];
        for (var i = 1; i < args.Length; i++)
        {
            if (CompareObjects(args[i], max) > 0) max = args[i];
        }
        return max;
    }

    public static object? MbApiIsRunning(params object?[] args)
    {
        try
        {
            if (string.Equals(OptionValue("Sage100APILevel"), "v18", StringComparison.OrdinalIgnoreCase))
            {
                var type = Type.GetTypeFromProgID("MBAPI.IMBXML");
                return type != null;
            }
            return true;
        }
        catch
        {
            return false;
        }
    }

    public static void MergeSortStrings(params object?[] args)
    {
        SortStringsInPlace(GetArg(args, 0));
    }

    public static void MergeSortStringsStart(params object?[] args)
    {
        SortStringsInPlace(GetArg(args, 0));
    }

    public static void MergeStrings(params object?[] args)
    {
        SortStringsInPlace(GetArg(args, 0));
    }

    public static object? Min(params object?[] args)
    {
        if (args.Length == 0) return null;
        object? min = args[0];
        for (var i = 1; i < args.Length; i++)
        {
            if (CompareObjects(args[i], min) < 0) min = args[i];
        }
        return min;
    }

    public static object? MouseX(params object?[] args)
    {
        return 0;
    }

    public static object? MouseY(params object?[] args)
    {
        return 0;
    }

    public static object? MyLookUp(params object?[] args)
    {
        var table = GetArgString(args, 0);
        var outField = GetArgString(args, 1);
        if (string.IsNullOrWhiteSpace(table) || string.IsNullOrWhiteSpace(outField)) return null;

        var where = new List<string>();
        for (var i = 2; i + 1 < args.Length; i += 3)
        {
            var keyField = GetArgString(args, i);
            var value = GetArg(args, i + 1);
            var dtype = GetArgString(args, i + 2);
            if (string.IsNullOrWhiteSpace(keyField)) continue;

            var quoted = string.Equals(dtype, "S", StringComparison.OrdinalIgnoreCase)
                ? DbQuote(DbQuoteType.Str, value)
                : DbQuote(DbQuoteType.Num, value);
            where.Add($"{keyField}={quoted}");
        }

        var sql = $"SELECT {outField} FROM {table}";
        if (where.Count > 0)
            sql += " WHERE " + string.Join(" AND ", where);

        try
        {
            return db.ExecuteScalar<object>(sql);
        }
        catch
        {
            return null;
        }
    }

    public static object? PointerToStringA(params object?[] args)
    {
        if (args.Length == 0 || args[0] == null) return "";
        var ptr = new IntPtr(Convert.ToInt64(args[0], CultureInfo.InvariantCulture));
        return Marshal.PtrToStringAnsi(ptr) ?? "";
    }

    public static object? QualifyPath(params object?[] args)
    {
        var path = GetArgString(args, 0);
        if (string.IsNullOrEmpty(path)) return path;
        return path.EndsWith("\\", StringComparison.Ordinal) ? path : path + "\\";
    }

    public static void QuickSortStrings(params object?[] args)
    {
        SortStringsInPlace(GetArg(args, 0));
    }

    public static object? QuickSortStringsPartition(params object?[] args)
    {
        var list = GetArg(args, 0) as string[];
        var low = GetArgInt(args, 1) ?? 0;
        var high = GetArgInt(args, 2) ?? (list?.Length ?? 0) - 1;
        var order = GetArgInt(args, 3) ?? -1;
        var compareType = GetArgInt(args, 4) ?? 1;
        if (list == null || list.Length == 0) return low;

        var comparer = compareType == 0 ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase;
        var ascending = order <= 0;

        var pivot = low + (high - low) / 2;
        var pivotValue = list[pivot];
        list[pivot] = list[low];

        var iLow = low + 1;
        var iHigh = high;
        while (true)
        {
            while (iLow < iHigh)
            {
                var cmp = comparer.Compare(pivotValue, list[iLow]);
                if (ascending ? cmp < 0 : cmp > 0) break;
                iLow++;
            }

            while (iHigh >= iLow)
            {
                var cmp = comparer.Compare(list[iHigh], pivotValue);
                if (ascending ? cmp < 0 : cmp > 0) break;
                iHigh--;
            }

            if (iLow >= iHigh) break;
            (list[iLow], list[iHigh]) = (list[iHigh], list[iLow]);
        }

        list[low] = list[iHigh];
        list[iHigh] = pivotValue;
        return iHigh;
    }

    public static void QuickSortStringsStart(params object?[] args)
    {
        SortStringsInPlace(GetArg(args, 0));
    }

    public static object? RegEnumKeys(params object?[] args)
    {
        var hive = GetArg(args, 0);
        var section = GetArgString(args, 1);
        try
        {
            using var key = OpenRegistryKey(hive, section, false);
            if (key == null) return Array.Empty<string>();
            return key.GetValueNames();
        }
        catch
        {
            return Array.Empty<string>();
        }
    }

    public static object? RegGetKey(params object?[] args)
    {
        var hive = GetArg(args, 0);
        var section = GetArgString(args, 1);
        var keyName = GetArgString(args, 2);
        var def = GetArgString(args, 3);
        try
        {
            using var key = OpenRegistryKey(hive, section, false);
            if (key == null) return def;
            if (keyName == "*") keyName = "";
            return key.GetValue(keyName)?.ToString() ?? def;
        }
        catch
        {
            return def;
        }
    }

    public static object? RegSaveKey(params object?[] args)
    {
        var hive = GetArg(args, 0);
        var section = GetArgString(args, 1);
        var keyName = GetArgString(args, 2);
        var value = GetArgString(args, 3);
        try
        {
            using var key = OpenRegistryKey(hive, section, true) ?? Registry.LocalMachine.CreateSubKey(section);
            if (key == null) return false;
            if (string.IsNullOrEmpty(keyName)) keyName = "";
            key.SetValue(keyName, value ?? "");
            return true;
        }
        catch
        {
            return false;
        }
    }

    public static object? RegSectionExists(params object?[] args)
    {
        var hive = GetArg(args, 0);
        var section = GetArgString(args, 1);
        try
        {
            using var key = OpenRegistryKey(hive, section, false);
            return key != null;
        }
        catch
        {
            return false;
        }
    }

    public static object? SaveToCSV(params object?[] args)
    {
        var fileName = GetArgString(args, 0);
        if (string.IsNullOrWhiteSpace(fileName)) return "";
        var ext = Path.GetExtension(fileName);
        if (string.Equals(ext, ".csv", StringComparison.OrdinalIgnoreCase))
            return fileName;

        var temp = TempFile("csv");
        try
        {
            File.Copy(fileName, temp, true);
            return temp;
        }
        catch
        {
            return fileName;
        }
    }

    public static void SelectAll(params object?[] args)
    {
        var ctrl = GetArg(args, 0);
        TryInvoke(ctrl, "SelectAll");
    }

    public static object? SelectedOption(params object?[] args)
    {
        var options = GetArg(args, 0);
        if (options is bool[] array)
        {
            for (var i = 0; i < array.Length; i++)
                if (array[i]) return i;
        }
        if (options is System.Collections.IList list)
        {
            for (var i = 0; i < list.Count; i++)
                if (ToBool(list[i])) return i;
        }
        return -1;
    }

    public static void SelectionSortStrings(params object?[] args)
    {
        SortStringsInPlace(GetArg(args, 0));
    }

    public static void SendingWizard(params object?[] args)
    {
        var docType = GetArgString(args, 0);
        var job = GetArgString(args, 1);
        var poNumbers = GetArgString(args, 2);
        var specifiedOnly = GetArgBool(args, 3);
        var cmd = SyncCmd("SendingWizard", $"{docType}|{job}|{poNumbers}|{specifiedOnly}", true, true);
        Console.WriteLine($"[SendingWizard] {cmd}");
    }

    public static void SetComboBoxListIndex(params object?[] args)
    {
        var combo = GetArg(args, 0);
        var text = GetArgString(args, 1);
        var key = GetArgString(args, 2);
        var id = GetArg(args, 3);
        var index = GetArgInt(args, 4);

        if (combo is System.Collections.IList list)
        {
            if (index.HasValue && index.Value >= 0 && index.Value < list.Count)
            {
                TrySetProperty(combo, "SelectedIndex", index.Value);
                return;
            }
            if (!string.IsNullOrWhiteSpace(text))
            {
                for (var i = 0; i < list.Count; i++)
                {
                    if (string.Equals(list[i]?.ToString(), text, StringComparison.OrdinalIgnoreCase))
                    {
                        TrySetProperty(combo, "SelectedIndex", i);
                        return;
                    }
                }
            }
        }
        if (!string.IsNullOrWhiteSpace(text))
            TrySetProperty(combo, "Text", text);
    }

    public static void SetCtrlFocus(params object?[] args)
    {
        var ctrl = GetArg(args, 0);
        TryInvoke(ctrl, "Focus");
    }

    public static object? SetListIndex(params object?[] args)
    {
        var combo = GetArg(args, 0);
        var itemData = GetArg(args, 1);
        var text = GetArgString(args, 2);
        var list = TryGetProperty(combo, "Items") as System.Collections.IList;
        if (list == null)
            list = combo as System.Collections.IList;

        if (!string.IsNullOrEmpty(text))
        {
            if (list != null)
            {
                for (var i = 0; i < list.Count; i++)
                {
                    if (string.Equals(list[i]?.ToString(), text, StringComparison.OrdinalIgnoreCase))
                    {
                        TrySetProperty(combo, "SelectedIndex", i);
                        TrySetProperty(combo, "ListIndex", i);
                        return true;
                    }
                }
            }
        }
        else
        {
            if (itemData == null || Convert.ToInt64(itemData, CultureInfo.InvariantCulture) == 0)
            {
                TrySetProperty(combo, "Text", "");
                return true;
            }

            var dataList = TryGetProperty(combo, "ItemData") as System.Collections.IList;
            if (dataList != null)
            {
                for (var i = 0; i < dataList.Count; i++)
                {
                    if (Equals(dataList[i], itemData))
                    {
                        TrySetProperty(combo, "SelectedIndex", i);
                        TrySetProperty(combo, "ListIndex", i);
                        return true;
                    }
                }
            }
        }

        return false;
    }

    public static void SetToolbarIcons(params object?[] args)
    {
        // Toolbar icon assignment is UI-framework specific; no-op in Blazor.
    }

    public static void SetTopMostWindow(params object?[] args)
    {
        // Not applicable in Blazor.
    }

    public static void SetWindowFocus(params object?[] args)
    {
        // Not applicable in Blazor.
    }

    public static object? ShellFile(params object?[] args)
    {
        var fileName = GetArgString(args, 1);
        var printIt = GetArgBool(args, 4);
        if (string.IsNullOrWhiteSpace(fileName)) return "";
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = fileName,
                UseShellExecute = true
            };
            if (printIt) psi.Verb = "print";
            Process.Start(psi);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"ShellFile error: {ex.Message}");
        }
        return fileName;
    }

    public static object? ShowForm(params object?[] args)
    {
        var formName = GetArgString(args, 0);
        if (string.IsNullOrWhiteSpace(formName)) return false;
        return string.Equals(LoadedForm, formName, StringComparison.OrdinalIgnoreCase);
    }

    public static void ShowToolbarCaptions(params object?[] args)
    {
        // Toolbar captions are UI-framework specific; no-op in Blazor.
    }

    public static object? StripFormating(params object?[] args)
    {
        var s = GetArgString(args, 0);
        s = StripFormatting(s);
        s = s.Replace(((char)189).ToString(), "1/2");
        return s;
    }

    public static object? StripPathToRoot(params object?[] args)
    {
        var path = GetArgString(args, 0);
        if (string.IsNullOrWhiteSpace(path)) return "";
        var root = Path.GetPathRoot(path) ?? "";
        return root.TrimEnd('\\');
    }

    public static object? StripRootFromPath(params object?[] args)
    {
        var path = GetArgString(args, 0);
        if (string.IsNullOrWhiteSpace(path)) return "";
        var root = Path.GetPathRoot(path) ?? "";
        return path.StartsWith(root, StringComparison.OrdinalIgnoreCase) ? path[root.Length..] : path;
    }

    public static object? TrimNull(params object?[] args)
    {
        var s = GetArgString(args, 0);
        var idx = s.IndexOf('\0');
        return idx >= 0 ? s[..idx] : s;
    }

    public static object? trunc(params object?[] args)
    {
        var number = GetArg(args, 0);
        var precision = GetArgInt(args, 1) ?? 0;
        if (!double.TryParse(number?.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var val))
            return "0";
        var format = "0." + new string('0', Math.Max(0, precision));
        return val.ToString(format, CultureInfo.InvariantCulture);
    }

    public static object? UNC2Local(params object?[] args)
    {
        var unc = GetArgString(args, 0);
        if (string.IsNullOrWhiteSpace(unc) || !unc.StartsWith(@"\\", StringComparison.Ordinal)) return unc;
        var parts = unc.TrimStart('\\').Split('\\');
        if (parts.Length < 2) return unc;
        var server = parts[0];
        var share = parts[1];
        if (!string.Equals(server, MachineName(), StringComparison.OrdinalIgnoreCase)) return unc;

        if (share.EndsWith("$", StringComparison.Ordinal))
        {
            var drive = share[..^1] + ":";
            var remainder = string.Join("\\", parts.Skip(2));
            return Path.Combine(drive + "\\", remainder);
        }

        return unc;
    }

    public static object? ValidateField(params object?[] args)
    {
        var grid = GetArg(args, 0);
        var editText = GetArgString(args, 1);
        var warning = GetArgString(args, 2);
        var sql = GetArgString(args, 3);

        try
        {
            if (string.IsNullOrWhiteSpace(editText))
                return true;
            if (string.IsNullOrWhiteSpace(sql))
                return false;

            var dt = db.SqlExec(sql);
            if (dt.Rows.Count == 0)
            {
                if (!string.IsNullOrWhiteSpace(warning))
                    Console.WriteLine(warning);
                return false;
            }

            var row = dt.Rows[0];
            var value = row[0]?.ToString() ?? "";
            if (grid != null)
            {
                var currentRow = TryInvokeResult(grid, "GetCurrentRowIndex");
                var rowIndex = currentRow is int i ? i : -1;
                if (rowIndex >= 0)
                {
                    var dataSource = TryGetProperty(grid, "DataSource") as System.Collections.IList;
                    if (dataSource != null && rowIndex < dataSource.Count)
                    {
                        var item = dataSource[rowIndex];
                        var colField = TryGetProperty(grid, "CurrentColumnField")?.ToString();
                        if (!string.IsNullOrWhiteSpace(colField)) TrySetProperty(item, colField, value);
                    }
                }
            }

            return true;
        }
        catch (Exception ex)
        {
            if (!string.IsNullOrWhiteSpace(warning))
                Console.WriteLine($"{warning} ({ex.Message})");
            return false;
        }
    }

    public static object? vbBullet(params object?[] args)
    {
        return ((char)149).ToString();
    }

    public static object? vbQuote(params object?[] args)
    {
        return "\"";
    }

    public static object? WriteToFile(params object?[] args)
    {
        var fileName = GetArgString(args, 0);
        var description = GetArgString(args, 1);
        if (string.IsNullOrWhiteSpace(fileName)) return false;
        try
        {
            File.AppendAllText(fileName, description + Environment.NewLine);
            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"WriteToFile error: {ex.Message}");
            return false;
        }
    }

}
