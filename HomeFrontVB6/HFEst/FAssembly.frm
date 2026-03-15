Here is the complete VB6 file with TAG comment blocks inserted before every `Sub`, `Function`, `Property`, and `Begin` block:

```vb
VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{302C5C1A-C2E2-4302-9AF9-BAC87EEFDECE}#1.0#0"; "Panels.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
' TAG: Main MDI child form for the Model & Option Library, used to create/edit/view assembly masters and their detail items
' CONVERT: Create a Blazor page component at @page "/edit-model-assembly" with @rendermode InteractiveServer
' CONVERT: Use private fields for all form-level state (mDirty, mReadOnly, mAssemblyID, etc.) with StateHasChanged() for UI updates
' CONVERT: Replace MDIChild behavior with NavigationManager navigation and shared layout
' CONVERT: Use DbWrapperSqlServer for all database operations (QueryAsync, ExecuteAsync, etc.)
' CONVERT: Implement modal dialogs using _isVisible bool + overlay div pattern
Begin VB.Form FAssembly 
   Caption         =   "Model & Option Library"
   ClientHeight    =   9000
   ClientLeft      =   4905
   ClientTop       =   5070
   ClientWidth     =   20220
   Icon            =   "FAssembly.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   20220
   ' TAG: Horizontal slider/splitter control separating the components grid from the items grid
   ' CONVERT: Replace with a CSS flexbox or grid layout with a draggable splitter div
   ' CONVERT: Use a resizable panel pattern or a simple CSS resize handle between the two grid sections
   Begin Panels.Slider Slider 
      Height          =   60
      Left            =   105
      Top             =   4395
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   106
      Orientation     =   1
   End
   ' TAG: VSFlexGrid displaying component assemblies (child assemblies) linked to the current parent assembly
   ' CONVERT: Use <HfGrid TValue="AssemblyComponentRow"> with HfGridColumn components for community, model, option, assembly, description, qty, cost
   ' CONVERT: Load column layout via MMain.IniGetGridAsync("FAssembly", "gComponents", 0, staticNames: true)
   ' CONVERT: Support add/remove component rows via context menu with right-click handler
   ' CONVERT: Enable inline editing of the qty column only
   Begin VSFlex8Ctl.VSFlexGrid gComponents 
      Height          =   1020
      Left            =   105
      TabIndex        =   36
      Top             =   3105
      Width           =   12885
      _cx             =   22728
      _cy             =   1799
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   14336431
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   255
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FAssembly.frx":000C
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   5
      OutlineCol      =   1
      Ellipsis        =   0
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   1
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   2
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   1
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   ' TAG: Main toolbar with buttons for New, Open, Save, Takeoff operations, Mass Change, Pricelist Export, and Import
   ' CONVERT: Create a toolbar div with <button> elements for each action (New dropdown, Open, Save, Takeoff buttons, etc.)
   ' CONVERT: Use dropdown-menu-custom pattern for New and Import buttons with sub-menu items
   ' CONVERT: Bind disabled state to mReadOnly and mDirty flags
   ' CONVERT: Wire @onclick handlers to corresponding async methods (OnOpenClick, OnSaveClick, etc.)
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   20220
      _ExtentX        =   35666
      _ExtentY        =   1058
      ButtonWidth     =   1667
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Key             =   "New"
            Style           =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            Key             =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Save"
            Key             =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "One Time"
            Key             =   "TakeoffOneTime"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Item"
            Key             =   "TakeoffItem"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Assembly"
            Key             =   "TakeoffAssembly"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "PlanSwift"
            Key             =   "TakeoffPlanSwift"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Custom"
            Key             =   "TakeoffCustom"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Item Chart"
            Key             =   "TakeoffItemChart"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Create/Edit Chart"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Change"
            Key             =   "MassChange"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pricelist"
            Key             =   "PricelistExport"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Import"
            Key             =   "Import"
            Object.ToolTipText     =   "Import from Pipeline"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ImportModels"
                  Text            =   "Models and Options"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ImportGlobals"
                  Text            =   "Global Options"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      BorderStyle     =   1
      ' TAG: Timer control used to defer execution of tasks (e.g., opening attachment dialogs) via mTimerTask
      ' CONVERT: Replace with Task.Delay or InvokeAsync pattern to defer execution in Blazor
      Begin VB.Timer Timer1 
         Left            =   9510
         Top             =   0
      End
   End
   ' TAG: Header frame containing all assembly master fields (community, model, option, assembly, description, checkboxes, billing, pricing, etc.)
   ' CONVERT: Replace with a <div> containing inline-grid layout with three columns for left/middle/right panels
   ' CONVERT: Use form-field CSS classes with <input>, <select>, <textarea>, and <input type="checkbox"> elements
   ' CONVERT: Bind all fields to private C# properties with @bind directives
   Begin VB.Frame HeaderFrame 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   2280
      Left            =   -45
      TabIndex        =   41
      Top             =   585
      Width           =   21432
      ' TAG: Billing frame containing billing factor, billing code, and billing category fields for invoice posting
      ' CONVERT: Replace with a <div> or <fieldset> containing form-field rows for BillingFactor, BillingCode, BillingCategory
      ' CONVERT: Use <input type="number"> for factor and <select> dropdowns for code/category
      ' CONVERT: Show/hide based on mUseBilling flag
      Begin VB.Frame frmBilling 
         Appearance      =   0  'Flat
         Caption         =   " Billing "
         ForeColor       =   &H8000000D&
         Height          =   1245
         Left            =   14880
         TabIndex        =   80
         Top             =   990
         Width           =   3195
         ' TAG: Text input for the billing factor multiplier
         ' CONVERT: Use <input type="number" @bind="BillingFactor" step="0.01"> with disabled state based on mReadOnly
         Begin VB.TextBox txtBillingFactor 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   945
            MaxLength       =   10
            TabIndex        =   33
            Top             =   270
            Width           =   615
         End
         ' TAG: Dropdown combo for selecting the billing cost code
         ' CONVERT: Use <select @bind="BillingCodeValue"> populated from StandardCostCodes query
         Begin HFEst.VBCombo cboBillingCode 
            Height          =   240
            Left            =   945
            TabIndex        =   34
            Top             =   525
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   423
         End
         ' TAG: Dropdown combo for selecting the billing category
         ' CONVERT: Use <select @bind="BillingCategoryValue"> populated from StandardCategories query
         Begin HFEst.VBCombo cboBillingCategory 
            Height          =   240
            Left            =   945
            TabIndex        =   35
            Top             =   780
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   423
         End
         ' TAG: Label for the billing factor field
         ' CONVERT: Use <label> element with text "Factor"
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Factor"
            Height          =   195
            Left            =   435
            TabIndex        =   83
            Top             =   255
            Width           =   450
         End
         ' TAG: Label for the billing category field
         ' CONVERT: Use <label> element with text "Category"
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Category"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   82
            Top             =   810
            Width           =   630
         End
         ' TAG: Label for the billing code field
         ' CONVERT: Use <label> element with text "Code"
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Code"
            Height          =   195
            Index           =   1
            Left            =   495
            TabIndex        =   81
            Top             =   555
            Width           =   375
         End
      End
      ' TAG: Checkbox indicating whether this assembly is a base assembly (not published to sales)
      ' CONVERT: Use <input type="checkbox" @bind="IsBaseAssembly"> with label text
      Begin VB.CheckBox chkIsBaseAssembly 
         Caption         =   "Base assembly. Will not be published to sales."
         Height          =   195
         Left            =   11280
         TabIndex        =   31
         Top             =   1410
         Width           =   3705
      End
      ' TAG: Frame containing series and elevation fields, visible for model and option assembly types
      ' CONVERT: Replace with a conditional <div> shown when mAssemblyType is atModel or atoption
      ' CONVERT: Contains Series dropdown and Elevation text input
      Begin VB.Frame frmSeriesFields 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   525
         Left            =   5760
         TabIndex        =   53
         Top             =   0
         Visible         =   0   'False
         Width           =   3912
         ' TAG: Text input for the elevation value
         ' CONVERT: Use <input type="text" @bind="Elevation"> with disabled state based on mReadOnly
         Begin VB.TextBox txtElevation 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   1104
            MaxLength       =   20
            TabIndex        =   9
            Top             =   270
            Width           =   2400
         End
         ' TAG: Dropdown combo for selecting the series (e.g., Comfort Series, Estate Series)
         ' CONVERT: Use <select @bind="SeriesValue"> populated from tblSeries query
         Begin HFEst.VBCombo cboSeries 
            Height          =   240
            Left            =   1104
            TabIndex        =   8
            Top             =   12
            Width           =   2652
            _ExtentX        =   4683
            _ExtentY        =   423
         End
         ' TAG: Label for the elevation field
         ' CONVERT: Use <label> element with text "Elevation"
         Begin VB.Label lblElevation 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Elevation"
            Height          =   195
            Left            =   405
            TabIndex        =   59
            Top             =   300
            Width           =   660
         End
         ' TAG: Clickable label for the series field that opens the series editor dialog
         ' CONVERT: Use <a> or clickable <label> with @onclick to open series editor
         Begin VB.Label lblSeries 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Series"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   192
            Left            =   600
            TabIndex        =   54
            Top             =   36
            Width           =   432
         End
      End
      ' TAG: Text input for the graphic/image file path associated with the assembly
      ' CONVERT: Use <input type="text" @bind="GraphicPath"> with a browse button
      Begin VB.TextBox txtGraphicPath 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   1344
         MaxLength       =   200
         TabIndex        =   6
         Top             =   1020
         Width           =   4140
      End
      ' TAG: Hidden text input storing the style list ID value
      ' CONVERT: Use a private field (not rendered) to hold this value
      Begin VB.TextBox txtLStyle 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   18975
         MaxLength       =   30
         TabIndex        =   74
         Top             =   420
         Visible         =   0   'False
         Width           =   600
      End
      ' TAG: Hidden text input storing the other list ID value
      ' CONVERT: Use a private field (not rendered) to hold this value
      Begin VB.TextBox txtLOther 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   18975
         MaxLength       =   30
         TabIndex        =   73
         Top             =   900
         Visible         =   0   'False
         Width           =   600
      End
      ' TAG: Hidden text input storing the finish list ID value
      ' CONVERT: Use a private field (not rendered) to hold this value
      Begin VB.TextBox txtLFinish 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   18975
         MaxLength       =   30
         TabIndex        =   72
         Top             =   660
         Visible         =   0   'False
         Width           =   600
      End
      ' TAG: Hidden text input storing the color list ID value
      ' CONVERT: Use a private field (not rendered) to hold this value
      Begin VB.TextBox txtLColor 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   18975
         MaxLength       =   30
         TabIndex        =   71
         Top             =   180
         Visible         =   0   'False
         Width           =   600
      End
      ' TAG: Hidden text input storing the style attribute value
      ' CONVERT: Use a private field (not rendered) to hold this value
      Begin VB.TextBox txtAStyle 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   19590
         MaxLength       =   30
         TabIndex        =   70
         Top             =   420
         Visible         =   0   'False
         Width           =   600
      End
      ' TAG: Hidden text input storing the other attribute value
      ' CONVERT: Use a private field (not rendered) to hold this value
      Begin VB.TextBox txtAOther 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   19590
         MaxLength       =   30
         TabIndex        =   69
         Top             =   900
         Visible         =   0   'False
         Width           =   600
      End
      ' TAG: Hidden text input storing the finish attribute value
      ' CONVERT: Use a private field (not rendered) to hold this value
      Begin VB.TextBox txtAFinish 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   19590
         MaxLength       =   30
         TabIndex        =   68
         Top             =   660
         Visible         =   0   'False
         Width           =   600
      End
      ' TAG: Hidden text input storing the color attribute value
      ' CONVERT: Use a private field (not rendered) to hold this value
      Begin VB.TextBox txtAColor 
         BorderStyle     =   0  'None
         Height          =   230
         Left            =   19590
         MaxLength       =   30
         TabIndex        =   66
         Top             =   180
         Visible         =   0   'False
         Width           =   600
      End
      ' TAG: Frame containing the estimator notes text area, visible for option assembly types
      ' CONVERT: Replace with a <div> containing a <textarea @bind="Notes"> for estimator notes
      Begin VB.Frame frmEstimatorNotes 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   945
         Left            =   11115
         TabIndex        =   64
         Top             =   0
         Width           =   4215
         ' TAG: Multi-line text input for estimator notes
         ' CONVERT: Use <textarea rows="3" @bind="Notes" disabled="@mReadOnly">
         Begin VB.TextBox txtNotes 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   705
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   28
            Top             =   210
            Width           =   4140
         End
         ' TAG: Label for the estimator notes section
         ' CONVERT: Use <label> element with text "Estimator Notes"
         Begin VB.Label Label111 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Estimator Notes"
            Height          =   192
            Index           =   1
            Left            =   48
            TabIndex        =   65
            Top             =   0
            Width           =   1116
         End
      End
      ' TAG: Checkbox indicating whether the assembly is active
      ' CONVERT: Use <input type="checkbox" @bind="IsActive"> with label "Active"
      Begin VB.CheckBox chkActive 
         Caption         =   "Active"
         Height          =   195
         Left            =   3408
         TabIndex        =   4
         Top             =   555
         Width           =   2055
      End
      ' TAG: Checkbox for using normal sales quantity factors
      ' CONVERT: Use <input type="checkbox" @bind="UseNormalSalesQtyFactors">
      Begin VB.CheckBox chkUseNormalSalesQtyFactors 
         Caption         =   "Use normal sales quantity factors."
         Height          =   195
         Left            =   11280
         TabIndex        =   30
         Top             =   1185
         Width           =   3075
      End
      ' TAG: Checkbox indicating takeoff is required on every sale
      ' CONVERT: Use <input type="checkbox" @bind="IsTakeoffRequired">
      Begin VB.CheckBox chkTakeoffRequired 
         Caption         =   "Takeoff is required on every sale."
         Height          =   195
         Left            =   11280
         TabIndex        =   29
         Top             =   960
         Width           =   3075
      End
      ' TAG: Multi-line text input for assembly comments
      ' CONVERT: Use <textarea rows="3" @bind="Comments" disabled="@mReadOnly">
      Begin VB.TextBox txtComments 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   744
         Left            =   1344
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   1275
         Width           =   4140
      End
      ' TAG: Read-only text input displaying the option ID
      ' CONVERT: Use <input type="text" @bind="OptionValue" disabled="@(!_optionEnabled)">
      Begin VB.TextBox txtOption 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   3408
         MaxLength       =   20
         TabIndex        =   2
         Top             =   255
         Width           =   2055
      End
      ' TAG: Text input for the assembly description
      ' CONVERT: Use <input type="text" @bind="Description" disabled="@mReadOnly">
      Begin VB.TextBox txtDescription 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   1344
         MaxLength       =   200
         TabIndex        =   5
         Top             =   765
         Width           =   4140
      End
      ' TAG: Dropdown combo for selecting the model
      ' CONVERT: Use <select @bind="ModelValue" disabled="@(!_modelEnabled)"> populated from tblDBAssemblyMaster distinct models
      Begin HFEst.VBCombo cboModel 
         Height          =   240
         Left            =   1350
         TabIndex        =   1
         Top             =   255
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   423
         Style           =   2
      End
      ' TAG: Dropdown combo for selecting the community
      ' CONVERT: Use <select @bind="CommunityValue" disabled="@(!_communityEnabled)"> populated from tblLocality query
      Begin HFEst.VBCombo cboCommunity 
         Height          =   240
         Left            =   1344
         TabIndex        =   0
         Top             =   0
         Width           =   4188
         _ExtentX        =   7382
         _ExtentY        =   423
         Style           =   2
      End
      ' TAG: Dropdown combo for selecting the pricing community (used to display costs)
      ' CONVERT: Use <select @onchange="OnPriceCommunityChanged"> populated from LoadPricingCommunities query
      Begin HFEst.VBCombo cboPriceCommunity 
         Height          =   240
         Left            =   11130
         TabIndex        =   32
         Top             =   1920
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   423
         Style           =   2
      End
      ' TAG: Read-only text input displaying the assembly code
      ' CONVERT: Use <input type="text" @bind="AssemblyCode" disabled="@(!_assemblyEnabled)">
      Begin VB.TextBox txtAssembly 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   1344
         MaxLength       =   20
         TabIndex        =   3
         Top             =   510
         Width           =   2055
      End
      ' TAG: Frame containing option-specific fields (category, JC extra, qty/UOM, location, checkboxes, etc.)
      ' CONVERT: Replace with a conditional <div> shown when mAssemblyType is not atModel
      ' CONVERT: Contains category dropdown, JC extra input, qty/UOM, location, and option checkboxes
      Begin VB.Frame frmOptionFields 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Caption         =   "1"
         Height          =   1605
         Left            =   5760
         TabIndex        =   49
         Top             =   525
         Visible         =   0   'False
         Width           =   5280
         ' TAG: Checkbox for display total only option
         ' CONVERT: Use <input type="checkbox" @bind="IsDisplayTotalOnly">
         Begin VB.CheckBox chkDisplayTotalOnly 
            Caption         =   "Display total only"
            Height          =   215
            Left            =   3420
            TabIndex        =   27
            Top             =   945
            Width           =   1740
         End
         ' TAG: Checkbox for select by room option
         ' CONVERT: Use <input type="checkbox" @bind="IsSelectByRoom">
         Begin VB.CheckBox chkSelectByRoom 
            Caption         =   "Select by room"
            Height          =   215
            Left            =   3420
            TabIndex        =   26
            Top             =   720
            Width           =   1740
         End
         ' TAG: Checkbox for DC sales only option
         ' CONVERT: Use <input type="checkbox" @bind="IsDCSalesOnly">
         Begin VB.CheckBox chkDCSalesOnly 
            Caption         =   "DC Sales Only"
            Height          =   215
            Left            =   3420
            TabIndex        =   25
            Top             =   495
            Width           =   1740
         End
         ' TAG: Text input for the construction cutoff value
         ' CONVERT: Use <input type="number" @bind="ConstCutoff">
         Begin VB.TextBox txtConstCutoff 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   1104
            MaxLength       =   10
            TabIndex        =   22
            Top             =   765
            Width           =   300
         End
         ' TAG: Text input for the assembly quantity
         ' CONVERT: Use <input type="number" @bind="AssemblyQty">
         Begin VB.TextBox txtQty 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   1104
            MaxLength       =   10
            TabIndex        =   20
            Top             =   510
            Width           =   795
         End
         ' TAG: Checkbox for included option flag
         ' CONVERT: Use <input type="checkbox" @bind="IsIncludedOption">
         Begin VB.CheckBox chkIncludedOption 
            Caption         =   "Included Option"
            Height          =   215
            Left            =   3420
            TabIndex        =   24
            Top             =   270
            Width           =   1740
         End
         ' TAG: Text input for the JC extra field
         ' CONVERT: Use <input type="text" @bind="JCExtra">
         Begin VB.TextBox txtJCExtra 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   1104
            MaxLength       =   10
            TabIndex        =   19
            Top             =   255
            Width           =   1155
         End
         ' TAG: Dropdown combo for selecting the unit of measure
         ' CONVERT: Use <select @bind="UOMValue"> with common UOM options
         Begin HFEst.VBCombo cboUOM 
            Height          =   240
            Left            =   1920
            TabIndex        =   21
            Top             =   510
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   423
         End
         ' TAG: Dropdown combo for selecting the location (room)
         ' CONVERT: Use <select @bind="LocationValue"> populated from tblAreas query
         Begin HFEst.VBCombo cboLocation 
            Height          =   240
            Left            =   1110
            TabIndex        =   23
            Top             =   1020
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   423
            Style           =   2
         End
         ' TAG: Dropdown combo for selecting the option category
         ' CONVERT: Use <select @bind="CategoryValue"> populated from tblCategories query
         Begin HFEst.VBCombo cboCategory 
            Height          =   240
            Left            =   1104
            TabIndex        =   18
            Top             =   0
            Width           =   2652
            _ExtentX        =   4683
            _ExtentY        =   423
         End
         ' TAG: Label for the construction cutoff field
         ' CONVERT: Use <label> element with text "Const Cutoff"
         Begin VB.Label lblConstCutoff 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Const Cutoff"
            Height          =   195
            Left            =   75
            TabIndex        =   78
            Top             =   780
            Width           =   975
         End
         ' TAG: Clickable label to open the assembly attributes editor dialog
         ' CONVERT: Use <a> with @onclick to call FAssemblyAttributes.EditAttributes
         Begin VB.Label lblMoreAttributes 
            AutoSize        =   -1  'True
            Caption         =   "More Attributes..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   1110
            TabIndex        =   67
            Top             =   1275
            Width           =   1200
         End
         ' TAG: Clickable label for the location field that opens the room locations editor
         ' CONVERT: Use <a> with @onclick to open room locations editor dialog
         Begin VB.Label lblLocation 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Location"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   420
            TabIndex        =   63
            Top             =   1035
            Width           =   615
         End
         ' TAG: Separator label between category fields
         ' CONVERT: Use <span> with "/" text
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "/"
            Height          =   195
            Index           =   2
            Left            =   375
            TabIndex        =   57
            Top             =   -15
            Width           =   75
         End
         ' TAG: Clickable label for the option group/category that opens the categories editor
         ' CONVERT: Use <a> with @onclick to open option categories editor
         Begin VB.Label lblGroup 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cat"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   120
            TabIndex        =   56
            Top             =   -15
            Width           =   240
         End
         ' TAG: Label for the Qty/UOM fields
         ' CONVERT: Use <label> element with text "Qty/UOM"
         Begin VB.Label Label111 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Qty/UOM"
            Height          =   195
            Index           =   0
            Left            =   345
            TabIndex        =   52
            Top             =   525
            Width           =   690
         End
         ' TAG: Label for the JC Extra field
         ' CONVERT: Use <label> element with text "JC Extra"
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "JC Extra"
            Height          =   195
            Index           =   0
            Left            =   450
            TabIndex        =   51
            Top             =   270
            Width           =   585
         End
         ' TAG: Clickable label for the subcategory that opens the subcategories editor
         ' CONVERT: Use <a> with @onclick to open option subcategories editor
         Begin VB.Label lblCategory 
            AutoSize        =   -1  'True
            Caption         =   "Sub Cat"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   480
            TabIndex        =   50
            Top             =   -15
            Width           =   570
         End
      End
      ' TAG: Frame containing model-specific fields (style, schedule template, bedrooms, bathrooms, floor area, dimensions, spec doc)
      ' CONVERT: Replace with a conditional <div> shown when mAssemblyType is atModel
      ' CONVERT: Contains style combo/text, schedule template dropdown, dimension inputs, spec document path
      Begin VB.Frame frmModelFields 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1605
         Left            =   5760
         TabIndex        =   46
         Top             =   525
         Visible         =   0   'False
         Width           =   5265
         ' TAG: Text input for the specification document file path
         ' CONVERT: Use <input type="text" @bind="SpecDocument"> with a browse button
         Begin VB.TextBox txtSpecDocument 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   1104
            MaxLength       =   200
            TabIndex        =   17
            Top             =   1020
            Width           =   2412
         End
         ' TAG: Text input for the maximum width dimension
         ' CONVERT: Use <input type="number" @bind="MaxWidth">
         Begin VB.TextBox txtMaxWidth 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   1104
            TabIndex        =   15
            Top             =   765
            Width           =   630
         End
         ' TAG: Text input for the maximum length dimension
         ' CONVERT: Use <input type="number" @bind="MaxLength">
         Begin VB.TextBox txtMaxLength 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   1752
            TabIndex        =   16
            Top             =   765
            Width           =   630
         End
         ' TAG: Text input for the floor area (square footage)
         ' CONVERT: Use <input type="number" @bind="FloorArea">
         Begin VB.TextBox txtFloorArea 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   2400
            TabIndex        =   14
            Top             =   510
            Width           =   930
         End
         ' TAG: Text input for the number of bathrooms
         ' CONVERT: Use <input type="number" @bind="Bathrooms">
         Begin VB.TextBox txtBathrooms 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   1752
            TabIndex        =   13
            Top             =   510
            Width           =   630
         End
         ' TAG: Text input for the number of bedrooms
         ' CONVERT: Use <input type="number" @bind="Bedrooms">
         Begin VB.TextBox txtBedrooms 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   1104
            TabIndex        =   12
            Top             =   510
            Width           =   630
         End
         ' TAG: Dropdown combo for selecting the model style (used when SalesSystem is Builder1440)
         ' CONVERT: Use <select @bind="StyleValue"> with options like Single-Family, Condo, Townhouse
         Begin HFEst.VBCombo cboStyle 
            Height          =   240
            Left            =   1104
            TabIndex        =   10
            Top             =   0
            Width           =   2652
            _ExtentX        =   4683
            _ExtentY        =   423
         End
         ' TAG: Text input for the model style (used when SalesSystem is not Builder1440)
         ' CONVERT: Use <input type="text" @bind="StyleText"> as alternative to cboStyle
         Begin VB.TextBox txtStyle 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   1104
            MaxLength       =   30
            TabIndex        =   38
            Top             =   0
            Width           =   2652
         End
         ' TAG: Dropdown combo for selecting the schedule template
         ' CONVERT: Use <select @bind="ScheduleTemplateValue"> populated from app options
         Begin HFEst.VBCombo cboScheduleTemplate 
            Height          =   240
            Left            =   1110
            TabIndex        =   11
            Top             =   255
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   423
            Style           =   2
         End
         ' TAG: Label for the schedule template field
         ' CONVERT: Use <label> element with text "Sched Tmpl"
         Begin VB.Label lblScheduleTemplate 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Sched Tmpl"
            Height          =   195
            Left            =   195
            TabIndex        =   91
            Top             =   285
            Width           =   855
         End
         ' TAG: Clickable label to open the model dimensions editor
         ' CONVERT: Use <a> with @onclick to call FModelDimensions.ShowForm
         Begin VB.Label lblModelDimensions 
            AutoSize        =   -1  'True
            Caption         =   "Edit Dimensions..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   1125
            TabIndex        =   90
            Top             =   1335
            Width           =   1260
         End
         ' TAG: Label for the spec document field
         ' CONVERT: Use <label> element with text "Spec Doc"
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Spec Doc"
            Height          =   195
            Left            =   330
            TabIndex        =   77
            Top             =   1020
            Width           =   720
         End
         ' TAG: Browse button image for selecting the spec document file
         ' CONVERT: Use a <button> with "" text and @onclick to open file picker
         Begin VB.Image cmdBrowse 
            Height          =   240
            Index           =   1
            Left            =   3540
            Picture         =   "FAssembly.frx":0134
            Top             =   1035
            Width           =   240
         End
         ' TAG: Label for the width/length dimension fields
         ' CONVERT: Use <label> element with text "Width/Length"
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Width/Length"
            Height          =   195
            Left            =   120
            TabIndex        =   76
            Top             =   780
            Width           =   930
         End
         ' TAG: Label for the bedrooms/bathrooms/sqft fields
         ' CONVERT: Use <label> element with text "Bds/Bths/SqFt"
         Begin VB.Label Label177 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Bds/Bths/SqFt"
            Height          =   195
            Left            =   -15
            TabIndex        =   48
            Top             =   525
            Width           =   1065
         End
         ' TAG: Label for the style field
         ' CONVERT: Use <label> element with text "Style"
         Begin VB.Label Label166 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Style"
            Height          =   195
            Left            =   705
            TabIndex        =   47
            Top             =   30
            Width           =   345
         End
      End
      ' TAG: Label displaying the assembly cost (used when components are enabled, index 1)
      ' CONVERT: Use a <span> bound to AssemblyCostLabel
      Begin VB.Label lblAssemblyCost 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   1
         Left            =   12465
         TabIndex        =   89
         Top             =   1695
         Width           =   945
      End
      ' TAG: Label displaying the component cost value
      ' CONVERT: Use a <span> bound to ComponentCostLabel
      Begin VB.Label lblComponentCost 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   16815
         TabIndex        =   88
         Top             =   345
         Width           =   945
      End
      ' TAG: Static label "Component Cost:"
      ' CONVERT: Use <span> with text "Component Cost:"
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Component Cost:"
         Height          =   195
         Index           =   3
         Left            =   15465
         TabIndex        =   87
         Top             =   345
         Width           =   1215
      End
      ' TAG: Static label "Total Cost:"
      ' CONVERT: Use <span> with text "Total Cost:"
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total Cost:"
         Height          =   195
         Index           =   2
         Left            =   15915
         TabIndex        =   86
         Top             =   585
         Width           =   765
      End
      ' TAG: Static label "Assembly Cost:"
      ' CONVERT: Use <span> with text "Assembly Cost:"
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Assembly Cost:"
         Height          =   210
         Index           =   1
         Left            =   15615
         TabIndex        =   85
         Top             =   90
         Width           =   1065
      End
      ' TAG: Label displaying the total cost (assembly + component)
      ' CONVERT: Use a <span> bound to TotalCostLabel
      Begin VB.Label lblTotalCost 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   16815
         TabIndex        =   84
         Top             =   585
         Width           =   945
      End
      ' TAG: Bold label "Components" shown above the components grid
      ' CONVERT: Use <h5> or <strong> element with text "Components"
      Begin VB.Label lblComponents 
         AutoSize        =   -1  'True
         Caption         =   "Components"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   30
         TabIndex        =   79
         Top             =   2070
         Width           =   1050
      End
      ' TAG: Browse button image for selecting the graphic/image file
      ' CONVERT: Use a <button> with "" text and @onclick to open file picker
      Begin VB.Image cmdBrowse 
         Height          =   240
         Index           =   0
         Left            =   5505
         Picture         =   "FAssembly.frx":027E
         Top             =   1020
         Width           =   240
      End
      ' TAG: Label for the image/graphic path field
      ' CONVERT: Use <label> element with text "Image"
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Image"
         Height          =   195
         Left            =   810
         TabIndex        =   75
         Top             =   990
         Width           =   450
      End
      ' TAG: Label displaying the assembly cost value (index 0, used with component cost breakdown)
      ' CONVERT: Use a <span> bound to AssemblyCostLabel
      Begin VB.Label lblAssemblyCost 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   0
         Left            =   16815
         TabIndex        =   60
         Top             =   90
         Width           =   945
      End
      ' TAG: Static label "Show pricing for:" preceding the pricing community dropdown
      ' CONVERT: Use <span> with text "Show pricing for:"
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Show pricing for:"
         Height          =   195
         Index           =   0
         Left            =   11130
         TabIndex        =   58
         Top             =   1695
         Width           =   1185
      End
      ' TAG: Bold label for the Community field
      ' CONVERT: Use <label> with bold font and text from custom descriptions
      Begin VB.Label lblCommunity 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Community"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   360
         TabIndex        =   55
         Top             =   36
         Width           =   912
      End
      ' TAG: Label for the Comments field
      ' CONVERT: Use <label> element with text "Comments"
      Begin VB.Label Label111 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Comments"
         Height          =   210
         Index           =   10
         Left            =   480
         TabIndex        =   45
         Top             =   1245
         Width           =   795
      End
      ' TAG: Label for the Description field
      ' CONVERT: Use <label> element with text "Description"
      Begin VB.Label Label133 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Description"
         Height          =   195
         Left            =   480
         TabIndex        =   44
         Top             =   795
         Width           =   795
      End
      ' TAG: Bold label for the Model/Option field
      ' CONVERT: Use <label> with bold font, text changes based on assembly type
      Begin VB.Label Label144 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Model/Option"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   43
         Top             =   300
         Width           =   1170
      End
      ' TAG: Bold label for the Assembly field
      ' CONVERT: Use <label> with bold font and text "Assembly"
      Begin VB.Label Label131 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Assembly"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   480
         TabIndex        =   42
         Top             =   555
         Width           =   795
      End
   End
   ' TAG: Main VSFlexGrid displaying assembly detail items (phases, items, quantities, prices, vendors, WBS codes, etc.)
   ' CONVERT: Use <HfGrid TValue="AssemblyItemRow"> with HfGridColumn components for all item fields
   ' CONVERT: Load column layout via MMain.IniGetGridAsync("FAssembly", "gItems", 0, staticNames: true)
   ' CONVERT: Support grouping, subtotals, inline editing, context menu, and owner-draw via HfGrid features
   ' CONVERT: Use HfGridEvents for OnCellSave, RowSelected callbacks
   Begin VSFlex8Ctl.VSFlexGrid gItems 
      Height          =   2655
      Left            =   75
      TabIndex        =   37
      Top             =   4620
      Width           =   13515
      _cx             =   1990876447
      _cy             =   1990857291
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   14336431
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   255
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   37
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FAssembly.frx":03C8
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   5
      OutlineCol      =   1
      Ellipsis        =   0
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   1
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   2
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   1
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      ' TAG: Picture box used as a tooltip/warning message popup overlay on the items grid
      ' CONVERT: Replace with a positioned <div> tooltip overlay shown/hidden via bool flag
      Begin VB.PictureBox picWarningMessages 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   615