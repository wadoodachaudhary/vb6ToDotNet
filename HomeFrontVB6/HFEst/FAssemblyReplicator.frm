VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form FAssemblyReplicator 
   Caption         =   "Division Replication"
   ClientHeight    =   8880
   ClientLeft      =   3210
   ClientTop       =   2670
   ClientWidth     =   11505
   Icon            =   "FAssemblyReplicator.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   11505
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Index           =   1
      Left            =   10140
      Picture         =   "FAssemblyReplicator.frx":000C
      TabIndex        =   7
      ToolTipText     =   "Cancel"
      Top             =   8370
      Width           =   1215
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   8820
      Picture         =   "FAssemblyReplicator.frx":0596
      TabIndex        =   6
      ToolTipText     =   "Login"
      Top             =   8370
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Replicate Item DB"
      Enabled         =   0   'False
      Height          =   195
      Left            =   5490
      TabIndex        =   4
      Top             =   1680
      Value           =   1  'Checked
      Width           =   2805
   End
   Begin HFEst.WizHead WizHead1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   1588
      Caption         =   "Assembly Replication Wizard"
      Description     =   "Use the Assembly Replication Wizard to quickly setup a new division or community by copying the assemblies from another."
      Icon            =   "FAssemblyReplicator.frx":0B20
   End
   Begin HFEst.VBCombo cboSourceDivision 
      Height          =   240
      Left            =   1935
      TabIndex        =   0
      Top             =   1050
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   423
      Style           =   2
   End
   Begin HFEst.VBCombo cboSourceCommunity 
      Height          =   240
      Left            =   1935
      TabIndex        =   1
      Top             =   1320
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   423
      Style           =   2
   End
   Begin HFEst.VBCombo cboAssemblyType 
      Height          =   240
      Left            =   1935
      TabIndex        =   2
      Top             =   1590
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   423
      Style           =   2
   End
   Begin HFEst.VBCombo cboDestinationCommunity 
      Height          =   240
      Left            =   5310
      TabIndex        =   3
      Top             =   1320
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   423
      Style           =   2
   End
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   5835
      Left            =   120
      TabIndex        =   5
      Top             =   2415
      Width           =   11265
      _cx             =   19870
      _cy             =   10292
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FAssemblyReplicator.frx":13FA
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Assemblies"
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
      Index           =   4
      Left            =   135
      TabIndex        =   13
      Top             =   2160
      Width           =   945
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Destination Community"
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
      Index           =   3
      Left            =   5280
      TabIndex        =   11
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Assembly Type"
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
      Index           =   2
      Left            =   585
      TabIndex        =   10
      Top             =   1620
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Source Community"
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
      Index           =   1
      Left            =   285
      TabIndex        =   9
      Top             =   1350
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Source Division"
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
      Index           =   0
      Left            =   510
      TabIndex        =   8
      Top             =   1080
      Width           =   1350
   End
End
Attribute VB_Name = "FAssemblyReplicator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FAssemblyReplicator::"



Private Sub cmdNav_Click(Index As Integer)
On Error GoTo eh
    
    If Index = 0 Then
        Call SaveData
    Else
        Unload Me
    End If

Exit Sub
eh: Call errHandler(SRCFILE & "cmdNav_Click")
End Sub

Private Sub Form_Load()
On Error GoTo eh
    Dim s As String
    
    IniGetForm Me
    
    'load source divisions
    s = "select divisionname,'',divisionid from divisions order by 1"
    Call LoadComboBox(cboSourceDivision, HFApp.Databases(dbHomefront), s)
    
    'load destination communities
    s = ""
    s = s & "select '(Global)','',0" & vbCrLf
    s = s & "union all" & vbCrLf
    s = s & "select community + ' - ' + c.description,community,0" & vbCrLf
    s = s & "from divisioncommunities d" & vbCrLf
    s = s & "join tbllocality c on c.area=d.community" & vbCrLf
    s = s & "where d.divisionid=" & HFApp.DivisionID & vbCrLf
    s = s & "order by 1"
    Call LoadComboBox(cboDestinationCommunity, HFApp.Databases(dbHomefront), s)
    
    'load assembly types
    cboAssemblyType.Clear
    cboAssemblyType.AddItem "Models and Options"
    cboAssemblyType.AddItem "Global Options"
    cboAssemblyType.AddItem "Design Center Options"
    
    cboDestinationCommunity.ListIndex = 0

Exit Sub
eh: Call errHandler(SRCFILE & "Form_Load")
End Sub

Private Sub cboSourceDivision_Click()
    Dim s As String
    
    'load source communities
    s = ""
    s = s & "select '(Global)','',0" & vbCrLf
    s = s & "union all" & vbCrLf
    s = s & "select community + ' - ' + c.description,community,0" & vbCrLf
    s = s & "from divisioncommunities d" & vbCrLf
    s = s & "join tbllocality c on c.area=d.community" & vbCrLf
    s = s & "where d.divisionid=" & DbQuote(Num, GetComboBoxListID(cboSourceDivision)) & vbCrLf
    s = s & "order by 1"
    Call LoadComboBox(cboSourceCommunity, HFApp.Databases(dbHomefront), s)
    
    
    cboSourceCommunity.ListIndex = -1
    cboAssemblyType.ListIndex = -1

End Sub

Private Sub cboSourceCommunity_Click()
    Call LoadAssemblies
End Sub

Private Sub cboAssemblyType_Click()
    Call LoadAssemblies
End Sub



Private Sub Form_Resize()
On Error Resume Next
    Const margin = 120
    gData.Move margin, gData.Top, Me.ScaleWidth - 2 * margin, Me.ScaleHeight - gData.Top - 2 * margin - cmdNav(1).Height
    cmdNav(0).Move Me.ScaleWidth - margin - 2 * cmdNav(1).Width, Me.ScaleHeight - margin - cmdNav(1).Height
    cmdNav(1).Move Me.ScaleWidth - margin - 1 * cmdNav(1).Width, Me.ScaleHeight - margin - cmdNav(1).Height
End Sub

Private Sub LoadAssemblies()
On Error GoTo eh

    Dim s As String
    
    gData.Rows = 1
    
    s = ""
    s = s & "select 0 Selected,Community,Model,OptionID,Assembly,Description,AssemblyID" & vbCrLf
    s = s & "from tbldbassemblymaster" & vbCrLf
    s = s & "where divisionid = " & DbQuote(Num, GetComboBoxListID(cboSourceDivision)) & vbCrLf
    s = s & "and isnull(community,'') = " & DbQuote(Str, GetComboBoxListKey(cboSourceCommunity)) & vbCrLf
    s = s & "and isnull(isBaseAssembly,0)=0" & vbCrLf
    Select Case cboAssemblyType.ListIndex
    Case 0 'models and options
        s = s & "and assemblytype in(0,2)" & vbCrLf
    Case 1 'globals
        s = s & "and assemblytype =" & atGlobal & vbCrLf
    Case 2 'dcoptions
        s = s & "and assemblytype =" & atDesignCenter & vbCrLf
    Case Else 'not selected
        Exit Sub
    End Select
    s = s & "order by 2,3,4"
    
    
    gData.DataMode = flexDMBoundNoRowCount
    Set gData.DataSource = HFApp.SqlExec(s)
    
    
    'add filter row
    Call gData.AddItem("", 1)
    gData.Cell(flexcpBackColor, 1, 0, 1, gData.Cols - 1) = vbInfoBackground
    gData.FrozenRows = 1

Exit Sub
eh: Call errHandler(SRCFILE & "LoadAssemblies")
End Sub

Private Sub SaveData()
On Error GoTo eh

    Dim i As Long
    Dim SrcDiv As Integer
    Dim SrcAssemblyIDs As String
    Dim SrcComponentIDs As String
    Dim DstCommunity As String
    
    'get selection
    SrcDiv = GetComboBoxListID(Me.cboSourceDivision)
    DstCommunity = GetComboBoxListKey(cboDestinationCommunity)
    With gData
    For i = 2 To .Rows - 1
        If Not .RowHidden(i) And .Cell(flexcpChecked, i, .ColIndex("selected")) = flexChecked Then
            SrcAssemblyIDs = SrcAssemblyIDs & "," & .TextMatrix(i, .ColIndex("AssemblyID"))
        End If
    Next
    End With
    SrcAssemblyIDs = Mid(SrcAssemblyIDs, 2)
    If SrcAssemblyIDs = "" Then
        MsgBox "No assemblies are selected", vbExclamation, App.ProductName
        Exit Sub
    End If
    
    
    
    'must also replicate components if destination division is different than source division
    If HFApp.DivisionID <> SrcDiv Then SrcComponentIDs = GetComponentIDs(SrcAssemblyIDs)
    
    'replicate data
    Call CopyItemDB(SrcDiv)
    Call CopyAssemblies(SrcAssemblyIDs, SrcComponentIDs, DstCommunity)
    
    MsgBox "Replication Complete", vbInformation, App.ProductName

Exit Sub
eh: Call errHandler(SRCFILE & "SaveData")
End Sub

Private Function GetComponentIDs(ByRef SrcAssemblyIDs As String) As String
    'return list with leading comma  eg: ",id,id,id" or ""
    
    Dim SrcComponentIDs As String
    Dim s As String
    Dim rs As Recordset
    
    s = ""
    s = s & "select distinct ComponentAssemblyID" & vbCrLf
    s = s & "from tblDBAssemblyComponents " & vbCrLf
    s = s & "where parentassemblyid in(" & SrcAssemblyIDs & ")" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    While Not rs.EOF
        SrcComponentIDs = SrcComponentIDs & "," & rs(0)
        rs.MoveNext
    Wend
    
    GetComponentIDs = SrcComponentIDs
    
End Function
Private Sub CopyItemDB(SrcDiv As Integer)
    Dim s As String
    
    s = ""
    s = s & "insert tblphaseitem(DivisionID, Phase, Item, PriceLink, Description, Notes, POIndex, JCCostCode, JCCategory, OrderUOM, TakeoffUOM, UpdateEstimating, UStmp, TStmp, ConversionFactor, Price, TaxGroup, WastePercent, RoundDir, Roundto, PhaseSortOrder, ItemSortOrder, flag, PartNumber, CostCategory, ItemNumber, IsQuote, UseInFieldPO, AltJCCostCode, AltJCCategory, Formula, ModifiedDate, ModifiedBy, PriceModifiedDate, PriceModifiedBy, Location" & vbCrLf
    s = s & "                  , WBS01, WBS02, WBS03, WBS04, WBS05, WBS06, WBS07, WBS08, WBS09, WBS10, WBS11, WBS12, WBS13, WBS14, WBS15, WBS16, WBS17, WBS18, WBS19, WBS20, WBS21, WBS22, WBS23, WBS24, WBS25, WBS26, WBS27, WBS28, WBS29, WBS30, WBS31, WBS32, WBS33, WBS34, WBS35, WBS36, WBS37, WBS38, WBS39, WBS40, OptionID, Color, OptionCategory, RetailPretax, InversePhase, InverseItem, DeleteMe)" & vbCrLf
    s = s & "select " & DbQuote(Num, HFApp.DivisionID) & " DivisionID, src.Phase, src.Item, src.PriceLink, src.Description, src.Notes, src.POIndex, src.JCCostCode, src.JCCategory, src.OrderUOM, src.TakeoffUOM, src.UpdateEstimating, src.UStmp, src.TStmp, src.ConversionFactor, src.Price, src.TaxGroup, src.WastePercent, src.RoundDir, src.Roundto, src.PhaseSortOrder, src.ItemSortOrder, src.flag, src.PartNumber, src.CostCategory, src.ItemNumber, src.IsQuote, src.UseInFieldPO, src.AltJCCostCode, src.AltJCCategory, src.Formula, src.ModifiedDate, src.ModifiedBy, src.PriceModifiedDate, src.PriceModifiedBy, src.Location" & vbCrLf
    s = s & "      , src.WBS01, src.WBS02, src.WBS03, src.WBS04, src.WBS05, src.WBS06, src.WBS07, src.WBS08, src.WBS09, src.WBS10, src.WBS11, src.WBS12, src.WBS13, src.WBS14, src.WBS15, src.WBS16, src.WBS17, src.WBS18, src.WBS19, src.WBS20, src.WBS21, src.WBS22, src.WBS23, src.WBS24, src.WBS25, src.WBS26, src.WBS27, src.WBS28, src.WBS29, src.WBS30, src.WBS31, src.WBS32, src.WBS33, src.WBS34, src.WBS35, src.WBS36, src.WBS37, src.WBS38, src.WBS39, src.WBS40,src.OptionID, src.Color, src.OptionCategory, src.RetailPretax, src.InversePhase, src.InverseItem, src.DeleteMe" & vbCrLf
    s = s & "from tblPhaseItem src" & vbCrLf
    s = s & "left join tblphaseitem dst on(dst.divisionid=" & DbQuote(Num, HFApp.DivisionID) & " and dst.phase=src.phase and dst.item=src.item) " & vbCrLf
    s = s & "where src.DivisionID=" & DbQuote(Num, SrcDiv) & vbCrLf
    s = s & "and dst.divisionid is null" & vbCrLf
    
    s = s & vbCrLf & vbCrLf
    
    s = s & "insert tblEstPhases(DivisionID, Phase, Description, GroupPhase, UpdateEstimating, UStmp, TStmp, SortOrder, GroupPhaseValue, DeleteMe, flag)" & vbCrLf
    s = s & "select " & DbQuote(Num, HFApp.DivisionID) & " DivisionID, src.Phase, src.Description, src.GroupPhase, src.UpdateEstimating, src.UStmp, src.TStmp, src.SortOrder, src.GroupPhaseValue, src.DeleteMe, src.flag" & vbCrLf
    s = s & "from tblEstPhases src" & vbCrLf
    s = s & "left join tblEstPhases dst on(dst.divisionid=" & DbQuote(Num, HFApp.DivisionID) & " and dst.phase=src.phase)" & vbCrLf
    s = s & "where src.DivisionID=" & DbQuote(Num, SrcDiv) & vbCrLf
    s = s & "and dst.divisionid is null" & vbCrLf
    
    s = s & vbCrLf & vbCrLf
    
    s = s & "insert tblpoindex(DivisionID, POIndex, Notes, JCCostCode, JCCategory, UpdateEstimating, ForecastPercent1, ForecastPercent2, ForecastPercent3, ForecastPercent4, ForecastPercent5, ForecastPercent6, ForecastPercent7, ForecastPercent8, ForecastPercent9" & vbCrLf
    s = s & ", ForecastPercent10, ForecastPercent11, ForecastPercent12, StandardText, Description, POGroup, HideQty, HidePrice, TotalOnly, POFormat, SchedTaskID, ReleaseTaskID, PaymentTerm, FOB, ShipVia, Terms, UStmp, TStmp, RetainagePercent, POType, PayPoint1Percent" & vbCrLf
    s = s & ", PayPoint2Percent, PayPoint3Percent, PayPoint4Percent, PayPoint5Percent, PayPoint1SchedTask, PayPoint2SchedTask, PayPoint3SchedTask, PayPoint4SchedTask, PayPoint5SchedTask, RequiresPaymentApproval, BuildProEnabled, MPO, RequireLienRelease)" & vbCrLf
    s = s & "select " & DbQuote(Num, HFApp.DivisionID) & " DivisionID, src.POIndex, src.Notes, src.JCCostCode, src.JCCategory, src.UpdateEstimating, src.ForecastPercent1, src.ForecastPercent2, src.ForecastPercent3, src.ForecastPercent4, src.ForecastPercent5" & vbCrLf
    s = s & ", src.ForecastPercent6, src.ForecastPercent7, src.ForecastPercent8, src.ForecastPercent9, src.ForecastPercent10, src.ForecastPercent11, src.ForecastPercent12, src.StandardText, src.Description, src.POGroup, src.HideQty, src.HidePrice" & vbCrLf
    s = s & ", src.TotalOnly, src.POFormat, src.SchedTaskID, src.ReleaseTaskID, src.PaymentTerm, src.FOB, src.ShipVia, src.Terms, src.UStmp, src.TStmp, src.RetainagePercent, src.POType, src.PayPoint1Percent, src.PayPoint2Percent, src.PayPoint3Percent" & vbCrLf
    s = s & ", src.PayPoint4Percent, src.PayPoint5Percent, src.PayPoint1SchedTask, src.PayPoint2SchedTask, src.PayPoint3SchedTask, src.PayPoint4SchedTask, src.PayPoint5SchedTask, src.RequiresPaymentApproval, src.BuildProEnabled, src.MPO, src.RequireLienRelease" & vbCrLf
    s = s & "from tblPOIndex src" & vbCrLf
    s = s & "left join tblPOIndex dst on(dst.divisionid=" & DbQuote(Num, HFApp.DivisionID) & " and dst.poindex=src.poindex)" & vbCrLf
    s = s & "where src.DivisionID=" & DbQuote(Num, SrcDiv) & vbCrLf
    s = s & "and dst.divisionid is null" & vbCrLf
    
    HFApp.SqlExec s
    
End Sub

Private Sub CopyAssemblies(SrcAssemblyIDs As String, SrcComponentIDs As String, DstCommunity As String)
    Dim s As String


    s = ""
    s = s & "insert tbldbassemblymaster(SourceAssemblyID, DivisionID, Community, Assembly, OptionID, AssemblyType, Model, Elevation, Series, FloorArea, Bedrooms, Bathrooms, Style, Description, Notes, Category, Markup, Roundto, Cost, Pretax, COCost, COPretax, IncludeTax, TaxRate, Total_Selling, Tax, Total_CO_Selling, COTax, UStmp, TStmp, ConstCutoff," & vbCrLf
    s = s & "                           Margin, JCExtra, IncentiveRetail, TakeoffRequired, AssemblyUOM, Pretax2, Pretax3, Pretax4, Pretax5, Pretax6, Pretax7, Pretax8, Pretax9, Pretax10, Tax2, Tax3, Tax4, Tax5, Tax6, Tax7, Tax8, Tax9, Tax10, IncludedInSpec, IncludedInSpec2, IncludedInSpec3, IncludedInSpec4, IncludedInSpec5, IncludedInSpec6, " & vbCrLf
    s = s & "                           IncludedInSpec7, IncludedInSpec8, IncludedInSpec9, IncludedInSpec10, Status, UseNormalSalesQtyFactors, InActive,  Comments, Color, Location, Qty, IncludedOption, ColorListID, StyleListID, FinishListID, OtherListID, StyleValue, FinishValue, OtherValue, GraphicPath, SpecDocument, MaxWidth, " & vbCrLf
    s = s & "                           Maxlength, BillingFactor, BillingCode, BillingCategory, IsBaseAssembly, ScheduleTemplate, DesignCenterSalesOnly, SelectByRoom, DisplayTotalOnly, MainSF, UpperSF, LowerSF, GarageSF)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "select src.AssemblyID, " & DbQuote(Num, HFApp.DivisionID) & " DivisionID, " & DbQuote(Str, DstCommunity) & " Community, src.Assembly, src.OptionID, src.AssemblyType, src.Model, src.Elevation, src.Series, src.FloorArea, src.Bedrooms, src.Bathrooms, src.Style, src.Description, src.Notes, src.Category, src.Markup, src.Roundto, src.Cost, src.Pretax, src.COCost, src.COPretax, src.IncludeTax, src.TaxRate, src.Total_Selling, src.Tax, src.Total_CO_Selling, src.COTax, src.UStmp, src.TStmp, src.ConstCutoff, src." & vbCrLf
    s = s & "       Margin, src.JCExtra, src.IncentiveRetail, src.TakeoffRequired, src.AssemblyUOM, src.Pretax2, src.Pretax3, src.Pretax4, src.Pretax5, src.Pretax6, src.Pretax7, src.Pretax8, src.Pretax9, src.Pretax10, src.Tax2, src.Tax3, src.Tax4, src.Tax5, src.Tax6, src.Tax7, src.Tax8, src.Tax9, src.Tax10, src.IncludedInSpec, src.IncludedInSpec2, src.IncludedInSpec3, src.IncludedInSpec4, src.IncludedInSpec5, src.IncludedInSpec6, src." & vbCrLf
    s = s & "       IncludedInSpec7, src.IncludedInSpec8, src.IncludedInSpec9, src.IncludedInSpec10, src.Status, src.UseNormalSalesQtyFactors, src.InActive, src.Comments, src.Color, src.Location, src.Qty, src.IncludedOption, src.ColorListID, src.StyleListID, src.FinishListID, src.OtherListID, src.StyleValue, src.FinishValue, src.OtherValue, src.GraphicPath, src.SpecDocument, src.MaxWidth, src." & vbCrLf
    s = s & "       Maxlength, src.BillingFactor, src.BillingCode, src.BillingCategory, src.IsBaseAssembly, src.ScheduleTemplate, src.DesignCenterSalesOnly, src.SelectByRoom, src.DisplayTotalOnly, src.MainSF, src.UpperSF, src.LowerSF, src.GarageSF" & vbCrLf
    s = s & "from tblDBAssemblyMaster src" & vbCrLf
    s = s & "left join tblDBAssemblyMaster dst on dst.Divisionid=" & DbQuote(Num, HFApp.DivisionID) & " and dst.community=" & DbQuote(Str, DstCommunity) & " and dst.Assembly=src.Assembly and dst.model=Src.Model and dst.OptionID=src.OptionID" & vbCrLf
    s = s & "where dst.assemblyid is null" & vbCrLf
    'SrcComponentIDs will be formatted as ",id,id,id" or "" so is safe to just concatenate the lists together
    s = s & "and src.assemblyid in(" & SrcAssemblyIDs & SrcComponentIDs & ")" & vbCrLf
    
    s = s & vbCrLf & vbCrLf
    
    s = s & "insert tblDBAssemblyDetails(DivisionID, Community, AssemblyID, Assembly, OptionID, SourceSeq, Phase, Item, Vendor, Quantity, Rate, Cost, Notes, UStmp, TStmp, TakeoffQty, OrderQty, Model, Formula, ItemChart, Location, WBS01, WBS02, WBS03, WBS04, WBS05, WBS06, WBS07, WBS08, WBS09, WBS10, WBS11, WBS12, WBS13, WBS14, WBS15, WBS16, WBS17, WBS18, WBS19, " & vbCrLf
    s = s & "                            WBS20, WBS21, WBS22, WBS23, WBS24, WBS25, WBS26, WBS27, WBS28, WBS29, WBS30, WBS31, WBS32, WBS33, WBS34, WBS35, WBS36, WBS37, WBS38, WBS39, WBS40, POIndex, UseModelCost, ConversionFactor, Invertable)" & vbCrLf
    s = s & "select " & DbQuote(Num, HFApp.DivisionID) & " DivisionID, " & DbQuote(Str, DstCommunity) & " Community, dstM.AssemblyID, src.Assembly, src.OptionID, src.Sequence, src.Phase, src.Item, src.Vendor, src.Quantity, src.Rate, src.Cost, src.Notes, src.UStmp, src.TStmp, src.TakeoffQty, src.OrderQty, src.Model, src.Formula, src.ItemChart, src.Location, src.WBS01, src.WBS02, src.WBS03, src.WBS04, src.WBS05, src.WBS06, src.WBS07, src.WBS08, src.WBS09, src.WBS10, src.WBS11, src.WBS12, src.WBS13, src.WBS14, src.WBS15, src.WBS16, src.WBS17, src.WBS18, src.WBS19, " & vbCrLf
    s = s & "       src.WBS20, src.WBS21, src.WBS22, src.WBS23, src.WBS24, src.WBS25, src.WBS26, src.WBS27, src.WBS28, src.WBS29, src.WBS30, src.WBS31, src.WBS32, src.WBS33, src.WBS34, src.WBS35, src.WBS36, src.WBS37, src.WBS38, src.WBS39, src.WBS40, src.POIndex, src.UseModelCost, src.ConversionFactor, src.Invertable" & vbCrLf
    s = s & "from tblDBAssemblyDetails src" & vbCrLf
    s = s & "join tblDBAssemblyMaster dstM on src.AssemblyID=dstM.SourceAssemblyID" & vbCrLf
    s = s & "left join tblDBAssemblyDetails dst on dst.Divisionid=" & DbQuote(Num, HFApp.DivisionID) & " and dst.community=" & DbQuote(Str, DstCommunity) & " and dst.Assembly=src.Assembly and dst.model=Src.Model and dst.OptionID=src.OptionID and dst.SourceSeq=src.Sequence" & vbCrLf
    s = s & "where dst.divisionid is null" & vbCrLf
    s = s & "and src.assemblyid in(" & SrcAssemblyIDs & SrcComponentIDs & ")" & vbCrLf
    
    s = s & vbCrLf & vbCrLf
    
    'tblDBAssemblyComponents -- left join on dstY and isnull on companetassemblyid is because components dont need to be and are not copied when replicating inside the same division
    s = s & "insert tblDBAssemblyComponents(ParentAssemblyID, ComponentAssemblyID, Qty)" & vbCrLf
    s = s & "select dst.AssemblyID ParentAssemblyID, isnull(dstY.AssemblyID,srcX.ComponentAssemblyID) ComponentAssemblyID, srcX.Qty" & vbCrLf
    s = s & "from tblDBAssemblyMaster dst" & vbCrLf
    s = s & "join tblDBAssemblyComponents srcX on srcX.ParentAssemblyID=dst.SourceAssemblyID" & vbCrLf
    s = s & "left join tblDBAssemblyMaster dstY on dstY.SourceAssemblyID=srcX.ComponentAssemblyID" & vbCrLf
    s = s & "left join tblDBAssemblyComponents dstX on dstX.ParentAssemblyID=dst.AssemblyID and dstX.ComponentAssemblyID=dstY.AssemblyID" & vbCrLf
    s = s & "where dst.SourceAssemblyid in(" & SrcAssemblyIDs & ")" & vbCrLf
    s = s & "and dstX.ParentAssemblyID is null" & vbCrLf

    s = s & vbCrLf & vbCrLf
    
    
    
    HFApp.SqlExec s
    
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error GoTo eh
    IniPutForm Me
Exit Sub
eh: Call errHandler(SRCFILE & "Form_Unload")
End Sub

Private Sub gData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    Dim AllowEdit As Boolean

    Select Case True
        Case Row = 1 And Col = 0:  AllowEdit = True
        Case Row = 1 And Col <> 0: AllowEdit = True
        Case Row > 1 And Col = 0:  AllowEdit = True
        Case Row > 1 And Col <> 0: AllowEdit = False
    End Select
    
    If AllowEdit Then
        Cancel = False
        gData.AutoSearch = flexSearchNone
    Else
        Cancel = True
        gData.AutoSearch = flexSearchFromCursor
    End If
End Sub


Private Sub gData_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim r As Long
    Dim c As Long
    Dim s As String
    
    With gData
        If Row <> 1 Then Exit Sub
        
        If .ColKey(Col) = "Selected" Then
            .Cell(flexcpChecked, 2, Col, .Rows - 1, Col) = .Cell(flexcpChecked, 1, Col)
        Else
            'show rows that match
            .Redraw = flexRDNone
            For r = 2 To .Rows - 1
                .RowHidden(r) = False
                For c = 1 To .Cols - 1
                    If Not (UCase(.Cell(flexcpTextDisplay, r, c)) Like "*" & UCase(.Cell(flexcpTextDisplay, 1, c)) & "*") Then
                        .RowHidden(r) = True
                        Exit For
                    End If
                Next
            Next
            .Redraw = flexRDBuffered
        End If
    End With
End Sub
