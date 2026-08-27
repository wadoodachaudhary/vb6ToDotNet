VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form FAddPricelist 
   Caption         =   "Add Vendor Pricing"
   ClientHeight    =   8745
   ClientLeft      =   1905
   ClientTop       =   1635
   ClientWidth     =   14115
   ControlBox      =   0   'False
   Icon            =   "FAddPricelist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   14115
   Begin HFEst.Slider Slider 
      Height          =   3930
      Left            =   2565
      Top             =   2910
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   6932
   End
   Begin VB.TextBox txtVendor 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1425
      TabIndex        =   0
      Top             =   1260
      Width           =   5895
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "C&lose"
      CausesValidation=   0   'False
      Height          =   315
      Index           =   2
      Left            =   8505
      TabIndex        =   6
      Top             =   7125
      Visible         =   0   'False
      Width           =   1035
   End
   Begin HFEst.VBCombo cboCommunity 
      Height          =   240
      Left            =   1425
      TabIndex        =   1
      Top             =   1515
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   423
      Style           =   2
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&Save"
      Height          =   315
      Index           =   0
      Left            =   6345
      TabIndex        =   4
      Top             =   7125
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Clear List"
      Height          =   315
      Index           =   1
      Left            =   7425
      TabIndex        =   5
      Top             =   7125
      Width           =   1035
   End
   Begin HFEst.WizHead WizHead1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   14115
      _ExtentX        =   24897
      _ExtentY        =   1588
      Caption         =   "Create a Vendor Pricelist"
      Description     =   "Select the vendor and price list information, then selected purchase orders and enter prices."
      Icon            =   "FAddPricelist.frx":000C
   End
   Begin VSFlex8Ctl.VSFlexGrid gItems 
      Height          =   4155
      Left            =   2625
      TabIndex        =   3
      Top             =   2865
      Width           =   6915
      _cx             =   12197
      _cy             =   7329
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   16
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FAddPricelist.frx":08E6
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   6
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   5
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   1
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   -2147483643
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid gPOIndexes 
      Height          =   4155
      Left            =   105
      TabIndex        =   2
      Top             =   2865
      Width           =   2415
      _cx             =   4260
      _cy             =   7329
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FAddPricelist.frx":0B13
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   6
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   5
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   1
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   -2147483643
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Frame fItems 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   300
      Left            =   2640
      TabIndex        =   12
      Top             =   2625
      Width           =   6030
      Begin VB.CheckBox chkNewOnly 
         Caption         =   "show new items only"
         Height          =   195
         Left            =   3240
         TabIndex        =   13
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label lblItems 
         AutoSize        =   -1  'True
         Caption         =   "Item Prices"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   960
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid gAssemblies 
      Height          =   1200
      Left            =   7920
      TabIndex        =   15
      ToolTipText     =   "Add rows to create model or option specific prices"
      Top             =   1230
      Width           =   5610
      _cx             =   9895
      _cy             =   2117
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FAddPricelist.frx":0B43
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   6
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   5
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   1
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   -2147483643
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Image cmdVendor 
      Height          =   240
      Left            =   7335
      Picture         =   "FAddPricelist.frx":0BED
      Top             =   1260
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PO Index"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   6
      Left            =   105
      TabIndex        =   10
      Top             =   2625
      Width           =   795
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Community"
      Height          =   195
      Index           =   0
      Left            =   540
      TabIndex        =   8
      Top             =   1515
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Assemblies"
      Height          =   195
      Index           =   2
      Left            =   7935
      TabIndex        =   9
      Top             =   960
      Width           =   780
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Vendor"
      Height          =   195
      Index           =   4
      Left            =   795
      TabIndex        =   7
      Top             =   1260
      Width           =   510
   End
End
Attribute VB_Name = "FAddPricelist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FAddPriceList::"

Private DefaultVendor As String
Private DefaultCommunityPhase As String

Private Const mcITEM_RENUMBER = 0
Private Const mcITEM_NEWITEM = 1
Private Const mcITEM_DUPLICATE = 2
Private Const mcITEM_VIEWFILES = 4
Private Const mcITEM_INSERTFILE = 5
Private Const mcITEM_LINKTOFILE = 6
Private Const mcITEM_DELETE = 8
Private Const mcITEM_PRICEGROUPS = 10


Public Sub ShowForm(Vendor As String, community As String, CommunityPhase As String)
    Dim rs As Recordset
    
    DefaultVendor = Vendor
    DefaultCommunityPhase = community & Chr(2) & CommunityPhase
    Call ValidateVendor(Vendor)
    
    Me.Show vbModal
End Sub

Private Sub cboCommunity_Click()
        
    If cboCommunity.ListIndex = 0 Then
        'Corporate, no assembly allowed
        gAssemblies.Enabled = False
        Label1(2).Enabled = False
    Else
        gAssemblies.Enabled = True
        Label1(2).Enabled = True
    End If
    
End Sub

Private Sub cboCommunity_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete, vbKeyBack
            cboCommunity.ListIndex = 1
    End Select
End Sub

Private Sub chkNewOnly_Click()
    Call ShowItems
End Sub



Private Sub gAssemblies_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    Dim s As String
    Dim i As Long
    Dim key As String
    Dim multipick As Boolean

    With gAssemblies
    ' can multipick if on last row or on only row
    multipick = (.Row = .Rows - 1) Or (.Row = 0 And .Rows = 2)
    
    ' if you make changes here, also do them in FAssembly.Toolbar_ButtonClick
    If HFApp.Options.ValueByName("BuilderType") = "Commercial" Then
        s = ""
        s = s & "Master Assemblies" & Chr(1) & "select c.area AreaID,c.description " & FMain.CD_Community & ",a.Model,a.Model modelid,a.OptionID,a.Assembly,a.Description,a.assemblytype,0 Source,a.Series,a.Style,a.Elevation from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) where a.DivisionID = " & HFApp.DivisionID & " and isnull(a.inactive,0)=0 and a.assemblytype=0 order by 1,2,3,4,5,6,7" & Chr(0)
        s = s & "Assembly Specific Extras" & Chr(1) & "select ct.Description Category,c.area AreaID,c.description " & FMain.CD_Community & ",a.Model,dm.description ModelDescription,a.Series,a.Model modelid,a.OptionID,a.OptionID [Option],a.Assembly,a.Description,a.assemblytype,0 Source from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) left outer join tblcategories ct on(a.category=ct.category) left outer join distinctmodelsbydivision dm on(a.model=dm.model and a.divisionid=dm.divisionid) where a.DivisionID = " & HFApp.DivisionID & " and isnull(a.inactive,0)=0 and a.assemblytype=2 order by 1,2,3,4,5,6,7" & Chr(0)
        s = s & "Design Assemblies" & Chr(1) & "select ct.Description Category,c.area AreaID,c.description " & FMain.CD_Community & ",a.Model modelid,a.OptionID,a.OptionID [Option],a.Assembly,a.Description,a.assemblytype,0 Source from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) left outer join tblcategories ct on(a.category=ct.category) where a.DivisionID = " & HFApp.DivisionID & " and isnull(a.inactive,0)=0 and a.assemblytype=4  order by 1,2,3,4,5,6,7" & Chr(0)
        s = s & "Global Assemblies" & Chr(1) & "select ct.Description Category,c.area AreaID,c.description " & FMain.CD_Community & ",a.Model modelid,a.OptionID,a.OptionID [Option],a.Assembly,a.Description,a.assemblytype,0 Source from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) left outer join tblcategories ct on(a.category=ct.category) where a.DivisionID = " & HFApp.DivisionID & " and isnull(a.inactive,0)=0 and a.assemblytype=3 order by 1,2,3,4,5,6,7" & Chr(0)
    Else
        s = ""
        s = s & "Models" & Chr(1) & "select c.area AreaID,c.description " & FMain.CD_Community & ",a.Model,a.Model modelid,a.OptionID,a.Assembly,a.Description,a.assemblytype,0 Source,a.Series,a.Style,a.Elevation from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) where a.DivisionID = " & HFApp.DivisionID & " and isnull(a.inactive,0)=0 and a.assemblytype=0 order by 1,2,3,4,5,6,7" & Chr(0)
        s = s & "Model Specific Options" & Chr(1)
            s = s & "select distinct m.Model,m.Description,0 Source,m.Model ModelID" & vbCrLf
            s = s & "from tbldbassemblymaster o" & vbCrLf
            s = s & "join DistinctModelsByDivision m on o.divisionid=m.divisionid and m.model=o.model" & vbCrLf
            s = s & "where o.assemblytype=2" & vbCrLf
            s = s & "and isnull(o.inactive,0)=0 and o.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf & Chr(0)
        s = s & "Design Center Options" & Chr(1) & "select ct.Description Category,c.area AreaID,c.description " & FMain.CD_Community & ",a.Model modelid,a.OptionID,a.OptionID [Option],a.Assembly,a.Description,a.assemblytype,0 Source from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) left outer join tblcategories ct on(a.category=ct.category) where a.DivisionID = " & HFApp.DivisionID & " and isnull(a.inactive,0)=0 and a.assemblytype=4  order by 1,2,3,4,5,6,7" & Chr(0)
        s = s & "Global Options" & Chr(1) & "select ct.Description Category,c.area AreaID,c.description " & FMain.CD_Community & ",a.Model modelid,a.OptionID,a.OptionID [Option],a.Assembly,a.Description,a.assemblytype,0 Source from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) left outer join tblcategories ct on(a.category=ct.category) where a.DivisionID = " & HFApp.DivisionID & " and isnull(a.inactive,0)=0 and a.assemblytype=3 order by 1,2,3,4,5,6,7" & Chr(0)
    End If
    
    If Not FPickList.Choose(HFApp.Databases(dbHomefront), IIf(HFApp.Options.ValueByName("BuilderType") = "Commercial", "Assemblies", "Model and Option"), s, , , , , "areaid,modelid,optionid,assemblytype,Source,Elevation", multipick) Then Exit Sub
    
    'if model option, now pick option
    If FPickList.SelectedView = 2 Then
        s = ""
        s = s & "select ct.Description Category,c.area AreaID,c.description " & FMain.CD_Community & ",a.Model,dm.description ModelDescription,a.Series,a.Model modelid,a.OptionID,a.OptionID [Option],a.Assembly,a.Description,a.assemblytype,0 Source" & vbCrLf
        s = s & "from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) left outer join tblcategories ct on(a.category=ct.category) left outer join distinctmodelsbydivision dm on(a.model=dm.model and a.divisionid=dm.divisionid)" & vbCrLf
        s = s & "where a.DivisionID = " & HFApp.DivisionID & vbCrLf
        s = s & " and isnull(a.inactive,0)=0" & vbCrLf
        s = s & " and a.assemblytype=2" & vbCrLf
        s = s & " and a.model=" & DbQuote(Str, FPickList.SelectedItem("modelid")) & vbCrLf
        s = s & " order by 1,2,3,4,5,6,7"
        If Not FPickList.Choose(HFApp.Databases(dbHomefront), IIf(HFApp.Options.ValueByName("BuilderType") = "Commercial", "Assemblies", "Model and Option"), s, , , , , "areaid,modelid,optionid,assemblytype,Source", multipick) Then Exit Sub
    End If
    
    For i = 1 To FPickList.SelectedItems
    
        key = FPickList.SelectedItem("assembly", i) & Chr(1) & FPickList.SelectedItem("modelid", i) & Chr(1) & FPickList.SelectedItem("optionid", i)
        If .FindRow(key, , .ColIndex("key")) = -1 Then
            
            If .Row = .Rows - 1 Then .AddItem ""
            
            .TextMatrix(.Row, .ColIndex("Assembly")) = FPickList.SelectedItem("assembly", i)
            .TextMatrix(.Row, .ColIndex("Model")) = FPickList.SelectedItem("modelid", i)
            .TextMatrix(.Row, .ColIndex("Option")) = FPickList.SelectedItem("optionid", i)
            .TextMatrix(.Row, .ColIndex("key")) = key
            
        
            s = "SELECT Description FROM DistinctAssemblies WHERE Model=" & DbQuote(Str, FPickList.SelectedItem("modelid", i)) & " and Assembly=" & DbQuote(Str, FPickList.SelectedItem("assembly", i)) & " and DivisionID = " & HFApp.DivisionID
            s = HFApp.SqlExec(s, dbHomefront)(0)
                    
            .TextMatrix(.Row, .ColIndex("Description")) = FPickList.SelectedItem("assembly", i) & " -- " & s
            .Row = .Row + 1
        End If
        
    Next
    End With

    gItems.Rows = 1
    gPOindexes.Cell(flexcpChecked, 0, 0, gPOindexes.Rows - 1, 0) = flexUnchecked


End Sub

Private Sub cmdNav_Click(Index As Integer)
    Select Case Index
        Case 0 'save
            Call SaveData(False)
        
        Case 1 'clear
            gItems.Rows = 1
            gPOindexes.Cell(flexcpChecked, 0, 0, gPOindexes.Rows - 1, 0) = flexUnchecked
            cmdNav(2).Cancel = True
            
        Case 2 'close
            Unload Me
            
    End Select
End Sub

Private Sub cmdVendor_Click()
    Call SetCtrlFocus(txtVendor)
    If FPickList.Choose(HFApp.Databases(dbHomefront), "Vendor", "select vendor_name CompanyName,vendor_id Vendor from tblvendors where inactive=0 and DivisionID = " & HFApp.DivisionID) Then
        txtVendor.Tag = FPickList.SelectedItem("Vendor")
        txtVendor.Text = "" & txtVendor.Tag & " - " & FPickList.SelectedItem("CompanyName")
    End If
End Sub

Private Sub Form_Load()
    Dim s As String
    Dim rs As Recordset

    cmdNav(2).Cancel = True
    
    Call IniGetForm(Me)
    Call IniGetGrid(Me, gItems)
    
    gItems.Rows = 1
       
    'load communities
    s = ""
    s = s & "SELECT c.Area+' -- '+ISNULL(c.Description,c.Area)+' (all phases)'" & vbCrLf
    s = s & "      ,c.Area+char(2),0" & vbCrLf
    s = s & "FROM tblLocality c" & vbCrLf
    s = s & "JOIN DivisionCommunities dc on c.Area = dc.Community" & vbCrLf
    s = s & "WHERE c.Inactive<>1 and dc.DivisionID = " & HFApp.DivisionID & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT c.Area + '.' + p.CommunityPhase + ' -- ' + ISNULL(c.Description + ' ' + p.Description,c.Description)" & vbCrLf
    s = s & "      ,c.Area+char(2)+p.CommunityPhase,0" & vbCrLf
    s = s & "FROM tblLocality c JOIN CommunityPhase p ON(c.Area=p.Community)" & vbCrLf
    s = s & "JOIN DivisionCommunities dc on c.Area = dc.Community" & vbCrLf
    s = s & "WHERE c.Inactive<>1 and dc.DivisionID = " & HFApp.DivisionID & vbCrLf
    s = s & "ORDER BY 1" & vbCrLf
    Call LoadComboBox(cboCommunity, HFApp.Databases(dbHomefront), s)
    
    cboCommunity.AddItem "Global (any " & FMain.CD_Community & ")", 0
    cboCommunity.Tag = "Global (any " & FMain.CD_Community & ")" & Chr(1) & cboCommunity.Tag
    
    cboCommunity.AddItem "Corporate (any Division)", 0
    cboCommunity.Tag = "Corporate (any Division)" & Chr(1) & cboCommunity.Tag
    
    Call SetComboBoxListIndex(cboCommunity, , DefaultCommunityPhase)
    If cboCommunity.ListIndex = -1 Then cboCommunity.ListIndex = 1


    
'    'load default assembly
'    If DefaultAssembly <> "" Then
'        txtAssembly.Tag = DefaultAssembly
'        txtAssembly.Text = DefaultAssembly & " -- " & DefaultAssemblyDesc
'    End If
    
    'load POindexes
    s = "select poindex,description from tblpoindex where DivisionID = " & HFApp.DivisionID & " order by 1"
    Set rs = HFApp.SqlExec(s)
    With gPOindexes
        .Rows = 0
        While Not rs.EOF
            s = "" & rs(0)
            If s <> "" Then
                s = s & " " & rs(1)
                Call .AddItem(s)
                .RowData(.Rows - 1) = "K" & rs(0)
            End If
            Call rs.MoveNext
        Wend
        If .Rows > 0 Then .Cell(flexcpChecked, 0, 0, .Rows - 1, 0) = flexUnchecked
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gItems)
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Const margin = 120
    
    Slider.Height = Me.ScaleHeight - gPOindexes.Top - 2 * margin - cmdNav(0).Height
    Slider.Top = gPOindexes.Top
    
    gPOindexes.Move gPOindexes.Left, gPOindexes.Top, Slider.Left - gPOindexes.Left, Slider.Height
    gItems.Move Slider.Left + Slider.Width, gItems.Top, Me.ScaleWidth - Slider.Left - Slider.Width - margin, gPOindexes.Height
    fItems.Left = gItems.Left
    
    cmdNav(0).Move Me.ScaleWidth - (3 * (margin + cmdNav(0).Width)), Me.ScaleHeight - cmdNav(0).Height - margin
    cmdNav(1).Move Me.ScaleWidth - (2 * (margin + cmdNav(0).Width)), Me.ScaleHeight - cmdNav(0).Height - margin
    cmdNav(2).Move Me.ScaleWidth - (1 * (margin + cmdNav(0).Width)), Me.ScaleHeight - cmdNav(0).Height - margin
    
    cmdNav(0).Visible = True
    cmdNav(1).Visible = True
    cmdNav(2).Visible = True
End Sub



Private Sub gAssemblies_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim r As Long
    With gAssemblies
        If KeyCode <> vbKeyDelete Then Exit Sub
        If .Row = .Rows - 1 Then Exit Sub
        
        'remove prices from grid
        For r = gItems.Rows - 1 To 1 Step -1
            If .TextMatrix(.Row, .ColIndex("key")) = gItems.TextMatrix(r, gItems.ColIndex("RowKey")) Then
                gItems.RemoveItem r
            End If
        Next
            
        'remove assembly from grid
        .RemoveItem
        
        If .Rows = 1 Then
            'you have removed all the rows so restore the unit prices row
            .AddItem ""
            .TextMatrix(0, .ColIndex("Description")) = "Unit Prices"
            .TextMatrix(0, .ColIndex("Assembly")) = ""
            .TextMatrix(0, .ColIndex("Model")) = ""
            .TextMatrix(0, .ColIndex("Option")) = ""
            .TextMatrix(0, .ColIndex("key")) = ""
        
            gItems.Rows = 1
            gPOindexes.Cell(flexcpChecked, 0, 0, gPOindexes.Rows - 1, 0) = flexUnchecked
            
        End If
    End With
End Sub

Private Sub gItems_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
On Error Resume Next
'change selection to only single column
    Static bInHere As Boolean
    If bInHere Then Exit Sub
    bInHere = True
    gItems.ColSel = gItems.Col
    bInHere = False
End Sub

Private Sub gItems_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gItems
        .EditMaxLength = 0
        .ComboList = ""
        Select Case .ColKey(Col)
            Case "Price"
            Case "TaxGroup":         .ComboList = "|..."
            Case "TaxGroupDesc":     .ComboList = "..."
            Case "SKU":              .EditMaxLength = 50
            Case Else:                Cancel = True
        End Select
    End With
End Sub

Private Sub gItems_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    Dim r As Long
    If Button <> vbRightButton Then Exit Sub
    
    Cancel = True
    If gItems.MouseRow = 0 Then
        Call FMain.ShowColumnMenu(gItems)
    Else
        
        FMain.mnuFItemsItemsSub(mcITEM_RENUMBER).Enabled = False
        FMain.mnuFItemsItemsSub(mcITEM_NEWITEM).Enabled = False
        FMain.mnuFItemsItemsSub(mcITEM_DUPLICATE).Enabled = False
        
        FMain.mnuFItemsItemsSub(mcITEM_VIEWFILES).Enabled = False
        
        FMain.mnuFItemsItemsSub(mcITEM_DELETE).Enabled = gItems.Row > 0
        FMain.mnuFItemsItemsSub(mcITEM_PRICEGROUPS).Enabled = False
        PopupMenu FMain.mnuFItemsItems
    End If

End Sub



Private Sub gItems_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim s As String
    With gItems
        Select Case .ColKey(Col)
            Case "TaxGroup", "TaxGroupDesc"
                s = "SELECT TaxGroup,Description FROM TaxGroups where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Tax Groups", s, gItems) Then
                    .Cell(flexcpText, Row, .ColIndex("TaxGroup"), .RowSel, .ColIndex("TaxGroup")) = FPickList.SelectedItem("TaxGroup")
                    .Cell(flexcpText, Row, .ColIndex("TaxGroupDesc"), .RowSel, .ColIndex("TaxGroupDesc")) = FPickList.SelectedItem("Description")
                End If
        End Select
    End With
End Sub

Private Sub gItems_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo eh
    Dim r As Long
        
    Select Case True
        Case KeyCode = vbKeyF And Shift = vbCtrlMask
            Call FFind.ShowForm(gItems)
            
        Case KeyCode = vbKeyDelete And Shift = vbCtrlMask
            With gItems
                If .Row < 1 Then Exit Sub
                For r = Max(.Row, .RowSel) To Min(.Row, .RowSel) Step -1
                    Call .RemoveItem(r)
                Next
            End With
            
    End Select
eh:
End Sub


Private Sub gItems_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim s As String
    With gItems
        s = .EditText
        Select Case .ColKey(Col)
            Case "Price": Cancel = Not IsNumeric(.EditText)
            Case "SKU"
            Case "TaxGroup":          Cancel = Not ValidateField(gItems, s, "Tax Group not found", "SELECT TaxGroup,Description FROM TaxGroups WHERE DivisionID = " & HFApp.DivisionID & " and TaxGroup=" & DbQuote(Str, s), "TaxGroupDesc")
            Case Else:    Cancel = True
        End Select
        .EditText = s
    End With
End Sub

Private Sub gPOIndexes_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Static inHere As Boolean
If inHere Then Exit Sub
inHere = True
Screen.MousePointer = vbHourglass

    Dim r As Long
    Dim po As String
    Dim s As String
    Dim rs As Recordset
    
    Dim Vendor         As String
    Dim community      As String
    Dim CommunityPhase As String
    Dim Model          As String
    Dim Assembly       As String
    Dim OptionID       As String
    
    Vendor = txtVendor.Tag
    s = GetComboBoxListKey(cboCommunity)
    community = Parse(s, 1, Chr(2))
    CommunityPhase = Parse(s, 2, Chr(2))
    If IsIn(community, "Corporate (any Division)", "Global (any " & FMain.CD_Community & ")") Then community = ""

    
    po = Mid(gPOindexes.RowData(Row), 2)
    If gPOindexes.Cell(flexcpChecked, Row, Col) <> flexChecked Then
        'MsgBox "you just UNchecked it"
        With gItems
            For r = .Rows - 1 To 1 Step -1
                If po = .TextMatrix(r, .ColIndex("POIndex")) Then
                    .RemoveItem r
                End If
            Next
        End With
    Else
        'MsgBox "you just checked it"
        With gAssemblies
            s = ""
            For r = 0 To .Rows - 2
                Model = .TextMatrix(r, .ColIndex("Model"))
                Assembly = .TextMatrix(r, .ColIndex("Assembly"))
                OptionID = .TextMatrix(r, .ColIndex("Option"))
                s = s & "UNION ALL" & vbCrLf
                s = s & "SELECT " & DbQuote(Str, .TextMatrix(r, .ColIndex("Description"))) & " AssemblyDescription, " & DbQuote(Str, Model) & " Model, " & DbQuote(Str, OptionID) & " OptionID, " & DbQuote(Str, Assembly) & " Assembly, " & DbQuote(Str, .TextMatrix(r, .ColIndex("Key"))) & " RowKey " & vbCrLf
                s = s & ",i.POIndex,i.Phase,i.Item,i.Description,i.OrderUOM,i.PriceLink,t.TaxGroup,t.Description TaxGroupDesc,i.PartNumber" & vbCrLf
                s = s & ",dbo.Purch_GetItemRate(0,0," & DbQuote(Str, community) & "," & DbQuote(Str, CommunityPhase) & "," & DbQuote(Str, Assembly) & "," & DbQuote(Str, Model) & "," & DbQuote(Str, OptionID) & ",i.phase,i.item,0," & DbQuote(Str, Vendor) & ",getdate()," & HFApp.DivisionID & ") Price" & vbCrLf
                s = s & ",dbo.Purch_GetItemRateQuality(0,0," & DbQuote(Str, community) & "," & DbQuote(Str, CommunityPhase) & "," & DbQuote(Str, Assembly) & "," & DbQuote(Str, Model) & "," & DbQuote(Str, OptionID) & ",i.phase,i.item,0," & DbQuote(Str, Vendor) & ",getdate()," & HFApp.DivisionID & ") Quality" & vbCrLf
                s = s & ",dbo.Purch_GetVendorSKU(" & DbQuote(Str, Vendor) & "," & DbQuote(Str, community) & "," & DbQuote(Str, Assembly) & "," & DbQuote(Str, Model) & ",i.phase,i.item," & HFApp.DivisionID & ") VendorSKU" & vbCrLf
                s = s & "FROM tblPhaseItem i" & vbCrLf
                s = s & "LEFT OUTER JOIN TaxGroups t ON(i.DivisionID = t.DivisionID and i.TaxGroup=t.TaxGroup)" & vbCrLf
                s = s & "WHERE i.DivisionID = " & HFApp.DivisionID & " and POIndex=" & DbQuote(Str, po) & vbCrLf
            Next
        End With
        s = Mid(s, 12)
        Set rs = HFApp.SqlExec(s, dbHomefront)
        With gItems
            .Redraw = flexRDNone
            r = .Rows - 1
            While Not rs.EOF
                r = r + 1
                .AddItem ""
                
                .Cell(flexcpForeColor, r, 0, r, .Cols - 1) = IIf("" & rs("Quality") = "12" Or "" & rs("Quality") = "", HFApp.Options(Format_NonSysRate_ForeColor), vbDefault)
                .Cell(flexcpBackColor, r, 0, r, .Cols - 1) = IIf("" & rs("Quality") = "12" Or "" & rs("Quality") = "", HFApp.Options(Format_NonSysRate_BackColor), vbDefault)
                .Cell(flexcpFontBold, r, 0, r, .Cols - 1) = False
                .Cell(flexcpFontItalic, r, 0, r, .Cols - 1) = False
                If "" & rs("Quality") = "12" Or "" & rs("Quality") = "" Then
                    Select Case HFApp.Options(Format_NonSysRate_FontStyle)
                        Case "Bold"
                            .Cell(flexcpFontBold, r, 0, r, .Cols - 1) = True
                            .Cell(flexcpFontItalic, r, 0, r, .Cols - 1) = False
                        Case "Italic"
                            .Cell(flexcpFontBold, r, 0, r, .Cols - 1) = False
                            .Cell(flexcpFontItalic, r, 0, r, .Cols - 1) = True
                        Case "Bold Italic"
                            .Cell(flexcpFontBold, r, 0, r, .Cols - 1) = True
                            .Cell(flexcpFontItalic, r, 0, r, .Cols - 1) = True
                    End Select
                End If
                
                
                .TextMatrix(r, .ColIndex("Quality")) = "" & rs("Quality")
                
                
                .TextMatrix(r, .ColIndex("AssemblyDescription")) = "" & rs("AssemblyDescription")
                .TextMatrix(r, .ColIndex("Model")) = "" & rs("Model")
                .TextMatrix(r, .ColIndex("Assembly")) = "" & rs("Assembly")
                .TextMatrix(r, .ColIndex("Option")) = "" & rs("OptionID")
                .TextMatrix(r, .ColIndex("RowKey")) = "" & rs("RowKey")
                
                
                .TextMatrix(r, .ColIndex("POIndex")) = "" & rs("POIndex")
                .TextMatrix(r, .ColIndex("Phase")) = "" & rs("Phase")
                .TextMatrix(r, .ColIndex("Item")) = "" & rs("Item")
                .TextMatrix(r, .ColIndex("PriceLink")) = "" & rs("PriceLink")
                .TextMatrix(r, .ColIndex("TaxGroup")) = "" & rs("TaxGroup")
                .TextMatrix(r, .ColIndex("TaxGroupDesc")) = "" & rs("TaxGroupDesc")
                If "" & rs("VendorSKU") <> "" Then
                    .TextMatrix(r, .ColIndex("SKU")) = "" & rs("VendorSKU")
                Else
                    .TextMatrix(r, .ColIndex("SKU")) = "" & rs("PartNumber")
                End If
                .TextMatrix(r, .ColIndex("Description")) = "" & rs("Description")
                .TextMatrix(r, .ColIndex("Price")) = "" & rs("Price")
                .TextMatrix(r, .ColIndex("OrderUOM")) = "" & rs("OrderUOM")
                rs.MoveNext
            Wend
            
            Call ShowItems
            .Redraw = flexRDBuffered
        End With
    End If
    
    If gItems.Rows = 1 Then
        cmdNav(2).Cancel = True
    Else
        cmdNav(1).Cancel = True
    End If
    
Screen.MousePointer = vbDefault
inHere = False
End Sub

Private Sub ShowItems()
    Dim i As Long
    With Me.gItems
        For i = 1 To .Rows - 1
            .RowHidden(i) = .TextMatrix(i, .ColIndex("Quality")) <> "13" And .TextMatrix(i, .ColIndex("Quality")) <> "" And chkNewOnly.value = vbChecked
        Next
    End With
End Sub

Private Sub Slider_Move()
    Call Form_Resize
End Sub

Public Function SaveData(prompt As Boolean) As Boolean
On Error GoTo eh
    Dim i As Long
    Dim s As String
    
    Dim SaveZeros      As Boolean
    
    Dim Vendor         As String
    Dim Corporate      As Boolean
    Dim community      As String
    Dim CommunityPhase As String
    
    If gItems.Rows = 1 Then Exit Function
    If Not ValidateData Then Exit Function
    
    Vendor = txtVendor.Tag
    s = GetComboBoxListKey(cboCommunity)
    community = Parse(s, 1, Chr(2))
    CommunityPhase = Parse(s, 2, Chr(2))
    Corporate = community = "Corporate (any Division)"
    If IsIn(community, "Corporate (any Division)", "Global (any " & FMain.CD_Community & ")") Then community = ""
    
    
    
    'save the vendor prices
    With Me.gItems
    
        For i = 1 To .Rows - 1
            If .ValueMatrix(i, .ColIndex("Price")) = 0 Then
                Select Case MsgBox("This list has items with a price of $0.00." & vbCrLf & vbCrLf & "Do you want to save them?", vbYesNoCancel + vbQuestion + vbDefaultButton2, App.ProductName)
                    Case vbYes
                        SaveZeros = True
                        Exit For
                    Case vbNo
                        SaveZeros = False
                        Exit For
                    Case Else
                        Exit Function
                End Select
            End If
        Next
    
        Screen.MousePointer = vbHourglass
        For i = 1 To .Rows - 1
            
            If .ValueMatrix(i, .ColIndex("Price")) <> 0 Or SaveZeros And .RowHidden(i) = False Then
                
                s = ""
                s = s & "INSERT INTO tblVendorCost(DivisionID,Community,CommunityPhase,Assembly,Model,Phase,Item,Vendor)" & vbCrLf
                s = s & "VALUES(" & IIf(Corporate, 0, HFApp.DivisionID) & vbCrLf
                s = s & "      ," & DbQuote(Str, community) & vbCrLf
                s = s & "      ," & DbQuote(Str, CommunityPhase) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Assembly"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Model"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Phase"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Item"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, Vendor) & ")" & vbCrLf
                On Error Resume Next
                Call HFApp.SqlExec(s)
                On Error GoTo eh
            
                s = ""
                s = s & "UPDATE tblVendorCost" & vbCrLf
                s = s & "SET PriceLink=" & DbQuote(Num, IIf(.TextMatrix(i, .ColIndex("Assembly")) <> "", 0, .ValueMatrix(i, .ColIndex("PriceLink")))) & vbCrLf
                s = s & "   ,Current_Cost=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Price"))) & vbCrLf
                s = s & "   ,Forecast1=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Price"))) & vbCrLf
                s = s & "   ,Forecast2=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Price"))) & vbCrLf
                s = s & "   ,Forecast3=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Price"))) & vbCrLf
                s = s & "   ,Forecast4=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Price"))) & vbCrLf
                s = s & "   ,Forecast5=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Price"))) & vbCrLf
                s = s & "   ,Forecast6=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Price"))) & vbCrLf
                s = s & "   ,Forecast7=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Price"))) & vbCrLf
                s = s & "   ,Forecast8=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Price"))) & vbCrLf
                s = s & "   ,Forecast9=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Price"))) & vbCrLf
                s = s & "   ,Forecast10=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Price"))) & vbCrLf
                s = s & "   ,Forecast11=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Price"))) & vbCrLf
                s = s & "   ,Forecast12=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Price"))) & vbCrLf
                s = s & "   ,TaxGroup=" & DbQuote(Str, .TextMatrix(i, .ColIndex("TaxGroup"))) & vbCrLf
                s = s & "   ,PartNumber=" & DbQuote(Str, .TextMatrix(i, .ColIndex("SKU"))) & vbCrLf
                s = s & "WHERE Community=" & DbQuote(Str, community) & vbCrLf
                s = s & "  AND CommunityPhase=" & DbQuote(Str, CommunityPhase) & vbCrLf
                s = s & "  AND Assembly=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Assembly"))) & vbCrLf
                s = s & "  AND Model=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Model"))) & vbCrLf
                s = s & "  AND Phase=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Phase"))) & vbCrLf
                s = s & "  AND Item=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Item"))) & vbCrLf
                s = s & "  AND Vendor=" & DbQuote(Str, Vendor) & vbCrLf
                s = s & "  AND DivisionID = " & IIf(Corporate, 0, HFApp.DivisionID) & vbCrLf
                Call HFApp.SqlExec(s)
                
            End If
        Next
        Screen.MousePointer = vbDefault
    End With
    
    
    SaveData = True

    
Exit Function
eh: Call errHandler(SRCFILE & "SaveData", s)
End Function





Private Function ValidateData() As Boolean
    Dim s As String
    
    If txtVendor.Text = "" Then s = s & " " & vbBullet & " Vendor is required" & vbCrLf
    
    If s = "" Then
        ValidateData = True
    Else
        MsgBox "Unable to save data" & vbCrLf & vbCrLf & s, vbExclamation, App.ProductName
    End If
    
End Function



Private Function ValidateVendor(Vendor As String) As Boolean
    Dim rs As Recordset
    
    If Vendor = "" Then
        txtVendor.Tag = ""
        txtVendor.Text = ""
        ValidateVendor = True
    Else
        Set rs = HFApp.SqlExec("select vendor_id,vendor_name from tblvendors where inactive=0 and DivisionID = " & HFApp.DivisionID & " and vendor_ID=" & DbQuote(Str, Vendor))
        If rs.EOF Then
            MsgBox "unknown vendor", vbCritical, App.ProductName
            SetCtrlFocus txtVendor
            SelectAll txtVendor
        Else
            txtVendor.Tag = "" & rs(0)
            txtVendor.Text = "" & rs(0) & " - " & rs(1)
            ValidateVendor = True
        End If
    End If
    
End Function

Private Sub txtVendor_GotFocus()
    SelectAll txtVendor
End Sub

Private Sub txtVendor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 And Shift = 0 Then Call cmdVendor_Click
End Sub

Private Sub txtVendor_Validate(Cancel As Boolean)
    Cancel = Not ValidateVendor(Parse(txtVendor.Text, 1, " - "))
End Sub

Public Sub mnuFItemsItemsSub_Click(Index As Integer)
    Select Case Index
        Case mcITEM_DELETE:   Call gItems_KeyDown(vbKeyDelete, vbCtrlMask)
    End Select
End Sub

