VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form FItems 
   Caption         =   "Item Database"
   ClientHeight    =   5280
   ClientLeft      =   7845
   ClientTop       =   1530
   ClientWidth     =   12960
   Icon            =   "FItems.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5280
   ScaleWidth      =   12960
   Begin HFEst.Slider Slider 
      Height          =   3855
      Left            =   2700
      Top             =   600
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   6800
      Max             =   7275
   End
   Begin VSFlex8Ctl.VSFlexGrid gPhases 
      Height          =   2655
      Left            =   60
      TabIndex        =   1
      Top             =   600
      Width           =   2595
      _cx             =   1982796257
      _cy             =   1982796363
      Appearance      =   2
      BorderStyle     =   0
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
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FItems.frx":000C
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   7
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
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
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
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin VB.Image imgShortcut 
         Height          =   240
         Index           =   0
         Left            =   2145
         Picture         =   "FItems.frx":005E
         Stretch         =   -1  'True
         ToolTipText     =   "Expand All"
         Top             =   0
         Width           =   240
      End
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   12960
      _ExtentX        =   22860
      _ExtentY        =   1058
      ButtonWidth     =   1984
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "Save"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "View"
            Key             =   "View"
            Style           =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Import DB"
            Key             =   "Import"
            Object.ToolTipText     =   "Import Estimating Item DB"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Import Items"
            Key             =   "ExcelImport"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Export Items"
            Key             =   "ExcelExport"
            Object.ToolTipText     =   "Export items to Excel"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.Timer Timer1 
         Left            =   11070
         Top             =   90
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid gItems 
      Height          =   2655
      Left            =   2820
      TabIndex        =   2
      Top             =   660
      Width           =   10575
      _cx             =   1982810333
      _cy             =   1982796363
      Appearance      =   2
      BorderStyle     =   0
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
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   36
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FItems.frx":05E8
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
      OwnerDraw       =   0
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
      FrozenRows      =   1
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   12648447
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "FItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FItems::"

Private Type ViewDefs
    Name         As String
    KeyFlds      As String
    SortFlds     As String
    DisplayFlds  As String
    FromWhere    As String
    IconKeys     As String
End Type
Private mDirty                  As Boolean
Private mViews()                As ViewDefs

Private mViewIndex As Long

Private SageEstimatingDB As Connection
Private SageEstimatingDBName As String

Private mTimerTask As String

Private Const mcPHASE_RENAME = 0
Private Const mcPHASE_RENUMBER = 1
Private Const mcPHASE_NEWGROUP = 3
Private Const mcPHASE_NEWPHASE = 4
Private Const mcPHASE_DUPLICATE = 5
Private Const mcPHASE_DELETE = 7

Private Const mcITEM_RENUMBER = 0
Private Const mcITEM_NEWITEM = 1
Private Const mcITEM_DUPLICATE = 2
Private Const mcITEM_VIEWFILES = 4
Private Const mcITEM_DELETE = 8
Private Const mcITEM_PRICEGROUPS = 10

'used purely for debugging
Private mViewQuery As String

Private Sub gItems_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'dirty flags are different here. DIRTY means update HF tables only DDIRTY means update HF and TL

    Dim i As Long, r As Long, c As Long
    
    With gItems
        If Row = 1 Then
            .Redraw = flexRDNone
            For r = 2 To .Rows - 1
                If .RowData(r) <> "Deleted" Then
                    .RowHidden(r) = False
                End If
                'If .EditText <> "" Then
                    For c = 0 To .Cols - 1
                        If Not (UCase(.Cell(flexcpTextDisplay, r, c)) Like "*" & UCase(.Cell(flexcpTextDisplay, 1, c)) & "*") Then
                            
                            .RowHidden(r) = True
                            
                            Exit For
                        End If
                    Next
                'End If
            Next
            .Redraw = flexRDBuffered
            Exit Sub
        End If
    End With
    
    With gItems
        
        
        If Row < Min(.Row, .RowSel) Or Row > Max(.Row, .RowSel) Then
            .Row = Row
        End If
        
        mDirty = True
        
        
        Select Case .ColKey(Col)
    
            Case "Formula", "AltJCCostCode", "AltJCCostCodeDesc", "AltJCCategory", "AltJCCategoryDesc", "TaxGroup", _
                 "JCCostCode", "JCCostCodeDesc", "JCCategory", "JCCategoryDesc", "PartNumber", "IsQuote", "PriceLink", _
                 "UseInFieldPO", "Color", "OptionCategory", "RetailPretax"
                 
                For i = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                    If "" & .RowData(i) = "" Then .RowData(i) = "DIRTY"
                Next
            
            
            Case "POIndexDescription", "POIndex"
                For i = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                    If .TextMatrix(i, .ColIndex("CostCategory")) = "M" Then
                        If "" & .RowData(i) = "" Then .RowData(i) = "DDIRTY"
                    Else
                        If "" & .RowData(i) = "" Then .RowData(i) = "DIRTY"
                    End If
                Next
            
            
            Case Else
                For i = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                    If .RowData(i) <> "Delete" Then
                        .RowData(i) = "DDIRTY"
                    End If
                Next
                
        End Select
    End With
End Sub


Private Sub gItems_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error Resume Next
    With gItems
        If OldRow > 1 Then .Cell(flexcpBackColor, OldRow, 0, OldRow, .Cols - 1) = vbWindowBackground
        If NewRow > 1 Then .Cell(flexcpBackColor, NewRow, 0, NewRow, .Cols - 1) = vbButtonFace
        
        .Cell(flexcpBackColor, 1, 0, 1, .Cols - 1) = vbInfoBackground
        
    End With
    
    
End Sub

Private Sub gItems_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
'On Error Resume Next
'change selection to only single column
    Static bInHere As Boolean
    If bInHere Then Exit Sub
    bInHere = True
    gItems.ColSel = gItems.Col
    bInHere = False
End Sub

Private Sub gItems_AfterSort(ByVal Col As Long, Order As Integer)
gItems.FixedRows = 1
End Sub

Private Sub gItems_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gItems
        .ComboList = ""
        .AutoSearch = flexSearchNone
    If Row = 1 Then
        gItems.ComboList = ""
        Cancel = gItems.ColDataType(Col) = flexDTBoolean
    Else
        Select Case .ColKey(Col)
        
            Case "Phase", "Item", "PhaseDesc"
                Cancel = True
                .AutoSearch = flexSearchFromCursor
            
            
            
            Case "RoundDir", "Price", "ConversionFactor", "WastePercent", "RoundTo", "IsQuote", "UseInFieldPO"
            
            Case "Color"
                .EditMaxLength = 30
                            
            Case "InverseItemDesc"
                .ComboList = "..."
            
            
            Case "OptionID"
                .EditMaxLength = 20
            Case "RetailPretax"
                Cancel = .TextMatrix(Row, .ColIndex("OptionID")) = ""
            Case "OptionCategory"
                .EditMaxLength = 15
                .ComboList = "|..."
                Cancel = .TextMatrix(Row, .ColIndex("OptionID")) = ""
                
                
                
            Case "PartNumber"
                .EditMaxLength = 50
                
            Case "ItemDesc"
                .EditMaxLength = 200
                .ComboList = "|..."
                
            Case "Notes"
                .EditMaxLength = 4000
                .ComboList = "|..."
                
            Case "Formula"
                .ComboList = "|..."
                
            Case "TakeoffUOM", "OrderUOM"
                .EditMaxLength = 10
                .ComboList = "|..."
                
            Case "POIndex", "Phase", "TaxGroup", "JCCostCode", "JCCategory", "AltJCCostCode", "AltJCCategory"
                .ComboList = "|..."
                
            Case "JCCostCodeDesc", "JCCategoryDesc", "POIndexDescription", "AltJCCostCodeDesc", "AltJCCategoryDesc"
                .ComboList = "..."
                .AutoSearch = flexSearchFromCursor
            
            
            Case "Location", "WBS01", "WBS02", "WBS03", "WBS04", "WBS05", "WBS06", "WBS07", "WBS08", "WBS09", "WBS10", "WBS11", "WBS12", "WBS13", "WBS14", "WBS15", "WBS16", "WBS17", "WBS18", "WBS19", "WBS20", "WBS21", "WBS22", "WBS23", "WBS24", "WBS25", "WBS26", "WBS27", "WBS28", "WBS29", "WBS30", "WBS31", "WBS32", "WBS33", "WBS34", "WBS35", "WBS36", "WBS37", "WBS38", "WBS39", "WBS40"
                .EditMaxLength = 50
            
            Case Else
                Cancel = True
        End Select
    End If
    End With
End Sub

Private Sub gItems_BeforeSort(ByVal Col As Long, Order As Integer)
gItems.FixedRows = 2
End Sub

Private Sub gItems_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim Phase As String
    Dim Item As String
    Dim InvPhase As String
    Dim InvItem As String
    

    Dim s As String
    Dim i As Long
    With gItems
    
        Select Case .ColKey(Col)
            
            Case "InverseItemDesc"
                s = "select Phase,Item,Description,POIndex from tblphaseitem where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Inverse Item", s, , , , , , , , "Phase=" & .Cell(flexcpText, .Row, .ColIndex("phase"))) Then
                    Phase = .Cell(flexcpText, .Row, .ColIndex("phase"))
                    Item = .Cell(flexcpText, .Row, .ColIndex("item"))
                    InvPhase = FPickList.SelectedItem("phase")
                    InvItem = FPickList.SelectedItem("item")
                        
                    If Phase = InvPhase And Item = InvItem Then Exit Sub
                        
                        
                    s = ""
                    s = s & "update tblphaseitem set" & vbCrLf
                    s = s & " inversephase=" & DbQuote(Str, InvPhase) & vbCrLf
                    s = s & ",inverseitem=" & DbQuote(Str, InvItem) & vbCrLf
                    s = s & "where divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
                    s = s & "and phase=" & DbQuote(Str, Phase) & vbCrLf
                    s = s & "and item=" & DbQuote(Str, Item) & vbCrLf
                    s = s & "" & vbCrLf
                    s = s & "update tblphaseitem set" & vbCrLf
                    s = s & " inversephase=" & DbQuote(Str, Phase) & vbCrLf
                    s = s & ",inverseitem=" & DbQuote(Str, Item) & vbCrLf
                    s = s & "where divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
                    s = s & "and phase=" & DbQuote(Str, InvPhase) & vbCrLf
                    s = s & "and item=" & DbQuote(Str, InvItem) & vbCrLf
                    s = s & "and isnull(inversephase,'')=''" & vbCrLf
                    s = s & "and isnull(inverseitem,'')=''" & vbCrLf

                    Call HFApp.SqlExec(s)
                                            
                    .Cell(flexcpText, .Row, .ColIndex("InverseItemDesc")) = InvPhase & "/" & InvItem & " - " & FPickList.SelectedItem("Description")
                    Call UpdateInverseItem(InvPhase, InvItem, Phase & "/" & Item & " - " & .Cell(flexcpText, .Row, .ColIndex("itemdesc")))

                    Exit Sub
                    
                End If

            
            
            Case "Formula"
                s = gItems.Text
                If FFormulaEditor.EditFormula("", s) Then
                    .Text = s
                End If
            
            Case "AltJCCostCode"
                s = "SELECT CostCode,Description FROM StandardCostCodes where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Cost Code", s, gItems) Then
                    .Cell(flexcpText, .Row, .ColIndex("AltJCCostCode"), .RowSel, .ColIndex("AltJCCostCode")) = FPickList.SelectedItem("CostCode")
                    .Cell(flexcpText, .Row, .ColIndex("AltJCCostCodeDesc"), .RowSel, .ColIndex("AltJCCostCodeDesc")) = FPickList.SelectedItem("Description")
                End If
                
            Case "AltJCCostCodeDesc"
                s = "SELECT Description,CostCode FROM StandardCostCodes where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Cost Code", s, gItems) Then
                    .Cell(flexcpText, .Row, .ColIndex("AltJCCostCode"), .RowSel, .ColIndex("AltJCCostCode")) = FPickList.SelectedItem("CostCode")
                    .Cell(flexcpText, .Row, .ColIndex("AltJCCostCodeDesc"), .RowSel, .ColIndex("AltJCCostCodeDesc")) = FPickList.SelectedItem("Description")
                End If
            
            Case "AltJCCategory"
                s = "SELECT Category,Description FROM StandardCategories where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Category", s, gItems) Then
                    .Cell(flexcpText, .Row, .ColIndex("AltJCCategory"), .RowSel, .ColIndex("AltJCCategory")) = FPickList.SelectedItem("Category")
                    .Cell(flexcpText, .Row, .ColIndex("AltJCCategoryDesc"), .RowSel, .ColIndex("AltJCCategoryDesc")) = FPickList.SelectedItem("Description")
                End If
                
            Case "AltJCCategoryDesc"
                s = "SELECT Description,Category FROM StandardCategories where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Category", s, gItems) Then
                    .Cell(flexcpText, .Row, .ColIndex("AltJCCategory"), .RowSel, .ColIndex("AltJCCategory")) = FPickList.SelectedItem("Category")
                    .Cell(flexcpText, .Row, .ColIndex("AltJCCategoryDesc"), .RowSel, .ColIndex("AltJCCategoryDesc")) = FPickList.SelectedItem("Description")
                End If
            
            Case "OptionCategory"
                s = "SELECT Category,Description FROM tblCategories"
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Category", s, gItems) Then
                    .Cell(flexcpText, .Row, .ColIndex("OptionCategory"), .RowSel, .ColIndex("OptionCategory")) = FPickList.SelectedItem("Category")
                End If
                
                
            
            Case "JCCostCode"
                s = "SELECT CostCode,Description FROM StandardCostCodes where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Cost Code", s, gItems) Then
                    .Cell(flexcpText, .Row, .ColIndex("JCCostCode"), .RowSel, .ColIndex("JCCostCode")) = FPickList.SelectedItem("CostCode")
                    .Cell(flexcpText, .Row, .ColIndex("JCCostCodeDesc"), .RowSel, .ColIndex("JCCostCodeDesc")) = FPickList.SelectedItem("Description")
                End If
                
            Case "JCCostCodeDesc"
                s = "SELECT Description,CostCode FROM StandardCostCodes where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Cost Code", s, gItems) Then
                    .Cell(flexcpText, .Row, .ColIndex("JCCostCode"), .RowSel, .ColIndex("JCCostCode")) = FPickList.SelectedItem("CostCode")
                    .Cell(flexcpText, .Row, .ColIndex("JCCostCodeDesc"), .RowSel, .ColIndex("JCCostCodeDesc")) = FPickList.SelectedItem("Description")
                End If
            
            Case "JCCategory"
                s = "SELECT Category,Description FROM StandardCategories where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Category", s, gItems) Then
                    .Cell(flexcpText, .Row, .ColIndex("JCCategory"), .RowSel, .ColIndex("JCCategory")) = FPickList.SelectedItem("Category")
                    .Cell(flexcpText, .Row, .ColIndex("JCCategoryDesc"), .RowSel, .ColIndex("JCCategoryDesc")) = FPickList.SelectedItem("Description")
                End If
                
            Case "JCCategoryDesc"
                s = "SELECT Description,Category FROM StandardCategories where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Category", s, gItems) Then
                    .Cell(flexcpText, .Row, .ColIndex("JCCategory"), .RowSel, .ColIndex("JCCategory")) = FPickList.SelectedItem("Category")
                    .Cell(flexcpText, .Row, .ColIndex("JCCategoryDesc"), .RowSel, .ColIndex("JCCategoryDesc")) = FPickList.SelectedItem("Description")
                End If
            
            Case "OrderUOM"
                s = "SELECT DISTINCT OrderUOM Unit FROM tblPhaseItem where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Units", s, gItems) Then
                    .Cell(flexcpText, .Row, Col, .RowSel, Col) = FPickList.SelectedItem("Unit")
                End If
                
            Case "TakeoffUOM"
                s = "SELECT DISTINCT TakeoffUOM Unit FROM tblPhaseItem where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Units", s, gItems) Then
                    .Cell(flexcpText, .Row, Col, .RowSel, Col) = FPickList.SelectedItem("Unit")
                End If
                
            Case "TaxGroup"
                s = "SELECT TaxGroup,Description FROM TaxGroups where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Tax Groups", s, gItems) Then
                    .Cell(flexcpText, .Row, Col, .RowSel, Col) = FPickList.SelectedItem("TaxGroup")
                End If
        
            Case "ItemDesc"
                s = gItems.Text
                If FComments.Edit(s, gItems, , "Description", 200) Then
                    .Text = s
                End If
                
            Case "Notes"
                s = gItems.Text
                If FComments.Edit(s, gItems, , "Comments", 4000) Then
                    .Text = s
                End If
                    
            Case "POIndex"
                s = "SELECT POIndex,Description FROM tblPOIndex where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Purchase Order", s, gItems) Then
                    .Cell(flexcpText, .Row, Col, .RowSel, Col) = FPickList.SelectedItem("POIndex")
                    .Cell(flexcpText, .Row, .ColIndex("POIndexDescription"), .RowSel, .ColIndex("POIndexDescription")) = FPickList.SelectedItem("Description")
                End If
                
            Case "POIndexDescription"
                s = "SELECT Description,POIndex FROM tblPOIndex where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Purchase Order", s, gItems) Then
                    .Cell(flexcpText, .Row, Col, .RowSel, Col) = FPickList.SelectedItem("Description")
                    .Cell(flexcpText, .Row, .ColIndex("POIndex"), .RowSel, .ColIndex("POIndex")) = FPickList.SelectedItem("POIndex")
                End If
                
        End Select
    End With
    Call gItems_AfterEdit(Row, Col)
End Sub

Private Sub gItems_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim r As Long
    Dim i As Long
    Dim s As String
    
    
    Select Case True
        Case KeyCode = vbKeyF And Shift = vbCtrlMask
            Call FFind.ShowForm(gItems)
            
        Case KeyCode = vbKeyInsert And Shift = vbCtrlMask
            Call mnuFItemsItemsSub_Click(mcITEM_NEWITEM)
            
        Case KeyCode = vbKeyDelete And Shift = vbCtrlMask
            Call mnuFItemsItemsSub_Click(mcITEM_DELETE)
            
        Case KeyCode = vbKeyDelete
            With gItems
            Select Case .ColKey(.Col)
                Case "Phase", "Item", "RoundDir", "ConversionFactor", "RoundTo", "PhaseDesc", "PriceLink"
                    'no
                    
                Case "InverseItemDesc"
                    For r = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                        s = ""
                        s = s & "update tblphaseitem " & vbCrLf
                        s = s & "set inversephase='',inverseitem=''" & vbCrLf
                        s = s & "where phase=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Phase"))) & vbCrLf
                        s = s & "and item=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Item"))) & vbCrLf
                        Call HFApp.SqlExec(s)
                        .Text = ""
                    Next
              
                                    
                
                
                Case Else
                    .Text = ""
                    For r = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                        Call gItems_ValidateEdit(r, .Col, False)
                        Call gItems_AfterEdit(r, .Col)
                    Next
            End Select
            End With
            
            
    End Select
EXITSUB: Exit Sub
End Sub

Private Sub gItems_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim s As String
    Dim i As Long
    Dim Description As String
    Dim rs As Recordset
    Dim PriceLink As Long
    If Row = 1 Then Exit Sub
    With gItems
        s = .EditText
        Select Case .ColKey(Col)
            Case "Phase":                   Cancel = Not ValidateField(gItems, s, "Phase not found", "SELECT Phase,Description FROM tblEstPhases WHERE DivisionID = " & HFApp.DivisionID & " and GroupPhase=0 AND Phase=" & DbQuote(Str, s), "PhaseDesc")
            Case "OptionCategory":          Cancel = Not ValidateField(gItems, s, "Category not found", "SELECT Category FROM tblCategories WHERE DivisionID = " & HFApp.DivisionID & " and Category=" & DbQuote(Str, s))
            Case "JCCostCode":              Cancel = Not ValidateField(gItems, s, "Cost Code not found", "SELECT CostCode,Description FROM StandardCostCodes WHERE DivisionID = " & HFApp.DivisionID & " and CostCode=" & DbQuote(Str, s), "JCCostCodeDesc")
            Case "JCCategory":              Cancel = Not ValidateField(gItems, s, "Category not found", "SELECT Category,Description FROM StandardCategories WHERE DivisionID = " & HFApp.DivisionID & " and Category=" & DbQuote(Str, s), "JCCategoryDesc")
            Case "AltJCCostCode":           Cancel = Not ValidateField(gItems, s, "Cost Code not found", "SELECT CostCode,Description FROM StandardCostCodes WHERE DivisionID = " & HFApp.DivisionID & " and CostCode=" & DbQuote(Str, s), "AltJCCostCodeDesc")
            Case "AltJCCategory":           Cancel = Not ValidateField(gItems, s, "Category not found", "SELECT Category,Description FROM StandardCategories WHERE DivisionID = " & HFApp.DivisionID & " and Category=" & DbQuote(Str, s), "AltJCCategoryDesc")
            Case "TaxGroup":                Cancel = Not ValidateField(gItems, s, "Tax Group not found", "SELECT TaxGroup FROM TaxGroups WHERE DivisionID = " & HFApp.DivisionID & " and TaxGroup=" & DbQuote(Str, s))
            Case "ConversionFactor":        s = Val(s): If Val(s) <= 0 Then s = 1
            Case "WastePercent":            s = Abs(Round(Val(s), 0)): If Val(s) >= 100 Then s = 99
            Case "Roundto":                 s = Abs(Val(s))
            Case "POIndex":                 Cancel = Not ValidateField(gItems, s, "PO Index not found", "SELECT POIndex,Description POIndexDescription FROM tblPOIndex WHERE DivisionID = " & HFApp.DivisionID & " and POIndex=" & DbQuote(Str, s), "POIndexDescription")
            
            'for these descriptions the event only gets fired on delete key press
            Case "JCCostCodeDesc":     .TextMatrix(Row, .ColIndex("JCCostCode")) = ""
            Case "JCCategoryDesc":     .TextMatrix(Row, .ColIndex("JCCategory")) = ""
            Case "AltJCCostCodeDesc":  .TextMatrix(Row, .ColIndex("AltJCCostCode")) = ""
            Case "AltJCCategoryDesc":  .TextMatrix(Row, .ColIndex("AltJCCategory")) = ""
            Case "POIndexDescription": .TextMatrix(Row, .ColIndex("POIndex")) = ""
            
            Case "OptionID"
                s = Trim(s)
            
            Case "RetailPretax"
                s = Val(s)
            
            Case "Price"
                s = Val(s)
                If Val(s) < 0 Then s = 0
                PriceLink = Val(.Cell(flexcpData, Row, .ColIndex("PriceLink")))
                If PriceLink <> 0 Then
                    For i = 1 To .Rows - 1
                        If PriceLink = Val(.Cell(flexcpData, i, .ColIndex("PriceLink"))) Then
                            .TextMatrix(i, Col) = s
                        End If
                    Next
                End If
                
        End Select
        .EditText = s
        mDirty = True
    End With
End Sub


Private Sub gPhases_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    Dim KeyFld   As String
    Dim KeyValue As String
    
    With gPhases
    
    
        If ActiveControl Is gPhases Then
        Else
            .SetFocus
            If .MouseRow > -1 Then .Row = .MouseRow
            If .MouseCol > -1 Then .Col = .MouseCol
        End If
        
        If Button = vbRightButton Then
            If .MouseRow >= -1 Then
                .Row = .MouseRow
                KeyValue = Trim(.Cell(flexcpText, Max(.Row, 0), 1))
                KeyFld = Trim(.Cell(flexcpData, Max(.Row, 0), 1))
        
                Select Case KeyFld
                    Case "i.GrpPhase", "Phase", ""
'                        FMain.mnuFItemsPhasesSub(mcPHASE_DUPLICATE).Enabled = KeyFld = "Phase"
                        PopupMenu FMain.mnuFItemsPhases
                End Select
                
            End If
        End If
    End With
End Sub

Private Sub Timer1_Timer()
    Call HFApp.RunTask(mTimerTask)
    Timer1.Enabled = False
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo eh
    
    Dim i As Long
    Dim s As String
    Dim f As Form
    
    Dim Vendor    As String
    Dim community As String
    Dim CommunityPhase As String
    Dim Assembly  As String
    
    Select Case True
            
        Case Button.Key = "Save"
            Call SaveData(False)
        
        Case Button.Key = "View"
            Call Toolbar_ButtonDropDown(Button)


        Case Button.Key = "ExcelImport"
            If HFApp.ImportData("tblPhaseItem", "Phase,ItemNumber,CostCategory", "DB Items", "Description,TakeoffUOM,ConversionFactor,OrderUOM", "Phase_code=phase,Item=,Item_number=ItemNumber,phase_code=Phase") Then
                s = ""
                s = s & "update tblphaseitem" & vbCrLf
                s = s & "set itemsortorder = case when isnumeric(itemnumber)=1 then cast(itemnumber as float) else 0 end" & vbCrLf
                s = s & "   ,phasesortorder = case when isnumeric(phase)=1 then cast(phase as float) else 0 end" & vbCrLf
                s = s & " where DivisionID = " & HFApp.DivisionID
                Call HFApp.SqlExec(s, dbHomefront)
                
                s = ""
                s = s & "insert into tblestphases(DivisionID,phase,description,groupphase,ustmp,tstmp)" & vbCrLf
                s = s & "select distinct i.DivisionID,i.phase,i.phase,0,'ME',getdate()" & vbCrLf
                s = s & "from tblphaseitem i" & vbCrLf
                s = s & "left outer join tblestphases p on(i.DivisionID = p.DivisionID and i.phase=p.phase)" & vbCrLf
                s = s & "where i.DivisionID = " & HFApp.DivisionID & " and p.phase is null" & vbCrLf
                Call HFApp.SqlExec(s, dbHomefront)
                
                Call FixGroupPhaseValue
            
            End If
        
        Case Button.Key = "ExcelExport"
            Call FDataExport.ExportData("SELECT Phase,ItemNumber,CostCategory,PriceLink,Description,Notes,POIndex,JCCostCode,JCCategory,TaxGroup,TakeoffUOM,OrderUOM,ConversionFactor,Price,WastePercent,RoundDir,Roundto,PartNumber,IsQuote FROM tblPhaseItem where DivisionID = " & HFApp.DivisionID & " ORDER BY Phase,ItemNumber,CostCategory", "Export Items to")

            

        'Pipeline import
        Case Button.Key = "Import" And HFApp.Options.ValueByName("EstimatingSystem") = esPipeline
            If Not SaveData(True) Then Exit Sub
            Call ImportPipelineProducts
            Call LoadPhases(True)
            gItems.Rows = 2

        'Sage100 import
        Case Button.Key = "Import" And HFApp.Options(AccountingSystem) = asMasterBuilder
            i = HFApp.SqlExec("select count(*) from prtcls,tkfprt where prtcls.indent=0 and prtcls.recnum=tkfprt.prtcls", dbAccounting)(0)
            If i = 0 Then
                i = MsgBox("This will copy the parts from Master Builder's database into Precision Builder's item database." & vbCrLf & vbCrLf & "Are you sure you want to continue?", vbYesNo + vbQuestion, App.ProductName)
            Else
                i = MsgBox("This will copy the parts from Master Builder's database into Precision Builder's item database." & vbCrLf & vbCrLf & "Master Builder has parts in a top level class. These items will not be available in Precision Builder's item list." & vbCrLf & vbCrLf & "Are you sure you want to continue?", vbExclamation + vbOKCancel, App.ProductName)
            End If
            If i = vbOK Or i = vbYes Then
                Call ImportPhases("Sage100")
                Call ImportItems("Sage100")
                Call LoadPhases(True)
            End If

        'Timberline import
        Case Button.Key = "Import" And HFApp.Databases(dbEstimating).State = adStateOpen
            If vbYes = MsgBox("This will copy the items from your Sage Estimating database into Precision Builder's database." & vbCrLf & vbCrLf & "Are you sure you want to continue?", vbYesNo + vbQuestion, App.ProductName) Then
                                
                Set SageEstimatingDB = New Connection
                                
                'choose what to import from
                SageEstimatingDB.Open HFApp.ConnectionString(dbEstimating)
                s = "select name from sys.databases where name not like('%_AddressBook') and name not in ('master','tempdb','model','msdb'," & DbQuote(Str, HFApp.Options.ValueByName("SageSqlEstDatabase")) & ")"
                If Not FPickList.Choose(SageEstimatingDB, "Standard Database", s) Then Exit Sub
                SageEstimatingDBName = "use " & FPickList.SelectedItem(1) & " "
                
                Call ImportPhases("SageEstimating")
                Call ImportItems("SageEstimating")
                Call LoadPhases(True)
            End If
        
        
        
    End Select
Exit Sub
eh: Call errHandler(SRCFILE & "Toolbar_ButtonClick")
End Sub

Private Sub Toolbar_ButtonDropDown(ByVal Button As MSComctlLib.Button)
On Error GoTo eh
    Dim ParentMenu As Long
    Dim i As Long
    
    Select Case Button.Key
        Case "View"
            With FMain.PopMenu
                ParentMenu = .MenuIndex("mnuPriceListViews")
                Call .ClearSubMenusOfItem(ParentMenu)
                For i = 0 To UBound(mViews)
                    .AddItem mViews(i).Name, "PriceListView" & i, , i, ParentMenu, , i = mViewIndex
                Next
            End With
            PopupMenu FMain.mnuPriceListViews, , Button.Left, Button.Top + Button.Height
    End Select
    Exit Sub
eh: Call errHandler(SRCFILE & "Toolbar_ButtonDropDown")
End Sub
Public Sub mnuPriceListViewsSub_Click(Index As Integer)
    If SaveData(True) Then
        gItems.Rows = 1
        mDirty = False
        Call IniPutGrid(Me, gItems, , mViewIndex)
        mViewIndex = Index
        Call LoadPhases(True)
        Call IniGetGrid(Me, gItems, , , mViewIndex)
    End If
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    Call SetToolbarIcons(Toolbar, FMain.LargeIcons)
    
'    Toolbar.Buttons("Export").Enabled = HFApp.Databases(dbEstimating).State = adStateOpen
    Toolbar.Buttons("Import").Enabled = HFApp.Databases(dbEstimating).State = adStateOpen _
                                       Or HFApp.Options(AccountingSystem) = asMasterBuilder _
                                       Or HFApp.Options.ValueByName("EstimatingSystem") = esPipeline
        
    Toolbar.Buttons("Import").Visible = Toolbar.Buttons("Import").Enabled
 '   Toolbar.Buttons("Export").Visible = Toolbar.Buttons("Export").Enabled
    
    
    mViewIndex = IniGet(AppIni, Me.Name, "mViewIndex", 0)
    Call LoadWBSDescriptions
    Call IniGetGrid(Me, gItems, , , mViewIndex)
    Call IniGetForm(Me)
    Call LoadViews
    Me.Show

End Sub


Private Sub LoadViews()
    Dim i As Long
    ReDim mViews(2) As ViewDefs
  

    mViews(i).Name = "Group and Phase"
    mViews(i).KeyFlds = "GrpPhase, Phase"
    mViews(i).DisplayFlds = "GrpPhase + ' - ' + GrpDesc,Phase + ' - ' + PhaseDesc"
    mViews(i).SortFlds = "GrpSortOrder,PhaseSortOrder,ItemSortOrder"
    mViews(i).FromWhere = " FROM EstimatingItems i LEFT OUTER JOIN PriceGroups ON(PriceLink=PriceGroup and i.DivisionID = PriceGroups.divisionID)"
    mViews(i).IconKeys = "groupphase,phase"
    i = i + 1
            
    mViews(i).Name = "Purchase Order"
    mViews(i).KeyFlds = "POIndex"
    mViews(i).DisplayFlds = "POIndex + ' ' + POIndexDescription"
    mViews(i).SortFlds = "POIndex"
    mViews(i).FromWhere = " FROM EstimatingItems i"
    mViews(i).IconKeys = "purchaseorder"
    i = i + 1
    
    mViews(i).Name = "Cost Code"
    mViews(i).KeyFlds = "JCCostCode"
    mViews(i).DisplayFlds = "ltrim(JCCostCode) + ' - ' + JCCostCodeDesc"
    mViews(i).SortFlds = "ltrim(i.JCCostCode)"
    mViews(i).FromWhere = " FROM EstimatingItems i"
    mViews(i).IconKeys = "costcode"
    i = i + 1
        
    Call LoadPhases(True)

End Sub



Private Sub LoadPhases(ClearTree As Boolean)
    Dim rs As Recordset
    Dim r As Long
    Dim s       As String
    Dim i       As Long
    Dim levels  As Long
    Dim Level   As Long
    Dim n       As VSFlexNode
    Dim WhereClause As String
    Dim KeyFld     As String
    Dim DisplayFld As String
    Dim KeyValue   As String
    Dim DisplayValue   As String
    Dim SortFld    As String
    Dim IconKey As String
    

'PROPERTY                  CONTAINS
'.rowdata()                Where clause
'.cell(flexcpText, r, 0)   DisplayFld Value
'.cell(flexcpData, r, 0)   DisplayFld Field Name (also used as tooltip)
'.Cell(flexcpText, r, 1)   KeyFld Value
'.Cell(flexcpData, r, 1)   KeyFld Field Name

    
    
    With gPhases
        .Redraw = flexRDNone
        .TextMatrix(0, 0) = mViews(mViewIndex).Name
        levels = Parse(mViews(mViewIndex).DisplayFlds)
        
        If ClearTree Then .Rows = 1
        
        If .Rows = 1 Then
            Level = -1
        Else
            Level = .RowOutlineLevel(.Row)
        End If
        
        If Level + 1 = levels Then
            'user has opened the lowest level so do nothing
        Else
            On Error Resume Next
            If .TextMatrix(.GetNodeRow(.Row, flexNTFirstChild), 0) = "dummy" Then
                .RemoveItem .GetNodeRow(.Row, flexNTFirstChild)
            Else
                .Redraw = flexRDBuffered
                Exit Sub
            End If
            On Error GoTo 0
            
            KeyFld = Parse(mViews(mViewIndex).KeyFlds, Level + 2)
            IconKey = Parse(mViews(mViewIndex).IconKeys, Level + 2)
            DisplayFld = Parse(mViews(mViewIndex).DisplayFlds, Level + 2)
            SortFld = Parse(mViews(mViewIndex).SortFlds, Level + 2)
            
            
            s = ""
            s = s & "SELECT DISTINCT " & KeyFld & "," & DisplayFld & vbCrLf
            If KeyFld <> SortFld Then s = s & "," & SortFld & vbCrLf
            s = s & mViews(mViewIndex).FromWhere
            On Error Resume Next
            WhereClause = .RowData(.Row)
            On Error GoTo 0
            If WhereClause <> "" Then
                s = s & " WHERE " & Mid(WhereClause, 6) & " and i.DivisionID = " & HFApp.DivisionID & vbCrLf
            Else
                s = s & " WHERE i.DivisionID = " & HFApp.DivisionID & vbCrLf
            End If
            s = s & " and GrpPhase<>Phase" & vbCrLf 'Added March 15, 2014
            s = s & " ORDER BY " & SortFld
            
            mViewQuery = s
            Set rs = HFApp.SqlExec(s)
            
            While Not rs.EOF
                KeyValue = "" & rs(0)
                If KeyValue <> "" Then
                    
                    DisplayValue = Trim("" & rs(1))
                    
                    If Level = -1 Then
                        r = .Rows
                        Call .AddItem(DisplayValue, r)
                        .Cell(flexcpData, r, 0) = DisplayFld
                        .Cell(flexcpText, r, 1) = KeyValue
                        .Cell(flexcpData, r, 1) = "i." & KeyFld
                        
                        Set .Cell(flexcpPicture, r, 0) = FMain.SmallIcons.ListImages(IconKey).Picture
    
                        
                        .IsSubtotal(r) = True
                        .RowOutlineLevel(r) = Level + 1
                        If Level + 2 <> levels Then
                            Call .GetNode(r).AddNode(flexNTFirstChild, "dummy") ' add dummy child row so grid knows this one can be opened
                        End If
                        Set n = .GetNode(r)
                        n.Expanded = False
                        
                        .RowData(r) = WhereClause & " AND i." & KeyFld & " = " & DbQuote(Str, KeyValue)
                        
                    Else
                        r = .Row
                        
                        If DisplayValue <> "" Then
                            Set n = .GetNode(r).AddNode(flexNTLastChild, DisplayValue)
                            .Cell(flexcpData, n.Row, 0) = DisplayFld
                            .Cell(flexcpText, n.Row, 1) = KeyValue
                            .Cell(flexcpData, n.Row, 1) = KeyFld
                            Set .Cell(flexcpPicture, n.Row, 0) = FMain.SmallIcons.ListImages(IconKey).Picture
                            If Level + 2 < levels Then
                                Call .GetNode(n.Row).AddNode(flexNTFirstChild, "dummy") ' add dummy child row so grid knows this one can be opened
                            End If
                            n.Expanded = False
                            If Left(WhereClause, 13) = " AND GrpPhase" Then
                                .RowData(n.Row) = " AND " & KeyFld & " = " & DbQuote(Str, KeyValue)
                            Else
                                .RowData(n.Row) = WhereClause & " AND i." & Trim(KeyFld) & " = " & DbQuote(Str, KeyValue)
                            End If
                        End If
                    End If
                    
                End If
                rs.MoveNext
            
            Wend
            
        End If
        
        Call .AutoSize(0)
        .Redraw = flexRDBuffered
    End With
End Sub



Private Sub gPhases_BeforeCollapse(ByVal Row As Long, ByVal State As Integer, Cancel As Boolean)
Static bInHere As Boolean
    Dim r As Long
    Dim levels As Long
    Dim Level As Long
    
    With gPhases
        If State = flexOutlineExpanded Then
            .Row = Row
            Call LoadPhases(False)
        End If
    End With

End Sub


Private Sub gPhases_DblClick()
On Error Resume Next
    If gPhases.GetNode.Expanded Or gPhases.GetNode.Children = 0 Then
        Call gPhases_KeyDown(vbKeyReturn, 0)
    Else
        gPhases.IsCollapsed(gPhases.Row) = flexOutlineExpanded
    End If
End Sub

Private Sub gPhases_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim s As String
    Dim i As Long
    If InIde And gPhases.MouseRow = 0 Then
        s = ""
        s = s & "View Definition:" & vbCrLf
        s = s & "------------------------------------" & vbCrLf
        For i = 1 To Parse(mViews(mViewIndex).KeyFlds)
            s = s & String(4 * (i - 1), " ") & Parse(mViews(mViewIndex).KeyFlds, i) & vbCrLf
        Next
        s = s & vbCrLf & vbCrLf
        s = s & "Last Query:" & vbCrLf
        s = s & "------------------------------------" & vbCrLf
        s = s & mViewQuery
        MsgBox s, vbInformation, "View Definition"
    End If
End Sub

Private Sub gPhases_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Dim i As Long
    i = gPhases.MouseRow
    If i = -1 Then
        gPhases.ToolTipText = ""
    Else
        gPhases.ToolTipText = PrettyName(gPhases.Cell(flexcpData, gPhases.MouseRow, 0))
    End If
End Sub

Private Sub Slider_Move()
    Form_Resize
End Sub

Private Sub Form_Resize()
On Error Resume Next

    Slider.Min = 960
    Slider.Max = Me.ScaleWidth - 960
    Slider.Move Slider.Left, Toolbar.Height, Slider.Width, Me.ScaleHeight - Toolbar.Height

    gPhases.Move 0, Slider.Top, Slider.Left, Slider.Height
    gItems.Move Slider.Left + Slider.Width, Slider.Top, Me.ScaleWidth - Slider.Left - Slider.Width, Slider.Height

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not SaveData(True) Then
        Cancel = True
        Exit Sub
    End If
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gItems, , mViewIndex)
    Call IniPut(AppIni, Me.Name, "mViewIndex", mViewIndex)
    Unload FComments
End Sub

Private Sub gPhases_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim Level As Long
    Dim levels As Long
    
    With gPhases
        If .Row < 0 Then Exit Sub
        Select Case KeyCode

            Case vbKeyReturn
            
                levels = Parse(mViews(mViewIndex).DisplayFlds)
                If .Rows > 1 Then
                    Level = .RowOutlineLevel(.Row)
                End If
                .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = False
                .Cell(flexcpFontBold, .Row, 0) = True
                Call .AutoSize(0, .Cols - 1)
                Call LoadItems
                KeyCode = 0

            Case vbKeyLeft
                If .GetNodeRow(.Row, flexNTFirstChild) <> -1 And .IsCollapsed(.Row) <> flexOutlineCollapsed Then
                    .IsCollapsed(.Row) = flexOutlineCollapsed
                Else
                    If .GetNodeRow(.Row, flexNTParent) <> -1 Then .Row = .GetNodeRow(.Row, flexNTParent)
                End If

            Case vbKeyRight
                If .GetNodeRow(.Row, flexNTFirstChild) <> -1 Then
                    If .IsCollapsed(.Row) = flexOutlineCollapsed Then
                        .IsCollapsed(.Row) = flexOutlineExpanded
                    Else
                        .Row = .GetNodeRow(.Row, flexNTFirstChild)
                    End If
                End If

        End Select
    End With
End Sub

Private Sub LoadItems()
On Error GoTo eh
    Dim s As String
    Dim WhereClause As String
    Dim rs As Recordset
    Dim w As Long
    Dim r As Long
    Dim i As Long

Dim maxWBS As Long
maxWBS = Val("" & HFApp.SqlExec("select max(right(Item,2)) from customDescriptions where item like 'WBS__' and Custom_Description<>''")(0))




    If Not SaveData(True) Then Exit Sub
    
    With gItems
        Screen.MousePointer = vbHourglass
        .Redraw = flexRDNone
        
        .Rows = 1 'dont delete this. it clears the filter bar
        .Rows = 2
        
        s = ""
        s = s & "select" & vbCrLf
        s = s & " x.jccostcode ItemCostCode" & vbCrLf
        s = s & ",y.description ItemCostCodeDesc" & vbCrLf
        s = s & ",v.phase + '/'+ v.item + ' - ' + isnull(v.Description,'') InverseItemDesc" & vbCrLf
        s = s & ",i.*,g.* " & vbCrLf
        s = s & "from EstimatingItems i" & vbCrLf
        s = s & "left outer join PriceGroups g ON(i.PriceLink=g.PriceGroup and i.DivisionID = g.DivisionID)" & vbCrLf
        s = s & "left outer join tblphaseitem x on i.DivisionID=x.DivisionID and i.phase=x.phase and i.item=x.item" & vbCrLf
        s = s & "left outer join tblphaseitem v on x.DivisionID=v.DivisionID and x.inversephase=v.phase and x.inverseitem=v.item" & vbCrLf
        s = s & "left outer join standardcostcodes y on i.DivisionID=y.DivisionID and x.jccostcode=y.costcode" & vbCrLf
        On Error Resume Next
        WhereClause = gPhases.RowData(gPhases.Row)
        
        WhereClause = Replace(WhereClause, "and Phase =", "and i.Phase =")
        
        
        On Error GoTo eh
        If WhereClause <> "" Then
            s = s & " WHERE i.DivisionID = " & HFApp.DivisionID & " and " & Mid(WhereClause, 6) & vbCrLf
        Else
            s = s & " WHERE i.DivisionID = " & HFApp.DivisionID
        End If
        s = s & "ORDER BY " & mViews(mViewIndex).SortFlds
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
            If "" & rs("Item") <> "" Then
                .AddItem ""
                r = .Rows - 1
    
               
               
                .Cell(flexcpData, r, .ColIndex("PriceLink")) = "" & rs("PriceLink")
                .Cell(flexcpText, r, .ColIndex("PriceLink")) = "" & rs("PriceGroupDesc")
                If Val(.Cell(flexcpData, r, .ColIndex("PriceLink"))) <> 0 Then
                    .Cell(flexcpPicture, r, .ColIndex("PriceLink")) = FMain.SmallIcons.ListImages.Item("links").Picture
                End If
                
                .TextMatrix(r, .ColIndex("OptionCategory")) = "" & rs("OptionCategory")
                .TextMatrix(r, .ColIndex("OptionID")) = "" & rs("OptionID")
                .TextMatrix(r, .ColIndex("Color")) = "" & rs("Color")
                .TextMatrix(r, .ColIndex("RetailPretax")) = "" & rs("RetailPretax")
                
                .TextMatrix(r, .ColIndex("Formula")) = "" & rs("Formula")
                .TextMatrix(r, .ColIndex("TaxGroup")) = "" & rs("TaxGroup")
                .TextMatrix(r, .ColIndex("ItemNumber")) = "" & rs("ItemNumber")
                .TextMatrix(r, .ColIndex("CostCategory")) = "" & rs("CostCategory")
                .TextMatrix(r, .ColIndex("Item")) = "" & rs("Item")
                .TextMatrix(r, .ColIndex("ItemDesc")) = "" & rs("ItemDesc")
                .TextMatrix(r, .ColIndex("Notes")) = "" & rs("Notes")
                .TextMatrix(r, .ColIndex("Phase")) = "" & rs("Phase")
                .TextMatrix(r, .ColIndex("PartNumber")) = "" & rs("PartNumber")
                .TextMatrix(r, .ColIndex("PhaseDesc")) = "" & rs("PhaseDesc")
                .TextMatrix(r, .ColIndex("POIndex")) = "" & rs("POIndex")
                .TextMatrix(r, .ColIndex("POIndexDescription")) = "" & rs("POIndexDescription")
                .TextMatrix(r, .ColIndex("JCCostCode")) = "" & rs("ItemCostCode")
                .TextMatrix(r, .ColIndex("JCCostCodeDesc")) = "" & rs("ItemCostCodeDesc")
                .TextMatrix(r, .ColIndex("JCCategory")) = "" & rs("JCCategory")
                .TextMatrix(r, .ColIndex("JCCategoryDesc")) = "" & rs("JCCategoryDesc")
                
                .TextMatrix(r, .ColIndex("AltJCCostCode")) = "" & rs("AltJCCostCode")
                .TextMatrix(r, .ColIndex("AltJCCostCodeDesc")) = "" & rs("AltJCCostCodeDesc")
                .TextMatrix(r, .ColIndex("AltJCCategory")) = "" & rs("AltJCCategory")
                .TextMatrix(r, .ColIndex("AltJCCategoryDesc")) = "" & rs("AltJCCategoryDesc")
                
                .TextMatrix(r, .ColIndex("InverseItemDesc")) = "" & rs("InverseItemDesc")
                
                .TextMatrix(r, .ColIndex("ItemDesc")) = "" & rs("ItemDesc")
                .TextMatrix(r, .ColIndex("Price")) = "" & rs("Price")
                .TextMatrix(r, .ColIndex("TakeoffUOM")) = "" & rs("TakeoffUOM")
                .TextMatrix(r, .ColIndex("OrderUOM")) = "" & rs("OrderUOM")
                .TextMatrix(r, .ColIndex("ConversionFactor")) = "" & rs("ConversionFactor")
                .TextMatrix(r, .ColIndex("WastePercent")) = "" & rs("WastePercent")
                .TextMatrix(r, .ColIndex("RoundDir")) = "" & rs("RoundDir")
                .TextMatrix(r, .ColIndex("Roundto")) = "" & rs("RoundTo")
                
                .TextMatrix(r, .ColIndex("IsQuote")) = "" & rs("IsQuote")
                .TextMatrix(r, .ColIndex("UseInFieldPO")) = "" & rs("UseInFieldPO")
                
                .TextMatrix(r, .ColIndex("Location")) = "" & rs("Location")
                
                
                
                For w = 1 To maxWBS
                    .TextMatrix(r, .ColIndex("WBS" & format(w, "00"))) = "" & rs("WBS" & format(w, "00"))
                Next
                
                If .ValueMatrix(r, .ColIndex("ConversionFactor")) <= 0 Then .TextMatrix(r, .ColIndex("ConversionFactor")) = "1"
            End If
            rs.MoveNext
        Wend
        .Redraw = flexRDBuffered
        Screen.MousePointer = vbDefault
    End With
    mDirty = False
Exit Sub
eh: Call errHandler(SRCFILE & "LoadItems", s)
End Sub

Public Function SaveData(prompt As Boolean) As Boolean
On Error GoTo eh
    Dim i As Long
    Dim w As Long
    Dim s As String
    Dim Conversion As Double
    Dim MultDiv As String
    Dim categories As String
    Dim rs As Recordset
    
    If Not mDirty Then
        SaveData = True
        Exit Function
    End If
    If prompt Then
        Select Case MsgBox(Me.Caption & " has changed." & vbCrLf & vbCrLf & "Do you want to save these changes?" & vbCrLf, vbExclamation + vbYesNoCancel, App.ProductName)
            Case vbNo
                SaveData = True
                Exit Function
            Case vbCancel
                SaveData = False
                Exit Function
        End Select
    End If

    Screen.MousePointer = vbHourglass

    With gItems
    
        'make sure final edit to checkbox is accepted
        Call .FinishEditing(False)
        If .Row = 2 Then
            .Row = 0
        Else
            .Row = 2
        End If
    
        For i = .Rows - 1 To 2 Step -1
            If .RowData(i) = "Deleted" Then
            
                Call HFApp.SqlExec("DELETE FROM tblPhaseItem           WHERE Divisionid = " & HFApp.DivisionID & " and Phase=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Phase"))) & " AND Item=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Item"))))
                Call HFApp.SqlExec("DELETE FROM tblDBAssemblyDetails   WHERE DivisionID = " & HFApp.DivisionID & " and Phase=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Phase"))) & " AND Item=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Item"))))
                Call HFApp.SqlExec("DELETE FROM tblSalesSheetCosts     WHERE Phase=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Phase"))) & " AND Item=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Item"))))
                Call HFApp.SqlExec("DELETE FROM tblVendorCost          WHERE DivisionID = " & HFApp.DivisionID & " and Phase=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Phase"))) & " AND Item=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Item"))))
                
                
                .RemoveItem i
            End If
        Next
        
        
        For i = 2 To .Rows - 1
            .Row = i
            
            If IsIn(.RowData(i), "DIRTY", "DDIRTY") Then
                Conversion = .ValueMatrix(i, .ColIndex("ConversionFactor"))
                
                s = ""
                s = s & "UPDATE tblPhaseItem" & vbCrLf
                s = s & "SET PhaseSortOrder=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Phase"))) & vbCrLf
                s = s & "   ,ItemSortOrder=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Item"))) & vbCrLf
                s = s & "   ,Description=" & DbQuote(Str, .TextMatrix(i, .ColIndex("ItemDesc"))) & vbCrLf
                s = s & "   ,Formula=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Formula"))) & vbCrLf
                s = s & "   ,Notes=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Notes"))) & vbCrLf
                s = s & "   ,PartNumber=" & DbQuote(Str, .TextMatrix(i, .ColIndex("PartNumber"))) & vbCrLf
                s = s & "   ,POIndex=" & DbQuote(Str, .TextMatrix(i, .ColIndex("POIndex"))) & vbCrLf
                s = s & "   ,JCCostCode=" & DbQuote(Str, .TextMatrix(i, .ColIndex("JCCostCode"))) & vbCrLf
                s = s & "   ,JCCategory=" & DbQuote(Str, .TextMatrix(i, .ColIndex("JCCategory"))) & vbCrLf
                s = s & "   ,AltJCCostCode=" & DbQuote(Str, .TextMatrix(i, .ColIndex("AltJCCostCode"))) & vbCrLf
                s = s & "   ,AltJCCategory=" & DbQuote(Str, .TextMatrix(i, .ColIndex("AltJCCategory"))) & vbCrLf
                s = s & "   ,OrderUOM=" & DbQuote(Str, .TextMatrix(i, .ColIndex("OrderUOM"))) & vbCrLf
                s = s & "   ,TakeoffUOM=" & DbQuote(Str, .TextMatrix(i, .ColIndex("TakeoffUOM"))) & vbCrLf
                s = s & "   ,OptionID=" & DbQuote(Str, .TextMatrix(i, .ColIndex("OptionID"))) & vbCrLf
                s = s & "   ,OptionCategory=" & DbQuote(Str, .TextMatrix(i, .ColIndex("OptionCategory"))) & vbCrLf
                s = s & "   ,Color=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Color"))) & vbCrLf
                s = s & "   ,Location=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Location"))) & vbCrLf
                s = s & "   ,RetailPretax=" & DbQuote(Num, .TextMatrix(i, .ColIndex("RetailPretax"))) & vbCrLf
                s = s & "   ,UStmp=" & DbQuote(Str, HFApp.LoginID) & vbCrLf
                s = s & "   ,TStmp=getdate()" & vbCrLf
                s = s & "   ,ConversionFactor=" & DbQuote(Num, Conversion) & vbCrLf
                s = s & "   ,WastePercent=" & DbQuote(Num, Abs(.ValueMatrix(i, .ColIndex("WastePercent")))) & vbCrLf
                s = s & "   ,RoundDir=" & DbQuote(Num, .ValueMatrix(i, .ColIndex("RoundDir"))) & vbCrLf
                s = s & "   ,RoundTo=" & DbQuote(Num, .ValueMatrix(i, .ColIndex("Roundto"))) & vbCrLf
                s = s & "   ,Price=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Price"))) & vbCrLf
                s = s & "   ,TaxGroup=" & DbQuote(Str, .TextMatrix(i, .ColIndex("TaxGroup"))) & vbCrLf
                s = s & "   ,IsQuote=" & DbQuote(Bit, .ValueMatrix(i, .ColIndex("IsQuote"))) & vbCrLf
                s = s & "   ,UseInFieldPO=" & DbQuote(Bit, .ValueMatrix(i, .ColIndex("UseInFieldPO"))) & vbCrLf
                For w = 1 To 40
                    s = s & "   ,WBS" & format(w, "00") & "=" & DbQuote(Str, .TextMatrix(i, .ColIndex("WBS" & format(w, "00")))) & vbCrLf
                Next
                s = s & "WHERE DivisionID = " & HFApp.DivisionID & " and Phase=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Phase"))) & vbCrLf
                s = s & "  AND Item=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Item"))) & vbCrLf
                Call HFApp.SqlExec(s, dbHomefront)
                If Val(.Cell(flexcpData, i, .ColIndex("PriceLink"))) <> 0 Then
                    s = ""
                    s = s & "update tblPhaseItem" & vbCrLf
                    s = s & "set Price=" & DbQuote(Num, .TextMatrix(i, .ColIndex("Price"))) & vbCrLf
                    s = s & "   ,PartNumber=" & DbQuote(Str, .TextMatrix(i, .ColIndex("PartNumber"))) & vbCrLf
                    s = s & "where DivisionID = " & HFApp.DivisionID & " and pricelink=" & DbQuote(Num, .Cell(flexcpData, i, .ColIndex("PriceLink"))) & vbCrLf
                    Call HFApp.SqlExec(s)
                End If
                
                
                Call UpdateItems(.TextMatrix(i, .ColIndex("Phase")), .TextMatrix(i, .ColIndex("Item")))
                
                
            End If
            
            .RowData(i) = ""
        
        Next
    End With
    
    SaveData = True
    mDirty = False
    Screen.MousePointer = vbDefault

Exit Function
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Screen.MousePointer = vbDefault
        MsgBox Trim(gItems.TextMatrix(i, gItems.ColIndex("Phase"))) & "/" & Trim(gItems.TextMatrix(i, gItems.ColIndex("Item"))) & " already exists. Please specify a different value.", vbInformation, App.ProductName
        Call SetCtrlFocus(gItems)
    Else
        Call errHandler(SRCFILE & "SaveData", s)
        Screen.MousePointer = vbDefault
    End If
End Function

Private Function PrettyName(fieldname As String) As String
    Dim s As String
    s = Trim(fieldname)
    Select Case s
        Case "CostCategory":                                      s = "Cost Type"
        Case "GrpPhase + ' - ' + GrpDesc":                        s = "Group"
        Case "POIndex + ' ' + POIndexDescription":                s = "PO Index"
        Case "Phase + ' - ' + PhaseDesc":                         s = "Phase"
        Case "Item + ' - ' + ItemDesc":                           s = "Item"
        Case "JCCostCode + ' - ' + JCCostCodeDesc":               s = "Cost Code"
        Case "ltrim(JCCostCode) + ' - ' + JCCostCodeDesc":        s = "Cost Code"
        Case "JCCategory + ' - ' + JCCategoryDesc":               s = "Category"
        Case "space(20-len(ltrim(item)))+ltrim(item) + ' - ' + ItemDesc":    s = "Item"
        Case "":                                                  s = ""
        Case "POIndex":
        Case Else
            MsgBox "STOP!!!" & vbCrLf & vbCrLf & "YOU NEED TO TRANSLATE THIS" & vbCrLf & vbCrLf & s, vbExclamation, App.ProductName
    End Select
    PrettyName = s
End Function

Private Sub gItems_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
Dim mr As Long
    With gItems
        Select Case True
        
            Case .MouseRow = 0 And Button = vbRightButton
                Cancel = True
                Call FMain.ShowColumnMenu(gItems)
                
            Case Button = vbRightButton
                Cancel = True
                
                FMain.mnuFItemsItemsSub(mcITEM_RENUMBER).Enabled = .Row > 1
                FMain.mnuFItemsItemsSub(mcITEM_NEWITEM).Enabled = True
                FMain.mnuFItemsItemsSub(mcITEM_DUPLICATE).Enabled = .Row > 1
                FMain.mnuFItemsItemsSub(mcITEM_DELETE).Enabled = .Row > 1
                FMain.mnuFItemsItemsSub(mcITEM_PRICEGROUPS).Enabled = .Row > 1
                
                PopupMenu FMain.mnuFItemsItems
                
        End Select
    End With
End Sub


Private Sub ImportItems(Source As String)
On Error GoTo eh

    Dim s As String
    Dim Row As Long
    Dim i As Long
    Dim rs As Recordset
    Dim Count As Long
    Dim Conversion As Double
    Dim categories As String
    Dim Category   As String
    
    Call FProgress.Progress("Synchronizing items...", "querying...", 1, 2)
    
    
    
    'mark items so we know which ones aren't touched
    Call HFApp.SqlExec("update tblPhaseItem set flag=1 where DivisionID = " & HFApp.DivisionID)

    Row = 0
    Select Case Source
    
    Case "Sage100"
        s = "SELECT COUNT(*) FROM tkfprt"
        Count = Val("" & HFApp.SqlExec(s, dbAccounting)(0))
        s = ""
        s = s & "select tkfprt.prtcls Phase" & vbCrLf
        s = s & "      ,tkfprt.recnum Item" & vbCrLf
        s = s & "      ,tkfprt.prtnme ItemDesc" & vbCrLf
        s = s & "      ,tkfprt.prtunt UOM" & vbCrLf
        s = s & "      ,tkfprt.prtcst Price" & vbCrLf
        s = s & "      ,tkfprt.cstcde JCCostCode" & vbCrLf
        s = s & "      ,tkfprt.csttyp JCCategory" & vbCrLf
        s = s & "      ,tkfprt.recnum PartNumber" & vbCrLf
        s = s & "  from prtcls,tkfprt" & vbCrLf
        s = s & "where prtcls.indent<>0 and prtcls.recnum=tkfprt.prtcls" & vbCrLf
        Set rs = HFApp.SqlExec(s, dbAccounting)
        While Not rs.EOF
            Row = Row + 1
            Call FProgress.Progress(, "updating database...", Row, Count)
            Call ImportItem(False, _
                            True, _
                            "" & rs("Phase"), _
                            "" & rs("Item"), _
                            "M", _
                            "" & rs("ItemDesc"), _
                            "", _
                            0, _
                            0, _
                            "", _
                            0, _
                            "", _
                            "" & rs("uom"), _
                            1, _
                            "" & rs("uom"), _
                            Val("" & rs("price")), _
                            "" & rs("jccostcode"), _
                            "" & rs("jccategory"), _
                            "" & rs("partnumber"))
            
            Call UpdateItems("" & rs("Phase"), "" & rs("Item"))
            
            rs.MoveNext
        Wend
    
    Case "SageEstimating"
        
        s = "SELECT COUNT(*) FROM Item"
        Count = Val("" & SageEstimatingDB.Execute(SageEstimatingDBName & s)(0))
        
        s = ""
        s = s & "SELECT ltrim(i.PhaseCode) Phase" & vbCrLf
        s = s & "      ,ltrim(i.ItemCode) Item" & vbCrLf
        s = s & "      ,i.Description ItemDesc" & vbCrLf
        s = s & "      ,i.TakeoffUnitName TakeoffUOM" & vbCrLf
        s = s & "      ,i.CategoryCodes Category_Codes" & vbCrLf
        s = s & "      ,i.BomClass POIndex" & vbCrLf
        s = s & "      ,case when i.LabConversionFactor=0 then 1 when i.LabMultDiv='D' then 1/i.LabConversionFactor else i.LabConversionFactor end LConversion" & vbCrLf
        s = s & "      ,i.LabOrderUnitName LOrderUOM" & vbCrLf
        s = s & "      ,i.LabPrice LPrice" & vbCrLf
        s = s & "      ,i.LabJcCostCode LJCCostCode" & vbCrLf
        s = s & "      ,i.LabJcCategory LJCCategory" & vbCrLf
        s = s & "      ,case when i.LabUseWaste=1 then i.WastePercent else 0 end LWastePercent" & vbCrLf
        s = s & "      ,i.LabRoundDirection LRoundDir" & vbCrLf
        s = s & "      ,i.LabRoundUnit LRoundTo" & vbCrLf
        s = s & "      ,i.LabPriceLink LPriceLink" & vbCrLf
        s = s & "      ,case when i.MatConversionFactor=0 then 1 when i.MatMultDiv='D' then 1/i.MatConversionFactor else i.MatConversionFactor end MConversion" & vbCrLf
        s = s & "      ,i.MatOrderUnitName MOrderUOM" & vbCrLf
        s = s & "      ,i.MatPrice MPrice" & vbCrLf
        s = s & "      ,i.MatJcCostCode MJCCostCode" & vbCrLf
        s = s & "      ,i.MatJcCategory MJCCategory" & vbCrLf
        s = s & "      ,case when i.MatUseWaste=1 then i.WastePercent else 0 end MWastePercent" & vbCrLf
        s = s & "      ,i.MatRoundDirection MRoundDir" & vbCrLf
        s = s & "      ,i.MatRoundUnit MRoundTo" & vbCrLf
        s = s & "      ,i.MatPriceLink MPriceLink" & vbCrLf
        s = s & "      ,case when i.SubConversionFactor=0 then 1 when i.SubMultDiv='D' then 1/i.SubConversionFactor else i.SubConversionFactor end SConversion" & vbCrLf
        s = s & "      ,i.SubOrderUnitName SOrderUOM" & vbCrLf
        s = s & "      ,i.SubPrice SPrice" & vbCrLf
        s = s & "      ,i.SubJcCostCode SJCCostCode" & vbCrLf
        s = s & "      ,i.SubJcCategory SJCCategory" & vbCrLf
        s = s & "      ,case when i.SubUseWaste=1 then i.WastePercent else 0 end SWastePercent" & vbCrLf
        s = s & "      ,i.SubRoundDirection SRoundDir" & vbCrLf
        s = s & "      ,i.SubRoundUnit SRoundTo" & vbCrLf
        s = s & "      ,0 SPriceLink" & vbCrLf
        s = s & "      ,case when i.EqpConversionFactor=0 then 1 when i.EqpMultDiv='D' then 1/i.EqpConversionFactor else i.EqpConversionFactor end EConversion" & vbCrLf
        s = s & "      ,i.EqpOrderUnitName EOrderUOM" & vbCrLf
        s = s & "      ,i.EqpPrice EPrice" & vbCrLf
        s = s & "      ,i.EqpJcCostCode EJCCostCode" & vbCrLf
        s = s & "      ,i.EqpJcCategory EJCCategory" & vbCrLf
        s = s & "      ,case when i.EqpUseWaste=1 then i.WastePercent else 0 end EWastePercent" & vbCrLf
        s = s & "      ,i.EqpRoundDirection ERoundDir" & vbCrLf
        s = s & "      ,i.EqpRoundUnit ERoundTo" & vbCrLf
        s = s & "      ,i.EqpPriceLink EPriceLinK" & vbCrLf
        s = s & "      ,case when i.OthConversionFactor=0 then 1 when i.OthMultDiv='D' then 1/i.OthConversionFactor else i.OthConversionFactor end OConversion" & vbCrLf
        s = s & "      ,i.OthOrderUnitName OOrderUOM" & vbCrLf
        s = s & "      ,i.OthPrice OPrice" & vbCrLf
        s = s & "      ,i.OthJcCostCode OJCCostCode" & vbCrLf
        s = s & "      ,i.OthJcCategory OJCCategory" & vbCrLf
        s = s & "      ,case when i.OthUseWaste=1 then i.WastePercent else 0 end OWastePercent" & vbCrLf
        s = s & "      ,i.OthRoundDirection ORoundDir" & vbCrLf
        s = s & "      ,i.OthRoundUnit ORoundTo" & vbCrLf
        s = s & "      ,0 OPriceLink" & vbCrLf
        s = s & "      ,i.Note Notes" & vbCrLf
        s = s & "FROM Item i" & vbCrLf
        Set rs = SageEstimatingDB.Execute(SageEstimatingDBName & s)
        
        While Not rs.EOF
            Row = Row + 1
            Call FProgress.Progress(, "updating database...", Row, Count)
            
            
            categories = "" & rs("category_codes")
            For i = 1 To 5
                Category = Mid(categories, i, 1)
                
'                If Category = "M" And "" & rs("POIndex") <> "" Then Stop
                
                Select Case Category
                    Case "M", "L", "S", "E", "O"
                        Call ImportItem(True, _
                                        False, _
                                        "" & rs("Phase"), _
                                        "" & rs("Item"), _
                                        Category, _
                                        "" & rs("ItemDesc"), _
                                        "" & rs("Notes"), _
                                        Val("" & rs(Category & "PriceLink")), _
                                        Val("" & rs(Category & "WastePercent")), _
                                        "" & rs(Category & "RoundDir"), _
                                        Val("" & rs(Category & "RoundTo")), _
                                        "" & rs("poindex"), _
                                        "" & rs("takeoffuom"), _
                                        Val("" & rs(Category & "conversion")), _
                                        "" & rs(Category & "orderuom"), _
                                        Val("" & rs(Category & "price")), _
                                        "" & rs(Category & "jccostcode"), _
                                        "" & rs(Category & "jccategory"), _
                                        "")
                End Select
            Next
            rs.MoveNext
        Wend
        
        
    End Select
    
    s = "update tblphaseitem set notes=dbo.rtf2text(notes)"
    HFApp.SqlExec s
    
    Unload FProgress
    
    
    
    'check if we need to remove any items
    Count = HFApp.SqlExec("select count(*) from tblphaseitem where flag=1 and DivisionID = " & HFApp.DivisionID)(0)
    If Count <> 0 Then
        If vbYes = MsgBox("Precision Builder's database has " & Count & " items that are not in your estimating database. Do you want to remove them?", vbQuestion + vbYesNo, App.ProductName) Then
            Call HFApp.SqlExec("delete from tblPhaseItem where flag=1 and DivisionID = " & HFApp.DivisionID)
        End If
    End If
    
    
Exit Sub
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Call errHandler(SRCFILE & "ImportItems", s)
        Unload FProgress
    End If
End Sub

Private Sub ImportItem(UpdateOtherInfo As Boolean, UpdatePartNumber As Boolean, Phase As String, ItemNumber As String, CostCategory As String, Description As String, Notes As String, PriceLink As Long, WastePercent As Long, RoundDir As String, RoundTo As Double, POIndex As String, TakeoffUOM As String, Conversion As Double, OrderUOM As String, price As Double, JCCostCode As String, JCCategory As String, PartNumber As String)
On Error GoTo eh
    Dim s As String
    
    If POIndex <> "" Then
        Call HFApp.SqlExec("INSERT INTO tblPOIndex(DivisionID,POIndex) VALUES(" & HFApp.DivisionID & "," & DbQuote(Str, POIndex) & ")")
    End If
    
    
    
    Conversion = Abs(Conversion)
    If Conversion = 0 Then Conversion = 1
    
    s = ""
    s = s & "INSERT INTO tblPhaseItem(DivisionID,Flag,PartNumber,Phase,Item,ItemNumber,CostCategory,Description,Notes,WastePercent,RoundDir,Roundto" & vbCrLf
    If CostCategory = "M" Then
        s = s & "                         ,PriceLink,POIndex" & vbCrLf
    End If
    s = s & "                         ,TakeoffUOM,ConversionFactor,OrderUOM,Price,JCCostCode,JCCategory)" & vbCrLf
    s = s & "VALUES(" & HFApp.DivisionID & ",0" & vbCrLf
    s = s & "      ," & DbQuote(Str, PartNumber, , True, 50) & vbCrLf
    s = s & "      ," & DbQuote(Str, Phase, , True, 20) & vbCrLf
    s = s & "      ," & DbQuote(Str, ItemNumber & CostCategory, , True, 16) & vbCrLf
    s = s & "      ," & DbQuote(Str, ItemNumber, , True, 15) & vbCrLf
    s = s & "      ," & DbQuote(Str, CostCategory, , True, 1) & vbCrLf
    s = s & "      ," & DbQuote(Str, Description, , True, 150) & vbCrLf
    s = s & "      ," & DbQuote(Str, Notes, , True, 4000) & vbCrLf
    s = s & "      ," & DbQuote(Num, WastePercent) & vbCrLf
    Select Case RoundDir
        Case "U":  s = s & "      ,1" & vbCrLf
        Case "D":  s = s & "      ,2" & vbCrLf
        Case "C":  s = s & "      ,3" & vbCrLf
        Case Else: s = s & "      ,0" & vbCrLf
    End Select
    s = s & "      ," & DbQuote(Num, RoundTo) & vbCrLf
    If CostCategory = "M" Then
        s = s & "      ," & DbQuote(Num, PriceLink) & vbCrLf
        s = s & "      ," & DbQuote(Str, POIndex, , True, 20) & vbCrLf
    End If
    s = s & "      ," & DbQuote(Str, TakeoffUOM, , True, 10) & vbCrLf
    s = s & "      ," & DbQuote(Num, Conversion) & vbCrLf
    s = s & "      ," & DbQuote(Str, OrderUOM, , True, 10) & vbCrLf
    s = s & "      ," & DbQuote(Num, price) & vbCrLf
    s = s & "      ," & DbQuote(Str, HFApp.FormatCostCode(JCCostCode), , , 15) & vbCrLf
    s = s & "      ," & DbQuote(Str, JCCategory, , , 3) & ")"
    Call HFApp.SqlExec(s)
    
    
    s = ""
    s = s & "UPDATE tblPhaseItem" & vbCrLf
    s = s & "SET Flag=0" & vbCrLf
    s = s & "   ,PhaseSortOrder=" & DbQuote(Num, Phase) & vbCrLf
    s = s & "   ,ItemSortOrder=" & DbQuote(Num, ItemNumber) & vbCrLf
    s = s & "   ,Description=" & DbQuote(Str, Description, , True, 150) & vbCrLf
    If UpdateOtherInfo Then
        s = s & "   ,Notes=" & DbQuote(Str, Notes, , True, 4000) & vbCrLf
        s = s & "   ,ConversionFactor=" & DbQuote(Num, Conversion) & vbCrLf
        If CostCategory = "M" Then
'            s = s & "   ,PriceLink=" & DbQuote(Num, PriceLink) & vbCrLf
            s = s & "   ,POIndex=" & DbQuote(Str, POIndex, , True, 20) & vbCrLf
        End If
        s = s & "   ,WastePercent=" & DbQuote(Num, WastePercent) & vbCrLf
        Select Case RoundDir
            Case "U":  s = s & "   ,RoundDir=1" & vbCrLf
            Case "D":  s = s & "   ,RoundDir=2" & vbCrLf
            Case "C":  s = s & "   ,RoundDir=3" & vbCrLf
            Case Else: s = s & "   ,RoundDir=0" & vbCrLf
        End Select
        s = s & "   ,RoundTo=" & DbQuote(Num, RoundTo) & vbCrLf
        s = s & "   ,OrderUOM=" & DbQuote(Str, OrderUOM, , True) & vbCrLf
'        s = s & "   ,Price=" & DbQuote(Num, Price) & vbCrLf
        s = s & "   ,TakeoffUOM=" & DbQuote(Str, TakeoffUOM, , True) & vbCrLf
    End If
    If UpdatePartNumber Then
        s = s & "   ,PartNumber=" & DbQuote(Str, PartNumber, , 50) & vbCrLf
    End If
    s = s & "WHERE DivisionID = " & HFApp.DivisionID & " and Phase=" & DbQuote(Str, Phase, , True) & vbCrLf
    s = s & "  AND Item=" & DbQuote(Str, ItemNumber & CostCategory, , True) & vbCrLf
    Call HFApp.SqlExec(s)
    

Exit Sub
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub
Private Sub ImportPhases(Source As String)
On Error GoTo eh

    Dim s As String
    Dim Row As Long
    Dim rs As Recordset
    Dim Count As Long
    Dim Conversion As Double

    
    Call FProgress.Progress("Synchronizing phases...", "querying...", 1, 2)
    
    
    'mark phases so we know which ones aren't touched
    Call HFApp.SqlExec("update tblEstPhases set flag=1 where DivisionID = " & HFApp.DivisionID)
    
    
    ' get phases
    Row = 0
    Select Case Source
    Case "Sage100"
        s = "select count(*) from prtcls"
        Count = Val("" & HFApp.SqlExec(s, dbAccounting)(0))
        
        s = ""
        s = s & "select recnum Phase" & vbCrLf
        s = s & "      ,clsnme PhaseDesc" & vbCrLf
        s = s & "      ,1 GroupPhase" & vbCrLf
        s = s & "  from prtcls" & vbCrLf
        s = s & " where indent=0" & vbCrLf
        s = s & "UNION ALL" & vbCrLf
        s = s & "select recnum Phase" & vbCrLf
        s = s & "      ,clsnme PhaseDesc" & vbCrLf
        s = s & "      ,0 GroupPhase" & vbCrLf
        s = s & "  from prtcls" & vbCrLf
        s = s & " where indent<>0" & vbCrLf
        Set rs = HFApp.SqlExec(s, dbAccounting)
    
    Case "SageEstimating"
            
    
        s = "SELECT COUNT(*) FROM Phase"
        Count = Val("" & SageEstimatingDB.Execute(SageEstimatingDBName & s)(0))
        
        s = "select ltrim(PhaseCode) Phase,Description PhaseDesc,IsGroup GroupPhase from Phase"
        Set rs = SageEstimatingDB.Execute(SageEstimatingDBName & s)
    
    End Select
    
    While Not rs.EOF
        
        Call FProgress.Progress(, "updating database...", Row, Count)
        
        s = ""
        s = s & "INSERT INTO tblEstPhases(flag,DivisionID,Phase,Description,GroupPhase,SortOrder)" & vbCrLf
        s = s & "VALUES(0," & HFApp.DivisionID & "," & DbQuote(Str, "" & rs("Phase")) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("PhaseDesc")) & vbCrLf
        s = s & "      ," & DbQuote(Bit, "" & rs("GroupPhase")) & vbCrLf
        s = s & "      ," & DbQuote(Num, "" & rs("Phase")) & ")"
        Call HFApp.SqlExec(s)
        
        s = ""
        s = s & "UPDATE tblEstPhases" & vbCrLf
        s = s & "SET flag=0,Description=" & DbQuote(Str, "" & rs("PhaseDesc")) & vbCrLf
        s = s & "   ,GroupPhase=" & DbQuote(Bit, "" & rs("GroupPhase")) & vbCrLf
        s = s & "   ,SortOrder=" & DbQuote(Num, "" & rs("Phase")) & vbCrLf
        s = s & "WHERE DivisionID = " & HFApp.DivisionID & " and Phase=" & DbQuote(Str, "" & rs("Phase")) & vbCrLf
        Call HFApp.SqlExec(s)
        
        Row = Row + 1
        rs.MoveNext
    Wend
    
    'check if we need to remove any phases
    Count = HFApp.SqlExec("select count(*) from tblestphases where flag=1 and DivisionID = " & HFApp.DivisionID)(0)
    If Count <> 0 Then
        If vbYes = MsgBox("Precision Builder's database has " & Count & " phases that are not in your estimating database. Do you want to remove them?", vbQuestion + vbYesNo, App.ProductName) Then
            Call HFApp.SqlExec("delete from tblestphases where flag=1 and DivisionID = " & HFApp.DivisionID)
        End If
    End If
    
    
    Call FixGroupPhaseValue
    Unload FProgress

Exit Sub
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Call errHandler(SRCFILE & "ImportPhases", s)
        Unload FProgress
    End If
End Sub

Public Sub mnuFItemsPriceGroupSub_Click(Index As Integer)
    Dim i As Long
    Dim HF As String
    
    Dim TL_M As String
    Dim TL_L As String
    Dim TL_E As String
    
    Dim PriceGroup As Long
    Dim PriceGroupDesc As String
    Dim rs As Recordset
    
    If Not SaveData(True) Then Exit Sub
    
    With gItems
        HF = ""
        TL_M = ""
        TL_L = ""
        TL_E = ""
        For i = Min(.Row, .RowSel) To Max(.Row, .RowSel)
            HF = HF & " or (Phase=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Phase"))) & " and Item=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Item"))) & ")"
            Select Case .TextMatrix(i, .ColIndex("CostCategory"))
                Case "L":  TL_L = TL_L & " or (Phase_Code=" & DbQuote(Str, HFApp.FormatPhase(.TextMatrix(i, .ColIndex("Phase")))) & " and Item_Number=" & DbQuote(Str, HFApp.FormatItem(.TextMatrix(i, .ColIndex("ItemNumber")))) & ")"
                Case "E":  TL_E = TL_E & " or (Phase_Code=" & DbQuote(Str, HFApp.FormatPhase(.TextMatrix(i, .ColIndex("Phase")))) & " and Item_Number=" & DbQuote(Str, HFApp.FormatItem(.TextMatrix(i, .ColIndex("ItemNumber")))) & ")"
                Case Else: TL_M = TL_M & " or (Phase_Code=" & DbQuote(Str, HFApp.FormatPhase(.TextMatrix(i, .ColIndex("Phase")))) & " and Item_Number=" & DbQuote(Str, HFApp.FormatItem(.TextMatrix(i, .ColIndex("ItemNumber")))) & ")"
            End Select
        Next
        HF = "(" & Mid(HF, 5) & ")"
        TL_M = "(" & Mid(TL_M, 5) & ")"
        TL_L = "(" & Mid(TL_L, 5) & ")"
        TL_E = "(" & Mid(TL_E, 5) & ")"
        
        If TL_M = "()" Then TL_M = "(1=2)"
        If TL_L = "()" Then TL_L = "(1=2)"
        If TL_E = "()" Then TL_E = "(1=2)"
        
    End With
    
    
    Select Case Index
        Case 0 'make group
            PriceGroup = HFApp.SqlExec("SELECT ISNULL(MAX(PriceLink),0)+1 FROM tblPhaseItem where DivisionID = " & HFApp.DivisionID, dbHomefront)(0)
            Call HFApp.SqlExec("update tblPhaseItem SET PriceLink=" & DbQuote(Num, PriceGroup) & " WHERE " & HF & " and DivisionID = " & HFApp.DivisionID)
            
            
            
            With gItems
                Set rs = HFApp.SqlExec("select PriceGroup,PriceGroupDesc from tblphaseitem left outer join PriceGroups on(PriceLink=PriceGroup) where tblPhaseItem.DivisionID = " & HFApp.DivisionID & " and Phase=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("Phase"))) & " and Item=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("Item"))))
                PriceGroup = rs("PriceGroup")
                PriceGroupDesc = rs("PriceGroupDesc")
                For i = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                    .Cell(flexcpData, i, .ColIndex("PriceLink")) = PriceGroup
                    .Cell(flexcpText, i, .ColIndex("PriceLink")) = PriceGroupDesc
                    If Val(.Cell(flexcpData, i, .ColIndex("PriceLink"))) <> 0 Then
                        .Cell(flexcpPicture, i, .ColIndex("PriceLink")) = FMain.SmallIcons.ListImages.Item("links").Picture
                    End If
                Next
            End With
        
        Case 1 'join group
            If FPickList.Choose(HFApp.Databases(dbHomefront), "Pricing Group", "select PriceGroup, PriceGroupDesc from PriceGroups where divisionid=" & DbQuote(Num, HFApp.DivisionID), , , , , "PriceGroup") Then
                PriceGroup = FPickList.SelectedItem("PriceGroup")
                PriceGroupDesc = FPickList.SelectedItem("PriceGroupDesc")
                Call HFApp.SqlExec("update tblPhaseItem SET PriceLink=" & DbQuote(Num, PriceGroup) & " WHERE " & HF & " and DivisionID = " & HFApp.DivisionID)
                With gItems
                    For i = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                        .Cell(flexcpData, i, .ColIndex("PriceLink")) = PriceGroup
                        .Cell(flexcpText, i, .ColIndex("PriceLink")) = PriceGroupDesc
                        If Val(.Cell(flexcpData, i, .ColIndex("PriceLink"))) <> 0 Then
                            .Cell(flexcpPicture, i, .ColIndex("PriceLink")) = FMain.SmallIcons.ListImages.Item("links").Picture
                        End If
                    Next
                End With
            End If
        
        Case 2 'leave group
            Call HFApp.SqlExec("update tblPhaseItem SET PriceLink=0 WHERE " & HF & " and DivisionID = " & HFApp.DivisionID)
                        
            
            With gItems
                For i = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                    .Cell(flexcpData, i, .ColIndex("PriceLink")) = 0
                    .Cell(flexcpText, i, .ColIndex("PriceLink")) = ""
                    .Cell(flexcpPicture, i, .ColIndex("PriceLink")) = Nothing
                Next
            End With
        
    End Select
    
End Sub

Private Sub UpdateItems(Phase As String, Item As String)
    Dim s As String

    s = ""
    s = s & "update custompreestimateitems" & vbCrLf
    s = s & "set takeoffuom=i.takeoffuom" & vbCrLf
    s = s & "   ,orderuom=i.orderuom" & vbCrLf
    s = s & "   ,conversionfactor=i.conversionfactor" & vbCrLf
    s = s & "from custompreestimateitems e join tblphaseitem i on(i.Divisionid = " & HFApp.DivisionID & " and e.phase=i.phase and e.item=i.item)" & vbCrLf
    s = s & "WHERE e.estitemid=0" & vbCrLf
    s = s & "  and i.phase=" & DbQuote(Str, Phase) & vbCrLf
    s = s & "  AND i.item=" & DbQuote(Str, Item) & vbCrLf
    Call HFApp.SqlExec(s)
    
    s = ""
    s = s & "update estimateitems" & vbCrLf
    s = s & "set takeoffuom=i.takeoffuom" & vbCrLf
    s = s & "   ,orderuom=i.orderuom" & vbCrLf
    s = s & "   ,conversionfactor=i.conversionfactor" & vbCrLf
    s = s & "from estimateitems e join tblphaseitem i on(e.DivisionID = i.DivisionID and e.phase=i.phase and e.item=i.item)" & vbCrLf
    s = s & "where e.budgetgenerated=0 and e.pogenbatch=0" & vbCrLf
    s = s & "  and i.phase=" & DbQuote(Str, Phase) & vbCrLf
    s = s & "  AND i.item=" & DbQuote(Str, Item) & vbCrLf
    s = s & "  AND i.DivisionID = " & HFApp.DivisionID
    
    Call HFApp.SqlExec(s)


End Sub




















Public Sub mnuFItemsPhasesSub_Click(Index As Integer)
On Error GoTo eh
    Dim Name As String
    Dim oldcode As String
    Dim newcode As String
    
    Dim HFcsv As String
    
    Dim s As String
    Dim rs As Recordset
    Dim i As Long
    
    If Not SaveData(False) Then Exit Sub
    
    Select Case Index
    
        Case mcPHASE_RENAME
            With gPhases
                oldcode = .Cell(flexcpText, .Row, 1)
                Name = Mid(.Cell(flexcpText, .Row, 0), Len(oldcode) + 4)
                Name = Left(InputBox(vbCrLf & vbCrLf & "Enter a new description for " & oldcode, App.ProductName, Name), HFApp.EstFieldSize("PhaseDesc"))
                If Name = "" Then Exit Sub
                .Cell(flexcpText, .Row, 0) = oldcode & " - " & Name
                
                Call HFApp.SqlExec("UPDATE tblEstPhases SET Description=" & DbQuote(Str, Name) & "WHERE DivisionID = " & HFApp.DivisionID & " and Phase=" & DbQuote(Str, oldcode))
                
            End With
        
        Case mcPHASE_RENUMBER
            With gPhases
                oldcode = .Cell(flexcpText, .Row, 1)
                Name = Mid(.Cell(flexcpText, .Row, 0), Len(oldcode) + 4)
                newcode = Left(InputBox(vbCrLf & vbCrLf & "Enter a new number for " & Name, App.ProductName, oldcode), HFApp.EstFieldSize("Phase"))
                If newcode = "" Then Exit Sub
                
                Screen.MousePointer = vbHourglass
                Call HFApp.SqlExec("UPDATE tblEstPhases           SET Phase=" & DbQuote(Str, newcode) & " WHERE DivisionID = " & HFApp.DivisionID & " and Phase=" & DbQuote(Str, oldcode))
                Call HFApp.SqlExec("UPDATE tblPhaseItem           SET Phase=" & DbQuote(Str, newcode) & " WHERE DivisionID = " & HFApp.DivisionID & " and Phase=" & DbQuote(Str, oldcode))
                Call HFApp.SqlExec("UPDATE tblDBAssemblyDetails   SET Phase=" & DbQuote(Str, newcode) & " WHERE DivisionID = " & HFApp.DivisionID & " and Phase=" & DbQuote(Str, oldcode))
                Call HFApp.SqlExec("UPDATE tblSalesSheetCosts     SET Phase=" & DbQuote(Str, newcode) & " WHERE Phase=" & DbQuote(Str, oldcode))
                Call HFApp.SqlExec("UPDATE tblVendorCost          SET Phase=" & DbQuote(Str, newcode) & " WHERE DivisionID = " & HFApp.DivisionID & " and Phase=" & DbQuote(Str, oldcode))
                Call FixGroupPhaseValue
                
                If gPhases.RowOutlineLevel(gPhases.Row) = 0 Then
                    gPhases.RowData(gPhases.Row) = " AND GrpPhase = " & DbQuote(Str, newcode)
                Else
                    gPhases.RowData(gPhases.Row) = " AND Phase = " & DbQuote(Str, newcode)
                    Call LoadItems
                End If
                .Cell(flexcpText, .Row, 1) = newcode
                .Cell(flexcpText, .Row, 0) = newcode & " - " & Name
                
                
               
            End With
        
        
        Case mcPHASE_NEWGROUP
            If FPhase.ShowForm(True) Then Call mnuPriceListViewsSub_Click(CInt(mViewIndex))
        
        Case mcPHASE_NEWPHASE
            If FPhase.ShowForm(False) Then Call mnuPriceListViewsSub_Click(CInt(mViewIndex))
        
        Case mcPHASE_DUPLICATE
            Call FDuplicateItems.ShowForm("Phase=" & DbQuote(Str, Trim(gPhases.Cell(flexcpText, gPhases.Row, 1))))
            
            
        Case mcPHASE_DELETE
            oldcode = gPhases.Cell(flexcpText, gPhases.Row, 1)
            
            If gPhases.RowOutlineLevel(gPhases.Row) = 0 Then
                i = MsgBox("Do you want to delete the underlying phases and items?", vbYesNoCancel + vbExclamation, "Confirm Delete")
            Else
                i = MsgBox("This will delete the underlying items." & vbCrLf & vbCrLf & "Are you sure you want to continue?", vbOKCancel + vbExclamation, "Confirm Delete")
            End If
            
            'double check deleting
            If i = vbCancel Then Exit Sub
            If i = vbYes Then
                If Not FConfirmationMsgBox.ShowWarning("This action could affect your model and option library. If you click" & vbCrLf & _
                                                       "OK, these phases and items will be permanently deleted and" & vbCrLf & _
                                                       "removed from your database and all your model and option" & vbCrLf & _
                                                       "assemblies." & vbCrLf & vbCrLf & _
                                                       "Are you sure this is what you want to do?") Then Exit Sub
                
            End If
            
            
            Screen.MousePointer = vbHourglass
            'build csv list of phases
            HFcsv = DbQuote(Str, oldcode)
            If gPhases.RowOutlineLevel(gPhases.Row) = 0 And i = vbYes Then
                Set rs = HFApp.SqlExec("select phase from tblestphases where DivisionID = " & HFApp.DivisionID & " and groupphasevalue=" & DbQuote(Str, oldcode))
                While Not rs.EOF
                    HFcsv = HFcsv & "," & DbQuote(Str, "" & rs(0))
                    rs.MoveNext
                Wend
            End If
            
            'delete phases and items in the list
            Call HFApp.SqlExec("DELETE FROM tblEstPhases WHERE DivisionID = " & HFApp.DivisionID & " and Phase in(" & HFcsv & ")")
            Call FixGroupPhaseValue
            If i = vbYes Then
                Call HFApp.SqlExec("DELETE FROM tblPhaseItem           WHERE DivisionID = " & HFApp.DivisionID & " and phase in(" & HFcsv & ")")
                Call HFApp.SqlExec("DELETE FROM tblDBAssemblyDetails   WHERE DivisionID = " & HFApp.DivisionID & " and phase in(" & HFcsv & ")")
                Call HFApp.SqlExec("DELETE FROM tblSalesSheetCosts     WHERE phase in(" & HFcsv & ")")
                Call HFApp.SqlExec("DELETE FROM tblVendorCost          WHERE DivisionID = " & HFApp.DivisionID & " and phase in(" & HFcsv & ")")
            End If
            
            
            Call FixGroupPhaseValue
            If gPhases.RowOutlineLevel(gPhases.Row) = 0 Then
                Call gPhases.RemoveItem
                Call mnuPriceListViewsSub_Click(CInt(mViewIndex))
            Else
                Call gPhases.RemoveItem
            End If
            gItems.Rows = 1
        
    End Select
    
    
EXITSUB:
    Screen.MousePointer = vbDefault
    Exit Sub
    
eh:
    If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Screen.MousePointer = vbDefault
        MsgBox Trim(newcode) & " already exists. Please specify a different value.", vbInformation, App.ProductName
        Call SetCtrlFocus(gItems)
    Else
        Call errHandler(SRCFILE & "mnuFItemsPhasesSub_Click", s)
        Screen.MousePointer = vbDefault
    End If
End Sub


Public Sub mnuFItemsItemsSub_Click(Index As Integer)
On Error GoTo eh

    Dim i As Long
    Dim r As Long
    Dim s As String
    Dim Phase As String
    Dim Item As String
    Dim Name As String
    Dim rs As Recordset
    Dim ObjectID As String
    
    With gItems
    Select Case Index
    
        Case mcITEM_VIEWFILES
            mTimerTask = "EditAttachments|" & "ITM~" & .TextMatrix(.Row, .ColIndex("Phase")) & "~" & .TextMatrix(.Row, .ColIndex("Item")) & "|Item Database"
            Timer1.Enabled = True
            Timer1.Interval = 10
            
        
            
        Case mcITEM_RENUMBER
            Call SaveData(False)
            If .Row > 0 Then
                If FItem.Renumber(.TextMatrix(.Row, .ColIndex("Phase")), .TextMatrix(.Row, .ColIndex("Item"))) Then Call LoadItems
            End If


        Case mcITEM_NEWITEM
            Call SaveData(False)
            Phase = ""
            If .Row > 0 Then Phase = .TextMatrix(.Row, .ColIndex("Phase"))
            If FItem.Add(Phase, .TextMatrix(.Row, .ColIndex("Item")), .TextMatrix(.Row, .ColIndex("ItemDesc"))) Then Call LoadItems


        Case mcITEM_DUPLICATE
            If .Row < 1 Then Exit Sub
            Call SaveData(False)
            s = ""
            For r = Max(.Row, .RowSel) To Min(.Row, .RowSel) Step -1
            If Not .RowHidden(r) Then
                s = s & " or (phase=" & DbQuote(Str, .TextMatrix(r, .ColIndex("phase"))) & " and item=" & DbQuote(Str, .TextMatrix(r, .ColIndex("item"))) & ")"
            End If
            Next
            s = "(" & Mid(s, 5) & ")"
            If FDuplicateItems.ShowForm(s) Then Call LoadItems
            
            
        Case mcITEM_DELETE
            If .Row < 1 Then Exit Sub
            
            For r = Max(.Row, .RowSel) To Min(.Row, .RowSel) Step -1
            If Not .RowHidden(r) Then
                .RowHidden(r) = True
                .RowData(r) = "Deleted"
                mDirty = True
            End If
            Next

            For r = Min(.Row, .RowSel) To 1 Step -1
                If Not .RowHidden(r) Then
                    .Row = r
                    Exit Sub
                End If
            Next
            For r = Max(.Row, .RowSel) To .Rows - 1
                If Not .RowHidden(r) Then
                    .Row = r
                    Exit Sub
                End If
            Next

    End Select
    End With
Exit Sub
eh: Call errHandler(SRCFILE & "mnuFItemsItemsSub_Click")
End Sub


Private Sub UpdateInverseItem(Phase As String, Item As String, NewValue As String)
    Dim r As Long
    With gItems
    
    For r = 1 To .Rows - 1
        If .TextMatrix(r, .ColIndex("phase")) = Phase And .TextMatrix(r, .ColIndex("item")) = Item And .TextMatrix(r, .ColIndex("InverseItemDesc")) = "" Then
            .TextMatrix(r, .ColIndex("InverseItemDesc")) = NewValue
            Exit Sub
        End If
    Next
    End With
End Sub


Private Sub LoadWBSDescriptions()
    Dim i As Long
    Dim c As Long
    Dim s As String
    Dim rs As Recordset
    
    With gItems
    
        'add columns
        c = .Cols - 1
        .Cols = .Cols + 40
        For i = 1 To 40
            .ColKey(c + i) = "WBS" & format(i, "00")
            .ColHidden(c + i) = True
        Next
        
        'read column names
        s = "select Item,custom_description Description from customDescriptions where isnull(custom_description,'')<>'' and item like 'WBS__' order by item"
        Set rs = HFApp.SqlExec(s, dbHomefront)
        While Not rs.EOF
            c = .ColIndex("" & rs(0))
            If c <> -1 Then
                .TextMatrix(0, c) = "" & rs(1)
                .ColHidden(c) = False
            End If
            rs.MoveNext
        Wend
    
    
    End With

End Sub



Private Sub imgShortcut_Click(Index As Integer)
    Select Case Index
    
        Case 0
            Call GridExpandALL(Me.gPhases)
    
    End Select

End Sub
