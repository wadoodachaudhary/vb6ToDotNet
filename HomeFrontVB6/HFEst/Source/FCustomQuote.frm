VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FCustomQuote 
   Caption         =   "Custom Option Quote"
   ClientHeight    =   5115
   ClientLeft      =   5280
   ClientTop       =   2505
   ClientWidth     =   13185
   Icon            =   "FCustomQuote.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   13185
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   13185
      _ExtentX        =   23257
      _ExtentY        =   1058
      ButtonWidth     =   1984
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "Save"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Preview"
            Key             =   "Preview"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "One Time"
            Key             =   "TakeoffOneTime"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Item"
            Key             =   "TakeoffItem"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Assembly"
            Key             =   "TakeoffAssembly"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Custom"
            Key             =   "TakeoffCustom"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Attachments"
            Key             =   "Attachments"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   2895
      TabIndex        =   8
      Top             =   3255
      Width           =   7515
      Begin HFEst.VBCombo cboUOM 
         Height          =   240
         Left            =   3015
         TabIndex        =   19
         Top             =   1470
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   423
         Style           =   2
         Text            =   "Combo1"
      End
      Begin VB.TextBox txtSalesQty 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   2580
         TabIndex        =   5
         Text            =   "3"
         Top             =   1470
         Width           =   420
      End
      Begin VB.TextBox txtMargin 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   5580
         TabIndex        =   3
         Text            =   "0"
         Top             =   660
         Width           =   1200
      End
      Begin VB.TextBox txtSell 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   5580
         TabIndex        =   4
         Text            =   "0.00"
         Top             =   900
         Width           =   1200
      End
      Begin VB.TextBox txtMarkup 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   5580
         TabIndex        =   2
         Text            =   "0"
         Top             =   420
         Width           =   1200
      End
      Begin VB.TextBox txtCost 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   5580
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "0.00"
         Top             =   180
         Width           =   1200
      End
      Begin VB.Label lblPrice 
         AutoSize        =   -1  'True
         Caption         =   "@ 11.11 = $22.22"
         Height          =   195
         Left            =   3735
         TabIndex        =   18
         Top             =   1485
         Width           =   1290
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sell "
         Height          =   195
         Index           =   2
         Left            =   2190
         TabIndex        =   17
         Top             =   1485
         Width           =   300
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Custom Request"
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
         Left            =   90
         TabIndex        =   16
         Top             =   225
         Width           =   1395
      End
      Begin VB.Label lblOptDesc 
         Caption         =   "Add 3 GFI plugs in laundry room nearest the door beside the sink in the back counter."
         Height          =   930
         Left            =   90
         TabIndex        =   15
         Top             =   465
         Width           =   4515
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Margin"
         Height          =   195
         Index           =   2
         Left            =   5040
         TabIndex        =   14
         Top             =   690
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sell"
         Height          =   195
         Index           =   3
         Left            =   5265
         TabIndex        =   11
         Top             =   930
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Markup"
         Height          =   195
         Index           =   1
         Left            =   4980
         TabIndex        =   10
         Top             =   450
         Width           =   540
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cost"
         Height          =   195
         Index           =   0
         Left            =   5205
         TabIndex        =   9
         Top             =   210
         Width           =   315
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid gInfoPanel 
      Height          =   1215
      Left            =   120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3255
      Width           =   5775
      _cx             =   2000758730
      _cy             =   2000750687
      Appearance      =   0
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
      BackColor       =   -2147483633
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483633
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FCustomQuote.frx":06EA
      ScrollTrack     =   0   'False
      ScrollBars      =   0
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
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
   End
   Begin VSFlex8Ctl.VSFlexGrid gItems 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   7995
      _cx             =   2000762646
      _cy             =   2000753227
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
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   41
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FCustomQuote.frx":0759
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
      TabBehavior     =   1
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
      Begin VB.PictureBox picWarningMessages 
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   2040
         ScaleHeight     =   615
         ScaleWidth      =   3795
         TabIndex        =   6
         Top             =   960
         Visible         =   0   'False
         Width           =   3795
         Begin VB.TextBox lblWarningMessages 
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   360
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   7
            Text            =   "FCustomQuote.frx":0D51
            Top             =   60
            Width           =   1095
         End
         Begin VB.Shape shpWarningMessages 
            BorderColor     =   &H80000017&
            Height          =   555
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Width           =   3675
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   60
            Picture         =   "FCustomQuote.frx":0D57
            Top             =   60
            Width           =   240
         End
      End
      Begin VB.Image imgWarning 
         Height          =   240
         Left            =   0
         Picture         =   "FCustomQuote.frx":12E1
         Top             =   240
         Visible         =   0   'False
         Width           =   240
      End
   End
End
Attribute VB_Name = "FCustomQuote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const SRCFILE = "FCustomQuote::"

Private mReadOnly As Boolean
Private mSaved As Boolean
Private mDirty  As Boolean

Private mFunnyFlag As Boolean  ' used to ignore the first mouse move event see gitems_mousemove and gitems_rowcolchanged
Private mEditFlag As Boolean   ' used to ignore margin/markup change events caused by each other

Private mAttachmentID   As String
Private mCommunity      As String
Private mCommunityPhase As String
Private mAssemblyDesc   As String
Private mSalesQty       As Double
Private mSalesUOM       As String
Private mSeq            As Long
Private mJob            As String
Private mExtra          As String
Private mCustomer       As String
Private mLocation       As Long
Private mModel          As String
Private mOptionID       As String

Private mCost           As Double
Private mSell           As Double
Private Const mcITEM_COPYITEMS = 0
Private Const mcITEM_SPLITITEMS = 1
Private Const mcITEM_SUBSTITUEITEM = 2
Private Const mcITEM_REMOVEITEMS = 3
Private Const mcITEM_VIEWFILES = 5
Private Const mcITEM_CANCELBUDGETS = 8
Private Const mcITEM_SAVEONETIMETODB = 10
Private Const mcITEM_UPDATEPRICELIST = 11
Private Const mcITEM_COMPAREPRICES = 12
Private Const mcITEM_FORMATTING = 13



Public Function BuildQuote(ReadOnly As Boolean, AttachmentID As String, Location As String, Model As String, Opt As String, Job As String, Customer As String, community As String, CommunityPhase As String, Extra As String, Description As String, seq As Long, Category As String, _
                           CostAmount As Double, SellAmount As Double, SalesQty As Double, SalesUOM As String) As Boolean
On Error GoTo eh
   
    mReadOnly = ReadOnly
    mAttachmentID = AttachmentID
    mCommunity = community
    mCommunityPhase = CommunityPhase
    mAssemblyDesc = Description
    mSalesQty = SalesQty
    mSalesUOM = SalesUOM
    
    mSeq = seq
    mJob = Replace(Job, "-", "")
    mExtra = Extra
    mCustomer = Customer
    
    Select Case Location
        Case "ADDENDUM":     mLocation = olScheduleB
        Case "CHANGEORDER":  mLocation = olChangeOrder
        Case "DESIGNCENTER": mLocation = olDesignCenter
    End Select
    mModel = Model
    mOptionID = Opt
    mCost = CostAmount
    mSell = SellAmount
    
    Load Me
    
    On Error Resume Next
    
    Call SetListIndex(cboUOM, , mSalesUOM)
    
    mEditFlag = True
    txtMarkup.Text = ""
    txtMarkup.Text = Val(HFApp.SqlExec("select comarkup from tblcategories where category=" & DbQuote(Str, Category))(0))
    If txtMarkup.Text = "0" Then txtMarkup.Text = ""
    
    
    CostAmount = mCost
    If CostAmount <> 0 And SellAmount <> 0 Then
        txtMarkup.Text = Round((SellAmount - CostAmount) / CostAmount * 100, 2)
        txtMargin.Text = Round((SellAmount - CostAmount) / SellAmount * 100, 0)
    End If
    
    On Error GoTo eh
    txtCost.Text = format(CostAmount, "#,##0.00")
    txtSell.Text = format(SellAmount, "#,##0.00")
    mEditFlag = False
    
    mSaved = False
    mDirty = False
    Me.Show vbModal
    BuildQuote = mSaved
    CostAmount = mCost
    SellAmount = mSell
    SalesQty = mSalesQty
    SalesUOM = mSalesUOM
    
Exit Function
eh: Call errHandler(SRCFILE & "BuildQuote")
End Function

Private Sub cboUOM_Click()
    mSalesUOM = cboUOM.Text
End Sub

Private Sub Form_Load()
On Error Resume Next
    Call SetToolbarIcons(Toolbar, FMain.LargeIcons)
    
    Call IniGetForm(Me)
    
    Call LoadComboBox(cboUOM, HFApp.Databases(dbHomefront), "SELECT DISTINCT OrderUOM Unit,'',0 FROM tblPhaseItem where DivisionID = " & HFApp.DivisionID & " order by 1")

    Call LoadWBSDescriptions
    Call IniGetGrid(Me, gItems)
    Call LoadData
    Call LoadInfoPanel

    ReadOnly = mReadOnly

End Sub


Private Property Let ReadOnly(RHS As Boolean)
    Toolbar.Buttons("Save").Enabled = Not RHS
    Toolbar.Buttons("TakeoffOneTime").Enabled = Not RHS
    Toolbar.Buttons("TakeoffItem").Enabled = Not RHS
    Toolbar.Buttons("TakeoffAssembly").Enabled = Not RHS
    Toolbar.Buttons("TakeoffCustom").Enabled = Not RHS
    Me.gItems.Editable = IIf(RHS, flexEDNone, flexEDKbdMouse)
    txtCost.Locked = RHS
    txtMarkup.Locked = RHS
    txtMargin.Locked = RHS
    txtSell.Locked = RHS
End Property

Private Sub Form_Resize()
On Error Resume Next
    gItems.Move 0, Toolbar.Height - Screen.TwipsPerPixelY, Me.ScaleWidth, Me.ScaleHeight - Frame1.Height - Toolbar.Height
    Frame1.Move Me.ScaleWidth - Frame1.Width, gItems.Top + gItems.Height
    gInfoPanel.Move 120, Frame1.Top + 120, Me.ScaleWidth - 120 - Frame1.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo eh
    If Not SaveData(True) Then
        Cancel = True
        Exit Sub
    End If
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gItems)
Exit Sub
eh:
If Err.Number <> 0 Then
    MsgBox "FCustomQuote: Unload " & vbCrLf & Err.Description, vbExclamation, App.ProductName
End If
End Sub

Public Sub AddItem(Assembly As String, _
                   AssemblyDescription As String, _
                   Model As String, _
                   Phase As String, _
                   Item As String, Sequence As Long, _
                   Description As String, _
                   OrderQty As Double, _
                   OrderUOM As String, _
                   TakeoffQty As Double, _
                   TakeoffUOM As String, _
                   ConversionFactor As Double, RoundTo As Double, RoundDir As Long, WastePercent As Long, _
                   JCExtra As String, _
                   JCCostCode As String, _
                   JCCostCodeDesc As String, JCCategory As String, _
                   JCCategoryDesc As String, Vendor As String, _
                   VendorName As String, price As Double, _
                   TaxGroup As String, TaxGroupName As String, _
                   JCTaxRate As Double, NJCTaxRate As Double, _
                   POIndex As String, _
                   Comments As String, Formula As String, SalesQty As Double, Location As String, _
                   WBS, Optional Job As String)

'new columns have been added but are not yet implemented in this forms code.
'roundto,rounddir,wastepercent

On Error GoTo eh
    
    Dim i As Long
    Dim w As Long
    Dim s As String
    
    mDirty = True
    With gItems
        .AddItem ""
        i = .Rows - 1
        
        .RowData(i) = "NEW"
        .Cell(flexcpChecked, i, .ColIndex("Selected")) = flexNoCheckbox
        .TextMatrix(i, .ColIndex("Customer_No")) = mCustomer
        .TextMatrix(i, .ColIndex("seq")) = mSeq
        .TextMatrix(i, .ColIndex("Location")) = mLocation
        .TextMatrix(i, .ColIndex("Assembly")) = Assembly
        .TextMatrix(i, .ColIndex("HFDescription")) = mAssemblyDesc
        .TextMatrix(i, .ColIndex("Job_No")) = mJob
        .TextMatrix(i, .ColIndex("JCExtra")) = mExtra
        .TextMatrix(i, .ColIndex("EstPhase")) = Phase
        .TextMatrix(i, .ColIndex("EstItem")) = Item
        .TextMatrix(i, .ColIndex("ItemDesc")) = Description
        .TextMatrix(i, .ColIndex("ConversionFactor")) = ConversionFactor
        .TextMatrix(i, .ColIndex("JCCostCode")) = JCCostCode
        .TextMatrix(i, .ColIndex("JCCostCodeDesc")) = JCCostCodeDesc
        .TextMatrix(i, .ColIndex("JCCategory")) = JCCategory
        .TextMatrix(i, .ColIndex("JCCategoryDesc")) = JCCategoryDesc
        .TextMatrix(i, .ColIndex("POIndex")) = POIndex
        .TextMatrix(i, .ColIndex("RoundTo")) = RoundTo
        .TextMatrix(i, .ColIndex("RoundDir")) = RoundDir
        .TextMatrix(i, .ColIndex("ItemComments")) = Comments
        .TextMatrix(i, .ColIndex("TakeoffQty")) = TakeoffQty
        .TextMatrix(i, .ColIndex("TakeoffUOM")) = TakeoffUOM
        .TextMatrix(i, .ColIndex("OrderUOM")) = OrderUOM
        .TextMatrix(i, .ColIndex("OptionID")) = mOptionID
        .TextMatrix(i, .ColIndex("BudgetQty")) = OrderQty
        .TextMatrix(i, .ColIndex("BudgetVendor")) = Vendor
        .TextMatrix(i, .ColIndex("BudgetVendorName")) = VendorName
        .TextMatrix(i, .ColIndex("BudgetRate")) = price
        .TextMatrix(i, .ColIndex("BudgetPretax")) = Round(OrderQty * price, 2)
        .TextMatrix(i, .ColIndex("BudgetTaxGroup")) = TaxGroup
        .TextMatrix(i, .ColIndex("BudgetJCTaxRate")) = JCTaxRate
        .TextMatrix(i, .ColIndex("BudgetNJCTaxRate")) = NJCTaxRate
        .TextMatrix(i, .ColIndex("BudgetJCTax")) = Round(.ValueMatrix(i, .ColIndex("BudgetPretax")) * JCTaxRate / 100, 2)
        .TextMatrix(i, .ColIndex("BudgetNJCTax")) = Round(.ValueMatrix(i, .ColIndex("BudgetPretax")) * NJCTaxRate / 100, 2)
        .TextMatrix(i, .ColIndex("BudgetTax")) = .ValueMatrix(i, .ColIndex("BudgetNJCTax")) + .ValueMatrix(i, .ColIndex("BudgetJCTax"))

        .TextMatrix(i, .ColIndex("RoomLocation")) = Location
        For w = 1 To 40
            .TextMatrix(i, .ColIndex("WBS" & format(w, "00"))) = WBS(w)
        Next

    End With
    CalcTotals
    
Exit Sub
eh: Call errHandler(SRCFILE & "AddItem")
End Sub





Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim s As String
    
    Select Case Button.Key
    
        Case "Attachments"
            s = "EditAttachments|" & mAttachmentID & "|Files"
            Call HFApp.RunTask(s)
        
        Case "Save"
            If SaveData(False) Then Unload Me
        
        Case "Preview"
            s = PathAppend(HFApp.SystemFolder, "System\Reports\CustomPreEstimate.rpt")
            'Call FRptViewer.ShowReport(s, True, True, "vCustomer", mCustomer, "Seq", mSeq)
        
            Dim c As New ZybUtil.Crystal
            Call c.LoadODBCReport(s, HFApp.LoginDSN, HFApp.LoginDBUID, HFApp.LoginDBPWD)
            On Error Resume Next
            Call c.ParameterValue("DivisionID", HFApp.DivisionID)
            Call c.ParameterValue("vCustomer", mCustomer)
            Call c.ParameterValue("Seq", mSeq)
            On Error GoTo 0
            Call c.PrintPreview("Print Preview")
        
        Case "TakeoffOneTime", "TakeoffItem", "TakeoffAssembly", "TakeoffCustom"
            Call FTakeoff.Takeoff(Me, mLocation <> olScheduleB, Mid(Button.Key, 8), 0, "", atCustomOption, mAssemblyDesc, mCommunity, mCommunityPhase, mModel, "", "", mJob)
    


    End Select
End Sub


Private Sub txtCost_GotFocus()
    SelectAll txtCost
End Sub

Private Sub txtMarkup_GotFocus()
    SelectAll txtMarkup
End Sub
Private Sub txtMarkup_Validate(Cancel As Boolean)
    If Val(txtMarkup.Text) <> 0 Then txtMarkup.Text = Val(txtMarkup.Text)
End Sub

Private Sub txtMargin_GotFocus()
    SelectAll txtMargin
End Sub
Private Sub txtMargin_Validate(Cancel As Boolean)
    If Val(txtMargin.Text) <> 0 Then txtMargin.Text = Val(txtMargin.Text)
End Sub





Private Sub txtSalesQty_Change()
    mDirty = True
End Sub

Private Sub txtSalesQty_GotFocus()
    SelectAll txtSalesQty
End Sub

Private Sub txtSalesQty_Validate(Cancel As Boolean)
    mSalesQty = Val(txtSalesQty.Text)
    txtSalesQty.Text = mSalesQty
    Call UpdatePrice
End Sub

Private Sub txtSell_GotFocus()
    SelectAll txtSell
End Sub


Private Sub LoadInfoPanel()
On Error GoTo eh
    Dim s As String
    Dim rs As Recordset
    s = ""
    s = s & "select isnull(c.community,'') + isnull('.' + c.phase,'') + ' - ' + isnull(l.description,'') + isnull(' ' + p.description,'') Community" & vbCrLf
    s = s & "      ,isnull(c.customer_no,'** not found **') + isnull(' - ' + c.Description,'') Customer" & vbCrLf
    s = s & "      ,isnull(c.model,'') + isnull(' - ' + m.description,'   ** not found **') Model" & vbCrLf
    s = s & "      ,isnull(c.Series,'') + isnull(' - ' + s.description,'   ** not found **') Series" & vbCrLf
    s = s & "      ,isnull(c.ModelAssembly,'') + isnull(' - ' + a.description,'   ** not found **') Assembly" & vbCrLf
    s = s & "from tblcustomers c" & vbCrLf
    s = s & "     LEFT OUTER JOIN tbllocality l on(c.community=l.area)" & vbCrLf
    s = s & "     LEFT OUTER JOIN communityphase p on(c.community=p.community and c.phase=p.communityphase)" & vbCrLf
    s = s & "     LEFT OUTER JOIN DistinctDivisionModelSeries m on(c.model=m.model and c.series=m.series and c.divisionid=m.divisionid)" & vbCrLf
    s = s & "     LEFT OUTER JOIN tblseries s on(m.series=s.series and m.DivisionID = s.DivisionID)" & vbCrLf
    s = s & "     LEFT OUTER JOIN tblDBAssemblyMaster a on(c.DivisionID = a.DivisionID and c.community=a.community and c.modelassembly=a.assembly and a.model=c.model and a.optionid='')" & vbCrLf
    s = s & "where c.customer_no=" & DbQuote(Str, mCustomer)
    
    Set rs = HFApp.SqlExec(s, dbHomefront)
    With gInfoPanel
        .TextMatrix(0, 1) = "" & rs(0)
        .TextMatrix(1, 1) = "" & rs(1)
        .TextMatrix(2, 1) = "" & rs(2)
        .TextMatrix(3, 1) = "" & rs(3)
        .TextMatrix(4, 1) = "" & rs(4)
    End With
    
    Me.Caption = "Prepare Quote"
    
    
    lblOptDesc.Caption = mAssemblyDesc
    txtSalesQty.Text = mSalesQty
    Call UpdatePrice
    
eh: Exit Sub
End Sub

Private Sub UpdatePrice()
    lblPrice.Caption = "@ " & format(mSell, "#,##0.00") & " = " & format(mSalesQty * mSell, "$#,##0.00")
End Sub
Private Sub txtSell_Validate(Cancel As Boolean)
    mSell = Dollars(txtSell.Text)
    txtSell.Text = format(mSell, "#,##0.00")
End Sub
Private Sub txtCost_Validate(Cancel As Boolean)
    mCost = Dollars(txtCost.Text)
    txtCost.Text = format(mCost, "#,##0.00")
End Sub

Private Sub gItems_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error Resume Next
    Dim r           As Long
    Dim rowBudgeted As Boolean
    Dim rowPOed     As Boolean
    Dim rate        As Double
    Dim rs          As Recordset
    Dim s           As String
    Dim i           As Long
    
    
    With gItems
    
    For r = Min(Row, .RowSel) To Max(Row, .RowSel)
    
        If r > 0 Then
    
        rowBudgeted = .ValueMatrix(r, .ColIndex("BudgetGenerated")) <> 0
        rowPOed = Trim(.TextMatrix(r, .ColIndex("PONumber"))) <> ""
    
        Select Case .ColKey(Col)
            Case "JCCategory", "JCCategoryDesc":
                s = ""
                s = s & "SELECT * FROM TaxGroups WHERE DivisionID = " & HFApp.DivisionID & " and TaxGroup=" & vbCrLf
                s = s & "dbo.Purch_GetDefaultTaxGroup(" & DbQuote(Str, .TextMatrix(r, .ColIndex("Job_No"))) & vbCrLf
                s = s & "   ," & DbQuote(Str, mCommunity) & vbCrLf
                s = s & "   ," & DbQuote(Str, mCommunityPhase) & vbCrLf
                s = s & "   ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Model"))) & vbCrLf
                s = s & "   ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Assembly"))) & vbCrLf
                s = s & "   ," & DbQuote(Str, .TextMatrix(r, .ColIndex("EstPhase"))) & vbCrLf
                s = s & "   ," & DbQuote(Str, .TextMatrix(r, .ColIndex("EstItem"))) & vbCrLf
                s = s & "   ," & DbQuote(Str, .TextMatrix(r, .ColIndex("BudgetVendor"))) & vbCrLf
                s = s & "   ," & DbQuote(Str, .TextMatrix(r, .ColIndex("JCCategory"))) & "," & HFApp.DivisionID & ")"
                Set rs = HFApp.SqlExec(s)
                If Not rs.EOF Then
                    If .Cell(flexcpChecked, r, .ColIndex("BudgetGenerated")) <> flexChecked Then
                        .TextMatrix(r, .ColIndex("BudgetTaxGroup")) = "" & rs("TaxGroup")
                        .TextMatrix(r, .ColIndex("BudgetJCTaxRate")) = "" & rs("JCRate")
                        .TextMatrix(r, .ColIndex("BudgetNJCTaxRate")) = "" & rs("NJCRate")
                        .TextMatrix(r, .ColIndex("BudgetJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetJCTaxRate")) / 100, 2)
                        .TextMatrix(r, .ColIndex("BudgetNJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetNJCTaxRate")) / 100, 2)
                        .TextMatrix(r, .ColIndex("BudgetTax")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetNJCTax"))
                    End If
                End If
                    
                    
            Case "ItemDesc":
            Case "ItemComments":
            
            Case "BudgetVendor", "BudgetVendorName":
                If GetVendorCost(r, .TextMatrix(r, .ColIndex("BudgetVendor")), rate) Then
                    .TextMatrix(r, .ColIndex("BudgetOverridden")) = "False"
                    .TextMatrix(r, .ColIndex("BudgetRate")) = rate
                    .TextMatrix(r, .ColIndex("BudgetPretax")) = Round(.ValueMatrix(r, .ColIndex("BudgetQty")) * .ValueMatrix(r, .ColIndex("BudgetRate")), 2)
                    .TextMatrix(r, .ColIndex("BudgetJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetJCTaxRate")) / 100, 2)
                    .TextMatrix(r, .ColIndex("BudgetNJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetNJCTaxRate")) / 100, 2)
                    .TextMatrix(r, .ColIndex("BudgetTax")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetNJCTax"))
                End If
            
            
            Case "TakeoffQty":
                .TextMatrix(r, .ColIndex("BudgetQty")) = RoundTo(.ValueMatrix(r, .ColIndex("TakeoffQty")) * (100 + .ValueMatrix(r, .ColIndex("WastePercent"))) / 100 * .ValueMatrix(r, .ColIndex("ConversionFactor")), .ValueMatrix(r, .ColIndex("RoundTo")), .ValueMatrix(r, .ColIndex("RoundDir")))
                .TextMatrix(r, .ColIndex("BudgetPretax")) = Round(.ValueMatrix(r, .ColIndex("BudgetQty")) * .ValueMatrix(r, .ColIndex("BudgetRate")), 2)
                .TextMatrix(r, .ColIndex("BudgetJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetJCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("BudgetNJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetNJCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("BudgetTax")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetNJCTax"))
                
            
            Case "BudgetQty":
                .TextMatrix(r, .ColIndex("TakeoffQty")) = .ValueMatrix(r, .ColIndex("BudgetQty")) / .ValueMatrix(r, .ColIndex("ConversionFactor"))
                .TextMatrix(r, .ColIndex("BudgetPretax")) = Round(.ValueMatrix(r, .ColIndex("BudgetQty")) * .ValueMatrix(r, .ColIndex("BudgetRate")), 2)
                .TextMatrix(r, .ColIndex("BudgetJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetJCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("BudgetNJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetNJCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("BudgetTax")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetNJCTax"))
            
            Case "BudgetRate":
                .TextMatrix(r, .ColIndex("BudgetOverridden")) = "True"
                .TextMatrix(r, .ColIndex("BudgetPretax")) = Round(.ValueMatrix(r, .ColIndex("BudgetQty")) * .ValueMatrix(r, .ColIndex("BudgetRate")), 2)
                .TextMatrix(r, .ColIndex("BudgetJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetJCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("BudgetNJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetNJCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("BudgetTax")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetNJCTax"))
            
            Case "BudgetPretax":
                If .ValueMatrix(r, .ColIndex("BudgetQty")) = 0 Then .TextMatrix(r, .ColIndex("BudgetQty")) = 1
                .TextMatrix(r, .ColIndex("BudgetRate")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) / .ValueMatrix(r, .ColIndex("BudgetQty")), 4)
                .TextMatrix(r, .ColIndex("BudgetOverridden")) = "True"
                .TextMatrix(r, .ColIndex("BudgetJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetJCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("BudgetNJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetNJCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("BudgetTax")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetNJCTax"))
            
            Case "BudgetTaxGroup":
                s = "SELECT JCRate,NJCRate FROM TaxGroups WHERE DivisionID = " & HFApp.DivisionID & " and TaxGroup=" & DbQuote(Str, .Cell(flexcpText, r, .ColIndex("BudgetTaxGroup")))
                Set rs = HFApp.SqlExec(s)
                .TextMatrix(r, .ColIndex("BudgetJCTaxRate")) = Val("" & rs("JCRate"))
                .TextMatrix(r, .ColIndex("BudgetNJCTaxRate")) = Val("" & rs("NJCRate"))
                .TextMatrix(r, .ColIndex("BudgetJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetJCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("BudgetNJCTax")) = Round(.ValueMatrix(r, .ColIndex("BudgetPretax")) * .ValueMatrix(r, .ColIndex("BudgetNJCTaxRate")) / 100, 2)
                .TextMatrix(r, .ColIndex("BudgetTax")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetNJCTax"))
                
            Case "BudgetJCTax":
                .TextMatrix(r, .ColIndex("BudgetTax")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetNJCTax"))
            
            Case "BudgetNJCTax":
                .TextMatrix(r, .ColIndex("BudgetTax")) = .ValueMatrix(r, .ColIndex("BudgetJCTax")) + .ValueMatrix(r, .ColIndex("BudgetNJCTax"))
            
        End Select
        
        Call ColorizeItems(-1)
        Call CalcTotals
        
        If .ColKey(Col) <> "Selected" Then
            mDirty = True
            For i = r To .RowSel
                If .RowData(i) <> "NEW" Then .RowData(i) = "DIRTY"
            Next
        End If
        
        End If
    Next
    End With
End Sub

Private Sub gItems_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
Static inHere As Boolean
If inHere Then Exit Sub
inHere = True
    gItems.ColSel = gItems.Col
inHere = False
End Sub

Private Sub gItems_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gItems
        .ComboList = ""
        .EditMaxLength = 0
        .AutoSearch = flexSearchNone
        
        
        Select Case .ColKey(Col)
            Case "JCExtra":            .ComboList = IIf(HFApp.Options(Use_Timberline), "|...", ""):   .EditMaxLength = 10
            Case "JCCostCode":         .ComboList = "|..."
            Case "JCCostCodeDesc":     .ComboList = "..."
            Case "JCCategory":         .ComboList = "|..."
            Case "JCCategoryDesc":     .ComboList = "..."
            Case "ItemDesc":           .ComboList = "|...":        .EditMaxLength = 200
            Case "ItemComments":       .ComboList = "|...":        .EditMaxLength = 2000
            Case "BudgetVendor":       .ComboList = "|..."
            Case "BudgetVendorName":   .ComboList = "..."
            Case "TakeoffQty":
            Case "BudgetQty":
            Case "BudgetRate":
            'Case "OrderUOM":           .ComboList = "|...":        .EditMaxLength = 10
            Case "BudgetPretax":
            Case "BudgetTaxGroup":     .ComboList = "|..."
            Case "BudgetJCTax":
            Case "BudgetNJCTax":
            Case "POIndex":            .ComboList = "|..."
            Case "RoomLocation", "WBS01", "WBS02", "WBS03", "WBS04", "WBS05", "WBS06", "WBS07", "WBS08", "WBS09", "WBS10", "WBS11", "WBS12", "WBS13", "WBS14", "WBS15", "WBS16", "WBS17", "WBS18", "WBS19", "WBS20", "WBS21", "WBS22", "WBS23", "WBS24", "WBS25", "WBS26", "WBS27", "WBS28", "WBS29", "WBS30", "WBS31", "WBS32", "WBS33", "WBS34", "WBS35", "WBS36", "WBS37", "WBS38", "WBS39", "WBS40"
                .EditMaxLength = 50
            Case Else:                 Cancel = True
                
        End Select
        
        .AutoSearch = IIf(Cancel, flexSearchFromCursor, flexSearchNone)
        
    End With
End Sub


Private Sub gItems_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error Resume Next
    If Button = vbRightButton Then
        Cancel = True
        If gItems.MouseRow = 0 Then
            Call FMain.ShowColumnMenu(gItems, , , , False)
        Else
            FMain.mnuEstimateItemsGridSub(mcITEM_REMOVEITEMS).Enabled = True
            FMain.mnuEstimateItemsGridSub(mcITEM_SPLITITEMS).Enabled = False
            FMain.mnuEstimateItemsGridSub(mcITEM_SUBSTITUEITEM).Enabled = False
            FMain.mnuEstimateItemsGridSub(mcITEM_COPYITEMS).Enabled = False
            FMain.mnuEstimateItemsGridSub(mcITEM_VIEWFILES).Enabled = False
            FMain.mnuEstimateItemsGridSub(mcITEM_CANCELBUDGETS).Enabled = False
            FMain.mnuEstimateItemsGridSub(mcITEM_SAVEONETIMETODB).Enabled = False
            FMain.mnuEstimateItemsGridSub(mcITEM_UPDATEPRICELIST).Enabled = False
            FMain.mnuEstimateItemsGridSub(mcITEM_COMPAREPRICES).Enabled = False
            FMain.mnuEstimateItemsGridSub(mcITEM_FORMATTING).Enabled = False
         
            PopupMenu FMain.mnuEstimateItemsGrid
            
            FMain.mnuEstimateItemsGridSub(mcITEM_VIEWFILES).Enabled = True
            
        End If
    End If

End Sub





Private Sub gItems_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    If Col < gItems.FrozenCols Then
        Position = Col
    Else
        If Position < gItems.FrozenCols Then
            Position = gItems.FrozenCols
        End If
    End If
End Sub


Private Sub gItems_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case gItems.ColKey(Col)
        Case "WarningMessages", "Selected"
            Cancel = True
    End Select
End Sub


Private Sub gItems_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim s As String
    Dim i As Long
    With gItems
        Select Case .ColKey(Col)
        
            Case "JCExtra"
                s = "select Extra,Description from jcm_master__extra where job=" & DbQuote(Str, HFApp.FormatJob(.TextMatrix(Row, .ColIndex("Job_No"))))
                If FPickList.Choose(HFApp.Databases(dbAccounting), "Extra", s, gItems) Then
                    .Cell(flexcpText, Row, .ColIndex("JCExtra"), .RowSel, .ColIndex("JCExtra")) = FPickList.SelectedItem("Extra")
                End If
                
            Case "JCCostCode"
                s = "SELECT CostCode,Description FROM StandardCostCodes where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Cost Code", s, gItems) Then
                    .Cell(flexcpText, Row, .ColIndex("JCCostCode"), .RowSel, .ColIndex("JCCostCode")) = FPickList.SelectedItem("CostCode")
                    .Cell(flexcpText, Row, .ColIndex("JCCostCodeDesc"), .RowSel, .ColIndex("JCCostCodeDesc")) = FPickList.SelectedItem("Description")
                End If
                
            Case "JCCostCodeDesc"
                s = "SELECT Description,CostCode FROM StandardCostCodes where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Cost Code", s, gItems) Then
                    .Cell(flexcpText, Row, .ColIndex("JCCostCode"), .RowSel, .ColIndex("JCCostCode")) = FPickList.SelectedItem("CostCode")
                    .Cell(flexcpText, Row, .ColIndex("JCCostCodeDesc"), .RowSel, .ColIndex("JCCostCodeDesc")) = FPickList.SelectedItem("Description")
                End If
            
            Case "JCCategory"
                s = "SELECT Category,Description FROM StandardCategories where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Category", s, gItems) Then
                    .Cell(flexcpText, Row, .ColIndex("JCCategory"), .RowSel, .ColIndex("JCCategory")) = FPickList.SelectedItem("Category")
                    .Cell(flexcpText, Row, .ColIndex("JCCategoryDesc"), .RowSel, .ColIndex("JCCategoryDesc")) = FPickList.SelectedItem("Description")
                End If
                
            Case "JCCategoryDesc"
                s = "SELECT Description,Category FROM StandardCategories where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Category", s, gItems) Then
                    .Cell(flexcpText, Row, .ColIndex("JCCategory"), .RowSel, .ColIndex("JCCategory")) = FPickList.SelectedItem("Category")
                    .Cell(flexcpText, Row, .ColIndex("JCCategoryDesc"), .RowSel, .ColIndex("JCCategoryDesc")) = FPickList.SelectedItem("Description")
                End If
            
            Case "OrderUOM"
                s = "SELECT DISTINCT OrderUOM Unit FROM tblPhaseItem where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Units", s, gItems) Then
                    .Cell(flexcpText, Row, .ColIndex("OrderUOM"), .RowSel, .ColIndex("OrderUOM")) = FPickList.SelectedItem("OrderUOM")
                End If
            
            Case "ItemComments"
                s = gItems.Text
                If FComments.Edit(s, gItems, , , 2000) Then
                    gItems.Text = s
                End If
                
            Case "ItemDesc":
                s = gItems.Text
                If FComments.Edit(s, gItems, , "Description", 200) Then
                    gItems.Text = s
                End If
                
            Case "BudgetVendor"
                s = "Approved Vendors" & Chr(1) & "SELECT DISTINCT v.Vendor_ID Vendor, v.Vendor_Name Company,v.City,v.Phone FROM tblVendors v JOIN tblVendorCost c ON(v.Vendor_ID=c.Vendor and v.DivisionID = c.DivisionID) WHERE v.DivisionID = " & HFApp.DivisionID & " and (ISNULL(Community,'')='' OR " & " Community=" & DbQuote(Str, mCommunity) & ") AND (ISNULL(Assembly,'')='' OR " & " Assembly=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("Assembly"))) & ") AND Phase=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("EstPhase"))) & " AND Item=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("EstItem"))) & Chr(0) & _
                    "All Vendors" & Chr(1) & "SELECT Vendor_ID Vendor, Vendor_Name Company,City,Phone FROM tblVendors WHERE DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Vendor", s, gItems, , , FMain.SmallIcons.ListImages("vendor").Picture) Then
                    .Cell(flexcpText, Row, .ColIndex("BudgetVendor"), .RowSel, .ColIndex("BudgetVendor")) = FPickList.SelectedItem("Vendor")
                    .Cell(flexcpText, Row, .ColIndex("BudgetVendorName"), .RowSel, .ColIndex("BudgetVendorName")) = FPickList.SelectedItem("Company")
                End If
                
            Case "BudgetVendorName"
                s = "Approved Vendors" & Chr(1) & "SELECT DISTINCT v.Vendor_Name Company, v.Vendor_ID Vendor,v.City,v.Phone FROM tblVendors v JOIN tblVendorCost c ON(v.Vendor_ID=c.Vendor and v.DivisionID = c.DivisionID) WHERE v.DivisionID = " & HFApp.DivisionID & " and (ISNULL(Community,'')='' OR " & " Community=" & DbQuote(Str, mCommunity) & ") AND (ISNULL(Assembly,'')='' OR " & " Assembly=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("Assembly"))) & ") AND Phase=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("EstPhase"))) & " AND Item=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("EstItem"))) & Chr(0) & _
                    "All Vendors" & Chr(1) & "SELECT Vendor_Name Company,Vendor_ID Vendor,City,Phone FROM tblVendors WHERE DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Vendor", s, gItems, , , FMain.SmallIcons.ListImages("vendor").Picture, "Vendor") Then
                    .Cell(flexcpText, Row, .ColIndex("BudgetVendor"), .RowSel, .ColIndex("BudgetVendor")) = FPickList.SelectedItem("Vendor")
                    .Cell(flexcpText, Row, .ColIndex("BudgetVendorName"), .RowSel, .ColIndex("BudgetVendorName")) = FPickList.SelectedItem("Company")
                End If
            
            Case "POVendor"
                s = "Approved Vendors" & Chr(1) & "SELECT DISTINCT v.Vendor_ID Vendor, v.Vendor_Name Company,v.City,v.Phone FROM tblVendors v JOIN tblVendorCost c ON(v.Vendor_ID=c.Vendor and v.DivisionID = c.DivisionID) WHERE  v.DivisionID = " & HFApp.DivisionID & " and(ISNULL(Community,'')='' OR " & " Community=" & DbQuote(Str, mCommunity) & ") AND (ISNULL(Assembly,'')='' OR " & " Assembly=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("Assembly"))) & ") AND Phase=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("EstPhase"))) & " AND Item=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("EstItem"))) & Chr(0) & _
                    "All Vendors" & Chr(1) & "SELECT Vendor_ID Vendor, Vendor_Name Company,City,Phone FROM tblVendors WHERE DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Vendor", s, gItems, , , FMain.SmallIcons.ListImages("vendor").Picture) Then
                    .Cell(flexcpText, Row, .ColIndex("POVendor"), .RowSel, .ColIndex("POVendor")) = FPickList.SelectedItem("Vendor")
                    .Cell(flexcpText, Row, .ColIndex("POVendorName"), .RowSel, .ColIndex("POVendorName")) = FPickList.SelectedItem("Company")
                End If
            
            Case "POVendorName"
                s = "Approved Vendors" & Chr(1) & "SELECT DISTINCT v.Vendor_Name Company, v.Vendor_ID Vendor,v.City,v.Phone FROM tblVendors v JOIN tblVendorCost c ON(v.Vendor_ID=c.Vendor and v.DivisionID = c.DivisionID) WHERE v.DivisionID = " & HFApp.DivisionID & " and (ISNULL(Community,'')='' OR " & " Community=" & DbQuote(Str, mCommunity) & ") AND (ISNULL(Assembly,'')='' OR " & " Assembly=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("Assembly"))) & ") AND Phase=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("EstPhase"))) & " AND Item=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("EstItem"))) & Chr(0) & _
                    "All Vendors" & Chr(1) & "SELECT Vendor_Name Company,Vendor_ID Vendor,City,Phone FROM tblVendors WHERE DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Vendor", s, gItems, , , FMain.SmallIcons.ListImages("vendor").Picture, "Vendor") Then
                    .Cell(flexcpText, Row, .ColIndex("POVendor"), .RowSel, .ColIndex("POVendor")) = FPickList.SelectedItem("Vendor")
                    .Cell(flexcpText, Row, .ColIndex("POVendorName"), .RowSel, .ColIndex("POVendorName")) = FPickList.SelectedItem("Company")
                End If

            Case "BudgetTaxGroup"
                s = "SELECT TaxGroup,Description FROM TaxGroups where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Tax Groups", s, gItems) Then
                    .Cell(flexcpText, Row, Col, .RowSel, Col) = FPickList.SelectedItem("TaxGroup")
                End If
        
            Case "POTaxGroup"
                s = "SELECT TaxGroup,Description FROM TaxGroups where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Tax Groups", s, gItems) Then
                    .Cell(flexcpText, Row, Col, .RowSel, Col) = FPickList.SelectedItem("TaxGroup")
                End If
            
            Case "POIndex"
                s = "SELECT POIndex FROM tblPOIndex where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "PO Index", s, gItems) Then
                    .Cell(flexcpText, Row, Col, .RowSel, Col) = FPickList.SelectedItem("POIndex")
                End If

        End Select
        
        Call gItems_AfterEdit(Row, Col)
        
    End With
End Sub




Private Sub gItems_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Dim r As Long
    
    With gItems
    
'        If .Row = 1 Then
'            Cancel = False
'            .AutoSearch = flexSearchNone
'        End If
        Select Case True
            Case KeyCode = vbKeyF And Shift = vbCtrlMask
                Call FFind.ShowForm(gItems)
                
            Case KeyCode = vbKeyDelete And Shift = vbCtrlMask And Not mReadOnly
                For r = Max(.Row, .RowSel) To Min(.Row, .RowSel) Step -1
                    mDirty = True
                    Call .RemoveItem(r)
                Next
                Call CalcTotals
        
        End Select
    End With
    
End Sub


Private Sub gItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If Button <> 0 Then Exit Sub
    With gItems
        If mFunnyFlag Then
            mFunnyFlag = False
            Exit Sub
        End If
        If .MouseRow >= Min(.Row, .RowSel) And .MouseRow <= Max(.Row, .RowSel) And .MouseCol = .Col Then
            Call gItems_SelChange
        Else
            Call ShowTip(.MouseRow, .MouseCol, MouseX(gItems.hwnd) * Screen.TwipsPerPixelX + 315, MouseY(gItems.hwnd) * Screen.TwipsPerPixelY + 315)
        End If
    End With
End Sub


Private Sub gItems_RowColChange()
On Error Resume Next
    With gItems
        mFunnyFlag = True
        Call ShowTip(.Row, .Col, .colPos(.Col) + 180, .RowPos(.Row) + .RowHeight(.Row) + 180)
    End With
End Sub

Private Sub ShowTip(Row As Long, Col As Long, X As Long, Y As Long, Optional tip As String)
    With gItems
        If Row < 0 Or Col < 0 Then
            picWarningMessages.Visible = False
        Else
            lblWarningMessages = IIf(tip <> "", tip, gItems.Cell(flexcpData, Row, Col))
            picWarningMessages.Visible = lblWarningMessages <> ""
            
            Set Me.Font = lblWarningMessages.Font
            
            lblWarningMessages.Alignment = IIf(tip <> "", 1, 0)
            lblWarningMessages.Width = Me.TextWidth(lblWarningMessages) + lblWarningMessages.Left
            lblWarningMessages.Height = Me.TextHeight(lblWarningMessages) + lblWarningMessages.Top
            
            picWarningMessages.Width = lblWarningMessages.Width + lblWarningMessages.Left + lblWarningMessages.Top
            picWarningMessages.Height = lblWarningMessages.Height + 2 * lblWarningMessages.Top
            If .Height - Y - picWarningMessages.Height - 365 < 0 Then Y = .Height - picWarningMessages.Height - 365
            If .Width - X - picWarningMessages.Width - 365 < 0 Then X = .Width - picWarningMessages.Width - 365
            picWarningMessages.Top = Y
            picWarningMessages.Left = X
            
            shpWarningMessages.Width = picWarningMessages.Width
            shpWarningMessages.Height = picWarningMessages.Height
        End If
    End With
End Sub




Private Sub gItems_SelChange()
    Dim r As Long
    
    Dim Count As Long
    Dim pretax As Double
    Dim tax As Double
    Dim tip As String
    
    With gItems
        If .Row <> .RowSel Then
            For r = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                Count = Count + 1
                pretax = pretax + .ValueMatrix(r, .ColIndex("BudgetPretax"))
                tax = tax + .ValueMatrix(r, .ColIndex("BudgetJCTax"))
            Next
            
            tip = "" & _
                  "Pretax:" & vbTab & format(pretax, "#,##0.00") & vbCrLf & _
                  "Tax:" & vbTab & format(tax, "#,##0.00") & vbCrLf & _
                  "Total:" & vbTab & format(pretax + tax, "#,##0.00")

            Call ShowTip(.MouseRow, .MouseCol, MouseX(gItems.hwnd) * Screen.TwipsPerPixelX + 315, MouseY(gItems.hwnd) * Screen.TwipsPerPixelY + 315, tip)
        End If
    End With
End Sub



Private Sub gItems_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim s As String
    Dim i As Long
    Dim Description As String
    Dim rs As Recordset
    
    With gItems
        s = .EditText
        Select Case .ColKey(Col)
            
            Case "JCExtra"
                If HFApp.Options(Use_Timberline) And Trim(s) <> "" Then
                    Set rs = HFApp.SqlExec("select extra from jcm_master__extra where job=" & DbQuote(Str, HFApp.FormatJob(.TextMatrix(Row, .ColIndex("Job_No")))) & " AND extra=" & DbQuote(Str, s), dbAccounting)
                    If rs.EOF Then
                        Description = InputBox(vbCrLf & "Extra '" & s & "' could not be found." & vbCrLf & "Do you want to add it?" & vbCrLf & vbCrLf & vbCrLf & "Enter a description for the extra. (30 characters or less)", App.ProductName, mAssemblyDesc)
                        If Description = "" Then
                            Cancel = True
                        Else
                            s = ""
                            s = s & "INSERT INTO jcm_master__extra(Job,Extra,Description,Status)" & vbCrLf
                            s = s & "VALUES(" & DbQuote(Str, HFApp.FormatJob(.TextMatrix(Row, .ColIndex("Job_No")))) & vbCrLf
                            s = s & "      ," & DbQuote(Str, .EditText) & vbCrLf
                            s = s & "      ," & DbQuote(Str, Left(Description, 30)) & vbCrLf
                            s = s & "      ,'In progress')" & vbCrLf
                            Set rs = HFApp.SqlExec(s, dbAccounting)
                            s = .EditText
                        End If
                    Else
                        s = "" & rs(0)
                    End If
                End If
            
                                   
            Case "JCCostCode":              Cancel = Not ValidateField(gItems, s, "Cost Code not found", "SELECT CostCode,Description FROM StandardCostCodes WHERE DivisionID = " & HFApp.DivisionID & " and CostCode=" & DbQuote(Str, s), "JCCostCodeDesc")
            'Case "JCCostCodeDesc":
            'case "OriginalJCCategory":
            Case "JCCategory":              Cancel = Not ValidateField(gItems, s, "Cost Code not found", "SELECT Category,Description FROM StandardCategories WHERE DivisionID = " & HFApp.DivisionID & " and Category=" & DbQuote(Str, s), "JCCategoryDesc")
            'Case "JCCategoryDesc":
            'case "HFDescription":
            'case "EstPhase":
            'case "EstItem":
            Case "ItemDesc":
            Case "ItemComments":
            Case "BudgetVendor":            Cancel = Not ValidateField(gItems, s, "Vendor not found", "SELECT Vendor_ID,Vendor_Name FROM tblVendors WHERE DivisionID = " & HFApp.DivisionID & " and Vendor_ID=" & DbQuote(Str, s), "BudgetVendorName")
            Case "BudgetQty":               Cancel = Not IsNumeric(s)
            Case "OrderUOM":
            Case "BudgetRate":              Cancel = Not IsNumeric(s)
            Case "BudgetPretax":            Cancel = Not IsNumeric(s)
            Case "BudgetTaxGroup":          Cancel = Not ValidateField(gItems, s, "Tax Group not found", "SELECT TaxGroup FROM TaxGroups WHERE DivisionID = " & HFApp.DivisionID & " and TaxGroup=" & DbQuote(Str, s))
            Case "BudgetJCTax":             Cancel = Not IsNumeric(s)
            Case "BudgetNJCTax":            Cancel = Not IsNumeric(s)
            Case "POIndex":                 Cancel = Not ValidateField(gItems, s, "PO Index not found", "SELECT POIndex FROM tblPOIndex WHERE DivisionID = " & HFApp.DivisionID & " and POIndex=" & DbQuote(Str, s))
        End Select
        .EditText = s
        If .ColKey(Col) <> "Selected" Then
            mDirty = True
            For i = Min(Row, .RowSel) To Max(Row, .RowSel)
                If .RowData(i) <> "NEW" Then .RowData(i) = "DIRTY"
            Next
        End If
    End With
End Sub




Private Function GetVendorCost(Row As Long, Vendor As String, rate As Double) As Boolean
    Dim s As String
    Dim r As Double
    With gItems
    
        s = ""
        s = s & "SELECT dbo.Purch_GetItemRate(" & vbCrLf
        s = s & "       0" & vbCrLf
        s = s & "      ,0" & vbCrLf
        s = s & "      ," & DbQuote(Str, mCommunity) & vbCrLf
        s = s & "      ," & DbQuote(Str, mCommunityPhase) & vbCrLf
        s = s & "      ," & DbQuote(Str, .TextMatrix(Row, .ColIndex("Assembly"))) & vbCrLf
        s = s & "      ," & DbQuote(Str, .TextMatrix(Row, .ColIndex("Model"))) & vbCrLf
        s = s & "      ," & DbQuote(Str, .TextMatrix(Row, .ColIndex("OptionID"))) & vbCrLf
        s = s & "      ," & DbQuote(Str, .TextMatrix(Row, .ColIndex("EstPhase"))) & vbCrLf
        s = s & "      ," & DbQuote(Str, .TextMatrix(Row, .ColIndex("EstItem"))) & vbCrLf
        s = s & "      ," & DbQuote(Num, .TextMatrix(Row, .ColIndex("Seq"))) & vbCrLf
        s = s & "      ," & DbQuote(Str, Vendor) & vbCrLf
        s = s & "      ,GETDATE()"
        s = s & "      ," & HFApp.DivisionID & ")"
        On Error Resume Next
        r = HFApp.SqlExec(s)(0)
        
        If r <> 0 Or HFApp.Options(ZeroRateOnChangeVendor) Then
            GetVendorCost = True
            rate = r
        End If
        
    
    End With
End Function




Public Sub ColorizeItems(r As Long)
    Dim s           As String
    Dim startRow    As Long
    Dim EndRow      As Long
        
    With gItems
        If r > 0 Then
            startRow = r
            EndRow = r
        Else
            startRow = 1
            EndRow = .Rows - 1
        End If
        For r = startRow To EndRow
        
            s = ""
            
            'clear all
            .Cell(flexcpForeColor, r, 0, r, .Cols - 1) = vbWindowText
            .Cell(flexcpBackColor, r, 0, r, .Cols - 1) = vbWindowBackground
            .Cell(flexcpFontItalic, r, 0, r, .Cols - 1) = False
            .Cell(flexcpFontBold, r, 0, r, .Cols - 1) = False
            
            
            
            .Cell(flexcpForeColor, r, 0, r, .Cols - 1) = vbWindowText
            
            ' less than zero
            If .ValueMatrix(r, .ColIndex("BudgetRate")) < 0 Or .ValueMatrix(r, .ColIndex("BudgetQty")) < 0 Then
                .Cell(flexcpForeColor, r, 0, r, .Cols - 1) = HFApp.Options(Format_QtyRateLTZero_ForeColor)
                .Cell(flexcpBackColor, r, 0, r, .Cols - 1) = HFApp.Options(Format_QtyRateLTZero_BackColor)
                .Cell(flexcpFontItalic, r, 0, r, .Cols - 1) = InStr(1, HFApp.Options(Format_QtyRateLTZero_FontStyle), "Italic")
                .Cell(flexcpFontBold, r, 0, r, .Cols - 1) = InStr(1, HFApp.Options(Format_QtyRateLTZero_FontStyle), "Bold")
            End If
            
            ' equal to zero
            If .ValueMatrix(r, .ColIndex("BudgetPretax")) = 0 Then
                .Cell(flexcpForeColor, r, 0, r, .Cols - 1) = HFApp.Options(Format_QtyRateEQZero_ForeColor)
                .Cell(flexcpBackColor, r, 0, r, .Cols - 1) = HFApp.Options(Format_QtyRateEQZero_BackColor)
                .Cell(flexcpFontItalic, r, 0, r, .Cols - 1) = InStr(1, HFApp.Options(Format_QtyRateEQZero_FontStyle), "Italic")
                .Cell(flexcpFontBold, r, 0, r, .Cols - 1) = InStr(1, HFApp.Options(Format_QtyRateEQZero_FontStyle), "Bold")
            End If
            
            
            ' if overridden value
            If .TextMatrix(r, .ColIndex("BudgetOverridden")) = "True" Then
                .Cell(flexcpForeColor, r, 0, r, .Cols - 1) = HFApp.Options(Format_NonSysRate_ForeColor)
                .Cell(flexcpBackColor, r, 0, r, .Cols - 1) = HFApp.Options(Format_NonSysRate_BackColor)
                .Cell(flexcpFontItalic, r, 0, r, .Cols - 1) = InStr(1, HFApp.Options(Format_NonSysRate_FontStyle), "Italic")
                .Cell(flexcpFontBold, r, 0, r, .Cols - 1) = InStr(1, HFApp.Options(Format_NonSysRate_FontStyle), "Bold")
            End If
            
            If .TextMatrix(r, .ColIndex("POOverridden")) = "True" Then
                .Cell(flexcpForeColor, r, 0, r, .Cols - 1) = HFApp.Options(Format_NonSysRate_ForeColor)
                .Cell(flexcpBackColor, r, 0, r, .Cols - 1) = HFApp.Options(Format_NonSysRate_BackColor)
                .Cell(flexcpFontItalic, r, 0, r, .Cols - 1) = InStr(1, HFApp.Options(Format_NonSysRate_FontStyle), "Italic")
                .Cell(flexcpFontBold, r, 0, r, .Cols - 1) = InStr(1, HFApp.Options(Format_NonSysRate_FontStyle), "Bold")
            End If
            
            
            ' figure out if there are errors
            If Trim(.TextMatrix(r, .ColIndex("POIndex"))) = "" Then s = s & vbCrLf & "PO Index not specified"
            If Trim(.TextMatrix(r, .ColIndex("JCCostCode"))) = "" Then s = s & vbCrLf & "Cost Code not specified"
            If Trim(.TextMatrix(r, .ColIndex("JCCategory"))) = "" Then s = s & vbCrLf & "Category not specified"
            If Trim(.TextMatrix(r, .ColIndex("BudgetVendor"))) = "" Then s = s & vbCrLf & "Vendor not specified"
            If Trim(.TextMatrix(r, .ColIndex("BudgetTaxGroup"))) = "" Then s = s & vbCrLf & "Tax Group not specified"
            gItems.Cell(flexcpData, r, 0, r, .Cols - 1) = Mid(s, 3)
            If s = "" Then
                .Cell(flexcpPicture, r, .ColIndex("WarningMessages")) = Nothing
            Else
                .Cell(flexcpPicture, r, .ColIndex("WarningMessages")) = imgWarning.Picture
                .Cell(flexcpForeColor, r, 0, r, .Cols - 1) = HFApp.Options(Format_InvalidData_ForeColor)
                .Cell(flexcpBackColor, r, 0, r, .Cols - 1) = HFApp.Options(Format_InvalidData_BackColor)
                .Cell(flexcpFontItalic, r, 0, r, .Cols - 1) = InStr(1, HFApp.Options(Format_InvalidData_FontStyle), "Italic")
                .Cell(flexcpFontBold, r, 0, r, .Cols - 1) = InStr(1, HFApp.Options(Format_InvalidData_FontStyle), "Bold")
            End If
        Next
    End With
End Sub


Public Function SaveData(prompt As Boolean) As Boolean
On Error GoTo eh
    Dim i As Long
    Dim w As Long
    Dim s As String
    
    If Not mDirty Then
        SaveData = True
        Exit Function
    End If
    If prompt Then
        Select Case MsgBox("This data has changed." & vbCrLf & vbCrLf & "Do you want to save these changes?" & vbCrLf, vbExclamation + vbYesNoCancel, App.ProductName)
            Case vbNo
                SaveData = True
                Exit Function
            Case vbCancel
                SaveData = False
                Exit Function
        End Select
    End If
    If Not ValidateData Then Exit Function
    
    Screen.MousePointer = vbHourglass
    
    mSaved = True
                
    With gItems
            
        s = ""
        s = s & "delete from CustomPreEstimateItems" & vbCrLf
        s = s & "where customerno=" & DbQuote(Str, mCustomer) & vbCrLf
        s = s & "  and location=" & DbQuote(Num, mLocation) & vbCrLf
        s = s & "  and seq=" & DbQuote(Num, mSeq) & vbCrLf
        HFApp.SqlExec s
            
        For i = 1 To .Rows - 1
            s = ""
            s = s & "insert into CustomPreEstimateItems(Assembly,CustomerNo,Location,OptionID,Seq,POIndex,Phase,Item,Description,Comments,TakeoffUOM,TakeoffQty,ConversionFactor,OrderUOM,Job,JCExtra,JCCostCode,JCCategory,BudgetVendor,BudgetQty,BudgetRate,BudgetPretax,BudgetTaxGroup,BudgetJCTax,BudgetJCTaxRate,BudgetNJCTax,BudgetNJCTaxRate" & vbCrLf
            s = s & "   ,RoomLocation"
            For w = 1 To 40
                s = s & ",WBS" & format(w, "00")
            Next
            s = s & ")" & vbCrLf
            s = s & "values(" & DbQuote(Str, .TextMatrix(i, .ColIndex("Assembly")), , , 20) & vbCrLf
            s = s & "      ," & DbQuote(Str, mCustomer) & vbCrLf
            s = s & "      ," & DbQuote(Num, mLocation) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("OptionID")), , , 20) & vbCrLf
            s = s & "      ," & DbQuote(Num, mSeq) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("POIndex")), , , 20) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("EstPhase")), , , 20) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("EstItem")), , , 15) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("ItemDesc")), , , 200) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("ItemComments")), , , 4000) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("TakeoffUOM")), , , 4) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("TakeoffQty"))) & vbCrLf
            If .ValueMatrix(i, .ColIndex("ConversionFactor")) = 0 Then
                s = s & "      ," & DbQuote(Num, 1) & vbCrLf
            Else
                s = s & "      ," & DbQuote(Num, .ValueMatrix(i, .ColIndex("ConversionFactor"))) & vbCrLf
            End If
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("OrderUOM")), , , 4) & vbCrLf
            s = s & "      ," & DbQuote(Str, mJob, , , 12) & vbCrLf
            s = s & "      ," & DbQuote(Str, mExtra, , , 10) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("JCCostCode"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("JCCategory"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("BudgetVendor"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetQty"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetRate"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetPretax"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("BudgetTaxGroup"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetJCTax"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetJCTaxRate"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetNJCTax"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("BudgetNJCTaxRate"))) & vbCrLf
            s = s & "     ," & DbQuote(Str, .TextMatrix(i, .ColIndex("RoomLocation"))) & vbCrLf
            For w = 1 To 40
            s = s & "     ," & DbQuote(Str, .TextMatrix(i, .ColIndex("WBS" & format(w, "00")))) & vbCrLf
            Next
            s = s & ")"
            HFApp.SqlExec s
        Next
    End With
    
    SaveData = True
    mDirty = False
    Screen.MousePointer = vbDefault
    
Exit Function
eh: Call errHandler(SRCFILE & "SaveData", s)
    s = ""
    s = s & "delete from CustomPreEstimateItems" & vbCrLf
    s = s & "where customerno=" & DbQuote(Str, mCustomer) & vbCrLf
    s = s & "  and location=" & DbQuote(Num, mLocation) & vbCrLf
    s = s & "  and seq=" & DbQuote(Num, mSeq) & vbCrLf
    HFApp.SqlExec s
End Function

Public Function ValidateData() As Boolean
    ValidateData = True
End Function

Private Sub LoadData()
    Dim i As Long
    Dim w As Long
    Dim s As String
    Dim rs As Recordset
    
    s = ""
    s = s & "select cc.Description JCCostCodeDesc, cat.Description JCCategoryDesc, v.Vendor_Name BudgetVendorName,pi.roundto,pi.rounddir, i.*" & vbCrLf
    s = s & "  from CustomPreEstimateItems i" & vbCrLf
    s = s & "  left outer join tblphaseitem pi ON(pi.DivisionID = " & DbQuote(Num, HFApp.DivisionID) & " and i.phase=pi.phase and i.item=pi.item)" & vbCrLf
    s = s & "  left outer join tblVendors v   ON(i.BudgetVendor=v.Vendor_ID and v.DivisionID = " & HFApp.DivisionID & ")" & vbCrLf
    s = s & "  left outer join StandardCostCodes cc   ON(cc.DivisionID=" & DbQuote(Num, HFApp.DivisionID) & " and i.JCCostCode=cc.CostCode)" & vbCrLf
    s = s & "  left outer join StandardCategories cat ON(cat.DivisionID=" & DbQuote(Num, HFApp.DivisionID) & " and i.JCCategory=cat.Category)" & vbCrLf
    s = s & " where i.customerno=" & DbQuote(Str, mCustomer) & vbCrLf
    s = s & "   and i.location=" & DbQuote(Num, mLocation) & vbCrLf
    s = s & "   and i.seq=" & DbQuote(Num, mSeq) & vbCrLf
    Set rs = HFApp.SqlExec(s)
    
    If mCost = 0 Then
        txtMarkup.Text = 0
    Else
        txtMarkup.Text = 100 * (mSell - mCost) / mCost
    End If
    txtCost.Text = format(mCost, "#,##0.00")
    txtSell.Text = format(mSell, "#,##0.00")
    
    With gItems
        .Rows = 1
        While Not rs.EOF
        
            .AddItem ""
            i = .Rows - 1
            
            
            .TextMatrix(i, .ColIndex("Customer_No")) = mCustomer
            .TextMatrix(i, .ColIndex("Location")) = mLocation
            .TextMatrix(i, .ColIndex("OptionID")) = mOptionID
            .TextMatrix(i, .ColIndex("HFDescription")) = mAssemblyDesc
            
            .TextMatrix(i, .ColIndex("POIndex")) = "" & rs("POIndex")
            .TextMatrix(i, .ColIndex("RoundTo")) = "" & rs("RoundTo")
            .TextMatrix(i, .ColIndex("RoundDir")) = "" & rs("RoundDir")
            .TextMatrix(i, .ColIndex("ConversionFactor")) = "" & rs("ConversionFactor")
            .TextMatrix(i, .ColIndex("EstPhase")) = "" & rs("Phase")
            .TextMatrix(i, .ColIndex("EstItem")) = "" & rs("Item")
            .TextMatrix(i, .ColIndex("ItemDesc")) = "" & rs("Description")
            .TextMatrix(i, .ColIndex("ItemComments")) = "" & rs("Comments")
            .TextMatrix(i, .ColIndex("TakeoffQty")) = "" & rs("TakeoffQty")
            .TextMatrix(i, .ColIndex("TakeoffUOM")) = "" & rs("TakeoffUOM")
            .TextMatrix(i, .ColIndex("ConversionFactor")) = "" & rs("ConversionFactor")
            .TextMatrix(i, .ColIndex("OrderUOM")) = "" & rs("OrderUOM")
            .TextMatrix(i, .ColIndex("Job_No")) = mJob
            .TextMatrix(i, .ColIndex("JCExtra")) = mExtra
            .TextMatrix(i, .ColIndex("JCCostCode")) = "" & rs("JCCostCode")
            .TextMatrix(i, .ColIndex("JCCategory")) = "" & rs("JCCategory")
            .TextMatrix(i, .ColIndex("BudgetVendor")) = "" & rs("BudgetVendor")
            .TextMatrix(i, .ColIndex("BudgetQty")) = "" & rs("BudgetQty")
            .TextMatrix(i, .ColIndex("BudgetRate")) = "" & rs("BudgetRate")
            .TextMatrix(i, .ColIndex("BudgetPretax")) = "" & rs("BudgetPretax")
            .TextMatrix(i, .ColIndex("BudgetTaxGroup")) = "" & rs("BudgetTaxGroup")
            .TextMatrix(i, .ColIndex("BudgetJCTax")) = "" & rs("BudgetJCTax")
            .TextMatrix(i, .ColIndex("BudgetJCTaxRate")) = "" & rs("BudgetJCTaxRate")
            .TextMatrix(i, .ColIndex("BudgetNJCTax")) = "" & rs("BudgetNJCTax")
            .TextMatrix(i, .ColIndex("BudgetNJCTaxRate")) = "" & rs("BudgetNJCTaxRate")
            .TextMatrix(i, .ColIndex("seq")) = mSeq
            .TextMatrix(i, .ColIndex("Assembly")) = "" & rs("Assembly")
            .TextMatrix(i, .ColIndex("JCCostCodeDesc")) = "" & rs("JCCostCodeDesc")
            .TextMatrix(i, .ColIndex("JCCategoryDesc")) = "" & rs("JCCategoryDesc")
            .TextMatrix(i, .ColIndex("BudgetVendorName")) = "" & rs("BudgetVendorName")
            
            .TextMatrix(i, .ColIndex("BudgetTax")) = .ValueMatrix(i, .ColIndex("BudgetNJCTax")) + .ValueMatrix(i, .ColIndex("BudgetJCTax"))
            
            .TextMatrix(i, .ColIndex("RoomLocation")) = "" & rs("RoomLocation")
            For w = 1 To 40
            .TextMatrix(i, .ColIndex("WBS" & format(w, "00"))) = "" & rs("WBS" & format(w, "00"))
            Next
            
            rs.MoveNext
        Wend
    End With
    Call CalcTotals
    mDirty = False
End Sub

Private Sub LoadWBSDescriptions()
On Error GoTo eh
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
Exit Sub
eh: Call errHandler(SRCFILE & "LoadWBSDescriptions")
End Sub



Public Sub mnuEstimateItemsGridSub_Click(Index As Integer)
    Dim rc As Double
    Dim rc2 As Long
    Dim r As Long
    Dim c As Long
    Dim s As String
    Dim NewRow As Long
    Dim Phase As String
    Dim Item As String
    Dim Description As String
    
    With gItems
        If .Row <= 0 Then Exit Sub
        
        If Index = 3 Then
            Call gItems_KeyDown(vbKeyDelete, vbCtrlMask)
        End If
        

    End With
End Sub


















Private Sub CalcTotals(Optional CalcMargin As Boolean = True, Optional CalcMarkup As Boolean = True)
On Error Resume Next

    Dim r As Long
    
    mEditFlag = True
    
    If gItems.Rows = 1 Then
        mCost = Dollars(txtCost.Text)
    Else
        With gItems
            mCost = 0
            For r = 1 To .Rows - 1
                mCost = mCost + .ValueMatrix(r, .ColIndex("BudgetPretax")) + .ValueMatrix(r, .ColIndex("BudgetJCTax"))
            Next
        End With
        txtCost.Text = format(mCost, "#,##0.00")
    End If
    
    mSell = Dollars(txtSell.Text)
    
    If CalcMargin Then
        txtMargin.Text = ""
        txtMargin.Text = Round((mSell - mCost) / mSell * 100, 0)
    End If
    If CalcMarkup Then
        txtMarkup.Text = ""
        txtMarkup.Text = Round((mSell - mCost) / mCost * 100, 0)
    End If
    Call UpdatePrice
    mEditFlag = False
    
End Sub

Private Sub txtMargin_Change()
    If mEditFlag Then Exit Sub
    mEditFlag = True

    If Val(txtMargin.Text) >= 100 Then
        mSell = mCost
    Else
        mSell = 100 * mCost / (100 - Val(txtMargin.Text))
    End If
    txtSell.Text = format(mSell, "#,##0.00")
    Call CalcTotals(False, True)

    mDirty = True
    mEditFlag = False
End Sub
Private Sub txtMarkup_Change()
    If mEditFlag Then Exit Sub
    mEditFlag = True

    mSell = mCost * (1 + Val(txtMarkup.Text) / 100)
    txtSell.Text = format(mSell, "#,##0.00")
    Call CalcTotals(True, False)

    mDirty = True
    mEditFlag = False
End Sub


Private Sub txtSell_Change()
If mEditFlag Then Exit Sub
    mEditFlag = True
    
    Call CalcTotals

    mDirty = True
    mEditFlag = False
End Sub

Private Sub txtCost_Change()
If mEditFlag Then Exit Sub
    mEditFlag = True
    
    Call CalcTotals

    mDirty = True
    mEditFlag = False
End Sub

Private Function Dollars(s As String) As Double
    Dollars = Val(Replace(Replace(Trim(s), ",", ""), "$", ""))
End Function
