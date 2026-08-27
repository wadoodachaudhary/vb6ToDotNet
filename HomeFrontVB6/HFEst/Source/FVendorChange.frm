VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form FVendorChange 
   Caption         =   "Work Reassignment Wizard"
   ClientHeight    =   6165
   ClientLeft      =   840
   ClientTop       =   1230
   ClientWidth     =   7305
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FVendorChange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   7305
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Height          =   4755
      Index           =   2
      Left            =   7140
      TabIndex        =   7
      Top             =   3330
      Visible         =   0   'False
      Width           =   7395
      Begin VSFlex8Ctl.VSFlexGrid gData 
         Height          =   3555
         Left            =   930
         TabIndex        =   11
         Top             =   930
         Width           =   6165
         _cx             =   1982802554
         _cy             =   1982797951
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FVendorChange.frx":000C
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   1
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   7
         PicturesOver    =   0   'False
         FillStyle       =   0
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
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "Change the assigned vendor. You can sort the rows and select and update multiple PO's at a time."
         Height          =   435
         Index           =   2
         Left            =   1050
         TabIndex        =   15
         Top             =   420
         UseMnemonic     =   0   'False
         Width           =   5955
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reassign Work"
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
         Index           =   5
         Left            =   900
         TabIndex        =   12
         Top             =   150
         UseMnemonic     =   0   'False
         Width           =   1305
      End
      Begin VB.Image imgOK 
         Height          =   480
         Index           =   4
         Left            =   300
         OLEDropMode     =   1  'Manual
         Picture         =   "FVendorChange.frx":00E1
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Height          =   4755
      Index           =   1
      Left            =   4230
      TabIndex        =   17
      Top             =   2400
      Visible         =   0   'False
      Width           =   7395
      Begin VB.CheckBox chkVendor 
         Caption         =   "Vendor"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1170
         TabIndex        =   24
         Top             =   1110
         Value           =   1  'Checked
         Width           =   3435
      End
      Begin VB.CheckBox chkJob 
         Caption         =   "Job"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1170
         TabIndex        =   23
         Top             =   1350
         Value           =   1  'Checked
         Width           =   3435
      End
      Begin VB.CheckBox chkCategory 
         Caption         =   "Category"
         Height          =   195
         Left            =   1170
         TabIndex        =   22
         Top             =   2070
         Width           =   3435
      End
      Begin VB.CheckBox chkCostcode 
         Caption         =   "Cost Code"
         Height          =   195
         Left            =   1170
         TabIndex        =   21
         Top             =   1830
         Width           =   3435
      End
      Begin VB.CheckBox chkPOIndex 
         Caption         =   "PO Index"
         Height          =   195
         Left            =   1170
         TabIndex        =   20
         Top             =   1590
         Width           =   3435
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grouped by"
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
         Index           =   7
         Left            =   1050
         TabIndex        =   25
         Top             =   810
         UseMnemonic     =   0   'False
         Width           =   990
      End
      Begin VB.Image imgOK 
         Height          =   480
         Index           =   1
         Left            =   300
         OLEDropMode     =   1  'Manual
         Picture         =   "FVendorChange.frx":09AB
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Work Filter"
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
         Left            =   900
         TabIndex        =   19
         Top             =   150
         UseMnemonic     =   0   'False
         Width           =   945
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "How do you want to view the incomplete work?"
         Height          =   435
         Index           =   3
         Left            =   1050
         TabIndex        =   18
         Top             =   420
         UseMnemonic     =   0   'False
         Width           =   5955
      End
   End
   Begin HFEst.WizHead WizHead1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   1588
      Caption         =   "Change Vendors"
      Description     =   "The work reassignment wizard will help you to quickly reassign incomplete work."
      Icon            =   "FVendorChange.frx":1275
   End
   Begin VB.PictureBox WizFoot 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   7305
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5580
      Width           =   7305
      Begin VB.CommandButton cmdNav 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Index           =   0
         Left            =   1860
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "< &Back"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   3060
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Next >"
         Height          =   375
         Index           =   2
         Left            =   4200
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "Commi&t"
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   5400
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   -60
         X2              =   26420
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   26480
         Y1              =   15
         Y2              =   15
      End
   End
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Height          =   4755
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   870
      Visible         =   0   'False
      Width           =   7395
      Begin VB.TextBox txtWhereClause 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   1920
         TabIndex        =   16
         Top             =   3480
         Visible         =   0   'False
         Width           =   3915
      End
      Begin VB.TextBox txtVendors 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   1755
         Left            =   1920
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   1050
         Width           =   3915
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "Click Next to continue."
         Height          =   525
         Index           =   6
         Left            =   1050
         TabIndex        =   14
         Top             =   3150
         UseMnemonic     =   0   'False
         Width           =   5865
      End
      Begin VB.Image cmdChooseItem 
         Height          =   240
         Index           =   5
         Left            =   5850
         Picture         =   "FVendorChange.frx":1B4F
         Top             =   1050
         Width           =   240
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "Choose one or more vendors who have work that needs to be reassigned."
         Height          =   435
         Index           =   1
         Left            =   1050
         TabIndex        =   10
         Top             =   420
         UseMnemonic     =   0   'False
         Width           =   5865
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vendor(s)"
         Height          =   195
         Index           =   0
         Left            =   1140
         TabIndex        =   9
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Choose Vendor"
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
         Left            =   900
         TabIndex        =   8
         Top             =   150
         UseMnemonic     =   0   'False
         Width           =   1305
      End
      Begin VB.Image imgOK 
         Height          =   480
         Index           =   0
         Left            =   300
         Picture         =   "FVendorChange.frx":1C99
         Top             =   240
         Width           =   480
      End
   End
End
Attribute VB_Name = "FVendorChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FVendorChange::"

Private mRecordset As Recordset

Public Function ShowForm()
    Me.Show vbModal
End Function


Private Sub Form_Load()
    Call IniGetForm(Me)
    chkPOIndex.value = IIf(IniGet(AppIni, SRCFILE, "GroupByPOIndex") = "True", vbChecked, vbUnchecked)
    chkCostcode.value = IIf(IniGet(AppIni, SRCFILE, "GroupByCostCode") = "True", vbChecked, vbUnchecked)
    chkCategory.value = IIf(IniGet(AppIni, SRCFILE, "GroupByCategory") = "True", vbChecked, vbUnchecked)
    CurrentFrame = 0
End Sub


Private Property Get CurrentFrame() As Long
    Dim i As Long
    For i = 0 To WizFrame.UBound
        If WizFrame(i).Visible = True Then
            CurrentFrame = i
            Exit Property
        End If
    Next
    CurrentFrame = WizFrame.LBound
End Property

Private Property Let CurrentFrame(RHS As Long)
On Error GoTo eh
    Static AsmblyLoaded As Boolean
    Static ItemsLoaded  As Boolean
    Dim i As Long
    Dim s As String
    Dim b As Boolean
    
    Screen.MousePointer = vbHourglass
    For i = 0 To WizFrame.UBound
        WizFrame(i).Visible = i = RHS
    Next
    cmdNav(1).Enabled = RHS > 0
    cmdNav(2).Enabled = RHS < WizFrame.UBound
    cmdNav(3).Enabled = RHS = WizFrame.UBound
        
        
    Select Case RHS
        Case 0 'choose vendors
            cmdNav(2).Enabled = txtWhereClause.Text <> ""
            
        Case 1 'filter
            
        Case 2 'reassign work
            s = ""
            s = s & "select distinct budgetvendor CurrentVendor,budgetvendor Vendor,budgetvendorname CompanyName,Community,CommunityDesc,Job_No Job,JobDesc" & vbCrLf
            If chkPOIndex.value = vbChecked Then s = s & "      ,POIndex" & vbCrLf
            If chkCostcode.value = vbChecked Then s = s & "      ,JCCostCode CostCode,JCCostCodeDesc CostCodeDesc" & vbCrLf
            If chkCategory.value = vbChecked Then s = s & "      ,JCCategory Category,JCCategoryDesc CategoryDesc" & vbCrLf
            s = s & "from estimateditems" & vbCrLf
            s = s & "where budgetgenerated=0 and DivisionID = " & HFApp.DivisionID & " and budgetvendor in(" & txtWhereClause & ")" & vbCrLf
            s = s & "union" & vbCrLf
            s = s & "select distinct povendor CurrentVendor,povendor,povendorname,Community,CommunityDesc,Job_No,JobDesc" & vbCrLf
            If chkPOIndex.value = vbChecked Then s = s & "      ,POIndex" & vbCrLf
            If chkCostcode.value = vbChecked Then s = s & "      ,JCCostCode,JCCostCodeDesc" & vbCrLf
            If chkCategory.value = vbChecked Then s = s & "      ,JCCategory,JCCategoryDesc" & vbCrLf
            s = s & "from estimateditems" & vbCrLf
            s = s & "where pogenbatch=0 and DivisionID = " & HFApp.DivisionID & " and povendor in(" & txtWhereClause & ")" & vbCrLf
            s = s & "order by 1,2,3,4,5,6" & vbCrLf
            Set mRecordset = New Recordset
            Call mRecordset.Open(s, HFApp.Databases(dbHomefront))
            
            With gData
                Set .DataSource = mRecordset
                If .ColKey(0) = "" Then
                    On Error Resume Next
                    For i = 0 To .Cols - 1
                        .ColKey(i) = .TextMatrix(0, i)
                    Next
                    On Error GoTo eh
                End If
                Call .AutoSize(0, .Cols - 1)
                
                b = IsIn(HFApp.Options(AccountingSystem), asQuickBooks, asSimply)
                .ColHidden(.ColIndex("vendor")) = b
                .ColHidden(.ColIndex("CompanyName")) = False
                .ColHidden(.ColIndex("currentvendor")) = True
                If .ColIndex("costcode") <> -1 Then .ColHidden(.ColIndex("costcode")) = b
                If .ColIndex("category") <> -1 Then .ColHidden(.ColIndex("category")) = b
                
            End With
            
    End Select
    
    Screen.MousePointer = vbDefault
    
Exit Property
eh: Call errHandler(SRCFILE & "CurrentFrame")
End Property




Private Sub cmdChooseItem_Click(Index As Integer)
    Dim s As String
    Dim i As Long
    
    s = ""
    s = s & "select distinct v.vendor_name CompanyName,v.vendor_id Vendor" & vbCrLf
    s = s & "from estimateitems e join tblvendors v on(e.DivisionID = v.DivisionID and e.budgetvendor=v.vendor_id)" & vbCrLf
    s = s & "where budgetgenerated=0 and v.DivisionID = " & HFApp.DivisionID & vbCrLf
    s = s & "union" & vbCrLf
    s = s & "select distinct v.vendor_name,v.vendor_id" & vbCrLf
    s = s & "from estimateitems e join tblvendors v on(v.DivisionID = e.DivisionID and e.povendor=v.vendor_id)" & vbCrLf
    s = s & "where pogenbatch=0" & vbCrLf
    s = s & " and v.DivisionID = " & HFApp.DivisionID
    If FPickList.Choose(HFApp.Databases(dbHomefront), "Vendor", s, , , , , "vendor", True) Then
        txtVendors.Text = ""
        txtWhereClause.Text = ""
        For i = 1 To FPickList.SelectedItems
            txtVendors.Text = txtVendors.Text & vbCrLf & FPickList.SelectedItem("CompanyName", i)
            txtWhereClause.Text = txtWhereClause.Text & ", " & DbQuote(Str, FPickList.SelectedItem("Vendor", i))
        Next
        txtVendors.Text = Mid(txtVendors.Text, 3)
        txtWhereClause.Text = Mid(txtWhereClause.Text, 3)
        cmdNav(2).Enabled = Me.txtWhereClause.Text <> ""
    End If
    
End Sub

Private Sub cmdNav_Click(Index As Integer)
On Error Resume Next
    Select Case Index
        Case 0: Unload Me
        Case 1: CurrentFrame = CurrentFrame - 1
        Case 2: CurrentFrame = CurrentFrame + 1
        Case 3: Call SaveData
    End Select
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Const margin = 60
    Dim i As Long
    For i = 0 To WizFrame.UBound
        WizFrame(i).BorderStyle = 0
        WizFrame(i).Move 0, WizHead1.Height, Me.ScaleWidth, Me.ScaleHeight - WizHead1.Height - WizFoot.Height
    Next
    
    gData.Move gData.Left, gData.Top, WizFrame(0).Width - gData.Left - 180, WizFrame(0).Height - gData.Top - 180
    
    
    cmdNav(0).Move Me.ScaleWidth - (4 * (margin + cmdNav(0).Width)), 2 * margin
    cmdNav(1).Move Me.ScaleWidth - (3 * (margin + cmdNav(0).Width)), 2 * margin
    cmdNav(2).Move Me.ScaleWidth - (2 * (margin + cmdNav(0).Width)), 2 * margin
    cmdNav(3).Move Me.ScaleWidth - (1 * (margin + cmdNav(0).Width)), 2 * margin
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call IniPut(AppIni, SRCFILE, "GroupByPOIndex", chkPOIndex.value = vbChecked)
    Call IniPut(AppIni, SRCFILE, "GroupByCostCode", chkCostcode.value = vbChecked)
    Call IniPut(AppIni, SRCFILE, "GroupByCategory", chkCategory.value = vbChecked)
    Call IniPutForm(Me)
End Sub



Private Sub SaveData()
On Error GoTo eh
    Dim s As String
    Dim i As Long
    Dim ids As String
    Dim rs As Recordset
    
    With gData
    For i = 1 To .Rows - 1
        If .Cell(flexcpData, i, 0) = "DIRTY" Then
            
            ids = ""
            .Cell(flexcpData, i, 0) = ""
            
            
            'BUDGET ITEMS FIRST--------------------------------------------------------------------------------------
            'get list of budget items to be used in update and later in refresh costs
            s = ""
            s = s & "select estitemid from estimateitems" & vbCrLf
            s = s & " where budgetgenerated=0" & vbCrLf
            s = s & "   and DivisionID = " & HFApp.DivisionID & " and budgetvendor=" & DbQuote(Str, .TextMatrix(i, .ColIndex("currentvendor"))) & vbCrLf
            s = s & "   and job=" & DbQuote(Str, .TextMatrix(i, .ColIndex("job"))) & vbCrLf
            If chkPOIndex.value = vbChecked Then s = s & "   and poindex=" & DbQuote(Str, .TextMatrix(i, .ColIndex("poindex"))) & vbCrLf
            If chkCostcode.value = vbChecked Then s = s & "   and jccostcode=" & DbQuote(Str, .TextMatrix(i, .ColIndex("costcode"))) & vbCrLf
            If chkCategory.value = vbChecked Then s = s & "   and jccategory=" & DbQuote(Str, .TextMatrix(i, .ColIndex("category"))) & vbCrLf
            Set rs = HFApp.SqlExec(s, dbHomefront)
            While Not rs.EOF
                ids = ids & "," & rs(0)
                rs.MoveNext
            Wend
            ids = Mid(ids, 2)
            
            If ids <> "" Then
                'change vendors
                s = ""
                s = s & "UPDATE EstimateItems" & vbCrLf
                s = s & "   SET budgetvendor=" & DbQuote(Str, .TextMatrix(i, .ColIndex("vendor"))) & vbCrLf
                s = s & "      ,BudgetOverridden=0" & vbCrLf
                s = s & "      ,BudgetRate= case when ISNULL(dbo.Purch_GetItemRate(0,0,v.Community,v.CommunityPhase,v.Assembly, case when v.assemblytype=0 or v.assemblytype=2 then v.model else '' end,'',v.EstPhase,v.EstItem,0," & DbQuote(Str, .TextMatrix(i, .ColIndex("vendor"))) & ",getdate()," & HFApp.DivisionID & "),0) <>0 " & vbCrLf
                s = s & "                        then ISNULL(dbo.Purch_GetItemRate(0,0,v.Community,v.CommunityPhase,v.Assembly, case when v.assemblytype=0 or v.assemblytype=2 then v.model else '' end,'',v.EstPhase,v.EstItem,0," & DbQuote(Str, .TextMatrix(i, .ColIndex("vendor"))) & ",getdate()," & HFApp.DivisionID & "),0)" & vbCrLf
                s = s & "                        when 'True'=" & DbQuote(Str, "" & HFApp.Options(ZeroRateOnRefreshCosts)) & " and i.BudgetOverridden=0 then 0" & vbCrLf
                s = s & "                        else i.BudgetRate end" & vbCrLf
                s = s & "  FROM EstimateItems i JOIN EstimatedItems v ON(i.EstItemID=v.EstItemID)" & vbCrLf
                s = s & " WHERE i.BudgetDeleted=0" & vbCrLf
                s = s & "   AND i.BudgetGenerated=0" & vbCrLf
                s = s & "   AND i.estitemid in(" & ids & ")" & vbCrLf
                s = s & " AND i.DivisionID = " & HFApp.DivisionID
                Call HFApp.SqlExec(s, dbHomefront)
                
                'extend prices and taxes
                s = ""
                s = s & "UPDATE EstimateItems" & vbCrLf
                s = s & "   SET BudgetPretax = i.BudgetRate * i.BudgetQty" & vbCrLf
                s = s & "      ,BudgetJCTax = i.BudgetRate * i.BudgetQty * i.BudgetJCTaxRate/100" & vbCrLf
                s = s & "      ,BudgetNJCTax = i.BudgetRate * i.BudgetQty * i.BudgetNJCTaxRate/100" & vbCrLf
                s = s & "  FROM EstimateItems i JOIN EstimatedItems v ON(i.EstItemID=v.EstItemID)" & vbCrLf
                s = s & " WHERE v.BudgetDeleted=0" & vbCrLf
                s = s & "   AND v.BudgetGenerated=0" & vbCrLf
                s = s & " AND i.DivisionID = " & HFApp.DivisionID & vbCrLf
                s = s & "   AND i.estitemid in(" & ids & ")" & vbCrLf
                Call HFApp.SqlExec(s, dbHomefront)
            End If
                        
            'PO ITEMS SECOND--------------------------------------------------------------------------------------
            'get list of po item ids
            s = ""
            s = s & "select estitemid from estimateitems" & vbCrLf
            s = s & " where pogenbatch=0" & vbCrLf
            s = s & "   and povendor=" & DbQuote(Str, .TextMatrix(i, .ColIndex("currentvendor"))) & vbCrLf
            s = s & "   and job=" & DbQuote(Str, .TextMatrix(i, .ColIndex("job"))) & vbCrLf
            s = s & " AND DivisionID = " & HFApp.DivisionID & vbCrLf
            If chkPOIndex.value = vbChecked Then s = s & "   and poindex=" & DbQuote(Str, .TextMatrix(i, .ColIndex("poindex"))) & vbCrLf
            If chkCostcode.value = vbChecked Then s = s & "   and jccostcode=" & DbQuote(Str, .TextMatrix(i, .ColIndex("costcode"))) & vbCrLf
            If chkCategory.value = vbChecked Then s = s & "   and jccategory=" & DbQuote(Str, .TextMatrix(i, .ColIndex("category"))) & vbCrLf
            Set rs = HFApp.SqlExec(s, dbHomefront)
            While Not rs.EOF
                ids = ids & "," & rs(0)
                rs.MoveNext
            Wend
            ids = Mid(ids, 2)
            
            If ids <> "" Then
                'change vendor and refresh costs
                s = ""
                s = s & "UPDATE EstimateItems" & vbCrLf
                s = s & "   set povendor=" & DbQuote(Str, .TextMatrix(i, .ColIndex("vendor"))) & vbCrLf
                s = s & "      ,POOverridden=0" & vbCrLf
                s = s & "      ,PORate= case when ISNULL(dbo.Purch_GetItemRate(0,0,v.Community,v.CommunityPhase,v.Assembly, case when v.assemblytype=0 or v.assemblytype=2 then v.model else '' end,'',v.EstPhase,v.EstItem,0," & DbQuote(Str, .TextMatrix(i, .ColIndex("vendor"))) & ",getdate()," & HFApp.DivisionID & "),0) <>0 " & vbCrLf
                s = s & "                    then ISNULL(dbo.Purch_GetItemRate(0,0,v.Community,v.CommunityPhase,v.Assembly, case when v.assemblytype=0 or v.assemblytype=2 then v.model else '' end,'',v.EstPhase,v.EstItem,0," & DbQuote(Str, .TextMatrix(i, .ColIndex("vendor"))) & ",getdate()," & HFApp.DivisionID & "),0)" & vbCrLf
                s = s & "                    when 'True'=" & DbQuote(Str, "" & HFApp.Options(ZeroRateOnRefreshCosts)) & " and i.POOverridden=0 then 0" & vbCrLf
                s = s & "                    else i.PORate end" & vbCrLf
                s = s & "  FROM EstimateItems i JOIN EstimatedItems v ON(i.EstItemID=v.EstItemID)" & vbCrLf
                s = s & " WHERE i.PODeleted=0" & vbCrLf
                s = s & "   AND i.POGenBatch=0" & vbCrLf
                s = s & " AND i.DivisionID = " & HFApp.DivisionID
                s = s & "   AND i.estitemid in(" & ids & ")" & vbCrLf
                Call HFApp.SqlExec(s, dbHomefront, i)
                
                'extend prices and taxes
                s = ""
                s = s & "UPDATE EstimateItems" & vbCrLf
                s = s & "   SET POPretax = i.PORate * i.POQty" & vbCrLf
                s = s & "      ,POJCTax = i.PORate * i.POQty * i.POJCTaxRate/100" & vbCrLf
                s = s & "      ,PONJCTax = i.PORate * i.POQty * i.PONJCTaxRate/100" & vbCrLf
                s = s & "  FROM EstimateItems i JOIN EstimatedItems v ON(i.EstItemID=v.EstItemID)" & vbCrLf
                s = s & " WHERE i.PODeleted=0" & vbCrLf
                s = s & "   AND i.POGenBatch=0" & vbCrLf
                s = s & " AND i.DivisionID = " & HFApp.DivisionID
                s = s & "   AND i.estitemid in(" & ids & ")" & vbCrLf
                Call HFApp.SqlExec(s, dbHomefront)
            End If
            
        End If
    Next
    End With
    
    
    CurrentFrame = CurrentFrame
Exit Sub
eh: Call errHandler(SRCFILE & "SaveData", s)
End Sub


Private Sub gData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gData
        Select Case .ColKey(Col)
            Case "Vendor", "CompanyName"
                .ComboList = "..."
            Case Else
                Cancel = True
        End Select
    End With
End Sub

Private Sub gData_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    With gData
    Select Case .ColKey(Col)
        Case "Vendor", "CompanyName"
            If FPickList.Choose(HFApp.Databases(dbHomefront), "Vendor", "select Vendor_name Company,vendor_id Vendor from tblvendors where isnull(inactive,0)=0 and divisionid=" & DbQuote(Num, HFApp.DivisionID)) Then
                .Cell(flexcpText, Min(.Row, .RowSel), .ColIndex("Vendor"), Max(.Row, .RowSel), .ColIndex("Vendor")) = FPickList.SelectedItem("vendor")
                .Cell(flexcpText, Min(.Row, .RowSel), .ColIndex("CompanyName"), Max(.Row, .RowSel), .ColIndex("CompanyName")) = FPickList.SelectedItem("company")
                .Cell(flexcpData, Min(.Row, .RowSel), 0, Max(.Row, .RowSel), 0) = "DIRTY"
            End If
    End Select
    End With
End Sub

