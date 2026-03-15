VERSION 5.00
Begin VB.Form FTakeoffOneTime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "One Time Takeoff"
   ClientHeight    =   2715
   ClientLeft      =   7410
   ClientTop       =   2925
   ClientWidth     =   5640
   Icon            =   "FTakeoffOneTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCostCode 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1785
      Locked          =   -1  'True
      MaxLength       =   200
      TabIndex        =   5
      Top             =   1395
      Width           =   3495
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Height          =   375
      Index           =   0
      Left            =   3000
      TabIndex        =   7
      Top             =   2220
      Width           =   1215
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   4320
      TabIndex        =   8
      Top             =   2220
      Width           =   1215
   End
   Begin HFEst.VBCombo cboUOM 
      Height          =   240
      Left            =   2475
      TabIndex        =   2
      Top             =   510
      Width           =   975
      _ExtentX        =   979
      _ExtentY        =   423
   End
   Begin VB.TextBox txtPrice 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1800
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   765
      Width           =   915
   End
   Begin VB.TextBox txtQty 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1785
      TabIndex        =   1
      Text            =   "0"
      Top             =   510
      Width           =   675
   End
   Begin VB.TextBox txtDescription 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1785
      MaxLength       =   200
      TabIndex        =   0
      Top             =   255
      Width           =   3495
   End
   Begin HFEst.VBCombo cboPOIndex 
      Height          =   240
      Left            =   1785
      TabIndex        =   4
      Top             =   1140
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   423
      Style           =   2
   End
   Begin HFEst.VBCombo cboCategory 
      Height          =   240
      Left            =   1785
      TabIndex        =   6
      Top             =   1650
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   423
      Style           =   2
   End
   Begin VB.Image cmdBrowse 
      Height          =   240
      Index           =   0
      Left            =   5295
      Picture         =   "FTakeoffOneTime.frx":000C
      Top             =   1395
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Price"
      Height          =   195
      Index           =   5
      Left            =   1365
      TabIndex        =   14
      Top             =   810
      Width           =   360
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Category"
      Height          =   195
      Index           =   1
      Left            =   1095
      TabIndex        =   13
      Top             =   1680
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "FTakeoffOneTime.frx":0156
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Cost Code"
      Height          =   195
      Index           =   3
      Left            =   990
      TabIndex        =   12
      Top             =   1425
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Qty/Unit"
      Height          =   195
      Index           =   2
      Left            =   1125
      TabIndex        =   11
      Top             =   555
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "PO Index"
      Height          =   195
      Index           =   0
      Left            =   1065
      TabIndex        =   10
      Top             =   1185
      Width           =   660
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Description"
      Height          =   195
      Index           =   16
      Left            =   930
      TabIndex        =   9
      Top             =   300
      Width           =   795
   End
End
Attribute VB_Name = "FTakeoffOneTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SRCFILE = "FTakeoffOneTime::"

Private mParentForm          As Object
Private mAssemblyDescription As String
Private mCommunity           As String
Private mCommunityPhase      As String
Private mAssembly            As String
Private mModel               As String
Private mJob                 As String

Public Sub Takeoff(ParentForm As Form, _
                   AssemblyDescription As String, _
                   community As String, _
                   CommunityPhase As String, _
                   Model As String, _
                   Assembly As String, _
                   Job As String)
                    

    Set mParentForm = ParentForm
    mAssemblyDescription = AssemblyDescription
    mCommunity = community
    mCommunityPhase = CommunityPhase
    mModel = Model
    mAssembly = Assembly
    mJob = Job
    Me.Caption = "One Time Takeoff - " & AssemblyDescription
    Me.Show vbModal
    
End Sub


Private Sub cboPOIndex_Click()

    Dim s As String
    Dim rs As Recordset
    
    s = ""
    s = s & "select p.jccostcode,c.description,p.jccategory " & vbCrLf
    s = s & "from tblpoindex p" & vbCrLf
    s = s & "left outer join standardcostcodes c on p.jccostcode=c.costcode and p.divisionid=c.divisionid" & vbCrLf
    s = s & "where p.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
    s = s & "and p.poindex=" & DbQuote(Str, cboPOIndex.Text) & vbCrLf
    Set rs = HFApp.SqlExec(s)
    If Not rs.EOF Then
        txtCostCode.Tag = "" & rs("jccostcode")
        txtCostCode.Text = "" & rs("description")
        Call SetComboBoxListIndex(cboCategory, , "" & rs("jccategory"))
    End If


End Sub

Private Sub cmdBrowse_Click(Index As Integer)
    Dim s As String
    s = "select CostCode,Description from standardcostcodes where divisionid=" & DbQuote(Num, HFApp.DivisionID)
    If FPickList.Choose(HFApp.Databases(dbHomefront), "Cost Codes", s, txtCostCode.Tag) Then
        txtCostCode.Tag = FPickList.SelectedItem("Costcode")
        txtCostCode.Text = FPickList.SelectedItem("description")
    End If
    
End Sub

Private Sub cmdNav_Click(Index As Integer)
    If Index = 0 Then
        SaveData
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Call IniGetForm(Me)
    Call LoadComboBox(cboUOM, HFApp.Databases(dbHomefront), "SELECT DISTINCT OrderUOM Unit,'',0 FROM tblPhaseItem where DivisionID = " & HFApp.DivisionID & " order by 1")
    Call LoadComboBox(cboPOIndex, HFApp.Databases(dbHomefront), "select case when poindex=description then poindex else POIndex + isnull(' '+description,'') end,poindex,0 from tblPOIndex where DivisionID=" & HFApp.DivisionID & " order by 1")
    If HFApp.Options(AccountingSystem) = asQuickBooks Then
        Call LoadComboBox(cboCategory, HFApp.Databases(dbHomefront), "select Description,Category,0 from standardcategories where DivisionID = " & HFApp.DivisionID & " order by 1")
    Else
        Call LoadComboBox(cboCategory, HFApp.Databases(dbHomefront), "select Category + ISNULL(' - ' + Description,''),Category,0 from standardcategories where DivisionID = " & HFApp.DivisionID & " order by 1")
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub


Private Sub txtDescription_GotFocus()
    SelectAll txtDescription
End Sub
Private Sub txtQty_GotFocus()
    SelectAll txtQty
End Sub
Private Sub txtPrice_GotFocus()
    SelectAll txtPrice
End Sub


Private Sub txtQty_Validate(Cancel As Boolean)
    txtQty.Text = Val(txtQty.Text)
End Sub
Private Sub txtPrice_Validate(Cancel As Boolean)
    txtPrice.Text = format(Val(txtPrice.Text), "0.00###############")
End Sub



Private Sub SaveData()
    Dim s As String
    Dim rs As Recordset
    
    Dim Description As String
    Dim OrderQty As Double
    Dim OrderUOM As String
    Dim TakeoffQty As Double
    Dim TakeoffUOM As String
    Dim ConversionFactor As Double
    Dim JCCostCode As String
    Dim JCCostCodeDesc As String
    Dim JCCategory As String
    Dim JCCategoryDesc As String
    Dim Vendor As String
    Dim VendorName As String
    Dim price As Double
    Dim TaxGroup As String
    Dim TaxGroupName As String
    Dim JCTaxRate As Double
    Dim NJCTaxRate As Double
    Dim POIndex As String
    Dim Comments As String
    
    'validate
    If Trim(txtDescription.Text) = "" Then
        MsgBox "Description is required.", vbExclamation, App.ProductName
        txtDescription.SetFocus
        Exit Sub
    End If
    
    
    Description = txtDescription.Text
    OrderQty = Val(txtQty.Text)
    OrderUOM = cboUOM.Text
    TakeoffQty = Val(txtQty.Text)
    TakeoffUOM = cboUOM.Text
    ConversionFactor = 1
    JCCostCode = txtCostCode.Tag
    JCCategory = GetComboBoxListKey(cboCategory)
    
    POIndex = GetComboBoxListKey(cboPOIndex)
    
    
    Comments = ""
    price = Val(txtPrice.Text)

    On Error Resume Next
    JCCostCodeDesc = HFApp.SqlExec("SELECT Description FROM StandardCostCodes WHERE DivisionID = " & HFApp.DivisionID & " and CostCode=" & DbQuote(Str, JCCostCode))(0)
    JCCategoryDesc = HFApp.SqlExec("SELECT Description FROM StandardCategories WHERE DivisionID = " & HFApp.DivisionID & " and Category=" & DbQuote(Str, JCCategory))(0)
    On Error GoTo 0




    'get vendor
    s = ""
    s = s & "SELECT Vendor_ID,Vendor_Name" & vbCrLf
    s = s & "  FROM tblVendors" & vbCrLf
    s = s & " WHERE DivisionID = " & HFApp.DivisionID & " and Vendor_id=dbo.Purch_GetCommunityVendor(" & DbQuote(Str, mCommunity) & "," & DbQuote(Str, POIndex) & "," & HFApp.DivisionID & ")" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    If Not rs.EOF Then
        Vendor = "" & rs(0)
        VendorName = "" & rs(1)
    End If


    'get tax rates
    s = ""
    s = s & "SELECT *"
    s = s & " FROM TaxGroups "
    s = s & " WHERE DivisionID = " & HFApp.DivisionID & " and TaxGroup=dbo.Purch_GetDefaultTaxGroup("
    s = s & DbQuote(Str, mJob) & ","
    s = s & DbQuote(Str, mCommunity) & ","
    s = s & DbQuote(Str, mCommunityPhase) & ","
    s = s & DbQuote(Str, mModel) & ","
    s = s & DbQuote(Str, mAssembly) & ","
    s = s & "'','',"
    s = s & DbQuote(Str, Vendor) & ","
    s = s & DbQuote(Str, JCCategory) & "," & HFApp.DivisionID & ")"
    Set rs = HFApp.SqlExec(s)
    If Not rs.EOF Then
        TaxGroup = "" & rs("TaxGroup")
        TaxGroupName = "" & rs("Description")
        JCTaxRate = "" & rs("JCRate")
        NJCTaxRate = "" & rs("NJCRate")
    End If


    Dim WBS(40) As String
    Call mParentForm.AddItem("", "", "", "", "", 0, Description, OrderQty, OrderUOM, TakeoffQty, TakeoffUOM, ConversionFactor, 0, 0, 0, "", JCCostCode, JCCostCodeDesc, JCCategory, JCCategoryDesc, Vendor, VendorName, price, TaxGroup, TaxGroupName, JCTaxRate, NJCTaxRate, POIndex, Comments, "", 1, "", WBS)
    txtDescription.SetFocus
    
End Sub

