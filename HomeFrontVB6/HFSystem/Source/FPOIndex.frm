VERSION 5.00
Begin VB.Form FPOIndex 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Purchase Order"
   ClientHeight    =   7575
   ClientLeft      =   3855
   ClientTop       =   2745
   ClientWidth     =   8670
   Icon            =   "FPOIndex.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkLiens 
      Alignment       =   1  'Right Justify
      Caption         =   "Require Lien Release"
      Height          =   255
      Left            =   6060
      TabIndex        =   10
      Top             =   1875
      Width           =   1920
   End
   Begin VB.CheckBox chkMPO 
      Alignment       =   1  'Right Justify
      Caption         =   "Measurement PO"
      Height          =   255
      Left            =   6060
      TabIndex        =   11
      Top             =   2130
      Width           =   1920
   End
   Begin VB.CheckBox chkBuildPro 
      Alignment       =   1  'Right Justify
      Caption         =   "BuildPro"
      Height          =   255
      Left            =   6060
      TabIndex        =   12
      Top             =   2385
      Width           =   1920
   End
   Begin VB.CheckBox chkRequiresPaymentApproval 
      Alignment       =   1  'Right Justify
      Caption         =   "Approval Rqrd"
      Height          =   255
      Left            =   6060
      TabIndex        =   13
      Top             =   2640
      Width           =   1920
   End
   Begin HFSystem.WizHead WizHead1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   1588
      Caption         =   "PO Index Setup"
      Description     =   "Define your standard purchase orders."
      Icon            =   "FPOIndex.frx":058A
   End
   Begin VB.PictureBox picAttachments 
      BorderStyle     =   0  'None
      Height          =   2520
      Left            =   6060
      ScaleHeight     =   2520
      ScaleWidth      =   2295
      TabIndex        =   33
      Top             =   4485
      Width           =   2295
      Begin VB.ListBox lstAttachments 
         Appearance      =   0  'Flat
         Height          =   2550
         IntegralHeight  =   0   'False
         Left            =   -15
         Style           =   1  'Checkbox
         TabIndex        =   18
         Top             =   -15
         Width           =   2325
      End
   End
   Begin VB.CheckBox chkTotalOnly 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Only"
      Height          =   255
      Left            =   6060
      TabIndex        =   16
      Top             =   3390
      Width           =   1920
   End
   Begin VB.TextBox txtRetainagePercent 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   230
      Left            =   7605
      MaxLength       =   50
      TabIndex        =   17
      Text            =   "99%"
      Top             =   3690
      Width           =   375
   End
   Begin HFSystem.VBCombo cboJCCostCode 
      Height          =   240
      Left            =   1560
      TabIndex        =   7
      Top             =   6480
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   423
   End
   Begin VB.TextBox txtNotes 
      BorderStyle     =   0  'None
      Height          =   1785
      Left            =   1560
      MaxLength       =   2000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   4680
      Width           =   4155
   End
   Begin VB.TextBox txtPOIndex 
      BorderStyle     =   0  'None
      Height          =   230
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1080
      Width           =   3195
   End
   Begin VB.TextBox txtStandardText 
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   1560
      MaxLength       =   2000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1860
      Width           =   4155
   End
   Begin VB.TextBox txtShipVia 
      BorderStyle     =   0  'None
      Height          =   230
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   3
      Top             =   2400
      Width           =   3195
   End
   Begin VB.TextBox txtTerms 
      BorderStyle     =   0  'None
      Height          =   1785
      Left            =   1560
      MaxLength       =   2000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2880
      Width           =   4155
   End
   Begin VB.TextBox txtFOB 
      BorderStyle     =   0  'None
      Height          =   230
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   4
      Top             =   2640
      Width           =   3195
   End
   Begin VB.CheckBox chkHideQty 
      Alignment       =   1  'Right Justify
      Caption         =   "Hide Qty"
      Height          =   255
      Left            =   6060
      TabIndex        =   15
      Top             =   3150
      Width           =   1920
   End
   Begin VB.CheckBox chkHidePrice 
      Alignment       =   1  'Right Justify
      Caption         =   "Hide Price"
      Height          =   255
      Left            =   6060
      TabIndex        =   14
      Top             =   2895
      Width           =   1920
   End
   Begin VB.TextBox txtDescription 
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   1560
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1320
      Width           =   4155
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   315
      Index           =   0
      Left            =   6360
      TabIndex        =   19
      Top             =   7140
      Width           =   1035
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Index           =   1
      Left            =   7500
      TabIndex        =   20
      Top             =   7140
      Width           =   1035
   End
   Begin HFSystem.VBCombo cboJCCategory 
      Height          =   240
      Left            =   1560
      TabIndex        =   8
      Top             =   6735
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   423
   End
   Begin HFSystem.VBCombo cboPOType 
      Height          =   240
      Left            =   6030
      TabIndex        =   9
      Top             =   1245
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   423
      Style           =   2
   End
   Begin VB.Label Label1 
      Caption         =   "Type"
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
      Height          =   255
      Index           =   7
      Left            =   5940
      TabIndex        =   35
      Top             =   990
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "Attachments"
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
      Height          =   240
      Index           =   6
      Left            =   5940
      TabIndex        =   32
      Top             =   4215
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "Retainage"
      Height          =   255
      Index           =   5
      Left            =   6090
      TabIndex        =   31
      Top             =   3690
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Category"
      Height          =   240
      Index           =   4
      Left            =   540
      TabIndex        =   30
      Top             =   6780
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Cost Code"
      Height          =   240
      Index           =   3
      Left            =   540
      TabIndex        =   29
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Notes"
      Height          =   240
      Index           =   2
      Left            =   480
      TabIndex        =   28
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Purchase Order"
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
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   27
      Top             =   1080
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "Options"
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
      Height          =   240
      Index           =   0
      Left            =   5940
      TabIndex        =   26
      Top             =   1665
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Ship Via"
      Height          =   255
      Index           =   23
      Left            =   480
      TabIndex        =   25
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Terms"
      Height          =   240
      Index           =   24
      Left            =   480
      TabIndex        =   24
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Free On Board"
      Height          =   255
      Index           =   25
      Left            =   360
      TabIndex        =   23
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Description"
      Height          =   240
      Index           =   28
      Left            =   480
      TabIndex        =   22
      Top             =   1350
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Standard Text"
      Height          =   240
      Index           =   29
      Left            =   420
      TabIndex        =   21
      Top             =   1860
      Width           =   1035
   End
End
Attribute VB_Name = "FPOIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const SRCFILE = "FEditPOIndex::"

Private mPOIndex  As String
Private mCancel   As Boolean
Private mDirty    As Boolean
Private mSaved    As Boolean

Public Function Edit(POIndex) As Boolean
On Error Resume Next
    mCancel = False
    mSaved = False
    mPOIndex = POIndex
    Me.Show vbModal
    Edit = mSaved
End Function




Private Sub cboPOType_Click()
    mDirty = True
End Sub

Private Sub chkBuildPro_Click()
    mDirty = True
End Sub

Private Sub chkLiens_Click()
    mDirty = True
End Sub

Private Sub chkMPO_Click()
    mDirty = True
End Sub

Private Sub chkRequiresPaymentApproval_Click()
    mDirty = True
End Sub

Private Sub cmdNav_Click(Index As Integer)
    mCancel = True
    If SaveData(Index = 1) Then
        mCancel = False
        Unload Me
    End If
End Sub


Private Sub Form_Load()
    Dim s As String
    Dim rs As Recordset
    
    Call IniGetForm(Me)
    
    cboPOType.Clear
    cboPOType.AddItem "Purchase Order"
    cboPOType.AddItem "Subcontract"
    
    
    If HFApp.Options(AccountingSystem) = asQuickBooks Then
        Call LoadComboBox(cboJCCostCode, HFApp.Databases(dbHomefront), "SELECT ISNULL(Description,''),CostCode,0 FROM StandardCostCodes where DivisionID = " & HFApp.DivisionID & " ORDER BY 1")
        Call LoadComboBox(cboJCCategory, HFApp.Databases(dbHomefront), "SELECT ISNULL(Description,''),Category,0 FROM StandardCategories where DivisionID = " & HFApp.DivisionID & " ORDER BY 1")
    Else
        Call LoadComboBox(cboJCCostCode, HFApp.Databases(dbHomefront), "SELECT ISNULL(CostCode,'') + '    ' + ISNULL(Description,''),CostCode,0 FROM StandardCostCodes where DivisionID = " & HFApp.DivisionID & " ORDER BY 1")
        Call LoadComboBox(cboJCCategory, HFApp.Databases(dbHomefront), "SELECT ISNULL(Category,'') + '    ' + ISNULL(Description,''),Category,0 FROM StandardCategories where DivisionID = " & HFApp.DivisionID & " ORDER BY 1")
    End If
    
    
    Set rs = HFApp.SqlExec("select * from tblpoindex where DivisionID = " & HFApp.DivisionID & " and poindex=" & DbQuote(Str, mPOIndex))
    If rs.EOF Then
        txtPOIndex.Text = mPOIndex
        txtDescription.Text = ""
        txtStandardText.Text = ""
        txtShipVia.Text = ""
        txtFOB.Text = ""
        txtTerms.Text = ""
        txtNotes.Text = ""
        cboJCCostCode.ListIndex = -1
        cboJCCategory.ListIndex = -1
        cboPOType.ListIndex = 0
        chkBuildPro.Value = vbChecked
        chkMPO.Value = vbUnchecked
        chkRequiresPaymentApproval.Value = vbUnchecked
        chkHidePrice.Value = vbUnchecked
        chkHideQty.Value = vbUnchecked
        chkTotalOnly.Value = vbUnchecked
        txtRetainagePercent.Text = "0%"
    Else
        txtPOIndex.Text = "" & rs("POIndex")
        txtDescription.Text = "" & rs("Description")
        txtStandardText.Text = "" & rs("StandardText")
        txtShipVia.Text = "" & rs("ShipVia")
        txtFOB.Text = "" & rs("FOB")
        txtTerms.Text = "" & rs("Terms")
        txtNotes.Text = "" & rs("Notes")
        
        cboPOType.ListIndex = IIf("" & rs("POType") = "Subcontract", 1, 0)
        
        Call SetComboBoxListIndex(cboJCCostCode, , "" & rs("JCCostCode"))
        Call SetComboBoxListIndex(cboJCCategory, , "" & rs("JCCategory"))
        
        chkLiens.Value = IIf("" & rs("RequireLienRelease") = "True", vbChecked, vbUnchecked)
        chkMPO.Value = IIf("" & rs("MPO") = "True", vbChecked, vbUnchecked)
        chkBuildPro.Value = IIf("" & rs("BuildProEnabled") = "True", vbChecked, vbUnchecked)
        chkRequiresPaymentApproval.Value = IIf("" & rs("RequiresPaymentApproval") = "True", vbChecked, vbUnchecked)
        chkHidePrice.Value = IIf("" & rs("HidePrice") = "True", vbChecked, vbUnchecked)
        chkHideQty.Value = IIf("" & rs("HideQty") = "True", vbChecked, vbUnchecked)
        chkTotalOnly.Value = IIf("" & rs("TotalOnly") = "True", vbChecked, vbUnchecked)
        txtRetainagePercent.Text = Val("" & rs("RetainagePercent")) & "%"
        
    End If
    
    'now load doc attachments
    'NEW
    s = ""
    s = s & "SELECT DocumentClass" & vbCrLf
    s = s & " ,CASE WHEN p.Class IS NULL THEN 'False' ELSE 'True' END Selected " & vbCrLf
    s = s & "FROM dms_documentclasses c" & vbCrLf
    s = s & "LEFT OUTER JOIN POIndexDocuments p ON(c.DocumentClass=p.Class AND p.POIndex=" & DbQuote(Str, mPOIndex) & ")" & vbCrLf
    s = s & "ORDER BY 1" & vbCrLf
    
    'OLD
    s = ""
    s = s & "SELECT DISTINCT j.DocumentClass" & vbCrLf
    s = s & "      ,CASE WHEN p.Class IS NULL THEN 'False' ELSE 'True' END Selected " & vbCrLf
    s = s & "  FROM Attachments j " & vbCrLf
    s = s & "  LEFT OUTER JOIN POIndexDocuments p ON(j.DocumentClass=p.Class AND p.POIndex=" & DbQuote(Str, mPOIndex) & ")" & vbCrLf
    s = s & " WHERE ISNULL(j.DocumentClass,'')<>''" & vbCrLf
    s = s & "ORDER BY 1" & vbCrLf
    Set rs = HFApp.SqlExec(s, dbHomefront)
    With lstAttachments
        .Clear
        While Not rs.EOF
            .AddItem "" & rs(0)
            .Selected(.NewIndex) = "" & rs(1) = "True"
            rs.MoveNext
        Wend
    End With
    
    
    mDirty = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub


Private Sub lstAttachments_ItemCheck(item As Integer)
    mDirty = True
End Sub

Private Sub txtPOIndex_GotFocus()
    SelectAll txtPOIndex
End Sub
Private Sub txtDescription_GotFocus()
    SelectAll txtDescription
End Sub

Private Sub txtPOIndex_Validate(Cancel As Boolean)
    txtPOIndex.Text = Trim(txtPOIndex.Text)
End Sub

Private Sub txtRetainagePercent_Change()
    mDirty = True
End Sub

Private Sub txtRetainagePercent_GotFocus()
    SelectAll txtRetainagePercent
End Sub

Private Sub txtRetainagePercent_Validate(Cancel As Boolean)
    txtRetainagePercent.Text = Min(Max(Val(txtRetainagePercent.Text), 0), 100) & "%"
End Sub

Private Sub txtStandardText_GotFocus()
    SelectAll txtStandardText
End Sub
Private Sub txtShipVia_GotFocus()
    SelectAll txtShipVia
End Sub
Private Sub txtFOB_GotFocus()
    SelectAll txtFOB
End Sub
Private Sub txtTerms_GotFocus()
    SelectAll txtTerms
End Sub
Private Sub txtNotes_GotFocus()
    SelectAll txtNotes
End Sub

Private Sub txtPOIndex_Change()
    mDirty = True
End Sub
Private Sub txtDescription_Change()
    mDirty = True
End Sub
Private Sub txtStandardText_Change()
    mDirty = True
End Sub
Private Sub txtShipVia_Change()
    mDirty = True
End Sub
Private Sub txtFOB_Change()
    mDirty = True
End Sub
Private Sub txtTerms_Change()
    mDirty = True
End Sub
Private Sub txtNotes_Change()
    mDirty = True
End Sub
Private Sub chkHidePrice_Click()
    mDirty = True
End Sub
Private Sub chkHideQty_Click()
    mDirty = True
End Sub
Private Sub chkTotalOnly_Click()
    mDirty = True
End Sub
Private Sub cboJCCostCode_Click()
    mDirty = True
End Sub
Private Sub cboJCCategory_Click()
    mDirty = True
End Sub



Public Function SaveData(prompt As Boolean) As Boolean
On Error GoTo eh:

    Dim rc As Long
    Dim s As String
    Dim i As Long
    
    If Not mDirty Then
        SaveData = True
        Exit Function
    End If
    rc = vbYes
    If prompt Then rc = MsgBox("This data has changed." & vbCrLf & vbCrLf & "Do you want to save the changes?", vbExclamation + vbYesNoCancel, Me.Caption)
    Select Case rc
        Case vbNo:     SaveData = True:     Exit Function
        Case vbCancel: Exit Function
    End Select
    
    
    If Not ValidateData() Then Exit Function
    
    Screen.MousePointer = vbHourglass
    
    s = ""
    s = s & "UPDATE tblPOIndex" & vbCrLf
    s = s & "SET POIndex=" & DbQuote(Str, txtPOIndex.Text) & vbCrLf
    s = s & "   ,Description=" & DbQuote(Str, txtDescription.Text) & vbCrLf
    s = s & "   ,StandardText=" & DbQuote(Str, txtStandardText.Text) & vbCrLf
    s = s & "   ,ShipVia=" & DbQuote(Str, txtShipVia.Text) & vbCrLf
    s = s & "   ,POType=" & DbQuote(Str, cboPOType.Text) & vbCrLf
    s = s & "   ,FOB=" & DbQuote(Str, txtFOB.Text) & vbCrLf
    s = s & "   ,Terms=" & DbQuote(Str, txtTerms.Text) & vbCrLf
    s = s & "   ,Notes=" & DbQuote(Str, txtNotes.Text) & vbCrLf
    s = s & "   ,RetainagePercent=" & DbQuote(Num, txtRetainagePercent.Text) & vbCrLf
    s = s & "   ,JCCostCode=" & DbQuote(Str, GetComboBoxListKey(cboJCCostCode)) & vbCrLf
    s = s & "   ,JCCategory=" & DbQuote(Str, GetComboBoxListKey(cboJCCategory)) & vbCrLf
    s = s & "   ,MPO=" & DbQuote(Bit, chkMPO.Value = vbChecked) & vbCrLf
    s = s & "   ,BuildProEnabled=" & DbQuote(Bit, chkBuildPro.Value = vbChecked) & vbCrLf
    s = s & "   ,RequireLienRelease=" & DbQuote(Bit, chkLiens.Value = vbChecked) & vbCrLf
    s = s & "   ,RequiresPaymentApproval=" & DbQuote(Bit, chkRequiresPaymentApproval.Value = vbChecked) & vbCrLf
    s = s & "   ,HidePrice=" & DbQuote(Bit, chkHidePrice.Value = vbChecked) & vbCrLf
    s = s & "   ,HideQty=" & DbQuote(Bit, chkHideQty.Value = vbChecked) & vbCrLf
    s = s & "   ,TotalOnly=" & DbQuote(Bit, chkTotalOnly.Value = vbChecked) & vbCrLf
    s = s & "   ,UStmp=" & DbQuote(Str, HFApp.LoginID) & vbCrLf
    s = s & "   ,TStmp=GETDATE()" & vbCrLf
    s = s & "WHERE DivisionID = " & HFApp.DivisionID & " and POIndex=" & DbQuote(Str, mPOIndex) & vbCrLf
    Call HFApp.SqlExec(s)

    'update doc attachments
    s = "DELETE FROM POIndexDocuments WHERE POIndex=" & DbQuote(Str, mPOIndex)
    Call HFApp.SqlExec(s, dbHomefront)
    With lstAttachments
    For i = 0 To .ListCount - 1
        If .Selected(i) Then
            s = "INSERT INTO POIndexDocuments(POIndex,Class) VALUES(" & DbQuote(Str, txtPOIndex.Text) & "," & DbQuote(Str, .list(i)) & ")"
            Call HFApp.SqlExec(s, dbHomefront)
        End If
    Next
    End With
    
    
    'update retainage amount for un-issued items
    s = ""
    s = s & "UPDATE JobPOIndex" & vbCrLf
    s = s & "   SET RetainagePercent=" & DbQuote(Num, txtRetainagePercent.Text) & vbCrLf
    s = s & "  FROM EstimateItems e JOIN JobPOIndex j ON(e.Job=j.Job_No AND e.POIndex=j.POIndex)" & vbCrLf
    s = s & " WHERE e.POGenBatch=0 " & vbCrLf
    s = s & "   AND e.POIndex=" & DbQuote(Str, mPOIndex) & vbCrLf
    Call HFApp.SqlExec(s, dbHomefront)
    
    
    'update foreign tables if poindex name has changed
    If txtPOIndex.Text <> mPOIndex Then
    
        On Error Resume Next
        s = ""
        Call HFApp.SqlExec("UPDATE CustomPreEstimateItems SET POIndex=" & DbQuote(Str, txtPOIndex.Text) & " WHERE POIndex=" & DbQuote(Str, mPOIndex))
        Call HFApp.SqlExec("UPDATE EstimateItems          SET POIndex=" & DbQuote(Str, txtPOIndex.Text) & " WHERE DivisionID = " & HFApp.DivisionID & " and POIndex=" & DbQuote(Str, mPOIndex))
        Call HFApp.SqlExec("UPDATE JobPOIndex             SET POIndex=" & DbQuote(Str, txtPOIndex.Text) & " WHERE DivisionID = " & HFApp.DivisionID & " and POIndex=" & DbQuote(Str, mPOIndex))
        Call HFApp.SqlExec("UPDATE PayPoints              SET POIndex=" & DbQuote(Str, txtPOIndex.Text) & " WHERE POIndex=" & DbQuote(Str, mPOIndex))
        Call HFApp.SqlExec("UPDATE POAreaVendor           SET POIndex=" & DbQuote(Str, txtPOIndex.Text) & " WHERE DivisionID = " & HFApp.DivisionID & " and POIndex=" & DbQuote(Str, mPOIndex))
        Call HFApp.SqlExec("UPDATE POMaster               SET POIndex=" & DbQuote(Str, txtPOIndex.Text) & " WHERE DivisionID = " & HFApp.DivisionID & " and POIndex=" & DbQuote(Str, mPOIndex))
        Call HFApp.SqlExec("UPDATE tblPhaseItem           SET POIndex=" & DbQuote(Str, txtPOIndex.Text) & " WHERE DivisionID = " & HFApp.DivisionID & " and POIndex=" & DbQuote(Str, mPOIndex))
        Call HFApp.SqlExec("UPDATE tbldbAssemblyDetails   SET POIndex=" & DbQuote(Str, txtPOIndex.Text) & " WHERE DivisionID = " & HFApp.DivisionID & " and POIndex=" & DbQuote(Str, mPOIndex))
        On Error GoTo eh
        
    End If

    s = ""
    s = s & "insert into DivisionPOIndexes(POIndex,DivisionID,POFormat)" & vbCrLf
    s = s & "select p.POIndex,d.DivisionID,p.POFormat from tblPOindex p" & vbCrLf
    s = s & "join Divisions d on 1 = 1" & vbCrLf
    s = s & "Left Outer join DivisionPoIndexes dp on (dp.POIndex = p.POIndex and dp.DivisionID = d.DivisionID)" & vbCrLf
    s = s & "where dp.DivisionID is null and p.DivisionID = " & HFApp.DivisionID & vbCrLf
    Call HFApp.SqlExec(s, dbHomefront)

    mSaved = True
    mDirty = False
    SaveData = True
    Screen.MousePointer = vbDefault
    Exit Function
    
eh: Select Case True
    Case InStr(1, Err.Description, "duplicate") > 0:
        Screen.MousePointer = vbDefault
        MsgBox "Purchase order """ & Trim(txtPOIndex.Text) & """ has already been entered. Please specify a different value.", vbInformation, App.ProductName
    Case Else:  Call errHandler(SRCFILE & "SaveData", s)
    End Select
End Function

Private Function ValidateData() As Boolean
    Dim s As String
    
    ValidateData = True
    If Trim(txtPOIndex.Text) = "" Then s = s & "Purchase order is required"
    
    
    If s <> "" Then
        ValidateData = False
        MsgBox s, vbExclamation, App.ProductName
    End If
End Function
