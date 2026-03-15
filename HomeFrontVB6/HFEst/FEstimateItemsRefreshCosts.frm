VERSION 5.00
Begin VB.Form FEstimateItemsRefreshCosts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Refresh Costs"
   ClientHeight    =   4635
   ClientLeft      =   4875
   ClientTop       =   1665
   ClientWidth     =   5475
   Icon            =   "FEstimateItemsRefreshCosts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkRefreshOverridden 
      Caption         =   "Refresh costs on manually updated items."
      Height          =   195
      Left            =   960
      TabIndex        =   10
      Top             =   240
      Width           =   4215
   End
   Begin HFEst.VBCombo cboVarianceCat 
      Height          =   240
      Left            =   1950
      TabIndex        =   1
      Top             =   810
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   423
   End
   Begin VB.CheckBox chkSetVarianceCat 
      Caption         =   "Set the variance category on all changed items."
      Height          =   195
      Left            =   960
      TabIndex        =   0
      Top             =   540
      Width           =   4215
   End
   Begin VB.TextBox txtEffectiveDate 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1950
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1440
      Width           =   1515
   End
   Begin VB.OptionButton optForecast 
      Caption         =   "Refresh costs using a forecast"
      Height          =   240
      Left            =   930
      TabIndex        =   4
      Top             =   1800
      Width           =   3915
   End
   Begin VB.OptionButton optCurrent 
      Caption         =   "Refresh costs using current prices "
      Height          =   240
      Left            =   930
      TabIndex        =   2
      Top             =   1140
      Value           =   -1  'True
      Width           =   3915
   End
   Begin VB.ListBox lstForecasts 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   1785
      ItemData        =   "FEstimateItemsRefreshCosts.frx":000C
      Left            =   1230
      List            =   "FEstimateItemsRefreshCosts.frx":0034
      TabIndex        =   5
      Top             =   2100
      Width           =   3675
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   4110
      TabIndex        =   7
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Height          =   375
      Index           =   0
      Left            =   2850
      TabIndex        =   6
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Category:"
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   9
      Top             =   810
      Width           =   855
   End
   Begin VB.Image cmdEffectiveDate 
      Height          =   240
      Left            =   3480
      Picture         =   "FEstimateItemsRefreshCosts.frx":00CB
      Top             =   1440
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Effective:"
      Height          =   255
      Index           =   9
      Left            =   1050
      TabIndex        =   8
      Top             =   1440
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   240
      Picture         =   "FEstimateItemsRefreshCosts.frx":0215
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "FEstimateItemsRefreshCosts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FEstimateItemsRefreshCosts::"

Private mCancel As Boolean
Private mCostBasis As CostBasisTypes
Private mEffectiveDate As Date
Private mApplyVarianceCat As Boolean
Private mVarianceCat As String
Private mRefreshOverridden As Boolean

Public Function ShowForm(CostBasis As Long, EffectiveDate As Date, ApplyVarianceCat As Boolean, VarianceCat As String, RefreshOverridden As Boolean) As Boolean
    mCancel = False
    mVarianceCat = VarianceCat
    Me.Show vbModal
    ShowForm = Not mCancel
    CostBasis = mCostBasis
    EffectiveDate = mEffectiveDate
    ApplyVarianceCat = mApplyVarianceCat
    VarianceCat = mVarianceCat
    RefreshOverridden = mRefreshOverridden
    
End Function


Private Sub cmdNav_Click(Index As Integer)
    Dim i As Long
    If Index = 0 Then
        If (optCurrent And Not IsDate(txtEffectiveDate)) Then MsgBox "You must specify an effective date", vbInformation, App.ProductName:  Exit Sub
        If (optForecast And lstForecasts.ListIndex < 0) Then MsgBox "You must select a forecast", vbInformation, App.ProductName:           Exit Sub
        
        mCancel = False
        If optCurrent.value Then
            mCostBasis = cbCurrent
            mEffectiveDate = DateValue(txtEffectiveDate)
        Else
            mCostBasis = cbForecast1 + lstForecasts.ListIndex
        End If
        
        mRefreshOverridden = chkRefreshOverridden = vbChecked
        mApplyVarianceCat = chkSetVarianceCat = vbChecked
        mVarianceCat = GetComboBoxListKey(cboVarianceCat)
    
        Unload Me
        
    Else
        mCancel = True
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim s As String
    Dim rs As Recordset
    
    Call IniGetForm(Me)
    txtEffectiveDate.Text = format(Now, HFApp.Options(DateFormat))
    
    s = ""
    s = s & "SELECT MAX(Custom_Description),0  FROM CustomDescriptions WHERE Item ='Forecast1' UNION ALL" & vbCrLf
    s = s & "SELECT MAX(Custom_Description),1  FROM CustomDescriptions WHERE Item ='Forecast2' UNION ALL" & vbCrLf
    s = s & "SELECT MAX(Custom_Description),2  FROM CustomDescriptions WHERE Item ='Forecast3' UNION ALL" & vbCrLf
    s = s & "SELECT MAX(Custom_Description),3  FROM CustomDescriptions WHERE Item ='Forecast4' UNION ALL" & vbCrLf
    s = s & "SELECT MAX(Custom_Description),4  FROM CustomDescriptions WHERE Item ='Forecast5' UNION ALL" & vbCrLf
    s = s & "SELECT MAX(Custom_Description),5  FROM CustomDescriptions WHERE Item ='Forecast6' UNION ALL" & vbCrLf
    s = s & "SELECT MAX(Custom_Description),6  FROM CustomDescriptions WHERE Item ='Forecast7' UNION ALL" & vbCrLf
    s = s & "SELECT MAX(Custom_Description),7  FROM CustomDescriptions WHERE Item ='Forecast8' UNION ALL" & vbCrLf
    s = s & "SELECT MAX(Custom_Description),8  FROM CustomDescriptions WHERE Item ='Forecast9' UNION ALL" & vbCrLf
    s = s & "SELECT MAX(Custom_Description),9  FROM CustomDescriptions WHERE Item ='Forecast10' UNION ALL" & vbCrLf
    s = s & "SELECT MAX(Custom_Description),10 FROM CustomDescriptions WHERE Item ='Forecast11' UNION ALL" & vbCrLf
    s = s & "SELECT MAX(Custom_Description),11 FROM CustomDescriptions WHERE Item ='Forecast12'" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    While Not rs.EOF
        lstForecasts.list(rs(1)) = "" & rs(0)
        rs.MoveNext
    Wend
    
    
    chkRefreshOverridden.value = IIf(HFApp.Options.ValueByName("RefreshOverridden") = "True", vbChecked, vbUnchecked)
    chkSetVarianceCat.value = IIf(HFApp.Options.ValueByName("ApplyVarianceCategory") = "True", vbChecked, vbUnchecked)
    Call LoadComboBox(cboVarianceCat, HFApp.Databases(dbHomeFront), "select category+' - '+description,category,0 from standardcategories where isvariance=1 and divisionid=" & DbQuote(str, HFApp.DivisionID) & " order by description")
    Call SetComboBoxListIndex(cboVarianceCat, , mVarianceCat)
    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call IniPutForm(Me)
    HFApp.Options.ValueByName("ApplyVarianceCategory") = chkSetVarianceCat.value = vbChecked
    HFApp.Options.ValueByName("RefreshOverridden") = chkRefreshOverridden.value = vbChecked

End Sub

Private Sub optCurrent_Click()

    txtEffectiveDate.Enabled = optCurrent.value
    txtEffectiveDate.BackColor = IIf(txtEffectiveDate.Enabled, vbWindowBackground, vbButtonFace)
    
    lstForecasts.Enabled = Not optCurrent.value
    lstForecasts.BackColor = IIf(lstForecasts.Enabled, vbWindowBackground, vbButtonFace)
    lstForecasts.ListIndex = IIf(lstForecasts.Enabled, 0, -1)
    
End Sub

Private Sub optForecast_Click()
    Call optCurrent_Click
End Sub

Private Sub cmdEffectiveDate_Click()
On Error Resume Next
    Call txtEffectiveDate_KeyDown(vbKeyF4, 0)
End Sub
Private Sub txtEffectiveDate_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo eh
    If KeyCode = vbKeyF4 And Shift = 0 Then Call DCalendar.Popup(txtEffectiveDate)
Exit Sub
eh: Call errHandler(SRCFILE & "txtEffectiveDate_KeyDown")
End Sub
Private Sub txtEffectiveDate_Validate(Cancel As Boolean)
On Error GoTo eh
    If IsDate(txtEffectiveDate) Or txtEffectiveDate = "" Then
        txtEffectiveDate.Text = format(txtEffectiveDate, HFApp.Options(DateFormat))
    Else
        MsgBox "Not a valid date", vbExclamation
        Call SelectAll(txtEffectiveDate)
        Cancel = True
    End If
    Exit Sub
eh: Call errHandler(SRCFILE & "txtEffectiveDate_Validate")
End Sub
