VERSION 5.00
Begin VB.Form FDuplicateItems 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Duplicate Items"
   ClientHeight    =   3420
   ClientLeft      =   1935
   ClientTop       =   1620
   ClientWidth     =   5745
   Icon            =   "FDuplicateItems.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboPOIndex 
      Enabled         =   0   'False
      Height          =   240
      Left            =   1560
      TabIndex        =   5
      Top             =   2280
      Width           =   3225
   End
   Begin VB.CheckBox chkRenumber 
      Caption         =   "Re-number the new items."
      Height          =   195
      Left            =   1320
      TabIndex        =   1
      Top             =   1140
      Width           =   2295
   End
   Begin VB.CheckBox chkChangePO 
      Caption         =   "Change the PO index"
      Height          =   195
      Left            =   1320
      TabIndex        =   4
      Top             =   2040
      Width           =   2985
   End
   Begin VB.TextBox txtStart 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   215
      Left            =   2820
      TabIndex        =   2
      Text            =   "0010"
      Top             =   1380
      Width           =   600
   End
   Begin VB.TextBox txtIncrement 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   215
      Left            =   2820
      TabIndex        =   3
      Text            =   "10"
      Top             =   1620
      Width           =   600
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   1560
      Picture         =   "FDuplicateItems.frx":000C
      TabIndex        =   6
      ToolTipText     =   "Login"
      Top             =   2910
      Width           =   1215
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   2865
      Picture         =   "FDuplicateItems.frx":0596
      TabIndex        =   7
      ToolTipText     =   "Cancel"
      Top             =   2910
      Width           =   1215
   End
   Begin VB.ComboBox cboPhase 
      Height          =   240
      Left            =   1560
      TabIndex        =   0
      Top             =   750
      Width           =   3225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Phase"
      Height          =   195
      Index           =   4
      Left            =   1020
      TabIndex        =   11
      Top             =   750
      Width           =   450
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   210
      Picture         =   "FDuplicateItems.frx":0B20
      Top             =   150
      Width           =   480
   End
   Begin VB.Label lblIncrement 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Increment By"
      Enabled         =   0   'False
      Height          =   195
      Left            =   1770
      TabIndex        =   10
      Top             =   1620
      Width           =   930
   End
   Begin VB.Label lblStart 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Starting number"
      Enabled         =   0   'False
      Height          =   195
      Left            =   1590
      TabIndex        =   9
      Top             =   1380
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Make a copy of these items."
      Height          =   195
      Index           =   2
      Left            =   960
      TabIndex        =   8
      Top             =   240
      Width           =   1995
   End
End
Attribute VB_Name = "FDuplicateItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SRCFILE = "FDuplicateItems::"

Private mCancel As Boolean
Private mWhereClause As String

Public Function ShowForm(WhereClause As String) As Boolean
    
    mCancel = True
    mWhereClause = WhereClause
    Me.Show vbModal
    ShowForm = Not mCancel
    
End Function

Private Sub chkChangePO_Click()
    cboPOIndex.Enabled = chkChangePO.value = vbChecked
End Sub


Private Sub chkRenumber_Click()
    Dim b As Boolean
    b = chkRenumber.value = vbChecked
    
    txtStart.Enabled = b
    txtStart.BackColor = IIf(b, vbWindowBackground, vbButtonFace)
    lblStart.Enabled = b
    
    txtIncrement.Enabled = b
    txtIncrement.BackColor = IIf(b, vbWindowBackground, vbButtonFace)
    lblIncrement.Enabled = b
    
End Sub

Private Sub cmdNav_Click(Index As Integer)
    Select Case Index
        Case 0 'ok
            If SaveData() Then
                mCancel = False
                Unload Me
            End If
            
        Case 1 'cancel
            mCancel = True
            Unload Me
    End Select
End Sub

Private Function SaveData() As Boolean
On Error GoTo eh
    Dim rs As Recordset
    Dim s As String
    Dim i As Long
    Dim Phase As String
    Dim POIndex As String
    Dim Start As Double
    Dim Incr  As Double
    Dim Decimals As Long
    Dim Digits   As Long
    Dim ItemFormat As String
    Dim Item As String
    
    Phase = GetComboBoxListKey(cboPhase)
    'phase=trim(phase)
    POIndex = GetComboBoxListKey(cboPOIndex)
    
    Start = Val(txtStart.Text)
    Incr = Val(txtIncrement.Text)
    Decimals = Max(Len(Trim(Parse(txtStart.Text, 2, "."))), Len(Trim(Parse(txtIncrement.Text, 2, "."))))
    Digits = Max(Len(Trim(Parse(txtStart.Text, 1, "."))), Len(Trim(Parse(txtIncrement.Text, 1, "."))))
    ItemFormat = String(Digits, "0") & IIf(Decimals < 1, "", ".") & String(Decimals, "0")
    
    s = "select * from tblPhaseItem where " & mWhereClause & IIf(mWhereClause <> "", " and DivisionID = " & HFApp.DivisionID, "DivisionID = " & HFApp.DivisionID) & " order by item"
    Set rs = HFApp.SqlExec(s, dbHomeFront)
    i = 0
    While Not rs.EOF
        
        s = ""
        s = s & "insert into tblPhaseItem(DivisionID,Phase,Item,ItemNumber,CostCategory,POIndex,PriceLink,Description,Notes,JCCostCode,JCCategory,OrderUOM,TakeoffUOM,ConversionFactor,Price,TaxGroup,WastePercent,RoundDir,RoundTo,PhaseSortOrder,ItemSortOrder)" & vbCrLf
        s = s & "values(" & HFApp.DivisionID & "," & DbQuote(str, Phase) & vbCrLf
        If chkRenumber.value = vbChecked Then
            Item = format(Start + i * Incr, ItemFormat)
            s = s & "      ," & DbQuote(str, Item & rs("CostCategory")) & vbCrLf
            s = s & "      ," & DbQuote(str, Item) & vbCrLf
        Else
            Item = "" & rs("Item")
            s = s & "      ," & DbQuote(str, "" & rs("Item")) & vbCrLf
            s = s & "      ," & DbQuote(str, "" & rs("ItemNumber")) & vbCrLf
        End If
        s = s & "      ," & DbQuote(str, "" & rs("CostCategory")) & vbCrLf
        If chkChangePO.value = vbChecked Then
            s = s & "      ," & DbQuote(str, POIndex) & vbCrLf
        Else
            s = s & "      ," & DbQuote(str, "" & rs("POIndex")) & vbCrLf
        End If
        s = s & "      ," & DbQuote(Num, "" & rs("PriceLink")) & vbCrLf
        s = s & "      ," & DbQuote(str, "" & rs("Description")) & vbCrLf
        s = s & "      ," & DbQuote(str, "" & rs("Notes")) & vbCrLf
        s = s & "      ," & DbQuote(str, "" & rs("JCCostCode")) & vbCrLf
        s = s & "      ," & DbQuote(str, "" & rs("JCCategory")) & vbCrLf
        s = s & "      ," & DbQuote(str, "" & rs("OrderUOM")) & vbCrLf
        s = s & "      ," & DbQuote(str, "" & rs("TakeoffUOM")) & vbCrLf
        s = s & "      ," & DbQuote(Num, "" & rs("ConversionFactor")) & vbCrLf
        s = s & "      ," & DbQuote(Num, "" & rs("Price")) & vbCrLf
        s = s & "      ," & DbQuote(str, "" & rs("TaxGroup")) & vbCrLf
        s = s & "      ," & DbQuote(Num, "" & rs("WastePercent")) & vbCrLf
        s = s & "      ," & DbQuote(Num, "" & rs("RoundDir")) & vbCrLf
        s = s & "      ," & DbQuote(Num, "" & rs("RoundTo")) & vbCrLf
        s = s & "      ," & DbQuote(Num, "" & rs("Phase")) & vbCrLf
        s = s & "      ," & DbQuote(Num, "" & rs("ItemNumber")) & ")"
        Call HFApp.SqlExec(s, dbHomeFront)
        
        i = i + 1
        rs.MoveNext
    Wend

    
    
    SaveData = True


EXITSUB:
    Screen.MousePointer = vbDefault
    Exit Function
    
eh: Select Case True
    Case InStr(1, Err.Description, "duplicate") > 0:
        Screen.MousePointer = vbDefault
        MsgBox "Precision Builder already has an item numbered " & rs("phase") & "/" & rs("item") & ". Please specify a different value.", vbInformation, App.ProductName
    Case Else:  Call errHandler(SRCFILE & "SaveData", s)
    End Select
    Exit Function
    
End Function


Private Sub Form_Load()
    Dim s As String
    
    
    Call IniGetForm(Me)
    s = "select case when poindex=description then poindex else poindex + isnull(' ' + nullif(description,''),'') end,poindex,0 from tblpoindex where DivisionID = " & HFApp.DivisionID & " order by 1"
    Call LoadComboBox(cboPOIndex, HFApp.Databases(dbHomeFront), s)
    
    
    s = "select isnull(phase,'') +' - '+ isnull(description,''),phase,0 from tblestphases where DivisionID = " & HFApp.DivisionID & " and groupphase=0 order by 1"
    Call LoadComboBox(cboPhase, HFApp.Databases(dbHomeFront), s)
    
    txtStart.BackColor = vbButtonFace
    txtIncrement.BackColor = vbButtonFace
        
On Error Resume Next
    cboPhase.ListIndex = 0
    cboPOIndex.ListIndex = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub

Private Sub txtIncrement_GotFocus()
    SelectAll txtIncrement
End Sub

Private Sub txtIncrement_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric(txtIncrement.Text)
End Sub

Private Sub txtStart_GotFocus()
    SelectAll txtStart
End Sub

Private Sub txtStart_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric(txtStart.Text)
End Sub
