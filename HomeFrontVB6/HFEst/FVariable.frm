VERSION 5.00
Begin VB.Form FVariable 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Variable"
   ClientHeight    =   4635
   ClientLeft      =   3105
   ClientTop       =   2145
   ClientWidth     =   4830
   Icon            =   "FVariable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmNumericOptions 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1695
      Left            =   1890
      TabIndex        =   10
      Top             =   660
      Width           =   2895
      Begin VB.Frame frmTOSystem 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   795
         Left            =   0
         TabIndex        =   21
         Top             =   960
         Width           =   3075
         Begin VB.Label lblConditionName 
            AutoSize        =   -1  'True
            Caption         =   "lblConditionName"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   90
            TabIndex        =   23
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblCondition 
            AutoSize        =   -1  'True
            Caption         =   "Condition..."
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
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Width           =   780
         End
      End
      Begin VB.TextBox txtDefault 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1965
         TabIndex        =   5
         Text            =   "0"
         Top             =   240
         Width           =   675
      End
      Begin VB.ComboBox cboUOM 
         Height          =   315
         ItemData        =   "FVariable.frx":000C
         Left            =   1965
         List            =   "FVariable.frx":0034
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   570
         Width           =   675
      End
      Begin VB.TextBox txtMax 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   615
         TabIndex        =   4
         Text            =   "9999"
         Top             =   570
         Width           =   675
      End
      Begin VB.TextBox txtMin 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   615
         TabIndex        =   3
         Text            =   "-9999"
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Default"
         Height          =   195
         Index           =   9
         Left            =   1380
         TabIndex        =   14
         Top             =   270
         Width           =   510
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "UOM"
         Height          =   195
         Index           =   8
         Left            =   1515
         TabIndex        =   13
         Top             =   630
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Max"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   300
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Min"
         Height          =   195
         Index           =   4
         Left            =   285
         TabIndex        =   11
         Top             =   270
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Properties"
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
         Index           =   0
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   870
      End
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   750
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
   Begin VB.TextBox txtHelp 
      Height          =   1485
      Left            =   270
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   2460
      Width           =   4335
   End
   Begin VB.OptionButton optVariableType 
      Caption         =   "List of values"
      Height          =   225
      Index           =   1
      Left            =   390
      TabIndex        =   2
      Top             =   1140
      Width           =   1365
   End
   Begin VB.OptionButton optVariableType 
      Caption         =   "Numeric"
      Height          =   225
      Index           =   0
      Left            =   390
      TabIndex        =   1
      Top             =   900
      Value           =   -1  'True
      Width           =   1395
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Height          =   375
      Index           =   0
      Left            =   2220
      Picture         =   "FVariable.frx":006B
      TabIndex        =   8
      ToolTipText     =   "Cancel"
      Top             =   4110
      Width           =   1215
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3480
      Picture         =   "FVariable.frx":05F5
      TabIndex        =   9
      ToolTipText     =   "Cancel"
      Top             =   4110
      Width           =   1215
   End
   Begin VB.TextBox txtListOfValues 
      Height          =   1455
      Left            =   1890
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Top             =   870
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name"
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
      Index           =   1
      Left            =   210
      TabIndex        =   19
      Top             =   270
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Help Notes"
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
      Index           =   7
      Left            =   270
      TabIndex        =   17
      Top             =   2220
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Values"
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
      Left            =   1890
      TabIndex        =   16
      Top             =   660
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   3
      Left            =   300
      TabIndex        =   15
      Top             =   660
      Width           =   435
   End
End
Attribute VB_Name = "FVariable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mName   As String
Private mDirty  As Boolean
Private mCancel As Boolean

Private mConditionName As String
Private mConditionType As String
Private mConditionUOM  As String


Private Sub cboUOM_Click()
    mDirty = True
End Sub

Private Sub cmdNav_Click(Index As Integer)
    If SaveData(Index = 1) Then Unload Me
End Sub

Private Sub Form_Load()
    Dim s As String
    
    Call IniGetForm(Me)
    
    Select Case TakeoffSystem
        Case tsOnScreen:   frmTOSystem.Visible = OpenOnScreen()
        Case tsPlanSwift:  frmTOSystem.Visible = True
        Case Else:         frmTOSystem.Visible = False
    End Select
    lblConditionName.Caption = ""
    
    s = "select uom,null,null from uom order by 1"
    Call LoadComboBox(cboUOM, HFApp.Databases(dbHomefront), s)
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub

Private Sub lblCondition_Click()
    Dim s As String
    Dim ps As Planswift

    Select Case TakeoffSystem
        Case tsOnScreen
            s = ""
            s = s & "select c.Name" & vbCrLf
            s = s & "      ,q.description as Type" & vbCrLf
            s = s & "      ,u.description as UOM" & vbCrLf
            s = s & "  from bidconditions c" & vbCrLf
            s = s & "      ,HF_UnitCodes u" & vbCrLf
            s = s & "      ,HF_QtyCodes  q" & vbCrLf
            s = s & " where c.biduid is null" & vbCrLf
            s = s & "   and c.uom1=u.code" & vbCrLf
            s = s & "   and c.quantity1=q.code" & vbCrLf
            s = s & "UNION" & vbCrLf
            s = s & "select c.name" & vbCrLf
            s = s & "      ,q.description as type" & vbCrLf
            s = s & "      ,u.description as uom" & vbCrLf
            s = s & "  from bidconditions c" & vbCrLf
            s = s & "      ,HF_UnitCodes u" & vbCrLf
            s = s & "      ,HF_QtyCodes  q" & vbCrLf
            s = s & " where c.biduid is null" & vbCrLf
            s = s & "   and c.uom2=u.code" & vbCrLf
            s = s & "   and c.quantity2=q.code" & vbCrLf
            s = s & "UNION" & vbCrLf
            s = s & "select c.name" & vbCrLf
            s = s & "      ,q.description as type" & vbCrLf
            s = s & "      ,u.description as uom" & vbCrLf
            s = s & "  from bidconditions c" & vbCrLf
            s = s & "      ,HF_UnitCodes u" & vbCrLf
            s = s & "      ,HF_QtyCodes  q" & vbCrLf
            s = s & " where c.biduid is null" & vbCrLf
            s = s & "   and c.uom3=u.code" & vbCrLf
            s = s & "   and c.quantity3=q.code" & vbCrLf
            If FPickList.Choose(OnScreenConnection, "Condition", s) Then
                mDirty = True
                mConditionName = FPickList.SelectedItem("Name")
                mConditionType = FPickList.SelectedItem("Type")
                mConditionUOM = FPickList.SelectedItem("UOM")
                lblConditionName.Caption = mConditionName & vbCrLf & mConditionType & ", " & mConditionUOM
            End If

        Case tsPlanSwift
            Set ps = New Planswift
            If ps.SelectVariable(False) Then
                mDirty = True
                mConditionName = ps.SelectedItem("Path") & Chr(1) & ps.SelectedItem("VariableName") & Chr(1) & ps.SelectedItem("ValueName")
                mConditionType = ""
                mConditionUOM = ps.SelectedItem("UOM")
                lblConditionName.Caption = ps.SelectedItem("Path") & "/" & ps.SelectedItem("VariableName") & "/" & ps.SelectedItem("ValueName")
            End If


   End Select
    
End Sub

Private Sub optVariableType_Click(Index As Integer)
    mDirty = True
    frmNumericOptions.Visible = Index = 0
    txtListOfValues.Visible = Index = 1
End Sub

Private Sub txtMin_Validate(Cancel As Boolean)
    txtMin.Text = Val(txtMin.Text)
End Sub
Private Sub txtMax_Validate(Cancel As Boolean)
    txtMax.Text = Val(txtMax.Text)
End Sub
Private Sub txtDefault_Validate(Cancel As Boolean)
    txtDefault.Text = Val(txtDefault.Text)
End Sub


Public Function EditVariable(Name As String) As Boolean
    Dim s As String
    Dim rs As Recordset
    
    Load Me
        
    mName = Name
    txtName.Text = mName
    
    s = "select * from variables where name=" & DbQuote(Str, mName)
    Set rs = HFApp.SqlExec(s, dbHomefront)
    If rs.EOF Then
        txtMin.Text = 0
        txtMax.Text = 9999999
        txtDefault.Text = 0
    Else
        optVariableType(Val("" & rs("VariableType"))).value = True
        txtMin.Text = Val("" & rs("minimumvalue"))
        txtMax.Text = Val("" & rs("maximumvalue"))
        txtDefault.Text = Val("" & rs("defaultvalue"))
        txtListOfValues.Text = Replace("" & rs("ListOfValues"), "|", vbCrLf)
        txtHelp.Text = "" & rs("Help")
        
        mConditionName = "" & rs("Conditionname")
        mConditionType = "" & rs("Conditiontype")
        mConditionUOM = "" & rs("Conditionuom")
        
        If TakeoffSystem = tsPlanSwift Then
            lblConditionName.Caption = Parse(mConditionName, 1, Chr(1)) & "/" & Parse(mConditionName, 2, Chr(1)) & "/" & Parse(mConditionName, 3, Chr(1))
        Else
            lblConditionName.Caption = mConditionName & vbCrLf & mConditionType & ", " & mConditionUOM
        End If
        

        Call SetComboBoxListIndex(cboUOM, "" & rs("uom"))
    End If
    
    mDirty = False
    mCancel = True
    Me.Show vbModal
    EditVariable = Not mCancel
    
End Function

Private Function SaveData(prompt As Boolean) As Boolean
    Dim s As String
    Dim values As String
    
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
    
    
    If optVariableType(1).value Then
        values = Replace(txtListOfValues.Text, vbCrLf, "|")
        
        If Trim(Replace(Replace(values, "|", ""), " ", "")) = "" Then
            Call MsgBox("You must specify at least one value.", vbExclamation, App.ProductName)
            Exit Function
        End If
        
    Else
        values = ""
    End If

    If mName = "" Then
        s = ""
        s = s & "insert into variables(name,variabletype,minimumvalue,maximumvalue,defaultvalue,listofvalues,uom,help,conditionname,conditiontype,conditionuom)" & vbCrLf
        s = s & "values(" & DbQuote(Str, txtName.Text) & vbCrLf
        s = s & "      ," & DbQuote(Num, IIf(optVariableType(0).value, 0, 1)) & vbCrLf
        s = s & "      ," & DbQuote(Num, txtMin.Text) & vbCrLf
        s = s & "      ," & DbQuote(Num, txtMax.Text) & vbCrLf
        s = s & "      ," & DbQuote(Num, txtDefault.Text) & vbCrLf
        s = s & "      ," & DbQuote(Str, values) & vbCrLf
        s = s & "      ," & DbQuote(Str, cboUOM.Text) & vbCrLf
        s = s & "      ," & DbQuote(Str, txtHelp.Text) & vbCrLf
        s = s & "      ," & DbQuote(Str, mConditionName) & vbCrLf
        s = s & "      ," & DbQuote(Str, mConditionType) & vbCrLf
        s = s & "      ," & DbQuote(Str, mConditionUOM) & vbCrLf
        s = s & ")" & vbCrLf
        HFApp.SqlExec s
    Else
        s = ""
        s = s & "update variables set" & vbCrLf
        s = s & "   Name=" & DbQuote(Str, txtName.Text) & vbCrLf
        s = s & "  ,VariableType=" & DbQuote(Num, IIf(optVariableType(0).value, 0, 1)) & vbCrLf
        s = s & "  ,MinimumValue=" & DbQuote(Num, txtMin.Text) & vbCrLf
        s = s & "  ,MaximumValue=" & DbQuote(Num, txtMax.Text) & vbCrLf
        s = s & "  ,DefaultValue=" & DbQuote(Num, txtDefault.Text) & vbCrLf
        s = s & "  ,ListOfValues=" & DbQuote(Str, values) & vbCrLf
        
        s = s & "  ,UOM=" & DbQuote(Str, cboUOM.Text) & vbCrLf
        s = s & "  ,Help=" & DbQuote(Str, txtHelp.Text) & vbCrLf
        s = s & "  ,conditionname=" & DbQuote(Str, mConditionName) & vbCrLf
        s = s & "  ,conditiontype=" & DbQuote(Str, mConditionType) & vbCrLf
        s = s & "  ,conditionuom=" & DbQuote(Str, mConditionUOM) & vbCrLf
        s = s & "where name=" & DbQuote(Str, mName) & vbCrLf
        HFApp.SqlExec s
    End If
    
    mCancel = False
    SaveData = True
    
End Function


Private Sub txtName_GotFocus()
    SelectAll txtName
End Sub
Private Sub txtMin_GotFocus()
    SelectAll txtMin
End Sub
Private Sub txtMax_GotFocus()
    SelectAll txtMax
End Sub
Private Sub txtDefault_GotFocus()
    SelectAll txtDefault
End Sub
Private Sub txtListOfValues_GotFocus()
'    SelectAll txtListOfValues
End Sub
Private Sub txtHelp_GotFocus()
'    SelectAll txtHelp
End Sub


Private Sub txtName_Change()
    Dim s As String
    Dim i As Long
    mDirty = True
    
    'replace curly quotes with normal ascii double quote char
    'replace double quote with single quote
    
    s = txtName.Text
    i = txtName.SelStart
    If InStr(1, s, Chr(147)) Or InStr(1, s, Chr(148)) Or InStr(1, s, Chr(152)) Or InStr(1, s, Chr(34)) Then
        s = Replace(s, Chr(150), "-")
        s = Replace(s, Chr(147), Chr(34))
        s = Replace(s, Chr(148), Chr(34))
        s = Replace(s, Chr(152), Chr(34))
        s = Replace(s, Chr(34), Chr(39))
        txtName.Text = s
        txtName.SelStart = i
    End If
End Sub
Private Sub txtMin_Change()
    mDirty = True
End Sub
Private Sub txtMax_Change()
    mDirty = True
End Sub
Private Sub txtDefault_Change()
    mDirty = True
End Sub
Private Sub txtHelp_Change()
    mDirty = True
End Sub
Private Sub txtListOfValues_Change()
    Dim s As String
    Dim i As Long
    mDirty = True
    
    'replace curly quotes with normal ascii double quote char
    'replace double quote with single quote
    
    s = txtListOfValues.Text
    i = txtListOfValues.SelStart
    If InStr(1, s, Chr(147)) Or InStr(1, s, Chr(148)) Or InStr(1, s, Chr(152)) Or InStr(1, s, Chr(34)) Then
        s = Replace(s, Chr(150), "-")
        s = Replace(s, Chr(147), Chr(34))
        s = Replace(s, Chr(148), Chr(34))
        s = Replace(s, Chr(152), Chr(34))
        s = Replace(s, Chr(34), Chr(39))
        txtListOfValues.Text = s
        txtListOfValues.SelStart = i
    End If
End Sub
Private Sub cboUOM_Change()
    mDirty = True
End Sub



