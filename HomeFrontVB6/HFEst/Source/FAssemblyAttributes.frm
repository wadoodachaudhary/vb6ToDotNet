VERSION 5.00
Begin VB.Form FAssemblyAttributes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Attributes"
   ClientHeight    =   1620
   ClientLeft      =   3090
   ClientTop       =   2100
   ClientWidth     =   6795
   Icon            =   "FAssemblyAttributes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   Begin HFEst.VBCombo cboList 
      Height          =   240
      Index           =   0
      Left            =   870
      TabIndex        =   0
      Top             =   330
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   423
      Style           =   2
      DropDownWidth   =   8800
   End
   Begin HFEst.VBCombo cboList 
      Height          =   240
      Index           =   1
      Left            =   870
      TabIndex        =   2
      Top             =   585
      Width           =   2625
      _ExtentX        =   2619
      _ExtentY        =   423
      Style           =   2
      DropDownWidth   =   8800
   End
   Begin HFEst.VBCombo cboList 
      Height          =   240
      Index           =   2
      Left            =   870
      TabIndex        =   4
      Top             =   840
      Width           =   2625
      _ExtentX        =   2619
      _ExtentY        =   423
      Style           =   2
      DropDownWidth   =   8800
   End
   Begin HFEst.VBCombo cboList 
      Height          =   240
      Index           =   3
      Left            =   870
      TabIndex        =   6
      Top             =   1095
      Width           =   2625
      _ExtentX        =   2619
      _ExtentY        =   423
      Style           =   2
      DropDownWidth   =   8800
   End
   Begin HFEst.VBCombo cboValue 
      Height          =   240
      Index           =   0
      Left            =   3900
      TabIndex        =   1
      Top             =   330
      Width           =   2625
      _ExtentX        =   2619
      _ExtentY        =   423
      DropDownWidth   =   8800
   End
   Begin HFEst.VBCombo cboValue 
      Height          =   240
      Index           =   1
      Left            =   3900
      TabIndex        =   3
      Top             =   585
      Width           =   2625
      _ExtentX        =   2619
      _ExtentY        =   423
      DropDownWidth   =   8800
   End
   Begin HFEst.VBCombo cboValue 
      Height          =   240
      Index           =   2
      Left            =   3900
      TabIndex        =   5
      Top             =   840
      Width           =   2625
      _ExtentX        =   2619
      _ExtentY        =   423
      DropDownWidth   =   8800
   End
   Begin HFEst.VBCombo cboValue 
      Height          =   240
      Index           =   3
      Left            =   3900
      TabIndex        =   7
      Top             =   1095
      Width           =   2625
      _ExtentX        =   2619
      _ExtentY        =   423
      DropDownWidth   =   8800
   End
   Begin VB.Image cmdList 
      Height          =   240
      Index           =   3
      Left            =   3540
      Picture         =   "FAssemblyAttributes.frx":000C
      Top             =   1095
      Width           =   240
   End
   Begin VB.Image cmdList 
      Height          =   240
      Index           =   2
      Left            =   3540
      Picture         =   "FAssemblyAttributes.frx":0596
      Top             =   840
      Width           =   240
   End
   Begin VB.Image cmdList 
      Height          =   240
      Index           =   1
      Left            =   3540
      Picture         =   "FAssemblyAttributes.frx":0B20
      Top             =   585
      Width           =   240
   End
   Begin VB.Label Label111 
      AutoSize        =   -1  'True
      Caption         =   "Default Value"
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
      Left            =   3960
      TabIndex        =   13
      Top             =   90
      Width           =   1170
   End
   Begin VB.Image cmdList 
      Height          =   240
      Index           =   0
      Left            =   3540
      Picture         =   "FAssemblyAttributes.frx":10AA
      Top             =   330
      Width           =   240
   End
   Begin VB.Label Label111 
      AutoSize        =   -1  'True
      Caption         =   "Attribute List"
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
      Left            =   930
      TabIndex        =   12
      Top             =   90
      Width           =   1095
   End
   Begin VB.Label Label111 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Other"
      Height          =   195
      Index           =   3
      Left            =   420
      TabIndex        =   11
      Top             =   1118
      Width           =   390
   End
   Begin VB.Label Label111 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Finish"
      Height          =   195
      Index           =   2
      Left            =   405
      TabIndex        =   10
      Top             =   863
      Width           =   405
   End
   Begin VB.Label Label111 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Style"
      Height          =   195
      Index           =   1
      Left            =   465
      TabIndex        =   9
      Top             =   608
      Width           =   345
   End
   Begin VB.Label Label111 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Color"
      Height          =   195
      Index           =   0
      Left            =   450
      TabIndex        =   8
      Top             =   353
      Width           =   360
   End
End
Attribute VB_Name = "FAssemblyAttributes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub EditAttributes(f As FAssembly)
    Load Me
    Call IniGetForm(Me)
    
    Call LoadComboBox(cboList(0), HFApp.Databases(dbHomefront), "select name,'',listid from attributelists where isnull(inactive,0)=0 order by 1")
    Call LoadComboBox(cboList(1), HFApp.Databases(dbHomefront), "select name,'',listid from attributelists where isnull(inactive,0)=0 order by 1")
    Call LoadComboBox(cboList(2), HFApp.Databases(dbHomefront), "select name,'',listid from attributelists where isnull(inactive,0)=0 order by 1")
    Call LoadComboBox(cboList(3), HFApp.Databases(dbHomefront), "select name,'',listid from attributelists where isnull(inactive,0)=0 order by 1")
    
    Call SetComboBoxListIndex(cboList(0), , , Val(f.txtLColor.Text))
    Call SetComboBoxListIndex(cboList(1), , , Val(f.txtLStyle.Text))
    Call SetComboBoxListIndex(cboList(2), , , Val(f.txtLFinish.Text))
    Call SetComboBoxListIndex(cboList(3), , , Val(f.txtLOther.Text))
    
    Call cboList_Click(0)
    Call cboList_Click(1)
    Call cboList_Click(2)
    Call cboList_Click(3)
    cboValue(0).Text = f.txtAColor
    cboValue(1).Text = f.txtAStyle
    cboValue(2).Text = f.txtAFinish
    cboValue(3).Text = f.txtAOther

    Me.Show vbModal
    f.txtLColor = GetComboBoxListID(cboList(0))
    f.txtLStyle = GetComboBoxListID(cboList(1))
    f.txtLFinish = GetComboBoxListID(cboList(2))
    f.txtLOther = GetComboBoxListID(cboList(3))
    f.txtAColor = cboValue(0).Text
    f.txtAStyle = cboValue(1).Text
    f.txtAFinish = cboValue(2).Text
    f.txtAOther = cboValue(3).Text
    
    Unload Me

End Sub

Private Sub cboList_Click(Index As Integer)
    Dim b As Boolean
    
    On Error Resume Next
    b = "" & HFApp.SqlExec("select strictlist from attributelists where listid=" & GetComboBoxListID(cboList(Index)))(0) = "True"
    On Error GoTo 0
    
    'cboValue(Index).AutoCompleteListItemsOnly = b
    Call LoadComboBox(cboValue(Index), HFApp.Databases(dbHomefront), "select value,'',0 from attributelistvalues where isnull(inactive,0)=0  and listid=" & GetComboBoxListID(cboList(Index)) & " order by sortorder")
End Sub

Private Sub cboList_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Call SetComboBoxListIndex(cboList(Index), , , -1)
        cboValue(Index).Text = ""
    End If
End Sub

Private Sub cmdList_Click(Index As Integer)
    Dim i As Long
    Dim X As Integer
    Dim List As Long
    Dim value As String
    
    X = GetComboBoxListID(cboList(Index))
    Call HFApp.RunTask("EditAttributeLists|" & X)
    
    
    For X = 0 To 3
        List = GetComboBoxListID(cboList(X))
        value = cboValue(X).Text
        Call LoadComboBox(cboList(X), HFApp.Databases(dbHomefront), "select name,'',listid from attributelists order by 1")
        Call SetComboBoxListIndex(cboList(X), , , List)
        Call cboList_Click(X)
        cboValue(X).Text = value
    Next

End Sub

Private Sub Form_Load()
    Call LoadCustomDescriptions
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Cancel = True
        Me.Hide
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub

Private Sub LoadCustomDescriptions()
    Label111(0).Caption = FMain.CD_Color
End Sub
