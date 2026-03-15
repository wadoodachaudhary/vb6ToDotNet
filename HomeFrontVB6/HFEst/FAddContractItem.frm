VERSION 5.00
Begin VB.Form FAddContractItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Item"
   ClientHeight    =   3300
   ClientLeft      =   390
   ClientTop       =   6630
   ClientWidth     =   6330
   Icon            =   "FAddContractItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboChangeOrder 
      Height          =   240
      Left            =   1800
      TabIndex        =   2
      Top             =   1965
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ExtendedUI      =   0   'False
      DropDownWidth   =   0
   End
   Begin VB.TextBox txtHFDescription 
      BorderStyle     =   0  'None
      Height          =   230
      Left            =   1800
      MaxLength       =   200
      TabIndex        =   0
      Text            =   "New Item"
      Top             =   240
      Width           =   4155
   End
   Begin VB.TextBox txtHFComments 
      BorderStyle     =   0  'None
      Height          =   1470
      Left            =   1800
      MaxLength       =   4000
      TabIndex        =   1
      Top             =   480
      Width           =   4155
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   4980
      TabIndex        =   5
      Top             =   2820
      Width           =   1215
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Height          =   375
      Index           =   0
      Left            =   3660
      TabIndex        =   4
      Top             =   2820
      Width           =   1215
   End
   Begin VB.ComboBox cboJCExtra 
      Height          =   240
      Left            =   1800
      TabIndex        =   3
      Top             =   2220
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ExtendedUI      =   0   'False
      DropDownWidth   =   0
   End
   Begin VB.Label lblJCExtra 
      Alignment       =   1  'Right Justify
      Caption         =   "JC Extra"
      Height          =   255
      Left            =   540
      TabIndex        =   9
      Top             =   2250
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Description"
      Height          =   240
      Index           =   1
      Left            =   720
      TabIndex        =   8
      Top             =   270
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Comments"
      Height          =   240
      Index           =   2
      Left            =   720
      TabIndex        =   7
      Top             =   540
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "FAddContractItem.frx":000C
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Change Order"
      Height          =   195
      Index           =   0
      Left            =   705
      TabIndex        =   6
      Top             =   1980
      Width           =   990
   End
End
Attribute VB_Name = "FAddContractItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mCustomer As String
Private mCancel As Boolean

Private Function SaveData() As Boolean
    Dim s As String

    If Trim(txtHFDescription.Text) = "" Then
        MsgBox "You must specify a description", vbExclamation, App.ProductName
        SaveData = False
        Exit Function
    End If

    s = ""
    s = s & "INSERT INTO EstimateAssemblies(Customer_No,EstimateIndex,AssemblyType,OptionType,SalesQty,Job,HFComments,HFDescription,ChangeOrder,JCExtra)" & vbCrLf
    s = s & "SELECT Customer_No" & vbCrLf
    s = s & "      ,LastEstimateIndex+1" & vbCrLf
    s = s & "      ,-1" & vbCrLf
    s = s & "      ,-1" & vbCrLf
    s = s & "      ,1" & vbCrLf
    s = s & "      ,Job_No" & vbCrLf
    s = s & "      ," & DbQuote(Str, txtHFComments.Text, , True) & vbCrLf
    s = s & "      ," & DbQuote(Str, txtHFDescription.Text, , True) & vbCrLf
    s = s & "      ," & DbQuote(Str, cboChangeOrder.Text, , True, 10) & vbCrLf
    s = s & "      ," & DbQuote(Str, cboJCExtra.Text, , True, 10) & vbCrLf
    s = s & "  FROM tblCustomers" & vbCrLf
    s = s & " WHERE Customer_No=" & DbQuote(Str, mCustomer)
    Call HFApp.SqlExec(s)
    
    s = ""
    s = s & "UPDATE tblCustomers" & vbCrLf
    s = s & "   SET LastEstimateIndex=LastEstimateIndex+1" & vbCrLf
    s = s & " WHERE Customer_No=" & DbQuote(Str, mCustomer) & vbCrLf
    Call HFApp.SqlExec(s)
    
    SaveData = True
    
End Function

Public Function AddItem(Customer As String, ChangeOrder As String) As Boolean
    Dim job As String
    
    
    
    mCustomer = Customer
    
    On Error Resume Next
    Call LoadComboBox(cboChangeOrder, HFApp.Databases(dbHomeFront), "SELECT DISTINCT ISNULL(ChangeOrder,''),'',0  FROM EstimateAssemblies WHERE Job IN(SELECT Job_No FROM tblCustomers WHERE Customer_No=" & DbQuote(Str, mCustomer) & ")")
    cboChangeOrder.Text = ChangeOrder
    Call LoadComboBox(cboJCExtra, HFApp.Databases(dbHomeFront), "SELECT DISTINCT ISNULL(JCExtra,''),'',0 FROM EstimateAssemblies WHERE Job IN(SELECT Job_No FROM tblCustomers WHERE Customer_No=" & DbQuote(Str, mCustomer) & ")")
    cboJCExtra.Text = HFApp.SqlExec("SELECT JCExtra FROM EstimateAssemblies WHERE Customer_No=" & DbQuote(Str, mCustomer) & " AND ChangeOrder=" & DbQuote(Str, ChangeOrder) & " ORDER BY EstAssemblyID DESC", dbHomeFront)(0)

    
    mCancel = False
    Me.Show vbModal
    AddItem = Not mCancel
End Function

Private Sub cboChangeOrder_Validate(Cancel As Boolean)
    cboChangeOrder.Text = left(cboChangeOrder.Text, 10)
End Sub

Private Sub cboJCExtra_Validate(Cancel As Boolean)
    cboJCExtra.Text = left(cboJCExtra.Text, 10)
End Sub

Private Sub cmdNav_Click(Index As Integer)
    If Index = 0 Then If Not SaveData() Then Exit Sub
    mCancel = Index = 1
    Unload Me
End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
    Select Case HFApp.Options(AccountingSystem)
        Case asMasterBuilder:  lblJCExtra.Caption = "Phase"
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub

Private Sub txtHFDescription_GotFocus()
    SelectAll txtHFDescription
End Sub
Private Sub txtHFComments_GotFocus()
    SelectAll txtHFComments
End Sub

