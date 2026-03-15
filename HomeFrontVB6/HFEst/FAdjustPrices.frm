VERSION 5.00
Begin VB.Form FAdjustPrices 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Adjust Selling Prices"
   ClientHeight    =   3195
   ClientLeft      =   1065
   ClientTop       =   5835
   ClientWidth     =   6030
   Icon            =   "FAdjustPrices.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtValue 
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   3
      Left            =   1860
      MaxLength       =   50
      TabIndex        =   3
      Top             =   2100
      Width           =   1275
   End
   Begin VB.TextBox txtValue 
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   2
      Left            =   1860
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1800
      Width           =   1275
   End
   Begin VB.TextBox txtValue 
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   1
      Left            =   1860
      MaxLength       =   50
      TabIndex        =   1
      Top             =   840
      Width           =   675
   End
   Begin VB.TextBox txtValue 
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   0
      Left            =   1860
      MaxLength       =   50
      TabIndex        =   0
      Top             =   540
      Width           =   675
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   4740
      TabIndex        =   5
      Top             =   540
      Width           =   1215
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Height          =   375
      Index           =   0
      Left            =   4740
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Price Decrease"
      Height          =   195
      Index           =   3
      Left            =   720
      TabIndex        =   11
      Top             =   2100
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Price Increase"
      Height          =   195
      Index           =   2
      Left            =   795
      TabIndex        =   10
      Top             =   1800
      Width           =   1020
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Percent Decrease"
      Height          =   195
      Index           =   1
      Left            =   525
      TabIndex        =   9
      Top             =   840
      Width           =   1290
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Percent Change"
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
      Left            =   300
      TabIndex        =   8
      Top             =   240
      Width           =   1380
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   3
      X1              =   780
      X2              =   4260
      Y1              =   375
      Y2              =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   2
      X1              =   780
      X2              =   4260
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Dollar Value "
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
      Left            =   300
      TabIndex        =   7
      Top             =   1500
      Width           =   1110
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   780
      X2              =   4260
      Y1              =   1635
      Y2              =   1635
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   780
      X2              =   4260
      Y1              =   1620
      Y2              =   1620
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Percent Increase"
      Height          =   195
      Index           =   0
      Left            =   600
      TabIndex        =   6
      Top             =   540
      Width           =   1215
   End
End
Attribute VB_Name = "FAdjustPrices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mCancel As Boolean
Private mPercentChange As Double
Private mDollarChange As Double



Public Function ShowForm(PercentChange As Double, DollarChange As Double) As Boolean
    mCancel = False
    Me.Show vbModal
    ShowForm = (mPercentChange <> 0 Or mDollarChange <> 0) And Not mCancel
    PercentChange = mPercentChange
    DollarChange = IIf(mPercentChange = 0, mDollarChange, 0)
End Function


Private Sub cmdNav_Click(Index As Integer)
    Dim i As Long
    If Index = 0 Then
        mCancel = False
        
        mPercentChange = 0
        mDollarChange = 0
        
        If txtValue(0).Text <> "" Then mPercentChange = Abs(Val(txtValue(0).Text))
        If txtValue(1).Text <> "" Then mPercentChange = -1 * Abs(Val(txtValue(1).Text))
        If txtValue(2).Text <> "" Then mDollarChange = Abs(Val(Replace(Replace(txtValue(2).Text, ",", ""), "$", "")))
        If txtValue(3).Text <> "" Then mDollarChange = -1 * Abs(Val(Replace(Replace(txtValue(3).Text, ",", ""), "$", "")))
    Else
        mCancel = True
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub

Private Sub txtValue_Change(Index As Integer)
Static inHere As Boolean
If inHere Then Exit Sub
inHere = True
    Dim i As Integer
    For i = 0 To 3
        If i <> Index Then txtValue(i).Text = ""
    Next
inHere = False
End Sub

Private Sub txtValue_GotFocus(Index As Integer)
    SelectAll txtValue(Index)
End Sub

Private Sub txtValue_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = vbKeyReturn Then
        If txtValue(Index).Text <> "" Then
            cmdNav(0).SetFocus
        Else
            Call txtValue(Index + 1).SetFocus
        End If
    End If
End Sub

Private Sub txtValue_Validate(Index As Integer, Cancel As Boolean)
    If Trim(txtValue(Index).Text) <> "" Then
        If IsNumeric(txtValue(Index).Text) Then
            Cancel = False
            If Index < 2 Then
                'format percent
                txtValue(Index).Text = Abs(Val(Replace(txtValue(Index).Text, "%", ""))) & "%"
            Else
                'format dollars
                txtValue(Index).Text = format(Abs(Val(Replace(Replace(txtValue(Index).Text, ",", ""), "$", ""))), "#,##0.00")
            End If
        Else
            Cancel = True
        End If
    End If
End Sub
