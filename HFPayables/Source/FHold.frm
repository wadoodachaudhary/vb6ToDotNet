VERSION 5.00
Begin VB.Form FHold 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hold Invoice"
   ClientHeight    =   2460
   ClientLeft      =   2835
   ClientTop       =   1695
   ClientWidth     =   5160
   Icon            =   "FHold.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2520
      TabIndex        =   3
      Top             =   1980
      Width           =   1215
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3840
      TabIndex        =   2
      Top             =   1980
      Width           =   1215
   End
   Begin VB.TextBox txtComment 
      Height          =   1335
      IMEMode         =   3  'DISABLE
      Left            =   1020
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   540
      Width           =   4035
   End
   Begin VB.ComboBox cboApprover 
      Height          =   315
      Left            =   1020
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   180
      Width           =   3315
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Hold For"
      Height          =   255
      Index           =   0
      Left            =   -60
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Reason"
      Height          =   255
      Index           =   1
      Left            =   -60
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "FHold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SRCFILE = "FHold::"

Dim mCancel   As Boolean
Dim mApprover As String
Dim mComment  As String

Public Function Save(Approver As String, Comment As String) As Boolean
On Error GoTo eh
    Dim s As String
    Dim rs As Recordset
    
    mComment = Comment
    txtComment.Text = mComment
    
    mApprover = Approver
    cboApprover.Clear
    Set rs = HFApp.SqlExec("SELECT UID FROM PDUsers", dbHomeFront)
    While Not rs.EOF
        cboApprover.AddItem "" & rs(0)
        If Approver = "" & rs(0) Then cboApprover.ListIndex = cboApprover.NewIndex
        rs.MoveNext
    Wend

    Me.Show vbModal
    If Not mCancel Then
        Comment = mComment
        Approver = mApprover
    End If
    Save = Not mCancel

Exit Function
eh: Call ErrHandler(SRCFILE & "cmdNav_Click")
End Function

Private Sub cmdNav_Click(Index As Integer)
On Error GoTo eh
    mCancel = IIf(Index = 0, False, True)
    mApprover = IIf(Index = 0, cboApprover.Text, "")
    mComment = IIf(Index = 0, txtComment.Text, "")
    Unload Me
Exit Sub
eh: Call ErrHandler(SRCFILE & "cmdNav_Click")
End Sub

