VERSION 5.00
Begin VB.Form FCancelPO 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancel Purchase Order"
   ClientHeight    =   3435
   ClientLeft      =   7920
   ClientTop       =   2565
   ClientWidth     =   5505
   Icon            =   "FCancelPO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   2790
      Picture         =   "FCancelPO.frx":000C
      TabIndex        =   3
      ToolTipText     =   "Login"
      Top             =   2910
      Width           =   1215
   End
   Begin VB.TextBox txtReason 
      Height          =   2025
      Left            =   810
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   720
      Width           =   4485
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   4095
      Picture         =   "FCancelPO.frx":0596
      TabIndex        =   0
      ToolTipText     =   "Cancel"
      Top             =   2910
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   150
      Picture         =   "FCancelPO.frx":0B20
      Top             =   210
      Width           =   480
   End
   Begin VB.Label lblWarning 
      AutoSize        =   -1  'True
      Caption         =   "What is the reason for cancelling the purchase order?"
      Height          =   195
      Index           =   1
      Left            =   840
      TabIndex        =   1
      Top             =   300
      Width           =   3795
   End
End
Attribute VB_Name = "FCancelPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mCancel As Boolean
Private mReason As String

Public Function CancelPO(ByRef reason As String) As Boolean
    ' returns true if they clicked OK, false if the click cancel
    
    mReason = reason
    Load Me
    mCancel = True
    Me.cmdNav(0).Enabled = True
    
    Me.Show vbModal
    
    reason = mReason
    CancelPO = Not mCancel
    
End Function

Private Sub cmdNav_Click(Index As Integer)
    mCancel = Index = 1
    mReason = txtReason.Text
    Unload Me
End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
    txtReason.Text = mReason
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub
