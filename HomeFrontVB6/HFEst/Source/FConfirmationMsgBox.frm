VERSION 5.00
Begin VB.Form FConfirmationMsgBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WARNING!!"
   ClientHeight    =   2835
   ClientLeft      =   2325
   ClientTop       =   4395
   ClientWidth     =   5985
   Icon            =   "FConfirmationMsgBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   3300
      Picture         =   "FConfirmationMsgBox.frx":000C
      TabIndex        =   1
      ToolTipText     =   "Login"
      Top             =   2370
      Width           =   1215
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   4605
      Picture         =   "FConfirmationMsgBox.frx":0596
      TabIndex        =   2
      ToolTipText     =   "Cancel"
      Top             =   2370
      Width           =   1215
   End
   Begin VB.CheckBox chkConfirm 
      Caption         =   "I understand and wish to proceed."
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Label lblWarning 
      Caption         =   $"FConfirmationMsgBox.frx":0B20
      Height          =   1335
      Index           =   1
      Left            =   960
      TabIndex        =   4
      Top             =   450
      Width           =   4815
   End
   Begin VB.Label lblWarning 
      AutoSize        =   -1  'True
      Caption         =   "Warning!!"
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
      Index           =   0
      Left            =   960
      TabIndex        =   3
      Top             =   210
      Width           =   840
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "FConfirmationMsgBox.frx":0CCE
      Top             =   270
      Width           =   480
   End
End
Attribute VB_Name = "FConfirmationMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mCancel As Boolean

Public Function ShowWarning(Optional Warning As String = "Whoa! Thats scary! Are you sure?", Optional WindowCaption As String = "Warning!", Optional Title As String = "Warning!", Optional Confirmation As String = "I understand and wish to proceed.") As Boolean
    ' returns true if they clicked OK, false if the click cancel
    
    Load Me
    mCancel = True
    Me.Caption = WindowCaption
    Me.lblWarning(0).Caption = Title
    Me.lblWarning(1).Caption = Warning
    Me.chkConfirm.Caption = Confirmation
    Me.chkConfirm.Value = vbUnchecked
    Me.cmdNav(0).Enabled = False
    Me.Show vbModal
    
    ShowWarning = Not mCancel
    
End Function

Private Sub chkConfirm_Click()
    cmdNav(0).Enabled = chkConfirm.Value = vbChecked
End Sub

Private Sub cmdNav_Click(Index As Integer)
    mCancel = Index = 1
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub
