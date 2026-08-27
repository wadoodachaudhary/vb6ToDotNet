VERSION 5.00
Begin VB.Form FPOGenerationOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generation Options"
   ClientHeight    =   2595
   ClientLeft      =   6405
   ClientTop       =   5715
   ClientWidth     =   4815
   Icon            =   "FPOGenerationOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Option1 
      Caption         =   "One PO per Unit"
      Height          =   300
      Index           =   2
      Left            =   1230
      TabIndex        =   4
      Top             =   1455
      Width           =   2700
   End
   Begin VB.OptionButton Option1 
      Caption         =   "One PO per Job"
      Height          =   300
      Index           =   1
      Left            =   1230
      TabIndex        =   3
      Top             =   1080
      Value           =   -1  'True
      Width           =   2700
   End
   Begin VB.OptionButton Option1 
      Caption         =   "One PO per Vendor"
      Height          =   315
      Index           =   0
      Left            =   1230
      TabIndex        =   2
      Top             =   675
      Width           =   2700
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   3375
      Picture         =   "FPOGenerationOptions.frx":000C
      TabIndex        =   1
      ToolTipText     =   "Login"
      Top             =   2055
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   150
      Picture         =   "FPOGenerationOptions.frx":0596
      Top             =   210
      Width           =   480
   End
   Begin VB.Label lblWarning 
      AutoSize        =   -1  'True
      Caption         =   "How do you want to generate these POs?"
      Height          =   195
      Index           =   1
      Left            =   855
      TabIndex        =   0
      Top             =   300
      Width           =   2970
   End
End
Attribute VB_Name = "FPOGenerationOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mPer As String

Public Function Choose() As String
    
    Load Me
    Me.cmdNav(0).Enabled = True
    Me.Show vbModal
    
    Choose = mPer
    
End Function

Private Sub cmdNav_Click(Index As Integer)
    
    If Me.Option1(0).value Then mPer = "Vendor"
    If Me.Option1(1).value Then mPer = "Job"
    If Me.Option1(2).value Then mPer = "Unit"
    
    Unload Me
End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub
