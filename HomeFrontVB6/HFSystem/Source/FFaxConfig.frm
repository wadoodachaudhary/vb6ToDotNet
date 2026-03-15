VERSION 5.00
Begin VB.Form FFaxConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Provider Configuration"
   ClientHeight    =   2550
   ClientLeft      =   3270
   ClientTop       =   3990
   ClientWidth     =   6210
   Icon            =   "FFaxConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtBillingCode 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1860
      TabIndex        =   2
      Top             =   1110
      Width           =   3675
   End
   Begin VB.TextBox txtCoverPage 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1875
      TabIndex        =   1
      Top             =   810
      Width           =   3675
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   4830
      TabIndex        =   4
      Top             =   2010
      Width           =   1215
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Height          =   375
      Index           =   0
      Left            =   3510
      TabIndex        =   3
      Top             =   2010
      Width           =   1215
   End
   Begin VB.TextBox txtServer 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1875
      TabIndex        =   0
      Top             =   510
      Width           =   3675
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Billing Code"
      Height          =   195
      Index           =   1
      Left            =   945
      TabIndex        =   8
      Top             =   1125
      Width           =   825
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Cover Page"
      Height          =   195
      Index           =   0
      Left            =   945
      TabIndex        =   7
      Top             =   825
      Width           =   840
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Server"
      Height          =   195
      Index           =   43
      Left            =   1320
      TabIndex        =   6
      Top             =   525
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Windows Fax Services "
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
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   150
      Width           =   1995
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   8
      X1              =   600
      X2              =   6000
      Y1              =   270
      Y2              =   270
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   9
      X1              =   600
      X2              =   6000
      Y1              =   285
      Y2              =   285
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   270
      Picture         =   "FFaxConfig.frx":000C
      Top             =   510
      Width           =   480
   End
End
Attribute VB_Name = "FFaxConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Edit(ProviderName As String)
    Me.Show vbModal
End Sub

Private Sub cmdNav_Click(Index As Integer)
    If Index = 0 Then SaveData
    Unload Me
End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
    Call LoadData
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub


Private Sub LoadData()
    txtServer.Text = HFApp.Options.ValueByName("FaxServer")
    txtCoverPage.Text = HFApp.Options.ValueByName("FaxCoverPage")
    txtBillingCode.Text = HFApp.Options.ValueByName("FaxBillingCode")
End Sub

Private Sub SaveData()
    HFApp.Options.ValueByName("FaxServer") = txtServer.Text
    HFApp.Options.ValueByName("FaxCoverPage") = txtCoverPage.Text
    HFApp.Options.ValueByName("FaxBillingCode") = txtBillingCode.Text
End Sub

Private Sub txtBillingCode_GotFocus()
    SelectAll txtBillingCode
End Sub

Private Sub txtCoverPage_GotFocus()
    SelectAll txtCoverPage
End Sub

Private Sub txtServer_GotFocus()
    SelectAll txtServer
End Sub
