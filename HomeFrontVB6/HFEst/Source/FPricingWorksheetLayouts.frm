VERSION 5.00
Begin VB.Form FPricingWorksheetLayouts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Worksheet Layouts"
   ClientHeight    =   3195
   ClientLeft      =   3600
   ClientTop       =   1080
   ClientWidth     =   8160
   Icon            =   "FPricingWorksheetLayouts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtLayouts 
      Height          =   2175
      Left            =   300
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   300
      Width           =   3315
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Height          =   375
      Index           =   0
      Left            =   6780
      TabIndex        =   1
      Top             =   2220
      Width           =   1215
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   6780
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Enter one layout name per line."
      Height          =   675
      Index           =   3
      Left            =   3960
      TabIndex        =   5
      Top             =   1980
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   $"FPricingWorksheetLayouts.frx":000C
      Height          =   1095
      Index           =   2
      Left            =   3960
      TabIndex        =   4
      Top             =   1020
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   $"FPricingWorksheetLayouts.frx":00CD
      Height          =   675
      Index           =   1
      Left            =   3960
      TabIndex        =   3
      Top             =   300
      Width           =   3615
   End
End
Attribute VB_Name = "FPricingWorksheetLayouts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mCancel As Boolean
Private mLayouts As String

Public Function ShowForm(Layouts As String) As Boolean
    mCancel = False
    mLayouts = Layouts
    txtLayouts.Text = mLayouts
    Me.Show vbModal
    Layouts = mLayouts
    ShowForm = Not mCancel
End Function

Private Sub cmdNav_Click(Index As Integer)
    mCancel = Index = 1
    Unload Me
End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub

Private Sub txtLayouts_Change()
    mLayouts = txtLayouts.Text
End Sub

Private Sub txtLayouts_GotFocus()
    SelectAll txtLayouts
End Sub
