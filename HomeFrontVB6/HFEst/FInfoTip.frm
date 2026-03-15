VERSION 5.00
Begin VB.Form FInfoTip 
   BackColor       =   &H80000018&
   Caption         =   "no caption"
   ClientHeight    =   1950
   ClientLeft      =   6150
   ClientTop       =   4800
   ClientWidth     =   6585
   Icon            =   "FInfoTip.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1950
   ScaleWidth      =   6585
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Label1"
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "FInfoTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Sub Showform(TipText As String)
    Me.Show
    lblText.Caption = " " & Trim(TipText) & " "
    Me.Height = lblText.Height + Me.Height - Me.ScaleHeight
    Me.Width = lblText.Width + Me.Width - Me.ScaleWidth
End Sub
