VERSION 5.00
Begin VB.Form FList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4410
   ClientLeft      =   2610
   ClientTop       =   1695
   ClientWidth     =   7920
   Icon            =   "FList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Height          =   375
      Index           =   0
      Left            =   6450
      TabIndex        =   1
      Top             =   390
      Width           =   1215
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   6450
      TabIndex        =   3
      Top             =   810
      Width           =   1215
   End
   Begin VB.TextBox txtList 
      Height          =   3855
      Left            =   195
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "FList.frx":000C
      Top             =   360
      Width           =   6135
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Title"
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
      Left            =   150
      TabIndex        =   2
      Top             =   75
      Width           =   390
   End
End
Attribute VB_Name = "FList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FList::"

Private mCancel As Boolean
Private mCaption As String
Private mTitle As String
Private mList As String

Public Function Choose(ByVal FormCaption As String, ByVal Title As String, ByRef List As String, Optional Delimiter As String = ",") As Boolean
    mCaption = FormCaption
    mTitle = Title
    mList = Replace(List, Delimiter, vbCrLf)
    
    Me.Show vbModal

    If Not mCancel Then
        'remove dups
        mList = Replace(mList, vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf, vbCrLf)
        mList = Replace(mList, vbCrLf & vbCrLf & vbCrLf & vbCrLf, vbCrLf)
        mList = Replace(mList, vbCrLf & vbCrLf & vbCrLf, vbCrLf)
        mList = Replace(mList, vbCrLf & vbCrLf, vbCrLf)
        'return new list
        List = Replace(mList, vbCrLf, Delimiter)
    End If
    Choose = Not mCancel
    
End Function

Private Sub cmdNav_Click(Index As Integer)
On Error Resume Next
    
    Select Case Index
        
        Case 0 'OK
            mList = txtList.Text
            mCancel = False
            
        Case 1 'Cancel
            mCancel = True
            
    End Select
    Unload Me
    
Exit Sub
End Sub


Private Sub Form_Load()
    IniGetForm Me, , mTitle
    
    Me.Caption = mCaption
    lblTitle.Caption = mTitle
    txtList.Text = mList
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    IniPutForm Me, , mTitle
End Sub
