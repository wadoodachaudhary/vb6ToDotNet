VERSION 5.00
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Begin VB.Form FProgress 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Processing..."
   ClientHeight    =   1065
   ClientLeft      =   2325
   ClientTop       =   3855
   ClientWidth     =   3795
   ControlBox      =   0   'False
   Icon            =   "FProgress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   3795
   ShowInTaskbar   =   0   'False
   Begin vbalProgBarLib6.vbalProgressBar ProgressBar 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   660
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   450
      Picture         =   "FProgress.frx":000C
      ForeColor       =   0
      BarPicture      =   "FProgress.frx":0028
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      XpStyle         =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   360
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "FProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const SRCFILE = "FProgress"

Public Sub Progress(Optional Title As String, _
                    Optional Caption As String, _
                    Optional Row As Long, _
                    Optional count As Long, _
                    Optional Parent As Form)
On Error Resume Next


    Me.Show , IIf(Parent Is Nothing, FMain, Parent)
    
    DoEvents
    
    If Title <> "" Then Me.Caption = Title
    If Caption <> "" Then Label1.Caption = Caption
    Label2.Caption = "row " & Row & " of " & count
    Label1.Refresh
    Label2.Refresh
    If count = 0 Then
        ProgressBar.value = count
    Else
        ProgressBar.value = Row / count * 100
    End If
    
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
End Sub
