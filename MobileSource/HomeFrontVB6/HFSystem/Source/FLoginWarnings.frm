VERSION 5.00
Begin VB.Form FLoginWarnings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HomeFront"
   ClientHeight    =   5955
   ClientLeft      =   4650
   ClientTop       =   2310
   ClientWidth     =   9285
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      Caption         =   "Dont show me this again"
      Height          =   255
      Left            =   150
      TabIndex        =   4
      Top             =   5580
      Width           =   4155
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   7980
      Picture         =   "FLoginWarnings.frx":0000
      TabIndex        =   1
      ToolTipText     =   "Login"
      Top             =   5460
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5325
      Left            =   0
      ScaleHeight     =   5325
      ScaleWidth      =   9285
      TabIndex        =   0
      Top             =   0
      Width           =   9285
      Begin VB.Image Image1 
         Height          =   480
         Left            =   390
         Picture         =   "FLoginWarnings.frx":058A
         Top             =   330
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Errors...."
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   3
         Top             =   990
         Width           =   585
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "HomeFront is unable to connect to the following network resources. Please contact your system administrator for assistance."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   0
         Left            =   1200
         TabIndex        =   2
         Top             =   300
         Width           =   6045
      End
   End
End
Attribute VB_Name = "FLoginWarnings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mMessage As String

Public Sub ShowMessage(s As String)
    
    Dim f As String
    f = PathAppend(AppWorkingFolder, "warnings.ini")
    
    If IniGet(f, "Options", "SuppressConnectionWarnings", "False") <> "False" Then Exit Sub
    
    mMessage = s
    Me.Show vbModal
    
End Sub

Private Sub cmdNav_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    
    Call IniGetForm(Me)
    Me.Label1(1).Caption = mMessage
        
    Me.Width = Label1(1).left + Max(Label1(0).Width, Label1(1).Width) + 600 + Me.Width - Me.ScaleWidth
    Me.Height = Label1(1).Top + Label1(1).Height + 800 + Me.Height - Me.ScaleHeight
    Me.Picture1.Height = Me.ScaleHeight - 660
    Call WindowOnTop(Me, True)
    
End Sub

Private Sub Form_Resize()
    Const margin = 120
    cmdNav(0).Move Me.ScaleWidth - cmdNav(0).Width - margin, Me.ScaleHeight - cmdNav(0).Height - margin
    Check1.Move margin, Me.ScaleHeight - Check1.Height - margin
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim f As String
    f = PathAppend(AppWorkingFolder, "warnings.ini")
    Call IniPut(f, "Options", "SuppressConnectionWarnings", Check1.Value = vbChecked)
    Call IniPutForm(Me)
End Sub
