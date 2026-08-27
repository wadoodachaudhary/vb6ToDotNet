VERSION 5.00
Begin VB.Form FPhase 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estimating Phase"
   ClientHeight    =   2070
   ClientLeft      =   180
   ClientTop       =   2430
   ClientWidth     =   5340
   Icon            =   "FPhase.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPhase 
      Height          =   285
      Left            =   930
      MaxLength       =   100
      TabIndex        =   0
      Top             =   1350
      Width           =   4305
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Left            =   930
      MaxLength       =   100
      TabIndex        =   1
      Top             =   1680
      Width           =   4305
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   4260
      TabIndex        =   2
      Top             =   150
      Width           =   945
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   345
      Index           =   1
      Left            =   4260
      TabIndex        =   3
      Top             =   570
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter a number and description for the new phase."
      Height          =   195
      Index           =   0
      Left            =   210
      TabIndex        =   6
      Top             =   540
      Width           =   3570
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Number"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   330
      TabIndex        =   5
      Top             =   1380
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Description"
      Height          =   195
      Index           =   28
      Left            =   90
      TabIndex        =   4
      Top             =   1710
      Width           =   795
   End
End
Attribute VB_Name = "FPhase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const SRCFILE = "FPhase::"
Private mCancel   As Boolean
Private mGrpPhase As Boolean

Public Function ShowForm(GrpPhase As Boolean) As Boolean
On Error Resume Next
    mCancel = False
    mGrpPhase = GrpPhase
    Me.Show vbModal
    ShowForm = Not mCancel
End Function




Private Sub cmdNav_Click(Index As Integer)
    Select Case Index
        Case 0
            If SaveData Then
                mCancel = False
                Unload Me
            End If
        Case 1
            mCancel = True
            Unload Me
    End Select
End Sub


Private Sub Form_Load()
    Me.Caption = App.ProductName
    Call IniGetForm(Me)
    txtPhase.MaxLength = HFApp.EstFieldSize("Phase")
    txtDescription.MaxLength = HFApp.EstFieldSize("PhaseDesc")
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub


Private Sub txtPhase_GotFocus()
    SelectAll txtPhase
End Sub
Private Sub txtDescription_GotFocus()
    SelectAll txtDescription
End Sub



Public Function SaveData() As Boolean
On Error GoTo eh:

    Dim rc As Long
    Dim s As String
    Dim i As Long
    
    
    If Trim(txtPhase.Text) = "" Or Trim(txtDescription.Text) = "" Then
        MsgBox "Number and description are both required", vbExclamation, App.ProductName
        Exit Function
    End If
    
    
    
    Screen.MousePointer = vbHourglass
    
    
    s = ""
    s = s & "INSERT INTO tblEstPhases(DivisionID,Phase,Description,GroupPhase,UpdateEstimating)" & vbCrLf
    s = s & "VALUES(" & HFApp.DivisionID & "," & DbQuote(Str, txtPhase.Text) & vbCrLf
    s = s & "      ," & DbQuote(Str, txtDescription.Text) & vbCrLf
    s = s & "      ," & DbQuote(Bit, mGrpPhase) & vbCrLf
    s = s & "      ,0)"
    Call HFApp.SqlExec(s)
    
    Call FixGroupPhaseValue
    
  
    SaveData = True
    Screen.MousePointer = vbDefault
    Exit Function
    
    
eh: Select Case True
    Case InStr(1, Err.Description, "duplicate") > 0:
        Screen.MousePointer = vbDefault
        MsgBox "Phase """ & Trim(txtPhase.Text) & """ has already been entered. Please specify a different value.", vbInformation, App.ProductName
    Case Else:  Call errHandler(SRCFILE & "SaveData", s)
    End Select
End Function
