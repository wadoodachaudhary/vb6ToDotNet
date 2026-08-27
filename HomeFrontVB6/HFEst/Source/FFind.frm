VERSION 5.00
Begin VB.Form FFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   1485
   ClientLeft      =   1215
   ClientTop       =   2205
   ClientWidth     =   5280
   Icon            =   "FFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Direction"
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   2340
      TabIndex        =   8
      Top             =   600
      Width           =   1635
      Begin VB.OptionButton optDown 
         Caption         =   "&Down"
         Height          =   315
         Left            =   780
         TabIndex        =   4
         Top             =   300
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optUp 
         Caption         =   "&Up"
         Height          =   315
         Left            =   180
         TabIndex        =   3
         Top             =   300
         Width           =   555
      End
   End
   Begin VB.CheckBox chkMatchWholeWord 
      Caption         =   "Match whole &word only"
      Enabled         =   0   'False
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   660
      Width           =   1995
   End
   Begin VB.CheckBox chkMatchCase 
      Caption         =   "Match &case"
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1155
      TabIndex        =   0
      Top             =   255
      Width           =   2775
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   330
      Index           =   1
      Left            =   4140
      TabIndex        =   6
      Top             =   540
      Width           =   1035
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "Find Next"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   330
      Index           =   0
      Left            =   4140
      TabIndex        =   5
      Top             =   120
      Width           =   1035
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000011&
      Height          =   225
      Left            =   1140
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "&Find what:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   915
   End
End
Attribute VB_Name = "FFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mGrid As VSFlexGrid
Private mIsGrouped As Boolean

Private mOldHighlight As Long

Public Sub ShowForm(g As VSFlexGrid, Optional IsGrouped As Boolean)
    mIsGrouped = IsGrouped
    Set mGrid = g
    mOldHighlight = mGrid.HighLight
    mGrid.HighLight = flexHighlightAlways
    
    Me.Caption = "Find"
    Me.Show vbModal
End Sub

Private Sub cmdNav_Click(Index As Integer)
    Select Case Index
        Case 0 'find next
            If mIsGrouped Then
                Call FindNextHierarchy
            Else
                Call FindNextSimple
            End If
            
        Case 1 'cancel
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    chkMatchCase.Value = IniGet(AppIni, Me.Name, "chkMatchCase", vbUnchecked)
    chkMatchWholeWord.Value = IniGet(AppIni, Me.Name, "chkMatchWholeWord", vbUnchecked)
    optDown.Value = IniGet(AppIni, Me.Name, "optDown", True) = "True"
    optUp.Value = Not optDown.Value
    txtFind.Text = IniGet(AppIni, Me.Name, "txtFind", "")
    Call IniGetForm(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mGrid.HighLight = mOldHighlight
    Call IniPut(AppIni, Me.Name, "chkMatchCase", chkMatchCase.Value)
    Call IniPut(AppIni, Me.Name, "chkMatchWholeWord", chkMatchWholeWord.Value)
    Call IniPut(AppIni, Me.Name, "optDown", optDown.Value)
    Call IniPut(AppIni, Me.Name, "txtFind", txtFind.Text)
    Call IniPutForm(Me)
End Sub

Private Sub txtFind_Change()
    cmdNav(0).Enabled = Trim(txtFind.Text) <> ""
End Sub

Private Sub FindNextSimple()
On Error GoTo eh

    Dim bCase As Boolean
    Dim sWord As String
    Dim r     As Long
    Dim p     As Long
    
    Dim iTop     As Long
    Dim iBottom  As Long
    Dim iStep    As Long
    
    Dim iFoundIT As Long
    
    bCase = chkMatchCase.Value = vbChecked
    If chkMatchWholeWord.Value = vbChecked Then
        sWord = txtFind.Text
    Else
        sWord = "*" & txtFind.Text & "*"
    End If
    If Not bCase Then sWord = UCase(sWord)
    
    
    With mGrid
        
        iTop = IIf(optDown.Value, .Row + 1, .Row - 1)
        iBottom = IIf(optDown.Value, .Rows - 1, .FixedRows)
        iStep = IIf(optDown.Value, 1, -1)
        
        iTop = Min(Max(1, iTop), .Rows - 1)
        iBottom = Min(Max(1, iBottom), .Rows - 1)
        
        For r = iTop To iBottom Step iStep
            If (" " & .TextMatrix(r, .Col) & " " Like sWord And bCase) Or (UCase(" " & .TextMatrix(r, .Col) & " ") Like sWord And Not bCase) Then
                
                iFoundIT = r
                
                
                r = .GetNodeRow(r, flexNTParent)
                While r <> -1
                    .GetNode(r).Expanded = True
                    r = .GetNodeRow(r, flexNTParent)
                Wend
                
                Call .Select(iFoundIT, .Col)
                Call .ShowCell(iFoundIT, .Col)
                
                Exit Sub
            End If
        Next
        
    End With
    
Exit Sub
eh: Exit Sub
End Sub

Private Sub FindNextHierarchy()
On Error GoTo eh

    Dim bCase As Boolean
    Dim sWord As String
    Dim r     As Long
    Dim p     As Long
    
    bCase = chkMatchCase.Value = vbChecked
    If chkMatchWholeWord.Value = vbChecked Then
        sWord = txtFind.Text
    Else
        sWord = "*" & txtFind.Text & "*"
    End If
    If Not bCase Then sWord = UCase(sWord)
    
    
    With mGrid
        If optDown.Value Then
            For r = .Row + 1 To .Rows - 1
                If (" " & .TextMatrix(r, 0) & " " Like sWord And bCase) Or (UCase(" " & .TextMatrix(r, 0) & " ") Like sWord And Not bCase) Then
                    p = .GetNodeRow(r, flexNTParent)
                    While p > -1
                        .GetNode(p).Expanded = True
                        p = .GetNodeRow(p, flexNTParent)
                    Wend
                    Call .Select(r, 0)
                    Call .ShowCell(r, 0)
                    Exit Sub
                End If
            Next
        Else
            For r = .Row - 1 To 0 Step -1
                If (" " & .TextMatrix(r, 0) & " " Like sWord And bCase) Or (UCase(" " & .TextMatrix(r, 0) & " ") Like sWord And Not bCase) Then
                    p = .GetNodeRow(r, flexNTParent)
                    While p > -1
                        .GetNode(p).Expanded = True
                        p = .GetNodeRow(p, flexNTParent)
                    Wend
                    Call .Select(r, 0)
                    Call .ShowCell(r, 0)
                    Exit Sub
                End If
            Next
        End If
    End With
    MsgBox "Cannot find " & vbQuote & txtFind.Text & vbQuote, vbInformation, App.ProductName
Exit Sub
eh: Exit Sub
End Sub

Private Sub txtFind_GotFocus()
    SelectAll txtFind
End Sub
