VERSION 5.00
Begin VB.Form frmCalculator
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   3600
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDisplay
      Alignment       =   1  'Right Justify
      BeginProperty Font
         Name            =   "Segoe UI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton cmdClear
      Caption         =   "C"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   900
      Width           =   795
   End
   Begin VB.CommandButton cmdClearEntry
      Caption         =   "CE"
      Height          =   615
      Left            =   1020
      TabIndex        =   2
      Top             =   900
      Width           =   795
   End
   Begin VB.CommandButton cmdBackspace
      Caption         =   "<-"
      Height          =   615
      Left            =   1920
      TabIndex        =   3
      Top             =   900
      Width           =   795
   End
   Begin VB.CommandButton cmdSign
      Caption         =   "+/-"
      Height          =   615
      Left            =   2820
      TabIndex        =   4
      Top             =   900
      Width           =   795
   End
   Begin VB.CommandButton cmdNum
      Caption         =   "7"
      Height          =   615
      Index           =   7
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   795
   End
   Begin VB.CommandButton cmdNum
      Caption         =   "8"
      Height          =   615
      Index           =   8
      Left            =   1020
      TabIndex        =   6
      Top             =   1680
      Width           =   795
   End
   Begin VB.CommandButton cmdNum
      Caption         =   "9"
      Height          =   615
      Index           =   9
      Left            =   1920
      TabIndex        =   7
      Top             =   1680
      Width           =   795
   End
   Begin VB.CommandButton cmdOp
      Caption         =   "/"
      Height          =   615
      Index           =   3
      Left            =   2820
      TabIndex        =   8
      Top             =   1680
      Width           =   795
   End
   Begin VB.CommandButton cmdNum
      Caption         =   "4"
      Height          =   615
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   795
   End
   Begin VB.CommandButton cmdNum
      Caption         =   "5"
      Height          =   615
      Index           =   5
      Left            =   1020
      TabIndex        =   10
      Top             =   2400
      Width           =   795
   End
   Begin VB.CommandButton cmdNum
      Caption         =   "6"
      Height          =   615
      Index           =   6
      Left            =   1920
      TabIndex        =   11
      Top             =   2400
      Width           =   795
   End
   Begin VB.CommandButton cmdOp
      Caption         =   "*"
      Height          =   615
      Index           =   2
      Left            =   2820
      TabIndex        =   12
      Top             =   2400
      Width           =   795
   End
   Begin VB.CommandButton cmdNum
      Caption         =   "1"
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   3120
      Width           =   795
   End
   Begin VB.CommandButton cmdNum
      Caption         =   "2"
      Height          =   615
      Index           =   2
      Left            =   1020
      TabIndex        =   14
      Top             =   3120
      Width           =   795
   End
   Begin VB.CommandButton cmdNum
      Caption         =   "3"
      Height          =   615
      Index           =   3
      Left            =   1920
      TabIndex        =   15
      Top             =   3120
      Width           =   795
   End
   Begin VB.CommandButton cmdOp
      Caption         =   "-"
      Height          =   615
      Index           =   1
      Left            =   2820
      TabIndex        =   16
      Top             =   3120
      Width           =   795
   End
   Begin VB.CommandButton cmdNum
      Caption         =   "0"
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   3840
      Width           =   795
   End
   Begin VB.CommandButton cmdDecimal
      Caption         =   "."
      Height          =   615
      Left            =   1020
      TabIndex        =   18
      Top             =   3840
      Width           =   795
   End
   Begin VB.CommandButton cmdEquals
      Caption         =   "="
      Height          =   615
      Left            =   1920
      TabIndex        =   19
      Top             =   3840
      Width           =   795
   End
   Begin VB.CommandButton cmdOp
      Caption         =   "+"
      Height          =   615
      Index           =   0
      Left            =   2820
      TabIndex        =   20
      Top             =   3840
      Width           =   795
   End
End
Attribute VB_Name = "frmCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private dblAccumulator As Double
Private dblCurrentValue As Double
Private strCurrentOp As String
Private blnNewEntry As Boolean
Private blnDecimalUsed As Boolean
Private blnHasError As Boolean

Private Sub Form_Load()
    Call ClearAll
End Sub

Private Sub ClearAll()
    dblAccumulator = 0
    dblCurrentValue = 0
    strCurrentOp = ""
    blnNewEntry = True
    blnDecimalUsed = False
    blnHasError = False
    txtDisplay.Text = "0"
End Sub

Private Sub ClearEntry()
    dblCurrentValue = 0
    blnNewEntry = True
    blnDecimalUsed = False
    txtDisplay.Text = "0"
End Sub

Private Sub cmdClear_Click()
    Call ClearAll
End Sub

Private Sub cmdClearEntry_Click()
    Call ClearEntry
End Sub

Private Sub cmdBackspace_Click()
    Dim strText As String

    If blnHasError Then Exit Sub
    If blnNewEntry Then Exit Sub

    strText = txtDisplay.Text

    If Len(strText) > 1 Then
        ' Check if we are removing the decimal point
        If Right$(strText, 1) = "." Then
            blnDecimalUsed = False
        End If
        txtDisplay.Text = Left$(strText, Len(strText) - 1)
    Else
        txtDisplay.Text = "0"
        blnNewEntry = True
    End If

    dblCurrentValue = CDbl(txtDisplay.Text)
End Sub

Private Sub cmdSign_Click()
    If blnHasError Then Exit Sub

    dblCurrentValue = -dblCurrentValue
    txtDisplay.Text = CStr(dblCurrentValue)
End Sub

Private Sub cmdNum_Click(Index As Integer)
    If blnHasError Then Call ClearAll

    If blnNewEntry Then
        txtDisplay.Text = ""
        blnNewEntry = False
        blnDecimalUsed = False
    End If

    txtDisplay.Text = txtDisplay.Text & CStr(Index)
    dblCurrentValue = CDbl(txtDisplay.Text)
End Sub

Private Sub cmdDecimal_Click()
    If blnHasError Then Call ClearAll

    If blnDecimalUsed Then Exit Sub

    If blnNewEntry Then
        txtDisplay.Text = "0"
        blnNewEntry = False
    End If

    txtDisplay.Text = txtDisplay.Text & "."
    blnDecimalUsed = True
End Sub

Private Sub cmdOp_Click(Index As Integer)
    If blnHasError Then Exit Sub

    ' If there is a pending operation, compute it first
    If strCurrentOp <> "" And Not blnNewEntry Then
        Call PerformCalculation
        If blnHasError Then Exit Sub
    End If

    dblAccumulator = dblCurrentValue
    blnNewEntry = True

    Select Case Index
        Case 0
            strCurrentOp = "+"
        Case 1
            strCurrentOp = "-"
        Case 2
            strCurrentOp = "*"
        Case 3
            strCurrentOp = "/"
    End Select
End Sub

Private Sub cmdEquals_Click()
    If blnHasError Then Exit Sub
    If strCurrentOp = "" Then Exit Sub

    Call PerformCalculation
    strCurrentOp = ""
End Sub

Private Sub PerformCalculation()
    Dim dblResult As Double

    On Error GoTo CalcError

    Select Case strCurrentOp
        Case "+"
            dblResult = dblAccumulator + dblCurrentValue
        Case "-"
            dblResult = dblAccumulator - dblCurrentValue
        Case "*"
            dblResult = dblAccumulator * dblCurrentValue
        Case "/"
            If dblCurrentValue = 0 Then
                txtDisplay.Text = "Error: Div/0"
                blnHasError = True
                Exit Sub
            End If
            dblResult = dblAccumulator / dblCurrentValue
    End Select

    dblCurrentValue = dblResult
    dblAccumulator = dblResult
    txtDisplay.Text = FormatResult(dblResult)
    blnNewEntry = True
    blnDecimalUsed = InStr(txtDisplay.Text, ".") > 0

    Exit Sub

CalcError:
    txtDisplay.Text = "Error"
    blnHasError = True
End Sub

Private Function FormatResult(ByVal dblValue As Double) As String
    ' Remove trailing zeros after decimal point
    Dim strResult As String
    strResult = CStr(dblValue)

    ' If the result is very large or very small, use scientific notation
    If Abs(dblValue) >= 1E+15 Or (Abs(dblValue) < 0.000001 And dblValue <> 0) Then
        strResult = Format$(dblValue, "0.########E+0")
    ElseIf dblValue = Int(dblValue) Then
        strResult = CStr(CLng(dblValue))
    End If

    FormatResult = strResult
End Function
