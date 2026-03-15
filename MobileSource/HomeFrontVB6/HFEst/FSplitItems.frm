VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form FSplitItems 
   Caption         =   "Split Items"
   ClientHeight    =   4260
   ClientLeft      =   8535
   ClientTop       =   2775
   ClientWidth     =   4980
   Icon            =   "FSplitItems.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   4980
   Begin VB.TextBox txtIncrement 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   215
      Left            =   2520
      TabIndex        =   7
      Text            =   "1"
      Top             =   2970
      Width           =   510
   End
   Begin VB.TextBox txtStart 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   215
      Left            =   2520
      TabIndex        =   6
      Text            =   "100"
      Top             =   2730
      Width           =   510
   End
   Begin VB.CheckBox chkUseUnits 
      Caption         =   "Assign the new items to apt/unit numbers "
      Height          =   195
      Left            =   1020
      TabIndex        =   5
      Top             =   2430
      Width           =   4035
   End
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   525
      Left            =   1590
      TabIndex        =   4
      Top             =   1830
      Visible         =   0   'False
      Width           =   3225
      _cx             =   5689
      _cy             =   926
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   360
      ColWidthMax     =   360
      ExtendLastCol   =   0   'False
      FormatString    =   $"FSplitItems.frx":000C
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   1
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.OptionButton optAllocate 
      Caption         =   "Custom percentages"
      Height          =   195
      Index           =   1
      Left            =   1320
      TabIndex        =   3
      Top             =   1590
      Width           =   1935
   End
   Begin VB.OptionButton optAllocate 
      Caption         =   "Equal shares"
      Height          =   195
      Index           =   0
      Left            =   1320
      TabIndex        =   2
      Top             =   1350
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.CheckBox chkAllocate 
      Caption         =   "Allocate dollar amount across all copies."
      Height          =   195
      Left            =   1020
      TabIndex        =   1
      Top             =   1080
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.TextBox txtCopies 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   215
      Left            =   1620
      TabIndex        =   0
      Top             =   540
      Width           =   330
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   2625
      Picture         =   "FSplitItems.frx":008E
      TabIndex        =   9
      ToolTipText     =   "Cancel"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   1320
      Picture         =   "FSplitItems.frx":0618
      TabIndex        =   8
      ToolTipText     =   "Login"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label lblStart 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Starting number"
      Enabled         =   0   'False
      Height          =   195
      Left            =   1290
      TabIndex        =   13
      Top             =   2730
      Width           =   1110
   End
   Begin VB.Label lblIncrement 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Increment By"
      Enabled         =   0   'False
      Height          =   195
      Left            =   1470
      TabIndex        =   12
      Top             =   2970
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Copies"
      Height          =   195
      Index           =   0
      Left            =   1020
      TabIndex        =   11
      Top             =   540
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Split the selected items."
      Height          =   195
      Index           =   2
      Left            =   960
      TabIndex        =   10
      Top             =   240
      Width           =   1665
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   210
      Picture         =   "FSplitItems.frx":0BA2
      Top             =   150
      Width           =   480
   End
End
Attribute VB_Name = "FSplitItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const SRCFILE = "FSplitItems"

Private mCancel As Boolean

Public Function ShowForm(Percents() As Double, WriteUnits As Boolean, UnitStart As Long, UnitIncrement As Long) As Boolean
    Dim c As Long
    Dim i As Long
    
    mCancel = False
    Me.Show vbModal
    If Not mCancel Then
        ShowForm = Not mCancel
        c = Val(txtCopies.Text)
        ReDim Percents(c) As Double
        If chkAllocate.value = vbUnchecked Then
            Percents(1) = 1
        Else
            If optAllocate(0).value Then
                For i = 1 To c - 1
                    Percents(i) = 1 / c
                Next
                Percents(c) = 1 - ((c - 1) / c)
            Else
                For i = 1 To c
                    Percents(i) = gData.ValueMatrix(0, i - 1) / 100
                Next
            End If
        End If
        WriteUnits = chkUseUnits.value = vbChecked
        UnitStart = Val(txtStart.Text)
        UnitIncrement = Val(txtIncrement.Text)
    End If
    
    Unload Me
End Function



Private Sub gData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Col = gData.Cols - 1
End Sub
Private Sub optAllocate_Click(Index As Integer)
    Call chkAllocate_Click
    Call gData_AfterEdit(0, 0)
End Sub
Private Sub chkAllocate_Click()
    optAllocate(0).Enabled = chkAllocate.value = vbChecked
    optAllocate(1).Enabled = chkAllocate.value = vbChecked
    gData.Visible = optAllocate(1).value
End Sub
Private Sub chkUseUnits_Click()
    txtStart.Enabled = chkUseUnits.value = vbChecked
    txtIncrement.Enabled = chkUseUnits.value = vbChecked
    lblStart.Enabled = chkUseUnits.value = vbChecked
    lblIncrement.Enabled = chkUseUnits.value = vbChecked
End Sub



Private Sub Form_Load()
    Call IniGetForm(Me)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        mCancel = True
        Me.Hide
    End If
End Sub
Private Sub cmdNav_Click(Index As Integer)
    mCancel = Index = 1
    Me.Hide
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub

Private Sub txtCopies_GotFocus()
    SelectAll txtCopies
End Sub
Private Sub txtCopies_Validate(Cancel As Boolean)
'On Error Resume Next
    Dim i As Long
    i = Abs(Int(Val(txtCopies.Text)))
    txtCopies.Text = i
    gData.Cols = i
    Call gData_AfterEdit(0, 0)
        
    cmdNav(0).Enabled = Val(txtCopies.Text) <> 0
    
End Sub

Private Sub txtIncrement_GotFocus()
    SelectAll txtIncrement
End Sub

Private Sub txtIncrement_Validate(Cancel As Boolean)
    txtIncrement.Text = Val(txtIncrement.Text)
End Sub

Private Sub txtStart_GotFocus()
    SelectAll txtStart
End Sub
Private Sub txtStart_Validate(Cancel As Boolean)
    txtStart.Text = Val(txtStart.Text)
End Sub

Private Sub gData_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo eh
    Dim i As Long
    Dim t As Double
    With gData
        .Redraw = flexRDNone
        .Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbWindowText
        .Cell(flexcpForeColor, 0, .Cols - 1) = vbGrayText
        
        .ColWidthMin = 0
        .ColWidthMax = 0
        
        Call .AutoSize(0, .Cols - 1)
        
        t = 0
        For i = 0 To .Cols - 2
            If optAllocate(0).value Then
                .TextMatrix(0, i) = Int(1 / .Cols * 100)
            Else
                .TextMatrix(0, i) = .ValueMatrix(0, i)
            End If
            t = t + .ValueMatrix(0, i)
        Next
        .TextMatrix(0, i) = 100 - t
        
        .Redraw = flexRDBuffered
    End With
eh: Exit Sub
End Sub

