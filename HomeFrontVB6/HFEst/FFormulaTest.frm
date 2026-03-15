VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FFormulaTest 
   Caption         =   "Formula Test"
   ClientHeight    =   3765
   ClientLeft      =   5070
   ClientTop       =   3630
   ClientWidth     =   4905
   Icon            =   "FFormulaTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   4905
   Begin VB.CommandButton cmdNav 
      Caption         =   "&Test"
      Height          =   375
      Index           =   0
      Left            =   2340
      Picture         =   "FFormulaTest.frx":000C
      TabIndex        =   1
      ToolTipText     =   "Cancel"
      Top             =   3330
      Width           =   1215
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Index           =   1
      Left            =   3600
      Picture         =   "FFormulaTest.frx":0596
      TabIndex        =   0
      ToolTipText     =   "Cancel"
      Top             =   3330
      Width           =   1215
   End
   Begin VSFlex8Ctl.VSFlexGrid gVariables 
      Height          =   1995
      Left            =   90
      TabIndex        =   2
      Top             =   60
      Width           =   2505
      _cx             =   4419
      _cy             =   3519
      Appearance      =   2
      BorderStyle     =   1
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
      GridColorFixed  =   -2147483633
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   11
      FixedRows       =   0
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FFormulaTest.frx":0B20
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   5
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   2
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   1
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
      BackColorFrozen =   -2147483643
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "FFormulaTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mRawFormula As String
Private mResolvedFormula As String
Private mVariableList As String

Private ToolTip As New CToolTips


Public Sub TestFormula(Formula As String)

    Load Me
    mRawFormula = Formula
    mResolvedFormula = ResolveFormula(mRawFormula)
    mVariableList = ParseVariables(mResolvedFormula)
    gVariables.Rows = 0
    Call LoadVariables(True, gVariables, mVariableList)
    
    If gVariables.Rows = 0 Then
        Call cmdNav_Click(0)
        Unload Me
    Else
        Me.Show vbModal
    End If
End Sub









Public Function GetItemTotal(TotalType As String, POIndex As String, CostCode As String, Phase As String, Item As String) As Double
    GetItemTotal = 0
End Function

Private Sub cmdNav_Click(Index As Integer)
On Error GoTo eh
    Select Case Index
    
        Case 0 'test
            Dim d As Double
            d = CalcFormula(gVariables, mResolvedFormula, False, Me)
            MsgBox "Formula is valid." & vbCrLf & "return value:  " & Round(d, 3), vbInformation, App.ProductName

        Case 1 'close
            Unload Me
            
    End Select
eh: Exit Sub
End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
    Call ToolTip.Create(Me)
    ToolTip.MaxTipWidth = 3999
    ToolTip.DelayTime(ttDelayShow) = 20000
    Call ToolTip.AddTool(gVariables)
End Sub

Private Sub Form_Resize()
    
    With gVariables
    
        .Move 60, 60, Me.ScaleWidth - 120, Me.ScaleHeight - 180 - cmdNav(0).Height
    
    
        Call .AutoSize(0, 1, 2)
        .ColWidth(1) = Max(.ColWidth(1), 1020)
        .ColWidth(2) = Max(.ColWidth(2), 700)
        .ColWidth(0) = .Width - .ColWidth(1) - .ColWidth(2)
        
        .Col = 3
        .Sort = flexSortNumericAscending
        .Col = 1
    End With

    cmdNav(1).Move Me.ScaleWidth - 1 * (cmdNav(0).Width + 60), Me.ScaleHeight - cmdNav(0).Height - 60
    cmdNav(0).Move Me.ScaleWidth - 2 * (cmdNav(1).Width + 60), Me.ScaleHeight - cmdNav(0).Height - 60

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub




















Private Sub gVariables_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo eh
    Dim s As String
    Dim i As Long
    Dim w As Long
    
    With gVariables
        
        .Redraw = flexRDNone
        
        gVariables.SetFocus
        
        i = .ColWidth(0)
        Call .AutoSize(0)
        w = .ColWidth(0)
        .ColWidth(0) = i
        .Redraw = flexRDBuffered
        

        
        .Row = Min(.Row, GridLastVisibleRow(gVariables))
        
        cmdNav(0).Default = .Row = GridLastVisibleRow(gVariables)
        
    End With
eh: Exit Sub
End Sub

Private Sub gVariables_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gVariables
        .ComboList = .TextMatrix(Row, .ColIndex("ListOfValues"))
        Cancel = Col <> .ColIndex("value")
    End With
End Sub

Private Sub gVariables_GotFocus()
On Error Resume Next
    If gVariables.Row < 0 Then gVariables.Row = 0
    gVariables.Col = 1
End Sub

Private Sub gVariables_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        gVariables.Row = gVariables.Row + 1
        Call gVariables.ShowCell(gVariables.Row, gVariables.Col)
    End If
End Sub



Private Sub gVariables_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
On Error Resume Next
    If KeyCode = vbKeyReturn And gVariables.Row <> gVariables.Rows - 1 Then
        KeyCode = 0
        gVariables.Row = gVariables.Row + 1
    End If
    Call gVariables.ShowCell(gVariables.Row, gVariables.Col)
End Sub

Private Sub gVariables_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Dim s As String
    s = Trim(gVariables.TextMatrix(gVariables.MouseRow, 6))
    ToolTip.ToolText(gVariables) = s
End Sub

Private Sub gVariables_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim i As Long
    Dim d As Long
    With gVariables
        If .TextMatrix(Row, .ColIndex("ListOfValues")) <> "" Then
            If IsNumeric(.EditText) Then
                .EditText = Parse(.TextMatrix(Row, .ColIndex("ListOfValues")), Val(.EditText), "|")
            End If
        Else
            If Not IsNumeric(.EditText) Then
                Cancel = True
            Else
                If Val(.EditText) > .ValueMatrix(Row, 5) And .ValueMatrix(Row, 5) > 0 Then
                    MsgBox "The maximum value for " & .TextMatrix(Row, 0) & " is '" & format(.ValueMatrix(Row, 5), "0") & "'. The number you have entered is too large.", vbInformation, App.ProductName
                    Cancel = True
                End If
                If Val(.EditText) < .ValueMatrix(Row, 4) Then
                    MsgBox "The minimum value for " & .TextMatrix(Row, 0) & " is '" & format(.ValueMatrix(Row, 4), "0") & "'. The number you have entered is too small.", vbInformation, App.ProductName
                    Cancel = True
                End If
            End If
            .EditText = Val(.EditText)
        End If
        If Not Cancel Then
            On Error Resume Next
            .Col = 1
            .Row = Row
        End If
    End With
End Sub

