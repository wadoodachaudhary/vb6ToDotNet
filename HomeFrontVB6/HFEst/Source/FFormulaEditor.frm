VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{302C5C1A-C2E2-4302-9AF9-BAC87EEFDECE}#1.0#0"; "Panels.ocx"
Begin VB.Form FFormulaEditor 
   Caption         =   "Formula Editor"
   ClientHeight    =   7110
   ClientLeft      =   6780
   ClientTop       =   2220
   ClientWidth     =   9630
   Icon            =   "FFormulaEditor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   9630
   Begin Panels.Slider Slider 
      Height          =   45
      Index           =   2
      Left            =   90
      Top             =   2310
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   79
      Orientation     =   1
   End
   Begin Panels.Slider Slider 
      Height          =   2205
      Index           =   0
      Left            =   3150
      Top             =   0
      Width           =   45
      _ExtentX        =   79
      _ExtentY        =   3889
   End
   Begin Panels.Slider Slider 
      Height          =   2205
      Index           =   1
      Left            =   6300
      Top             =   0
      Width           =   45
      _ExtentX        =   79
      _ExtentY        =   3889
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&Test"
      Height          =   375
      Index           =   2
      Left            =   60
      Picture         =   "FFormulaEditor.frx":058A
      TabIndex        =   6
      ToolTipText     =   "Cancel"
      Top             =   6660
      Width           =   1215
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   7110
      Picture         =   "FFormulaEditor.frx":0B14
      TabIndex        =   5
      ToolTipText     =   "Cancel"
      Top             =   6660
      Width           =   1215
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Height          =   375
      Index           =   0
      Left            =   5850
      Picture         =   "FFormulaEditor.frx":109E
      TabIndex        =   4
      ToolTipText     =   "Cancel"
      Top             =   6660
      Width           =   1215
   End
   Begin VSFlex8Ctl.VSFlexGrid gFunctions 
      Height          =   2205
      Left            =   6420
      TabIndex        =   1
      Top             =   0
      Width           =   3045
      _cx             =   1966544123
      _cy             =   1966542641
      Appearance      =   1
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483643
      GridColorFixed  =   -2147483643
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FFormulaEditor.frx":1628
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   7
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   6
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
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
   Begin VB.TextBox txtFormula 
      BorderStyle     =   0  'None
      Height          =   3705
      Left            =   150
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "FFormulaEditor.frx":1665
      Top             =   2400
      Width           =   9645
   End
   Begin VSFlex8Ctl.VSFlexGrid gVariables 
      Height          =   2205
      Left            =   30
      TabIndex        =   2
      Top             =   30
      Width           =   3045
      _cx             =   1966544123
      _cy             =   1966542641
      Appearance      =   1
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483643
      GridColorFixed  =   -2147483643
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FFormulaEditor.frx":166B
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   7
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   6
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
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
   Begin VSFlex8Ctl.VSFlexGrid gFormulas 
      Height          =   2205
      Left            =   3240
      TabIndex        =   3
      Top             =   0
      Width           =   3045
      _cx             =   1966544123
      _cy             =   1966542641
      Appearance      =   1
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483643
      GridColorFixed  =   -2147483643
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FFormulaEditor.frx":16A8
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   6
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
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
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupSub 
         Caption         =   "Insert"
         Index           =   0
      End
      Begin VB.Menu mnuPopupSub 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuPopupSub 
         Caption         =   "Rename"
         Index           =   2
      End
      Begin VB.Menu mnuPopupSub 
         Caption         =   "New"
         Index           =   3
      End
      Begin VB.Menu mnuPopupSub 
         Caption         =   "Edit"
         Index           =   4
      End
      Begin VB.Menu mnuPopupSub 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuPopupSub 
         Caption         =   "Delete"
         Index           =   6
      End
   End
End
Attribute VB_Name = "FFormulaEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SRCFILE = "FFormulaEditor::"

Private mCancel As Boolean

Private Const mc_INSERT = 0
Private Const mc_RENAME = 2
Private Const mc_NEW = 3
Private Const mc_EDIT = 4
Private Const mc_DELETE = 6

Dim ToolTip As New CToolTips

Public Function EditFormula(ByVal Name As String, ByRef Text As String) As Boolean
    Load Me
    
    
    
    Me.Caption = "Formula Editor"
    If Name <> "" Then Me.Caption = Me.Caption & " - " & Name
    txtFormula.Text = Text
    
    mCancel = False
    
    
    Me.Show vbModal
    
    EditFormula = Not mCancel
    If Not mCancel Then Text = txtFormula.Text
    Unload Me

End Function


Private Sub cmdNav_Click(Index As Integer)
    Select Case Index
        Case 0 'ok
            mCancel = False
            Me.Hide
        Case 1 'cancel
            mCancel = True
            Me.Hide
        Case 2 'test
            Call FFormulaTest.TestFormula(txtFormula.Text)
            
    End Select
End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
    Call LoadFunctions
    Call LoadVariables
    Call LoadFormulas


    Call ToolTip.Create(Me)
    ToolTip.MaxTipWidth = 3999
    ToolTip.DelayTime(ttDelayShow) = 20000
    Call ToolTip.AddTool(gVariables)
    Call ToolTip.AddTool(gFormulas)
    Call ToolTip.AddTool(gFunctions)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        mCancel = True
        Me.Hide
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub

Private Sub gFormulas_DblClick()
    Call mnuPopupSub_Click(mc_INSERT)
End Sub

Private Sub gFormulas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 And gFormulas.Row > 0 Then gFormulas.EditCell
End Sub

Private Sub gFormulas_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim s As String
    
    With gFormulas
        s = "update formulas set name=" & DbQuote(Str, .EditText) & " where name=" & DbQuote(Str, .Text)
        On Error Resume Next
        Call HFApp.SqlExec(s, dbHomefront)
        If Err.Number <> 0 Then
            Cancel = True
            MsgBox "There is already a formula named " & .EditText & ". Please specify a unique name.", vbInformation, App.ProductName
        Else
            s = "update formulas set text = replace(text," & DbQuote(Str, "{" & .Text & "}") & "," & DbQuote(Str, "{" & .EditText & "}") & ")" & vbCrLf & _
                "update tblphaseitem set formula = replace(formula," & DbQuote(Str, "{" & .Text & "}") & "," & DbQuote(Str, "{" & .EditText & "}") & ")" & vbCrLf & _
                "update tbldbassemblydetails set formula = replace(formula," & DbQuote(Str, "{" & .Text & "}") & "," & DbQuote(Str, "{" & .EditText & "}") & ")"
            Call HFApp.SqlExec(s, dbHomefront)
        End If
    End With
End Sub

Private Sub gFunctions_DblClick()
    Call mnuPopupSub_Click(mc_INSERT)
End Sub

Private Sub gFunctions_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    With gFunctions
        If .Row < 0 Then Exit Sub
        Select Case True
        
            Case KeyCode = vbKeyF And Shift = vbCtrlMask
                Call FFind.ShowForm(gFunctions)
            
            Case KeyCode = vbKeyReturn
                
            Case KeyCode = vbKeyLeft
                If .IsSubtotal(.Row) And .IsCollapsed(.Row) <> flexOutlineCollapsed Then
                    .IsCollapsed(.Row) = flexOutlineCollapsed
                Else
                    If .GetNodeRow(.Row, flexNTParent) <> -1 Then .Row = .GetNodeRow(.Row, flexNTParent)
                End If
                
            Case KeyCode = vbKeyRight
                If .IsSubtotal(.Row) Then
                    If .IsCollapsed(.Row) = flexOutlineCollapsed Then
                        .IsCollapsed(.Row) = flexOutlineExpanded
                    Else
                        .Row = .GetNodeRow(.Row, flexNTFirstChild)
                    End If
                End If
                
        End Select
    End With
End Sub

Private Sub gVariables_DblClick()
    Call mnuPopupSub_Click(mc_INSERT)
End Sub



Private Sub gFormulas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbRightButton Then Exit Sub
    gFormulas.SetFocus
    mnuPopupSub(mc_INSERT).Enabled = txtFormula.Visible
    mnuPopupSub(mc_RENAME).Enabled = True
    mnuPopupSub(mc_NEW).Enabled = True
    mnuPopupSub(mc_EDIT).Enabled = True
    mnuPopupSub(mc_DELETE).Enabled = True
    PopupMenu mnuPopup, , , , mnuPopupSub(0)
End Sub
Private Sub gFunctions_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbRightButton Then Exit Sub
    Call gFunctions.SetFocus
    If gFunctions.Row < 0 Then Exit Sub
    mnuPopupSub(mc_INSERT).Enabled = txtFormula.Visible And Not gFunctions.IsSubtotal(gFunctions.Row)
    mnuPopupSub(mc_RENAME).Enabled = False
    mnuPopupSub(mc_NEW).Enabled = False
    mnuPopupSub(mc_EDIT).Enabled = False
    mnuPopupSub(mc_DELETE).Enabled = False
    PopupMenu mnuPopup, , , , mnuPopupSub(0)
End Sub

Private Sub gVariables_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 And gVariables.Row > 0 Then gVariables.EditCell
End Sub

Private Sub gVariables_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbRightButton Then Exit Sub
    Call gVariables.SetFocus
    If gVariables.Row < 0 Then Exit Sub
    mnuPopupSub(mc_INSERT).Enabled = txtFormula.Visible And Not gVariables.IsSubtotal(gVariables.Row)
    mnuPopupSub(mc_RENAME).Enabled = True
    mnuPopupSub(mc_NEW).Enabled = True
    mnuPopupSub(mc_EDIT).Enabled = Not gVariables.IsSubtotal(gVariables.Row)
    mnuPopupSub(mc_DELETE).Enabled = Not gVariables.IsSubtotal(gVariables.Row)
    PopupMenu mnuPopup, , , , mnuPopupSub(0)
End Sub


Private Sub gFunctions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Dim s As String
    s = gFunctions.TextMatrix(gFunctions.MouseRow, 1)
    ToolTip.ToolText(gFunctions) = Trim(s)
End Sub
Private Sub gVariables_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Dim s As String
    s = gVariables.TextMatrix(gVariables.MouseRow, 1)
    ToolTip.ToolText(gVariables) = Trim(s)
End Sub
Private Sub gFormulas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Dim s As String
    s = gFormulas.TextMatrix(gFormulas.MouseRow, 1)
    ToolTip.ToolText(gFormulas) = Trim(s)
End Sub


Private Sub gVariables_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim s As String
    
    With gVariables
        s = "update variables set name=" & DbQuote(Str, .EditText) & " where name=" & DbQuote(Str, .Text)
        On Error Resume Next
        Call HFApp.SqlExec(s, dbHomefront)
        If Err.Number <> 0 Then
            Cancel = True
            MsgBox "There is already a variable named " & .EditText & ". Please specify a unique name.", vbInformation, App.ProductName
        Else
            s = "update formulas set text = replace(text," & DbQuote(Str, "[" & .Text & "]") & "," & DbQuote(Str, "[" & .EditText & "]") & ")" & vbCrLf & _
                "update tblphaseitem set formula = replace(formula," & DbQuote(Str, "[" & .Text & "]") & "," & DbQuote(Str, "[" & .EditText & "]") & ")" & vbCrLf & _
                "update tbldbassemblydetails set formula = replace(formula," & DbQuote(Str, "[" & .Text & "]") & "," & DbQuote(Str, "[" & .EditText & "]") & ")"
            Call HFApp.SqlExec(s, dbHomefront)
        End If
    End With
End Sub

Private Sub mnuPopupSub_Click(Index As Integer)
On Error GoTo eh
    Dim s As String
    Dim r As Long
    Dim rs As Recordset
    Dim F As FFormulaEditor
    
    Select Case Index
    
        Case mc_INSERT
            With Screen.ActiveControl
                If .Row < 0 Then Exit Sub
                If Not .IsSubtotal(.Row) Then
                
                    If Screen.ActiveControl Is gVariables Then
                        s = "[" & .Text & "]"
                    ElseIf Screen.ActiveControl Is gFormulas Then
                        s = "{" & .Text & "}"
                    Else
                        s = Trim(.Text)
                    End If
                
                    txtFormula.SelText = s
                    SetCtrlFocus txtFormula
                End If
            End With
        
        Case mc_RENAME
            Call Screen.ActiveControl.EditCell
        
        
        
        Case mc_NEW
            If Screen.ActiveControl Is gVariables Then
                If FVariable.EditVariable("") Then Call LoadVariables
            Else
                s = InputBox("Formula name?", App.ProductName)
                If s = "" Then Exit Sub
                
                On Error Resume Next
                Call HFApp.SqlExec("insert into formulas(name) values(" & DbQuote(Str, s) & ")", dbHomefront)
                If Err.Number = 0 Then
                    Call LoadFormulas
                    r = gFormulas.FindRow(s, , 0)
                    If r > 0 Then
                        Call gFormulas.Select(r, 0)
                        Call gFormulas.ShowCell(r, 0)
                        gFormulas.SetFocus
                        Call mnuPopupSub_Click(mc_EDIT)
                        
                        
                    End If
                Else
                    MsgBox "A formula named " & s & " already exists. You must specify a unique name.", vbInformation, App.ProductName
                    Exit Sub
                End If
                
            End If
        
        
        
        Case mc_EDIT
            If Screen.ActiveControl Is gVariables Then
                If FVariable.EditVariable(gVariables.Text) Then Call LoadVariables
            ElseIf Screen.ActiveControl Is gFormulas Then
                s = "select * from formulas where name=" & DbQuote(Str, gFormulas.Text)
                Set rs = HFApp.SqlExec(s, dbHomefront)
                If rs.EOF Then Exit Sub
                s = "" & rs("text")
                Set F = New FFormulaEditor
                If F.EditFormula(gFormulas.Text, s) Then
                    gFormulas.TextMatrix(gFormulas.Row, 1) = s
                    s = "update formulas set text=" & DbQuote(Str, s) & " where name=" & DbQuote(Str, gFormulas.Text)
                    Set rs = HFApp.SqlExec(s, dbHomefront)
                End If
            End If
        
        
        
        Case mc_DELETE
            If Screen.ActiveControl Is gVariables Then
                If vbOK = MsgBox("Are you sure you want to delete " & gVariables.Text & "?", vbOKCancel + vbQuestion, App.ProductName) Then
                    s = "delete from variables where name=" & DbQuote(Str, gVariables.Text)
                    HFApp.SqlExec s
                    gVariables.RemoveItem
                End If
            ElseIf Screen.ActiveControl Is gFormulas Then
                If vbOK = MsgBox("Are you sure you want to delete " & gFormulas.Text & "?", vbOKCancel + vbQuestion, App.ProductName) Then
                    s = "delete from formulas where name=" & DbQuote(Str, gFormulas.Text)
                    HFApp.SqlExec s
                    gFormulas.RemoveItem
                End If
            End If
        
        
        
    End Select
eh:
Exit Sub
End Sub



Private Sub LoadFunctions()
    With gFunctions
        .Rows = 0
        
        .AddItem "Operators":
        .Cell(flexcpFontBold, .Rows - 1, 0) = True
        .RowOutlineLevel(.Rows - 1) = 0
        .IsSubtotal(.Rows - 1) = True
            .AddItem "+"
            .AddItem "-"
            .AddItem "*"
            .AddItem "/"
            .AddItem "^           " & vbTab & "Exponent. IE: 3 squared could be expressed as 3^2"
            .AddItem "%           " & vbTab & "Modulus. Returns the remainder of a division operation"
            
            
        .AddItem "Comparisons":
        .Cell(flexcpFontBold, .Rows - 1, 0) = True
        .RowOutlineLevel(.Rows - 1) = 0
        .IsSubtotal(.Rows - 1) = True
            .AddItem "="
            .AddItem ">"
            .AddItem "<"
            .AddItem "<="
            .AddItem ">="
            .AddItem "<>"
            
        
        
        .AddItem "Numeric Functions"
        .Cell(flexcpFontBold, .Rows - 1, 0) = True
        .RowOutlineLevel(.Rows - 1) = 0
        .IsSubtotal(.Rows - 1) = True
            .AddItem "ABS ( numeric )           " & vbTab & "A mathematical function that returns the" & vbCrLf & _
                                                            "absolute (positive) value of the specified" & vbCrLf & _
                                                            "numeric expression."
            .AddItem "ARCCOS ( numeric )        " & vbTab & "A mathematical function that returns the" & vbCrLf & _
                                                            "angle, in radians, whose cosine is the specified" & vbCrLf & _
                                                            "numeric expression; also called arccosine."
            .AddItem "ARCSIN ( numeric )        " & vbTab & "Returns the angle, in radians, whose sine is" & vbCrLf & _
                                                            "the specified numeric expression. This is" & vbCrLf & _
                                                            "also called arcsine."
            .AddItem "ARCTAN ( numeric )        " & vbTab & "Returns the angle in radians whose tangent" & vbCrLf & _
                                                            "is a specified numeric expression. This is" & vbCrLf & _
                                                            "also called arctangent."
            
            .AddItem "AVG ( numeric , numeric [, numeric] ) " & vbTab & "Returns the average of the 2 or 3 specified values."
            .AddItem "CEILING ( numeric )       " & vbTab & "Returns the smallest integer greater than, or" & vbCrLf & _
                                                            "equal to, the specified numeric expression."
            .AddItem "COS ( numeric )           " & vbTab & "Is a mathematical function that returns the" & vbCrLf & _
                                                            "trigonometric cosine of the specified angle," & vbCrLf & _
                                                            "in radians, in the specified expression."
            .AddItem "COT ( numeric )           " & vbTab & "A mathematical function that returns the trigonometric" & vbCrLf & _
                                                            "cotangent of the specified angle, in radians, in the" & vbCrLf & _
                                                            "specified numeric expression."
            .AddItem "DEGREES ( numeric )       " & vbTab & "Returns the corresponding angle in degrees for an angle" & vbCrLf & _
                                                            "specified in radians."
            .AddItem "EXP ( numeric )           " & vbTab & "Returns the exponential value of the specified expression."
            .AddItem "FLOOR ( numeric )         " & vbTab & "Returns the largest integer less than or equal" & vbCrLf & _
                                                            "to the specified numeric expression."
            .AddItem "LOG ( numeric )           " & vbTab & "Returns the natural logarithm of the specified" & vbCrLf & _
                                                            "numeric expression."
            .AddItem "LOG10 ( numeric )         " & vbTab & "Returns the base-10 logarithm of the specified" & vbCrLf & _
                                                            "numeric expression."
            .AddItem "MAX ( numeric , numeric [, numeric] ) " & vbTab & "Returns the maximum of the 2 or 3 specified values."
            .AddItem "MIN ( numeric , numeric [, numeric] ) " & vbTab & "Returns the minimum of the 2 or 3 specified values."
            .AddItem "MOD ( value , divisor )   " & vbTab & "Returns the integer remainder of value divided" & vbCrLf & _
                                                            "by divisor."
            .AddItem "PI ( )                    " & vbTab & "Returns the constant value of PI."
            .AddItem "POWER ( numeric , y )     " & vbTab & "Returns the value of the specified expression" & vbCrLf & _
                                                            "to the specified power."
            .AddItem "RADIANS ( numeric )       " & vbTab & "Returns radians when a numeric expression, in" & vbCrLf & _
                                                            "degrees, is entered."
            .AddItem "ROUND ( numeric , length )" & vbTab & "Returns a numeric value, rounded to the" & vbCrLf & _
                                                            "specified length or precision."
            .AddItem "ROUNDTO ( value , method , multiplier )" & vbTab & _
                                                            "Returns a numeric value, rounded to the" & vbCrLf & _
                                                            "specified length or precision using a specific method." & vbCrLf & _
                                                            "Method = 1, 0, -1 = round up, round closest, round down."
            .AddItem "SIN ( numeric )           " & vbTab & "Returns the trigonometric sine of the specified angle," & vbCrLf & _
                                                            "in radians, and in an approximate numeric, numeric," & vbCrLf & _
                                                            "expression."
            .AddItem "SQRT ( numeric )          " & vbTab & "Returns the square root of the specified numeric value."
            .AddItem "SQUARE ( numeric )        " & vbTab & "Returns the square of the specified numeric value."
            .AddItem "TAN ( numeric )           " & vbTab & "Returns the tangent of the input expression."
            .AddItem "TRUNC( value , decimals ) " & vbTab & "Returns value with the specified number of decimal places." & vbCrLf & _
                                                            "If decimals is negative, value is truncated to the left of" & vbCrLf & _
                                                            "the decimal point." & vbCrLf & vbCrLf & _
                                                            "   TRUNC(1538.3518, 3) = 1538.3510" & vbCrLf & _
                                                            "   TRUNC(1538.3518, 1) = 1538.3000" & vbCrLf & _
                                                            "   TRUNC(1538.3518, 0) = 1538.0000" & vbCrLf & _
                                                            "   TRUNC(1538.3518, -1) = 1530.0000" & vbCrLf & _
                                                            "   TRUNC(1538.3518, -3) = 1000.0000"
            
        .AddItem "Text Functions":
        .Cell(flexcpFontBold, .Rows - 1, 0) = True
        .RowOutlineLevel(.Rows - 1) = 0
        .IsSubtotal(.Rows - 1) = True
            .AddItem "CHAR ( integer )                           " & vbTab & "Converts an int ASCII code to a character."
            .AddItem "CHARINDEX ( text1 , text2 )                " & vbTab & "Searches text2 for text1 and returns its starting" & vbCrLf & _
                                                                             "position if found."
            .AddItem "INSTR ( text1 , text2 )                    " & vbTab & "Searches text2 for text1 and returns its starting" & vbCrLf & _
                                                                             "position if found."
            .AddItem "LEFT ( text , integer )                    " & vbTab & "Returns the left part of a character string with" & vbCrLf & _
                                                                             "the specified number of characters"
            .AddItem "LEN ( text )                               " & vbTab & "Returns the number of characters in text, excluding" & vbCrLf & _
                                                                             "trailing blanks."
            .AddItem "LOWER ( text )                             " & vbTab & "Returns a character expression after converting" & vbCrLf & _
                                                                             "uppercase character data to lowercase."
            .AddItem "LTRIM ( text )                             " & vbTab & "Returns a character expression after it removes" & vbCrLf & _
                                                                             "leading blanks."
            .AddItem "PATINDEX ( '%pattern%' , expression )      " & vbTab & "Returns the starting position of the first occurrence" & vbCrLf & _
                                                                             "of a pattern in a specified expression, or zeros if the" & vbCrLf & _
                                                                             "pattern is not found, on all valid text and character" & vbCrLf & _
                                                                             "data types."
            .AddItem "REPLACE ( text , pattern , replacement )   " & vbTab & "Replaces all occurrences of a specified string value" & vbCrLf & _
                                                                             "with another string value."
            .AddItem "REPLICATE ( text , integer )               " & vbTab & "Repeats a string value a specified number of times."
            .AddItem "RIGHT ( text , integer )                   " & vbTab & "Returns the right part of a character string with the" & vbCrLf & _
                                                                             "specified number of characters."
            .AddItem "RTRIM ( text )                             " & vbTab & "Returns a character string after truncating all" & vbCrLf & _
                                                                             "trailing blanks."
            .AddItem "SPACE ( text )                             " & vbTab & "Returns a string of repeated spaces."
            .AddItem "STR ( numeric [ , length [ , decimals ] ] )" & vbTab & "Returns character data converted from numeric data."
            .AddItem "SUBSTRING ( text ,start , length )         " & vbTab & "Returns part of a text string"
            .AddItem "UPPER ( text )                             " & vbTab & "Returns a character expression with lowercase character" & vbCrLf & _
                                                                             "data converted to uppercase."
            
        .AddItem "Other Functions":
        .Cell(flexcpFontBold, .Rows - 1, 0) = True
        .RowOutlineLevel(.Rows - 1) = 0
        .IsSubtotal(.Rows - 1) = True
            .AddItem "PI                                         "
            .AddItem "e                                          "
            .AddItem "Rnd                                        " & vbTab & "Returns a random number valued between 0 and 1"
            .AddItem "IF( expression, truevalue, falsevalue)     "
            
            
            .AddItem "AssemblyOrderQty ( community, assembly, model, optionid, phase, item )   " & vbTab & "Returns the total order quantity of a phase/item in an assembly."
            .AddItem "AssemblyTakeoffQty ( community, assembly, model, optionid, phase, item )   " & vbTab & "Returns the total takeoff quantity of a phase/item in an assembly."
            
            
            .AddItem "TotalOrderQty ( poindex, costcode, phase, item )   " & vbTab & "Returns the total order quantity of a group of items in a takeoff." & vbCrLf & _
                                                                                     "All 4 parameters are filters which can be used to limit the results" & vbCrLf & _
                                                                                     "to a specific set of items."
                                                                             
            .AddItem "TotalTakeoffQty ( poindex, costcode, phase, item )   " & vbTab & "Returns the total takeoff quantity of a group of items in a takeoff." & vbCrLf & _
                                                                                       "All 4 parameters are filters which can be used to limit the results" & vbCrLf & _
                                                                                       "to a specific set of items."
                                                                                     
            .AddItem "TotalCost ( poindex, costcode, phase, item )   " & vbTab & "Returns the total cost of a group of items in a takeoff. All 4" & vbCrLf & _
                                                                                 "parameters are filters which can be used to limit the results to" & vbCrLf & _
                                                                                 "a specific set of items."
            
            
        Call .Outline(0)
    End With
End Sub

Private Function LoadVariables()
    Dim s As String
    Dim rs As Recordset
    
    With gVariables
        .Rows = 0
        .AddItem "Variables":
        .Cell(flexcpFontBold, .Rows - 1, 0) = True
        .RowOutlineLevel(.Rows - 1) = 0
        .IsSubtotal(.Rows - 1) = True
        Set rs = HFApp.SqlExec("select name,help from variables order by name")
        While Not rs.EOF
            .AddItem "" & rs("name") & vbTab & rs("help")
            rs.MoveNext
        Wend
    End With
    
End Function

Private Function LoadFormulas()
    Dim s As String
    Dim rs As Recordset
    
    With gFormulas
        .Rows = 0
        .AddItem "Formulas":
        .Cell(flexcpFontBold, .Rows - 1, 0) = True
        .RowOutlineLevel(.Rows - 1) = 0
        .IsSubtotal(.Rows - 1) = True
        Set rs = HFApp.SqlExec("select name,text from formulas order by name")
        While Not rs.EOF
            .AddItem "" & rs("name") & vbTab & rs("text")
            rs.MoveNext
        Wend
    End With
    
End Function












Private Sub Slider_Move(Index As Integer)
    Call Form_Resize
End Sub
Private Sub Form_Resize()
On Error Resume Next

    Slider(0).ZOrder 0
    Slider(1).ZOrder 0
    Slider(2).ZOrder 0

    Slider(0).Visible = True
    Slider(1).Visible = True
    Slider(2).Visible = True


    Slider(0).Min = 270
    Slider(0).Max = Slider(1).Left - 270
    Slider(1).Min = Slider(0).Left + 45 + 270
    Slider(1).Max = Me.ScaleWidth - 270
    Slider(2).Min = 270
    Slider(2).Max = Me.ScaleHeight - 270
    Slider(2).Move 0, Slider(2).Top, Me.ScaleWidth, 45
    Slider(0).Move Slider(0).Left, 0, 45, Slider(2).Top
    Slider(1).Move Slider(1).Left, 0, 45, Slider(2).Top

    gVariables.Move 0, 0, Slider(0).Left, Slider(0).Height
    gFormulas.Move Slider(0).Left + 45, 0, Slider(1).Left - Slider(0).Left - 45, Slider(0).Height
    gFunctions.Move Slider(1).Left + 45, 0, Me.ScaleWidth - Slider(1).Left - 45, Slider(0).Height
    
    txtFormula.Move 0, Slider(2).Top + 45, Me.ScaleWidth, Me.ScaleHeight - Slider(2).Top - 45 - cmdNav(0).Height - 120
    
    cmdNav(1).Move Me.ScaleWidth - 1 * (cmdNav(0).Width + 60), Me.ScaleHeight - cmdNav(0).Height - 60
    cmdNav(0).Move Me.ScaleWidth - 2 * (cmdNav(1).Width + 60), Me.ScaleHeight - cmdNav(0).Height - 60
    cmdNav(2).Move 60, Me.ScaleHeight - cmdNav(0).Height - 60
    
End Sub

Private Sub txtFormula_Change()
    Dim s As String
    
    
    'replace curly quotes with normal ascii double quote char
    
    s = txtFormula.Text
    If InStr(1, s, Chr(147)) Or InStr(1, s, Chr(148)) Or InStr(1, s, Chr(152)) Then
        s = Replace(s, Chr(150), "-")
        s = Replace(s, Chr(147), Chr(34))
        s = Replace(s, Chr(148), Chr(34))
        s = Replace(s, Chr(152), Chr(34))
        txtFormula.Text = s
    End If
    
End Sub

