VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{302C5C1A-C2E2-4302-9AF9-BAC87EEFDECE}#1.0#0"; "Panels.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form FIntersection 
   Caption         =   "Option Intersections"
   ClientHeight    =   9000
   ClientLeft      =   6435
   ClientTop       =   1875
   ClientWidth     =   20040
   Icon            =   "FIntersection.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   20040
   Begin Panels.Slider VSlider 
      Height          =   5910
      Left            =   6015
      Top             =   1710
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   10425
      Max             =   20220
   End
   Begin Panels.Slider HSlider 
      Height          =   60
      Left            =   6840
      Top             =   3720
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   106
      Orientation     =   1
      Max             =   20220
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20040
      _ExtentX        =   35348
      _ExtentY        =   1058
      ButtonWidth     =   1535
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Save"
            Key             =   "Save"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Delete"
            Key             =   "Delete"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Add"
            Key             =   "Add"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Item"
            Key             =   "TakeoffItem"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Assembly"
            Key             =   "TakeoffAssembly"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VSFlex8Ctl.VSFlexGrid gIntersections 
      Height          =   6015
      Left            =   705
      TabIndex        =   1
      Top             =   1425
      Width           =   4560
      _cx             =   8043
      _cy             =   10610
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   255
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FIntersection.frx":000C
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
      OutlineBar      =   5
      OutlineCol      =   1
      Ellipsis        =   0
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   2
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   1
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
   Begin VSFlex8Ctl.VSFlexGrid gAssemblies 
      Height          =   1920
      Left            =   6840
      TabIndex        =   2
      Top             =   1320
      Width           =   10005
      _cx             =   17648
      _cy             =   3387
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   255
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FIntersection.frx":007A
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
      OutlineBar      =   5
      OutlineCol      =   1
      Ellipsis        =   0
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   1
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   2
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   1
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
   Begin VSFlex8Ctl.VSFlexGrid gItems 
      Height          =   1920
      Left            =   6885
      TabIndex        =   3
      Top             =   4950
      Width           =   10005
      _cx             =   17648
      _cy             =   3387
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   255
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   16
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FIntersection.frx":0192
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
      OutlineBar      =   5
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   1
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   2
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   1
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
End
Attribute VB_Name = "FIntersection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FIntersection::"

Private mIntersectID As Long
Private mDirty As Boolean
Private mAssemblyID As Long
Private mAssemblyDesc As String


Private Property Let Dirty(RHS As Boolean)
    mDirty = RHS
    Toolbar.Buttons("Save").Enabled = mDirty
End Property

Private Sub Form_Load()
    Call IniGetForm(Me)
    Call SetToolbarIcons(Toolbar, FMain.LargeIcons)
    Call IniGetGrid(Me, gAssemblies)
    Call IniGetGrid(Me, gItems)
    Call LoadIntersections
End Sub

Private Sub Form_Resize()
On Error Resume Next

    VSlider.Min = 300
    VSlider.Max = Me.ScaleWidth - 300
    VSlider.Move Min(VSlider.Left, VSlider.Max), Toolbar.Height, VSlider.Width, Me.ScaleHeight - Toolbar.Height
    
    HSlider.Min = Toolbar.Height + 600
    HSlider.Max = Me.ScaleHeight - 600

    HSlider.Move VSlider.Left + VSlider.Width, Max(Min(HSlider.Top, HSlider.Max), HSlider.Min), Me.ScaleWidth - VSlider.Left - VSlider.Width
    gIntersections.Move 0, Toolbar.Height, VSlider.Left, Me.ScaleHeight - Toolbar.Height
    gAssemblies.Move VSlider.Left + VSlider.Width, Toolbar.Height, Me.ScaleWidth - VSlider.Left - VSlider.Width, HSlider.Top - Toolbar.Height
    gItems.Move gAssemblies.Left, HSlider.Top + HSlider.Height, gAssemblies.Width, Me.ScaleHeight - HSlider.Top - HSlider.Height
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Not SaveData(True) Then
        Cancel = True
        Exit Sub
    End If
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gAssemblies)
    Call IniPutGrid(Me, gItems)
End Sub

Private Sub gAssemblies_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
  
    If Button <> vbRightButton Then Exit Sub
    If gAssemblies.MouseRow < 1 And GridLastVisibleRow(gAssemblies) <> 0 Then
        Cancel = True
        Call FMain.ShowColumnMenu(gAssemblies, , , , , False)
    End If
    
End Sub

Private Sub gAssemblies_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
    With gAssemblies
        mAssemblyID = .ValueMatrix(NewRowSel, .ColIndex("AssemblyID"))
        mAssemblyDesc = .TextMatrix(NewRowSel, .ColIndex("Description"))
        Toolbar.Buttons("TakeoffItem").Enabled = mAssemblyID <> 0
        Toolbar.Buttons("TakeoffAssembly").Enabled = mAssemblyID <> 0
    End With
End Sub

Private Sub gIntersections_AfterSort(ByVal Col As Long, Order As Integer)
    gIntersections.AddItem ""
End Sub

Private Sub gIntersections_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gIntersections
    
        If Row = .Rows - 1 Then
            .AutoSearch = flexSearchNone
        Else
            .AutoSearch = flexSearchFromCursor
        End If
    End With
End Sub

Private Sub gIntersections_BeforeSort(ByVal Col As Long, Order As Integer)
    Call gIntersections.RemoveItem(gIntersections.Rows - 1)
End Sub

Private Sub gIntersections_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim s As String
    With gIntersections
        If .Row = .Rows - 1 Or .Row < 1 Then Exit Sub
        If KeyCode = vbKeyDelete And Shift <> 0 Then
            
            If vbNo = MsgBox("Are you sure you want to delete this option intersection?", vbQuestion + vbYesNo, App.ProductName) Then Exit Sub
        
            s = ""
            s = s & "delete tbldbassemblyintersectassemblies where intersectid=" & DbQuote(Num, .TextMatrix(.Row, .ColIndex("intersectid"))) & vbCrLf
            s = s & "delete tbldbassemblyintersectitems where intersectid=" & DbQuote(Num, .TextMatrix(.Row, .ColIndex("intersectid"))) & vbCrLf
            s = s & "delete tbldbassemblyintersections where intersectid=" & DbQuote(Num, .TextMatrix(.Row, .ColIndex("intersectid"))) & vbCrLf
            Call HFApp.SqlExec(s)
            Call .RemoveItem(.Row)
            Call .Select(GridNextVisibleRow(gIntersections, .Row), .Col)
            .SetFocus
            Call LoadIntersect(.ValueMatrix(.Row, .ColIndex("intersectid")))
        End If
    End With
End Sub
Private Sub gAssemblies_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim s As String
    Dim AssemblyID As Long
    Dim i As Long
    With gAssemblies
        If .Row < 1 Then Exit Sub
        If KeyCode = vbKeyDelete And Shift <> 0 Then
                        
            AssemblyID = .ValueMatrix(.Row, .ColIndex("AssemblyID"))
            For i = 1 To gItems.Rows - 1
            If AssemblyID = gItems.ValueMatrix(i, gItems.ColIndex("AssemblyID")) Then
                gItems.RowData(i) = "delete"
                gItems.RowHidden(i) = True
            End If
            Next
            .RowHidden(.Row) = True
            .RowData(.Row) = "delete"
            Dirty = True
            Call .Select(GridNextVisibleRow(gAssemblies, .Row), .Col)
            .SetFocus
            
        End If
    End With
End Sub

Private Sub gItems_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim r As Long
    With gItems
    For r = Min(.Row, .RowSel) To Max(.Row, .RowSel)
    If .RowData(r) = "" Then .RowData(r) = "dirty"
    Dirty = True
    Select Case .ColKey(Col)
        Case "TakeoffQty", "ConversionFactor"
            .TextMatrix(r, .ColIndex("OrderQty")) = RoundTo(.ValueMatrix(r, .ColIndex("TakeoffQty")) * (100 + .ValueMatrix(r, .ColIndex("WastePercent"))) / 100 * .ValueMatrix(r, .ColIndex("ConversionFactor")), .ValueMatrix(r, .ColIndex("RoundTo")), .ValueMatrix(r, .ColIndex("RoundDir")))
            
        Case "OrderQty"
            .TextMatrix(r, .ColIndex("TakeoffQty")) = .ValueMatrix(r, .ColIndex("OrderQty")) / .ValueMatrix(r, .ColIndex("ConversionFactor"))
            
    End Select
    Next
    End With
End Sub

Private Sub gItems_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    gItems.ColSel = gItems.Col
End Sub

Private Sub gItems_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gItems
        .ComboList = ""
        Select Case .ColKey(Col)
            Case "TakeoffQty", "OrderQty", "ConversionFactor", "Notes"
            Case "POIndex":     .ComboList = "|..."
            Case Else:          Cancel = True
        End Select
    End With
End Sub

Private Sub gItems_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    
    If Button <> vbRightButton Then Exit Sub
    If gItems.MouseRow < 1 And GridLastVisibleRow(gItems) <> 0 Then
        Cancel = True
        Call FMain.ShowColumnMenu(gItems, , , , , False)
    End If

End Sub

Private Sub gItems_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    Dim s As String
    With gItems
        Select Case .ColKey(Col)
            Case "POIndex"
                s = "SELECT POIndex,case when poindex=description then '' else Description end Description FROM tblPOIndex where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "PO Index", s, gItems) Then
                    For i = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                        .TextMatrix(i, Col) = FPickList.SelectedItem("POIndex")
                        If .RowData(i) = "" Then .RowData(i) = "dirty"
                    Next
                    Dirty = True
                End If
        End Select
    End With
End Sub

Private Sub gItems_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim s As String
    With gItems
        If .Row < 1 Then Exit Sub
        If KeyCode = vbKeyDelete And Shift <> 0 Then
            .RowData(.Row) = "delete"
            .RowHidden(.Row) = True
            Dirty = True
            Call .Select(GridNextVisibleRow(gItems, .Row), .Col)
            .SetFocus
        End If
    End With
End Sub

Private Sub gIntersections_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo eh
    Dim s As String
    With gIntersections
       .EditText = Trim(.EditText)
        If .EditText = "" Then
            Cancel = True
            Exit Sub
        End If
    
        If Row = .Rows - 1 Then
            'add
            s = "insert tbldbassemblyintersections(divisionid,description) values(" & DbQuote(Num, HFApp.DivisionID) & "," & DbQuote(Str, .EditText) & ")"
            Call HFApp.SqlExec(s)
            .TextMatrix(Row, .ColIndex("intersectid")) = HFApp.SqlIdentity("tbldbassemblyintersections")
            Call LoadIntersect(.ValueMatrix(Row, .ColIndex("intersectid")))
            .AddItem ""
        Else
            'add
            s = "update tbldbassemblyintersections set description=" & DbQuote(Str, .EditText) & " where intersectid=" & DbQuote(Num, .TextMatrix(Row, .ColIndex("intersectid")))
            Call HFApp.SqlExec(s)
        End If
    End With
Exit Sub
eh: Call errHandler(SRCFILE & "gIntersections_ValidateEdit")
End Sub

Private Sub gItems_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim s As String
    With gItems
        s = .EditText
        Select Case .ColKey(Col)
            Case "TakeoffQty", "OrderQty", "ConversionFactor"
                s = Val(s)
            Case "Notes"
            Case "POIndex"
                Cancel = Not ValidateField(gItems, s, "PO Index not found", "SELECT POIndex FROM tblPOIndex WHERE DivisionID = " & HFApp.DivisionID & " and POIndex=" & DbQuote(Str, s))
        End Select
        .EditText = s
    End With
End Sub

Private Sub HSlider_Move()
    Call Form_Resize
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
    
    Case "Add"
        Call AddAssembly
        
    Case "Save"
        Call SaveData(False)
        
    Case "TakeoffOneTime", "TakeoffItem", "TakeoffAssembly", "TakeoffCustom", "TakeoffPlanSwift"
        Call FTakeoff.Takeoff(Me, False, Mid(Button.Key, 8), 0, "", atoption, "", "", "", "", "", "", "")
        
    Case "Delete"
        If Screen.ActiveControl Is gIntersections Then Call gIntersections_KeyDown(vbKeyDelete, 1)
        If Screen.ActiveControl Is gAssemblies Then Call gAssemblies_KeyDown(vbKeyDelete, 1)
        If Screen.ActiveControl Is gItems Then Call gItems_KeyDown(vbKeyDelete, 1)
            
    
    End Select
    
End Sub

Private Function AssemblyExists(AssemblyID As Long) As Boolean
    'return true if assemblyid exists in gAssemblies
    Dim i As Long
    With gAssemblies
    For i = 1 To .Rows - 1
        If .ValueMatrix(i, .ColIndex("AssemblyID")) = AssemblyID And .RowData(i) <> "delete" Then
            AssemblyExists = True
            Exit Function
        End If
    Next
    End With
End Function


Private Sub AddAssembly()
    Dim i As Long
    Dim r As Long
    Dim s As String
    
    Dim Model As String
    With gAssemblies
    For i = 1 To .Rows - 1
        If .TextMatrix(i, .ColIndex("model")) <> "" And Not .RowHidden(i) Then Model = .TextMatrix(i, .ColIndex("model"))
    Next
                    
    s = ""
    If Model = "" Then
        s = s & "Models" & Chr(1)
        s = s & "select a.Community,a.AssemblyID,a.Assembly,a.Model,a.Description,a.assemblytypedesc" & vbCrLf
        s = s & "from tbldbassemblymaster a" & vbCrLf
        s = s & "where a.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        s = s & "and isnull(a.inactive,0)=0" & vbCrLf
        s = s & "and a.assemblytype=0" & vbCrLf & Chr(0)
    End If
    s = s & "Model Options" & Chr(1)
    s = s & "select a.Category,a.Community,a.AssemblyID,a.Assembly,a.Model,a.OptionID,a.Description,a.assemblytypedesc" & vbCrLf
    s = s & "from tbldbassemblymaster a" & vbCrLf
    s = s & "where a.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
    s = s & "and isnull(a.inactive,0)=0" & vbCrLf
    If Model <> "" Then
        s = s & "and a.model=" & DbQuote(Str, Model) & vbCrLf
    End If
    s = s & "and a.assemblytype=2" & vbCrLf & Chr(0)
    s = s & "Global Options" & Chr(1)
    s = s & "select a.Category,a.Community,a.AssemblyID,a.Assembly,a.OptionID,a.Description,a.assemblytypedesc" & vbCrLf
    s = s & "from tbldbassemblymaster a" & vbCrLf
    s = s & "where a.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
    s = s & "and isnull(a.inactive,0)=0" & vbCrLf
    s = s & "and a.assemblytype=3" & vbCrLf & Chr(0)
    s = s & "Design Center Options" & Chr(1)
    s = s & "select a.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
    s = s & "and a.Category,a.Community,a.AssemblyID,a.Assembly,a.OptionID,a.Description,a.assemblytypedesc" & vbCrLf
    s = s & "from tbldbassemblymaster a" & vbCrLf
    s = s & "where isnull(a.inactive,0)=0" & vbCrLf
    s = s & "and a.assemblytype=4"

    If Not FPickList.Choose(HFApp.Databases(dbHomefront), "Assembly", s, , , , , "AssemblyID,assemblytypedesc", True) Then Exit Sub
    For i = 1 To FPickList.SelectedItems
    If AssemblyExists(FPickList.SelectedItem("AssemblyID", i)) Then
        MsgBox """" & FPickList.SelectedItem("Description", i) & """ has already been added.", vbInformation, App.ProductName
    Else
    
        Dirty = True
        .AddItem ""
        r = .Rows - 1
        .RowData(r) = "new"
        .TextMatrix(r, .ColIndex("AssemblyID")) = FPickList.SelectedItem("AssemblyID", i)
        .TextMatrix(r, .ColIndex("AssemblyType")) = FPickList.SelectedItem("AssemblyTypedesc", i)
        .TextMatrix(r, .ColIndex("Community")) = FPickList.SelectedItem("Community", i)
        .TextMatrix(r, .ColIndex("Assembly")) = FPickList.SelectedItem("Assembly", i)
        .TextMatrix(r, .ColIndex("Model")) = FPickList.SelectedItem("Model", i)
        .TextMatrix(r, .ColIndex("Option")) = FPickList.SelectedItem("OptionID", i)
        .TextMatrix(r, .ColIndex("Description")) = FPickList.SelectedItem("Description", i)
        
        'cant prevent user from selecting 3 models at once, so only accept the first.
        If .TextMatrix(r, .ColIndex("AssemblyType")) = "Model" Then Exit Sub
        
        
        .Row = r
    End If
    Next
    
    End With

End Sub
Private Sub VSlider_Move()
    Call Form_Resize
End Sub

Private Sub LoadIntersections()
    Dim s As String
    Dim rs As Recordset
    Dim r As Long
    
    With gIntersections
        .Rows = 1
        
        s = "select Intersectid,description from tbldbassemblyintersections where divisionid=" & DbQuote(Num, HFApp.DivisionID) & " order by description"
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
            .AddItem ""
            r = .Rows - 1
            .TextMatrix(r, .ColIndex("intersectid")) = "" & rs("intersectid")
            .TextMatrix(r, .ColIndex("description")) = "" & rs("description")
            rs.MoveNext
        Wend
        .AddItem ""
    End With
    
End Sub
Private Sub gIntersections_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
With gIntersections
    If .Rows = 1 Then Exit Sub
    If OldRowSel > 0 Then
        Cancel = Not SaveData(True)
        If Cancel Then Exit Sub
    End If
    Call LoadIntersect(.ValueMatrix(NewRowSel, .ColIndex("intersectid")))
End With
End Sub

Private Sub LoadIntersect(IntersectID As Long)
    Dim s As String
    Dim rs As Recordset
    Dim r As Long
    
    mIntersectID = IntersectID
    Toolbar.Buttons("TakeoffItem").Enabled = False
    Toolbar.Buttons("TakeoffAssembly").Enabled = False
    
    With gAssemblies
        .Rows = 1
        
        s = ""
        s = s & "select m.assemblyid,m.community,m.assembly,m.model,m.optionid,m.description,m.assemblytypedesc" & vbCrLf
        s = s & "from tbldbassemblyintersectassemblies xa" & vbCrLf
        s = s & "join tbldbassemblymaster m on xa.assemblyid=m.assemblyid" & vbCrLf
        s = s & "where xa.intersectid=" & DbQuote(Str, IntersectID) & vbCrLf
        s = s & "order by m.description"
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
            .AddItem ""
            r = .Rows - 1
            .TextMatrix(r, .ColIndex("assemblyid")) = "" & rs("assemblyid")
            .TextMatrix(r, .ColIndex("community")) = "" & rs("community")
            .TextMatrix(r, .ColIndex("assemblytype")) = "" & rs("assemblytypedesc")
            .TextMatrix(r, .ColIndex("assembly")) = "" & rs("assembly")
            .TextMatrix(r, .ColIndex("model")) = "" & rs("model")
            .TextMatrix(r, .ColIndex("option")) = "" & rs("optionid")
            .TextMatrix(r, .ColIndex("description")) = "" & rs("description")
            rs.MoveNext
        Wend
    End With

    With gItems
        .Rows = 1
        
        s = ""
        s = s & "select  m.description assemblydesc,i.itemdesc,i.wastepercent,i.rounddir,i.roundto,xi.*" & vbCrLf
        s = s & "from tblDBAssemblyIntersections x" & vbCrLf
        s = s & "join tblDBAssemblyIntersectItems xi on x.intersectid=xi.intersectid" & vbCrLf
        s = s & "join tblDBAssemblyMaster m on xi.assemblyid=m.assemblyid" & vbCrLf
        s = s & "join EstimatingItems i on i.divisionid=x.divisionid and i.phase=xi.phase and i.item=xi.item" & vbCrLf
        s = s & "where x.intersectid=" & DbQuote(Str, IntersectID) & vbCrLf
        s = s & "order by xi.itemid"
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
            .AddItem ""
            r = .Rows - 1
            .TextMatrix(r, .ColIndex("itemid")) = "" & rs("itemid")
            .TextMatrix(r, .ColIndex("assemblyid")) = "" & rs("assemblyid")
            .TextMatrix(r, .ColIndex("assemblydesc")) = "" & rs("assemblydesc")
            .TextMatrix(r, .ColIndex("phase")) = "" & rs("phase")
            .TextMatrix(r, .ColIndex("item")) = "" & rs("item")
            .TextMatrix(r, .ColIndex("takeoffqty")) = "" & rs("takeoffqty")
            .TextMatrix(r, .ColIndex("OrderQty")) = "" & rs("OrderQty")
            .TextMatrix(r, .ColIndex("ConversionFactor")) = "" & rs("ConversionFactor")
            .TextMatrix(r, .ColIndex("wastepercent")) = "" & rs("wastepercent")
            .TextMatrix(r, .ColIndex("rounddir")) = "" & rs("rounddir")
            .TextMatrix(r, .ColIndex("roundto")) = "" & rs("roundto")
            .TextMatrix(r, .ColIndex("takeoffUOM")) = "" & rs("takeoffUOM")
            .TextMatrix(r, .ColIndex("OrderUOM")) = "" & rs("OrderUOM")
            .TextMatrix(r, .ColIndex("poindex")) = "" & rs("poindex")
            .TextMatrix(r, .ColIndex("notes")) = "" & rs("notes")
            .TextMatrix(r, .ColIndex("itemdesc")) = "" & rs("itemdesc")
            rs.MoveNext
        Wend
    End With
    Dirty = False
    Toolbar.Buttons("Add").Enabled = IntersectID <> 0
    Toolbar.Buttons("Delete").Enabled = IntersectID <> 0

End Sub

Private Function SaveData(prompt As Boolean) As Boolean
On Error GoTo eh

    Dim s As String
    Dim r As Long

    If Not mDirty Then
        SaveData = True
        Exit Function
    End If
    If prompt Then
        Select Case MsgBox("This data has changed." & vbCrLf & vbCrLf & "Do you want to save these changes?" & vbCrLf, vbExclamation + vbYesNoCancel, App.ProductName)
            Case vbNo
                SaveData = True
                Exit Function
            Case vbCancel
                SaveData = False
                Exit Function
        End Select
    End If
    
    With gAssemblies
    For r = .Rows - 1 To 1 Step -1
        Select Case .RowData(r)
        Case "delete"
            s = ""
            s = s & "delete tbldbassemblyintersectassemblies where intersectid=" & DbQuote(Num, mIntersectID) & " and assemblyid=" & DbQuote(Num, .TextMatrix(r, .ColIndex("assemblyid"))) & vbCrLf
            s = s & "delete tbldbassemblyintersectitems where intersectid=" & DbQuote(Num, mIntersectID) & " and assemblyid=" & DbQuote(Num, .TextMatrix(r, .ColIndex("assemblyid"))) & vbCrLf
            Call HFApp.SqlExec(s)
            Call .RemoveItem(r)
        End Select
    Next
    For r = 1 To .Rows - 1
        Select Case .RowData(r)
        Case "new"
            s = ""
            s = s & "insert tbldbassemblyintersectassemblies(intersectid,assemblyid) values(" & DbQuote(Num, mIntersectID) & "," & DbQuote(Num, .TextMatrix(r, .ColIndex("assemblyid"))) & ")"
            Call HFApp.SqlExec(s)
            .RowData(r) = ""
        
        End Select
    Next
    End With


    With gItems
    For r = .Rows - 1 To 1 Step -1
        Select Case .RowData(r)
        Case "delete"
            s = "delete tbldbassemblyintersectitems where itemid=" & DbQuote(Num, .TextMatrix(r, .ColIndex("itemid")))
            Call HFApp.SqlExec(s)
            Call .RemoveItem(r)
        End Select
    Next
    For r = 1 To .Rows - 1
        Select Case .RowData(r)
            
        Case "new"
            s = ""
            s = s & "insert tbldbassemblyintersectitems(IntersectID, AssemblyID, Phase, Item, TakeoffQty, TakeoffUOM, ConversionFactor, OrderQty, OrderUOM, POIndex, Notes) values(" & vbCrLf
            s = s & DbQuote(Num, mIntersectID)
            s = s & "," & DbQuote(Num, .TextMatrix(r, .ColIndex("AssemblyID")))
            s = s & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("Phase")))
            s = s & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("Item")))
            s = s & "," & DbQuote(Num, .TextMatrix(r, .ColIndex("TakeoffQty")))
            s = s & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("TakeoffUOM")))
            s = s & "," & DbQuote(Num, .TextMatrix(r, .ColIndex("ConversionFactor")))
            s = s & "," & DbQuote(Num, .TextMatrix(r, .ColIndex("OrderQty")))
            s = s & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("OrderUOM")))
            s = s & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("POIndex")))
            s = s & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("Notes")))
            s = s & ")"
            Call HFApp.SqlExec(s)
            .TextMatrix(r, .ColIndex("ItemID")) = HFApp.SqlIdentity("tbldbassemblyintersectitems")
            .RowData(r) = ""
        
        Case "dirty"
            s = ""
            s = s & "update tbldbassemblyintersectitems set"
            s = s & " takeoffqty=" & DbQuote(Num, .TextMatrix(r, .ColIndex("TakeoffQty")))
            s = s & ",conversionfactor=" & DbQuote(Num, .TextMatrix(r, .ColIndex("ConversionFactor")))
            s = s & ",orderqty=" & DbQuote(Num, .TextMatrix(r, .ColIndex("OrderQty")))
            s = s & ",poindex=" & DbQuote(Str, .TextMatrix(r, .ColIndex("POIndex")))
            s = s & ",notes=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Notes")))
            s = s & "where itemid=" & DbQuote(Num, .TextMatrix(r, .ColIndex("ItemID")))
            Call HFApp.SqlExec(s)
            .RowData(r) = ""
        
        End Select
    Next
    End With

    Dirty = False
    SaveData = True
Exit Function
eh: Call errHandler(SRCFILE & "SaveIntersect", s)
End Function

Public Sub AddItem(Assembly As String, _
                   AssemblyDescription As String, _
                   Model As String, _
                   Phase As String, _
                   Item As String, Sequence As Long, _
                   Description As String, _
                   OrderQty As Double, _
                   OrderUOM As String, _
                   TakeoffQty As Double, _
                   TakeoffUOM As String, _
                   ConversionFactor As Double, RoundTo As Double, RoundDir As Long, WastePercent As Long, _
                   JCExtra As String, _
                   JCCostCode As String, _
                   JCCostCodeDesc As String, JCCategory As String, _
                   JCCategoryDesc As String, Vendor As String, _
                   VendorName As String, price As Double, _
                   TaxGroup As String, TaxGroupName As String, _
                   JCTaxRate As Double, NJCTaxRate As Double, _
                   POIndex As String, _
                   Comments As String, Formula As String, SalesQty As Double, Location As String, _
                   WBS, Optional Job As String)
On Error GoTo eh
    Dim r As Long
    With gItems
        .AddItem ""
        r = .Rows - 1
        If .RowData(r) = "" Then .RowData(r) = "new"
        .TextMatrix(r, .ColIndex("assemblyid")) = mAssemblyID
        .TextMatrix(r, .ColIndex("assemblydesc")) = mAssemblyDesc
        .TextMatrix(r, .ColIndex("phase")) = Phase
        .TextMatrix(r, .ColIndex("item")) = Item
        .TextMatrix(r, .ColIndex("takeoffqty")) = TakeoffQty
        .TextMatrix(r, .ColIndex("ConversionFactor")) = ConversionFactor
        .TextMatrix(r, .ColIndex("OrderQty")) = OrderQty
        .TextMatrix(r, .ColIndex("takeoffuom")) = TakeoffUOM
        .TextMatrix(r, .ColIndex("OrderUOM")) = OrderUOM
        .TextMatrix(r, .ColIndex("poindex")) = POIndex
        .TextMatrix(r, .ColIndex("itemdesc")) = Description
        .TextMatrix(r, .ColIndex("WastePercent")) = WastePercent
        .TextMatrix(r, .ColIndex("RoundDir")) = RoundDir
        .TextMatrix(r, .ColIndex("RoundTo")) = RoundTo
        
    End With
    Dirty = True

Exit Sub
eh: Stop
End Sub

