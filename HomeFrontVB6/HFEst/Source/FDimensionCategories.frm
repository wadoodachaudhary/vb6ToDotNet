VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{302C5C1A-C2E2-4302-9AF9-BAC87EEFDECE}#1.0#0"; "Panels.ocx"
Begin VB.Form FDimensionCategories 
   Caption         =   "Dimension Categories"
   ClientHeight    =   6270
   ClientLeft      =   10110
   ClientTop       =   4770
   ClientWidth     =   14415
   Icon            =   "FDimensionCategories.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   14415
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear &All"
      Height          =   330
      Left            =   12855
      TabIndex        =   5
      Top             =   135
      Width           =   1035
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Save"
      Height          =   375
      Index           =   0
      Left            =   11310
      TabIndex        =   1
      Top             =   5145
      Width           =   1095
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&Close"
      Height          =   375
      Index           =   1
      Left            =   12480
      TabIndex        =   0
      Top             =   5145
      Width           =   1095
   End
   Begin VSFlex8Ctl.VSFlexGrid gRooms 
      Height          =   4365
      Left            =   9135
      TabIndex        =   2
      Top             =   480
      Width           =   4755
      _cx             =   8387
      _cy             =   7699
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
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   3
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FDimensionCategories.frx":000C
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
      OwnerDraw       =   0
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
   Begin VSFlex8Ctl.VSFlexGrid gCategories 
      Height          =   4365
      Left            =   180
      TabIndex        =   3
      Top             =   480
      Width           =   7140
      _cx             =   12594
      _cy             =   7699
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
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FDimensionCategories.frx":005B
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
      OwnerDraw       =   0
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
   Begin VSFlex8Ctl.VSFlexGrid gUnits 
      Height          =   4365
      Left            =   7560
      TabIndex        =   4
      Top             =   480
      Width           =   1365
      _cx             =   2408
      _cy             =   7699
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
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FDimensionCategories.frx":00E7
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
      OwnerDraw       =   0
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
   Begin Panels.Slider Slider 
      Height          =   1860
      Index           =   0
      Left            =   7395
      Top             =   480
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   3281
   End
   Begin Panels.Slider Slider 
      Height          =   1860
      Index           =   1
      Left            =   8985
      Top             =   480
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   3281
   End
End
Attribute VB_Name = "FDimensionCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FDimensionCategories::"

Private mDirty As Boolean

Private mCategory As String
Private mUOM As String



Private Sub cmdClear_Click()
On Error Resume Next
    Dim i As Long
    With gRooms
        .Cell(flexcpChecked, 1, .ColIndex("room"), .Rows - 1, .ColIndex("room")) = flexUnchecked
        Dirty = True
    End With
End Sub

Private Sub cmdNav_Click(Index As Integer)
    If Index = 0 Then
        Call SaveData(False)
    Else
        If SaveData(True) Then
            Dirty = False
            Unload Me
        End If
    End If
End Sub

Public Sub ShowForm()
    Me.Show vbModal
End Sub

Private Function GetComboList(sql As String, Optional LimitToList As Boolean = True) As String
    Dim i As Long
    Dim s As String
    Dim rs As Recordset
    
    Set rs = HFApp.SqlExec(sql, dbHomefront)
    s = ""
    While Not rs.EOF
        If Trim("" & rs(0)) <> "" Then s = s & "|" & rs(0)
        rs.MoveNext
    Wend
    If LimitToList Then s = Mid(s, 2)
    If s = "|" Then s = ""
    GetComboList = s
    
End Function



Private Property Let Dirty(RHS As Boolean)
    mDirty = RHS
    cmdNav(0).Enabled = mDirty
End Property

Private Sub Form_Load()
    Call IniGetForm(Me)
    Call IniGetGrid(Me, gCategories)
    Call LoadCategories
    Call LoadRooms("")
    
    gUnits.ComboList = GetComboList("SELECT DISTINCT ISNULL(AssemblyUOM,''),'',0 FROM tblDBAssemblyMaster where ISNULL(AssemblyUOM,'')<>'' and divisionid=" & DbQuote(Num, HFApp.DivisionID) & " ORDER BY 1", False)

End Sub

Private Sub Form_Resize()
On Error Resume Next
    Const margin = 120
    
    Slider(0).Min = 1800
    Slider(1).Min = Slider(0).Left + 1800
    Slider(0).Max = Slider(1).Left - 1800
    Slider(1).Max = Me.ScaleWidth - 1200
    Slider(0).Height = Me.ScaleHeight - Slider(0).Top - 2 * margin - cmdNav(0).Height
    Slider(1).Height = Slider(0).Height
    Slider(0).ZOrder 0
    Slider(1).ZOrder 0
    
    gCategories.Move gCategories.Left, Slider(0).Top, Slider(0).Left - gCategories.Left, Slider(0).Height
    gUnits.Move Slider(0).Left + Slider(0).Width, Slider(0).Top, Slider(1).Left - Slider(0).Left - Slider(0).Width, Slider(0).Height
    gRooms.Move Slider(1).Left + Slider(1).Width, Slider(1).Top, Me.ScaleWidth - margin - Slider(1).Left - Slider(1).Width, Slider(1).Height
    cmdClear.Move gRooms.Left + gRooms.Width - cmdClear.Width, gRooms.Top - cmdClear.Height
    
    cmdNav(0).Move Me.ScaleWidth - (2 * (margin + cmdNav(0).Width)), Me.ScaleHeight - margin - cmdNav(0).Height
    cmdNav(1).Move Me.ScaleWidth - (1 * (margin + cmdNav(0).Width)), Me.ScaleHeight - margin - cmdNav(0).Height

End Sub



Private Sub Form_Unload(Cancel As Integer)
    If Not SaveData(True) Then
        Cancel = True
        Exit Sub
    End If
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gCategories)
End Sub




Private Sub gCategories_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
    Static inHere As Boolean
    If inHere Then Exit Sub
    inHere = True
    
    With gCategories
        If .Rows = 1 Then GoTo ExitSub
        If OldRowSel > 0 Then
            Cancel = Not SaveData(False)
            If Cancel Then GoTo ExitSub
        End If
        mCategory = ""
        mUOM = ""
        Call .Select(NewRowSel, 0)
        Call LoadUnits(.TextMatrix(NewRowSel, .ColIndex("Category")))
    End With

ExitSub:
    inHere = False
End Sub



Private Sub LoadCategories()
    Dim s As String
    Dim rs As Recordset
    Dim r As Long
    
    With gCategories
        .Rows = 1
        
        s = ""
        s = s & "select c.Category,g.description GroupDescription,c.Description CategoryDescription" & vbCrLf
        s = s & "from tblMajorGroups g" & vbCrLf
        s = s & "join tblCategories c on g.major_group=c.group_code" & vbCrLf
        s = s & "order by g.description,c.description" & vbCrLf
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
            .AddItem ""
            r = .Rows - 1
            .TextMatrix(r, .ColIndex("category")) = "" & rs("category")
            .TextMatrix(r, .ColIndex("groupdescription")) = "" & rs("groupdescription")
            .TextMatrix(r, .ColIndex("categorydescription")) = "" & rs("categorydescription")
            rs.MoveNext
        Wend
        
        If .Rows > 1 Then .Row = 1
    End With
        
    
End Sub
Private Sub LoadUnits(Category As String)
    Dim s As String
    Dim rs As Recordset
    Dim r As Long
    
    mCategory = Category
    
    With gUnits
        .Rows = 1
        
        
        s = ""
        s = s & "select distinct c.uom" & vbCrLf
        s = s & "from RoomMaster r" & vbCrLf
        s = s & "join RoomMasterSubCategory c on r.roomid=c.roomid and c.subcategory=" & DbQuote(Str, mCategory) & vbCrLf
        s = s & "where isnull(r.inactive,0)=0 and r.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        s = s & "order by c.uom"
        
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
            .AddItem ""
            r = .Rows - 1
            .TextMatrix(r, .ColIndex("uom")) = "" & rs("uom")
            rs.MoveNext
        Wend
    
        If .Rows > 1 Then .Row = 1
        If .Rows < 2 And gRooms.Rows > 1 Then
            gRooms.Cell(flexcpChecked, 1, gRooms.ColIndex("room"), gRooms.Rows - 1, gRooms.ColIndex("room")) = flexUnchecked
        End If
        
        .AddItem ""
        
    End With

    Dirty = False
    
End Sub

Private Sub LoadRooms(UOM As String)
    Dim s As String
    Dim rs As Recordset
    Dim r As Long
    
    mUOM = UOM
    
    With gRooms
       
        
        s = ""
        s = s & "select r.roomid,r.room" & vbCrLf
        s = s & "     , cast(case when c.roomid is null then 0 else 1 end as bit) selected" & vbCrLf
        s = s & "from RoomMaster r" & vbCrLf
        s = s & "left outer join RoomMasterSubCategory c on r.roomid=c.roomid " & vbCrLf
        s = s & "                                       and c.subcategory=" & DbQuote(Str, mCategory) & vbCrLf
        s = s & "                                       and c.uom=" & DbQuote(Str, mUOM) & vbCrLf
        s = s & "where isnull(r.inactive,0)=0 and r.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        s = s & "order by r.Room"
        Set rs = HFApp.SqlExec(s)
        .Rows = 1
        .Redraw = flexRDNone
        While Not rs.EOF
            .AddItem ""
            r = .Rows - 1
            .TextMatrix(r, .ColIndex("roomid")) = "" & rs("roomid")
            .TextMatrix(r, .ColIndex("room")) = "" & rs("room")
            .Cell(flexcpChecked, r, .ColIndex("room")) = IIf(rs("selected"), flexChecked, flexUnchecked)
            rs.MoveNext
        Wend
        .Redraw = flexRDBuffered
        Dirty = False
        
    End With


ExitSub:
    Dirty = False
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


    With gRooms
        s = ""
        s = s & "delete c from roommastersubcategory c join roommaster r on c.roomid=r.roomid" & vbCrLf
        s = s & "where c.Subcategory=" & DbQuote(Str, mCategory) & " and c.uom=" & DbQuote(Str, mUOM) & vbCrLf
        s = s & "and r.divisionid=" & HFApp.DivisionID & vbCrLf
        For r = 1 To .Rows - 1
        If .Cell(flexcpChecked, r, .ColIndex("room")) = flexChecked And mUOM <> "" Then
            s = s & "insert RoomMasterSubCategory(SubCategory,UOM,RoomID) values(" & DbQuote(Str, mCategory) & "," & DbQuote(Str, mUOM) & "," & DbQuote(Num, .TextMatrix(r, .ColIndex("roomid"))) & ")" & vbCrLf
        End If
        Next
        Call HFApp.SqlExec(s)
    End With

    Dirty = False
    SaveData = True

Exit Function
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Call errHandler(SRCFILE & "SaveData", s)
    End If
End Function



Private Sub gRooms_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gUnits
        Cancel = .TextMatrix(.Row, .ColIndex("UOM")) = "" Or .Row < 1
    End With
End Sub

Private Sub gRooms_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dirty = True
End Sub

Private Sub gUnits_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With gUnits
        If Row <> .Rows - 1 Then Exit Sub
        If .TextMatrix(Row, Col) = "" Then Exit Sub
        .AddItem ""
        mUOM = .TextMatrix(Row, Col)
        Dirty = True
    End With
End Sub

Private Sub gUnits_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Row < gUnits.Rows - 1
End Sub

Private Sub gUnits_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
    Static inHere As Boolean
    If inHere Then Exit Sub
    inHere = True
    
    With gUnits
        If .Rows = 1 Then GoTo ExitSub
        If OldRowSel > 0 Then
            Cancel = Not SaveData(False)
            If Cancel Then GoTo ExitSub
        End If
        Call .Select(NewRowSel, 0)
        Call LoadRooms(.TextMatrix(NewRowSel, .ColIndex("UOM")))
    End With

ExitSub:
    inHere = False
End Sub

Private Sub gUnits_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim s As String
    
    If Not (KeyCode = vbKeyDelete And Shift <> 0) Then Exit Sub
    If MsgBox("Are you sure you want to delete this UOM from this category?", vbQuestion + vbYesNo, App.ProductName) = vbNo Then Exit Sub
    With gUnits
        If .Row < 1 Then Exit Sub
        s = "delete RoomMasterSubCategory where Subcategory=" & DbQuote(Str, mCategory) & " and uom=" & DbQuote(Str, mUOM)
        mUOM = ""
        Call HFApp.SqlExec(s)
        Call .RemoveItem(.Row)
        gUnits.Row = 0
    End With

End Sub

Private Sub gUnits_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim i As Long
    
    With gUnits
        For i = 1 To .Rows - 2
            If .TextMatrix(i, Col) = .EditText Then
                Cancel = True
                Exit Sub
            End If
        Next
    End With

End Sub

Private Sub Slider_Move(Index As Integer)
    Call Form_Resize
End Sub
