VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{302C5C1A-C2E2-4302-9AF9-BAC87EEFDECE}#1.0#0"; "Panels.ocx"
Begin VB.Form FModelDimensions 
   Caption         =   "Model Dimensions"
   ClientHeight    =   5385
   ClientLeft      =   7875
   ClientTop       =   2745
   ClientWidth     =   15210
   Icon            =   "FModelDimensions.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   15210
   Begin Panels.Slider Slider 
      Height          =   1860
      Left            =   5895
      Top             =   180
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   3281
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&Save"
      Height          =   375
      Index           =   0
      Left            =   9705
      TabIndex        =   1
      Top             =   4650
      Width           =   1095
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&Close"
      Height          =   375
      Index           =   1
      Left            =   10875
      TabIndex        =   0
      Top             =   4650
      Width           =   1095
   End
   Begin VSFlex8Ctl.VSFlexGrid gModels 
      Height          =   4365
      Left            =   180
      TabIndex        =   2
      Top             =   180
      Width           =   5655
      _cx             =   9975
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
      FormatString    =   $"FModelDimensions.frx":000C
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
      OutlineBar      =   5
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   7
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
      Left            =   6045
      TabIndex        =   3
      Top             =   165
      Width           =   5985
      _cx             =   10557
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
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FModelDimensions.frx":006C
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
   Begin VB.Menu mnuModels 
      Caption         =   "mnuModels"
      Visible         =   0   'False
      Begin VB.Menu mnuModelsSub 
         Caption         =   "Add Model..."
         Index           =   0
      End
      Begin VB.Menu mnuModelsSub 
         Caption         =   "Add Room..."
         Index           =   1
      End
      Begin VB.Menu mnuModelsSub 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuModelsSub 
         Caption         =   "Delete"
         Index           =   3
      End
   End
   Begin VB.Menu mnuCategories 
      Caption         =   "mnuCategories"
      Visible         =   0   'False
      Begin VB.Menu mnuCategoriesSub 
         Caption         =   "Add Room..."
         Index           =   0
      End
      Begin VB.Menu mnuCategoriesSub 
         Caption         =   "Remove Room"
         Index           =   1
      End
   End
End
Attribute VB_Name = "FModelDimensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FModelDimensions::"

Private mOrigModel As String
Private mModel As String
Private mDirty As Boolean

Const mnuMODEL_ADDMODEL = 0
Const mnuMODEL_ADDROOM = 1
Const mnuMODEL_DELETE = 3

Const mnuCATEGORIES_ADDROOM = 0
Const mnuCATEGORIES_REMOVEROOM = 1


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

Public Sub ShowForm(Optional Model As String = "")
    mOrigModel = Model
    mModel = Model
    Me.Show vbModal
End Sub


Private Property Let Dirty(RHS As Boolean)
    mDirty = RHS
    cmdNav(0).Enabled = mDirty
End Property

Private Sub Form_Load()
    Call IniGetForm(Me)
    Call IniGetGrid(Me, gCategories, , , , , True)
    Call IniGetGrid(Me, gModels, , , , , True)
    Call LoadModels
End Sub


Private Sub Form_Resize()
On Error Resume Next
    Const margin = 120
    
    Slider.Min = 1200
    Slider.Max = Me.ScaleWidth - 1200
    Slider.Height = Me.ScaleHeight - Slider.Top - 2 * margin - cmdNav(0).Height
    
    gModels.Move gModels.Left, Slider.Top, Slider.Left - gModels.Left, Slider.Height
    gCategories.Move Slider.Left + Slider.Width, Slider.Top, Me.ScaleWidth - margin - Slider.Left - Slider.Width, Slider.Height
    
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
    Call IniPutGrid(Me, gModels)
End Sub






Private Sub gCategories_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gCategories
    Select Case .ColKey(Col)
    Case "Qty"
    Case Else
        Cancel = True
    End Select
    End With
End Sub

Private Sub gCategories_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gCategories
        .EditText = Val(.EditText)
        .RowData(Row) = "dirty"
    End With
End Sub

Private Sub gModels_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim s As String
    With gModels
        If .Row < 1 Then Exit Sub
        If KeyCode = vbKeyDelete And Shift <> 0 Then
            DeleteModel
        End If
    End With
End Sub
Private Sub DeleteModel()
    Dim s As String
    If vbNo = MsgBox("Are you sure you want to delete the dimensions for this model?", vbQuestion + vbYesNo, App.ProductName) Then Exit Sub
    With gModels
        's = "delete RoomQtyByModel where model=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("Model"))) & vbCrLf
        s = "update RoomQtyByModel set inactive=1 where model=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("Model"))) & vbCrLf
        Call HFApp.SqlExec(s)
        Call .RemoveItem(.Row)
        Call .Select(GridNextVisibleRow(gModels, .Row), .Col)
        .SetFocus
        Call LoadModel(.TextMatrix(.Row, .ColIndex("Model")))
    End With
End Sub
Private Sub gCategories_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim r As Long
    With gCategories
        For r = Max(.Row, .RowSel) To Min(.Row, .RowSel) Step -1
            If .RowData(r) = "" Then .RowData(r) = "dirty"
            Dirty = True
        Next
    End With
End Sub

Private Sub gCategories_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    gCategories.ColSel = gCategories.Col
End Sub







Private Sub LoadModels()
    Dim s As String
    Dim rs As Recordset
    Dim r As Long
    
    With gModels
        .Rows = 1
        
        s = ""
        If mModel <> "" Then
            s = s & "select model,min(description) description" & vbCrLf
            s = s & "from tbldbassemblymaster " & vbCrLf
            s = s & "where assemblytype=0 and isbaseassembly=0 and inactive=0 and divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
            s = s & "and model=" & DbQuote(Str, mModel) & vbCrLf
            s = s & "group by model" & vbCrLf
        Else
            s = s & "select q.model,min(m.description) description" & vbCrLf
            s = s & "from RoomQtyByModel q" & vbCrLf
            s = s & "join roommaster r on q.roomid=r.roomid" & vbCrLf
            s = s & "join tbldbassemblymaster m on m.assemblytype=0 and m.model=q.model and r.divisionid=m.divisionid and m.isbaseassembly=0 and m.inactive=0" & vbCrLf
            s = s & "where isnull(r.inactive,0)=0 and isnull(q.inactive,0)=0 and r.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
            s = s & "group by q.model" & vbCrLf
            s = s & "order by 1,2" & vbCrLf
        End If
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
            .AddItem ""
            r = .Rows - 1
            .TextMatrix(r, .ColIndex("description")) = "" & rs("description")
            .TextMatrix(r, .ColIndex("model")) = "" & rs("model")
            rs.MoveNext
        Wend
        
        If .Rows > 1 Then .Row = 1
        
    End With
    
End Sub
Private Sub gModels_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
    Static inHere As Boolean
    If inHere Then Exit Sub
    inHere = True
    
    With gModels
        If .Rows = 1 Then GoTo ExitSub
        If OldRowSel > 0 Then
            Cancel = Not SaveData(True)
            If Cancel Then GoTo ExitSub
        End If
        Call .Select(NewRowSel, 0)
        Call LoadModel(.TextMatrix(NewRowSel, .ColIndex("Model")))
    End With

ExitSub:
    inHere = False
End Sub

Private Sub LoadModel(Model As String)
    Dim s As String
    Dim rs As Recordset
    Dim r As Long
    
    mModel = Model
    
    With gCategories
        .Rows = 1
        
        s = ""
        s = s & "select q.*,m.Room,c.Description CatDesc" & vbCrLf
        s = s & "from RoomQtyByModel q" & vbCrLf
        s = s & "join RoomMaster m on q.roomid=m.roomid" & vbCrLf
        s = s & "join tblcategories c on q.subcategory=c.category" & vbCrLf
        s = s & "where isnull(q.inactive,0)=0 and isnull(m.inactive,0)=0 and q.model=" & DbQuote(Str, mModel) & vbCrLf
        s = s & "order by m.room,q.subcategory,q.uom"
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
            .AddItem ""
            r = .Rows - 1
            .TextMatrix(r, .ColIndex("roomid")) = "" & rs("roomid")
            .TextMatrix(r, .ColIndex("room")) = "" & rs("room")
            .TextMatrix(r, .ColIndex("subcategory")) = "" & rs("subcategory")
            .TextMatrix(r, .ColIndex("catdesc")) = "" & rs("catdesc")
            .TextMatrix(r, .ColIndex("qty")) = "" & rs("qty")
            .TextMatrix(r, .ColIndex("uom")) = "" & rs("uom")
            rs.MoveNext
        Wend
    End With

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
    
    With gCategories
    For r = 1 To .Rows - 1
    If .RowData(r) = "dirty" Then
    
        s = ""
        s = s & "update RoomQtyByModel set qty=" & DbQuote(Num, .TextMatrix(r, .ColIndex("Qty"))) & vbCrLf
        s = s & "where model=" & DbQuote(Str, mModel) & vbCrLf
        s = s & " and roomid=" & DbQuote(Num, .TextMatrix(r, .ColIndex("roomid"))) & vbCrLf
        s = s & " and subcategory=" & DbQuote(Str, .TextMatrix(r, .ColIndex("SubCategory"))) & vbCrLf
        s = s & " and uom=" & DbQuote(Str, .TextMatrix(r, .ColIndex("UOM"))) & vbCrLf
        Call HFApp.SqlExec(s)
        .RowData(r) = ""
        
    End If
    Next
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



Private Sub gModels_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Exit Sub
    With gModels
        mnuModelsSub(mnuMODEL_ADDMODEL).Visible = mOrigModel = ""
        mnuModelsSub(mnuMODEL_DELETE).Visible = mOrigModel = ""
        
        mnuModelsSub(mnuMODEL_ADDROOM).Enabled = .Row > 0
        mnuModelsSub(mnuMODEL_DELETE).Enabled = .Row > 0
        PopupMenu mnuModels
    End With
End Sub

Private Sub gCategories_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Exit Sub
    
    If gCategories.Row <> gCategories.MouseRow Then gCategories.Row = gCategories.MouseRow
    
    mnuCategoriesSub(mnuCATEGORIES_ADDROOM).Enabled = mModel <> ""
    mnuCategoriesSub(mnuCATEGORIES_REMOVEROOM).Enabled = gCategories.Row > 0
    PopupMenu mnuCategories
End Sub


Private Sub mnuCategoriesSub_Click(Index As Integer)
    
    With gCategories
    Select Case Index
    Case mnuCATEGORIES_REMOVEROOM
        Call SaveData(False)
        Call DeleteRoom
        Call LoadModel(mModel)
    
    Case mnuCATEGORIES_ADDROOM
        Call mnuModelsSub_Click(mnuMODEL_ADDROOM)
    
    End Select
    End With

End Sub


Private Sub DeleteRoom()
    Dim s As String
    
    With gCategories
        's = "delete RoomQtyByModel where model=" & DbQuote(Str, mModel) & " and roomid=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("roomid")))
        s = "update RoomQtyByModel set inactive=1 where model=" & DbQuote(Str, mModel) & " and roomid=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("roomid")))
        Call HFApp.SqlExec(s)
    End With
    
End Sub

Private Sub mnuModelsSub_Click(Index As Integer)
    Dim s As String
    Dim ids As String
    Dim i As Long
    
    
    With gModels
    Select Case Index
    Case mnuMODEL_DELETE
        Call DeleteModel
    
    Case mnuMODEL_ADDMODEL
        s = ""
        s = s & "select model,min(description) description" & vbCrLf
        s = s & "from tbldbassemblymaster " & vbCrLf
        s = s & "where divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        s = s & "and assemblytype=0 and isbaseassembly=0 and inactive=0" & vbCrLf
        s = s & "group by model" & vbCrLf
        If Not FPickList.Choose(HFApp.Databases(dbHomefront), "Model", s) Then Exit Sub
        Call SaveData(False)
        Call .AddItem("")
        .TextMatrix(.Rows - 1, .ColIndex("model")) = FPickList.SelectedItem("model")
        .TextMatrix(.Rows - 1, .ColIndex("description")) = FPickList.SelectedItem("description")
        Call .Select(.Rows - 1, 0)
        Call mnuModelsSub_Click(mnuMODEL_ADDROOM)
        
    Case mnuMODEL_ADDROOM
        's = "select RoomID,Room from RoomMaster where divisionid=" & DbQuote(Num, HFApp.DivisionID)
        
        s = ""
        s = s & "select distinct r.RoomID,r.Room " & vbCrLf
        s = s & "from RoomMasterSubCategory c" & vbCrLf
        s = s & "join RoomMaster r on r.roomid=c.roomid" & vbCrLf
        s = s & "where isnull(r.inactive,0)=0 and divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        s = s & "order by 1 asc" & vbCrLf
        If Not FPickList.Choose(HFApp.Databases(dbHomefront), "Room", s, , , , , "RoomID", True) Then Exit Sub
        Call SaveData(False)
        
        ids = ""
        For i = 1 To FPickList.SelectedItems
            ids = ids & "," & FPickList.SelectedItem("RoomID", i)
        Next
        
        
        
        s = ""
        s = s & "Insert RoomQtyByModel(RoomID, Model, SubCategory, UOM, Qty)" & vbCrLf
        s = s & "select q.RoomID," & DbQuote(Str, .TextMatrix(.Row, 0)) & ",q.subcategory,q.uom,0 " & vbCrLf
        s = s & "from RoomMasterSubCategory q" & vbCrLf
        s = s & "left outer join roomQtyByModel m on m.model=" & DbQuote(Str, .TextMatrix(.Row, 0)) & " and m.roomid=q.roomid and m.subcategory=q.subcategory and m.uom=q.uom" & vbCrLf
        s = s & "where m.roomid is null" & vbCrLf
        s = s & "and q.roomID in(" & Mid(ids, 2) & ")" & vbCrLf
        
        s = s & "update m set inactive=0" & vbCrLf
        s = s & "from RoomMasterSubCategory q" & vbCrLf
        s = s & "join roomQtyByModel m on m.model=" & DbQuote(Str, .TextMatrix(.Row, 0)) & " and m.roomid=q.roomid and m.subcategory=q.subcategory and m.uom=q.uom" & vbCrLf
        s = s & "where q.roomID in(" & Mid(ids, 2) & ")" & vbCrLf
        
        Call HFApp.SqlExec(s)
        Call LoadModel(.TextMatrix(.Row, 0))
        
    End Select
    End With
    
End Sub

Private Sub Slider_Move()
    Call Form_Resize
End Sub
