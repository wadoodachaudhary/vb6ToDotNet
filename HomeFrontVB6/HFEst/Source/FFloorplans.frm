VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form FFloorplans 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Floorplan Dimensions"
   ClientHeight    =   5385
   ClientLeft      =   7875
   ClientTop       =   2745
   ClientWidth     =   9705
   Icon            =   "FFloorplans.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   9705
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   7275
      TabIndex        =   2
      Top             =   4815
      Width           =   1095
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   8445
      TabIndex        =   1
      Top             =   4815
      Width           =   1095
   End
   Begin VSFlex8Ctl.VSFlexGrid gModels 
      Height          =   4365
      Left            =   300
      TabIndex        =   3
      Top             =   345
      Width           =   2985
      _cx             =   5265
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
      FormatString    =   $"FFloorplans.frx":000C
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
      Left            =   3540
      TabIndex        =   4
      Top             =   330
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
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FFloorplans.frx":005E
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
   Begin VB.Label lblContacts 
      AutoSize        =   -1  'True
      Caption         =   "Models"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupSub 
         Caption         =   "Add Model..."
         Index           =   0
      End
      Begin VB.Menu mnuPopupSub 
         Caption         =   "Add Room..."
         Index           =   1
      End
   End
End
Attribute VB_Name = "FFloorplans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FModelDimensions::"

Private mModelID As String
Private mDirty As Boolean


Private Sub cmdNav_Click(Index As Integer)
    If SaveData(Index = 1) Then
        Dirty = False
        Unload Me
    End If
End Sub

Public Sub ShowForm()
    Me.Show vbModal
End Sub


Private Property Let Dirty(RHS As Boolean)
    mDirty = RHS
    cmdNav(0).Enabled = mDirty
End Property

Private Sub Form_Load()
    Call IniGetForm(Me)
    Call IniGetGrid(Me, gCategories)
    Call IniGetGrid(Me, gModels)
    Call LoadModels
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


Private Sub gCategories_AfterSort(ByVal Col As Long, Order As Integer)
    Call gCategories.AddItem("")
End Sub

Private Sub gCategories_BeforeSort(ByVal Col As Long, Order As Integer)
    Call gCategories.RemoveItem(gCategories.Rows - 1)
End Sub




Private Sub gModels_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim s As String
    With gModels
        If .Row = .Rows - 1 Or .Row < 1 Then Exit Sub
        If KeyCode = vbKeyDelete And Shift <> 0 Then
            
            If vbNo = MsgBox("Are you sure you want to delete the dimensions for this model?", vbQuestion + vbYesNo, App.ProductName) Then Exit Sub
        
            s = ""
            s = s & "delete RoomQtyByModel where RoomID=" & DbQuote(Num, .TextMatrix(.Row, .ColIndex("ModelID"))) & vbCrLf
            Call HFApp.SqlExec(s)
            Call .RemoveItem(.Row)
            Call .Select(GridNextVisibleRow(gModels, .Row), .Col)
            .SetFocus
            Call LoadModel(.TextMatrix(.Row, .ColIndex("ModelID")))
        End If
    End With
End Sub

Private Sub gCategories_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim r As Long
    With gCategories
        For r = Max(.Row, .RowSel) To Min(.Row, .RowSel) Step -1
            
            If r = .Rows - 1 Then
                .AddItem ""
                .RowData(r) = "new"
            End If
            
            If .RowData(r) = "" Then .RowData(r) = "dirty"
            Dirty = True
        Next
    End With
End Sub

Private Sub gCategories_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    gCategories.ColSel = gCategories.Col
End Sub




Private Sub gCategories_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim s As String
    With gCategories
        If .Row < 1 Then Exit Sub
        If .Row = .Rows - 1 Then Exit Sub
        If KeyCode = vbKeyDelete And Shift <> 0 Then
            .RowData(.Row) = "delete"
            .RowHidden(.Row) = True
            Dirty = True
            Call .Select(GridNextVisibleRow(gCategories, .Row), .Col)
            .SetFocus
        End If
    End With
End Sub



Private Sub LoadModels()
    Dim s As String
    Dim rs As Recordset
    Dim r As Long
    
    With gModels
        .Rows = 1
        
        s = ""
        s = s & "select distinct q.model " & vbCrLf
        s = s & "from RoomQtyByModel q" & vbCrLf
        s = s & "join roommaster r on q.roomid=r.roomid" & vbCrLf
        s = s & "where r.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        s = s & "order by q.model" & vbCrLf
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
            .AddItem ""
            r = .Rows - 1
            .TextMatrix(r, .ColIndex("ModelID")) = "" & rs("Model")
            .TextMatrix(r, .ColIndex("model")) = "" & rs("model")
            rs.MoveNext
        Wend
    End With
    
End Sub
Private Sub gModels_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
With gModels
    If .Rows = 1 Then Exit Sub
    If OldRowSel > 0 Then
        Cancel = Not SaveData(True)
        If Cancel Then Exit Sub
    End If
    Call LoadModel(.TextMatrix(NewRowSel, .ColIndex("ModelID")))
End With
End Sub

Private Sub LoadModel(ModelID As String)
    Dim s As String
    Dim rs As Recordset
    Dim r As Long
    
    mModelID = ModelID
    
    With gCategories
        .Rows = 1
        
        s = ""
        s = s & "select q.*,m.Room" & vbCrLf
        s = s & "from RoomQtyByModel q" & vbCrLf
        s = s & "join RoomMaster m on q.roomid=m.roomid" & vbCrLf
        s = s & "where q.model=" & DbQuote(Str, ModelID) & vbCrLf
        s = s & "order by m.room,q.subcategory,q.uom"
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
            .AddItem ""
            r = .Rows - 1
            .TextMatrix(r, .ColIndex("room")) = "" & rs("room")
            .TextMatrix(r, .ColIndex("subcategory")) = "" & rs("subcategory")
            .TextMatrix(r, .ColIndex("qty")) = "" & rs("qty")
            .TextMatrix(r, .ColIndex("uom")) = "" & rs("uom")
            rs.MoveNext
        Wend
        .AddItem ""
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
    
    s = "delete roommastersubcategory where roomid=" & DbQuote(Num, mModelID)
    Call HFApp.SqlExec(s)
    
    With gCategories
    For r = 1 To .Rows - 2
    If .RowData(r) <> "delete" Then
    
        s = ""
        s = s & "insert roommastersubcategory(ModelID, SubCategory, UOM) values(" & vbCrLf
        s = s & DbQuote(Num, mModelID)
        s = s & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("SubCategory")))
        s = s & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("UOM")))
        s = s & ")"
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
        mnuPopupSub(1).Enabled = .Row > 0
        PopupMenu mnuPopup
    End With
End Sub
