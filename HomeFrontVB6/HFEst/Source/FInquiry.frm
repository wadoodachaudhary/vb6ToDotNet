VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form FInquiry 
   Caption         =   "Inquiry"
   ClientHeight    =   2235
   ClientLeft      =   6135
   ClientTop       =   5010
   ClientWidth     =   3045
   Icon            =   "FInquiry.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2235
   ScaleWidth      =   3045
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   1815
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2475
      _cx             =   4366
      _cy             =   3201
      Appearance      =   2
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
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483633
      TreeColor       =   -2147483632
      FloodColor      =   -2147483635
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FInquiry.frx":000C
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
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
End
Attribute VB_Name = "FInquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private GroupedColumns As Long 'number of grouped columns
Private mFunnyFlag As Boolean ' used to ignore the first mouse move event see gitems_mousemove and gitems_rowcolchanged-
Private mFilename As String


Private Sub Form_Resize()
On Error Resume Next
    Const margin = 0
    gData.Move margin, margin, Me.ScaleWidth - 2 * margin, Me.ScaleHeight - 2 * margin
End Sub


Public Sub ShowInquiry(sFileName As String)
On Error GoTo eh
    
    Dim i      As Long
    Dim rs     As Recordset
    
    Static sql As String
    Static opCode As String
    Static opFunc As Long
    Static opCol  As Long
    Static grpCol As Long
    Static expand As Boolean
    mFilename = sFileName
    Screen.MousePointer = vbHourglass
    On Error Resume Next
    Me.Show
    Call IniGetForm(Me, FileName(mFilename) & ".ini")
    Me.Caption = Replace(FileName(sFileName), "&", "")
    On Error GoTo eh

    'read query from file
    i = FreeFile
    Open sFileName For Input As #i
    sql = input(LOF(i), #i)
    Close #i
    
    'parse statement parts
    opCode = Trim(Parse(UCase(sql), 2, "GROUP ON"))
    grpCol = Val(Parse(opCode, 1, " ")) - 1
    opCol = Val(Parse(opCode, 3, " ")) - 1
    expand = "EXPAND" = Trim(Parse(opCode, 4, " "))
    opCode = Trim(Parse(opCode, 2, " "))
    opCode = Replace(opCode, vbNullChar, "")
    opCode = Replace(opCode, vbCr, "")
    opCode = Replace(opCode, vbLf, "")
    sql = Parse(sql, 1, "GROUP ON")
    sql = Replace(sql, "{DivisionID}", HFApp.DivisionID)
       
    With gData
        .Redraw = flexRDNone
        .Cell(flexcpFontName, 0, 0) = Me.FontName
        .OwnerDraw = flexODContent
        If LCase(Left(sql, 6)) <> "select" Then
            .GridLines = flexGridNone
            .FocusRect = flexFocusNone
            .HighLight = flexHighlightNever
            .Rows = 1
            .Cols = 1
            .FixedRows = 0
            Call HFApp.SqlExec(sql, , i)
            .Cell(flexcpAlignment, 0, 0) = flexAlignLeftTop
            .TextMatrix(0, 0) = vbCrLf & vbCrLf & String(8, " ") & "mCurrentCaption" & " has executed successfully." & vbCrLf & String(8, " ") & i & " row(s) were affected."
            .RowHeight(0) = .Height
        Else
            'load it into the grid
            .GridLines = flexGridFlat
            .FocusRect = flexFocusLight
            .HighLight = flexHighlightWithFocus
            .Rows = 1
            .FixedRows = 1
            .DataMode = flexDMBoundNoRowCount
            .RowHeight(0) = 240
            Set rs = HFApp.SqlExec(sql)
            Set .DataSource = rs
            
            'shouldnt have to do this
            .Cols = rs.fields.Count
            'set up outline
            Call .SubTotal(flexSTClear)
            Select Case opCode
                Case "NONE":    opFunc = SubtotalSettings.flexSTNone
                Case "SUM":     opFunc = SubtotalSettings.flexSTSum
                Case "PERCENT": opFunc = SubtotalSettings.flexSTPercent
                Case "COUNT":   opFunc = SubtotalSettings.flexSTCount
                Case "AVG":     opFunc = SubtotalSettings.flexSTAverage
                Case "MAX":     opFunc = SubtotalSettings.flexSTMax
                Case "MIN":     opFunc = SubtotalSettings.flexSTMin
                Case "STD":     opFunc = SubtotalSettings.flexSTStd
                Case "VAR":     opFunc = SubtotalSettings.flexSTVar
                Case Else:      opFunc = -1
            End Select
            
            If opFunc >= 0 Then
                .MergeCells = flexMergeOutline
                If opCol > -1 Then
                    Call .SubTotal(opFunc, grpCol, opCol, , , , True, "%s")
                Else
                    Call .SubTotal(opFunc, grpCol, , , , , True, "%s")
                End If
                If Not expand Then Call .Outline(grpCol)
            End If
        End If
        
        For i = 0 To .Rows - 1
            If .IsSubtotal(i) Then
                .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = vbButtonFace
                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = True
            End If
        Next
        
        Call GroupGrid
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1)
'        Call IniGetGrid(Me, gData, , , FileName(mFilename) & ".ini")
        
        .ExtendLastCol = False
        .Redraw = flexRDBuffered
    End With
    Call SetCtrlFocus(gData)
    Screen.MousePointer = vbDefault

    Exit Sub
eh:

With gData

        .Rows = 1
        .Cols = 1
        .FixedRows = 0
        .Cell(flexcpAlignment, 0, 0) = flexAlignLeftTop
        .Cell(flexcpFontName, 0, 0) = "Courier"
        .TextMatrix(0, 0) = vbCrLf & "Unable to process inquiry." & _
                            vbCrLf & Err.Description & _
                            vbCrLf & vbCrLf & sql
        .RowHeight(0) = .Height
        .Redraw = flexRDBuffered
    End With
    Screen.MousePointer = vbDefault
End Sub



Private Sub Form_Unload(Cancel As Integer)
Call IniPutForm(Me, FileName(mFilename) & ".ini")
'Call IniPutGrid(Me, gData, , FileName(mFilename) & ".ini")
End Sub

Private Sub gData_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, done As Boolean)
    With gData
        If .ColData(Col) = "GROUPED" And Row > .FixedRows And Not .IsSubtotal(Row) Then
            done = True
        End If
    End With
End Sub

Private Sub gData_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case True
        Case KeyCode = vbKeyF And Shift = vbCtrlMask
            Call FFind.ShowForm(gData)
    End Select
End Sub


Public Sub GroupGrid()
On Error GoTo eh
    
    Dim i As Long
    Dim Row As Long
    Dim Col As Long
    
    
    
    With Me.gData
        Row = .Row
        Col = .Col
        
        GroupedColumns = 0
        For i = 0 To .Cols - 1
            If .ColData(i) = "GROUPED" Then
                .ColPosition(i) = GroupedColumns
                GroupedColumns = GroupedColumns + 1
            End If
        Next
        
        If GroupedColumns = 0 Then
            'Call .Subtotal(flexSTClear)
        Else
            '.Subtotal flexSTClear
            mFunnyFlag = True
            .Col = 0
            .ColSel = GroupedColumns - 1
            .Sort = flexSortGenericAscending
            mFunnyFlag = False
            
            
            .OutlineCol = 0
            '.SubtotalPosition = flexSTAbove
            For i = 0 To GroupedColumns - 1
                '.Subtotal flexSTSum, i, .ColIndex("ExtendedAmount"), "$(#,###.00)", , vbHighlight, True, "%s", 0
            Next
        End If
    
        On Error Resume Next
        .Row = Row
        .Col = Col
    End With
Exit Sub
eh: Call errHandler("FInquiry " & "GroupGrid")
End Sub

Private Sub gData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then Call FMain.ShowColumnMenu(gData, False, False, False, False, False)

End Sub
