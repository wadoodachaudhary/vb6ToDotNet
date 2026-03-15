VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FInquiry 
   Caption         =   "Inquiry"
   ClientHeight    =   2235
   ClientLeft      =   6135
   ClientTop       =   5010
   ClientWidth     =   3045
   Icon            =   "FInquiry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
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

Private mcaption As String
Private sql    As String
Private opCode As String
Private opFunc As Long
Private opCol  As Long
Private grpCol As Long
Private expand As Boolean
    



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Const margin = 0
    gData.Move margin, margin, Me.ScaleWidth - 2 * margin, Me.ScaleHeight - 2 * margin
End Sub


Public Sub ShowInquiry(Caption As String, query As String)
    
    Screen.MousePointer = vbHourglass
    mcaption = Caption
    
    'parse statement parts
    sql = query
    opCode = Trim(Parse(UCase(sql), 2, "GROUP ON"))
    grpCol = Val(Parse(opCode, 1, " ")) - 1
    opCol = Val(Parse(opCode, 3, " ")) - 1
    expand = "EXPAND" = Trim(Parse(opCode, 4, " "))
    opCode = Trim(Parse(opCode, 2, " "))
    opCode = Replace(opCode, vbNullChar, "")
    opCode = Replace(opCode, vbCr, "")
    opCode = Replace(opCode, vbLf, "")
    sql = Parse(sql, 1, "GROUP ON")
    Me.Show vbModal

End Sub
    
Private Sub Form_Load()
On Error GoTo eh
    
    Dim i      As Long
    Dim rs     As Recordset
    
    Me.Caption = mcaption
    Call IniGetForm(Me, , mcaption)
    
    With gData
        .Redraw = flexRDNone
        .Cell(flexcpFontName, 0, 0) = Me.FontName
        If LCase(left(sql, 6)) <> "select" Then
            .GridLines = flexGridNone
            .FocusRect = flexFocusNone
            .HighLight = flexHighlightNever
            .Rows = 1
            .cols = 1
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
            .cols = rs.fields.Count
            'set up outline
            Call .Subtotal(flexSTClear)
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
                    Call .Subtotal(opFunc, grpCol, opCol, , , , True, "%s")
                Else
                    Call .Subtotal(opFunc, grpCol, , , , , True, "%s")
                End If
                If Not expand Then Call .Outline(grpCol)
            End If
        End If
        
        For i = 0 To .Rows - 1
            If .IsSubtotal(i) Then
                .Cell(flexcpBackColor, i, 0, i, .cols - 1) = vbButtonFace
                .Cell(flexcpFontBold, i, 0, i, .cols - 1) = True
            End If
        Next
        
        
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .cols - 1)
        
        
        If opCode <> "" Then .ColWidth(0) = 210
        
        .ExtendLastCol = False
        .Redraw = flexRDBuffered
    End With
    Call SetCtrlFocus(gData)
    Screen.MousePointer = vbDefault
    Call WindowOnTop(Me, True)
Exit Sub
eh:
With gData
        .Rows = 1
        .cols = 1
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
    Call IniPutForm(Me, , mcaption)
End Sub

Private Sub gData_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case True
        Case KeyCode = vbKeyF And Shift = vbCtrlMask
            Call FFind.Showform(gData)
    End Select
End Sub
