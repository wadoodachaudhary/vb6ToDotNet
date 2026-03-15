VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form FDefaultVendors 
   Caption         =   "Default Vendors"
   ClientHeight    =   3765
   ClientLeft      =   9660
   ClientTop       =   3585
   ClientWidth     =   6165
   FillColor       =   &H00808000&
   Icon            =   "FDefaultVendors.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3765
   ScaleWidth      =   6165
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   2355
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      _cx             =   6800
      _cy             =   4154
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
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483633
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FDefaultVendors.frx":000C
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
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   1
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin VB.Image Image2 
         Height          =   240
         Left            =   2400
         Picture         =   "FDefaultVendors.frx":006D
         Top             =   1410
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   2850
         Picture         =   "FDefaultVendors.frx":05F7
         Top             =   1440
         Visible         =   0   'False
         Width           =   240
      End
   End
End
Attribute VB_Name = "FDefaultVendors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FDefaultVendors::"
Const DIRTYCOLOR = vbRed
Private mDirty   As Boolean
Private Const mcEDIT_COPY = 0
Private Const mcEDIT_PASTE = 1

Private Sub Form_Load()
    IniGetForm Me
    Call LoadData
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not SaveData(True) Then Cancel = True
End Sub

Private Sub Form_Resize()
On Error Resume Next
    gData.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub
Private Sub LoadData()
    Dim rs As Recordset
    Dim r As Long
    Dim c As Long
    Dim s As String
    
    
    'this screen is very sensitive to orphan records so delete them first.
    On Error Resume Next
    Call HFApp.SqlExec("delete from poareavendor where DivisionID =" & HFApp.DivisionID & " and poindex in(select a.poindex from poareavendor a left outer join tblpoindex b on a.poindex=b.poindex and a.DivisionID = b.DivisionID where b.poindex is null and a.DivisionID =" & HFApp.DivisionID & ")", dbHomefront)
    Call HFApp.SqlExec("delete from poareavendor where DivisionID =" & HFApp.DivisionID & " and isnull(area,'')<>'' and area in(select a.area from poareavendor a left outer join tbllocality b on a.area=b.area where b.area is null and a.DivisionID =" & HFApp.DivisionID & ")", dbHomefront)
    'Call HFApp.SqlExec("delete from poareavendor where vendor in(select a.vendor from poareavendor a left outer join tblvendors b on (a.vendor=b.vendor_id and b.DivisionID = " & HFApp.DivisionID & ") where b.vendor_id is null)", dbHomeFront)
    Call HFApp.SqlExec("delete from poareavendor where DivisionID =" & HFApp.DivisionID & " and vendor in(select a.vendor from poareavendor a left outer join tblvendors b on (a.vendor=b.vendor_id and a.DivisionID = b.DivisionID) where b.vendor_id is null and a.DivisionID =" & HFApp.DivisionID & ")", dbHomefront)
    On Error GoTo 0


    
    
    
    With gData
        .Redraw = flexRDNone
        
        'load po indexes
        .Rows = 1
        r = 1
        Set rs = HFApp.SqlExec("SELECT POIndex,Description FROM tblPOIndex where divisionid = " & HFApp.DivisionID & " ORDER BY POIndex")
        While Not rs.EOF
            .AddItem "" & rs(0) & " " & IIf("" & rs(0) = "" & rs(1), "", "" & rs(1))
            .RowData(.Rows - 1) = "K" & rs(0)
            
            Set .Cell(flexcpPicture, r, 0) = Image1.Picture
            .ColKey(0) = "POIndex"
            rs.MoveNext
            r = r + 1
        Wend
        
        'load communities
        .Cols = 2
        .FrozenCols = 1
        s = ""
        If HFApp.DivisionID <> "" Then
            s = ""
            s = s & "SELECT Area,Description FROM tblLocality join DivisionCommunities d on d.Community=Area" & vbCrLf
            s = s & "WHERE Area<>'N/A' AND Inactive=0 AND ISNULL(Area,'')<>'' and d.DivisionID = " & HFApp.DivisionID & vbCrLf
            s = s & "ORDER BY 2" & vbCrLf
        Else
            s = "SELECT Area,Description FROM tblLocality WHERE Area<>'N/A' AND Inactive=0 AND ISNULL(Area,'')<>'' ORDER BY 2"
        End If
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
            c = .Cols
            .Cols = c + 1
            .ColAlignment(c) = flexAlignLeftCenter
            .ColComboList(c) = "..."
            .ColKey(c) = "" & rs(0)
            .TextMatrix(0, c) = Trim("" & rs(1))
            rs.MoveNext
        Wend
        
        'now load vendors
        s = ""
        s = s & "SELECT l.Area,POIndex,Vendor,Vendor_Name" & vbCrLf
        s = s & "FROM POAreaVendor JOIN tblVendors ON(Vendor=Vendor_ID and tblvendors.divisionid = POAreaVendor.DivisionID) LEFT OUTER JOIN tblLocality l ON(POAreaVendor.Area=l.Area)" & vbCrLf
        s = s & "WHERE POAreaVendor.DivisionID =" & HFApp.DivisionID & " and (l.Inactive=0 or l.Area IS NULL)" & vbCrLf
        s = s & "ORDER BY POIndex" & vbCrLf
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
            r = FindRow("" & rs("POIndex"))
            c = .ColIndex("" & rs("Area"))
            If c < 1 Or r < 1 Then
'                MsgBox "unknown area or poindex " & rs("Area") & " : " & rs("POIndex")
            Else
                .Cell(flexcpData, r, c) = Trim("" & rs("Vendor"))
                .Cell(flexcpText, r, c) = Trim("" & rs("Vendor_Name"))
            End If
            rs.MoveNext
        Wend

        Call .AutoSize(0, .Cols - 1)
        .Cell(flexcpForeColor, 0, 0, .Rows - 1, .Cols - 1) = vbWindowText
        .Redraw = flexRDBuffered
    End With
    
    Call IniGetGrid(Me, gData, , , , True, True)
    
End Sub
Public Function SaveData(prompt As Boolean) As Boolean
On Error GoTo eh:
'    Dim wms As New ADODB.Connection
    Dim rc As Long
    If Not mDirty Then
        SaveData = True
        Exit Function
    End If
    rc = vbYes
    If prompt Then rc = MsgBox("This data has changed." & vbCrLf & vbCrLf & "Do you want to save the changes?", vbExclamation + vbYesNoCancel, Me.Caption)
    Select Case rc
        Case vbNo:     SaveData = True:     Exit Function
        Case vbCancel: Exit Function
    End Select
    
    
    Dim s As String
    Dim r As Long
    Dim c As Long
    Dim rs As Recordset
    Dim Area    As String
    Dim vendor  As String
    Dim POIndex As String
    Screen.MousePointer = vbHourglass
    
    With gData
        For c = 1 To .Cols - 1
            Area = .ColKey(c)
            For r = 1 To .Rows - 1
                If .Cell(flexcpForeColor, r, c) = DIRTYCOLOR Then
                    
                    s = ""
                    s = s & "DELETE FROM POAreaVendor"
                    s = s & " WHERE DivisionID = " & HFApp.DivisionID & " and Area=" & DbQuote(Str, Area)
                    s = s & "   AND POIndex=" & DbQuote(Str, Mid(.RowData(r), 2))
                    Call HFApp.SqlExec(s)
                    
                    vendor = Trim(.Cell(flexcpData, r, c))
                    If vendor <> "" Then
                        
                        s = ""
                        s = s & "SELECT Vendor_ID" & vbCrLf
                        s = s & "  FROM tblVendors" & vbCrLf
                        s = s & " WHERE Vendor_ID=" & DbQuote(Str, vendor) & " and DivisionID = " & HFApp.DivisionID & vbCrLf
                        Set rs = HFApp.SqlExec(s)
                        If rs.EOF Then
                            .Row = r
                            .Col = c
                            Screen.MousePointer = vbDefault
                            MsgBox "unknown vendor", vbExclamation, Me.Caption
                            Exit Function
                        End If
                        vendor = "" & rs("Vendor_ID")
                        POIndex = Mid(.RowData(r), 2)
                    
                        s = ""
                        s = s & "INSERT INTO POAreaVendor(DivisionID,Area,POIndex,Vendor)"
                        s = s & " VALUES(" & HFApp.DivisionID & "," & DbQuote(Str, Area)
                        s = s & "," & DbQuote(Str, POIndex)
                        s = s & "," & DbQuote(Str, vendor) & ")"
                        Call HFApp.SqlExec(s)
                        
                    End If
                    
                End If
            Next
        Next
        .Cell(flexcpForeColor, 0, 0, .Rows - 1, .Cols - 1) = vbWindowText
    End With
    Screen.MousePointer = vbDefault
    mDirty = False
    SaveData = True
    
    Exit Function
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Screen.MousePointer = False
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Function




Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gData)
End Sub

Private Sub gData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'    if col=0
    Cancel = Col = 0
    
    
End Sub



Private Sub gData_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    If gData.ColKey(Col) = "POIndex" Then
        Position = 0
    Else
        Position = Max(Position, 1)
    End If
End Sub

Private Sub gData_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal newrow As Long, ByVal NewCol As Long, Cancel As Boolean)
    Cancel = NewCol = 0
End Sub


Private Sub gData_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim s As String
    Dim r1 As Long
    Dim r2 As Long
    Dim c1 As Long
    Dim c2 As Long
    With gData
        s = .Text
        If FPickList.Choose(HFApp.Databases(dbHomefront), "Vendor", "SELECT vendorgroupid Trade,Vendor_Name Company ,Vendor_ID Vendor ,City,Phone FROM tblVendors where divisionid = " & HFApp.DivisionID & " and Inactive = 0", s) Then
            Call .GetSelection(r1, c1, r2, c2)
            .Cell(flexcpText, r1, c1, r2, c2) = FPickList.SelectedItem("Company")
            .Cell(flexcpData, r1, c1, r2, c2) = FPickList.SelectedItem("Vendor")
            .Cell(flexcpForeColor, r1, c1, r2, c2) = DIRTYCOLOR
            mDirty = True
        End If
    End With
End Sub

Private Sub gData_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'source clip
    Static ClipRow1 As Long
    Static ClipRow2 As Long
    Static ClipCol1 As Long
    Static ClipCol2 As Long
    Static ClipCols As Long
    Static ClipRows As Long

    'destination clip
    Dim DestRow1 As Long
    Dim DestRow2 As Long
    Dim DestCol1 As Long
    Dim DestCol2 As Long
    
    'used in tiling across region
    Dim SourceRow As Long
    Dim SourceCol As Long
    
    'loop counters
    Dim r As Long
    Dim c As Long
    
    With gData
    Select Case True
        
        Case KeyCode = vbKeyS And Shift = vbCtrlMask
            Call SaveData(False)
        
        Case KeyCode = vbKeyF And Shift = vbCtrlMask
            Call FFind.ShowForm(gData)
        
        Case KeyCode = vbKeyF4 And Shift = 0
            If .Row > 0 And .Col > 0 Then Call gData_CellButtonClick(.Row, .Col)
            
        Case KeyCode = vbKeyDelete
            .Cell(flexcpData, .Row, .Col, .RowSel, .ColSel) = ""
            .Cell(flexcpText, .Row, .Col, .RowSel, .ColSel) = ""
            .Cell(flexcpForeColor, .Row, .Col, .RowSel, .ColSel) = DIRTYCOLOR
            mDirty = True
            
        Case Shift = vbCtrlMask And KeyCode = vbKeyC
            Call .GetSelection(ClipRow1, ClipCol1, ClipRow2, ClipCol2)
            ClipRows = ClipRow2 - ClipRow1 + 1
            ClipCols = ClipCol2 - ClipCol1 + 1
            
            
        Case Shift = vbCtrlMask And KeyCode = vbKeyV
            If .Row < 1 Or .Col < 1 Then Exit Sub
            If .Col < 1 Or .ColSel < 1 Then Exit Sub
            If ClipRow1 = 0 Then Exit Sub
            
            Call .GetSelection(DestRow1, DestCol1, DestRow2, DestCol2)
            If DestRow1 = DestRow2 And DestCol1 = DestCol2 Then
                'user has a single destination cell selected so paste entire clip
                For c = 1 To ClipCols
                    For r = 1 To ClipRows
                        If .Row + r > .Rows Or .Col + c > .Cols Then
                            'cell is not on sheet
                        Else
                            .Cell(flexcpText, DestRow1 + r - 1, DestCol1 + c - 1) = .Cell(flexcpText, ClipRow1 + r - 1, ClipCol1 + c - 1)
                            .Cell(flexcpData, DestRow1 + r - 1, DestCol1 + c - 1) = .Cell(flexcpData, ClipRow1 + r - 1, ClipCol1 + c - 1)
                            .Cell(flexcpForeColor, DestRow1 + r - 1, DestCol1 + c - 1) = DIRTYCOLOR
                        End If
                    Next
                Next
            Else
                'user has a range of cells selected so "tile" the clip across the region
                For c = DestCol1 To DestCol2
                    For r = DestRow1 To DestRow2
                        
                        SourceCol = (c - DestCol1 + 1) Mod ClipCols: If SourceCol = 0 Then SourceCol = ClipCols
                        SourceRow = (r - DestRow1 + 1) Mod ClipRows: If SourceRow = 0 Then SourceRow = ClipRows
                            
                        .Cell(flexcpText, r, c) = .Cell(flexcpText, ClipRow1 + SourceRow - 1, ClipCol1 + SourceCol - 1)
                        .Cell(flexcpData, r, c) = .Cell(flexcpData, ClipRow1 + SourceRow - 1, ClipCol1 + SourceCol - 1)
                        .Cell(flexcpForeColor, r, c) = DIRTYCOLOR
                        
                    Next
                Next
            End If
            
            
            
            mDirty = True
            
    End Select
    End With
End Sub


Private Function FindRow(POIndex As String) As Long
Static r As Long
    With gData
        If r < .Rows Then
            If Mid(.RowData(r), 2) = POIndex Then
                FindRow = r
                Exit Function
            End If
        End If
        For r = 1 To .Rows - 1
            If Mid(.RowData(r), 2) = POIndex Then
                FindRow = r
                Exit Function
            End If
        Next
    End With
    FindRow = -1
End Function

