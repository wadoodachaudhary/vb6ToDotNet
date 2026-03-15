VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{302C5C1A-C2E2-4302-9AF9-BAC87EEFDECE}#1.0#0"; "Panels.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FApprovals 
   Caption         =   "Approvals"
   ClientHeight    =   6120
   ClientLeft      =   4695
   ClientTop       =   1425
   ClientWidth     =   8940
   Icon            =   "FApprovals.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6120
   ScaleWidth      =   8940
   Begin Panels.Slider Slider 
      Height          =   30
      Left            =   1380
      Top             =   2490
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   53
      Orientation     =   1
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   1005
      ButtonWidth     =   1402
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Preview"
            Key             =   "Preview"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Send POs"
            Key             =   "SendPOs"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VSFlex8Ctl.VSFlexGrid gHeaders 
      Height          =   1785
      Left            =   90
      TabIndex        =   1
      Top             =   630
      Width           =   8535
      _cx             =   1999125199
      _cy             =   1999113293
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
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   4
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FApprovals.frx":000C
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
      TabBehavior     =   1
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
   Begin VSFlex8Ctl.VSFlexGrid gDetails 
      Height          =   3135
      Left            =   90
      TabIndex        =   2
      Top             =   2730
      Width           =   8535
      _cx             =   1999125199
      _cy             =   1999115674
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
      Rows            =   5
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FApprovals.frx":0185
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
End
Attribute VB_Name = "FApprovals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SRCFILE = "FApprovals"

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        Call LoadData
    End If
End Sub

Private Sub Form_Load()
    Call SetToolbarIcons(Toolbar, FMain.LargeIcons)
    Call IniGetForm(Me)
    Call ClearGroups
    Call IniGetGrid(Me, gHeaders, , , , True)
    Call IniGetGrid(Me, gDetails)
    
    Toolbar.Buttons("SendPOs").Enabled = HFApp.UserPermission("SendPO")
    
    If HFApp.Options(AccountingSystem) = asQuickBooks Then
        gHeaders.TextMatrix(0, gHeaders.ColIndex("Vendor")) = ""
        gHeaders.ColHidden(gHeaders.ColIndex("Vendor")) = True
    Else
        If gHeaders.TextMatrix(0, gHeaders.ColIndex("Vendor")) = "" Then gHeaders.TextMatrix(0, gHeaders.ColIndex("Vendor")) = "Vendor"
    End If
    
    Call LoadData
End Sub

Private Sub Form_Resize()
On Error Resume Next
    
    Slider.Min = 1500
    Slider.Max = Me.ScaleHeight - 700
    Slider.Move 0, Slider.Top, Me.ScaleWidth
    gHeaders.Move 0, Toolbar.Height - Screen.TwipsPerPixelY, Me.ScaleWidth, Slider.Top - Toolbar.Height - Screen.TwipsPerPixelY
    gDetails.Move 0, Slider.Top + Slider.Height, Me.ScaleWidth, Me.ScaleHeight - Slider.Top - Slider.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gHeaders)
    Call IniPutGrid(Me, gDetails)
End Sub

Private Sub LoadData()
    Dim s As String
    Dim rs As Recordset

    s = ""
    s = s & "select m.Job" & vbCrLf
    s = s & "      ,m.JobDesc" & vbCrLf
    s = s & "      ,m.Vendor" & vbCrLf
    s = s & "      ,m.VendorDesc" & vbCrLf
    s = s & "      ,m.PONumber" & vbCrLf
    s = s & "      ,m.POIndex" & vbCrLf
    s = s & "      ,m.PODesc" & vbCrLf
    s = s & "      ,m.POFormat" & vbCrLf
    s = s & "      ,sum(i.pretax) Pretax" & vbCrLf
    s = s & "      ,sum(i.jctax + i.njctax) Tax" & vbCrLf
    s = s & "      ,sum(i.pretax + i.jctax + i.njctax) Total" & vbCrLf
    s = s & "  from purchaseorders m" & vbCrLf
    s = s & "       join poitems i on(m.ponumber=i.ponumber and m.DivisionID = i.DivisionID)" & vbCrLf
    s = s & " where isfieldpo=1" & vbCrLf
    s = s & "   and deliverydate is null" & vbCrLf
    s = s & "   and m.DivisionID = " & HFApp.DivisionID
    s = s & "group by m.Job" & vbCrLf
    s = s & "      ,m.JobDesc" & vbCrLf
    s = s & "      ,m.Vendor" & vbCrLf
    s = s & "      ,m.VendorDesc" & vbCrLf
    s = s & "      ,m.PONumber" & vbCrLf
    s = s & "      ,m.POIndex" & vbCrLf
    s = s & "      ,m.PODesc,m.POFormat" & vbCrLf
    s = s & "order by m.jobdesc,m.vendordesc,m.ponumber" & vbCrLf
    Set rs = HFApp.SqlExec(s, dbHomefront)

    With gHeaders
    .Rows = 1
    While Not rs.EOF

        .AddItem ""
        .TextMatrix(.Rows - 1, .ColIndex("Job")) = "" & rs("Job")
        .TextMatrix(.Rows - 1, .ColIndex("JobDesc")) = "" & rs("JobDesc")
        .TextMatrix(.Rows - 1, .ColIndex("Vendor")) = "" & rs("Vendor")
        .TextMatrix(.Rows - 1, .ColIndex("VendorDesc")) = "" & rs("VendorDesc")
        .TextMatrix(.Rows - 1, .ColIndex("PONumber")) = "" & rs("PONumber")
        .TextMatrix(.Rows - 1, .ColIndex("POIndex")) = "" & rs("POIndex")
        .TextMatrix(.Rows - 1, .ColIndex("PODesc")) = "" & rs("PODesc")
        .TextMatrix(.Rows - 1, .ColIndex("Pretax")) = "" & rs("Pretax")
        .TextMatrix(.Rows - 1, .ColIndex("Tax")) = "" & rs("Tax")
        .TextMatrix(.Rows - 1, .ColIndex("Total")) = "" & rs("Total")

        Call rs.MoveNext
    Wend
    End With
    Call GroupGrid

End Sub

Private Sub gHeaders_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, done As Boolean)
    With gHeaders
        If .ColData(Col) = "GROUPED" And Row > 0 And (Col < .RowOutlineLevel(Row) Or Not .IsSubtotal(Row)) Then
            done = True
        End If
        
    End With
End Sub

Private Sub gHeaders_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton And gHeaders.MouseRow = 0 Then
        Call FMain.ShowColumnMenu(gHeaders, , , , True)
    End If
End Sub

Private Sub Slider_Move()
    Call Form_Resize
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "SendPOs":     Call FSendingWizard.ShowForm("PO")
        Case "Preview":     Call ShowSelectedPOs
    End Select
End Sub


Private Sub ShowSelectedPOs()
On Error GoTo eh
    Dim s As String
    Dim r As Long
    Dim rs As Recordset
    Dim pos As String
    Dim lastRpt As String
    'Dim f As FRptViewer
    Dim c As ZybUtil.Crystal
    
    With gHeaders
        s = ""
        s = s & "select poformat,ponumber" & vbCrLf
        s = s & "  from purchaseorders" & vbCrLf
        s = s & " where 1=2" & vbCrLf
        For r = Min(.Row, .RowSel) To Max(.Row, .RowSel)
            s = s & "    or (" & Mid(IIf(.TextMatrix(r, .ColIndex("Job")) = "", "", " and job=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Job")))) & _
                                     IIf(.TextMatrix(r, .ColIndex("Vendor")) = "", "", " and vendor=" & DbQuote(Str, .TextMatrix(r, .ColIndex("vendor")))) & _
                                     IIf(.TextMatrix(r, .ColIndex("ponumber")) = "", "", " and ponumber=" & DbQuote(Str, .TextMatrix(r, .ColIndex("ponumber")))), 6) & ")" & vbCrLf
        Next
        s = s & "order by 1,2"
    End With
    
    Set rs = HFApp.SqlExec(s)
    'loop thru this set building csv list of ponumbers in each format then showing rpt
    While Not rs.EOF
        If lastRpt <> "" & rs("poformat") Then
            If pos <> "" And lastRpt <> "" Then
                'Set f = New FRptViewer
                s = PathAppend(HFApp.SystemFolder, "Estimating\PO Formats", lastRpt & ".rpt")
                'Call f.ShowReport(s, True, False, "PONumber", Mid(pos, 2))
            
                Set c = New ZybUtil.Crystal
'testme
                Call c.LoadODBCReport(s, HFApp.LoginDSN, HFApp.LoginDBUID, HFApp.LoginDBPWD)
                On Error Resume Next
                Call c.ParameterValue("DivisionID", HFApp.DivisionID)
                Call c.ParameterValue("PONumber", Mid(pos, 2))
                On Error GoTo eh
                Call c.PrintPreview("Print Preview")
            
            End If
            lastRpt = "" & rs("poformat")
            pos = ""
        End If
        pos = pos & "," & DbQuote(Str, "" & rs("PONumber"))
        rs.MoveNext
    Wend
    If pos <> "" And lastRpt <> "" Then
        'Set f = New FRptViewer
        s = PathAppend(HFApp.SystemFolder, "Estimating\PO Formats", lastRpt & ".rpt")
        'Call f.ShowReport(s, True, False, "PONumber", Mid(pos, 2))
    
        Set c = New ZybUtil.Crystal
'testme
        Call c.LoadODBCReport(s, HFApp.LoginDSN, HFApp.LoginDBUID, HFApp.LoginDBPWD)
        On Error Resume Next
        Call c.ParameterValue("DivisionID", HFApp.DivisionID)
        Call c.ParameterValue("PONumber", Mid(pos, 2))
        On Error GoTo eh
        Call c.PrintPreview("Print Preview")
    End If
Exit Sub
eh:
    Select Case Err.Number
    Case 438 'File not found.
        MsgBox "Unable to open PO format """ & lastRpt & """." & vbCrLf & vbCrLf & "File not found:" & vbCrLf & PathAppend(HFApp.SystemFolder, "Estimating\PO Formats", lastRpt & ".rpt"), vbExclamation, App.ProductName
    Case Else: Call errHandler(SRCFILE & "ShowSelectedPOs", s)
    End Select
End Sub

Public Sub GroupGrid()
On Error GoTo eh

    Dim GroupedColumns As Long
    Dim i As Long
    
    With gHeaders
    .Redraw = flexRDNone
        
        GroupedColumns = 0
        For i = 0 To .cols - 1
            If .ColData(i) = "GROUPED" Then
                .ColPosition(i) = GroupedColumns
                GroupedColumns = GroupedColumns + 1
            End If
        Next
        
        If GroupedColumns = 0 Then
            Call .SubTotal(flexSTClear)
        Else
            Call .SubTotal(flexSTClear)
            .Col = 0
            .ColSel = GroupedColumns - 1
            .Sort = flexSortGenericAscending
            .OutlineCol = 0
            .SubtotalPosition = flexSTAbove
            For i = 0 To GroupedColumns - 1
                .SubTotal flexSTSum, i, .ColIndex("Total"), , , vbHighlight, True, "%s", 0
            Next
        End If
    
    .Redraw = flexRDBuffered
    End With
    
Exit Sub
eh: Call errHandler(SRCFILE & "GroupGrid")
End Sub

Private Sub ClearGroups()
    Dim i As Long
    With gHeaders
    For i = 0 To .cols - 1
        .ColData(i) = ""
    Next
    End With
End Sub



Private Sub gHeaders_AfterMoveColumn(ByVal Col As Long, Position As Long)
    Call GroupGrid
End Sub
