VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FCommunityStandards 
   Caption         =   "Community Standards"
   ClientHeight    =   3180
   ClientLeft      =   8910
   ClientTop       =   3435
   ClientWidth     =   8685
   FillColor       =   &H00808000&
   Icon            =   "CommunityStandards.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3180
   ScaleWidth      =   8685
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   1815
      Left            =   60
      TabIndex        =   0
      Top             =   600
      Width           =   8595
      _cx             =   15161
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
      Rows            =   4
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"CommunityStandards.frx":000C
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
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   1058
      ButtonWidth     =   820
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "Save"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList LargeIcons 
         Left            =   4260
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   39
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":0145
               Key             =   "EditAssembly"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":0A1F
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":12F9
               Key             =   "ExcelImport"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":1BD3
               Key             =   "ExcelExport"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":24AD
               Key             =   "Publish"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":2D87
               Key             =   "Forecast"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":3661
               Key             =   "ViewPOs"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":3F3B
               Key             =   "ViewBudgets"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":4815
               Key             =   "Open"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":50EF
               Key             =   "Preview"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":59C9
               Key             =   "Send"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":62A3
               Key             =   "TakeoffOneTime"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":6B7D
               Key             =   "Estimate"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":7457
               Key             =   "TakeoffAssembly"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":7D31
               Key             =   "NewAssembly"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":860B
               Key             =   "New"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":8EE5
               Key             =   "TakeoffItem"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":97BF
               Key             =   "TakeoffCustom"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":A099
               Key             =   "Save"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":A973
               Key             =   "SaveAs"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":B24D
               Key             =   "Delete"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":BB27
               Key             =   "RePrice"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":C401
               Key             =   "PricebookSearch"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":CCDB
               Key             =   "Pricebook"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":D5B5
               Key             =   "PricebookEdit"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":DE8F
               Key             =   "PricebookExport"
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":E769
               Key             =   "PricebookImport"
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":F043
               Key             =   "PricebookNew"
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":F91D
               Key             =   "View"
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":101F7
               Key             =   "Vendor1"
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":10AD1
               Key             =   "Vendor"
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":113AB
               Key             =   "Add"
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":11C85
               Key             =   "Attachments"
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":11F9F
               Key             =   ""
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":12879
               Key             =   ""
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":13153
               Key             =   "Design Center Options"
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":13A2D
               Key             =   "Global Options"
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":14307
               Key             =   "Models and Options"
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CommunityStandards.frx":14BE1
               Key             =   "Generate"
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FCommunityStandards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FCommunityStandards::"

Private mDirty   As Boolean
Private Sub Form_Load()
On Error GoTo eh
    Call SetToolbarIcons(Toolbar, LargeIcons)
    Call IniGetForm(Me)
    Call IniGetGrid(Me, gData)
    Call LoadCustomDescriptions
    Call LoadData
Exit Sub
eh: Call errHandler(SRCFILE & "Form_Load")
End Sub
Private Sub Form_Resize()
On Error Resume Next
    gData.Move 0, Toolbar.Height, Me.ScaleWidth, Me.ScaleHeight - Toolbar.Height
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    If Not SaveData(True) Then
        Cancel = True
        Exit Sub
    End If
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gData)
End Sub

Private Sub gData_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    mDirty = True
    gData.RowData(Row) = "dirty"
End Sub

Private Sub gData_AfterMoveColumn(ByVal Col As Long, Position As Long)
    gData.OutlineCol = gData.ColIndex("StdItemDesc")
End Sub

Private Sub gData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gData
        .ComboList = ""
        Select Case True
            Case gData.RowOutlineLevel(Row) = 0:  Cancel = True
            Case Col = .ColIndex("Phase"):        .ComboList = "..."
            Case Col = .ColIndex("Item"):         .ComboList = "..."
            Case Col = .ColIndex("ItemDesc"):     .ComboList = "..."
            
            Case Col = .ColIndex("Notes"):        .ComboList = "|..."
                                                  Cancel = .TextMatrix(Row, .ColIndex("Item")) = ""
                                                  
            Case Else:                            Cancel = True
        End Select
    End With
End Sub

Private Sub gData_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
Dim mr As Long
    With gData
        Select Case True
        
            Case .MouseRow = 0 And Button = vbRightButton
                Cancel = True
                Call FMain.ShowColumnMenu(gData)
                
        End Select
    End With

End Sub

Private Sub gData_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim s As String
    With gData
        Select Case .ColKey(Col)
            Case "Phase", "Item", "ItemDesc":
                s = "select phase,item,description,notes from tblphaseitem where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Item", s) Then
                    .TextMatrix(Row, .ColIndex("Phase")) = FPickList.SelectedItem("Phase")
                    .TextMatrix(Row, .ColIndex("Item")) = FPickList.SelectedItem("Item")
                    .TextMatrix(Row, .ColIndex("ItemDesc")) = FPickList.SelectedItem("Description")
                    .TextMatrix(Row, .ColIndex("Notes")) = FPickList.SelectedItem("Notes")
                    Call gData_AfterEdit(Row, Col)
                End If
                
            Case "Notes"
                s = .Text
                If FComments.Edit(s, gData, , "Notes", 4000) Then
                    .Text = s
                    Call gData_AfterEdit(Row, Col)
                End If
                
        End Select
    End With

End Sub

Private Sub gData_KeyDown(KeyCode As Integer, Shift As Integer)
    With gData
    Select Case KeyCode
        
        Case vbKeyLeft
            If .GetNodeRow(.Row, flexNTFirstChild) <> -1 And .IsCollapsed(.Row) <> flexOutlineCollapsed Then
                .IsCollapsed(.Row) = flexOutlineCollapsed
                .Col = .ColIndex("StdItemDesc")
            ElseIf .Col = GridNextVisibleColumn(gData, 0) Then
                If .GetNodeRow(.Row, flexNTParent) <> -1 Then .Row = .GetNodeRow(.Row, flexNTParent)
            End If

        Case vbKeyRight
            If .GetNodeRow(.Row, flexNTFirstChild) <> -1 Then
                If .IsCollapsed(.Row) = flexOutlineCollapsed Then
                    .IsCollapsed(.Row) = flexOutlineExpanded
                   .Col = .ColIndex("StdItemDesc")
                Else
                    .Row = .GetNodeRow(.Row, flexNTFirstChild)
                End If
            End If
        
        Case vbKeyDelete
            If .RowOutlineLevel(.Row) < 1 Then Exit Sub
            Select Case .ColKey(.Col)
                Case "Phase", "Item", "ItemDesc"
                    .TextMatrix(.Row, .ColIndex("Phase")) = ""
                    .TextMatrix(.Row, .ColIndex("Item")) = ""
                    .TextMatrix(.Row, .ColIndex("ItemDesc")) = ""
                    .TextMatrix(.Row, .ColIndex("Notes")) = ""
                    Call gData_AfterEdit(.Row, .Col)

                Case "Notes"
                    .Text = ""
                    mDirty = True
                    Call gData_AfterEdit(.Row, .Col)
                    
            End Select
            
    End Select
    End With
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case UCase(Trim(Button.key))
        Case "SAVE"
            Call SaveData(False)
    End Select
End Sub






Private Sub LoadData()
    Dim s As String
    Dim rs As Recordset
    Dim i As Long
    Dim comdesc As String
    
    s = ""
    'community phases
    s = s & "SELECT c.Area Community" & vbCrLf
    s = s & "      ,p.CommunityPhase" & vbCrLf
    s = s & "      ,isnull(c.Description,'')+isnull(' ' + p.description,'') CommunityDesc" & vbCrLf
    s = s & "      ,ti.phase StdPhase" & vbCrLf
    s = s & "      ,ti.item StdItem" & vbCrLf
    s = s & "      ,ti.description StdItemDesc" & vbCrLf
    s = s & "      ,si.Phase" & vbCrLf
    s = s & "      ,si.Item" & vbCrLf
    s = s & "      ,si.description ItemDesc" & vbCrLf
    s = s & "      ,s.Notes" & vbCrLf
    s = s & "  FROM tblLocality c " & vbCrLf
    s = s & "       join DivisionCommunities d on d.Community = c.Area " & vbCrLf
    s = s & "       left outer join CommunityPhase p on(c.area=p.community and 1=" & IIf("True" = HFApp.Options.ValueByName("CommunityStandardsArePhaseSpecific"), 1, 2) & ")" & vbCrLf
    s = s & "       left outer join CommunityStandardItems t on(1=1)" & vbCrLf
    s = s & "       join tblPhaseItem ti on(ti.DivisionID = " & HFApp.DivisionID & " and t.phase=ti.phase and t.item=ti.item)" & vbCrLf
    s = s & "       left outer join CommunityStandards s on(s.StdPhase=t.Phase and s.stditem=t.item and s.community=c.area and s.communityphase=isnull(p.communityphase,''))" & vbCrLf
    s = s & "       left outer join tblPhaseItem si on(si.DivisionID = " & HFApp.DivisionID & " and s.phase=si.phase and s.item=si.item)" & vbCrLf
    s = s & " WHERE isnull(c.inactive,0)=0 and d.DivisionID = " & HFApp.DivisionID & vbCrLf
    s = s & "UNION" & vbCrLf
    'communities with no phase
    s = s & "SELECT c.Area Community" & vbCrLf
    s = s & "      ,null CommunityPhase" & vbCrLf
    s = s & "      ,isnull(c.Description,'') CommunityDesc" & vbCrLf
    s = s & "      ,ti.phase StdPhase" & vbCrLf
    s = s & "      ,ti.item StdItem" & vbCrLf
    s = s & "      ,ti.description StdItemDesc" & vbCrLf
    s = s & "      ,si.Phase" & vbCrLf
    s = s & "      ,si.Item" & vbCrLf
    s = s & "      ,si.description ItemDesc" & vbCrLf
    s = s & "      ,s.Notes" & vbCrLf
    s = s & "  FROM tblLocality c " & vbCrLf
    s = s & "       join DivisionCommunities d on d.Community = c.Area " & vbCrLf
    s = s & "       left outer join CommunityStandardItems t on(1=1)" & vbCrLf
    s = s & "       join tblPhaseItem ti on(ti.DivisionID = " & HFApp.DivisionID & " and t.phase=ti.phase and t.item=ti.item)" & vbCrLf
    s = s & "       left outer join CommunityStandards s on(s.StdPhase=t.Phase and s.stditem=t.item and s.community=c.area and s.communityphase='')" & vbCrLf
    s = s & "       left outer join tblPhaseItem si on(si.DivisionID = " & HFApp.DivisionID & " and s.phase=si.phase and s.item=si.item)" & vbCrLf
    s = s & " WHERE isnull(c.inactive,0)=0 and d.DivisionID = " & HFApp.DivisionID & vbCrLf
    s = s & "ORDER BY 1,2,4,5" & vbCrLf

    Set rs = HFApp.SqlExec(s, dbHomefront)
    
    With gData
        .Redraw = flexRDNone
        .Rows = 1
        .OutlineCol = .ColIndex("StdItemDesc")
        While Not rs.EOF
            
            If comdesc <> "" & rs("CommunityDesc") Then
                comdesc = "" & rs("CommunityDesc")
                .AddItem ""
                i = .Rows - 1
                .TextMatrix(i, .ColIndex("StdItemDesc")) = "" & rs("CommunityDesc")
                .Cell(flexcpPicture, i, .ColIndex("StdItemDesc")) = FMain.SmallIcons.ListImages.Item("folder").Picture
                .RowOutlineLevel(i) = 0
                .IsSubtotal(i) = True
            End If
            
            
            .AddItem ""
            i = .Rows - 1
            .RowOutlineLevel(i) = 1
            .IsSubtotal(i) = True
            .TextMatrix(i, .ColIndex("CommunityPhase")) = "" & rs("CommunityPhase")
            .TextMatrix(i, .ColIndex("Community")) = "" & rs("Community")
            .TextMatrix(i, .ColIndex("StdPhase")) = "" & rs("StdPhase")
            .TextMatrix(i, .ColIndex("StdItem")) = "" & rs("StdItem")
            .TextMatrix(i, .ColIndex("StdItemDesc")) = "" & rs("StdItemDesc")
            .TextMatrix(i, .ColIndex("Phase")) = "" & rs("Phase")
            .TextMatrix(i, .ColIndex("Item")) = "" & rs("Item")
            If "" & rs("Item") <> "" Then
                .TextMatrix(i, .ColIndex("ItemDesc")) = "" & rs("ItemDesc")
            End If
            .TextMatrix(i, .ColIndex("Notes")) = "" & rs("Notes")
    
    
            rs.MoveNext
        Wend
        Call .Outline(0)
        .Redraw = flexRDBuffered
    End With
    
    
End Sub

Private Function SaveData(prompt As Boolean) As Boolean
On Error GoTo eh
    Dim i As Long
    Dim s As String

    If Not mDirty Then
        SaveData = True
        Exit Function
    End If
    If prompt Then
        Select Case MsgBox("Community Standards have changed." & vbCrLf & vbCrLf & "Do you want to save these changes?" & vbCrLf, vbExclamation + vbYesNoCancel, App.ProductName)
            Case vbNo
                SaveData = True
                Exit Function
            Case vbCancel
                SaveData = False
                Exit Function
        End Select
    End If

    
    
    Screen.MousePointer = vbHourglass
    With gData
        For i = 1 To .Rows - 1
            If .RowOutlineLevel(i) = 1 And .RowData(i) <> "" Then
                s = ""
                s = s & "INSERT INTO CommunityStandards(Community,CommunityPhase,StdPhase,StdItem,Phase,Item,Notes)" & vbCrLf
                s = s & "VALUES(" & DbQuote(Str, .TextMatrix(i, .ColIndex("Community"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("CommunityPhase"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("StdPhase"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("StdItem"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Phase"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Item"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Notes"))) & ")"
                Call HFApp.SqlExec(s, dbHomefront)
            
                s = ""
                s = s & "UPDATE CommunityStandards" & vbCrLf
                s = s & "SET Phase=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Phase"))) & vbCrLf
                s = s & "   ,Item=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Item"))) & vbCrLf
                If .TextMatrix(.Row, .ColIndex("Item")) = "" Then
                    s = s & "   ,Notes=''" & vbCrLf
                Else
                    s = s & "   ,Notes=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Notes"))) & vbCrLf
                End If
                s = s & "WHERE Community=" & DbQuote(Str, .TextMatrix(i, .ColIndex("Community"))) & vbCrLf
                s = s & "  AND CommunityPhase=" & DbQuote(Str, .TextMatrix(i, .ColIndex("CommunityPhase"))) & vbCrLf
                s = s & "  AND StdPhase=" & DbQuote(Str, .TextMatrix(i, .ColIndex("StdPhase"))) & vbCrLf
                s = s & "  AND StdItem=" & DbQuote(Str, .TextMatrix(i, .ColIndex("StdItem"))) & vbCrLf
                Call HFApp.SqlExec(s, dbHomefront)
                
                s = "Delete from CommunityStandards where isnull(Phase,'')='' and isnull(Item,'')='' and isnull(StdPhase,'')<>'' and isnull(Community,'')<>''"
                Call HFApp.SqlExec(s, dbHomefront)
                
                .RowData(i) = ""
            End If
        Next
    End With
    SaveData = True
    mDirty = False
    Screen.MousePointer = vbDefault

Exit Function
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Call errHandler(SRCFILE & "SaveData", s)
    End If
End Function


Private Sub LoadCustomDescriptions()
    Me.Caption = FMain.CD_Community & " Standards"
End Sub

