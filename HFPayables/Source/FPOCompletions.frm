VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FPOCompletions 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "PO Completions"
   ClientHeight    =   6150
   ClientLeft      =   2235
   ClientTop       =   2190
   ClientWidth     =   9105
   Icon            =   "FPOCompletions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   Begin HFPayables.VBCombo cboCommunity 
      Height          =   240
      Left            =   2160
      TabIndex        =   5
      Top             =   240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   423
   End
   Begin VB.CheckBox chkShowClosed 
      Caption         =   "Show closed PO's"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   5580
      Width           =   2475
   End
   Begin VB.CheckBox chkShowCompleted 
      Caption         =   "Show completed PO's"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   5400
      Width           =   2475
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Index           =   1
      Left            =   4320
      TabIndex        =   2
      Top             =   5460
      Width           =   1215
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&Save"
      Height          =   375
      Index           =   0
      Left            =   3000
      TabIndex        =   1
      Top             =   5460
      Width           =   1215
   End
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   4395
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   5115
      _cx             =   1999119166
      _cy             =   1999117896
      Appearance      =   1
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FPOCompletions.frx":000C
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
      OutlineBar      =   0
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
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "FPOCompletions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private Const SRCFILE = "FPOCompletions::"
Private mDirty As Boolean




Private Sub cboCommunity_Click()
    Call LoadData
End Sub

Private Sub chkShowClosed_Click()
    If SaveData(True) Then
        Call LoadData
        Call gData_ChangeEdit
    End If
End Sub

Private Sub chkShowCompleted_Click()
    If SaveData(True) Then
        Call LoadData
        Call gData_ChangeEdit
    End If
End Sub

Private Sub cmdNav_Click(Index As Integer)
    If Index = 0 Then If Not SaveData(False) Then Exit Sub
    Unload Me
End Sub

Private Sub Form_Load()
    Dim s As String
    
    chkShowClosed.Visible = TimberlineAccounting

    
    
    s = "select isnull(nullif(Custom_Description,''),'Community') as Description from customdescriptions where item ='Community'"
    Label1.Caption = "" & HFApp.SqlExec(s, dbHomefront)(0)
    Call IniGetForm(Me)
    
    
    Call LoadComboBox(cboCommunity, HFApp.Databases(dbHomefront), "select 'All','',0 union all SELECT isnull(area,'') + ' - ' + isnull(description,''),area,0 FROM tblLocality where inactive = 0 order by 1")
    'Call LoadData
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not SaveData(True) Then
        Cancel = True
        Exit Sub
    End If
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gData)
End Sub

Private Sub Form_Resize()
On Error Resume Next
    gData.Move gData.left, gData.Top, Me.ScaleWidth - 2 * gData.left, Me.ScaleHeight - gData.Top - cmdNav(0).Height - 2 * gData.left
    
    chkShowCompleted.Top = Me.ScaleHeight - chkShowCompleted.Height - chkShowClosed.Height - 60
    chkShowClosed.Top = Me.ScaleHeight - chkShowClosed.Height - 60
    
    cmdNav(0).Move Me.ScaleWidth - 2 * cmdNav(0).Width - 2 * gData.left, Me.ScaleHeight - cmdNav(0).Height - gData.left
    cmdNav(1).Move Me.ScaleWidth - cmdNav(0).Width - gData.left, Me.ScaleHeight - cmdNav(0).Height - gData.left
End Sub

Private Sub LoadData()
    Dim rs As New Recordset
    Dim s As String
    Dim Community As String
    Community = "" & Parse(GetComboBoxListKey(cboCommunity), 1, Chr(2))
    
    Screen.MousePointer = vbHourglass
    
    
    
    If TimberlineAccounting Then
        s = ""
        s = s & "SELECT sub " & App.Options(Caption_Commitment) & vbCrLf
        s = s & "      ,sdesc Description" & vbCrLf
        s = s & "      ,sVendor Vendor" & vbCrLf
        s = s & "      ,vname ""Company Name""" & vbCrLf
        s = s & "      ,samt+sapprco Amount" & vbCrLf
        s = s & "      ,samtinv Invoiced" & vbCrLf
        s = s & "      ,samt+sapprco-samtinv Remaining" & vbCrLf
        s = s & "      ,sactcd Completed" & vbCrLf
        s = s & "      ,sclosed Closed" & vbCrLf
        s = s & "  FROM master_jcm_record_12" & vbCrLf
        s = s & "      ,master_apm_record_9" & vbCrLf
        s = s & " WHERE svendor=vendor" & vbCrLf
        If chkShowCompleted.Value = vbUnchecked Then
            s = s & "   AND sactcd IS NULL" & vbCrLf
        End If
        If chkShowClosed.Value = vbUnchecked Then
            s = s & "   AND sclosed<>1" & vbCrLf
        End If
        Set gData.DataSource = HFApp.SqlExec(s, dbAccountingDictionary)
    Else
        gData.Cols = 1
        s = ""
        s = s & "select PO " & App.Options(Caption_Commitment) & vbCrLf
        s = s & "      ,POIndex                                   Description" & vbCrLf
        s = s & "      ,Vendor " & App.Options(Caption_Vendor) & vbCrLf
        s = s & "      ,VendorDesc" & vbCrLf
        s = s & "      ,linejobcommunity                          Community" & vbCrLf
        s = s & "      ,linejob                                   Job" & vbCrLf
        s = s & "      ,linejoblot                                Lot" & vbCrLf
        s = s & "      ,sum(linepretax)                           Amount" & vbCrLf
        's = s & "      ,poApproveddate                            Completed" & vbCrLf
        s = s & "      ,pocompleted                            Completed" & vbCrLf
        s = s & "      ,cast(case when poApprovedDate is null then 0 else 1 end as bit) Approved" & vbCrLf
        s = s & "      ,sum(lineinvoicedpretax)                   Invoiced" & vbCrLf
        s = s & "      ,sum(lineremainingpretax)                  Remaining" & vbCrLf
        s = s & "      ,sum(lineremainingtax)                     RemainingTax" & vbCrLf
        s = s & "      ,sum(lineremainingpretax+lineremainingtax) TotalRemaining" & vbCrLf
        s = s & "      ,podesc                                    Reference" & vbCrLf
        s = s & " from jcpodetails " & vbCrLf
        s = s & "where postingbatch<>0" & vbCrLf
        s = s & "  AND divisionid=" & DbQuote(num, HFApp.DivisionID) & vbCrLf
        If Community <> "" Then
            s = s & "   AND linejobcommunity=" & DbQuote(Str, Community) & vbCrLf
        End If
        
        If chkShowCompleted.Value = vbChecked Then
            s = s & "   AND pocompleted is not null" & vbCrLf
        Else
            s = s & "   AND pocompleted is null" & vbCrLf
        End If
'        If chkShowClosed.Value = vbUnchecked Then
'            s = s & "   AND sclosed<>1" & vbCrLf
'        End If
        
        s = s & "group by po,poindex,Vendor,VendorDesc,linejobcommunity,linejob,linejoblot,poApproveddate,pocompleted,podesc" & vbCrLf
        s = s & "having sum(lineremainingpretax)<>0" & vbCrLf
    
        Set gData.DataSource = HFApp.SqlExec(s, dbHomefront)
        gData.ColDataType(4) = flexDTString

    End If
    
    gData.DataMode = flexDMBoundNoRowCount

    Call gData.AddItem("", 1)
    gData.ColDataType(4) = flexDTString
    gData.Cell(flexcpBackColor, 1, 0, 1, gData.Cols - 1) = vbInfoBackground
    gData.FrozenRows = 1
    
    If Not TimberlineAccounting Then
        gData.ColKey(0) = "PONumber"
        gData.ColKey(1) = "Description"
        gData.ColKey(2) = "Vendor"
        gData.ColKey(3) = "VendorDesc"
        gData.ColKey(4) = "Community"
        gData.ColKey(5) = "Job"
        gData.ColKey(6) = "Lot"
        gData.ColKey(7) = "Amount"
        gData.ColKey(8) = "Completed"
        gData.ColKey(9) = "Approved"
        gData.ColKey(10) = "VPO"
        gData.ColKey(11) = "Invoiced"
        gData.ColKey(12) = "Remaining"
        gData.ColKey(13) = "RemainingTax"
        gData.ColKey(14) = "TotalRemaining"
        gData.ColKey(14) = "Reference"
        
        gData.ColDataType(8) = flexDTDate
        gData.ColDataType(7) = flexDTCurrency
        gData.ColFormat(4) = "&&&&&&&&&&"
        gData.ColFormat(5) = "&&&&&&&&&&"
        gData.ColFormat(6) = "&&&&&&&"
        gData.ColFormat(11) = "#,###.##"
        gData.ColFormat(12) = "#,###.##"
        gData.ColFormat(13) = "#,###.##"
        gData.ColFormat(14) = "#,###.##"
        gData.ColFormat(7) = "#,###.##"
        gData.ColAlignment(7) = flexAlignRightCenter
        gData.ColFormat(8) = App.Options(TimberlineDateFormat)
        Call IniGetGrid(Me, gData)
    Else
        gData.ColFormat(7) = App.Options(TimberlineDateFormat)
    End If
    
    
    mDirty = False
    cmdNav(0).Enabled = False
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub gData_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim r As Long
Dim c As Long
    With gData
        If Row = 1 Then
            .Redraw = flexRDNone
            For r = 2 To .Rows - 1
                .RowHidden(r) = False
                'If .EditText <> "" Then
                    For c = 0 To .Cols - 1
                        If c = .ColIndex("Job") And Val(.Cell(flexcpTextDisplay, 1, c)) <> 0 Then
                            If Not (UCase(.Cell(flexcpTextDisplay, r, c)) Like "*" & UCase(Val(.Cell(flexcpTextDisplay, 1, c))) & "*") Then
                                .RowHidden(r) = True
                                Exit For
                            End If

                        Else
                            If Not (UCase(.Cell(flexcpTextDisplay, r, c)) Like "*" & UCase(.Cell(flexcpTextDisplay, 1, c)) & "*") Then
                                .RowHidden(r) = True
                                Exit For
                            End If
                        End If
                    Next
                'End If
            Next
            .Redraw = flexRDBuffered
            Exit Sub
        End If
    End With
 
 
 With gData
    If Not TimberlineAccounting Then
        For r = Min(Row, .RowSel) To Max(Row, .RowSel)
            If Col = .ColIndex("Completed") Then
                 .Cell(flexcpChecked, r, .ColIndex("Approved"), r, .ColIndex("Approved")) = IIf(.Cell(flexcpText, r, .ColIndex("Approved")) <> "", flexChecked, flexUnchecked)
            ElseIf Col = .ColIndex("Approved") Then
                If .Cell(flexcpChecked, r, .ColIndex("Approved"), r, .ColIndex("Approved")) = flexUnchecked Then
                    .TextMatrix(r, .ColIndex("Completed")) = ""
                End If
                If .TextMatrix(r, .ColIndex("Completed")) = "" Then
                    .TextMatrix(r, .ColIndex("Completed")) = "" & IIf(.Cell(flexcpChecked, r, .ColIndex("Approved")) = flexChecked, Format(Now(), App.Options(TimberlineDateFormat)), "")
                End If
            End If
        Next
    End If
End With
End Sub

Private Sub gData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gData
        .AutoSearch = flexSearchNone
        .ComboList = ""
        Select Case True
            Case Row = 0
                Cancel = True
            
            Case Row = 1
                Cancel = True
            
            Case .ColIndex("Approved") = Col ' = 9
            Case .ColIndex("Completed") = Col '=8
                .ComboList = "..."
                
            Case Else
                Cancel = True
                .AutoSearch = flexSearchFromCursor
                
        End Select
    End With
End Sub

Private Sub gData_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    With gData
        Call DCalendar.Popup(gData, .RowPos(.Row) + .RowHeight(.Row), .colPos(.Col))
        If Not TimberlineAccounting Then
'        For i = Min(.Row, .RowSel) To Max(.Row, .RowSel)
'            If .RowHidden(i) = False Then
'                .RowStatus(i) = flexrsModified
'            End If
'        Next
            .Cell(flexcpChecked, Min(Row, .RowSel), .ColIndex("Approved"), Max(Row, .RowSel), .ColIndex("Approved")) = flexChecked
        End If
        For i = Min(.Row, .RowSel) To Max(.Row, .RowSel)
            If .RowHidden(i) = False Then
                .RowStatus(i) = flexrsModified
            End If
        Next
        cmdNav(0).Enabled = True
        mDirty = True
    End With
End Sub

Private Sub gData_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
On Error Resume Next
'change selection to only single column
    Static bInHere As Boolean
    If bInHere Then Exit Sub
    bInHere = True
    If gData.Row = 1 And gData.RowSel <> 1 Then gData.RowSel = 1
    If gData.RowSel = 1 And gData.Row <> 1 Then gData.RowSel = 2
    gData.ColSel = gData.Col
    bInHere = False
End Sub

Private Sub gData_AfterSort(ByVal Col As Long, Order As Integer)
    gData.FixedRows = 1
End Sub
Private Sub gData_BeforeSort(ByVal Col As Long, Order As Integer)
    gData.FixedRows = 2
End Sub
Private Sub gData_ChangeEdit()
' apply filter to hide stuff that doesn't match
'    Dim i As Long
'    With gData
'        If .Row <> 1 Then Exit Sub
'        .Redraw = flexRDNone
'
'        For i = 0 To 2
'            If i <> .Col Then .TextMatrix(1, i) = ""
'        Next
'
'
'        For i = 2 To .Rows - 1
'            .RowHidden(i) = Not (.Cell(flexcpTextDisplay, i, .Col) Like "*" & .EditText & "*")
'        Next
'        .Redraw = flexRDBuffered
'    End With
End Sub

Private Sub ShowColumnMenu(Grid As VSFlexGrid, Optional Sortable As Boolean = True, Optional Hideable As Boolean = True)
On Error GoTo eh
    Dim i As Long
    Dim j As Long
    
    'save this stuff for menu click
    Set MouseGrid = Grid
    MouseCol = Grid.MouseCol
    
    'set these

    
    'load Grid column names
    FMain.mnuColumnsSub(0).Visible = True
    For i = FMain.mnuColumnsSub.UBound To 1 Step -1
        Unload FMain.mnuColumnsSub(i)
    Next
    For i = 0 To MouseGrid.Cols - 1
        If MouseGrid.ColHidden(i) And MouseGrid.TextMatrix(0, i) <> "" Then
            j = j + 1
            Load FMain.mnuColumnsSub(j)
            FMain.mnuColumnsSub(j).tag = MouseGrid.ColKey(i)
            FMain.mnuColumnsSub(j).Caption = MouseGrid.TextMatrix(0, i)
            FMain.mnuColumnsSub(j).Visible = True
            FMain.mnuColumnsSub(j).Enabled = True
        End If
    Next
    If j = 0 Then FMain.mnuColumnsSub(0).Caption = "(none available)"
    FMain.mnuColumnsSub(0).Visible = j = 0
    FMain.mnuColumnsSub(0).Enabled = False
    
    'show menu
    PopupMenu FMain.mnuGrid

    Exit Sub
eh: Call ErrHandler(SRCFILE & "ShowColumnMenu")
End Sub

Private Sub gData_DblClick()
On Error Resume Next
Dim s As String
    Dim rs As Recordset
'    Dim v As New HFPrinter.ReportViewer
    Dim PONumber As String
    If gData.Row <> 1 Then
    
        PONumber = gData.TextMatrix(gData.Row, 0)
        
        s = ""
        s = s & "SELECT rpt = CASE " & vbCrLf
        s = s & "               WHEN ISNULL(tblvendors.poformat,'')<>'' THEN tblvendors.poformat" & vbCrLf
        s = s & "               WHEN ISNULL(tblpoindex.poformat,'')<>'' THEN tblpoindex.poformat" & vbCrLf
        s = s & "               ELSE " & DbQuote(Str, HFApp.Options(POFormat)) & vbCrLf
        s = s & "             END" & vbCrLf
        s = s & "  FROM POMaster" & vbCrLf
        s = s & "       LEFT OUTER JOIN tblPOIndex ON(POMaster.DivisionID = tblPOIndex.DivisionID and POMaster.POIndex=tblPOIndex.POIndex)" & vbCrLf
        s = s & "       LEFT OUTER JOIN tblVendors ON(POMaster.DivisionID = tblVendors.DivisionID and POMaster.Vendor=tblVendors.Vendor_id)" & vbCrLf
        s = s & " WHERE POMaster.DivisionID =" & HFApp.DivisionID & " and POMaster.PONumber=" & DbQuote(Str, PONumber) & vbCrLf
        Set rs = HFApp.SqlExec(s)
        
        Screen.MousePointer = vbHourglass
        s = PathAppend(HFApp.SystemFolder, "Estimating\PO Formats", "" & rs(0) & ".rpt")
        'Call v.ShowReport(HFApp.ConnectionString(dbHomeFront), s, rvPreview, "", "", "PONumber", PONumber)
        
        Dim c As New ZybUtil.Crystal
        Call c.LoadODBCReport(s, HFApp.LoginDSN, HFApp.LoginDBUID, HFApp.LoginDBPWD)
        On Error Resume Next
        Call c.ParameterValue("DivisionID", HFApp.DivisionID)
        Call c.ParameterValue("PONumber", PONumber)
        On Error GoTo 0
        Call c.PrintPreview("Print Preview")
        
        
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub gData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With gData
        If Button <> vbRightButton Or .MouseRow < 0 Or .MouseRow > .Rows - 2 Or .MouseCol < 0 Then Exit Sub
        
        If .MouseRow = 0 Then
            If gData.MouseRow = 0 Then Call ShowColumnMenu(gData, , gData.TextMatrix(0, gData.MouseCol) <> "")
        Else
            Call .Select(.MouseRow, .MouseCol)
            'mnuEditSub(mcEDIT_INQUIRY).Enabled = .TextMatrix(.Row, .ColIndex("commitment")) <> ""
            PopupMenu FMain.mnuEdit
        End If
    
    End With

End Sub

Private Sub gData_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim i As Long
    With gData
        Select Case Col
            Case .ColIndex("Approved")
                mDirty = True
                cmdNav(0).Enabled = True
'                If .Cell(flexcpChecked, Row, 8) = flexUnchecked Then
' '                   If .RowSel >= 2 Then
'                        '.Cell(flexcpText, Min(Row, Max(.RowSel, 2)), 7, Max(Row, Max(.RowSel, 2)), 7) = "" & Format(Now(), App.Options(TimberlineDateFormat))
' '                   Else
'
'                        .Cell(flexcpText, Row, 7, .RowSel, 7) = "" & Format(Now(), App.Options(TimberlineDateFormat))
' '                   End If
'
'                    '.TextMatrix(Row, 7) = Format(Now(), App.Options(TimberlineDateFormat))
'                Else
'                    .Cell(flexcpText, Row, 7, .RowSel, 7) = ""
'                    '.TextMatrix(Row, 7) = ""
'                End If
            Case .ColIndex("Closed")
                mDirty = True
                cmdNav(0).Enabled = True
                
            Case .ColIndex("Completed") '7
                If IsDate(.EditText) Or .EditText = "" Then
                    .EditText = Format(.EditText, App.Options(TimberlineDateFormat))
                    mDirty = True
                    cmdNav(0).Enabled = True
                Else
                    Cancel = True
                End If
        End Select
        For i = Min(.Row, .RowSel) To Max(.Row, .RowSel)
            If .RowHidden(i) = False Then
                .RowStatus(i) = flexrsModified
            End If
        Next
    End With
End Sub

Private Function SaveData(prompt As Boolean) As Boolean
    Dim s As String
    Dim i As Long
    
    
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
    
    
    With gData
        For i = .Rows - 1 To 2 Step -1
            If .RowStatus(i) = flexrsModified Then
'THIS IS WRONG. DONT CHANGE THE APPROVED STAMPS
'                If TimberlineAccounting Then
'                    s = ""
'                    s = s & "UPDATE master_jcm_record_12" & vbCrLf
'                    s = s & "SET sclosed=" & IIf(.Cell(flexcpChecked, i, 8) = flexChecked, 1, 0) & vbCrLf
'                    s = s & "   ,sactcd=" & DbQuote(Date, .TextMatrix(i, 7)) & vbCrLf
'                    s = s & "WHERE sub=" & DbQuote(Str, .TextMatrix(i, 0)) & vbCrLf
'                    Call HFApp.SqlExec(s, dbAccountingDictionary)
'                    If .TextMatrix(i, 7) <> "" Then
'                        s = ""
'                        s = s & "UPDATE POMaster" & vbCrLf
'                        s = s & "SET Approvedby=" & DbQuote(Str, HFApp.LoginID) & vbCrLf
'                        s = s & "   ,ApprovedDate=" & DbQuote(Date, .TextMatrix(i, 7)) & vbCrLf
'                        s = s & "WHERE PONumber=" & DbQuote(Str, .TextMatrix(i, 0)) & vbCrLf
'                    Else
'                        s = ""
'                        s = s & "UPDATE POMaster" & vbCrLf
'                        s = s & "SET Approvedby=''" & vbCrLf
'                        s = s & "   ,ApprovedDate=null" & vbCrLf
'                        s = s & "WHERE PONumber=" & DbQuote(Str, .TextMatrix(i, 0)) & vbCrLf
'
'                    End If
'                Else
'                If .Cell(flexcpChecked, i, 9) = flexChecked Then
'                    s = ""
'                    s = s & "UPDATE POMaster" & vbCrLf
'                    s = s & "SET Approvedby=" & DbQuote(Str, HFApp.LoginID) & vbCrLf
'                    s = s & "   ,ApprovedDate=" & DbQuote(Date, .TextMatrix(i, 8)) & vbCrLf
'                    s = s & "WHERE PONumber=" & DbQuote(Str, .TextMatrix(i, 0)) & vbCrLf
'                Else
'                    s = ""
'                    s = s & "UPDATE POMaster" & vbCrLf
'                    s = s & "SET Approvedby=''" & vbCrLf
'                    s = s & "   ,ApprovedDate=null" & vbCrLf
'                    s = s & "WHERE PONumber=" & DbQuote(Str, .TextMatrix(i, 0)) & vbCrLf
'
'                End If
'                End If
'
'                Call HFApp.SqlExec(s, dbHomefront)
'                Call .RemoveItem(i)
            End If
        Next
    End With
    
    
    SaveData = True
    mDirty = False
    cmdNav(0).Enabled = False
    
    
End Function
