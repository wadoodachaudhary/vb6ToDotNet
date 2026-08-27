VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form FRFPWizard 
   Caption         =   "RFQ Wizard"
   ClientHeight    =   4785
   ClientLeft      =   1185
   ClientTop       =   2265
   ClientWidth     =   6330
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FRFPWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   6330
   Begin VB.PictureBox WizFoot 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   6330
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4200
      Width           =   6330
      Begin VB.CommandButton cmdNav 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Index           =   0
         Left            =   1860
         TabIndex        =   0
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "< &Back"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   3060
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Next >"
         Height          =   375
         Index           =   2
         Left            =   4200
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Finish"
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   5400
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   26540
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   26480
         Y1              =   15
         Y2              =   15
      End
   End
   Begin HFEst.WizHead WizHead1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   1588
      Caption         =   "Generate RFQ's"
      Description     =   "This wizard will help you create request for quotes."
      Icon            =   "FRFPWizard.frx":000C
   End
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3660
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   870
      Width           =   6345
      Begin VSFlex8Ctl.VSFlexGrid gData 
         Height          =   1500
         Index           =   0
         Left            =   1320
         TabIndex        =   7
         Top             =   1560
         Width           =   4965
         _cx             =   8758
         _cy             =   2646
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
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   3
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FRFPWizard.frx":08E6
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   1
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   5
         OutlineCol      =   0
         Ellipsis        =   1
         ExplorerBar     =   2
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
         BackColorFrozen =   -2147483624
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label Label3 
         Caption         =   "Click Next to continue, where you can choose which RFQ's you would like to produce."
         Height          =   465
         Index           =   7
         Left            =   1260
         TabIndex        =   14
         Top             =   1080
         Width           =   4710
      End
      Begin VB.Label Label3 
         Caption         =   $"FRFPWizard.frx":0938
         Height          =   795
         Index           =   6
         Left            =   1260
         TabIndex        =   13
         Top             =   360
         Width           =   4710
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "How should the RFQ's be created?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1170
         TabIndex        =   8
         Top             =   90
         Width           =   2985
      End
      Begin VB.Image imgError 
         Height          =   480
         Left            =   330
         Picture         =   "FRFPWizard.frx":09E4
         Top             =   270
         Width           =   480
      End
   End
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Caption         =   "3"
      Height          =   3600
      Index           =   1
      Left            =   1800
      TabIndex        =   9
      Top             =   4710
      Width           =   6315
      Begin VSFlex8Ctl.VSFlexGrid gData 
         Height          =   2280
         Index           =   1
         Left            =   1170
         TabIndex        =   10
         Top             =   1320
         Width           =   5175
         _cx             =   9128
         _cy             =   4022
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
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   3
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FRFPWizard.frx":12AE
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   1
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   5
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   2
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
         BackColorFrozen =   -2147483624
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.CheckBox chkSelectAll 
         Caption         =   "Select All"
         Height          =   315
         Index           =   1
         Left            =   5280
         TabIndex        =   11
         Top             =   990
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   $"FRFPWizard.frx":12D7
         Height          =   795
         Index           =   1
         Left            =   1260
         TabIndex        =   15
         Top             =   360
         Width           =   4950
      End
      Begin VB.Image Image2 
         Height          =   480
         Index           =   2
         Left            =   330
         Picture         =   "FRFPWizard.frx":13BA
         Top             =   270
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Select RFQ's"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   1170
         TabIndex        =   12
         Top             =   90
         Width           =   1125
      End
   End
End
Attribute VB_Name = "FRFPWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FRFPWizard::"

Private mJob    As String
Private bInHere As Boolean



Private Sub chkSelectAll_Click(Index As Integer)
On Error Resume Next
    If Not bInHere Then
        gData(Index).Cell(flexcpChecked, 1, 1, gData(Index).Rows - 1, 1) = IIf(chkSelectAll(Index).value = vbChecked, flexChecked, flexUnchecked)
    End If
End Sub

Public Sub ShowForm(Job As String)
    mJob = Job
    bInHere = False
    Me.Show vbModal
End Sub


Private Sub Form_Load()
    Dim i As Long
    Dim s As String
    Dim rs As Recordset
    
    Call IniGetForm(Me)
    
    'show fields
    Set rs = HFApp.SqlExec("select * from tbljobs where DivisionID = " & HFApp.DivisionID & " and job_no=" & DbQuote(Str, mJob))
    With gData(0)
        .Rows = 4
        .RowData(0) = "POIndex"
        .RowData(1) = "Location"
        .RowData(2) = "HFDescription"
        .RowData(3) = "JCCostCodeDesc"
        If Not rs.EOF Then
            For i = 1 To 40
                If "" & rs("WBSDesc" & format(i, "00")) <> "" Then
                    .AddItem "" & rs("WBSDesc" & format(i, "00"))
                    .RowData(.Rows - 1) = "WBS" & format(i, "00")
                End If
            Next
        End If
        For i = 0 To .Rows - 1
            .Cell(flexcpChecked, i, 0) = flexUnchecked
        Next
    End With
    
    Call SetCurrentFrame(0)
End Sub

Private Function GetCurrentFrame() As Long
    Dim i As Long
    For i = 0 To WizFrame.UBound
        If WizFrame(i).Visible = True Then
            GetCurrentFrame = i
            Exit Function
        End If
    Next
    GetCurrentFrame = WizFrame.LBound
End Function

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub

Private Sub cmdNav_Click(Index As Integer)
On Error Resume Next
    Select Case Index
        Case 0: Unload Me
        Case 1: Call SetCurrentFrame(GetCurrentFrame - 1)
        Case 2: Call SetCurrentFrame(GetCurrentFrame + 1)
        Case 3: If CreateRFPs Then Unload Me
    End Select
End Sub

Private Sub gData_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    
    If Index <> 1 Then Exit Sub
    bInHere = True
    With gData(Index)
        For i = 2 To .Rows - 1
            If .Cell(flexcpChecked, 1, 1) <> .Cell(flexcpChecked, i, 1) Then
                chkSelectAll(Index).value = vbGrayed
                bInHere = False
                Exit Sub
            End If
        Next
        chkSelectAll(Index).value = IIf(.Cell(flexcpChecked, 0, 0) = flexChecked, vbChecked, vbUnchecked)
    End With
    bInHere = False

End Sub


Private Sub gData_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gData(Index)
        .ComboList = ""
        .AutoSearch = flexSearchNone
        If Index = 0 Then
            .AutoSearch = flexSearchFromCursor
        Else
            Select Case Col
                Case 1
                Case .Cols - 1
                    .ComboList = "..."
                Case Else
                    .AutoSearch = flexSearchFromCursor
                    Cancel = True
            End Select
        End If
    End With
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Const margin = 60
    Dim i As Long
    For i = 0 To WizFrame.UBound
        WizFrame(i).Move 0, WizHead1.Height, Me.ScaleWidth, Me.ScaleHeight - WizHead1.Height - Me.WizFoot.Height
        gData(i).Move gData(i).Left, gData(i).Top, WizFrame(i).Width - gData(i).Left - 2 * margin, WizFrame(i).Height - gData(i).Top - 2 * margin
        chkSelectAll(i).Left = gData(i).Left + gData(i).Width - chkSelectAll(i).Width
    Next
    cmdNav(0).Move Me.ScaleWidth - (4 * (margin + cmdNav(0).Width)), 2 * margin
    cmdNav(1).Move Me.ScaleWidth - (3 * (margin + cmdNav(0).Width)), 2 * margin
    cmdNav(2).Move Me.ScaleWidth - (2 * (margin + cmdNav(0).Width)), 2 * margin
    cmdNav(3).Move Me.ScaleWidth - (1 * (margin + cmdNav(0).Width)), 2 * margin
    
End Sub

Private Function GroupByClause(Named As Boolean) As String
    Dim s As String
    Dim i As Long
    s = ""
    With gData(0)
        For i = 0 To .Rows - 1
            If .Cell(flexcpChecked, i, 0) = flexChecked Then
                s = s & "," & .RowData(i)
                
'                If Named Then s = s & " " & vbQuote & .TextMatrix(i, 0) & vbQuote
            End If
        Next
    End With
    GroupByClause = Mid(s, 2)
End Function


Private Sub gData_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim s As String
    Dim v As String
    Dim i As Long
    With gData(Index)
        
        s = "select vendorgroupid Type , vendor_id Vendor,Vendor_name Company from tblvendors where inactive=0 and divisionid=" & DbQuote(Num, HFApp.DivisionID)
        
        If Not FPickList.Choose(HFApp.Databases(dbHomefront), "Vendor", s, , , , , IIf(HFApp.Options(AccountingSystem) = asQuickBooks, "Vendor", ""), True) Then Exit Sub
        s = ""
        v = ""
        For i = 1 To FPickList.SelectedItems
            s = s & "," & FPickList.SelectedItem("Company", i)
            v = v & "," & DbQuote(Str, FPickList.SelectedItem("Vendor", i))
        Next
        .TextMatrix(Row, Col) = Mid(s, 2)
        .Cell(flexcpData, Row, Col) = Mid(v, 2)
        .Cell(flexcpChecked, Row, 1) = flexChecked
    End With
End Sub

Private Sub gData_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And Index = 1 And gData(Index).Col = gData(Index).Cols - 1 Then gData(Index).Text = ""
End Sub

Private Sub SetCurrentFrame(RHS As Long)
    Dim i As Long
    Dim r As Long
    Dim c As Long
    Dim s As String
    Dim rs As Recordset
    
    If RHS = 1 And GroupByClause(False) = "" Then
        MsgBox "You have to pick at least one field.", vbInformation, App.ProductName
        RHS = 0
    End If
    
    
    For i = 0 To WizFrame.UBound
        WizFrame(i).Visible = i = RHS
    Next
    cmdNav(1).Enabled = RHS > 0
    cmdNav(2).Enabled = RHS < WizFrame.UBound
    cmdNav(3).Enabled = RHS = WizFrame.UBound
    
    
    Select Case RHS
        Case 0
        Case 1
            'load unique values
            s = ""
            s = s & "select distinct isnull(" & Replace(GroupByClause(False), ",", ",'')+ ', ' + isnull(") & ",'un-named'),'', " & GroupByClause(True) & ",'' Vendors" & vbCrLf
            s = s & "  from estimateditems " & vbCrLf
            s = s & " where isnull(RFP,'')=''" & vbCrLf
            s = s & "   and divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
            s = s & "   and job_no=" & DbQuote(Str, mJob) & vbCrLf
            s = s & "order by " & GroupByClause(False) & vbCrLf
            With gData(1)
                'clear
                .Rows = 1
                .FixedRows = 1
                .ExplorerBar = flexExSortShow
                
                
                r = 1
                .Redraw = flexRDNone
                Set rs = HFApp.SqlExec(s)
                'load columns
                .Cols = rs.fields.Count
                For c = 0 To .Cols - 1
                    .TextMatrix(0, c) = rs.fields(c).Name
                    .ColData(c) = rs.fields(c).Name
                    .ColHidden(c) = c = 0
                Next
                'load data
                While Not rs.EOF
                    .AddItem ""
                    For c = 0 To .Cols - 1
                        .TextMatrix(r, c) = "" & rs(c)
                    Next
                    r = r + 1
                    rs.MoveNext
                Wend
                'set checkboxes
                .Cell(flexcpChecked, 1, 1, .Rows - 1, 1) = flexUnchecked
                chkSelectAll(1).value = vbUnchecked
                .Redraw = flexRDBuffered
                Call .AutoSize(0, .Cols - 1)
                .ColWidth(1) = 255
            End With
            
    End Select
    
End Sub


Private Function CreateRFPs() As Boolean
On Error GoTo eh
    
    Dim r As Long
    Dim c As Long
    Dim s As String
    Dim d As String
    Dim vendors As String
    
    With gData(1)
    For r = 1 To .Rows - 1
        If .Cell(flexcpChecked, r, 1) = flexChecked Then
            
            vendors = .Cell(flexcpData, r, .Cols - 1)
            
            'insert RFP then update EstimateItems
            s = ""
            s = s & "INSERT INTO RFPs(divisionid,Job,Title) VALUES(" & DbQuote(Num, HFApp.DivisionID) & "," & DbQuote(Str, mJob) & "," & DbQuote(Str, .TextMatrix(r, 0), , True, 250) & ")" & vbCrLf
            s = s & vbCrLf
            s = s & "UPDATE EstimateItems" & vbCrLf
            s = s & "   SET RFP=IDENT_CURRENT('RFPs')" & vbCrLf
            s = s & " WHERE EstItemID IN(SELECT EstItemID " & vbCrLf
            s = s & "                      FROM EstimatedItems " & vbCrLf
            s = s & "                     WHERE ISNULL(RFP,0)=0" & vbCrLf
            s = s & "                       AND divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
            s = s & "                       AND job_no=" & DbQuote(Str, mJob) & vbCrLf
            For c = 2 To .Cols - 2
            s = s & "                       AND " & .ColData(c) & "=" & DbQuote(Str, .TextMatrix(r, c)) & vbCrLf
            Next
            s = s & ")"
            s = s & "" & vbCrLf
            
            If vendors <> "" Then
                s = s & "INSERT INTO VendorBids(RFP,Vendor,CreatedDate)" & vbCrLf
                s = s & "SELECT IDENT_CURRENT('RFPs'),vendor_id,getdate()" & vbCrLf
                s = s & "FROM tblVendors" & vbCrLf
                s = s & "WHERE divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
                s = s & "and vendor_id in(" & vendors & ")" & vbCrLf
                s = s & vbCrLf
                
                s = s & "INSERT INTO VendorBidItems(RFP,Vendor,EstItemID)" & vbCrLf
                s = s & "SELECT rfp,vendor_id,estitemid" & vbCrLf
                s = s & "FROM tblvendors v" & vbCrLf
                s = s & "    ,estimateitems e" & vbCrLf
                s = s & "WHERE e.rfp=IDENT_CURRENT('RFPs')" & vbCrLf
                s = s & "  AND v.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
                s = s & "  AND v.vendor_id in(" & .Cell(flexcpData, r, .Cols - 1) & ")" & vbCrLf
                s = s & vbCrLf
            End If
            
            Call HFApp.SqlExec(s)
            
            
        End If
    Next
    End With
    
    
    CreateRFPs = True
Exit Function
eh: Call errHandler(SRCFILE & "CreateRFPs", s)

End Function



