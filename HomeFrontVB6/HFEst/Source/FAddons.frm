VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FAddons 
   Caption         =   "Addons"
   ClientHeight    =   3525
   ClientLeft      =   1590
   ClientTop       =   4380
   ClientWidth     =   8730
   Icon            =   "FAddons.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   8730
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13515
      _cx             =   23839
      _cy             =   4683
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
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   18
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FAddons.frx":058A
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
      ExplorerBar     =   0
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
      Begin VB.Label lblMoveLine 
         BackColor       =   &H8000000D&
         Height          =   75
         Index           =   0
         Left            =   30
         TabIndex        =   1
         Top             =   330
         Visible         =   0   'False
         Width           =   75
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "<Popup>"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupSub 
         Caption         =   "Edit..."
         Index           =   0
      End
      Begin VB.Menu mnuPopupSub 
         Caption         =   "Insert..."
         Index           =   1
      End
      Begin VB.Menu mnuPopupSub 
         Caption         =   "Insert SubTotal"
         Index           =   2
      End
      Begin VB.Menu mnuPopupSub 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuPopupSub 
         Caption         =   "Delete"
         Index           =   4
      End
   End
End
Attribute VB_Name = "FAddons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private Const SRCFILE = "FAddons::"


Private mJob        As String
Private mAssemblyID As Long

Private mMat As Double
Private mLab As Double
Private mSub As Double
Private mOth As Double
Private mEq As Double

Private Const mc_EDIT = 0
Private Const mc_INSERTADDON = 1
Private Const mc_INSERTSUBTOTAL = 2
Private Const mc_DELETE = 4

Private Sub gData_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call ReCalc
End Sub

'------------------------------------
' cost basis's (including non-addons)
'------------------------------------
'LumpSum
'Category        budgeted cost of 1 or more cost types plus markup
'LastSubTotal    previous subtotal line plus markup
'RunningTotal    current total plus markup
'Cost            budgeted cost per category
'SubTotal        interface only no effect on costs
'Total           interface only no effect on costs


Private Sub gData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gData
    .EditMaxLength = 0
    .ComboList = ""
    Select Case .ColKey(Col)
        Case "Title"
            .EditMaxLength = 75
            Cancel = IsIn(.TextMatrix(.Row, .ColIndex("Basis")), "Cost", "SubTotal", "Total")
        
        Case "Rate"
            Cancel = IsIn(.TextMatrix(.Row, .ColIndex("Basis")), "Cost", "SubTotal", "Total", "LumpSum")
        
        Case "Amount"
            Cancel = IsIn(.TextMatrix(.Row, .ColIndex("Basis")), "Cost", "SubTotal", "Total")
        
        Case Else
            Cancel = True
    End Select
    End With
End Sub


Private Sub gData_DblClick()
    Call mnuPopupSub_Click(mc_EDIT)
End Sub

Private Sub gData_KeyDown(KeyCode As Integer, Shift As Integer)
    With gData
    Select Case True
        Case KeyCode = vbKeyDelete And Shift <> 0
            Call mnuPopupSub_Click(mc_DELETE)
    End Select
    End With
End Sub

Private Sub gData_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim a1 As Double
    Dim a2 As Double
    Dim p As Double
    
    With gData
    Select Case .ColKey(Col)
        Case "Title"
        
        
        Case "Amount"
            .TextMatrix(Row, .ColIndex("LumpSumAmt")) = Val(.EditText)
            .TextMatrix(Row, .ColIndex("RateType")) = "2" 'lumpsum
            .TextMatrix(Row, .ColIndex("Basis")) = "LumpSum"
            
        Case "Rate"
            If CBool(InStr(1, .EditText, "/", vbTextCompare)) Then
                .TextMatrix(Row, .ColIndex("Rate1")) = Val(Parse(.EditText, 1, "/"))
                .TextMatrix(Row, .ColIndex("Rate2")) = Val(Parse(.EditText, 2, "/"))
                .TextMatrix(Row, .ColIndex("RateType")) = "1" 'amt/amt
            Else
                .TextMatrix(Row, .ColIndex("Percentage")) = Val(.EditText)
                If Not IsIn(.TextMatrix(Row, .ColIndex("RateType")), "0", "3") Then
                    .TextMatrix(Row, .ColIndex("RateType")) = "0" 'percent
                End If
            End If
        
    End Select
    End With
End Sub

Private Sub LoadData()
    Dim s  As String
    Dim rs As Recordset
    Dim r  As Long
    
    s = ""
    s = s & "select *" & vbCrLf
    s = s & "  from addons" & vbCrLf
    s = s & " where job=" & DbQuote(str, mJob) & vbCrLf
    s = s & "   and assemblyid=" & DbQuote(Num, mAssemblyID) & vbCrLf
    s = s & "order by sequence" & vbCrLf
    Set rs = HFApp.SqlExec(s, dbHomefront)
    With gData
        .Rows = 1
        While Not rs.EOF
            If "" & rs("Basis") <> "Total" Then
                .AddItem ""
                r = r + 1
                .TextMatrix(r, .ColIndex("AddonID")) = "" & rs("AddonID")
                .TextMatrix(r, .ColIndex("Title")) = "" & rs("Title")
                .TextMatrix(r, .ColIndex("Basis")) = "" & rs("Basis")
                .TextMatrix(r, .ColIndex("RateType")) = "" & rs("RateType")
                .TextMatrix(r, .ColIndex("Percentage")) = Val("" & rs("Percentage"))
                .TextMatrix(r, .ColIndex("Rate1")) = Val("" & rs("AmountPer1"))
                .TextMatrix(r, .ColIndex("Rate2")) = Val("" & rs("AmountPer2"))
                .TextMatrix(r, .ColIndex("LumpSumAmt")) = Val("" & rs("LumpSumAmt"))
                .TextMatrix(r, .ColIndex("Mat")) = "" & rs("Mat")
                .TextMatrix(r, .ColIndex("Lab")) = "" & rs("Lab")
                .TextMatrix(r, .ColIndex("Sub")) = "" & rs("Sub")
                .TextMatrix(r, .ColIndex("Oth")) = "" & rs("Oth")
                .TextMatrix(r, .ColIndex("Eq")) = "" & rs("Eq")
            End If
            rs.MoveNext
        Wend
        
        'none found in db so is first time here. add costs
        If .Rows = 1 Then
            .AddItem ""
            r = r + 1
            .TextMatrix(r, .ColIndex("Basis")) = "Cost"
            .TextMatrix(r, .ColIndex("Lab")) = "True"
            .TextMatrix(r, .ColIndex("Title")) = "Labour"
            
            .AddItem ""
            r = r + 1
            .TextMatrix(r, .ColIndex("Basis")) = "Cost"
            .TextMatrix(r, .ColIndex("Mat")) = "True"
            .TextMatrix(r, .ColIndex("Title")) = "Material"
            
            .AddItem ""
            r = r + 1
            .TextMatrix(r, .ColIndex("Basis")) = "Cost"
            .TextMatrix(r, .ColIndex("Sub")) = "True"
            .TextMatrix(r, .ColIndex("Title")) = "Subcontract"
            
            .AddItem ""
            r = r + 1
            .TextMatrix(r, .ColIndex("Basis")) = "Cost"
            .TextMatrix(r, .ColIndex("Eq")) = "True"
            .TextMatrix(r, .ColIndex("Title")) = "Equipment"
            
            .AddItem ""
            r = r + 1
            .TextMatrix(r, .ColIndex("Basis")) = "Cost"
            .TextMatrix(r, .ColIndex("Oth")) = "True"
            .TextMatrix(r, .ColIndex("Title")) = "Other"
        End If
        
    
    
        'add grand total at bottom
        .AddItem ""
        r = r + 1
        .TextMatrix(r, .ColIndex("AddonID")) = -1
        .TextMatrix(r, .ColIndex("Basis")) = "Total"
    
            
    
     End With
    
    Call ReCalc
    
End Sub


Public Function SaveData(prompt As Boolean) As Boolean
On Error GoTo eh
    Dim r As Long
    Dim s As String
    
'    If Not mDirty Then
'        SaveData = True
'        Exit Function
'    End If
'    If prompt Then
'        Select Case MsgBox("This data has changed." & vbCrLf & vbCrLf & "Do you want to save these changes?" & vbCrLf, vbExclamation + vbYesNoCancel, App.ProductName)
'            Case vbNo
'                SaveData = True
'                Exit Function
'            Case vbCancel
'                SaveData = False
'                Exit Function
'        End Select
'    End If
    
    
    
    With gData
        
        s = ""
        s = s & "delete from addons" & vbCrLf
        s = s & " where job=" & DbQuote(str, mJob) & vbCrLf
        s = s & "   and assemblyid=" & DbQuote(Num, mAssemblyID) & vbCrLf
        Call HFApp.SqlExec(s, dbHomefront)
        
        
        For r = 1 To .Rows - 1
            s = ""
            s = s & "insert into addons(Job, AssemblyID, Sequence, Title, Basis, RateType, Percentage, AmountPer1, AmountPer2, LumpSumAmt, Mat, Eq, Lab, Sub, Oth,amount,totalamount,PercentOfTotal)" & vbCrLf
            s = s & "values(" & DbQuote(str, mJob) & vbCrLf
            s = s & "      ," & DbQuote(Num, mAssemblyID) & vbCrLf
            s = s & "      ," & DbQuote(Num, r) & vbCrLf
            s = s & "      ," & DbQuote(str, .TextMatrix(r, .ColIndex("Title"))) & vbCrLf
            s = s & "      ," & DbQuote(str, .TextMatrix(r, .ColIndex("Basis"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(r, .ColIndex("RateType"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(r, .ColIndex("Percentage"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(r, .ColIndex("Rate1"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(r, .ColIndex("Rate2"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(r, .ColIndex("LumpSumAmt"))) & vbCrLf
            s = s & "      ," & DbQuote(Bit, .TextMatrix(r, .ColIndex("Mat"))) & vbCrLf
            s = s & "      ," & DbQuote(Bit, .TextMatrix(r, .ColIndex("Eq"))) & vbCrLf
            s = s & "      ," & DbQuote(Bit, .TextMatrix(r, .ColIndex("Lab"))) & vbCrLf
            s = s & "      ," & DbQuote(Bit, .TextMatrix(r, .ColIndex("Sub"))) & vbCrLf
            s = s & "      ," & DbQuote(Bit, .TextMatrix(r, .ColIndex("Oth"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(r, .ColIndex("Amount"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(r, .ColIndex("Total"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(r, .ColIndex("PercentOfTotal"))) & ")"
            Call HFApp.SqlExec(s, dbHomefront)
        Next
        
    End With
Exit Function
eh: Call errHandler(SRCFILE & "SaveData", s)
End Function


Private Sub ReCalc()
    Dim r As Long
    
    Dim Amount       As Double
    Dim LastSubTotal As Double
    Dim CurrentTotal As Double
    Dim RunningTotal As Double
    
    With gData
    
        'clear formatting and celldata
        Call .Clear(0, 2)
        
        LastSubTotal = 0
        CurrentTotal = 0
        RunningTotal = 0
        For r = 1 To .Rows - 1
            
            'formatting
            Select Case .TextMatrix(r, .ColIndex("RateType"))
                Case 0, 3 'markup, margin
                    .TextMatrix(r, .ColIndex("Rate")) = .ValueMatrix(r, .ColIndex("Percentage")) & "%"
                Case 1 'amt/amt
                    .TextMatrix(r, .ColIndex("Rate")) = .ValueMatrix(r, .ColIndex("Rate1")) & " / " & .ValueMatrix(r, .ColIndex("Rate2"))
                Case 2 'lump
                    .TextMatrix(r, .ColIndex("Rate")) = ""
            End Select
            Select Case .TextMatrix(r, .ColIndex("Basis"))
                Case "Cost"
                    .TextMatrix(r, .ColIndex("RateType")) = "-1"
                    .TextMatrix(r, .ColIndex("Rate")) = ""
                Case "Subtotal"
                    .TextMatrix(r, .ColIndex("Title")) = ""
                    .TextMatrix(r, .ColIndex("Rate")) = ""
                    
                Case "Total"
                    .TextMatrix(r, .ColIndex("Title")) = "Total"
                    .TextMatrix(r, .ColIndex("Rate")) = ""
            End Select
            
            
            'calc amounts
            Amount = 0
            Select Case .TextMatrix(r, .ColIndex("Basis"))
                Case "Category", "Cost"
                    If .TextMatrix(r, .ColIndex("Mat")) = "True" Then Amount = Amount + mMat
                    If .TextMatrix(r, .ColIndex("Lab")) = "True" Then Amount = Amount + mLab
                    If .TextMatrix(r, .ColIndex("Sub")) = "True" Then Amount = Amount + mSub
                    If .TextMatrix(r, .ColIndex("Oth")) = "True" Then Amount = Amount + mOth
                    If .TextMatrix(r, .ColIndex("Eq")) = "True" Then Amount = Amount + mEq
                    .TextMatrix(r, .ColIndex("baseamount")) = Amount
                    Call ApplyMarkUp(r)
                    RunningTotal = RunningTotal + .ValueMatrix(r, .ColIndex("amount"))
                    CurrentTotal = CurrentTotal + .ValueMatrix(r, .ColIndex("amount"))
                    
                
                Case "LumpSum"
                    .TextMatrix(r, .ColIndex("baseamount")) = .TextMatrix(r, .ColIndex("LumpSumAmt"))
                    Call ApplyMarkUp(r)
                    RunningTotal = RunningTotal + .ValueMatrix(r, .ColIndex("amount"))
                    CurrentTotal = CurrentTotal + .ValueMatrix(r, .ColIndex("amount"))
                    
                Case "RunningTotal"
                    .TextMatrix(r, .ColIndex("baseamount")) = CurrentTotal
                    Call ApplyMarkUp(r)
                    RunningTotal = RunningTotal + .ValueMatrix(r, .ColIndex("amount"))
                    CurrentTotal = CurrentTotal + .ValueMatrix(r, .ColIndex("amount"))
                
                Case "LastSubTotal"
                    .TextMatrix(r, .ColIndex("baseamount")) = LastSubTotal
                    Call ApplyMarkUp(r)
                    RunningTotal = RunningTotal + .ValueMatrix(r, .ColIndex("amount"))
                    CurrentTotal = CurrentTotal + .ValueMatrix(r, .ColIndex("amount"))
                
                Case "SubTotal"
                    .TextMatrix(r, .ColIndex("amount")) = RunningTotal
                    .TextMatrix(r, .ColIndex("total")) = CurrentTotal
                    '.Cell(flexcpFontBold, r, .ColIndex("amount")) = True
                    '.Cell(flexcpFontBold, r, .ColIndex("total")) = True
                    '.Cell(flexcpFontBold, r, .ColIndex("Title")) = True
                    .Cell(flexcpFontBold, r, 0, r, .cols - 1) = True
                    
                    LastSubTotal = RunningTotal
                    RunningTotal = 0
                    
                
                Case "Total"
                    .TextMatrix(r, .ColIndex("total")) = CurrentTotal
                    '.Cell(flexcpFontBold, r, .ColIndex("total")) = True
                    '.Cell(flexcpFontBold, r, .ColIndex("Title")) = True
                    .Cell(flexcpFontBold, r, 0, r, .cols - 1) = True
                
            End Select
        Next
    
    
        For r = 1 To .Rows - 2
            If CurrentTotal <> 0 Then
                .TextMatrix(r, .ColIndex("percentoftotal")) = .ValueMatrix(r, .ColIndex("amount")) / CurrentTotal
            Else
                .TextMatrix(r, .ColIndex("percentoftotal")) = 0
            End If
        Next
    
        Call .AutoSize(0, .cols - 1, , 300)
    
    
    End With
    
End Sub

Private Function ApplyMarkUp(Row As Long) As Double
    Dim cost As Double
    Dim rate As Double

    With gData
        cost = .ValueMatrix(Row, .ColIndex("baseamount"))
        rate = .ValueMatrix(Row, .ColIndex("Percentage")) / 100
        
        Select Case .TextMatrix(Row, .ColIndex("RateType"))
            
            Case -1 '"Cost"
                .TextMatrix(Row, .ColIndex("amount")) = cost
            
            Case 1 '"AmountPerAmount"
                .TextMatrix(Row, .ColIndex("amount")) = Round(cost * .ValueMatrix(Row, .ColIndex("rate1")) / .ValueMatrix(Row, .ColIndex("rate2")), 2)
            
            Case 2 '"LumpSum"
                .TextMatrix(Row, .ColIndex("amount")) = cost
            
            Case 0 '"Markup"
                .TextMatrix(Row, .ColIndex("amount")) = Round(cost * rate, 2)
            
            Case 3 '"Margin"
                If rate = 1 Then rate = 0
                .TextMatrix(Row, .ColIndex("amount")) = (cost / (1 - rate)) - cost
        End Select
    End With
End Function

Public Sub ShowJob(Job As String, AssemblyID As Long)
    Dim s As String
    Dim rs As Recordset
    
    mJob = Job
    mAssemblyID = AssemblyID
    mMat = 0
    mLab = 0
    mSub = 0
    mOth = 0
    mEq = 0
    
    
    If AssemblyID <> 0 Then
        Set rs = HFApp.SqlExec("select hfdescription from estimateassemblies where estassemblyid=" & DbQuote(Num, AssemblyID))
        Me.Caption = "" & rs(0)
    Else
        Me.Caption = "Job"
    End If
    
    
    s = ""
    s = s & "select sum(e.budgetpretax+e.budgetjctax) pretax" & vbCrLf
    s = s & "      ,isnull(nullif(i.costcategory,''),case c.costtype when 1 then 'L' when 2 then 'M' when 3 then 'S' when 4 then 'E' else 'O' end) CostType" & vbCrLf
    s = s & "from estimateditems e " & vbCrLf
    s = s & "     left outer join tblphaseitem i on(e.DivisionID = i.DivisionID and e.estphase=i.phase and e.estitem=i.item)" & vbCrLf
    s = s & "     left outer join standardcategories c on(e.DivisionID = c.DivisionID and e.jccategory=c.category)" & vbCrLf
    s = s & "where e.DivisionID = " & HFApp.DivisionID & " and  isnull(e.Budgetdeleted,0) = 0 and e.job_no=" & DbQuote(str, Job) & vbCrLf
    If AssemblyID = 0 Then
        s = s & "  and e.isChangeRequest=0" & vbCrLf
    Else
        s = s & "  and e.estassemblyid=" & DbQuote(Num, AssemblyID) & vbCrLf
    End If
    s = s & "group by isnull(nullif(i.costcategory,''),case c.costtype when 1 then 'L' when 2 then 'M' when 3 then 'S' when 4 then 'E' else 'O' end)" & vbCrLf
    Set rs = HFApp.SqlExec(s, dbHomefront)
    While Not rs.EOF
        Select Case "" & rs("CostType")
        Case "M":  mMat = mMat + Val("" & rs("pretax"))
        Case "L":  mLab = mLab + Val("" & rs("pretax"))
        Case "E":  mEq = mEq + Val("" & rs("pretax"))
        Case "S":  mSub = mSub + Val("" & rs("pretax"))
        Case Else: mOth = mOth + Val("" & rs("pretax"))
        End Select
        rs.MoveNext
    Wend
    
    Me.Show vbModal
End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
    With gData
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
    End With
    Call LoadData
End Sub

Private Sub Form_Resize()
    Call gData.Move(0, 0, Me.ScaleWidth, Me.ScaleHeight)
    Call ReCalc
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveData(True)
    Call IniPutForm(Me)
End Sub


Private Sub gData_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error Resume Next
    With gData
    If Button = vbRightButton And .MouseCol > 0 And .MouseRow > 0 Then
        Call .Select(.MouseRow, .MouseCol)
        mnuPopupSub(mc_DELETE).Enabled = .Row > 0 And .Row <> .Rows - 1 And .TextMatrix(.Row, .ColIndex("Basis")) <> "Cost"
        mnuPopupSub(mc_EDIT).Enabled = mnuPopupSub(mc_DELETE).Enabled
        PopupMenu mnuPopup
    End If
    End With
End Sub



Private Sub lblMoveLine_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Top As Single
    Dim NewLine As Long
    Dim g As VSFlexGrid
    
    Select Case Index
        Case 0:  Set g = gData
    End Select
    If g.Editable = flexEDNone Then
        lblMoveLine(Index).Visible = False
        Exit Sub
    End If
    
    
    If Button <> 0 Then
        If g.Row <> g.RowSel Then g.Row = g.Row
        Top = lblMoveLine(Index).Top + Y
        NewLine = Top / g.RowHeight(0)
        
        If NewLine < 1 Then NewLine = 1
        
        
        If NewLine > g.Rows - 1 Then NewLine = g.Rows - 1
        Top = NewLine * g.RowHeight(0)
        lblMoveLine(Index).Move 0, Top, g.Width, 30
    End If



End Sub

Private Sub lblMoveLine_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Top As Single
    Dim NewLine As Long
    Dim g As VSFlexGrid
    Dim r As Long
    
    
    lblMoveLine(Index).Visible = False
    Select Case Index
        Case 0:  Set g = gData
    End Select
    If g.Editable = flexEDNone Then Exit Sub
    
    
    Top = lblMoveLine(Index).Top + 30
    NewLine = GetRowFromPoint(g, Top)
    If g.Row < NewLine Then NewLine = NewLine - 1
    
    If g.Row = NewLine Then
        lblMoveLine(0).Move g.CellLeft, g.CellTop, 60, 60
        lblMoveLine(Index).Visible = True
        Exit Sub
    End If
    
    g.RowPosition(g.Row) = NewLine
    g.Row = NewLine
    
    Call ReCalc
End Sub

Private Function GetRowFromPoint(g As VSFlexGrid, Y As Single) As Long
On Error GoTo done
    Dim r As Long
    With g
        r = 0
        While .RowPos(r) < Y
            r = r + 1
        Wend
    End With

done:
r = r - 1
GetRowFromPoint = r
End Function

Private Sub gData_RowColChange()
On Error Resume Next
    With gData
        If .Row = .Rows - 1 Then
            lblMoveLine(0).Visible = False
        Else
            lblMoveLine(0).Move .CellLeft, .CellTop, 60, 60
            lblMoveLine(0).Visible = True
        End If
    End With
End Sub

Private Sub gData_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    lblMoveLine(0).Visible = False
End Sub

Private Sub mnuPopupSub_Click(Index As Integer)
    Dim r As Long
    
    With gData
    Select Case Index
        Case mc_EDIT
            If .Row > 0 And .Row <> .Rows - 1 And .TextMatrix(.Row, .ColIndex("Basis")) <> "Cost" Then
                Call FAddon.EditAddon(gData)
            End If
            
        Case mc_INSERTADDON
            Call FAddon.CreateAddon(gData)
        
        Case mc_INSERTSUBTOTAL
            r = .Row
            If r = 0 Then r = 1
            Call .AddItem("", r)
            .TextMatrix(r, .ColIndex("Basis")) = "SubTotal"
            
            
        Case mc_DELETE
            If .Row > 0 And .Row <> .Rows - 1 And .TextMatrix(.Row, .ColIndex("Basis")) <> "Cost" Then .RemoveItem
            
    End Select
    End With
    Call ReCalc
End Sub

