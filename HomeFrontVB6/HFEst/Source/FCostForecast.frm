VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form FCostForecast 
   Caption         =   "Cost Forecasting"
   ClientHeight    =   7275
   ClientLeft      =   5610
   ClientTop       =   2310
   ClientWidth     =   7635
   Icon            =   "FCostForecast.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7275
   ScaleWidth      =   7635
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   180
      TabIndex        =   1
      Top             =   3840
      Width           =   7395
      Begin VB.CheckBox chkCalcPrices 
         Caption         =   "Calculate new vendor price forecasts"
         Height          =   435
         Left            =   1560
         TabIndex        =   9
         Top             =   180
         Width           =   3195
      End
      Begin VB.ComboBox cboForecastStart 
         Enabled         =   0   'False
         Height          =   240
         Left            =   2460
         TabIndex        =   6
         Top             =   660
         Width           =   1815
      End
      Begin VB.CommandButton cmdUpdateForecasts 
         Caption         =   "Update Forecasted Prices"
         Height          =   735
         Left            =   4680
         TabIndex        =   2
         Top             =   780
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.ComboBox cboForecastEnd 
         Enabled         =   0   'False
         Height          =   240
         Left            =   2460
         TabIndex        =   7
         Top             =   915
         Width           =   1815
      End
      Begin VB.ComboBox cboCostBasis 
         Enabled         =   0   'False
         Height          =   240
         Left            =   2460
         TabIndex        =   8
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   660
         Top             =   660
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cost Basis"
         Height          =   195
         Index           =   2
         Left            =   1575
         TabIndex        =   5
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Calculate"
         Height          =   195
         Index           =   3
         Left            =   1725
         TabIndex        =   4
         Top             =   720
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "thru"
         Height          =   195
         Index           =   4
         Left            =   2100
         TabIndex        =   3
         Top             =   960
         Width           =   270
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   420
         Top             =   420
         Width           =   480
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   3435
      Left            =   180
      TabIndex        =   0
      Top             =   420
      Width           =   6735
      _cx             =   11880
      _cy             =   6059
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
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   13
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FCostForecast.frx":000C
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
Attribute VB_Name = "FCostForecast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const SRCFILE = "FCostForecast::"
Private mDirty   As Boolean

Private mMouseCol As Long

Private Const mcEDIT_COPY = 0
Private Const mcEDIT_PASTE = 1

Private Sub chkCalcPrices_Click()
    Dim b As Boolean
    b = chkCalcPrices.value = vbChecked
    cboCostBasis.Enabled = b
    cboForecastStart.Enabled = b
    cboForecastEnd.Enabled = b
    Label1(2).ForeColor = IIf(b, vbWindowText, vbGrayText)
    Label1(3).ForeColor = IIf(b, vbWindowText, vbGrayText)
    Label1(4).ForeColor = IIf(b, vbWindowText, vbGrayText)
    
    cmdUpdateForecasts.Visible = chkCalcPrices.value = vbChecked
    
End Sub

Private Sub Form_Load()
    Image1.Picture = FMain.LargeIcons.ListImages("RePrice").Picture
    Image2.Picture = FMain.LargeIcons.ListImages("Vendor").Picture
    Call IniGetGrid(Me, gData)
    Call LoadData
End Sub


Public Function SaveData(prompt As Boolean) As Boolean
On Error GoTo eh:

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
    With gData
        For r = 1 To .Rows - 1
            s = ""
            s = s & "UPDATE tblPOIndex" & vbCrLf
            s = s & "   SET ForecastPercent1=" & DbQuote(Num, .ValueMatrix(r, .ColIndex("Forecast1"))) & vbCrLf
            s = s & "      ,ForecastPercent2=" & DbQuote(Num, .ValueMatrix(r, .ColIndex("Forecast2"))) & vbCrLf
            s = s & "      ,ForecastPercent3=" & DbQuote(Num, .ValueMatrix(r, .ColIndex("Forecast3"))) & vbCrLf
            s = s & "      ,ForecastPercent4=" & DbQuote(Num, .ValueMatrix(r, .ColIndex("Forecast4"))) & vbCrLf
            s = s & "      ,ForecastPercent5=" & DbQuote(Num, .ValueMatrix(r, .ColIndex("Forecast5"))) & vbCrLf
            s = s & "      ,ForecastPercent6=" & DbQuote(Num, .ValueMatrix(r, .ColIndex("Forecast6"))) & vbCrLf
            s = s & "      ,ForecastPercent7=" & DbQuote(Num, .ValueMatrix(r, .ColIndex("Forecast7"))) & vbCrLf
            s = s & "      ,ForecastPercent8=" & DbQuote(Num, .ValueMatrix(r, .ColIndex("Forecast8"))) & vbCrLf
            s = s & "      ,ForecastPercent9=" & DbQuote(Num, .ValueMatrix(r, .ColIndex("Forecast9"))) & vbCrLf
            s = s & "      ,ForecastPercent10=" & DbQuote(Num, .ValueMatrix(r, .ColIndex("Forecast10"))) & vbCrLf
            s = s & "      ,ForecastPercent11=" & DbQuote(Num, .ValueMatrix(r, .ColIndex("Forecast11"))) & vbCrLf
            s = s & "      ,ForecastPercent12=" & DbQuote(Num, .ValueMatrix(r, .ColIndex("Forecast12"))) & vbCrLf
            s = s & " WHERE DivisionID in(0," & HFApp.DivisionID & ") and POIndex=" & DbQuote(Str, Mid(.Cell(flexcpData, r, .ColIndex("POIndex")), 2)) & vbCrLf
            Call HFApp.SqlExec(s)
        Next
    End With
    mDirty = False
    SaveData = True
    
    Exit Function
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Function




Private Sub LoadData()
    Dim s As String
    Dim rs As Recordset
    Dim i As Long
    
    
    
    
    
    With gData
        .Redraw = flexRDNone
        .Rows = 1
        
        
        'get forecast names
        Call LoadCostTypes(cboCostBasis)
        Call LoadCostTypes(cboForecastStart, , True)
        Call LoadCostTypes(cboForecastEnd, , True)
        For i = 1 To 12
            .TextMatrix(0, i) = cboForecastStart.list(i - 1)
        Next
        cboCostBasis.ListIndex = 0
        cboForecastStart.ListIndex = 0
        cboForecastEnd.ListIndex = 0
        
        s = ""
        s = s & "SELECT POIndex,description" & vbCrLf
        s = s & "      ,ForecastPercent1" & vbCrLf
        s = s & "      ,ForecastPercent2" & vbCrLf
        s = s & "      ,ForecastPercent3" & vbCrLf
        s = s & "      ,ForecastPercent4" & vbCrLf
        s = s & "      ,ForecastPercent5" & vbCrLf
        s = s & "      ,ForecastPercent6" & vbCrLf
        s = s & "      ,ForecastPercent7" & vbCrLf
        s = s & "      ,ForecastPercent8" & vbCrLf
        s = s & "      ,ForecastPercent9" & vbCrLf
        s = s & "      ,ForecastPercent10" & vbCrLf
        s = s & "      ,ForecastPercent11" & vbCrLf
        s = s & "      ,ForecastPercent12" & vbCrLf
        s = s & "  FROM tblPOIndex where DivisionID = " & HFApp.DivisionID
        Set rs = HFApp.SqlExec(s)
        i = 0
        While Not rs.EOF
            i = i + 1
            .AddItem ""
            
            .Cell(flexcpData, i, .ColIndex("POIndex")) = "K" & rs("POIndex")
            s = "" & rs("POIndex")
            If s <> "" & rs("Description") Then s = s & " " & rs("Description")
            .TextMatrix(i, .ColIndex("POIndex")) = s
            .Cell(flexcpPicture, i, .ColIndex("POIndex")) = FMain.SmallIcons.ListImages("purchaseorder").Picture
            
            .TextMatrix(i, .ColIndex("Forecast1")) = "" & rs("ForecastPercent1")
            .TextMatrix(i, .ColIndex("Forecast2")) = "" & rs("ForecastPercent2")
            .TextMatrix(i, .ColIndex("Forecast3")) = "" & rs("ForecastPercent3")
            .TextMatrix(i, .ColIndex("Forecast4")) = "" & rs("ForecastPercent4")
            .TextMatrix(i, .ColIndex("Forecast5")) = "" & rs("ForecastPercent5")
            .TextMatrix(i, .ColIndex("Forecast6")) = "" & rs("ForecastPercent6")
            .TextMatrix(i, .ColIndex("Forecast7")) = "" & rs("ForecastPercent7")
            .TextMatrix(i, .ColIndex("Forecast8")) = "" & rs("ForecastPercent8")
            .TextMatrix(i, .ColIndex("Forecast9")) = "" & rs("ForecastPercent9")
            .TextMatrix(i, .ColIndex("Forecast10")) = "" & rs("ForecastPercent10")
            .TextMatrix(i, .ColIndex("Forecast11")) = "" & rs("ForecastPercent11")
            .TextMatrix(i, .ColIndex("Forecast12")) = "" & rs("ForecastPercent12")
            rs.MoveNext
        Wend
        Call .AutoSize(0)
        Call .AutoSize(1, .Cols - 1, True)
        .Redraw = flexRDBuffered
    End With
    
End Sub



Private Sub Form_Resize()
On Error Resume Next
    Const margin = 60
    gData.Move margin, margin, Me.ScaleWidth - 2 * margin, Me.ScaleHeight - margin - Me.Frame1.Height
    Frame1.Move margin, gData.Top + gData.Height
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Not SaveData(True) Then Cancel = True
    Call IniPutGrid(Me, gData)
End Sub

Private Sub gData_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim r As Long
    Dim c As Long
    
    With gData
        For r = 1 To .Rows - 1
            For c = 1 To .Cols - 1
                .TextMatrix(r, c) = .ValueMatrix(r, c)
            Next
        Next
    End With

End Sub

Public Sub RenameForecast()
    Dim s As String
    Dim newname As String
    Dim c As Long
    With gData
        c = mMouseCol
        If c > 0 Then
            newname = Left(Trim(InputBox(vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Change description from """ & Trim(MouseCtrl.TextMatrix(0, MouseCol)) & """ to:", SpaceCase(MouseCtrl.ColKey(MouseCol)))), 30)
'            newname = left(Trim(InputBox("Enter a new name for this forecast.", App.ProductName, .TextMatrix(0, c))), 30)
            If newname <> "" Then
                .TextMatrix(0, c) = newname
                cboCostBasis.list(c + 5) = newname
                cboForecastEnd.list(c - 1) = newname
                cboForecastStart.list(c - 1) = newname
                
                s = ""
                s = s & "UPDATE CustomDescriptions" & vbCrLf
                s = s & "SET Custom_Description=" & DbQuote(Str, newname) & vbCrLf
                s = s & "WHERE Item=" & DbQuote(Str, "Forecast" & c) & vbCrLf
                Call HFApp.SqlExec(s)
            End If
        End If
    End With
End Sub

Private Sub gData_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim s As String
    Dim r As Long
    Dim c As Long
    
    With gData
        Select Case True
                
            Case KeyCode = vbKeyF And Shift = vbCtrlMask
                Call FFind.ShowForm(gData)
            
            Case Shift = vbCtrlMask And KeyCode = vbKeyC
                s = Replace(.Clip, vbCr, vbCrLf)
                Call Clipboard.SetText(s)
                
            Case Shift = vbCtrlMask And KeyCode = vbKeyV And gData.Row > 0 And gData.Col > 0
                s = Clipboard.GetText
                If .RowSel = .Row And .ColSel = .Col Then
                    s = Replace(Replace(s, vbLf, ""), vbCr, vbCrLf)
                    While Right(s, 2) = vbCrLf
                        s = Mid(s, 1, Len(s) - 2)
                    Wend
                    If s = "" Then Exit Sub
                    r = Parse(s, , vbCrLf)
                    For i = 1 To r
                        c = Max(c, Parse(Parse(s, r, vbCrLf), , vbTab))
                    Next
                    mDirty = True
                    .Cell(flexcpData, .Row, .Col, Min(.Rows - 1, .Row + r - 1), Min(.Cols - 1, .Col + c - 1)) = "DIRTY"
                    Call .Select(.Row, .Col, Min(.Rows - 1, .Row + r - 1), Min(.Cols - 1, .Col + c - 1))
                    .Clip = s
                    Call gData_AfterEdit(0, 0)
                Else
                    .Text = s
                    .Cell(flexcpData, .Row, .Col, .RowSel, .ColSel) = "DIRTY"
                End If
                
        End Select
    End With

End Sub



Private Sub gData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Col = 0
End Sub

Private Sub gData_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error Resume Next
    If Button = vbRightButton Then
        Cancel = True
        Call FMain.ShowColumnMenu(gData, False, False, gData.ColKey(gData.MouseCol) <> "POIndex")
    End If
End Sub

Private Sub gData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mMouseCol = gData.MouseCol
End Sub

Private Sub gData_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not IsNumeric(gData.EditText) Then
        Cancel = True
    Else
        mDirty = True
    End If
End Sub



Private Sub cmdUpdateForecasts_Click()
On Error GoTo eh
    Dim i         As Long
    Dim s         As String
    Dim u         As String
    Dim fcstBasis As String
    
    
    
    s = ""
    s = s & "This will update your forecasted cost database." & vbCrLf
    s = s & "It will re-calculate the following forecast indexes" & vbCrLf
    s = s & "for all vendors. Are you sure this is what you want" & vbCrLf
    s = s & "to do?" & vbCrLf & vbCrLf
    For i = cboForecastStart.ListIndex + 1 To cboForecastEnd.ListIndex + 1
        s = s & "   " & vbBullet & "  " & gData.TextMatrix(0, i) & vbCrLf
    Next
    If vbNo = MsgBox(s, vbYesNo + vbExclamation, App.ProductName) Then Exit Sub
    
    
    
    If Not SaveData(False) Then Exit Sub
    
    Screen.MousePointer = vbHourglass
        
    Select Case cboCostBasis.ListIndex
        Case 0:  fcstBasis = "Current_Cost"
        Case 1:  fcstBasis = "Next_Cost1"
        Case 2:  fcstBasis = "Next_Cost2"
        Case 3:  fcstBasis = "Last_Cost1"
        Case 4:  fcstBasis = "Last_Cost2"
        Case 5:  fcstBasis = "Last_Cost3"
        Case 6:  fcstBasis = "Forecast1"
        Case 7:  fcstBasis = "Forecast2"
        Case 8:  fcstBasis = "Forecast3"
        Case 9:  fcstBasis = "Forecast4"
        Case 10: fcstBasis = "Forecast5"
        Case 11: fcstBasis = "Forecast6"
        Case 12: fcstBasis = "Forecast7"
        Case 13: fcstBasis = "Forecast8"
        Case 14: fcstBasis = "Forecast9"
        Case 15: fcstBasis = "Forecast10"
        Case 16: fcstBasis = "Forecast11"
        Case 17: fcstBasis = "Forecast12"
    End Select
    
    'build update clause
    For i = cboForecastStart.ListIndex + 1 To cboForecastEnd.ListIndex + 1
        u = u & "      ,Forecast" & i & " = " & fcstBasis & " + (" & fcstBasis & " * ForecastPercent" & i & " / 100)" & vbCrLf
    Next
    u = "   SET " & Mid(u, 8)
    
    
    If u <> "" Then
        
       'Updated March 15, 2014
        s = ""
        s = s & "alter table tblvendorcost disable trigger all" & vbCrLf
        Call HFApp.SqlExec(s)
        s = ""
        s = s & "" & vbCrLf
        s = s & "UPDATE tblVendorCost" & vbCrLf
        s = s & u
        s = s & "FROM tblVendorCost t1" & vbCrLf
        s = s & "INNER JOIN tblPhaseItem t2 ON (t1.DivisionID = t2.DivisionID and t1.Phase=t2.Phase and t1.Item = t2.Item)" & vbCrLf
        s = s & "INNER JOIN tblPOIndex p on (p.DivisionID = t2.DivisionID and p.POIndex = t2.POIndex)" & vbCrLf
        s = s & "WHERE t1.DivisionID =" & HFApp.DivisionID & vbCrLf
        Call HFApp.SqlExec(s)
        s = ""
        s = s & "alter table tblvendorcost enable trigger all" & vbCrLf
        
        
        Call HFApp.SqlExec(s)

    
    End If
    Screen.MousePointer = vbDefault
    
Exit Sub
eh: Call errHandler(SRCFILE & "cmdUpdateForecasts_Click", s)
End Sub


Public Sub RenameColumn(ColumnIndex As Long)
    Dim s As String
    Dim i As Long
    Dim newname As String
    newname = InputBox(vbCrLf & vbCrLf & "Enter a new description for " & gData.ColKey(ColumnIndex) & ".", "Change Description", gData.TextMatrix(0, ColumnIndex))
    If newname <> "" Then
    
        i = Mid(gData.ColKey(ColumnIndex), 9)
        i = Max(1, Min(12, i))
        
        cboForecastStart.list(i - 1) = newname
        cboForecastEnd.list(i - 1) = newname
        cboCostBasis.list(i + 5) = newname
        
        gData.TextMatrix(0, ColumnIndex) = newname
        
        

        s = ""
        s = s & "UPDATE CustomDescriptions" & vbCrLf
        s = s & "SET Custom_Description=" & DbQuote(Str, newname) & vbCrLf
        s = s & "WHERE Item=" & DbQuote(Str, gData.ColKey(ColumnIndex)) & vbCrLf
        Call HFApp.SqlExec(s)
        
        
        
        
    End If
    
End Sub
