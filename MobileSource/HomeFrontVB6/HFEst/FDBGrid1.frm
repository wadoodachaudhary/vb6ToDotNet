VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form FDBGrid 
   ClientHeight    =   3090
   ClientLeft      =   10065
   ClientTop       =   4980
   ClientWidth     =   10305
   Icon            =   "FDBGrid.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   10305
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   6555
      Left            =   0
      TabIndex        =   1
      Top             =   750
      Width           =   19035
      _cx             =   33576
      _cy             =   11562
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
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
      Editable        =   0
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
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   1005
      ButtonWidth     =   1296
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "   Save   "
            Key             =   "Save"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Export"
            Key             =   "ExcelExport"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Key             =   "New"
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
               Picture         =   "FDBGrid.frx":08CA
               Key             =   "EditAssembly"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":11A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":1A7E
               Key             =   "ExcelImport"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":2358
               Key             =   "ExcelExport"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":2C32
               Key             =   "Publish"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":350C
               Key             =   "Forecast"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":3DE6
               Key             =   "ViewPOs"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":46C0
               Key             =   "ViewBudgets"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":4F9A
               Key             =   "Open"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":5874
               Key             =   "Preview"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":614E
               Key             =   "Send"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":6A28
               Key             =   "TakeoffOneTime"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":7302
               Key             =   "Estimate"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":7BDC
               Key             =   "TakeoffAssembly"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":84B6
               Key             =   "NewAssembly"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":8D90
               Key             =   "New"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":966A
               Key             =   "TakeoffItem"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":9F44
               Key             =   "TakeoffCustom"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":A81E
               Key             =   "Save"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":B0F8
               Key             =   "SaveAs"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":B9D2
               Key             =   "Delete"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":C2AC
               Key             =   "RePrice"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":CB86
               Key             =   "PricebookSearch"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":D460
               Key             =   "Pricebook"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":DD3A
               Key             =   "PricebookEdit"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":E614
               Key             =   "PricebookExport"
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":EEEE
               Key             =   "PricebookImport"
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":F7C8
               Key             =   "PricebookNew"
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":100A2
               Key             =   "View"
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":1097C
               Key             =   "Vendor1"
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":11256
               Key             =   "Vendor"
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":11B30
               Key             =   "Add"
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":1240A
               Key             =   "Attachments"
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":12724
               Key             =   ""
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":12FFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":138D8
               Key             =   "Design Center Options"
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":141B2
               Key             =   "Global Options"
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":14A8C
               Key             =   "Models and Options"
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":15366
               Key             =   "Generate"
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FDBGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private mFormCaption As String
Private mTableName As String
Private mSQLStatement As String
Private mReadOnly As Boolean
Private mDirty As Boolean
Private mLockedColumns As String
Private mKeyColumns    As String
Private mAllowAddDelete As Boolean

Private Sub Form_Load()
    Call SetToolbarIcons(Me.Toolbar, Me.LargeIcons)
    Toolbar.Buttons("New").Visible = mAllowAddDelete And Not mReadOnly
    Call IniGetForm(Me, , mFormCaption)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not SaveData(True) Then
        Cancel = True
        Exit Sub
    End If
    Call IniPutForm(Me, , mFormCaption)
End Sub

Private Sub Form_Resize()
On Error Resume Next
    gData.Move 0, Toolbar.Height, Me.ScaleWidth, Me.ScaleHeight - Toolbar.Height
End Sub

Private Sub gData_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    gData.ColSel = gData.Col
End Sub

Private Sub gData_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    If Shift = 0 And gData.MouseRow > 0 Then
        gData.Row = gData.MouseRow
    End If
End Sub

Private Sub gData_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    Select Case True
        Case KeyCode = vbKeyDelete
            If Not mLockedColumns Like "*," & gData.ColKey(gData.Col) & ",*" Then
                gData.Text = ""
                mDirty = True
            End If
            
        Case KeyCode = vbKeyS And Shift = vbCtrlMask
            Call Toolbar_ButtonClick(Toolbar.Buttons("Save"))
            
    End Select
End Sub

Private Sub gData_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim i As Long
    
    With gData
        
        'you have to do this to catch mouse clicks into checkboxes. stupid.
        If Not IsBetween(Row, Min(.Row, .RowSel), Max(.Row, .RowSel)) Then
            .Row = Row
        End If
    
        For i = Min(.Row, .RowSel) To Max(.Row, .RowSel)
            If gData.RowData(i) = "" Then gData.RowData(i) = "Update"
        Next
    End With
    mDirty = True
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "ExcelExport"
            If SaveData(False) Then
                Call HFApp.ExportData(mSQLStatement, "Export " & Me.Caption & " to")
            End If
            
        Case "Save"
            Call SaveData(False)
            
        Case "New"
            Call gData.AddItem("")
            gData.Row = gData.Rows - 1
            gData.RowData(gData.Rows - 1) = "Insert"
            gData.SetFocus
            mDirty = True
            
    End Select
End Sub


Public Sub ShowForm(FormCaption As String, SQLStatement As String, TableName As String, KeyColumns As String, Optional ReadOnly As Boolean = False, Optional HiddenColumns As String, Optional LockedColumns As String, Optional AllowAddDelete As Boolean = True)
On Error Resume Next
    Dim i As Long
    Dim r As Long
    Dim c As Long
    Dim rs As Recordset
    
    
    mFormCaption = FormCaption
    mSQLStatement = SQLStatement
    mReadOnly = ReadOnly
    mLockedColumns = "," & LockedColumns & ","
    mTableName = TableName
    mKeyColumns = KeyColumns
    mAllowAddDelete = AllowAddDelete
        
    With gData
        .Editable = IIf(mReadOnly, flexEDNone, flexEDKbdMouse)
        Set rs = HFApp.SqlExec("" & SQLStatement, dbHomeFront)
        
        Set .DataSource = rs
        .FixedRows = 1
        
        'column translations
        c = .ColIndex("CostType")
        If c <> -1 Then
            .ColComboList(c) = "#1;labour|#2;material|#3;subcontract|#4;equipment|#5;overhead|#6;other"
        End If
        
        'setup formatting
        For c = 0 To rs.fields.Count - 1
            
            'hidden columns
            If "," & HiddenColumns & "," Like "*," & .ColKey(c + 1) & ",*" Then
                .ColHidden(c + 1) = True
            End If
            
            'format columns so excel doesnt break them and so we know how to write them back to db
            Select Case rs.fields(c).Type
                'boolean
                Case adBoolean
                    .ColDataType(c + 1) = flexDTBoolean
                
                'date/times
                Case adDate, adDBDate, adFileTime, adDBTime, adDBTimeStamp
                    .ColDataType(c + 1) = flexDTDate
                'numeric
                Case adVarNumeric, adVarWChar, adCurrency, adDecimal, adBinary, adBoolean, adDouble, adWChar, adInteger, adBigInt, adNumeric, adSingle, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                    .ColDataType(c + 1) = flexDTDouble
                'text
                Case Else
                    .ColDataType(c + 1) = flexDTString
            End Select
            
        Next
        
        'preserve key column values incase the user changes them.
        For i = 1 To Parse(mKeyColumns)
            c = .ColIndex(Trim(Parse(mKeyColumns, i)))
            If c <> -1 Then
                For r = 1 To .Rows
                    'have to add a leading char as the grid blows up when data is a number.
                    .Cell(flexcpData, r, c) = "k" & .Cell(flexcpText, r, c)
                Next
            End If
        Next
        
    End With
    
    mDirty = False
    Me.Caption = FormCaption
    Me.Show vbModal
    
End Sub


Private Sub gData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim s As String
    With gData
    
        If mLockedColumns Like "*," & .ColKey(Col) & ",*" Then
            Cancel = True
            Exit Sub
        End If
        
        
        Select Case .ColKey(Col)
            Case "DebitAccount": s = "..."
            Case Else:           s = ""
        End Select
        .ComboList = s
        
    End With
End Sub

Private Sub gData_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    
    Dim c As String
    Dim s As String
    Dim r As Long
        
    With gData
        Select Case .ColKey(Col)
            Case "DebitAccount":    c = "Accounts":           s = "SELECT Account,Description FROM GLAccounts WHERE divisionid=" & DbQuote(Num, HFApp.DivisionID)
        End Select
        
        If s <> "" Then
            If FPickList.Choose(HFApp.Databases(dbHomeFront), c, s, gData.TextMatrix(Row, Col)) Then
                For r = Min(Row, .RowSel) To Max(Row, .RowSel)
                    .Cell(flexcpText, r, Col, r, Col) = FPickList.SelectedItem(1)
                    If .RowData(r) = "" Then .RowData(r) = "Update"
                Next
            End If
            mDirty = True
        End If
    End With
    
End Sub



Private Function SaveData(prompt As Boolean) As Boolean
On Error GoTo eh
    
    Dim s As String
    Dim ic As String 'insert columns clause
    Dim iv As String 'insert values clause
    Dim r As Long
    Dim c As Long
    Dim whereclause As String
    Dim newkey As String
    
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
    
    Screen.MousePointer = vbHourglass
    With gData
    For r = 1 To .Rows - 1
    Select Case .RowData(r)
    Case "Insert"
        
        'get new key values for error handler dupkey message and enforce key columns are entered
        newkey = ""
        For c = 1 To Parse(mKeyColumns)
            s = Trim(Parse(mKeyColumns, c))
            Select Case True
                Case s = ""
                Case .ColKey(c) = "DivisionID":        newkey = newkey & " and [" & s & "]=" & DbQuote(Num, HFApp.DivisionID)
                Case .ColDataType(c) = flexDTBoolean:  newkey = newkey & " and [" & s & "]=" & DbQuote(Bit, .Cell(flexcpChecked, r, .ColIndex(s)) = flexChecked)
                Case .ColDataType(c) = flexDTDate:     newkey = newkey & " and [" & s & "]=" & DbQuote(Date, .Cell(flexcpText, r, .ColIndex(s)))
                Case .ColDataType(c) = flexDTDouble:   newkey = newkey & " and [" & s & "]=" & DbQuote(Num, .Cell(flexcpText, r, .ColIndex(s)))
                Case .ColDataType(c) = flexDTString:   newkey = newkey & " and [" & s & "]=" & DbQuote(Str, .Cell(flexcpText, r, .ColIndex(s)))
            End Select
            If s <> "" And s <> "divisionid" Then
                If .Cell(flexcpText, r, .ColIndex(s)) = "" Then
                    Call MsgBox(s & " cannot be blank.", vbExclamation + vbOK, App.ProductName)
                    SaveData = False
                    Screen.MousePointer = vbNormal
                    Exit Function
                End If
            End If
        Next
        If newkey <> "" Then newkey = Mid(newkey, 6)
        
        
        'build insert statement
        s = ""
        ic = ""
        iv = ""
        For c = 1 To .Cols - 1
            ic = ic & ",[" & Trim(.ColKey(c)) & "]"
            Select Case True
                Case .ColKey(c) = "DivisionID":          iv = iv & "," & DbQuote(Num, HFApp.DivisionID)
                Case .ColDataType(c) = flexDTBoolean:    iv = iv & "," & DbQuote(Bit, .Cell(flexcpChecked, r, c) = flexChecked)
                Case .ColDataType(c) = flexDTDate:       iv = iv & "," & DbQuote(Date, .TextMatrix(r, c))
                Case .ColDataType(c) = flexDTDouble:     iv = iv & "," & DbQuote(Num, .TextMatrix(r, c))
                Case Else:                               iv = iv & "," & DbQuote(Str, .TextMatrix(r, c))
            End Select
        Next
        s = "insert " & mTableName & "(" & Mid(ic, 2) & ") values(" & Mid(iv, 2) & ")"
        Call HFApp.SqlExec(s)
        .RowData(r) = ""
    
    
    Case "Update"
                
        'figure out where clause and get new key values for error handler dupkey message
        whereclause = ""
        newkey = ""
        For c = 1 To Parse(mKeyColumns)
            s = Trim(Parse(mKeyColumns, c))
            Select Case True
                Case s = ""
                Case .ColKey(c) = "DivisionID"
                    whereclause = whereclause & " and [" & s & "]=" & DbQuote(Num, HFApp.DivisionID)
                    newkey = newkey & " and [" & s & "]=" & DbQuote(Num, HFApp.DivisionID)
                
                Case .ColDataType(c) = flexDTBoolean
                    whereclause = whereclause & " and [" & s & "]=" & DbQuote(Bit, .Cell(flexcpChecked, r, .ColIndex(s)) = flexChecked)
                    newkey = newkey & " and [" & s & "]=" & DbQuote(Bit, .Cell(flexcpChecked, r, .ColIndex(s)) = flexChecked)
                
                Case .ColDataType(c) = flexDTDate
                    whereclause = whereclause & " and [" & s & "]=" & DbQuote(Date, Mid(.Cell(flexcpData, r, .ColIndex(s)), 2))
                    newkey = newkey & " and [" & s & "]=" & DbQuote(Date, .Cell(flexcpText, r, .ColIndex(s)))
                    
                Case .ColDataType(c) = flexDTDouble
                    whereclause = whereclause & " and [" & s & "]=" & DbQuote(Num, Mid(.Cell(flexcpData, r, .ColIndex(s)), 2))
                    newkey = newkey & " and [" & s & "]=" & DbQuote(Num, .Cell(flexcpText, r, .ColIndex(s)))
                    
                Case .ColDataType(c) = flexDTString
                    whereclause = whereclause & " and [" & s & "]=" & DbQuote(Str, Mid(.Cell(flexcpData, r, .ColIndex(s)), 2))
                    newkey = newkey & " and [" & s & "]=" & DbQuote(Str, .Cell(flexcpText, r, .ColIndex(s)))
                    
            End Select
        Next
        whereclause = Mid(whereclause, 6)
        If whereclause <> "" Then
            whereclause = " where " & whereclause
            newkey = Mid(newkey, 6)
        End If
        
        'build update statement
        s = ""
        For c = 1 To .Cols - 1
            Select Case True
                Case .ColDataType(c) = flexDTBoolean:    s = s & ",[" & Trim(.ColKey(c)) & "]=" & DbQuote(Bit, .Cell(flexcpChecked, r, c) = flexChecked) & vbCrLf
                Case .ColDataType(c) = flexDTDate:       s = s & ",[" & Trim(.ColKey(c)) & "]=" & DbQuote(Date, .TextMatrix(r, c)) & vbCrLf
                Case .ColDataType(c) = flexDTDouble:     s = s & ",[" & Trim(.ColKey(c)) & "]=" & DbQuote(Num, .TextMatrix(r, c)) & vbCrLf
                Case .ColDataType(c) = flexDTString:     s = s & ",[" & Trim(.ColKey(c)) & "]=" & DbQuote(Str, .TextMatrix(r, c)) & vbCrLf
            End Select
        Next
        s = "update " & mTableName & vbCrLf & " set " & Mid(s, 2) & whereclause
        Call HFApp.SqlExec(s)
        .RowData(r) = ""
    
    End Select
    Next
    End With
    Screen.MousePointer = vbNormal
    
    mDirty = False
    SaveData = True
    
Exit Function
eh:
    Screen.MousePointer = vbNormal
    If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        MsgBox "Unable to save duplicate record with key" & vbCrLf & vbCrLf & newkey, vbInformation, App.ProductName
    Else
        Call errHandler("SaveData", s)
    End If
End Function



