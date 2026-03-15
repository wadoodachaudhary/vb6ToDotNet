VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FDBGrid 
   ClientHeight    =   3090
   ClientLeft      =   1995
   ClientTop       =   4020
   ClientWidth     =   10305
   Icon            =   "FDBGrid.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   10305
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   1935
      Left            =   0
      TabIndex        =   1
      Top             =   750
      Width           =   5085
      _cx             =   8969
      _cy             =   3413
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
      SheetBorder     =   -2147483632
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   5
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
      Height          =   600
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   1058
      ButtonWidth     =   1085
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save   "
            Key             =   "Save"
            Object.ToolTipText     =   "Ctrl+S"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Key             =   "Delete"
            Object.ToolTipText     =   "Ctrl+Delete"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Export"
            Key             =   "ExcelExport"
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
               Picture         =   "FDBGrid.frx":000C
               Key             =   "EditAssembly"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":08E6
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":11C0
               Key             =   "ExcelImport"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":1A9A
               Key             =   "ExcelExport"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":2374
               Key             =   "Publish"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":2C4E
               Key             =   "Forecast"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":3528
               Key             =   "ViewPOs"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":3E02
               Key             =   "ViewBudgets"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":46DC
               Key             =   "Open"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":4FB6
               Key             =   "Preview"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":5890
               Key             =   "Send"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":616A
               Key             =   "TakeoffOneTime"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":6A44
               Key             =   "Estimate"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":731E
               Key             =   "TakeoffAssembly"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":7BF8
               Key             =   "NewAssembly"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":84D2
               Key             =   "New"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":8DAC
               Key             =   "TakeoffItem"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":9686
               Key             =   "TakeoffCustom"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":9F60
               Key             =   "Save"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":A83A
               Key             =   "SaveAs"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":B114
               Key             =   "Delete"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":B9EE
               Key             =   "RePrice"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":C2C8
               Key             =   "PricebookSearch"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":CBA2
               Key             =   "Pricebook"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":D47C
               Key             =   "PricebookEdit"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":DD56
               Key             =   "PricebookExport"
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":E630
               Key             =   "PricebookImport"
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":EF0A
               Key             =   "PricebookNew"
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":F7E4
               Key             =   "View"
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":100BE
               Key             =   "Vendor1"
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":10998
               Key             =   "Vendor"
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":11272
               Key             =   "Add"
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":11B4C
               Key             =   "Attachments"
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":11E66
               Key             =   ""
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":12740
               Key             =   ""
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":1301A
               Key             =   "Design Center Options"
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":138F4
               Key             =   "Global Options"
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":141CE
               Key             =   "Models and Options"
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDBGrid.frx":14AA8
               Key             =   "Generate"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupSub 
         Caption         =   "Cut"
         Index           =   1
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuPopupSub 
         Caption         =   "Copy"
         Index           =   2
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPopupSub 
         Caption         =   "Paste"
         Index           =   3
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuPopupSub 
         Caption         =   "Delete"
         Index           =   4
         Shortcut        =   {DEL}
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
Private mKeyColumns As String
Private mAddRecords As Boolean
Private mAllowExport As Boolean
Private mFields As adodb.fields
Private mHasIdentityColumn As Boolean


Const mcCUT = 1
Const mcCOPY = 2
Const mcPASTE = 3
Const mcDELETE = 4


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case True
        Case KeyCode = vbKeyDelete And Shift = vbCtrlMask:      Call Toolbar_ButtonClick(Toolbar.Buttons("Delete"))
        Case KeyCode = vbKeyS And Shift = vbCtrlMask:           Call Toolbar_ButtonClick(Toolbar.Buttons("Save"))
    End Select
End Sub

Private Sub Form_Load()
    Me.Caption = mFormCaption
    Call IniGetForm(Me, , mFormCaption)
    Call SetToolbarIcons(Me.Toolbar, Me.LargeIcons)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not SaveData(True) Then
        Cancel = True
        Exit Sub
    End If
    Call IniPutForm(Me, , Me.Caption)
End Sub

Private Sub Form_Resize()
On Error Resume Next
    gData.Move 0, Toolbar.Height, Me.ScaleWidth, Me.ScaleHeight - Toolbar.Height
End Sub

Private Sub gData_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    gData.ColSel = gData.Col
End Sub

Private Sub gData_AfterSort(ByVal Col As Long, Order As Integer)
    If mAddRecords Then gData.AddItem ""
End Sub

Private Sub gData_BeforeSort(ByVal Col As Long, Order As Integer)
    If mAddRecords Then gData.RemoveItem (gData.Rows - 1)
End Sub

Private Sub gData_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Dim r As Long
    With gData
    Select Case True
    
        Case KeyCode = vbKeyC And Shift = vbCtrlMask
            Call Clipboard.Clear
            Call Clipboard.SetText(.TextMatrix(.Row, .Col))
    
        Case KeyCode = vbKeyV And Shift = vbCtrlMask
            If mLockedColumns Like "*," & .ColKey(.Col) & ",*" Then Exit Sub
            For r = Min(.Row, .RowSel) To Max(.Row, .RowSel)
                .TextMatrix(r, .Col) = Clipboard.GetText
                If .RowData(r) = "" Then .RowData(r) = "update"
            Next
            mDirty = True
    
        Case KeyCode = vbKeyDelete And Shift <> 0
            If Not mAddRecords Or .Row = .Rows - 1 Then Exit Sub
            If .Row < 1 Then Exit Sub
            If MsgBox("Are you sure you want to delete this record?" & vbCrLf & vbCrLf & "You will not be able to undo this action.", vbYesNo + vbExclamation, App.ProductName) = vbNo Then Exit Sub
            
            .RowHidden(.Row) = True
            .RowData(.Row) = "delete"
            mDirty = True
        
        Case KeyCode = vbKeyDelete
            If mLockedColumns Like "*," & .ColKey(.Col) & ",*" Then Exit Sub
            If mReadOnly Then Exit Sub
            
            .Text = ""
            If .RowData(.Row) = "" Then .RowData(.Row) = "update"
            mDirty = True
            Select Case .ColKey(.Col)
                Case "crewtype":    .Cell(flexcpText, .Row, .ColIndex("crewtypeid"), .RowSel, .ColIndex("crewtypeid")) = ""
                Case "technician":  .Cell(flexcpText, .Row, .ColIndex("techid"), .RowSel, .ColIndex("techid")) = ""
            End Select
            
        Case KeyCode = vbKeyS And Shift = vbCtrlMask
            Call Toolbar_ButtonClick(Toolbar.Buttons("Save"))
            
    End Select
    End With
End Sub

Private Sub gData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu mnuPopup
End Sub

Private Sub gData_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    With gData
        If .ColKey(.MouseCol) = "Sage300GLPrefixLength" Then
            .ToolTipText = "Number of digits of the gl account that make up the company number"
        Else
            .ToolTipText = ""
        End If
    End With
End Sub

Private Sub gData_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim r As Long
    
    With gData
    
        'you have to do this to catch mouse clicks into checkboxes. stupid.
        If Not IsBetween(Row, Min(.Row, .RowSel), Max(.Row, .RowSel)) Then
            .Row = Row
        End If
        
        
        If mAddRecords And Row = .Rows - 1 Then
            .RowData(.Rows - 1) = "insert"
            If .ColIndex("WalletPartyID") <> -1 Then
                .TextMatrix(.Rows - 1, .ColIndex("WalletPartyID")) = CreateGUID()
            End If
            .AddItem ""
        End If
        
        
        Select Case True
            'int col
            Case mFields(.ColKey(Col)).Type = adInteger
                .EditText = Int(Val(.EditText))
            
            'time col
            Case mFields(.ColKey(Col)).Type = adVarWChar '<-- yup thats what it is.
               If .EditText <> "" Then
                    If IsNumeric(.EditText) And Val(.EditText) < 24 Then
                        .EditText = Format(CDate(Val(.EditText) / 24), "Medium Time")
                        .TextMatrix(Row, Col) = .EditText
                    Else
                        If IsDate(.EditText) Then
                            .EditText = Format(.EditText, "Medium Time")
                            .TextMatrix(Row, Col) = .EditText
                        Else
                            MsgBox "invalid time entered", vbInformation, App.ProductName
                            Cancel = True
                        End If
                    End If
                End If
                
            'date col
            Case .ColDataType(Col) = flexDTDate
               If .EditText <> "" Then
                    If Not IsDate(.EditText) Then
                        MsgBox "invalid time entered", vbInformation, App.ProductName
                        Cancel = True
                    End If
                End If
                
        End Select
        
        For r = Min(.Row, .RowSel) To Max(.Row, .RowSel)
            If .RowData(r) = "" Then .RowData(r) = "Update"
        Next
        mDirty = True
        
    End With
End Sub

Private Sub mnuPopupSub_Click(Index As Integer)
    Select Case Index
        Case mcCUT:      Call gData_KeyDown(vbKeyDelete, 0)
        Case mcCOPY:     Call gData_KeyDown(vbKeyC, vbCtrlMask)
        Case mcPASTE:    Call gData_KeyDown(vbKeyV, vbCtrlMask)
        Case mcDELETE:   Call gData_KeyDown(vbKeyDelete, vbCtrlMask)
    End Select
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "ExcelExport"
            If SaveData(False) Then
                Call FDataExport.ExportData(mSQLStatement)
            End If
            
        Case "Delete"
            Call gData_KeyDown(vbKeyDelete, vbCtrlMask)
            
        Case "Save"
            Call SaveData(False)
    
    End Select
End Sub


Public Sub ShowForm(FormCaption As String, SQLStatement As String, TableName As String, KeyColumns As String, _
                    Optional ReadOnly As Boolean = False, Optional HiddenColumns As String, Optional LockedColumns As String, _
                    Optional AddRecords As Boolean = True, Optional AllowExport As Boolean = True)
On Error Resume Next
    Dim i As Long
    Dim r As Long
    Dim c As Long
    Dim rs As Recordset
    
    mSQLStatement = SQLStatement
    mReadOnly = ReadOnly
    mLockedColumns = "," & LockedColumns & ","
    mTableName = TableName
    mKeyColumns = KeyColumns
    mAddRecords = AddRecords
    mFormCaption = FormCaption
    mAllowExport = AllowExport
    
    Toolbar.Buttons.item("Delete").Visible = mAddRecords
    Toolbar.Buttons.item("ExcelExport").Visible = mAllowExport
    
    mnuPopupSub(mcCUT).Enabled = Not mReadOnly
    mnuPopupSub(mcCOPY).Enabled = True
    mnuPopupSub(mcPASTE).Enabled = Not mReadOnly
    mnuPopupSub(mcDELETE).Enabled = mAddRecords
    
    
    With gData
        .Editable = IIf(mReadOnly, flexEDNone, flexEDKbdMouse)
        
        'Set rs = HFApp.SqlExec("" & SQLStatement, dbHomefront)
        Set rs = New Recordset
        Call rs.Open(SQLStatement, HFApp.Databases(dbHomeFront), adOpenDynamic)

        
        Set mFields = rs.fields
        
        Set .DataSource = rs
        .FixedRows = 1
        
        If mAddRecords Then .AddItem ""
        
        
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
      

            
            
            
            'stash column size
            'stash basetablename
            'stash identity...
            .ColData(c + 1) = rs.fields(c).DefinedSize & Chr(1) & _
                              rs.fields(c).Properties("BASETABLENAME") & Chr(1) & _
                              rs.fields(c).Properties("ISAUTOINCREMENT") & Chr(1) & _
                              rs.fields(c).Properties("BASECOLUMNNAME")
                              
            If rs.fields(c).Properties("ISAUTOINCREMENT") Then mHasIdentityColumn = True
            
            'format columns so excel doesnt break them and so we know how to write them back to db
            Select Case True
                'boolean
                Case rs.fields(c).Type = adBoolean
                    .ColDataType(c + 1) = flexDTBoolean
                
                
                'time only
                Case rs.fields(c).Type = adVarWChar '<-- yup thats what it is.
                    .ColFormat(c + 1) = "Medium Time"
                    For i = 1 To .Rows - 2
                        .TextMatrix(i, c + 1) = Format(Parse(.TextMatrix(i, c + 1), 1, "."), "Medium Time")
                    Next
                
                'date/times
                Case IsIn(rs.fields(c).Type, adDate, adDBDate, adFileTime, adDBTime, adDBTimeStamp)
                    .ColFormat(c + 1) = HFApp.Options(DateFormat)
                    .ColDataType(c + 1) = flexDTDate
                    
                'numeric
                Case IsIn(rs.fields(c).Type, adVarNumeric, adVarWChar, adCurrency, adDecimal, adBinary, adBoolean, adDouble, adWChar, adInteger, adBigInt, adNumeric, adSingle, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt)
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
        
        
        
        
        'do color translations
        i = .ColIndex("color")
        If i <> -1 Then
            .FixedAlignment(i) = flexAlignCenterCenter
            For r = 1 To .Rows
                .Cell(flexcpBackColor, r, i) = Val(.Cell(flexcpText, r, i))
                .Cell(flexcpText, r, i) = ""
            Next
        End If
        
        
        
        
        
        
    End With
    
    mDirty = False
    Me.Show vbModal
    
End Sub


Private Sub gData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim s As String
    With gData
    
        If mLockedColumns Like "*," & .ColKey(Col) & ",*" Then
            Cancel = True
            Exit Sub
        End If
        
        
        s = ""
        Select Case True
            Case .ColDataType(Col) = flexDTDate
                s = "|..."
                
            Case IsIn(.ColKey(Col), "DebitAccount", "Group_Code")
                s = "..."
            
            Case IsIn(.ColKey(Col), "Sage300GLPrefixes")
                If .ValueMatrix(Row, .ColIndex("Sage300glPrefixlength")) > 0 Then
                    s = "..."
                Else
                    Cancel = True
                    Exit Sub
                End If
            
        
        End Select
        
        
        
        .ComboList = s
        .EditMaxLength = Val(Parse(.ColData(Col), 1, Chr(1)))
        
    End With
End Sub

Private Sub gData_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    
    Dim c As String
    Dim s As String
    Dim r As Long
        
    With gData
        Select Case True
        
        Case .ColDataType(Col) = flexDTDate
            r = .Row
            Call DCalendar.Popup(gData, .RowPos(r) + .RowHeight(r), .colPos(.Col))
            If mAddRecords And r = .Rows - 1 Then
                .RowData(r) = "insert"
                .AddItem ""
            End If
        
        Case .ColKey(Col) = "Group_Code"
            If FPickList.Choose(HFApp.Databases(dbHomeFront), "Group", "select Major_Group,Description from tblmajorgroups", gData.TextMatrix(Row, Col)) Then
                For r = Min(Row, .RowSel) To Max(Row, .RowSel)
                    .Cell(flexcpText, r, Col, r, Col) = FPickList.SelectedItem("major_group")
                    If mAddRecords And r = .Rows - 1 Then
                        .RowData(r) = "insert"
                        .AddItem ""
                    End If
                    If .RowData(r) = "" Then .RowData(r) = "Update"
                Next
                
            End If
            mDirty = True
        
        
        Case .ColKey(Col) = "DebitAccount"
            If FPickList.Choose(HFApp.Databases(dbHomeFront), "GL Account", "select Account, Description from GLAccounts where divisionid=" & DbQuote(Num, HFApp.DivisionID), gData.TextMatrix(Row, Col)) Then
                For r = Min(Row, .RowSel) To Max(Row, .RowSel)
                    .Cell(flexcpText, r, Col, r, Col) = FPickList.SelectedItem("account")
                    If mAddRecords And r = .Rows - 1 Then
                        .RowData(r) = "insert"
                        .AddItem ""
                    End If
                    If .RowData(r) = "" Then .RowData(r) = "Update"
                Next
                
            End If
            mDirty = True
            
            
        Case .ColKey(Col) = "Sage300GLPrefixes"
            Dim f As New FDBGrid
            s = "select DataSourceID,GLPrefix,Description,WalletCompanyID from Wallet_Sage300GLPrefixes where datasourceid=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("WalletPartyID"))) & " and len(glprefix)=" & .TextMatrix(Row, .ColIndex("Sage300glPrefixlength"))
            Call f.ShowForm("GLPrefixes", s, "Wallet_Sage300GLPrefixes", "DataSourceID,GLPrefix", False, "DataSourceID", "GLPrefix,Description", False)

        
        End Select
    End With
    
End Sub



Private Function SaveData(prompt As Boolean) As Boolean
On Error GoTo eh
    Dim s As String
    Dim r As Long
    Dim c As Long
    Dim i As Long
    Dim ValuesClause As String
    Dim WhereClause As String
    Dim newkey As String
    Dim basecolname As String
    
    
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
    
    'do deletes first---------------------------------------------------------------------------
    With gData
    For r = .Rows - 1 To 1 Step -1
    If .RowData(r) = "delete" Then
        WhereClause = ""
        For c = 1 To Parse(mKeyColumns)
            s = Trim(Parse(mKeyColumns, c))
            Select Case True
                Case s = ""
                Case .ColDataType(c) = flexDTBoolean:  WhereClause = WhereClause & " and [" & s & "]=" & DbQuote(Bit, .Cell(flexcpChecked, r, .ColIndex(s)) = flexChecked)
                Case .ColDataType(c) = flexDTDate:     WhereClause = WhereClause & " and [" & s & "]=" & DbQuote(Date, Mid(.Cell(flexcpData, r, .ColIndex(s)), 2))
                Case .ColDataType(c) = flexDTDouble:   WhereClause = WhereClause & " and [" & s & "]=" & DbQuote(Num, Mid(.Cell(flexcpData, r, .ColIndex(s)), 2))
                Case .ColDataType(c) = flexDTString:   WhereClause = WhereClause & " and [" & s & "]=" & DbQuote(Str, Mid(.Cell(flexcpData, r, .ColIndex(s)), 2))
            End Select
        Next
        WhereClause = Mid(WhereClause, 6)
        If WhereClause <> "" Then
            s = "delete " & mTableName & " where " & WhereClause
            Call HFApp.SqlExec(s, dbHomeFront)
            Call .RemoveItem(r)
        End If
    End If
    Next
    
    
    
    'do updates second---------------------------------------------------------------------------
    For r = 1 To .Rows - 1
    If .RowData(r) = "Update" Then
        'build where clause and get new keys for dupkey errors
        WhereClause = ""
        newkey = ""
        For c = 1 To Parse(mKeyColumns)
            s = Trim(Parse(mKeyColumns, c))
            Select Case True
                Case s = ""
                Case .ColDataType(c) = flexDTBoolean
                    WhereClause = WhereClause & " and [" & s & "]=" & DbQuote(Bit, .Cell(flexcpChecked, r, .ColIndex(s)) = flexChecked)
                    newkey = newkey & " and [" & s & "]=" & DbQuote(Bit, .Cell(flexcpChecked, r, .ColIndex(s)) = flexChecked)
                Case .ColDataType(c) = flexDTDate
                    WhereClause = WhereClause & " and [" & s & "]=" & DbQuote(Date, Mid(.Cell(flexcpData, r, .ColIndex(s)), 2))
                    newkey = newkey & " and [" & s & "]=" & DbQuote(Date, .Cell(flexcpText, r, .ColIndex(s)))
                Case .ColDataType(c) = flexDTDouble
                    WhereClause = WhereClause & " and [" & s & "]=" & DbQuote(Num, Mid(.Cell(flexcpData, r, .ColIndex(s)), 2))
                    newkey = newkey & " and [" & s & "]=" & DbQuote(Num, .Cell(flexcpText, r, .ColIndex(s)))
                Case .ColDataType(c) = flexDTString
                    WhereClause = WhereClause & " and [" & s & "]=" & DbQuote(Str, Mid(.Cell(flexcpData, r, .ColIndex(s)), 2))
                    newkey = newkey & " and [" & s & "]=" & DbQuote(Str, .Cell(flexcpText, r, .ColIndex(s)))
            End Select
        Next
        WhereClause = Mid(WhereClause, 6)
        newkey = Mid(newkey, 6)
        If WhereClause <> "" Then
            'build update clause
            ValuesClause = ""
            For c = 1 To .Cols - 1
            'column is on the update table and is not autoincrement
            basecolname = Parse(.ColData(c), 4, Chr(1))
            If Parse(.ColData(c), 2, Chr(1)) = mTableName And Parse(.ColData(c), 3, Chr(1)) <> "True" Then
                Select Case True
                    Case .ColKey(c) = "color":               ValuesClause = ValuesClause & ",[" & Trim(basecolname) & "]=" & DbQuote(Num, .Cell(flexcpBackColor, r, c))
                    Case .ColDataType(c) = flexDTBoolean:    ValuesClause = ValuesClause & ",[" & Trim(basecolname) & "]=" & DbQuote(Bit, .Cell(flexcpChecked, r, c) = flexChecked)
                    
                    Case mFields(c - 1).Type = adVarWChar:
                        ValuesClause = ValuesClause & ",[" & Trim(basecolname) & "]=" & DbQuote(Time, .TextMatrix(r, c), True)
                    
                    Case .ColDataType(c) = flexDTDate:       ValuesClause = ValuesClause & ",[" & Trim(basecolname) & "]=" & DbQuote(Date, .TextMatrix(r, c), True)
                    Case .ColDataType(c) = flexDTDouble:     ValuesClause = ValuesClause & ",[" & Trim(basecolname) & "]=" & DbQuote(Num, .TextMatrix(r, c), True)
                    Case .ColDataType(c) = flexDTString:     ValuesClause = ValuesClause & ",[" & Trim(basecolname) & "]=" & DbQuote(Str, .TextMatrix(r, c), True)
                End Select
            End If
            Next
            s = "update " & mTableName & " set " & Mid(ValuesClause, 2) & " where " & WhereClause
            Call HFApp.SqlExec(s, dbHomeFront)
            .RowData(r) = ""
            
            'preserve key column values incase the user changes them.
            For i = 1 To Parse(mKeyColumns)
                c = .ColIndex(Trim(Parse(mKeyColumns, i)))
                If c <> -1 Then
                    'have to add a leading char as the grid blows up when data is a number.
                    .Cell(flexcpData, r, c) = "k" & .Cell(flexcpText, r, c)
                End If
            Next
        End If
    End If
    Next
    
    
    'do inserts third---------------------------------------------------------------------------
    For r = 1 To .Rows - 1
    If .RowData(r) = "insert" Then
        'build get new keys for dupkey errors
        newkey = ""
        For c = 1 To Parse(mKeyColumns)
            s = Trim(Parse(mKeyColumns, c))
            Select Case True
                Case s = ""
                Case s = "DivisionID":                 newkey = newkey & " and [" & s & "]=" & DbQuote(Num, HFApp.DivisionID)
                Case .ColDataType(c) = flexDTBoolean:  newkey = newkey & " and [" & s & "]=" & DbQuote(Bit, .Cell(flexcpChecked, r, .ColIndex(s)) = flexChecked)
                Case .ColDataType(c) = flexDTDate:     newkey = newkey & " and [" & s & "]=" & DbQuote(Date, .Cell(flexcpText, r, .ColIndex(s)), True)
                Case .ColDataType(c) = flexDTDouble:   newkey = newkey & " and [" & s & "]=" & DbQuote(Num, .Cell(flexcpText, r, .ColIndex(s)), True)
                Case .ColDataType(c) = flexDTString:   newkey = newkey & " and [" & s & "]=" & DbQuote(Str, .Cell(flexcpText, r, .ColIndex(s)), True)
                Case .ColDataType(c) = 130:            newkey = newkey & " and [" & s & "]=" & DbQuote(Time, .Cell(flexcpText, r, .ColIndex(s)), True)
            End Select
        Next
        newkey = Mid(newkey, 6)
        
        'build insert clause
        ValuesClause = "insert into " & mTableName & "("
        For c = 1 To .Cols - 1
        If Parse(.ColData(c), 2, Chr(1)) = mTableName And Parse(.ColData(c), 3, Chr(1)) <> "True" Then
            basecolname = Parse(.ColData(c), 4, Chr(1))
            ValuesClause = ValuesClause & "[" & Trim(basecolname) & "],"
        End If
        Next
        ValuesClause = Mid(ValuesClause, 1, Len(ValuesClause) - 1) & ") values("
        For c = 1 To .Cols - 1
        'column is on the update table and is not autoincrement
        If Parse(.ColData(c), 2, Chr(1)) = mTableName And Parse(.ColData(c), 3, Chr(1)) <> "True" Then
            Select Case True
                Case .ColKey(c) = "DivisionID":          ValuesClause = ValuesClause & DbQuote(Num, HFApp.DivisionID) & ",":    .TextMatrix(r, c) = HFApp.DivisionID
                Case .ColDataType(c) = flexDTBoolean:    ValuesClause = ValuesClause & DbQuote(Bit, .Cell(flexcpChecked, r, c) = flexChecked) & ","
                Case .ColDataType(c) = flexDTDate:       ValuesClause = ValuesClause & DbQuote(Date, .TextMatrix(r, c), True) & ","
                Case .ColDataType(c) = flexDTDouble:     ValuesClause = ValuesClause & DbQuote(Num, .TextMatrix(r, c), True) & ","
                Case .ColDataType(c) = flexDTString:     ValuesClause = ValuesClause & DbQuote(Str, .TextMatrix(r, c), True) & ","
                Case .ColDataType(c) = 130:              ValuesClause = ValuesClause & DbQuote(Time, .TextMatrix(r, c), True) & ","
            End Select
        End If
        Next
        s = Mid(ValuesClause, 1, Len(ValuesClause) - 1) & ")"
        Call HFApp.SqlExec(s, dbHomeFront)
        .RowData(r) = ""
        
        'get new identity value
        If Parse(mKeyColumns) = 1 Then
            c = .ColIndex(Trim(Parse(mKeyColumns, 1)))
            i = HFApp.SqlIdentity(mTableName) 'returns zero if table doesnt have identity
            If i > 0 Then .Cell(flexcpText, r, c) = i
        End If
        
        'preserve key column values incase the user changes them.
        For i = 1 To Parse(mKeyColumns)
            c = .ColIndex(Trim(Parse(mKeyColumns, i)))
            If c <> -1 Then
                'have to add a leading char as the grid blows up when data is a number.
                .Cell(flexcpData, r, c) = "k" & .Cell(flexcpText, r, c)
            End If
        Next
        
    End If
    Next
    
    End With
    Screen.MousePointer = vbNormal
    
    mDirty = False
    SaveData = True
    
Exit Function
eh:
    Screen.MousePointer = vbNormal
    Select Case True
    
    Case InStr(1, Err.Description, "duplicate", vbTextCompare)
        If mHasIdentityColumn Then
            MsgBox "Unable to save duplicate record." & vbCrLf & vbCrLf & Err.Description, vbInformation, App.ProductName
        Else
            MsgBox "Unable to save duplicate record with key" & vbCrLf & vbCrLf & newkey, vbInformation, App.ProductName
        End If
        
        
    Case InStr(1, Err.Description, "column does not allow nulls", vbTextCompare)
        MsgBox Parse(Err.Description, 2, "'") & " is required.", vbInformation, App.ProductName
    
    Case Else
        Call errHandler("SaveData", s)
    End Select
End Function
