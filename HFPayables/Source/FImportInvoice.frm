VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form FImportInvoice 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Payables Desk"
   ClientHeight    =   2790
   ClientLeft      =   14550
   ClientTop       =   2805
   ClientWidth     =   7035
   Icon            =   "FImportInvoice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   315
      Left            =   450
      TabIndex        =   24
      Top             =   1980
      Visible         =   0   'False
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.TextBox txtDescription 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1260
      MaxLength       =   30
      TabIndex        =   5
      Top             =   1410
      Width           =   2535
   End
   Begin VB.TextBox txtInvoice 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1260
      TabIndex        =   1
      Top             =   450
      Width           =   1695
   End
   Begin VB.TextBox txtVendor 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1260
      TabIndex        =   0
      Top             =   210
      Width           =   1695
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6570
      Picture         =   "FImportInvoice.frx":000C
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Browse folders"
      Top             =   1650
      Width           =   250
   End
   Begin VB.TextBox txtFileName 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1260
      TabIndex        =   6
      Top             =   1650
      Width           =   5265
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Height          =   375
      Index           =   0
      Left            =   2490
      TabIndex        =   8
      Top             =   2340
      Width           =   1095
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Index           =   1
      Left            =   3660
      TabIndex        =   9
      Top             =   2340
      Width           =   1095
   End
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   7605
      Left            =   510
      TabIndex        =   10
      Top             =   2820
      Visible         =   0   'False
      Width           =   14475
      _cx             =   1975673788
      _cy             =   1975661670
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   0   'False
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
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
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
      ExplorerBar     =   3
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
   Begin TDBDate6Ctl.TDBDate dteReceived 
      Height          =   225
      Left            =   1260
      TabIndex        =   3
      Top             =   930
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   397
      Calendar        =   "FImportInvoice.frx":0596
      Caption         =   "FImportInvoice.frx":068B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FImportInvoice.frx":06F0
      Keys            =   "FImportInvoice.frx":070E
      Spin            =   "FImportInvoice.frx":077A
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd-mmm-yy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   -1
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "dd-mmm-yy"
      HighlightText   =   2
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__-___-__"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   3.66819414244724E-316
      CenturyMode     =   0
   End
   Begin TDBDate6Ctl.TDBDate dteInvoice 
      Height          =   225
      Left            =   1260
      TabIndex        =   2
      Top             =   690
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   397
      Calendar        =   "FImportInvoice.frx":07A2
      Caption         =   "FImportInvoice.frx":0897
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FImportInvoice.frx":08FC
      Keys            =   "FImportInvoice.frx":091A
      Spin            =   "FImportInvoice.frx":097E
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd-mmm-yy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   -1
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "dd-mmm-yy"
      HighlightText   =   2
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__-___-__"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   1.10537306944062E-317
      CenturyMode     =   0
   End
   Begin TDBDate6Ctl.TDBDate dteAccounting 
      Height          =   225
      Left            =   1260
      TabIndex        =   4
      Top             =   1170
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   397
      Calendar        =   "FImportInvoice.frx":09A6
      Caption         =   "FImportInvoice.frx":0A9B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FImportInvoice.frx":0B00
      Keys            =   "FImportInvoice.frx":0B1E
      Spin            =   "FImportInvoice.frx":0B8A
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd-mmm-yy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   -1
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "dd-mmm-yy"
      HighlightText   =   2
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__-___-__"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   1.10541259469229E-317
      CenturyMode     =   0
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Index           =   5
      Left            =   390
      TabIndex        =   23
      Top             =   1410
      Width           =   795
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company"
      Height          =   195
      Index           =   3
      Left            =   3150
      TabIndex        =   22
      Top             =   210
      Width           =   675
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   195
      Index           =   0
      Left            =   3240
      TabIndex        =   21
      Top             =   435
      Width           =   585
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Terms"
      Height          =   195
      Index           =   4
      Left            =   3390
      TabIndex        =   20
      Top             =   1020
      Width           =   435
   End
   Begin VB.Label lblTerms 
      BackColor       =   &H80000005&
      Height          =   465
      Left            =   3870
      TabIndex        =   19
      Top             =   1050
      UseMnemonic     =   0   'False
      Width           =   2955
   End
   Begin VB.Label lblAddress 
      BackColor       =   &H80000005&
      Height          =   615
      Left            =   3870
      TabIndex        =   18
      Top             =   435
      UseMnemonic     =   0   'False
      Width           =   2955
   End
   Begin VB.Label lblCompany 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3870
      TabIndex        =   17
      Top             =   210
      UseMnemonic     =   0   'False
      Width           =   2955
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File Name"
      Height          =   195
      Index           =   2
      Left            =   480
      TabIndex        =   16
      Top             =   1650
      Width           =   705
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice"
      Height          =   195
      Index           =   1
      Left            =   660
      TabIndex        =   15
      Top             =   450
      Width           =   525
   End
   Begin VB.Label lblVendor 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   675
      TabIndex        =   14
      Top             =   210
      Width           =   510
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inv Date"
      Height          =   195
      Index           =   8
      Left            =   570
      TabIndex        =   13
      Top             =   690
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accounting"
      Height          =   195
      Index           =   10
      Left            =   375
      TabIndex        =   12
      Top             =   1170
      Width           =   810
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Received"
      Height          =   195
      Index           =   7
      Left            =   495
      TabIndex        =   11
      Top             =   930
      Width           =   690
   End
End
Attribute VB_Name = "FImportInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FImportInvoice::"

Private mVendor  As String
Private mInvoice As String
Private mAccountingDate As Date
Private mVendorType As String
Private mVDefPhase        As String
Private mVDefCategory     As String
Private mVDefDebitAccount As String
Private mVDefTaxGroup     As String
Private mDiscDays    As Long
Private mDiscPercent As Double
Private mTermsDays   As Long
Private mTermsType   As Long
Private mWCBRate     As Double


Public Function ImportInvoice(Vendor As String, Invoice As String, AccountingDate) As Boolean
    mVendor = ""
    mInvoice = ""
    
    If IsNull(AccountingDate) Then
        mAccountingDate = VBA.Date
    Else
        mAccountingDate = AccountingDate
    End If
    
    Me.Show vbModal
    Vendor = mVendor
    Invoice = mInvoice
    ImportInvoice = mVendor <> ""
End Function

Private Sub cmdBrowse_Click()
    Dim s As String
    txtFileName.SetFocus
    s = txtFileName.Text
    If VBGetOpenFileName(s, , , , , True, "Excel Files (*.xls)|*.xls|Text Files (*.txt;*.csv)|*.txt;*.csv|All Files (*.*)|*.*", , , , , FMain.hWnd) Then
        txtFileName.Text = s
    End If
End Sub

Private Sub cmdNav_Click(Index As Integer)
    Select Case Index
        Case 0:
            If FileExt(txtFileName.Text) = "xls" Then
                If ImportVisaInvoice Then Unload Me
            Else
                If ImportProfitMasterSalesReport Then Unload Me
            End If
            
            
        Case Else: Unload Me
    End Select
End Sub

Private Sub dteInvoice_Change()
    If App.Options(AcctDateDefaultsToInvDate) Then dteAccounting.Value = dteInvoice.Value
    If App.Options(AcctDateDefaultsToRecDate) Then dteAccounting.Value = dteReceived.Value
End Sub

Private Sub dteReceived_Change()
    If App.Options(AcctDateDefaultsToInvDate) Then dteAccounting.Value = dteInvoice.Value
    If App.Options(AcctDateDefaultsToRecDate) Then dteAccounting.Value = dteReceived.Value
End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
    txtFileName.Text = IniGet(AppIni, "ImportInvoice", "FileName")
    txtVendor.Text = IniGet(AppIni, "ImportInvoice", "Vendor")
    txtDescription.Text = IniGet(AppIni, "ImportInvoice", "Description")
    
    dteInvoice.Value = VBA.Date
    dteReceived.Value = VBA.Date
    dteAccounting.Value = mAccountingDate
    If App.Options(AcctDateDefaultsToInvDate) Then dteAccounting.Value = dteInvoice.Value
    If App.Options(AcctDateDefaultsToRecDate) Then dteAccounting.Value = dteReceived.Value
    dteInvoice.Format = App.Options.Value(TimberlineDateFormat)
    dteInvoice.DisplayFormat = App.Options.Value(TimberlineDateFormat)
    dteReceived.Format = App.Options.Value(TimberlineDateFormat)
    dteReceived.DisplayFormat = App.Options.Value(TimberlineDateFormat)
    dteAccounting.Format = App.Options.Value(TimberlineDateFormat)
    dteAccounting.DisplayFormat = App.Options.Value(TimberlineDateFormat)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
    Call IniPut(AppIni, "ImportInvoice", "FileName", txtFileName.Text)
    Call IniPut(AppIni, "ImportInvoice", "Vendor", txtVendor.Text)
    Call IniPut(AppIni, "ImportInvoice", "Description", txtDescription.Text)
End Sub

Private Function TextMatrix(Row As Long, Optional Col As Long, Optional ColKey As String) As String
On Error Resume Next
    If ColKey = "" Then
        TextMatrix = gData.TextMatrix(Row, Col)
    Else
        TextMatrix = gData.TextMatrix(Row, gData.ColIndex(ColKey))
    End If
End Function






Private Sub txtDescription_GotFocus()
SelectAll txtDescription
End Sub

Private Sub txtFileName_GotFocus()
SelectAll txtFileName
End Sub

Private Sub txtInvoice_GotFocus()
SelectAll txtInvoice
End Sub

Private Sub txtInvoice_Validate(Cancel As Boolean)
    Dim s  As String
    Dim rs As Recordset
    
    If TimberlineAccounting Then
        'check TL new file
        s = ""
        s = s & "select * " & vbCrLf
        s = s & "  from new_api_record_1" & vbCrLf
        s = s & " where oivnd=" & DbQuote(Str, txtVendor.Text) & vbCrLf
        s = s & "   and oiinv=" & DbQuote(Str, txtInvoice.Text) & vbCrLf
        On Error Resume Next
        Set rs = HFApp.SqlExec(s, dbAccountingDictionary)
        If err.Number = 0 Then
            If Not rs.EOF Then
                Call MsgBox(vbQuote & txtVendor & vbQuote & " invoice " & vbQuote & txtInvoice & vbQuote & " has already been entered.", vbExclamation + vbOK, App.ProductName)
                Cancel = True
                Exit Sub
            End If
        End If
        
        'check TL master file
        s = ""
        s = s & "select * " & vbCrLf
        s = s & "  from master_apm_record_1" & vbCrLf
        s = s & " where oivnd=" & DbQuote(Str, txtVendor) & vbCrLf
        s = s & "   and oiinv=" & DbQuote(Str, txtInvoice) & vbCrLf
        Set rs = HFApp.SqlExec(s, dbAccountingDictionary)
        If Not rs.EOF Then
            Call MsgBox(vbQuote & txtVendor & vbQuote & " invoice " & vbQuote & txtInvoice & vbQuote & " has already been entered.", vbExclamation + vbOK, App.ProductName)
            Cancel = True
            Exit Sub
        End If
    End If

    'check payables desk db
    s = ""
    s = s & "select * " & vbCrLf
    s = s & "  from invoices" & vbCrLf
    s = s & " where DivisionID =" & HFApp.DivisionID & " and vendor=" & DbQuote(Str, txtVendor) & vbCrLf
    s = s & "   and invoice=" & DbQuote(Str, txtInvoice) & vbCrLf
    Set rs = HFApp.SqlExec(s, dbHomeFront)
    If Not rs.EOF Then
        Call MsgBox(vbQuote & txtVendor & vbQuote & " invoice " & vbQuote & txtInvoice & vbQuote & " has already been entered.", vbExclamation + vbOK, App.ProductName)
        Cancel = True
        Exit Sub
    End If

End Sub

Private Sub txtVendor_GotFocus()
SelectAll txtVendor
End Sub

Private Sub txtVendor_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
        Case vbKeyF4:     If Shift = 0 Then Call lblVendor_Click
    End Select
End Sub

Private Sub lblVendor_Click()
    Dim c As Boolean
    Dim s As String
    If FPickList.Choose(HFApp.Databases(dbHomeFront), App.Options(Caption_Vendor), "SELECT Vendor_id " & Quote(App.Options(Caption_Vendor)) & ", Vendor_Name Name FROM tblvendors where DivisionID =" & HFApp.DivisionID, txtVendor.Text) Then
        s = FPickList.SelectedItem(1)
        txtVendor.Text = s
        On Error Resume Next
        Call txtVendor_Validate(c)
        txtInvoice.SetFocus
    End If
End Sub


Private Sub txtVendor_Validate(Cancel As Boolean)
On Error GoTo eh

    
    Dim s  As String
    Dim rs As Recordset
    
    If Trim(txtVendor.Text) <> "" Then
    
        s = ""
        s = s & "SELECT vendor_id vendor" & vbCrLf
        s = s & "      ,vendor_name   Name" & vbCrLf
        s = s & "      ,tradetype VendorType  " & vbCrLf
        s = s & "      ,addr1  Addr1" & vbCrLf
        s = s & "      ,addr2  Addr2" & vbCrLf
        s = s & "      ,city   City" & vbCrLf
        s = s & "      ,state  Prov" & vbCrLf
        s = s & "      ,zip    Postal" & vbCrLf
        s = s & "      ,PayTermDisc DiscPercent" & vbCrLf
        s = s & "      ,PayTermDiscDays DiscDays" & vbCrLf
        s = s & "      ,PayTermNetDays TermsDays" & vbCrLf
        s = s & "      ,PayTermType TermsType" & vbCrLf
        s = s & "      ,MiscDeductionRate WCBRate" & vbCrLf
        s = s & "      ,CostCode  DefPhase" & vbCrLf
        s = s & "      ,Category  DefCategory" & vbCrLf
        s = s & "      ,GlInsRequired,AutoInsRequired,WcInsRequired,UmbInsRequired" & vbCrLf
        s = s & "      ,DebitAccount  DefDebitAccount" & vbCrLf
        s = s & "      ,MaterialTaxGroup DefTaxGroup" & vbCrLf
        s = s & "      ,glinsexpdate   GlInsExpDate" & vbCrLf
        s = s & "      ,wcinsexpdate   WcInsExpDate" & vbCrLf
        s = s & "      ,umbinsexpdate  UmbInsExpDate" & vbCrLf
        s = s & "      ,autoinsexpdate AutoInsExpDate" & vbCrLf
        s = s & "from tblvendors" & vbCrLf
        s = s & " WHERE DivisionID =" & HFApp.DivisionID & " and vendor_id=" & DbQuote(Str, txtVendor) & vbCrLf
        Set rs = HFApp.SqlExec(s, dbHomeFront)
        If rs.EOF Then
            lblAddress = ""
            MsgBox "Vendor not found", vbExclamation, App.ProductName
            Cancel = True
        Else
            txtVendor = Trim("" & rs("vendor"))
            lblCompany = Trim("" & rs("name"))
            lblAddress = Trim("" & rs("Addr1")) & vbCrLf & _
                         IIf(Trim("" & rs("Addr2")) & vbCrLf = vbCrLf, "", Trim("" & rs("Addr2")) & vbCrLf) & _
                         Trim("" & rs("City")) & ", " & Trim("" & rs("Prov")) & vbCrLf & _
                         Trim("" & rs("Postal")) & vbCrLf

            mVendorType = Val("" & rs("vendortype"))
            
            'discount and terms
            s = ""
            mDiscDays = Val("" & rs("DiscDays"))
            mDiscPercent = Val("" & rs("DiscPercent"))
            If mDiscPercent <> 0 Then
                s = s & mDiscPercent & "% discount"
                If mDiscDays <> 0 Then
                    s = s & " if paid in " & mDiscDays & Choose(Min(mDiscDays, 2), " day", " days")
                End If
                s = s & vbCrLf
            End If
            mTermsDays = Val("" & rs("TermsDays"))
            mTermsType = Val("" & rs("TermsType"))
            If mTermsDays <> 0 Then
                If mTermsType = 1 Then
                    s = s & "Net due in " & mTermsDays & Choose(Min(mTermsDays, 2), " day", " days") & vbCrLf
                Else
                    s = s & "Net due on the " & mTermsDays & Choose(Min(mTermsDays, 4), "st", "nd", "rd", "th") & vbCrLf
                End If
            End If
            lblTerms = s
            
            

            
            'check insurance
            s = ""
            If rs("GlInsRequired") <> 0 And IsDate(rs("GlInsExpDate")) Then
                If Now > rs("GlInsExpDate") Then s = s & vbCrLf & "General Liability insurance expired on " & Format(rs("GlInsExpDate"), App.Options(TimberlineDateFormat))
            End If
            If rs("AutoInsRequired") <> 0 And IsDate(rs("AutoInsExpDate")) Then
                If Now > rs("AutoInsExpDate") Then s = s & vbCrLf & "Automobile insurance expired on " & Format(rs("AutoInsExpDate"), App.Options(TimberlineDateFormat))
            End If
            If rs("WcInsRequired") <> 0 And IsDate(rs("WcInsExpDate")) Then
                If Now > rs("WcInsExpDate") Then s = s & vbCrLf & "Workers' Comp insurance expired on " & Format(rs("WcInsExpDate"), App.Options(TimberlineDateFormat))
            End If
            If rs("UmbInsRequired") <> 0 And IsDate(rs("UmbInsExpDate")) Then
                If Now > rs("UmbInsExpDate") Then s = s & vbCrLf & "Umbrella insurance expired on " & Format(rs("UmbInsExpDate"), App.Options(TimberlineDateFormat))
            End If
            If s <> "" Then
                s = "Warning!" & vbCrLf & vbCrLf & "Vendor's insurance is not up-to-date." & vbCrLf & s
                MsgBox s, vbExclamation, App.ProductName
            End If
            
        End If
        
    End If
    Exit Sub
eh: Call ErrHandler(SRCFILE & "txtVendor_Validate")

End Sub


Private Function SaveToCSV(FileName As String) As String
On Error Resume Next
    Dim s As String
    Dim xlSheet As Object 'Excel.Worksheet
    Set xlSheet = GetObject(FileName).Sheets(1)
    s = TempFile("", "csv")
    Kill s
    Call xlSheet.SaveAs(s, 6, , , , , False)
    Set xlSheet = Nothing
    SaveToCSV = s
End Function

Private Sub ReadText(FileName As String)
On Error GoTo eh
    Dim i As Long
    Dim s As String
    Dim rs As Recordset
    
    Select Case FileExt(FileName)
        Case "csv": Call gData.LoadGrid(FileName, flexFileCommaText)
        Case Else:  Call gData.LoadGrid(FileName, flexFileTabText)
    End Select
    
    
    
    With gData
    
    'set column indexs
    For i = 1 To .Cols - 1
        .ColKey(i) = Replace(.TextMatrix(1, i), " ", "")
        .TextMatrix(0, i) = .ColKey(i)
    Next
    Call .RemoveItem(0)
    .FixedRows = 1
    
    'translate column headings
    Call ChangeColKey("Net", "Pretax")
    Call ChangeColKey("GST", "Tax")
    Call ChangeColKey("Vendor", "Description")
    
    'add columns debitaccount,job,extra,costcode
    .Cols = .Cols + 1:    .ColKey(.Cols - 1) = "DrAcct"
    .Cols = .Cols + 1:    .ColKey(.Cols - 1) = "Job"
    .Cols = .Cols + 1:    .ColKey(.Cols - 1) = "Extra"
    .Cols = .Cols + 1:    .ColKey(.Cols - 1) = "CostCode"
    
    'write headings
    For i = 1 To .Cols - 1
        .TextMatrix(0, i) = .ColKey(i)
    Next
    
    'remove any empty rows
    For i = .Rows - 1 To 1 Step -1
        If .TextMatrix(i, .ColIndex("Account")) = "" Then Call .RemoveItem(i)
    Next
    
    'parse out debitaccount,job,extra,costcode
    For i = 1 To .Rows - 1
        s = .TextMatrix(i, .ColIndex("Account"))
        .TextMatrix(i, .ColIndex("DrAcct")) = Parse(s, 1) & Parse(s, 2)
        .TextMatrix(i, .ColIndex("Job")) = Parse(s, 3)
        .TextMatrix(i, .ColIndex("Extra")) = Parse(s, 4)
        .TextMatrix(i, .ColIndex("CostCode")) = Parse(s, 5)
    Next

    
    'check that required cols are available
'    If .ColIndex("Vendor_ID") = -1 Then: Err.Raise 5, , "Required column ""Description"" is missing"
    
    
    End With
Exit Sub
eh: MsgBox err.Description, vbExclamation, App.ProductName
End Sub


Private Function ImportVisaInvoice() As Boolean
    Dim s As String
    Dim i As Long
    Dim iID As Long
    
    s = ""
    If Trim(txtVendor.Text) = "" Then s = s & "Vendor is required" & vbCrLf
    If Trim(txtInvoice.Text) = "" Then s = s & "Invoice is required" & vbCrLf
    If App.Options(InvoiceDateRequired) And dteInvoice.ValueIsNull Then s = s & "Invoice Date is required" & vbCrLf
    If App.Options(ReceivedDateRequired) And dteReceived.ValueIsNull Then s = s & "Received Date is required" & vbCrLf
    If App.Options(AccountingDateRequired) And dteAccounting.ValueIsNull Then s = s & "Accounting Date is required" & vbCrLf
    If Trim(txtFileName.Text) = "" Then s = s & "File Name is required" & vbCrLf
    If Not FileExists(txtFileName.Text) Then s = s & "File not found" & vbCrLf
    If s <> "" Then
        MsgBox s, vbInformation, App.ProductName
        ImportVisaInvoice = False
        Exit Function
    End If
        
    'load file
    s = txtFileName.Text
    Select Case FileExt(s)
        Case "xls":  Call ReadText(SaveToCSV(s))
        Case Else:   Call ReadText(s)
    End Select
        
    mVendor = txtVendor.Text
    mInvoice = txtInvoice.Text
    
    'insert invoice
    s = ""
    s = s & "insert into invoices(DivisionID,vendor,vendortype,vendorname,invoice,Status,ReceivedDate" & vbCrLf
    s = s & "                    ,InvoiceDate,AccountingDate,Description,UStmp,DStmp,TStmp)" & vbCrLf
    s = s & "values(" & HFApp.DivisionID & "," & DbQuote(Str, txtVendor.Text, , True) & vbCrLf
    s = s & "      ," & DbQuote(Str, mVendorType, , True) & vbCrLf
    s = s & "      ," & DbQuote(Str, lblCompany.Caption, , True) & vbCrLf
    s = s & "      ," & DbQuote(Str, txtInvoice.Text, , True) & vbCrLf
    s = s & "      ,'Pending'" & vbCrLf
    s = s & "      ," & DbQuote(Date, dteReceived.Value) & vbCrLf
    s = s & "      ," & DbQuote(Date, dteInvoice.Value) & vbCrLf
    s = s & "      ," & DbQuote(Date, dteAccounting.Value) & vbCrLf
    s = s & "      ," & DbQuote(Str, txtDescription.Text, , True) & vbCrLf
    s = s & "      ," & DbQuote(Str, HFApp.LoginID) & vbCrLf
    s = s & "      ," & DbQuote(Date, Now) & vbCrLf
    s = s & "      ," & DbQuote(Time, Now) & ")"
    Call HFApp.SqlExec(s, dbHomeFront)
    iID = HFApp.SqlIdentity("Invoices", dbHomeFront)
    
    'insert invoicedetails
    With gData
        For i = 1 To .Rows - 1
            s = ""
            s = s & "insert into invoiceitems(DivisionID,InvoiceID,Vendor,Invoice,Job,Extra,Phase,Category,DebitAccount" & vbCrLf
            s = s & "                         ,TaxGroup,TaxRate,TaxGroupDesc,Pretax,Tax,Description,UStmp,DStmp,TStmp)" & vbCrLf
            s = s & "values(" & HFApp.DivisionID & "," & DbQuote(Num, iID) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtVendor.Text, , True) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtInvoice.Text, , True) & vbCrLf
            
            If Trim(.TextMatrix(i, .ColIndex("Job"))) <> "" Then
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Job")), , True) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Extra")), , True) & vbCrLf
                s = s & "      ," & DbQuote(Str, Replace(.TextMatrix(i, .ColIndex("CostCode")), "-", ""), , True) & vbCrLf
                s = s & "      ,'M'" & vbCrLf
                s = s & "      ," & DbQuote(Str, Replace(.TextMatrix(i, .ColIndex("DrAcct")), "-", ""), , True) & vbCrLf
            Else
                s = s & "      ,''" & vbCrLf
                s = s & "      ,''" & vbCrLf
                s = s & "      ,''" & vbCrLf
                s = s & "      ,''" & vbCrLf
                s = s & "      ," & DbQuote(Str, Replace(.TextMatrix(i, .ColIndex("DrAcct")), "-", ""), , True) & vbCrLf
            End If
            
            s = s & "      ," & DbQuote(Str, App.Options(DefaultTaxGroup)) & vbCrLf
            s = s & "      ," & DbQuote(Num, App.Options(DefaultTaxGroupRate)) & vbCrLf
            s = s & "      ," & DbQuote(Str, App.Options(DefaultTaxGroupDesc)) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("Pretax")), , True) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("Tax")), , True) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Description")), , True) & vbCrLf
            s = s & "      ," & DbQuote(Str, HFApp.LoginID) & vbCrLf
            s = s & "      ," & DbQuote(Date, Now) & vbCrLf
            s = s & "      ," & DbQuote(Time, Now) & ")"
            Call HFApp.SqlExec(s, dbHomeFront)
        Next
    End With

    s = ""
    s = s & "update invoices " & vbCrLf
    s = s & "set pretax = (select sum(pretax) from invoiceitems where invoiceitems.invoiceid=invoices.invoiceid)" & vbCrLf
    s = s & "   ,tax = (select sum(tax) from invoiceitems where invoiceitems.invoiceid=invoices.invoiceid)" & vbCrLf
    s = s & "where DivisionID =" & HFApp.DivisionID & " and invoiceid=" & DbQuote(Num, iID) & vbCrLf
    Call HFApp.SqlExec(s, dbHomeFront)

    ImportVisaInvoice = True

End Function





Private Sub ChangeColKey(OldColKey As String, NewColKey As String)
On Error Resume Next
    With gData
        If .ColIndex(NewColKey) <> -1 Then .ColKey(.ColIndex(NewColKey)) = ""
        If .ColIndex(OldColKey) <> -1 Then .ColKey(.ColIndex(OldColKey)) = NewColKey
    End With
End Sub





Private Function ImportProfitMasterSalesReport() As Boolean
On Error GoTo eh:
    Dim s As String
    
    s = ""
    If Trim(txtVendor.Text) = "" Then s = s & "Vendor is required" & vbCrLf
    If App.Options(InvoiceDateRequired) And dteInvoice.ValueIsNull Then s = s & "Invoice Date is required" & vbCrLf
    If App.Options(ReceivedDateRequired) And dteReceived.ValueIsNull Then s = s & "Received Date is required" & vbCrLf
    If App.Options(AccountingDateRequired) And dteAccounting.ValueIsNull Then s = s & "Accounting Date is required" & vbCrLf
    If Trim(txtFileName.Text) = "" Then s = s & "File Name is required" & vbCrLf
    If Not FileExists(txtFileName.Text) Then s = s & "File not found" & vbCrLf
    If s <> "" Then
        MsgBox s, vbInformation, App.ProductName
        ImportProfitMasterSalesReport = False
        Exit Function
    End If
        
    Dim i     As Long
    Dim bSkip As Boolean
    Dim iFile As Integer
    Dim sFile As String
    Dim sBuff As String
    
    Dim SO     As String
    Dim SODesc As String
    Dim Inv    As String
    Dim OldInv As String
    Dim Part   As String
    Dim Desc   As String
    Dim Qty    As Double
    Dim Cost   As Double
    
    Dim rs      As Recordset
    Dim Job     As String
    Dim JobDesc As String
    Dim InvoiceID As String
    
    'Stupid unix files. Lines are ended with a vbLf instead of vbCrLf.
    'read it into a string variable
    iFile = FreeFile
    Open txtFileName.Text For Input As iFile
    While Not EOF(iFile)
        Line Input #iFile, sBuff
        sFile = sFile & sBuff & vbLf
    Wend
    Close iFile
        
    'the file has 2 dbl quotes where there should be 1. either a bug or another unix encoding difference?
    sFile = Replace(sFile, vbQuote & vbQuote, vbQuote)
        
    'parse each line out of string variable
    ProgressBar.Visible = True
    ProgressBar.Max = Parse(sFile, , vbLf)
    For i = 1 To Parse(sFile, , vbLf)
        ProgressBar.Value = i
        ProgressBar.Refresh
        sBuff = Parse(sFile, i, vbLf)
        bSkip = False
        
        Select Case True
            
            Case Trim(sBuff) = "":                               bSkip = True 'ignore empty line
            Case Mid(sBuff, 10, 4) = "THRU":                     bSkip = True 'ignore heading line 1
            Case Mid(sBuff, 56, 21) = "SALES ANALYSIS REPORT":   bSkip = True 'ignore heading line 2
            Case Mid(sBuff, 1, 8) = "RUN DATE":                  bSkip = True 'ignore heading line 3
            Case Mid(sBuff, 35, 7) = "Product":                  bSkip = True 'ignore heading line 4
            Case Mid(sBuff, 1, 10) = "----------":               bSkip = True 'ignore heading line 5
            Case Mid(sBuff, 34, 11) = "***********":             bSkip = True 'ignore summary line
            Case Mid(sBuff, 1, 12) = "************":             bSkip = True 'ignore all customers trailer line
            Case Trim(Mid(sBuff, 8, 6)) = SO:                    bSkip = True 'ignore service order trailer line
            
            'Service order line
            Case Trim(Mid(sBuff, 1, 13)) <> ""
                SO = Trim(Mid(sBuff, 1, 13))
                SODesc = Trim(Mid(sBuff, 14))
                bSkip = True
                
                'look up job
                s = ""
                s = s & "select j.job,j.description " & vbCrLf
                s = s & "from jcm_master__job_custom_fields c " & vbCrLf
                s = s & "inner join v j on(j.job=c.job)" & vbCrLf
                s = s & "where ltrim(rtrim(c.PM_Sales_Order))=" & DbQuote(Str, SO, , True)
                On Error Resume Next
                Set rs = HFApp.SqlExec(s, dbAccounting)
                If rs.EOF Then
                    Job = ""
                    JobDesc = ""
                Else
                    Job = Trim("" & rs("Job"))
                    JobDesc = Trim("" & rs("Description"))
                End If
                On Error GoTo eh
                
            Case Else
'                Inv = Trim(Mid(sBuff, 28, 8))
'                If Inv <> "" Then OldInv = Inv
'                Part = Trim(Mid(sBuff, 37, 11))
'                Desc = Trim(Mid(sBuff, 49, 25))
'                Qty = Val(Mid(sBuff, 77, 7))
'                Cost = Val(Mid(sBuff, 85, 10))
'                If Mid(sBuff, 95, 1) = "-" Then Cost = Cost * -1
                
'                Inv = Trim(Mid(sBuff, 25, 8))
'                If Inv <> "" Then OldInv = Inv
'                Part = Trim(Mid(sBuff, 34, 11))
'                Desc = Trim(Mid(sBuff, 46, 25))
'                Qty = Val(Mid(sBuff, 77, 7))
'                Cost = Val(Mid(sBuff, 85, 10))
'                If Mid(sBuff, 95, 1) = "-" Then Cost = Cost * -1
                
                Inv = Trim(Mid(sBuff, 27, 7))
                If Inv <> "" Then OldInv = Inv
                Part = Trim(Mid(sBuff, 34, 12))
                Desc = Trim(Mid(sBuff, 46, 26))
                Qty = Val(Mid(sBuff, 79, 5))
                Cost = Val(Mid(sBuff, 84, 11))
                If Mid(sBuff, 95, 1) = "-" Then Cost = Cost * -1
                
        End Select
        
    
        If Not bSkip Then
            
            If Inv <> "" Then
                
                'recalc old invoice
                If InvoiceID <> "" Then
                    s = ""
                    s = s & "update invoices " & vbCrLf
                    s = s & "set pretax = (select sum(pretax) from invoiceitems where invoiceitems.invoiceid=invoices.invoiceid)" & vbCrLf
                    s = s & "   ,tax = (select sum(tax) from invoiceitems where invoiceitems.invoiceid=invoices.invoiceid)" & vbCrLf
                    s = s & "where DivisionID =" & HFApp.DivisionID & " and invoiceid=" & DbQuote(Num, InvoiceID) & vbCrLf
                    Call HFApp.SqlExec(s, dbHomeFront)
                End If
                
                'insert invoice
                s = ""
                s = s & "insert into invoices(DivisionID,vendor,vendortype,vendorname,invoice,Status,ReceivedDate,Job,JobDesc" & vbCrLf
                s = s & "                    ,InvoiceDate,AccountingDate,Description,UStmp,DStmp,TStmp)" & vbCrLf
                s = s & "values(" & HFApp.DivisionID & "," & DbQuote(Str, txtVendor.Text, , True) & vbCrLf
                s = s & "      ," & DbQuote(Str, mVendorType, , True) & vbCrLf
                s = s & "      ," & DbQuote(Str, lblCompany.Caption, , True) & vbCrLf
                s = s & "      ," & DbQuote(Str, Inv) & vbCrLf
                s = s & "      ,'Hold'" & vbCrLf
                s = s & "      ," & DbQuote(Date, dteReceived.Value) & vbCrLf
                s = s & "      ," & DbQuote(Str, Job) & vbCrLf
                s = s & "      ," & DbQuote(Str, JobDesc) & vbCrLf
                s = s & "      ," & DbQuote(Date, dteInvoice.Value) & vbCrLf
                s = s & "      ," & DbQuote(Date, dteAccounting.Value) & vbCrLf
                s = s & "      ," & DbQuote(Str, SO & " " & SODesc, , True, 30) & vbCrLf
                s = s & "      ," & DbQuote(Str, HFApp.LoginID) & vbCrLf
                s = s & "      ," & DbQuote(Date, Now) & vbCrLf
                s = s & "      ," & DbQuote(Time, Now) & ")"
                Call HFApp.SqlExec(s, dbHomeFront)
                InvoiceID = HFApp.SqlIdentity("Invoices", dbHomeFront)
                
            End If
            
            s = ""
            s = s & "insert into invoiceitems(DivisionID,InvoiceID,Vendor,Invoice,Job,JobDesc,InvoicedQuantity" & vbCrLf
            s = s & "                         ,TaxGroup,TaxRate,TaxGroupDesc,Pretax,Tax,Description,UStmp,DStmp,TStmp)" & vbCrLf
            s = s & "values(" & HFApp.DivisionID & "," & DbQuote(Num, InvoiceID) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtVendor.Text, , True) & vbCrLf
            s = s & "      ," & DbQuote(Str, OldInv, , True) & vbCrLf
            s = s & "      ," & DbQuote(Str, Job) & vbCrLf
            s = s & "      ," & DbQuote(Str, JobDesc) & vbCrLf
            s = s & "      ," & DbQuote(Num, Qty) & vbCrLf
            s = s & "      ," & DbQuote(Str, App.Options(DefaultTaxGroup)) & vbCrLf
            s = s & "      ," & DbQuote(Num, App.Options(DefaultTaxGroupRate)) & vbCrLf
            s = s & "      ," & DbQuote(Str, App.Options(DefaultTaxGroupDesc)) & vbCrLf
            s = s & "      ," & DbQuote(Num, Cost) & vbCrLf
            s = s & "      ," & DbQuote(Num, Round(Cost * Val(App.Options.Value(DefaultTaxGroupRate)) / 100, 2)) & vbCrLf
            s = s & "      ," & DbQuote(Str, Desc) & vbCrLf
            s = s & "      ," & DbQuote(Str, HFApp.LoginID) & vbCrLf
            s = s & "      ," & DbQuote(Date, Now) & vbCrLf
            s = s & "      ," & DbQuote(Time, Now) & ")"
            Call HFApp.SqlExec(s, dbHomeFront)
            
        End If
    Next
    
    MsgBox ProgressBar.Max & " Held invoices have been added.", vbInformation, App.ProductName

    ImportProfitMasterSalesReport = True
    
Exit Function
eh:
Select Case True
    Case InStr(1, err.Description, "duplicate", vbTextCompare)
        s = ""
        s = s & "Unable to continue importing this file. It may already have been imported." & vbCrLf & vbCrLf
        s = s & "Invoice " & OldInv & " is already in HomeFronts database."
        MsgBox s, vbExclamation, App.ProductName
        ProgressBar.Visible = False
        
    Case Else
        Call ErrHandler(SRCFILE & "ImportProfitMasterSalesReport")
        ProgressBar.Visible = False
End Select
End Function

