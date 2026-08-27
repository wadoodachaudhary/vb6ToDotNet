VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form FImport 
   Caption         =   "Import"
   ClientHeight    =   9060
   ClientLeft      =   240
   ClientTop       =   1695
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   9060
   ScaleWidth      =   6585
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   1860
      Left            =   120
      TabIndex        =   0
      Top             =   825
      Width           =   4845
      _cx             =   8546
      _cy             =   3281
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
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FImport.frx":0000
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
End
Attribute VB_Name = "FImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FImport::"

Dim mSessionID As String

Public Function ShowForm() As Boolean
On Error GoTo eh:
    
    Dim Filename As String
    Dim filter As String
    Dim r As Long
    Dim s As String
        
    
    mSessionID = CreateGUID()
    
    filter = "Excel Files (*.csv;*.xls;*.xlsx)|*.csv;*.xls;*.xlsx"
    If Not VBGetOpenFileName(Filename, , , , , True, filter, , , "AP Import", , Screen.ActiveForm.hwnd) Then Exit Function
    If FileExt(Filename) = "csv" Then
        s = Filename
    Else
        s = SaveToCSV(Filename)
    End If
    Call LoadFile(s)
    On Error Resume Next
    If FileExt(Filename) <> "csv" Then Kill s
    On Error GoTo eh

    
    Call SaveToTables


Exit Function
eh: Call ErrHandler(SRCFILE & "ShowForm", s)
End Function


Private Sub LoadFile(Filename As String)
On Error GoTo eh:

    Dim i As Long
    Dim s As String

    With gData
        
        'load into grid
        .FixedRows = 0
        Call .LoadGrid(Filename, flexFileCommaText)
        .FixedRows = 1
        
        'set column keys
        '0,1,2 are common to both record types (I)nvoice and (D)istribution
        '3-12 are used by both types. After loading shuffle columns over for D types so they are in their own columns
        '3->13, 4->14,... 12->22
        .Cols = 30
        .ColKey(0) = "RecType"
        .ColKey(1) = "Vendor"
        .ColKey(2) = "Invoice"
        '--
        .ColKey(3) = "InvoiceDate"
        .ColKey(4) = "PaymentDate"
        .ColKey(5) = "ReceivedDate"
        .ColKey(6) = "AccountingDate"
        .ColKey(7) = "Discount"
        .ColKey(8) = "DiscountDate"
        .ColKey(9) = "InvoiceDescription"
        .ColKey(10) = "Comments"
        .ColKey(11) = "InvoiceCode1"
        .ColKey(12) = "InvoiceCode2"
        '--
        .ColKey(13) = "PO"
        .ColKey(14) = "POItem"
        .ColKey(15) = "Job"
        .ColKey(16) = "Extra"
        .ColKey(17) = "CostCode"
        .ColKey(18) = "Category"
        .ColKey(19) = "DebitAccount"
        .ColKey(20) = "Equipment"
        .ColKey(21) = "EqCostCode"
        .ColKey(22) = "TaxGroup"
        .ColKey(23) = "Qty"
        .ColKey(24) = "UnitPrice"
        .ColKey(25) = "Pretax"
        .ColKey(26) = "Tax"
        .ColKey(27) = "Retainage"
        .ColKey(28) = "Description"
        .ColKey(29) = "Billable"
        
        'set column titles
        For i = 0 To 29
            .TextMatrix(0, i) = .ColKey(i)
        Next
        Call .RemoveItem(1)
        
        'shuffle distribution columns over
        For i = 1 To .Rows - 1
        If .TextMatrix(i, .ColIndex("RecType")) = "D" Then
            .Cell(flexcpText, i, 13, i, 29) = .Cell(flexcpText, i, 3, i, 19)
            .Cell(flexcpText, i, 3, i, 12) = ""
        End If
        Next
        
    End With
    
Exit Sub
eh: Call ErrHandler(SRCFILE & "LoadFile", s)
End Sub

Private Sub SaveToTables()
On Error GoTo eh:
    Dim r As Long
    Dim s As String
    
    With gData
    
    s = "insert into ImportedAPInvoices(SessionID,DivisionID,Vendor,Invoice,InvoiceDate,PaymentDate,ReceivedDate,AccountingDate,Discount,DiscountDate,Description,Comments,InvoiceCode1,InvoiceCode2) values" & vbCrLf
    For r = 1 To .Rows - 1
    If .TextMatrix(r, .ColIndex("rectype")) = "I" Then
        s = s & "(" & DbQuote(Str, mSessionID)
        s = s & "," & DbQuote(num, HFApp.DivisionID)
        s = s & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("Vendor")))
        s = s & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("Invoice")))
        s = s & "," & DbQuote(Date, .TextMatrix(r, .ColIndex("InvoiceDate")))
        s = s & "," & DbQuote(Date, .TextMatrix(r, .ColIndex("PaymentDate")))
        s = s & "," & DbQuote(Date, .TextMatrix(r, .ColIndex("ReceivedDate")))
        s = s & "," & DbQuote(Date, .TextMatrix(r, .ColIndex("AccountingDate")))
        s = s & "," & DbQuote(num, .TextMatrix(r, .ColIndex("Discount")))
        s = s & "," & DbQuote(Date, .TextMatrix(r, .ColIndex("DiscountDate")))
        s = s & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("InvoiceDescription")))
        s = s & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("Comments")))
        s = s & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("InvoiceCode1")))
        s = s & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("InvoiceCode2")))
        s = s & ")," & vbCrLf
    End If
    Next
    s = left(s, Len(s) - 3)
    Call HFApp.SqlExec(s)
    
    
    s = "insert into ImportedAPDistributions(SessionID,DivisionID,Vendor,Invoice,PO,POItem,Job,Extra,CostCode,Category,DebitAccount,Equipment,EqCostCode,TaxGroup,Qty,UnitPrice,Pretax,Tax,Retainage,Description,Billable) values" & vbCrLf
    For r = 1 To .Rows - 1
    If .TextMatrix(r, .ColIndex("rectype")) = "D" Then
        s = s & "(" & DbQuote(Str, mSessionID)
        s = s & "," & DbQuote(num, HFApp.DivisionID)
        s = s & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("Vendor")))
        s = s & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("Invoice")))
        s = s & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("PO")))
        s = s & "," & DbQuote(num, .TextMatrix(r, .ColIndex("POItem")))
        s = s & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("Job")))
        s = s & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("Extra")))
        s = s & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("CostCode")))
        s = s & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("Category")))
        s = s & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("DebitAccount")))
        s = s & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("Equipment")))
        s = s & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("EqCostCode")))
        s = s & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("TaxGroup")))
        s = s & "," & DbQuote(num, .TextMatrix(r, .ColIndex("Qty")))
        s = s & "," & DbQuote(num, .TextMatrix(r, .ColIndex("UnitPrice")))
        s = s & "," & DbQuote(num, .TextMatrix(r, .ColIndex("Pretax")))
        s = s & "," & DbQuote(num, .TextMatrix(r, .ColIndex("Tax")))
        s = s & "," & DbQuote(num, .TextMatrix(r, .ColIndex("Retainage")))
        s = s & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("Description")))
        s = s & "," & DbQuote(Bit, .TextMatrix(r, .ColIndex("Billable")))
        s = s & ")," & vbCrLf
    End If
    Next
    s = left(s, Len(s) - 3)
    Call HFApp.SqlExec(s)
    
    
    s = "exec dbo.AP_ImportFromFile " & DbQuote(Str, mSessionID) & "," & DbQuote(Str, HFApp.LoginID)
    Dim rs As Recordset
    Set rs = HFApp.SqlExec(s)
    s = rs(0)
    If s = "" Then
        MsgBox "Import complete. Check the pending invoices list.", vbInformation + vbOKOnly, App.ProductName
    Else
        MsgBox "Import failed." & vbCrLf & vbCrLf & s, vbExclamation + vbOKOnly, App.ProductName
    End If
        
    On Error Resume Next
    Unload Me
    
    End With
    
Exit Sub
eh: Call ErrHandler(SRCFILE & "SaveToDB", s)
End Sub

Private Sub Form_Resize()
    gData.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub
