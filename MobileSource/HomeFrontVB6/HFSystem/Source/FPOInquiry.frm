VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form FPOInquiry 
   Caption         =   "Purchase Order Inquiry"
   ClientHeight    =   5295
   ClientLeft      =   5730
   ClientTop       =   1935
   ClientWidth     =   11160
   Icon            =   "FPOInquiry.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   11160
   Begin VB.TextBox txtRemaining 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   245
      Left            =   9540
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   " "
      Top             =   1230
      Width           =   1365
   End
   Begin VB.TextBox txtOriginal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      ForeColor       =   &H80000012&
      Height          =   245
      Left            =   9540
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   " "
      Top             =   210
      Width           =   1365
   End
   Begin VB.TextBox txtChanges 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   245
      Left            =   9540
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   " "
      Top             =   465
      Width           =   1365
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   245
      Left            =   9540
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   " "
      Top             =   720
      Width           =   1365
   End
   Begin VB.TextBox txtInvoiced 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   245
      Left            =   9540
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   " "
      Top             =   975
      Width           =   1365
   End
   Begin VB.TextBox txtDateDelivered 
      BorderStyle     =   0  'None
      Height          =   245
      Left            =   6150
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   " "
      Top             =   720
      Width           =   2445
   End
   Begin VB.TextBox txtDatePosted 
      BorderStyle     =   0  'None
      Height          =   245
      Left            =   6150
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   " "
      Top             =   465
      Width           =   2445
   End
   Begin VB.TextBox txtDateIssued 
      BorderStyle     =   0  'None
      Height          =   245
      Left            =   6150
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   " "
      Top             =   210
      Width           =   1845
   End
   Begin VB.TextBox txtJob 
      BorderStyle     =   0  'None
      Height          =   245
      Left            =   1980
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   " "
      Top             =   975
      Width           =   3225
   End
   Begin VB.TextBox txtVendor 
      BorderStyle     =   0  'None
      Height          =   245
      Left            =   1980
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   " "
      Top             =   720
      Width           =   3225
   End
   Begin VB.TextBox txtPOIndex 
      BorderStyle     =   0  'None
      Height          =   245
      Left            =   1980
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   " "
      Top             =   465
      Width           =   3225
   End
   Begin VB.TextBox txtPONumber 
      BorderStyle     =   0  'None
      ForeColor       =   &H80000012&
      Height          =   245
      Left            =   1980
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   " "
      Top             =   210
      Width           =   3225
   End
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   3135
      Left            =   240
      TabIndex        =   7
      Top             =   1770
      Width           =   10665
      _cx             =   18812
      _cy             =   5530
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
      ForeColor       =   -2147483630
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
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FPOInquiry.frx":000C
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
      ExplorerBar     =   0
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
   Begin VB.Label lblCancelled 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CANCELLED"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   345
      Left            =   5790
      TabIndex        =   16
      Top             =   1020
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Remaining"
      Height          =   195
      Index           =   11
      Left            =   8715
      TabIndex        =   26
      Top             =   1260
      Width           =   750
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Original"
      Height          =   195
      Index           =   10
      Left            =   8940
      TabIndex        =   24
      Top             =   240
      Width           =   525
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Changes"
      Height          =   195
      Index           =   9
      Left            =   8835
      TabIndex        =   23
      Top             =   510
      Width           =   630
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Total"
      Height          =   195
      Index           =   8
      Left            =   9105
      TabIndex        =   22
      Top             =   750
      Width           =   360
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Invoiced"
      Height          =   195
      Index           =   7
      Left            =   8850
      TabIndex        =   21
      Top             =   990
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Delivered"
      Height          =   195
      Index           =   6
      Left            =   5400
      TabIndex        =   15
      Top             =   750
      Width           =   675
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Posted"
      Height          =   195
      Index           =   5
      Left            =   5580
      TabIndex        =   14
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Issued"
      Height          =   195
      Index           =   4
      Left            =   5610
      TabIndex        =   13
      Top             =   240
      Width           =   465
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Job"
      Height          =   195
      Index           =   3
      Left            =   1650
      TabIndex        =   12
      Top             =   1020
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Vendor"
      Height          =   195
      Index           =   2
      Left            =   1395
      TabIndex        =   11
      Top             =   750
      Width           =   510
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "PO Index"
      Height          =   195
      Index           =   1
      Left            =   1245
      TabIndex        =   10
      Top             =   480
      Width           =   660
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "PO Number"
      Height          =   195
      Index           =   0
      Left            =   1080
      TabIndex        =   9
      Top             =   240
      Width           =   825
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   300
      Picture         =   "FPOInquiry.frx":00FD
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Accounting Documents"
      Height          =   285
      Left            =   270
      TabIndex        =   8
      Top             =   1500
      Width           =   2925
   End
End
Attribute VB_Name = "FPOInquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mPO As String


Private Sub Form_Load()
    Call IniGetForm(Me)
    Call LoadData
End Sub

Private Sub Form_Resize()
    With gData
        .Move .left, .Top, Me.ScaleWidth - 2 * .left, Me.ScaleHeight - .Top - .left
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub


Public Sub ShowForm(PO As String)
    mPO = PO
    Me.Show vbModal
End Sub

Private Sub LoadData()
    Dim s As String
    Dim rs As Recordset
    Dim c As Connections
    Dim docid As String
    Dim doctype As String
    Dim docRow As Long
    Dim pretax As Double
    Dim tax As Double


    'load po info
    s = ""
    s = s & "SELECT" & vbCrLf
    s = s & " m.PO PONumber" & vbCrLf
    s = s & ",m.PODesc" & vbCrLf
    s = s & ",m.POIndex" & vbCrLf
    s = s & ",m.Vendor" & vbCrLf
    s = s & ",m.VendorDesc" & vbCrLf
    s = s & ",m.POJob Job" & vbCrLf
    s = s & ",m.POJobDesc JobDesc" & vbCrLf
    s = s & ",cast(m.PODate as date) PODate" & vbCrLf
    s = s & ",b.TStmp PostingDate" & vbCrLf
    s = s & ",m.PostingBatch" & vbCrLf
    s = s & ",p.DeliveryDate" & vbCrLf
    s = s & ",case when p.DeliveryMethod=1 then 'email' else 'print' end DeliveryMethod" & vbCrLf
    s = s & ",p.DeliveryAddress" & vbCrLf
    s = s & ",p.Cancelled" & vbCrLf
    s = s & ",p.CancelledDate" & vbCrLf
    s = s & ",SUM(m.LinePretax) Original" & vbCrLf
    s = s & ",SUM(m.LineChangesPretax) Changes" & vbCrLf
    s = s & ",SUM(m.LineInvoicedPretax) Invoiced" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "FROM jcpodetails m" & vbCrLf
    s = s & "left outer join batches b on m.postingbatch=b.batch" & vbCrLf
    s = s & "left outer join pomaster p on m.po=p.ponumber and m.divisionid=p.divisionid" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "WHERE m.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
    s = s & "  and m.po=" & DbQuote(Str, mPO) & vbCrLf
    s = s & "GROUP BY" & vbCrLf
    s = s & " m.PO " & vbCrLf
    s = s & ",m.PODesc" & vbCrLf
    s = s & ",m.POIndex" & vbCrLf
    s = s & ",m.Vendor" & vbCrLf
    s = s & ",m.VendorDesc" & vbCrLf
    s = s & ",m.POJob " & vbCrLf
    s = s & ",m.POJobDesc " & vbCrLf
    s = s & ",cast(m.PODate as date) " & vbCrLf
    s = s & ",b.TStmp " & vbCrLf
    s = s & ",m.PostingBatch" & vbCrLf
    s = s & ",p.DeliveryDate" & vbCrLf
    s = s & ",p.DeliveryMethod" & vbCrLf
    s = s & ",p.DeliveryAddress" & vbCrLf
    s = s & ",p.Cancelled" & vbCrLf
    s = s & ",p.CancelledDate" & vbCrLf
    Set rs = HFApp.SqlExec(s, dbHomeFront)
    If Not rs.EOF Then
        txtPONumber.Text = "" & rs("PONumber") & "   " & rs("PODesc")
        txtPOIndex.Text = "" & rs("POIndex")
        txtVendor.Text = IIf(Val(HFApp.Options(AccountingSystem)) <> asQuickbooks, "" & rs("Vendor") & "   ", "") & rs("VendorDesc")
        txtJob.Text = "" & rs("Job") & "   " & rs("JobDesc")
        txtDateIssued.Text = Format("" & rs("PODate"), "mmmm d, yyyy")
        txtDatePosted.Text = Format("" & rs("PostingDate"), "mmmm d, yyyy") & "   Batch " & rs("PostingBatch")
        txtDateDelivered.Text = Format("" & rs("DeliveryDate"), "mmmm d, yyyy") & "   Via " & rs("DeliveryMethod")
        lblCancelled.Visible = "True" = "" & rs("Cancelled")
        lblCancelled.Caption = "CANCELLED " & Format("" & rs("CancelledDate"), "mmmm d, yyyy")
        txtOriginal.Text = Format(Val("" & rs("Original")), "#,##0.00")
        txtChanges.Text = Format(Val("" & rs("Changes")), "#,##0.00")
        txtTotal.Text = Format(Val("" & rs("Original")) + Val("" & rs("Changes")), "#,##0.00")
        txtInvoiced.Text = Format(Val("" & rs("Invoiced")), "#,##0.00")
        txtRemaining.Text = Format(Val("" & rs("Original")) + Val("" & rs("Changes")) - Val("" & rs("Invoiced")), "#,##0.00")
    End If




    'load document details
    s = ""
    s = s & "select" & vbCrLf
    s = s & " 'Purchase Order' DocType" & vbCrLf
    s = s & ",PO DocID" & vbCrLf
    s = s & ",PODate DocDate" & vbCrLf
    s = s & ",PODesc DocDesc" & vbCrLf
    s = s & ",LineDesc Description" & vbCrLf
    s = s & ",LinePretax+LineChangesPretax Amount" & vbCrLf
    s = s & ",LineTax+LineChangesTax Tax" & vbCrLf
    s = s & "from jcpodetails" & vbCrLf
    s = s & "where divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
    s = s & "  and po=" & DbQuote(Str, mPO) & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "select" & vbCrLf
    s = s & " 'Invoice' " & vbCrLf
    s = s & ",m.Invoice " & vbCrLf
    s = s & ",m.invoicedate " & vbCrLf
    s = s & ",m.Description " & vbCrLf
    s = s & ",d.description " & vbCrLf
    s = s & ",d.pretax " & vbCrLf
    s = s & ",d.tax" & vbCrLf
    s = s & "from Invoices m" & vbCrLf
    s = s & "left outer join InvoiceItems d on m.InvoiceID=d.InvoiceID" & vbCrLf
    s = s & "where m.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
    s = s & "  and d.commitment=" & DbQuote(Str, mPO) & vbCrLf
    With gData
        .Rows = 1
        .OutlineBar = flexOutlineBarCompleteLeaf
        .OutlineCol = .ColIndex("Description")
        docid = Chr(1)
        doctype = Chr(1)
        If s <> "" Then
            Set rs = HFApp.SqlExec(s, dbHomeFront)
            While Not rs.EOF
                
                If docid <> "" & rs("DocID") Or doctype <> "" & rs("DocType") Then
                    If docRow <> 0 Then
                        .TextMatrix(docRow, .ColIndex("Amount")) = pretax
                        .TextMatrix(docRow, .ColIndex("Tax")) = tax
                    End If
                    
                    doctype = "" & rs("DocType")
                    docid = "" & rs("DocID")
                    docRow = .Rows
                    pretax = 0
                    tax = 0
                    .AddItem ""
                    .IsSubtotal(.Rows - 1) = True
                    .RowOutlineLevel(.Rows - 1) = 0
                    .TextMatrix(.Rows - 1, .ColIndex("DocType")) = "" & rs("DocType")
                    .TextMatrix(.Rows - 1, .ColIndex("DocID")) = "" & rs("DocID")
                    .TextMatrix(.Rows - 1, .ColIndex("DocDate")) = Format("" & rs("DocDate"), "mmmm d, yyyy")
                    .TextMatrix(.Rows - 1, .ColIndex("Description")) = "" & rs("DocDesc")
                    .IsSubtotal(.Rows - 1) = True
                    .RowOutlineLevel(.Rows - 1) = 0
                End If
                
                .AddItem ""
                .IsSubtotal(.Rows - 1) = True
                .RowOutlineLevel(.Rows - 1) = 1
                .TextMatrix(.Rows - 1, .ColIndex("Description")) = "" & rs("Description")
                .TextMatrix(.Rows - 1, .ColIndex("Amount")) = "" & rs("Amount")
                .TextMatrix(.Rows - 1, .ColIndex("Tax")) = "" & rs("Tax")
                .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HC00000
                
                pretax = pretax + Val("" & rs("Amount"))
                tax = tax + Val("" & rs("Tax"))
                
                rs.MoveNext
            Wend
            If docRow <> 0 Then
                .TextMatrix(docRow, .ColIndex("Amount")) = pretax
                .TextMatrix(docRow, .ColIndex("Tax")) = tax
            End If
        End If
        Call .Outline(0)
        Call .AutoSize(0, .Cols - 1)
    End With
    
    
End Sub
