VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form FWalletBatches 
   Caption         =   "Wallet Batches"
   ClientHeight    =   8145
   ClientLeft      =   1680
   ClientTop       =   3510
   ClientWidth     =   13245
   Icon            =   "FWalletBatches.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   13245
   Begin VB.CommandButton cmdNav 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   1905
      TabIndex        =   4
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&Send"
      Height          =   375
      Index           =   0
      Left            =   285
      TabIndex        =   3
      Top             =   6705
      Width           =   1215
   End
   Begin VSFlex8Ctl.VSFlexGrid gBatches 
      Height          =   3255
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1815
      Width           =   5760
      _cx             =   10160
      _cy             =   5741
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
      BackColorSel    =   -2147483633
      ForeColorSel    =   -2147483640
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
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
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
      FormatString    =   $"FWalletBatches.frx":000C
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   1
      AutoSearchDelay =   3
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   3
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
   Begin VB.Frame FrameSearch 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "4"
      Height          =   1050
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   16845
      Begin VB.Label lblTotalSelected 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Selected: $0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2715
         TabIndex        =   6
         Top             =   720
         Width           =   1770
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000F&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   2.30016e6
      End
      Begin VB.Label lblCaption 
         BackStyle       =   0  'Transparent
         Caption         =   $"FWalletBatches.frx":00F4
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   900
         TabIndex        =   2
         Top             =   60
         Width           =   9075
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   210
         Picture         =   "FWalletBatches.frx":01CA
         Top             =   210
         Width           =   480
      End
   End
   Begin HFPayables.Slider Slider 
      Height          =   4335
      Left            =   5970
      Top             =   2460
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   7646
   End
   Begin VSFlex8Ctl.VSFlexGrid gPayments 
      Height          =   3255
      Left            =   6375
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1755
      Width           =   5760
      _cx             =   10160
      _cy             =   5741
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
      BackColorSel    =   -2147483633
      ForeColorSel    =   -2147483640
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
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FWalletBatches.frx":0A94
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   1
      AutoSearchDelay =   3
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
End
Attribute VB_Name = "FWalletBatches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private Const SRCFILE = "FWalletBatches::"

Private DataSourceID As String
Private WalletBankAccount As String

Private CompanyID As String
Private CompanyDescription As String

Public Function ShowForm()
On Error GoTo eh:
    Dim CMHCompanies As String
    Dim s As String
    Dim rs As Recordset
    Dim GLPrefixLength As Long
    
    s = ""
    s = s & "select s.DataSourceID,s.Sage300GLPrefixLength,s.walletbankaccount" & vbCrLf
    s = s & "from divisions d" & vbCrLf
    s = s & "join datasources s on d.bookofaccount=s.bookofaccount" & vbCrLf
    s = s & "where d.divisionid=" & HFApp.DivisionID
    Set rs = HFApp.SqlExec(s)
    DataSourceID = "" & rs("DataSourceID")
    GLPrefixLength = Val("" & rs("Sage300GLPrefixLength"))
    WalletBankAccount = "" & rs("WalletBankAccount")
    CompanyID = "1"
    CompanyDescription = ""
    
    
    
    If HFApp.Options(AccountingSystem) = asTimberline And GLPrefixLength <> 0 Then
        s = ""
        s = s & "select " & vbCrLf
        s = s & " a.CompanyID" & vbCrLf
        s = s & ",dbo.Wallet_Sage300GLPrefixDescription(a.datasourceid,min(a.glprefix)) Description" & vbCrLf
        s = s & ",count(distinct isnull(nullif(a.checknumber,''),a.paymentnumber)) [Payment Count]" & vbCrLf
        s = s & "from Wallet_AccountingAP a" & vbCrLf
        s = s & "left join wallet_payments p on a.datasourceid=p.datasourceid and a.transactionid=p.transactionid" & vbCrLf
        s = s & "where a.datasourceid=" & DbQuote(Str, DataSourceID) & vbCrLf
        s = s & "and a.bankaccount=" & DbQuote(Str, WalletBankAccount) & vbCrLf
        s = s & "and isnull(a.paymentvoided,0)=0" & vbCrLf
        s = s & "and p.transactionid is null" & vbCrLf
        s = s & "group by a.datasourceid,a.companyid" & vbCrLf
        If FPickList.Choose(HFApp.Databases(dbHomeFront), "Company", s) Then
            CompanyID = FPickList.SelectedItem("CompanyID")
            CompanyDescription = FPickList.SelectedItem("Description")
            Me.Show vbModal
        End If
    Else
        Me.Show vbModal
    End If
    


Exit Function
eh: Call ErrHandler(SRCFILE & "ShowForm", s)
End Function


Private Function TotalSelected() As Currency
    Dim i As Long
    Dim t As Currency
    t = 0
    With gBatches
    For i = 1 To .Rows - 1
        If .Cell(flexcpChecked, i, .ColIndex("selected")) = flexChecked Then t = t + .ValueMatrix(i, .ColIndex("BatchAmount"))
    Next
    End With
    TotalSelected = t

End Function

Private Sub cmdNav_Click(Index As Integer)
On Error GoTo eh
    
    Select Case Index
        Case 0 'generate
            Call SaveData(Index = 1)
            Unload Me
            
        Case 1 'close
            If SaveData(True) Then Unload Me

            
    End Select


Exit Sub
eh: Call ErrHandler(SRCFILE & "cmdNav_Click")
End Sub


Public Function SaveData(prompt As Boolean) As Boolean
On Error GoTo eh

    Dim s As String
    Dim b As String
    Dim r As Long
    Dim wb As Long
    
    With gBatches
        b = ""
        For r = 1 To .Rows - 1
            If .Cell(flexcpChecked, r, 1) = flexChecked Then
                b = b & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("AccountingBatch")))
            End If
        Next
        b = Mid(b, 2)
        If b = "" Then
            SaveData = True
            Exit Function
        End If
    End With
    
    If prompt Then
        Select Case MsgBox("You have batches selected but have not sent them." & vbCrLf & vbCrLf & "Do you want to send them?" & vbCrLf, vbExclamation + vbYesNoCancel, App.ProductName)
            Case vbNo
                SaveData = True
                Exit Function
            Case vbCancel
                SaveData = False
                Exit Function
        End Select
    End If


    If vbCancel = MsgBox("You are about to create a Wallet batch in the amount of " & Format(TotalSelected, "currency") & "." & vbCrLf, vbExclamation + vbOKCancel, App.ProductName) Then
        SaveData = False
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass

    s = "exec dbo.Wallet_CreateBatch " & DbQuote(Str, DataSourceID) & "," & DbQuote(Str, CompanyID) & "," & DbQuote(Str, b)
    Call HFApp.SqlExec(s)
    
    Screen.MousePointer = vbDefault

Exit Function
eh: Call ErrHandler(SRCFILE & "SaveData")
End Function



Private Sub Form_Load()
    Call IniGetForm(Me)
    
    If CompanyDescription <> "" Then Me.Caption = "Wallet Batches - " & CompanyDescription
    
    Call LoadData
End Sub

Private Sub LoadData()
    Dim s As String
    Dim rs As Recordset
    
    'load unbatched ap payments

    s = ""
    s = s & "select " & vbCrLf
    s = s & " a.AccountingBatch" & vbCrLf
    s = s & ",count(distinct isnull(nullif(a.paymentnumber,''),a.checknumber)) as PaymentCount" & vbCrLf
    s = s & ",sum(paymentAmount) BatchAmount" & vbCrLf
    s = s & "from" & vbCrLf
    s = s & "(" & vbCrLf
    s = s & "    select a.AccountingBatch,a.CheckNumber,a.PaymentNumber,a.NetAmount PaymentAmount" & vbCrLf
    s = s & "    from Wallet_AccountingAP a " & vbCrLf
    s = s & "    left join wallet_payments p on a.datasourceid=p.datasourceid and a.transactionid=p.transactionid" & vbCrLf
    s = s & "    where a.datasourceid=" & DbQuote(Str, DataSourceID) & vbCrLf
    s = s & "    and a.companyid=" & DbQuote(Str, CompanyID) & vbCrLf
    s = s & "    and a.bankaccount=" & DbQuote(Str, WalletBankAccount) & vbCrLf
    s = s & "    and isnull(a.paymentvoided,0)=0" & vbCrLf
    s = s & "    and p.transactionid is null" & vbCrLf
    s = s & ") a" & vbCrLf
    s = s & "group by a.AccountingBatch" & vbCrLf
    
    Set rs = HFApp.SqlExec(s)
    With gBatches
        .Rows = 1
        While Not rs.EOF
            .AddItem ""
            .TextMatrix(.Rows - 1, .ColIndex("AccountingBatch")) = "" & rs("AccountingBatch")
            .TextMatrix(.Rows - 1, .ColIndex("PaymentCount")) = "" & rs("PaymentCount")
            .TextMatrix(.Rows - 1, .ColIndex("BatchAmount")) = "" & rs("BatchAmount")
            rs.MoveNext
        Wend
    End With

End Sub

Private Sub Form_Resize()
On Error Resume Next
    Const margin = 60


    Slider.Min = 960
    Slider.Max = Me.ScaleWidth - 960
    If Slider.left > Slider.Max Then Slider.left = Slider.Max
    
    FrameSearch.Width = Me.ScaleWidth
    
    
    Slider.Move Slider.left, FrameSearch.Height, Slider.Width, Me.ScaleHeight - FrameSearch.Height - cmdNav(0).Height - 2 * margin
    gBatches.Move 0, FrameSearch.Height, Slider.left, Slider.Height
    gPayments.Move Slider.left + Slider.Width, Slider.Top, Me.ScaleWidth - Slider.left - Slider.Width, Slider.Height
    
    cmdNav(0).Move Me.ScaleWidth - 2 * cmdNav(0).Width - 2 * margin, Me.ScaleHeight - cmdNav(0).Height - margin
    cmdNav(1).Move Me.ScaleWidth - cmdNav(0).Width - margin, Me.ScaleHeight - cmdNav(0).Height - margin
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub

Private Sub gBatches_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    lblTotalSelected.Caption = "Total Selected:   " & Format(TotalSelected, "Currency")
End Sub

Private Sub gBatches_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow Then
    With gBatches
        Call LoadPayments(.TextMatrix(.Row, .ColIndex("AccountingBatch")))
    End With
    End If
End Sub

Private Sub gBatches_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gBatches
    Cancel = .ColKey(Col) <> "Selected"
    End With
End Sub

Private Sub LoadPayments(batch As String)
    Dim s As String
    Dim rs As Recordset
    
    
    s = ""
    s = s & "select" & vbCrLf
    's = s & " Payee," & vbCrLf
    s = s & " PayeeName" & vbCrLf
    s = s & ",PaymentNumber" & vbCrLf
    s = s & ",CheckNumber" & vbCrLf
    s = s & ",PaymentDate" & vbCrLf
    s = s & ",PONumber" & vbCrLf
    s = s & ",sum(NetAmount) AmountPaid" & vbCrLf
    s = s & ",InvoiceNumber" & vbCrLf
    s = s & ",InvoiceDate" & vbCrLf
    s = s & "from Wallet_AccountingAP" & vbCrLf
    s = s & "where DataSourceID=" & DbQuote(Str, DataSourceID) & vbCrLf
    s = s & "and CompanyID=" & DbQuote(Str, CompanyID) & vbCrLf
    s = s & "and AccountingBatch =" & DbQuote(Str, batch) & vbCrLf
    s = s & "group by Payee" & vbCrLf
    s = s & ",PayeeName" & vbCrLf
    s = s & ",PaymentNumber" & vbCrLf
    s = s & ",CheckNumber" & vbCrLf
    s = s & ",PaymentDate" & vbCrLf
    s = s & ",PaymentAmount" & vbCrLf
    s = s & ",PONumber" & vbCrLf
    s = s & ",InvoiceNumber" & vbCrLf
    s = s & ",InvoiceDate" & vbCrLf
    s = s & "order by PayeeName,PaymentNumber,CheckNumber,InvoiceNumber,PONumber" & vbCrLf
    
    gPayments.AutoResize = True
    gPayments.AutoSizeMode = flexAutoSizeColWidth
    Set rs = HFApp.SqlExec(s, dbHomeFront)
    Set gPayments.DataSource = rs
    
        
        
    gPayments.ColDataType(gPayments.ColIndex("PayeeName")) = flexDTString
    gPayments.ColDataType(gPayments.ColIndex("PaymentNumber")) = flexDTString
    gPayments.ColDataType(gPayments.ColIndex("CheckNumber")) = flexDTString
    gPayments.ColDataType(gPayments.ColIndex("PONumber")) = flexDTString
    gPayments.ColDataType(gPayments.ColIndex("InvoiceNumber")) = flexDTString
    gPayments.ColAlignment(gPayments.ColIndex("PayeeName")) = flexAlignLeftCenter
    gPayments.ColAlignment(gPayments.ColIndex("PaymentNumber")) = flexAlignLeftCenter
    gPayments.ColAlignment(gPayments.ColIndex("CheckNumber")) = flexAlignLeftCenter
    gPayments.ColAlignment(gPayments.ColIndex("PONumber")) = flexAlignLeftCenter
    gPayments.ColAlignment(gPayments.ColIndex("InvoiceNumber")) = flexAlignLeftCenter

    gPayments.ColDataType(gPayments.ColIndex("InvoiceDate")) = flexDTDate
    gPayments.ColDataType(gPayments.ColIndex("PaymentDate")) = flexDTDate

    gPayments.ColDataType(gPayments.ColIndex("AmountPaid")) = flexDTCurrency
    gPayments.ColAlignment(gPayments.ColIndex("AmountPaid")) = flexAlignRightCenter


End Sub

Private Sub gBatches_SelChange()
    gBatches.ColSel = gBatches.Col
End Sub

Private Sub Slider_Move()
    Call Form_Resize
End Sub
