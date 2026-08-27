VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form FCreateHBInvoice 
   Caption         =   "Create Invoice"
   ClientHeight    =   6630
   ClientLeft      =   2025
   ClientTop       =   2160
   ClientWidth     =   13050
   Icon            =   "FCreateHBInvoice.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   13050
   Begin VB.TextBox txtNotes 
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   6405
      MaxLength       =   50
      TabIndex        =   5
      Top             =   180
      Width           =   5925
   End
   Begin VB.TextBox txtDraw 
      BorderStyle     =   0  'None
      Height          =   230
      Left            =   1305
      MaxLength       =   50
      TabIndex        =   3
      Top             =   840
      Width           =   1965
   End
   Begin VB.TextBox txtDueDate 
      BorderStyle     =   0  'None
      Height          =   230
      Left            =   1305
      MaxLength       =   50
      TabIndex        =   2
      Top             =   600
      Width           =   1965
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&Save && Print"
      Default         =   -1  'True
      Height          =   375
      Index           =   2
      Left            =   11130
      Picture         =   "FCreateHBInvoice.frx":000C
      TabIndex        =   15
      Top             =   6150
      Width           =   1215
   End
   Begin VB.TextBox txtTerms 
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   1305
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1140
      Width           =   3525
   End
   Begin VB.TextBox txtScope 
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   6405
      MaxLength       =   50
      TabIndex        =   6
      Top             =   1200
      Width           =   5925
   End
   Begin VB.TextBox txtInvoiceDate 
      BorderStyle     =   0  'None
      Height          =   230
      Left            =   1305
      MaxLength       =   50
      TabIndex        =   0
      Top             =   60
      Width           =   1965
   End
   Begin VB.TextBox txtDescription 
      BorderStyle     =   0  'None
      Height          =   230
      Left            =   1305
      MaxLength       =   50
      TabIndex        =   1
      Top             =   360
      Width           =   3525
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   9750
      Picture         =   "FCreateHBInvoice.frx":0596
      TabIndex        =   9
      Top             =   6135
      Width           =   1215
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Height          =   375
      Index           =   0
      Left            =   8490
      Picture         =   "FCreateHBInvoice.frx":0B20
      TabIndex        =   8
      Top             =   6135
      Width           =   1215
   End
   Begin VSFlex8Ctl.VSFlexGrid gContract 
      Height          =   3765
      Left            =   270
      TabIndex        =   7
      Top             =   2250
      Width           =   12195
      _cx             =   1999131655
      _cy             =   1999116785
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
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   20
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FCreateHBInvoice.frx":10AA
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
      ExplorerBar     =   2
      PicturesOver    =   0   'False
      FillStyle       =   1
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   1
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
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Notes"
      Height          =   195
      Left            =   5895
      TabIndex        =   18
      Top             =   150
      Width           =   420
   End
   Begin VB.Image cmdBrowse 
      Height          =   240
      Index           =   0
      Left            =   3285
      Picture         =   "FCreateHBInvoice.frx":1253
      Top             =   600
      Width           =   240
   End
   Begin VB.Image cmdBrowse 
      Height          =   240
      Index           =   1
      Left            =   3285
      Picture         =   "FCreateHBInvoice.frx":139D
      Top             =   60
      Width           =   240
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Draw"
      Height          =   195
      Left            =   870
      TabIndex        =   17
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date"
      Height          =   195
      Left            =   555
      TabIndex        =   16
      Top             =   600
      Width           =   690
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Terms"
      Height          =   195
      Left            =   810
      TabIndex        =   14
      Top             =   1140
      Width           =   435
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scope of work"
      Height          =   195
      Left            =   5280
      TabIndex        =   13
      Top             =   1170
      Width           =   1035
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   135
      TabIndex        =   12
      Top             =   60
      Width           =   1110
   End
   Begin VB.Label Label112 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   270
      TabIndex        =   11
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "FCreateHBInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FCreateInvoice::"
Public EventTraps As Collection

Private mCancel As Boolean
Private mJob As String
Private mCustomer As String

Private mInvoice As Long

Private Sub cmdBrowse_Click(Index As Integer)
    
    If Not cmdBrowse(Index).Enabled Then Exit Sub
    
    Select Case Index
        Case 1:    Call DCalendar.Popup(txtInvoiceDate)
        Case 0:    Call DCalendar.Popup(txtDueDate)
        
    End Select

End Sub

Private Sub cmdNav_Click(Index As Integer)
    Dim s As String
'    Dim f As FRptViewer
    
    Select Case Index
    
        Case 0 'ok
            If SaveData Then
                Unload Me
            End If
            
            
        Case 1 'cancel
            Unload Me
            
            
        Case 2 ' save and print
            If SaveData Then
                s = HFApp.SystemFolder & "System\Reports\Estimating\ARInvoice.rpt"
                'Set f = New FRptViewer
                'Call f.ShowReport(s, True, True, "Invoice", mInvoice)
                
                Dim c As New ZybUtil.Crystal
                Call c.LoadODBCReport(s, HFApp.LoginDSN, HFApp.LoginDBUID, HFApp.LoginDBPWD)
                On Error Resume Next
                Call c.ParameterValue("DivisionID", HFApp.DivisionID)
                Call c.ParameterValue("Invoice", mInvoice)
                On Error GoTo 0
                Call c.PrintPreview("Print Preview")
                
                Unload Me
            End If
            
            
    End Select
    
End Sub

Private Sub Form_Load()
    Call IniGetForm(Me)
    Call IniGetGrid(Me, gContract)
    Call LoadData
End Sub

Private Sub Form_Resize()
On Error Resume Next
Const margin = 120
    txtNotes.Width = Me.ScaleWidth - txtNotes.Left - margin
    txtScope.Width = txtNotes.Width
    gContract.Move margin, gContract.Top, Me.ScaleWidth - 2 * margin, Me.ScaleHeight - gContract.Top - 2 * margin - cmdNav(0).Height
    cmdNav(0).Move Me.ScaleWidth - 3 * (cmdNav(0).Width + margin), Me.ScaleHeight - cmdNav(0).Height - margin
    cmdNav(1).Move Me.ScaleWidth - 2 * (cmdNav(0).Width + margin), Me.ScaleHeight - cmdNav(0).Height - margin
    cmdNav(2).Move Me.ScaleWidth - 1 * (cmdNav(0).Width + margin), Me.ScaleHeight - cmdNav(0).Height - margin
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gContract)
End Sub

Public Function Edit(Job As String) As Boolean
    Dim HoldBackDesc As String
    HoldBackDesc = "" & HFApp.Options.ValueByName("HoldbackDescription")
    If HoldBackDesc = "" Then HoldBackDesc = "Holdback"
    mJob = Job
    mCustomer = "" & HFApp.SqlExec("select ar_customer_deposit from tblcustomers where job_no=" & DbQuote(Str, mJob))(0)
    txtDescription.Text = "Bill " & HoldBackDesc
    mCancel = True
    Me.Caption = "Create " & HoldBackDesc & " Invoice"
    Me.Show vbModal
    Edit = Not mCancel
End Function



Private Function SaveData() As Boolean
On Error GoTo eh
    Dim s As String
    Dim r As Long
    
    
    If Me.txtInvoiceDate.Text = "" Or Me.txtDescription.Text = "" Then
        Call MsgBox("Description and invoice date are required.", vbInformation, App.ProductName)
        SaveData = False
        Exit Function
    End If
        
    
    s = ""
    s = s & "insert into arinvoices(DivisionID,invoicedate,description,Job,arcustomer,terms,scopeofwork,draw,DueDate,notes,RetainageInvoice)" & vbCrLf
    s = s & "values(" & HFApp.DivisionID & "," & DbQuote(Date, txtInvoiceDate.Text) & vbCrLf
    s = s & "      ," & DbQuote(Str, txtDescription.Text) & vbCrLf
    s = s & "      ," & DbQuote(Str, mJob) & vbCrLf
    s = s & "      ," & DbQuote(Str, mCustomer) & vbCrLf
    s = s & "      ," & DbQuote(Str, txtTerms.Text) & vbCrLf
    s = s & "      ," & DbQuote(Str, txtScope.Text) & vbCrLf
    s = s & "      ," & DbQuote(Num, txtDraw.Text) & vbCrLf
    s = s & "      ," & DbQuote(Date, txtDueDate.Text) & vbCrLf
    s = s & "      ," & DbQuote(Str, txtNotes.Text) & ",1)"
    Call HFApp.SqlExec(s, dbHomefront)
    mInvoice = HFApp.SqlIdentity("Arinvoices")
    
    With gContract
        For r = 1 To .Rows - 1
            If .ValueMatrix(r, .ColIndex("HoldBackAmount")) <> 0 Then
                s = ""
                s = s & "insert into arinvoiceitems(DivisionID,invoice,OriginalInvoice,billingitemid,description,holdbackamount,taxgroup,taxrate,RevenueAccount,tax,HoldBackBilled)" & vbCrLf
                s = s & "values(" & HFApp.DivisionID & "," & DbQuote(Num, mInvoice) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(r, .ColIndex("Invoice"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(r, .ColIndex("BillingItemID"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Description"))) & vbCrLf
                s = s & "      ," & DbQuote(Cur, .ValueMatrix(r, .ColIndex("HoldBackAmount")) * -1) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("TaxGroup"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(r, .ColIndex("TaxRate"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("RevenueAccount"))) & vbCrLf
                s = s & "      ," & DbQuote(Num, .ValueMatrix(r, .ColIndex("Tax"))) & vbCrLf
                s = s & "      ," & DbQuote(Cur, .ValueMatrix(r, .ColIndex("HoldBackAmount"))) & vbCrLf
                s = s & "      )"
                Call HFApp.SqlExec(s, dbHomefront)
                    
                s = ""
                s = s & "update billingitems" & vbCrLf
                s = s & "set HoldbackTotalHeld=HoldbackTotalHeld-" & DbQuote(Cur, .ValueMatrix(r, .ColIndex("HoldbackAmount"))) & vbCrLf
                s = s & "where billingitemid=" & DbQuote(Num, .TextMatrix(r, .ColIndex("BillingItemID"))) & vbCrLf
                Call HFApp.SqlExec(s, dbHomefront)
            End If
            mCancel = False
        
        
        Next
    End With
    SaveData = True
Exit Function
eh: Call errHandler(SRCFILE & "SaveData", s)
End Function


Private Sub LoadData()
    
    Dim s As String
    Dim r As Long
    Dim rs As Recordset
    
    s = ""
    s = s & "select m.Job,i.billingitemid,m.ARCustomer,i.Invoice,'Invoice '+cast(i.Invoice as varchar) Description,sum(i.HoldBackAmount-isnull(i.HoldBackBilled,0)) as HoldBackAmount,min(i.TaxRate) TaxRate,i.TaxGroup,sum(round((i.HoldBackAmount-isnull(i.HoldBackBilled,0))*i.TaxRate/100,2)) Tax,i.RevenueAccount,a.Description as AccountDesc" & vbCrLf
    s = s & "from arinvoiceitems i" & vbCrLf
    s = s & "join ArInvoices m on i.invoice = m.invoice" & vbCrLf
    s = s & "join tbljobs j on(m.job=j.job_no and j.DivisionID = m.divisionID)" & vbCrLf
    s = s & "left outer join glaccounts a on(a.DivisionID = i.DivisionID and a.account = isnull(nullif(i.revenueaccount,''),j.revenueaccount))" & vbCrLf
    s = s & "where i.DivisionID = " & HFApp.DivisionID & " and round(isnull(i.HoldbackAmount,0)-isnull(i.HoldbackBilled,0),2)<>0" & vbCrLf
    s = s & "and isnull(m.RetainageInvoice,0) = 0 and m.job=" & DbQuote(Str, mJob) & vbCrLf
    s = s & "Group by m.Job,i.billingitemid,m.ARCustomer,i.Invoice,'Invoice '+cast(i.Invoice as varchar),i.TaxGroup,i.RevenueAccount,a.Description"
    
    Set rs = HFApp.SqlExec(s)
    With gContract
        .Rows = 1
        r = 0
        While Not rs.EOF
            .AddItem ""
            r = r + 1
            .TextMatrix(r, .ColIndex("Invoice")) = "" & rs("Invoice")
            .TextMatrix(r, .ColIndex("BillingItemID")) = "" & rs("BillingItemID")
            
            .TextMatrix(r, .ColIndex("Description")) = "" & rs("Description")
            .TextMatrix(r, .ColIndex("HoldBackAmount")) = Val("" & rs("HoldBackAmount"))
            .TextMatrix(r, .ColIndex("taxgroup")) = "" & rs("taxgroup")
            .TextMatrix(r, .ColIndex("taxrate")) = Val("" & rs("taxrate"))
            .TextMatrix(r, .ColIndex("tax")) = Val("" & rs("tax"))
            
            .TextMatrix(r, .ColIndex("RevenueAccount")) = "" & rs("RevenueAccount")
            .TextMatrix(r, .ColIndex("RevenueAccountDesc")) = "" & rs("AccountDesc")

            rs.MoveNext
        Wend
    End With
    
    txtInvoiceDate.Text = format(VBA.Date(), "mmm d, yyyy")
    
End Sub


Private Sub gContract_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim r As Long
With gContract
If .ColKey(Col) = "HoldBackAmount" Then
    .TextMatrix(Row, .ColIndex("Tax")) = Round(.ValueMatrix(Row, .ColIndex("HoldBackAmount")) * .ValueMatrix(Row, .ColIndex("TaxRate")) / 100, 2)
End If
End With
End Sub

Private Sub gContract_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gContract
        Select Case .ColKey(Col)
            Case "AmtToBill", "PercentComplete", "Remaining", "HoldbackRate", "HoldbackAmount"
                .ComboList = ""
            Case "TaxGroup":        .ComboList = "|..."
            Case "RevenueAccountDesc":  .ComboList = "..."
            Case "RevenueAccount":  .ComboList = "|..."
            Case Else
                Cancel = True
        End Select
    End With
    
End Sub

Private Sub gContract_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    If Button = vbRightButton Then
        If gContract.MouseRow = 0 Then
            Call FMain.ShowColumnMenu(gContract)
        End If
    End If

End Sub

Private Sub gContract_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim s As String
    
    With gContract
        Select Case .ColKey(Col)
        
            Case "TaxGroup"
                s = "SELECT TaxGroup,Description,GroupRate FROM TaxGroups where DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Tax Groups", s, gContract) Then
                    .Cell(flexcpText, .Row, Col, .RowSel, Col) = FPickList.SelectedItem("TaxGroup")
                    .Cell(flexcpText, .Row, .ColIndex("TaxRate"), .RowSel, .ColIndex("TaxRate")) = FPickList.SelectedItem("GroupRate")
                End If
        
            Case "RevenueAccount", "RevenueAccountDesc"
                s = "SELECT Account,Description FROM GLAccounts where DivisionID =" & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Accounts", s, gContract) Then
                    .Cell(flexcpText, .Row, .ColIndex("RevenueAccount"), .RowSel, .ColIndex("RevenueAccount")) = FPickList.SelectedItem("Account")
                    .Cell(flexcpText, .Row, .ColIndex("RevenueAccountDesc"), .RowSel, .ColIndex("RevenueAccountDesc")) = FPickList.SelectedItem("description")
                End If
        
        End Select
    End With
End Sub

Private Sub gContract_SelChange()
   'limit selection to one column
    Static bInHere As Boolean
    If bInHere Then Exit Sub
    bInHere = True
    gContract.ColSel = gContract.Col




    Dim r As Long
    

    
    
    With gContract
        If .Row <> .RowSel Then

        End If
    End With
    
    
    bInHere = False
End Sub

Private Sub gContract_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rs As Recordset
    With gContract
        Select Case .ColKey(Col)
            Case "AmtToBill", "PercentComplete", "Remaining", "HoldbackAmount", "HoldbackRate"
                .EditText = Val(.EditText)
            
            Case "TaxGroup"
                Set rs = HFApp.SqlExec("Select taxgroup,description,grouprate from taxgroups where DivisionID = " & HFApp.DivisionID & " and taxgroup= " & DbQuote(Str, .EditText))
                If rs.EOF Then
                    MsgBox "Invalid Tax Group", vbInformation, App.ProductName
                    Cancel = True
                Else
                    .EditText = "" & rs("TaxGroup")
                    .Cell(flexcpText, .Row, .ColIndex("TaxRate"), .RowSel, .ColIndex("TaxRate")) = Val("" & rs("GroupRate"))
                End If

            
            Case "RevenueAccount"
                Set rs = HFApp.SqlExec("Select account,description from GLAccounts where DivisionID = " & HFApp.DivisionID & " and Account= " & DbQuote(Str, .EditText))
                If rs.EOF Then
                    MsgBox "Invalid Account", vbInformation, App.ProductName
                    Cancel = True
                Else
                    .EditText = Val(.EditText)
                    .Cell(flexcpText, Min(.Row, .RowSel), .ColIndex("RevenueAccountDesc"), Max(.Row, .RowSel), .ColIndex("RevenueAccountDesc")) = "" & rs("description")
                End If
            
            Case Else
                Cancel = True
        End Select
    End With
End Sub



Private Sub txtDraw_GotFocus()
    SelectAll txtDraw
End Sub

Private Sub txtDraw_Validate(Cancel As Boolean)
    txtDraw.Text = Val(txtDraw.Text)
End Sub

Private Sub txtInvoiceDate_GotFocus()
    SelectAll txtInvoiceDate
End Sub
Private Sub txtInvoiceDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 And Shift = 0 Then Call cmdBrowse_Click(1)
End Sub
Private Sub txtInvoiceDate_Validate(Cancel As Boolean)
On Error Resume Next
    If txtInvoiceDate.Text = "" Then Exit Sub
    If Not IsDate(txtInvoiceDate.Text) Then
        Cancel = True
    Else
        txtInvoiceDate.Text = format(txtInvoiceDate.Text, "mmm d, yyyy")
    End If
End Sub

Private Sub txtDueDate_GotFocus()
    SelectAll txtDueDate
End Sub
Private Sub txtDueDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 And Shift = 0 Then Call cmdBrowse_Click(0)
End Sub
Private Sub txtDueDate_Validate(Cancel As Boolean)
On Error Resume Next
    If txtDueDate.Text = "" Then Exit Sub
    If Not IsDate(txtDueDate.Text) Then
        Cancel = True
    Else
        txtDueDate.Text = format(txtDueDate.Text, "mmm d, yyyy")
    End If
End Sub

