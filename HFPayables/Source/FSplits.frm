VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FSplits 
   Caption         =   "Invoice Spliting"
   ClientHeight    =   2730
   ClientLeft      =   1725
   ClientTop       =   2220
   ClientWidth     =   7365
   Icon            =   "FSplits.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   7365
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   5970
      TabIndex        =   2
      Top             =   2190
      Width           =   1215
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   4650
      TabIndex        =   1
      Top             =   2190
      Width           =   1215
   End
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   1455
      Left            =   90
      TabIndex        =   0
      Top             =   600
      Width           =   7095
      _cx             =   12515
      _cy             =   2566
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
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FSplits.frx":000C
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
      ExplorerBar     =   2
      PicturesOver    =   0   'False
      FillStyle       =   0
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
   Begin VB.Label Label1 
      Caption         =   "To split this invoice add the subcontractors to the following list."
      Height          =   465
      Left            =   90
      TabIndex        =   3
      Top             =   180
      Width           =   7095
   End
End
Attribute VB_Name = "FSplits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SRCFILE = "FSplits::"

Private mCancelEdit  As Boolean
Private mInvoiceID   As Long
Private mVendor      As String
Private mInvoice     As String
Private mPretax      As Double
Private mTax         As Double


Public Sub SplitInvoice(vendor As String, invoice As String)
    mVendor = vendor
    mInvoice = invoice
    Me.Show vbModal
End Sub

Private Sub cmdNav_Click(Index As Integer)
On Error GoTo eh

    Dim i As Long
    Dim s As String
    Dim InvoiceID As Long
    
    With gData
    If Index = 0 Then
            
        'for each secondary vendor
        For i = 2 To .Rows - 2
        If .ValueMatrix(i, .ColIndex("Pretax")) <> 0 Then
        
            'create secondary invoice
            s = ""
            s = s & "insert into invoices(DivisionID,vendor,vendortype,vendorname,discount,invoice,job,jobdesc,deptid,status,discountdate,receiveddate,invoicedate,paymentdate,accountingdate,description,ustmp,dstmp,tstmp,errors,source,postingdate,invoicecode1,invoicecode2,approver,holdreason,approvercomments,dateapproved,advance,batchnumber)" & vbCrLf
            s = s & "select DivisionID," & DbQuote(Str, .TextMatrix(i, .ColIndex("Vendor"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("VendorType"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("VendorName"))) & vbCrLf
            s = s & "      ,0" & vbCrLf
            s = s & "      ,invoice,job,jobdesc,deptid,status,discountdate,receiveddate,invoicedate,paymentdate,accountingdate,description" & vbCrLf
            s = s & "      ,ustmp,dstmp,tstmp,errors,source,postingdate,invoicecode1,invoicecode2,approver,holdreason,approvercomments,dateapproved,advance,batchnumber" & vbCrLf
            s = s & " from invoices " & vbCrLf
            s = s & " where DivisionID =" & HFApp.DivisionID & " and vendor=" & DbQuote(Str, mVendor) & vbCrLf
            s = s & "   and invoice=" & DbQuote(Str, mInvoice) & vbCrLf
            Call HFApp.SqlExec(s, dbHomeFront)
            InvoiceID = HFApp.SqlIdentity("Invoices", dbHomeFront)
            
            'create secondary invoice items
            s = ""
            s = s & "insert into invoiceitems(DivisionID,InvoiceID,Vendor,Invoice,Pretax,Tax,taxgroup,taxgroupdesc,taxrate,job,jobdesc,extra,extradesc,phase,phasedesc,category,categorydesc,debitaccount,debitaccountdesc,description,retainagerate,taxretainagerate,ustmp,dstmp,tstmp)" & vbCrLf
            s = s & "select DivisionID," & DbQuote(Num, InvoiceID) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Vendor"))) & vbCrLf
            s = s & "      ,invoice" & vbCrLf
            s = s & "      ,round(pretax * " & .ValueMatrix(i, .ColIndex("Percent")) & " / 100,2)" & vbCrLf
            s = s & "      ,round(pretax * " & .ValueMatrix(i, .ColIndex("Percent")) & " / 100 * " & Val(.TextMatrix(i, .ColIndex("TaxRate"))) & " / 100,2)" & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("TaxGroup"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("TaxGroupDesc"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(i, .ColIndex("TaxRate"))) & vbCrLf
            s = s & "      ,job,jobdesc" & vbCrLf
            s = s & "      ,extra,extradesc" & vbCrLf
            s = s & "      ,phase,phasedesc" & vbCrLf
            s = s & "      ,category,categorydesc" & vbCrLf
            s = s & "      ,debitaccount,debitaccountdesc" & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Percent")) & "% split from " & mVendor) & vbCrLf
            s = s & "      ,retainagerate,taxretainagerate" & vbCrLf
            s = s & "      ,ustmp,dstmp,tstmp" & vbCrLf
            s = s & "  from invoiceitems" & vbCrLf
            s = s & " where DivisionID =" & HFApp.DivisionID & " and vendor=" & DbQuote(Str, mVendor) & vbCrLf
            s = s & "   and invoice=" & DbQuote(Str, mInvoice) & vbCrLf
            Call HFApp.SqlExec(s, dbHomeFront)
                
            'take wcb from secondary invoice
            Call RecalcInvoiceTotal(mInvoiceID)
                    
            'credit primary invoice
            s = ""
            s = s & "insert into invoiceitems(DivisionID,InvoiceID,Vendor,Invoice,Pretax,Tax,taxgroup,taxgroupdesc,taxrate,job,jobdesc,extra,extradesc,phase,phasedesc,category,categorydesc,debitaccount,debitaccountdesc,description,retainagerate,taxretainagerate,ustmp,dstmp,tstmp)" & vbCrLf
            s = s & "select DivisionID,InvoiceID" & vbCrLf
            s = s & "      ,Vendor" & vbCrLf
            s = s & "      ,Invoice" & vbCrLf
            s = s & "      ,round(pretax * -1 * " & .ValueMatrix(i, .ColIndex("Percent")) & " / 100,2)" & vbCrLf
            s = s & "      ,round(pretax * -1 * " & .ValueMatrix(i, .ColIndex("Percent")) & " / 100 * taxrate / 100,2)" & vbCrLf
            s = s & "      ,taxgroup,taxgroupdesc,taxrate" & vbCrLf
            s = s & "      ,job,jobdesc" & vbCrLf
            s = s & "      ,extra,extradesc" & vbCrLf
            s = s & "      ,phase,phasedesc" & vbCrLf
            s = s & "      ,category,categorydesc" & vbCrLf
            s = s & "      ,debitaccount,debitaccountdesc" & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(i, .ColIndex("Percent")) & "% split with " & .TextMatrix(i, .ColIndex("Vendor"))) & vbCrLf
            s = s & "      ,retainagerate,taxretainagerate" & vbCrLf
            s = s & "      ,ustmp,dstmp,tstmp" & vbCrLf
            s = s & "  from invoiceitems" & vbCrLf
            s = s & " where DivisionID =" & HFApp.DivisionID & " and vendor=" & DbQuote(Str, mVendor) & vbCrLf
            s = s & "   and invoice=" & DbQuote(Str, mInvoice) & vbCrLf
            Call HFApp.SqlExec(s, dbHomeFront)
            
        End If
        Next
        
            
    End If
    End With
    Unload Me
Exit Sub
eh: Call ErrHandler(SRCFILE & "cmdNav_Click")
End Sub

Private Sub Form_Activate()
Call gData.SetFocus
Call gData.Select(2, 1)
End Sub

Private Sub Form_Load()
On Error GoTo eh
    Dim s As String
    Dim rs As Recordset
    
    Call IniGetGrid(Me, gData)
    Call IniGetForm(Me)
    
    s = "select invoiceid,vendorname,VendorType,pretax,tax from invoices where DivisionID = " & HFApp.DivisionID & " and vendor=" & DbQuote(Str, mVendor) & " and invoice=" & DbQuote(Str, mInvoice)
    Set rs = HFApp.SqlExec(s, dbHomeFront)
    
    mInvoiceID = Val("" & rs("InvoiceID"))
    mPretax = Val("" & rs("Pretax"))
    mTax = Val("" & rs("Tax"))
    
    With gData
        .Cell(flexcpBackColor, 1, 0, 1, .Cols - 1) = vbButtonFace
        .TextMatrix(1, .ColIndex("Vendor")) = mVendor
        .TextMatrix(1, .ColIndex("VendorName")) = "" & rs("VendorName")
        .TextMatrix(1, .ColIndex("VendorType")) = "" & rs("VendorType")
        .TextMatrix(1, .ColIndex("Percent")) = 100
        .TextMatrix(1, .ColIndex("Pretax")) = mPretax
        .TextMatrix(1, .ColIndex("Tax")) = mTax
    End With
    
Exit Sub
eh: Call ErrHandler(SRCFILE & "cmdNav_Click")
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Const margin = 60
    gData.Move margin, gData.Top, Me.ScaleWidth - 2 * margin, Me.ScaleHeight - gData.Top - 2 * margin - cmdNav(0).Height
    cmdNav(0).Move Me.ScaleWidth - 2 * cmdNav(0).Width - 2 * margin, Me.ScaleHeight - cmdNav(0).Height - margin
    cmdNav(1).Move Me.ScaleWidth - cmdNav(0).Width - margin, Me.ScaleHeight - cmdNav(0).Height - margin
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutGrid(Me, gData)
    Call IniPutForm(Me)
End Sub


Private Sub gData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gData
        .ComboList = ""
        Select Case .ColKey(Col)
            Case "Percent":    .ComboList = "|50|40|33.33|25|20|10"
            Case "Vendor":     .ComboList = "|..."
            Case "Name":       .ComboList = "..."
            Case "TaxGroup":   .ComboList = "|..."
            Case "TaxRate":    Cancel = True
            Case "Pretax":
            Case "Tax":
        End Select
        If Row = 1 Then Cancel = True
    End With
End Sub

Private Sub gData_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim s As String
    With gData
        Select Case .ColKey(Col)
            Case "Vendor"
                s = "SELECT Vendor_id " & Quote(App.Options(Caption_Vendor)) & ", Vendor_Name Name FROM tblvendors where DivisionID =" & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomeFront), "", s) Then
                    .TextMatrix(Row, Col) = FPickList.SelectedItem(1)
                    Call gData_AfterEdit(Row, Col)
                End If
                
            Case "Name"
                s = "SELECT Vendor_Name Name, Vendor_id " & Quote(App.Options(Caption_Vendor)) & " FROM tblvendors where DivisionID =" & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomeFront), "", s) Then
                    .TextMatrix(Row, .ColIndex("Vendor")) = FPickList.SelectedItem(2)
                    Call gData_AfterEdit(Row, .ColIndex("Vendor"))
                End If
                
            Case "TaxGroup"
                s = "SELECT taxgroup " & Quote(App.Options(Caption_TaxGroup)) & ",Description,GroupRate TaxRate FROM taxgroups where DivisionID =" & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomeFront), "", s) Then
                    .TextMatrix(Row, Col) = FPickList.SelectedItem(1)
                    Call gData_AfterEdit(Row, Col)
                End If
                
        End Select
        
        If mCancelEdit Then
            Call .FinishEditing(True)
            .TextMatrix(Row, Col) = ""
        End If
    End With

End Sub

Private Sub gData_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim s As String
    Dim rs As Recordset
    
    With gData
        mCancelEdit = False
        
        If Row = .Rows - 1 Then
            .AddItem ""
        End If
    
    
        Select Case .ColKey(Col)
            
            Case "Vendor":
                s = ""
                s = s & "SELECT v.vendor_id,v.vendor_name,v.tradetype,t.description,t.taxgroup,t.grouprate" & vbCrLf
                s = s & "  FROM tblvendors v LEFT OUTER JOIN taxgroups t ON v.DivisionID = t.DivisionID and v.materialtaxgroup=t.taxgroup" & vbCrLf
                s = s & " WHERE v.DivisionID =" & HFApp.DivisionID & " and v.vendor_id=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("Vendor"))) & vbCrLf
                Set rs = HFApp.SqlExec(s, dbHomeFront)
                If rs.EOF Then
                    mCancelEdit = True
                    MsgBox App.Options(Caption_Vendor) & " not found", vbExclamation, App.ProductName
                Else
                    .TextMatrix(Row, .ColIndex("Vendor")) = Trim("" & rs("vendor_id"))
                    .TextMatrix(Row, .ColIndex("VendorName")) = Trim("" & rs("vendor_name"))
                    .TextMatrix(Row, .ColIndex("VendorType")) = Trim("" & rs("tradetype"))
                    .TextMatrix(Row, .ColIndex("TaxGroup")) = Trim("" & rs("taxgroup"))
                    If .TextMatrix(Row, .ColIndex("TaxGroup")) <> "" Then
                        .TextMatrix(Row, .ColIndex("TaxGroupDesc")) = Trim("" & rs("description"))
                        .TextMatrix(Row, .ColIndex("TaxRate")) = Val("" & rs("grouprate"))
                    Else
                        .TextMatrix(Row, .ColIndex("TaxGroupDesc")) = App.Options(DefaultTaxGroupDesc)
                        .TextMatrix(Row, .ColIndex("TaxGroup")) = App.Options(DefaultTaxGroup)
                        .TextMatrix(Row, .ColIndex("TaxRate")) = App.Options(DefaultTaxGroupRate)
                    End If
                    .TextMatrix(Row, .ColIndex("Tax")) = .ValueMatrix(Row, .ColIndex("PreTax")) * .ValueMatrix(Row, .ColIndex("TaxRate")) / 100
                End If
            
            Case "TaxGroup"
                s = "select taxgroup,grouprate,description FROM taxgroups where DivisionID =" & HFApp.DivisionID & " and taxgroup=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("TaxGroup")))
                Set rs = HFApp.SqlExec(s, dbHomeFront)
                If rs.EOF Then
                    mCancelEdit = True
                    MsgBox App.Options(Caption_TaxGroup) & " not found", vbExclamation, App.ProductName
                Else
                    .TextMatrix(Row, .ColIndex("TaxGroup")) = Trim("" & rs("taxgroup"))
                    .TextMatrix(Row, .ColIndex("TaxGroupDesc")) = Trim("" & rs("description"))
                    .TextMatrix(Row, .ColIndex("TaxRate")) = Val("" & rs("grouprate"))
                    .TextMatrix(Row, .ColIndex("Tax")) = .ValueMatrix(Row, .ColIndex("PreTax")) * Val("" & rs("grouprate")) / 100
                End If
                
            Case "Percent"
                .TextMatrix(Row, .ColIndex("Percent")) = Val(.TextMatrix(Row, .ColIndex("Percent")))
                .TextMatrix(Row, .ColIndex("Pretax")) = Round(mPretax * .ValueMatrix(Row, .ColIndex("Percent")) / 100, 2)
                .TextMatrix(Row, .ColIndex("Tax")) = Round(.ValueMatrix(Row, .ColIndex("PreTax")) * .ValueMatrix(Row, .ColIndex("TaxRate")) / 100, 2)
            
            Case "Pretax"
                .TextMatrix(Row, .ColIndex("Pretax")) = Val(.TextMatrix(Row, .ColIndex("Pretax")))
                .TextMatrix(Row, .ColIndex("Percent")) = Round(Val(.TextMatrix(Row, .ColIndex("Pretax"))) / mPretax * 100, 2)
                .TextMatrix(Row, .ColIndex("Tax")) = Round(.ValueMatrix(Row, .ColIndex("PreTax")) * .ValueMatrix(Row, .ColIndex("TaxRate")) / 100, 2)
            
            Case "Tax"
                .TextMatrix(Row, .ColIndex("Tax")) = Round(.ValueMatrix(Row, .ColIndex("Tax")), 2)
        
        End Select
    
    End With
    
    
    Call Recalc
    
    If mCancelEdit Then Call gData.EditCell
    
End Sub
Private Sub gData_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    Cancel = mCancelEdit
End Sub

Private Sub Recalc()
    Dim i As Long
    Dim pretax As Double
    Dim tax As Double
    With gData
        pretax = 0
        tax = 0
        For i = 2 To .Rows - 2
            pretax = pretax + .ValueMatrix(i, .ColIndex("Pretax"))
            tax = tax + .ValueMatrix(i, .ColIndex("Tax"))
        Next
        .TextMatrix(1, .ColIndex("Pretax")) = mPretax - pretax
        .TextMatrix(1, .ColIndex("Tax")) = mTax - tax
        If mPretax = 0 Then
            .TextMatrix(1, .ColIndex("Percent")) = 0
        Else
            .TextMatrix(1, .ColIndex("Percent")) = Round(.ValueMatrix(1, .ColIndex("Pretax")) / mPretax * 100, 2)
        End If
    End With
End Sub

Private Sub gData_KeyDown(KeyCode As Integer, Shift As Integer)
    With gData
    Select Case KeyCode
        Case vbKeyReturn
            KeyCode = 0
            If .Col = .Cols - 1 Then
                Call .Select(.Row + 1, 0)
            Else
                Call .Select(.Row, .Col + 1)
            End If
    End Select
    End With
End Sub
