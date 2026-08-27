Attribute VB_Name = "MIntacct"
Option Explicit
Option Compare Text
Private Const SRCFILE = "MIntacct::"
Const IntacctDateFormat = "mm\/dd\/yyyy"
    
Private Function POsArePosted(WhereClause As String) As Boolean
    Dim s As String
    Dim rs As Recordset
    
    
POsArePosted = True
Exit Function
'dont do any of this...  Some intacct setups don't require PO and clients like it that way.


    s = ""
    s = s & "select distinct pm.ponumber" & vbCrLf
    s = s & "from invoices " & vbCrLf
    s = s & "join tblvendors v on invoices.divisionid=v.divisionid and invoices.vendor=v.vendor_id" & vbCrLf
    s = s & "left join invoiceitems d on invoices.invoiceid=d.invoiceid" & vbCrLf
    s = s & "join pomaster pm on d.divisionid=pm.divisionid and d.commitment=pm.ponumber" & vbCrLf
    s = s & "where invoices.divisionid=" & DbQuote(num, HFApp.DivisionID) & vbCrLf
    s = s & "and nullif(pm.ExternalID,'') is null" & vbCrLf
    s = s & WhereClause & vbCrLf
    Set rs = HFApp.SqlExec(s)
    
    s = ""
    While Not rs.EOF
        s = s & vbCrLf & rs("PONumber")
        rs.MoveNext
    Wend
    
    If s = "" Then
        POsArePosted = True
    Else
        POsArePosted = False
        MsgBox "Unable to post the selected invoices. These POs have not been posted to Intacct." & vbCrLf & s, vbInformation, App.ProductName
    End If
    
End Function
    
Public Sub SendInvoicesToIntacct(WhereClause As String)
On Error GoTo eh
    Dim req As String
    Dim res As String
    Dim s As String
    Dim rs As Recordset
    Dim LastInvoice As String
    Dim ThisInvoice As String
    Dim Form1099 As String
    Dim Box1099 As String
    
    Dim TranType As String
    Dim OneTimeID As String
    Dim OneTimeUOM As String
    TranType = HFApp.Options.ValueByName("IntacctInvoiceType")
    OneTimeID = HFApp.Options.ValueByName("IntacctOneTimeItemID")
    OneTimeUOM = HFApp.Options.ValueByName("IntacctOneTimeItemUOM")

    Dim PostExtra As Boolean
    PostExtra = HFApp.Options.ValueByName("PostJCExtraAsSubJob") = "True"

    'get sourcelinekeys from intacct
    Call ReadPOItemIDs(WhereClause)
    
    If Not POsArePosted(WhereClause) Then Exit Sub

    
    Dim Intacct As New IntacctWrapper.IntacctWrapper
    Call Intacct.OpenMessage(HFApp.Options.ValueByName("IntacctCompanyID"), HFApp.Options.ValueByName("IntacctUID"), HFApp.Options.ValueByName("IntacctPWD"), HFApp.Options.ValueByName("IntacctEntity"))
    
    
    Dim RefFieldName As String
    Select Case HFApp.Options.ValueByName("IntacctInvReferenceFld")
        Case "Job Number":           RefFieldName = ",j.Job_no"
        Case "Invoice Description":  RefFieldName = ",invoices.Description"
        Case "PO Number":            RefFieldName = ",dbo.Payables_GetInvoicePONumber(invoices.invoiceid)"
        Case Else:                   RefFieldName = ",j.Description"
    End Select
        
    
    'post amount as pretax+jctax
    s = ""
    s = s & "select --non-polines" & vbCrLf
    s = s & " invoices.invoiceid,1 sortorder,d.itemid" & vbCrLf
    s = s & ",invoices.Vendor" & vbCrLf
    s = s & ",null PONumber" & vbCrLf
    s = s & RefFieldName & " Reference" & vbCrLf
    s = s & ",invoices.Invoice " & vbCrLf
    s = s & ",invoices.InvoiceDate" & vbCrLf
    s = s & ",invoices.AccountingDate" & vbCrLf
    s = s & ",invoices.PaymentDate" & vbCrLf
    s = s & ",isnull(isnull(nullif(cc.intacctitemid,''),nullif(c.intacctitemid,''))," & DbQuote(Str, OneTimeID) & ") IntacctItemID" & vbCrLf
    s = s & ",left( case when isnull(nullif(d.InvoicedQuantity,0),1)<>1 then rtrim(ltrim(d.Description)) + ' (qty ' + cast(isnull(nullif(d.InvoicedQuantity,0),1) as varchar(500)) + ')' else  d.Description  end,75) ItemDescription" & vbCrLf
    s = s & ",isnull(nullif(d.InvoicedQuantity,0),1) Qty" & vbCrLf
    s = s & ",(isnull(d.jctax,0)+isnull(d.Pretax,0)) / isnull(nullif(d.InvoicedQuantity,0),1) UnitPrice" & vbCrLf
    s = s & ",isnull(isnull(nullif(cc.intacctuom,''),nullif(c.intacctuom,'')),'Each') UOM" & vbCrLf
    s = s & ",j.gl_prefix      " & vbCrLf
    s = s & ",j.IntacctDepartment" & vbCrLf
    s = s & ",d.Job               " & vbCrLf
    s = s & ",d.Extra               " & vbCrLf
    s = s & ",d.Phase CostCode  --taskid" & vbCrLf
    s = s & ",d.Category        --costtypeid  " & vbCrLf
    s = s & ",null SourceLineKey --intacct recordno of poitem" & vbCrLf
    s = s & ",v.Form1099,v.Box1099" & vbCrLf
    s = s & ",j.arcustomer" & vbCrLf
    s = s & ",d.billable" & vbCrLf
    s = s & ",d.Retainage RetainageAmount" & vbCrLf
    s = s & ",iif(upcc.job is null,0,1) IsNewCostCode" & vbCrLf
    s = s & ",iif(upc.job is null,0,1) IsNewCategory" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "from invoices " & vbCrLf
    s = s & "join tblvendors v on invoices.divisionid=v.divisionid and invoices.vendor=v.vendor_id" & vbCrLf
    s = s & "left join invoiceitems d on invoices.invoiceid=d.invoiceid" & vbCrLf
    s = s & "left join tbljobs j on d.divisionid=j.divisionid and d.job=j.job_no" & vbCrLf
    s = s & "join StandardCostCodes cc on d.divisionid=cc.divisionid and d.phase=cc.costcode" & vbCrLf
    s = s & "join Standardcategories c on d.divisionid=c.divisionid and d.category=c.category" & vbCrLf
    s = s & "left join UnPostedJobCostCodes upcc on j.divisionid=upcc.divisionid and j.job_no=upcc.job and cc.costcode=upcc.jccostcode" & vbCrLf
    s = s & "left join UnPostedJobCostCategories upc on j.divisionid=upc.divisionid and j.job_no=upc.job and cc.costcode=upc.jccostcode and c.category=upc.jccategory" & vbCrLf
    s = s & "left join glaccounts a on d.divisionid=a.divisionid and d.debitaccount=a.account" & vbCrLf
    s = s & "where isnull(d.Commitment,'')=''" & vbCrLf
    s = s & "and invoices.divisionid=" & DbQuote(num, HFApp.DivisionID) & vbCrLf
    s = s & WhereClause & vbCrLf
    
    s = s & "" & vbCrLf
    s = s & "union all" & vbCrLf
    s = s & "" & vbCrLf
    
    s = s & "select --reglines" & vbCrLf
    s = s & " invoices.invoiceid,1 sortorder,d.itemid" & vbCrLf
    s = s & ",invoices.Vendor" & vbCrLf
    s = s & ",pm.IntacctTransactionType+'-'+pm.PONumber PONumber" & vbCrLf
    s = s & RefFieldName & " Reference" & vbCrLf
    s = s & ",invoices.Invoice " & vbCrLf
    s = s & ",invoices.InvoiceDate" & vbCrLf
    s = s & ",invoices.AccountingDate" & vbCrLf
    s = s & ",invoices.PaymentDate" & vbCrLf
    s = s & ",isnull(isnull(nullif(cc.intacctitemid,''),nullif(c.intacctitemid,''))," & DbQuote(Str, OneTimeID) & ") IntacctItemID" & vbCrLf
    s = s & ",left( case when isnull(nullif(d.InvoicedQuantity,0),1)<>1 then rtrim(ltrim(d.Description)) + ' (qty ' + cast(isnull(nullif(d.InvoicedQuantity,0),1) as varchar(500)) + ')' else  d.Description  end,75) ItemDescription" & vbCrLf
    s = s & ",isnull(nullif(d.InvoicedQuantity,0),1) Qty" & vbCrLf
    s = s & ",(isnull(d.jctax,0)+isnull(d.Pretax,0)) / isnull(nullif(d.InvoicedQuantity,0),1) UnitPrice" & vbCrLf
    s = s & ",isnull(isnull(nullif(cc.intacctuom,''),nullif(c.intacctuom,''))," & DbQuote(Str, OneTimeUOM) & ") UOM" & vbCrLf
    s = s & ",j.gl_prefix      " & vbCrLf
    s = s & ",j.IntacctDepartment" & vbCrLf
    s = s & ",d.Job" & vbCrLf
    s = s & ",d.Extra" & vbCrLf
    s = s & ",d.Phase CostCode  --taskid" & vbCrLf
    s = s & ",d.Category        --costtypeid  " & vbCrLf
    s = s & ",i.ExternalID SourceLineKey --intacct recordno of poitem" & vbCrLf
    s = s & ",v.Form1099,v.Box1099" & vbCrLf
    s = s & ",j.arcustomer" & vbCrLf
    s = s & ",d.billable" & vbCrLf
    s = s & ",d.Retainage RetainageAmount" & vbCrLf
    s = s & ",iif(upcc.job is null,0,1) IsNewCostCode" & vbCrLf
    s = s & ",iif(upc.job is null,0,1) IsNewCategory" & vbCrLf
    s = s & "from invoices " & vbCrLf
    s = s & "join tblvendors v on invoices.divisionid=v.divisionid and invoices.vendor=v.vendor_id" & vbCrLf
    s = s & "left join invoiceitems d on invoices.invoiceid=d.invoiceid" & vbCrLf
    s = s & "left join tbljobs j on d.divisionid=j.divisionid and d.job=j.job_no" & vbCrLf
    s = s & "left join pomaster pm on d.divisionid=pm.divisionid and d.commitment=pm.ponumber" & vbCrLf
    s = s & "join jcpodetails_reglines pi on d.divisionid=pi.divisionid and d.commitment=pi.ponumber and d.commitmentitem=pi.linenumber" & vbCrLf
    s = s & "join poitems  i on(pi.poitemseq=i.itemseq)" & vbCrLf
    s = s & "join StandardCostCodes cc on d.divisionid=cc.divisionid and d.phase=cc.costcode" & vbCrLf
    s = s & "join Standardcategories c on d.divisionid=c.divisionid and d.category=c.category" & vbCrLf
    s = s & "left join UnPostedJobCostCodes upcc on j.divisionid=upcc.divisionid and j.job_no=upcc.job and cc.costcode=upcc.jccostcode" & vbCrLf
    s = s & "left join UnPostedJobCostCategories upc on j.divisionid=upc.divisionid and j.job_no=upc.job and cc.costcode=upc.jccostcode and c.category=upc.jccategory" & vbCrLf
    s = s & "left join glaccounts a on d.divisionid=a.divisionid and d.debitaccount=a.account" & vbCrLf
    s = s & "where invoices.divisionid=" & DbQuote(num, HFApp.DivisionID) & vbCrLf
    s = s & WhereClause & vbCrLf
    
    s = s & "" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "" & vbCrLf
    
    s = s & "select --varlines" & vbCrLf
    s = s & " invoices.invoiceid,1 sortorder,d.itemid" & vbCrLf
    s = s & ",invoices.Vendor" & vbCrLf
    s = s & ",pm.IntacctTransactionType+'-'+pm.PONumber PONumber" & vbCrLf
    s = s & RefFieldName & " Reference" & vbCrLf
    s = s & ",invoices.Invoice " & vbCrLf
    s = s & ",invoices.InvoiceDate" & vbCrLf
    s = s & ",invoices.AccountingDate" & vbCrLf
    s = s & ",invoices.PaymentDate" & vbCrLf
    s = s & ",isnull(isnull(nullif(cc.intacctitemid,''),nullif(c.intacctitemid,''))," & DbQuote(Str, OneTimeID) & ") IntacctItemID" & vbCrLf
    s = s & ",left( case when isnull(nullif(d.InvoicedQuantity,0),1)<>1 then rtrim(ltrim(d.Description)) + ' (qty ' + cast(isnull(nullif(d.InvoicedQuantity,0),1) as varchar(500)) + ')' else  d.Description  end,75) ItemDescription" & vbCrLf
    s = s & ",isnull(nullif(d.InvoicedQuantity,0),1) Qty" & vbCrLf
    s = s & ",(isnull(d.jctax,0)+isnull(d.Pretax,0)) UnitPrice" & vbCrLf
    s = s & ",isnull(isnull(nullif(cc.intacctuom,''),nullif(c.intacctuom,''))," & DbQuote(Str, OneTimeUOM) & ") UOM" & vbCrLf
    s = s & ",j.gl_prefix      " & vbCrLf
    s = s & ",j.IntacctDepartment" & vbCrLf
    s = s & ",d.Job               " & vbCrLf
    s = s & ",d.Extra" & vbCrLf
    s = s & ",d.Phase CostCode  --taskid" & vbCrLf
    s = s & ",d.Category        --costtypeid  " & vbCrLf
    s = s & ",i.VarianceExternalID SourceLineKey --intacct recordno of poitem" & vbCrLf
    s = s & ",v.Form1099,v.Box1099" & vbCrLf
    s = s & ",j.arcustomer" & vbCrLf
    s = s & ",d.billable" & vbCrLf
    s = s & ",d.Retainage RetainageAmount" & vbCrLf
    s = s & ",iif(upcc.job is null,0,1) IsNewCostCode" & vbCrLf
    s = s & ",iif(upc.job is null,0,1) IsNewCategory" & vbCrLf
    s = s & "from invoices " & vbCrLf
    s = s & "join tblvendors v on invoices.divisionid=v.divisionid and invoices.vendor=v.vendor_id" & vbCrLf
    s = s & "left join invoiceitems d on invoices.invoiceid=d.invoiceid" & vbCrLf
    s = s & "left join tbljobs j on d.divisionid=j.divisionid and d.job=j.job_no" & vbCrLf
    s = s & "left join pomaster pm on d.divisionid=pm.divisionid and d.commitment=pm.ponumber" & vbCrLf
    s = s & "join jcpodetails_varlines pi on d.divisionid=pi.divisionid and d.commitment=pi.ponumber and d.commitmentitem=pi.linenumber" & vbCrLf
    s = s & "join poitems  i on(pi.poitemseq=i.itemseq)" & vbCrLf
    s = s & "join StandardCostCodes cc on d.divisionid=cc.divisionid and d.phase=cc.costcode" & vbCrLf
    s = s & "join Standardcategories c on d.divisionid=c.divisionid and d.category=c.category" & vbCrLf
    s = s & "left join UnPostedJobCostCodes upcc on j.divisionid=upcc.divisionid and j.job_no=upcc.job and cc.costcode=upcc.jccostcode" & vbCrLf
    s = s & "left join UnPostedJobCostCategories upc on j.divisionid=upc.divisionid and j.job_no=upc.job and cc.costcode=upc.jccostcode and c.category=upc.jccategory" & vbCrLf
    s = s & "left join glaccounts a on d.divisionid=a.divisionid and d.debitaccount=a.account" & vbCrLf
    s = s & "where invoices.divisionid=" & DbQuote(num, HFApp.DivisionID) & vbCrLf
    s = s & WhereClause & vbCrLf
    
    s = s & "" & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "" & vbCrLf
    
    s = s & "select --njctaxes" & vbCrLf
    s = s & " invoices.InvoiceID,3 sortorder,999999999999999 itemid" & vbCrLf
    s = s & ",invoices.Vendor" & vbCrLf
    s = s & ",pm.IntacctTransactionType+'-'+pm.PONumber PONumber" & vbCrLf
    s = s & RefFieldName & " Reference" & vbCrLf
    s = s & ",invoices.Invoice" & vbCrLf
    s = s & ",invoices.InvoiceDate" & vbCrLf
    s = s & ",invoices.AccountingDate" & vbCrLf
    s = s & ",invoices.PaymentDate" & vbCrLf
    s = s & ",o.optionvalue IntacctItemID" & vbCrLf
    s = s & ",'tax (non-costed)' ItemDescription" & vbCrLf
    s = s & ",1 Qty" & vbCrLf
    s = s & ",sum(d.njctax) UnitPrice" & vbCrLf
    s = s & ",ou.optionvalue UOM" & vbCrLf
    s = s & ",j.gl_prefix" & vbCrLf
    s = s & ",j.IntacctDepartment" & vbCrLf
    s = s & ",'' Job" & vbCrLf
    s = s & ",'' Extra" & vbCrLf
    s = s & ",'' CostCode" & vbCrLf
    s = s & ",'' Category" & vbCrLf
    s = s & ",pm.IntacctNJCTaxLineID SourceLineKey --intacct recordno of poitem" & vbCrLf
    s = s & ",v.Form1099,v.Box1099" & vbCrLf
    s = s & ",j.arcustomer" & vbCrLf
    s = s & ",d.billable" & vbCrLf
    ' tax retainage is prorated on all line items
    s = s & ",round(sum(d.njctax)*sum(d.Retainage)/sum(d.pretax+d.jctax),2) RetainageAmount" & vbCrLf
    s = s & ",0 IsNewCostCode" & vbCrLf
    s = s & ",0 IsNewCategory" & vbCrLf
    s = s & "from invoices" & vbCrLf
    s = s & "join tblvendors v on invoices.divisionid=v.divisionid and invoices.vendor=v.vendor_id" & vbCrLf
    s = s & "left join invoiceitems d on invoices.invoiceid=d.invoiceid" & vbCrLf
    s = s & "left join tbljobs j on d.divisionid=j.divisionid and d.job=j.job_no" & vbCrLf
    s = s & "left join pomaster pm on d.divisionid=pm.divisionid and d.commitment=pm.ponumber" & vbCrLf
    s = s & "left join poitems pi on d.divisionid=pi.divisionid and d.commitment=pi.ponumber and d.commitmentitem=pi.linenumber" & vbCrLf
    s = s & "left join standardcostcodes c on d.divisionid=c.divisionid and d.phase=c.costcode " & vbCrLf
    s = s & "left join appoptions o on invoices.divisionid=o.divisionid and o.optionname='IntacctGstItemID'" & vbCrLf
    s = s & "left join appoptions ou on invoices.divisionid=ou.divisionid and ou.optionname='IntacctGstItemUOM'" & vbCrLf
    s = s & "where invoices.divisionid=" & DbQuote(num, HFApp.DivisionID) & vbCrLf
    s = s & WhereClause & vbCrLf
    s = s & "and isnull(d.nJCTax,0)<>0" & vbCrLf
    s = s & "group by invoices.InvoiceID,invoices.Vendor,pm.IntacctTransactionType,j.gl_prefix,j.IntacctDepartment,pm.PONumber" & RefFieldName & vbCrLf
    s = s & ",invoices.Invoice,invoices.InvoiceDate,invoices.paymentdate,invoices.accountingdate,pm.IntacctNJCTaxLineID,o.optionvalue,d.DebitAccount,ou.optionvalue,v.Form1099,v.Box1099,j.arcustomer,d.billable" & vbCrLf
    s = s & "order by 1,2,3" & vbCrLf
    Set rs = HFApp.SqlExec(s)

    While Not rs.EOF
        ThisInvoice = "" & rs("InvoiceID")
        If ThisInvoice <> LastInvoice Then
            If LastInvoice <> "" Then Call Intacct.ClosePOTran
            Call Intacct.OpenPOTran(TranType, Format("" & rs("invoicedate"), IntacctDateFormat), Format("" & rs("accountingdate"), IntacctDateFormat), Format("" & rs("paymentdate"), IntacctDateFormat), _
                                              "" & rs("vendor"), "" & rs("PONumber"), "" & rs("Invoice"), "" & rs("Reference"))
            LastInvoice = ThisInvoice
            Form1099 = "" & rs("Form1099")
            Box1099 = "" & rs("Box1099")
        End If

        Call Intacct.AddInvoiceTranItem("" & rs("IntacctItemID"), "" & rs("ItemDescription"), "" & rs("Qty"), "" & rs("UOM"), Round(Val("" & rs("UnitPrice")), 10), "" & rs("gl_prefix"), "" & rs("Job"), IIf(PostExtra, "" & rs("Extra"), ""), _
                                        "" & rs("IsNewCostCode"), "" & rs("CostCode"), "" & rs("IsNewCategory"), "" & rs("Category"), IIf(Val("" & rs("SourceLineKey")) = 0, "", "" & rs("SourceLineKey")), "" & rs("IntacctDepartment"), Form1099, Box1099, "" & rs("arcustomer"), IIf("" & rs("billable"), "true", "false"), "" & rs("RetainageAmount"))
        rs.MoveNext
    Wend

    If ThisInvoice = "" Then
        'nothing in batch??
    
    Else
        Call Intacct.ClosePOTran
        Call Intacct.CloseMessage
        req = Intacct.xml()
        Call WriteLogFile("intacct.apinvoice.req.xml", req)

        s = Intacct.PostMessage(False)
        res = Intacct.lastResponse()
        Call WriteLogFile("intacct.apinvoice.res.xml", res)
    
        'remove failed records from batch
        s = "exec Intacct_ConfirmAPPosting " & DbQuote(num, HFApp.DivisionID) & "," & DbQuote(xml, req) & "," & DbQuote(xml, res)
        Call HFApp.SqlExec(s, dbHomeFront)
    End If





 Exit Sub
eh: Call ErrHandler(SRCFILE & "SendInvoicesToIntacct")
    Resume Next ' required. do not remove
End Sub

Private Function WriteLogFile(FileName As String, Text As String, Optional append As Boolean = False) As String
    Dim i As Integer
    Dim p As String
    
    If Not InIde() Then
        Text = Replace(Text, "<password>1Hyphensolutions!</password>", "<password>####</password>")
    End If
    
    i = FreeFile()
    p = PathAppend(GetFolderPath(CSIDL_LOCAL_APPDATA), "HomeFront")
    Call CreatePath("", p)
    If append Then
        Open PathAppend(p, FileName) For Append As #i
    Else
        Open PathAppend(p, FileName) For Output As #i
    End If
    Print #i, Text
    Close #i
    
    WriteLogFile = PathAppend(p, FileName)

End Function


Public Sub ReadPOItemIDs(WhereClause As String)
On Error GoTo eh:
    Dim s As String
    Dim poCSV As String
    Dim rs As Recordset
    Dim Intacct As New IntacctWrapper.IntacctWrapper
    
    
    s = ""
    s = s & "select distinct isnull(p.IntacctTransactionType,'')+'-'+invoiceitems.commitment DocID" & vbCrLf
    s = s & "  from invoices" & vbCrLf
    s = s & "  join invoiceitems on invoices.invoiceid=invoiceitems.invoiceid" & vbCrLf
    s = s & "  join pomaster p on invoiceitems.DivisionID = p.DivisionID and p.ponumber=invoiceitems.commitment" & vbCrLf
    s = s & " where invoices.DivisionID =" & HFApp.DivisionID & " and 1=1" & vbCrLf
    s = s & WhereClause & vbCrLf
    Set rs = HFApp.SqlExec(s)
    While Not rs.EOF
        poCSV = poCSV & "," & DbQuote(Str, "" & rs("DocID"))
        rs.MoveNext
    Wend
    poCSV = Mid(poCSV, 2)
    If poCSV = "" Then Exit Sub
    
    
    Call Intacct.OpenMessage(HFApp.Options.ValueByName("IntacctCompanyID"), HFApp.Options.ValueByName("IntacctUID"), HFApp.Options.ValueByName("IntacctPWD"), HFApp.Options.ValueByName("IntacctEntity"))
    Call Intacct.GetPOTranItems(poCSV)
    Call Intacct.CloseMessage
    Call WriteLogFile("intacct.getpoitems.req.xml", Intacct.xml())
    s = Intacct.PostMessage(False)
    Call WriteLogFile("intacct.getpoitems.res.xml", Intacct.lastResponse())
    If s <> "" Then
        s = "exec Intacct_SetPOItemIDs " & DbQuote(num, HFApp.DivisionID) & "," & DbQuote(Str, s)
        Call HFApp.SqlExec(s, dbHomeFront)
    End If
Exit Sub
eh: Call ErrHandler(SRCFILE & "ReadPOItemIDs")
End Sub

