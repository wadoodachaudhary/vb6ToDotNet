Attribute VB_Name = "MIntacct"
Option Explicit
Const SRCFILE = "MIntacct::"
Const IntacctDateFormat = "mm\/dd\/yyyy"

Public Sub SendEstimatesToIntacct(Batch As Long)
On Error GoTo eh
    Dim req As String
    Dim res As String
    Dim FileName As String
    Dim rs As Recordset
    Dim s As String
    Dim i As Integer
    Dim prevJob As String
    
    Dim OneTimeID As String
    OneTimeID = HFApp.Options.ValueByName("IntacctOneTimeItemID")
    
    Dim PostExtra As Boolean
    PostExtra = HFApp.Options.ValueByName("PostJCExtraAsSubJob") = "True"
    
    
    If HFApp.Options(PostSummarizedBudgets) Then
        s = ""
        s = s & "SELECT e.ChangeOrder,e.Job Job_No,'' JobDesc,e.JCExtra,null JCExtraDesc,e.JCCostCode,e.JCCategory,null Qty,'' UOM,round(SUM(e.Pretax+e.Tax),2) Amount" & vbCrLf
        s = s & ",isnull(isnull(nullif(g.intacctitemid,''),nullif(c.intacctitemid,''))," & DbQuote(Str, OneTimeID) & ") IntacctItemID,j.IntacctDepartment" & vbCrLf
        s = s & ",iif(upcc.job is null,0,1) IsNewCostCode" & vbCrLf
        s = s & ",iif(upc.job is null,0,1) IsNewCategory" & vbCrLf
        s = s & "FROM CancelledBudgets e" & vbCrLf
        s = s & "JOIN tbljobs j on e.divisionid=j.divisionid and e.job=j.job_no" & vbCrLf
        s = s & "JOIN StandardCostCodes c on e.divisionid=c.divisionid and e.jccostcode=c.costcode" & vbCrLf
        s = s & "JOIN StandardCategories g on e.divisionid=g.divisionid and e.jccategory=g.category" & vbCrLf
        s = s & "left join UnPostedJobCostCodes upcc on j.divisionid=upcc.divisionid and j.job_no=upcc.job and c.costcode=upcc.jccostcode" & vbCrLf
        s = s & "left join UnPostedJobCostCategories upc on j.divisionid=upc.divisionid and j.job_no=upc.job and c.costcode=upc.jccostcode and g.category=upc.jccategory" & vbCrLf
        s = s & "WHERE e.PostBatch=" & DbQuote(Num, Batch) & vbCrLf
        s = s & "GROUP BY e.ChangeOrder,e.Job,e.JCExtra,e.JCCostCode,e.JCCategory,c.intacctitemid,g.intacctitemid,j.IntacctDepartment,upcc.job,upc.job" & vbCrLf
        s = s & "UNION ALL" & vbCrLf
        s = s & "SELECT e.ChangeOrder,e.Job_No,e.JobDesc,e.JCExtra,left(e.AssemblyDescription,30) JCExtraDesc,e.JCCostCode,e.JCCategory,null Qty,'' UOM,round(SUM(e.BudgetPretax+e.BudgetJCTax),2) Amount" & vbCrLf
        s = s & ",isnull(isnull(nullif(g.intacctitemid,''),nullif(c.intacctitemid,''))," & DbQuote(Str, OneTimeID) & ") IntacctItemID,j.IntacctDepartment" & vbCrLf
        s = s & ",iif(upcc.job is null,0,1) IsNewCostCode" & vbCrLf
        s = s & ",iif(upc.job is null,0,1) IsNewCategory" & vbCrLf
        s = s & "FROM EstimatedItems e " & vbCrLf
        s = s & "JOIN tbljobs j on e.divisionid=j.divisionid and e.job_no=j.job_no" & vbCrLf
        s = s & "JOIN StandardCostCodes c on e.divisionid=c.divisionid and e.jccostcode=c.costcode" & vbCrLf
        s = s & "JOIN StandardCategories g on e.divisionid=g.divisionid and e.jccategory=g.category" & vbCrLf
        s = s & "left join UnPostedJobCostCodes upcc on j.divisionid=upcc.divisionid and j.job_no=upcc.job and c.costcode=upcc.jccostcode" & vbCrLf
        s = s & "left join UnPostedJobCostCategories upc on j.divisionid=upc.divisionid and j.job_no=upc.job and c.costcode=upc.jccostcode and g.category=upc.jccategory" & vbCrLf
        s = s & "WHERE e.BudgetPostingBatch=" & DbQuote(Num, Batch) & vbCrLf
        s = s & "GROUP BY e.ChangeOrder,e.Job_No,e.JobDesc,e.JCExtra,left(e.AssemblyDescription,30),e.JCCostCode,e.JCCategory,c.intacctitemid,g.intacctitemid,j.IntacctDepartment,upcc.job,upc.job" & vbCrLf
        s = s & "ORDER BY Job_No,e.JCExtra,e.JCCostCode,e.JCCategory" & vbCrLf
    Else
        s = ""
        s = s & "SELECT e.ChangeOrder,e.Job Job_No,'' JobDesc,e.JCExtra,null JCExtraDesc,e.JCCostCode,e.JCCategory,e.Qty,e.UOM,e.Pretax+e.Tax Amount" & vbCrLf
        s = s & ",isnull(isnull(nullif(g.intacctitemid,''),nullif(c.intacctitemid,''))," & DbQuote(Str, OneTimeID) & ") IntacctItemID,j.IntacctDepartment" & vbCrLf
        s = s & ",iif(upcc.job is null,0,1) IsNewCostCode" & vbCrLf
        s = s & ",iif(upc.job is null,0,1) IsNewCategory" & vbCrLf
        s = s & "FROM CancelledBudgets e" & vbCrLf
        s = s & "JOIN tbljobs j on e.divisionid=j.divisionid and e.job=j.job_no" & vbCrLf
        s = s & "JOIN StandardCostCodes c on e.divisionid=c.divisionid and e.jccostcode=c.costcode" & vbCrLf
        s = s & "JOIN StandardCategories g on e.divisionid=g.divisionid and e.jccategory=g.category" & vbCrLf
        s = s & "left join UnPostedJobCostCodes upcc on j.divisionid=upcc.divisionid and j.job_no=upcc.job and c.costcode=upcc.jccostcode" & vbCrLf
        s = s & "left join UnPostedJobCostCategories upc on j.divisionid=upc.divisionid and j.job_no=upc.job and c.costcode=upc.jccostcode and g.category=upc.jccategory" & vbCrLf
        s = s & "WHERE e.PostBatch=" & DbQuote(Num, Batch) & vbCrLf
        s = s & "UNION ALL" & vbCrLf
        s = s & "SELECT e.ChangeOrder,e.Job_No,e.JobDesc,e.JCExtra,left(e.AssemblyDescription,30) JCExtraDesc,e.JCCostCode,e.JCCategory,e.BudgetQty Qty,e.OrderUOM UOM,e.BudgetPretax+e.BudgetJCTax Amount" & vbCrLf
        s = s & ",isnull(isnull(nullif(g.intacctitemid,''),nullif(c.intacctitemid,''))," & DbQuote(Str, OneTimeID) & ") IntacctItemID,j.IntacctDepartment" & vbCrLf
        s = s & ",iif(upcc.job is null,0,1) IsNewCostCode" & vbCrLf
        s = s & ",iif(upc.job is null,0,1) IsNewCategory" & vbCrLf
        s = s & "FROM EstimatedItems e" & vbCrLf
        s = s & "JOIN tbljobs j on e.divisionid=j.divisionid and e.job_no=j.job_no" & vbCrLf
        s = s & "JOIN StandardCostCodes c on e.divisionid=c.divisionid and e.jccostcode=c.costcode" & vbCrLf
        s = s & "JOIN StandardCategories g on e.divisionid=g.divisionid and e.jccategory=g.category" & vbCrLf
        s = s & "left join UnPostedJobCostCodes upcc on j.divisionid=upcc.divisionid and j.job_no=upcc.job and c.costcode=upcc.jccostcode" & vbCrLf
        s = s & "left join UnPostedJobCostCategories upc on j.divisionid=upc.divisionid and j.job_no=upc.job and c.costcode=upc.jccostcode and g.category=upc.jccategory" & vbCrLf
        s = s & "WHERE e.BudgetPostingBatch=" & DbQuote(Num, Batch) & vbCrLf
        s = s & "ORDER BY Job_No,e.JCExtra,e.JCCostCode,e.JCCategory" & vbCrLf
    End If
    Set rs = HFApp.SqlExec(s)
    
    If rs.EOF Then Exit Sub
    
    Dim Intacct As New IntacctWrapper.IntacctWrapper
    Call Intacct.OpenMessage(HFApp.Options.ValueByName("IntacctCompanyID"), HFApp.Options.ValueByName("IntacctUID"), HFApp.Options.ValueByName("IntacctPWD"), HFApp.Options.ValueByName("IntacctEntity"))
    While Not rs.EOF
        If prevJob <> "" & rs("Job_No") & IIf(PostExtra, "-" & rs("jcextra"), "") Then
            If prevJob <> "" Then
                Call Intacct.CloseEstimate
            End If
            prevJob = "" & rs("Job_No") & IIf(PostExtra, "-" & rs("jcextra"), "")
            Call Intacct.OpenEstimate("" & rs("Job_No"), IIf(PostExtra, "" & rs("jcextra"), ""), HFApp.Options.ValueByName("IntacctEstimateType"), HFApp.Options.ValueByName("IntacctEstimateGLBudgetID"))
        End If
        Call Intacct.AddEstimateLine("" & rs("ChangeOrder") = "", "" & rs("IntacctItemID"), "" & rs("IsNewCostCode") = "1", "" & rs("JCCostCode"), "" & rs("IsNewCategory") = "1", "" & rs("JCCategory"), "" & rs("Qty"), "" & rs("UOM"), "" & rs("Amount"), "" & rs("ChangeOrder"), "" & rs("IntacctDepartment"))
        rs.MoveNext
    Wend
    If prevJob <> "" Then
        Call Intacct.CloseEstimate
        Call Intacct.CloseMessage
        req = Intacct.xml()
        FileName = WriteLogFile("intacct.estimates.req.xml", req)
        s = Intacct.PostMessage(False)
        res = Intacct.lastResponse()
        Call WriteLogFile("intacct.estimates.res.xml", res)
        
        s = "update estimateitems set BudgetPostingDate=getdate() where BudgetPostingBatch=" & DbQuote(Num, Batch)
        Call HFApp.SqlExec(s)

    End If

    
Exit Sub
eh: Call errHandler(SRCFILE & "SendEstimatesToIntacct", FileName)
    s = ""
    s = s & "update estimateitems set BudgetPostingBatch=0 where BudgetPostingBatch=" & DbQuote(Num, Batch) & vbCrLf
    s = s & "update CancelledBudgets set postbatch=0 where PostBatch=" & DbQuote(Num, Batch) & vbCrLf
    Call HFApp.SqlExec(s)
    Resume Next 'required do not remove
End Sub

Public Function CancelPOInIntacct(PONumber As String, reason As String) As Boolean
On Error GoTo eh
    
    Dim s As String
    Dim POTranType As String
    Dim Intacct As New IntacctWrapper.IntacctWrapper
    
    POTranType = HFApp.Options.ValueByName("IntacctPOType")
    
    Call Intacct.OpenMessage(HFApp.Options.ValueByName("IntacctCompanyID"), HFApp.Options.ValueByName("IntacctUID"), HFApp.Options.ValueByName("IntacctPWD"), HFApp.Options.ValueByName("IntacctEntity"))
    Call Intacct.DeletePOTran(POTranType, PONumber)
    Call Intacct.CloseMessage
    Call WriteLogFile("intacct.cancelpo.req.xml", Intacct.xml())
    s = Intacct.PostMessage(False)
    Call WriteLogFile("intacct.cancelpo.res.xml", Intacct.lastResponse())
    
    CancelPOInIntacct = True
    
Exit Function
eh:
    Select Case True
        Case Intacct.lastResponse() Like "*Document ID * does not exist*"
            CancelPOInIntacct = True
            Exit Function
        Case Else
            Call errHandler(SRCFILE & "CancelPOInIntacct")
            Call WriteLogFile("intacct.cancelpo.res.xml", Intacct.lastResponse())
    End Select
End Function


Public Sub SendPOsToIntacct(Batch As Long)
    Call WritePOs(Batch)
    Call WriteCOs(Batch)
End Sub


Private Sub WriteCOs(Batch As Long)
On Error GoTo eh
    Dim req As String
    Dim res As String
    
    Dim FileName As String
    Dim s As String
    Dim rs As Recordset
    Dim LastPO As String
    Dim ThisPO As String
    
    Dim TaxNJCAmount As Currency
    Dim TaxLocation As String
    Dim TaxLineID As String
    
    Dim Form1099 As String
    Dim Box1099 As String
    
    Dim POTranType As String
    Dim SCTranType As String
    Dim OneTimeID As String
    Dim OneTimeUOM As String
    Dim GstItemID As String
    Dim GstItemUOM As String
    POTranType = HFApp.Options.ValueByName("IntacctPOCOType")
    SCTranType = HFApp.Options.ValueByName("IntacctSubContractCOType")
    OneTimeID = HFApp.Options.ValueByName("IntacctOneTimeItemID")
    OneTimeUOM = HFApp.Options.ValueByName("IntacctOneTimeItemUOM")
    GstItemID = HFApp.Options.ValueByName("IntacctGstItemID")
    GstItemUOM = HFApp.Options.ValueByName("IntacctGstItemUOM")
   
    Dim PostExtra As Boolean
    PostExtra = HFApp.Options.ValueByName("PostJCExtraAsSubJob") = "True"

    Call ReadPOItemIDs(, Batch)

    Dim Intacct As New IntacctWrapper.IntacctWrapper
    Call Intacct.OpenMessage(HFApp.Options.ValueByName("IntacctCompanyID"), HFApp.Options.ValueByName("IntacctUID") _
                           , HFApp.Options.ValueByName("IntacctPWD"), HFApp.Options.ValueByName("IntacctEntity"))
    
    
    s = ""
    s = s & "select " & vbCrLf
    s = s & " po.ExternalID OrigDocID" & vbCrLf
    s = s & ",pi.ExternalID OrigLineID" & vbCrLf
    s = s & ",po.IntacctNJCTaxLineID OrigTaxLineID" & vbCrLf
    s = s & ",isnull(nullif(px.potype,''),'Purchase Order') potype,co.CODate" & vbCrLf
    s = s & ",co.seq CODesc   --dont change this. Intacct_ConfirmCOPosting requires it to join back to" & vbCrLf
    s = s & ",v.Vendor_id" & vbCrLf
    s = s & ",i.Description" & vbCrLf
    s = s & ",1 Qty" & vbCrLf
    s = s & ",isnull(isnull(nullif(cc.intacctitemid,''),nullif(c.intacctitemid,''))," & DbQuote(Str, OneTimeID) & ") IntacctItemID" & vbCrLf
    s = s & ",isnull(isnull(nullif(cc.intacctuom,''),nullif(c.intacctuom,''))," & DbQuote(Str, OneTimeUOM) & ") IntacctUOM" & vbCrLf
    s = s & ",i.popretax + i.pojctax Price" & vbCrLf
    s = s & ",i.ponjctax NJCTax " & vbCrLf
    s = s & ",j.gl_prefix Location" & vbCrLf
    s = s & ",i.Job" & vbCrLf
    s = s & ",i.JCExtra" & vbCrLf
    s = s & ",i.jcCostCode" & vbCrLf
    s = s & ",i.jcCategory" & vbCrLf
    s = s & ",j.IntacctDepartment Department" & vbCrLf
    s = s & ",v.Form1099" & vbCrLf
    s = s & ",v.Box1099" & vbCrLf
    s = s & ",j.arcustomer" & vbCrLf
    s = s & ",iif(upcc.job is null,0,1) IsNewCostCode" & vbCrLf
'    s = s & ",iif(upc.job is null,0,1) IsNewCategory" & vbCrLf
    s = s & ",1 IsNewCategory" & vbCrLf
    
    s = s & "from estimateitems i" & vbCrLf
    s = s & "join poitems pi on i.ChangeOrderOrigEstItemId=pi.estitemid and i.ponumber=pi.ponumber" & vbCrLf
    s = s & "join pomaster po on pi.divisionid=po.divisionid and pi.ponumber=po.ponumber" & vbCrLf
    s = s & "join tblpoindex px on po.divisionid=px.divisionid and po.poindex=px.poindex" & vbCrLf
    s = s & "join pochangeorders co on i.ponumber=co.ponumber and co.changeorder = i.pochangeordernumber" & vbCrLf
    s = s & "join StandardCostCodes cc on pi.divisionid=cc.divisionid and pi.jccostcode=cc.costcode" & vbCrLf
    s = s & "join Standardcategories c on pi.divisionid=c.divisionid and pi.jccategory=c.category" & vbCrLf
    s = s & "left join UnPostedJobCostCodes upcc on i.divisionid=upcc.divisionid and i.job=upcc.job and cc.costcode=upcc.jccostcode" & vbCrLf
' can't do this since it will find the original estimate posting which sent the category not the variance category
'    s = s & "left join UnPostedJobCostCategories upc on i.divisionid=upc.divisionid and i.job=upc.job and cc.costcode=upc.jccostcode and c.category=upc.jccategory" & vbCrLf
    
    s = s & "join tbljobs j on i.divisionid=j.divisionid and i.job=j.job_no" & vbCrLf
    s = s & "join tblvendors v on i.divisionid=v.divisionid and i.povendor=v.vendor_id" & vbCrLf
    s = s & "where isnull(po.ExternalID,'')<>'' and po.postingbatch<>0 and co.PostingBatch=" & DbQuote(Num, Batch) & vbCrLf
    s = s & "order by 1,2" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    While Not rs.EOF
        ThisPO = "" & rs("OrigDocID")
        If ThisPO <> LastPO Then
            
            If LastPO <> "" Then
                If TaxNJCAmount <> 0 Then Call Intacct.AddCOTranItem(ThisPO, "", GstItemID, "tax (non-costed)", 1, GstItemUOM, TaxNJCAmount, TaxLocation, "", "", False, "", False, "", "", Form1099, Box1099, "")
                Call Intacct.CloseCOTran
            End If
            
            If "" & rs("potype") = "Subcontract" Then
                Call Intacct.OpenCOTran(SCTranType, format("" & rs("codate"), IntacctDateFormat), format(Now(), IntacctDateFormat), "" & rs("vendor_id"), "" & rs("CODesc"))
            Else
                Call Intacct.OpenCOTran(POTranType, format("" & rs("codate"), IntacctDateFormat), format(Now(), IntacctDateFormat), "" & rs("vendor_id"), "" & rs("CODesc"))
            End If
            
            TaxNJCAmount = 0
            TaxLocation = ""
            TaxLineID = ""
            Form1099 = "" & rs("Form1099")
            Box1099 = "" & rs("Box1099")
            LastPO = ThisPO
                    
        End If
        
        Call Intacct.AddCOTranItem( _
            "" & rs("OrigDocID"), "" & rs("OrigLineID"), "" & rs("IntacctItemID"), "" & rs("Description"), Val("" & rs("Qty")), "" & rs("intacctuom"), Val("" & rs("price")), _
            "" & rs("Location"), "" & rs("job"), IIf(PostExtra, "" & rs("jcextra"), ""), "" & rs("IsNewCostCode") = "1", "" & rs("jccostcode"), "" & rs("IsNewCategory") = "1", "" & rs("jccategory"), "" & rs("Department"), _
            Form1099, Box1099, "" & rs("arcustomer"))
                                  
        TaxNJCAmount = TaxNJCAmount + Round(Val("" & rs("njctax")), 2)
        TaxLocation = "" & rs("Location")
        TaxLineID = "" & rs("OrigTaxLineID")
        
        rs.MoveNext
    Wend
    If ThisPO <> "" Then
        If TaxNJCAmount <> 0 Then Call Intacct.AddCOTranItem(ThisPO, TaxLineID, GstItemID, "tax (non-costed)", 1, GstItemUOM, TaxNJCAmount, TaxLocation, "", "", False, "", False, "", "", Form1099, Box1099, "")
        Call Intacct.CloseCOTran
        Call Intacct.CloseMessage
        req = Intacct.xml()
        FileName = WriteLogFile("intacct.pochangeorders.req.xml", req)
        
        s = Intacct.PostMessage(False)
        res = Intacct.lastResponse()
        Call WriteLogFile("intacct.pochangeorders.res.xml", res)
        
    End If
    
    'remove failed COs from batch
    If req = "" Then
        'nothing posted. remove all from batch
        s = "update pochangeorders set postingbatch=0,postingdate=null where postingbatch=" & DbQuote(Num, Batch)
        Call HFApp.SqlExec(s, dbHomefront)
    Else
        s = "exec Intacct_ConfirmCOPosting " & DbQuote(Num, HFApp.DivisionID) & "," & DbQuote(xml, req) & "," & DbQuote(xml, res)
        Call HFApp.SqlExec(s, dbHomefront)
    End If
    
    
Exit Sub
eh:
    Dim ejob As String
    Dim ecostcode As String
    Dim ecategory As String
    Select Case True
    Case Err.Description Like "*The task you selected is not associated with the project you selected. When you select a task dimension, that task must be associated with the project dimension that you selected at the line level. The task you selected is *. Go back and select another task.*"
        ecostcode = Parse(Parse(Err.Description, 3, "The task you selected is "), 1, ". Go back and select another task.")
        MsgBox "Intacct rejected this PO because cost code " & ecostcode & " has not been added to the job.", vbCritical, App.ProductName & " (version " & App.Major & "." & format(App.Minor, "00") & "." & format(App.Revision, "0000") & ")"

    Case Err.Description Like "*Cost type 'LAB--D60.10.30--ML390' specified is not valid.*"
        ejob = Parse(Parse(Err.Description, 3, "--"), 1, "'")
        ecostcode = Parse(Parse(Err.Description, 2, "--"), 1, "--")
        ecategory = Parse(Parse(Err.Description, 1, "--"), 2, "'")
        MsgBox "Intacct rejected this PO because cost type " & ecategory & " has not been added to cost code " & ecostcode & " for job " & ejob & ".", vbCritical, App.ProductName & " (version " & App.Major & "." & format(App.Minor, "00") & "." & format(App.Revision, "0000") & ")"
    
    Case Else
        Call errHandler(SRCFILE & "SendCOsToIntacct", FileName)
        
    End Select
    Resume Next 'this is required. do not remove
End Sub

Private Sub WritePOs(Batch As Long)
On Error GoTo eh
    Dim req As String
    Dim res As String
    Dim xmlCC As String
    
    Dim FileName As String
    Dim s As String
    Dim rs As Recordset
    Dim LastPO As String
    Dim ThisPO As String
    Dim poCSV As String
    
    Dim TaxNJCAmount As Currency
    Dim TaxLocation As String
    
    Dim Form1099 As String
    Dim Box1099 As String
    
    Dim POTranType As String
    Dim SCTranType As String
    Dim OneTimeID As String
    Dim OneTimeUOM As String
    Dim GstItemID As String
    Dim GstItemUOM As String
    POTranType = HFApp.Options.ValueByName("IntacctPOType")
    SCTranType = HFApp.Options.ValueByName("IntacctSubContractType")
    OneTimeID = HFApp.Options.ValueByName("IntacctOneTimeItemID")
    OneTimeUOM = HFApp.Options.ValueByName("IntacctOneTimeItemUOM")
    GstItemID = HFApp.Options.ValueByName("IntacctGstItemID")
    GstItemUOM = HFApp.Options.ValueByName("IntacctGstItemUOM")
   
    Dim Reference As String
    Dim ReferenceFld As String
    ReferenceFld = HFApp.Options.ValueByName("IntacctPOReferenceFld")

    Dim PostExtra As Boolean
    PostExtra = HFApp.Options.ValueByName("PostJCExtraAsSubJob") = "True"

    Dim Intacct As New IntacctWrapper.IntacctWrapper
    Call Intacct.OpenMessage(HFApp.Options.ValueByName("IntacctCompanyID"), HFApp.Options.ValueByName("IntacctUID") _
                           , HFApp.Options.ValueByName("IntacctPWD"), HFApp.Options.ValueByName("IntacctEntity"))
    
    
    s = ""
    s = s & "select d.po,d.pojob,d.line,d.podate,isnull(nullif(podesc,''),d.poindex) podesc,d.vendor,d.orderedby,d.terms,d.shipvia,d.fob,d.retainagerate" & vbCrLf
    s = s & ",d.linedesc,d.linejob,d.lineextra,d.linecostcode,d.linecategory,d.linetaxgroup,d.linejctax,d.linenjctax,d.linecommittedquantity,d.linecommittedunitprice" & vbCrLf
    s = s & ",d.linecommitteduom,isnull(nullif(d.potype,''),'Purchase Order') potype,d.Linepretax,d.linetax,j.gl_prefix" & vbCrLf
    s = s & ",isnull(isnull(nullif(cc.intacctitemid,''),nullif(c.intacctitemid,''))," & DbQuote(Str, OneTimeID) & ") IntacctItemID" & vbCrLf
    s = s & ",isnull(isnull(nullif(cc.intacctuom,''),nullif(c.intacctuom,''))," & DbQuote(Str, OneTimeUOM) & ") IntacctItemUOM" & vbCrLf
    s = s & ",j.IntacctDepartment" & vbCrLf
    s = s & ",v.Form1099,v.Box1099,j.arcustomer" & vbCrLf
    s = s & ",iif(upcc.job is null,0,1) IsNewCostCode" & vbCrLf
    s = s & ",iif(upc.job is null,0,1) IsNewCategory" & vbCrLf
    s = s & "from dbo.JCPODetails d" & vbCrLf
    s = s & "join tbljobs j on d.divisionid=j.divisionid and d.linejob=j.job_no" & vbCrLf
    s = s & "join tblvendors v on d.divisionid=v.divisionid and d.vendor=v.vendor_id" & vbCrLf
    s = s & "join StandardCostCodes cc on d.divisionid=cc.divisionid and d.linecostcode=cc.costcode" & vbCrLf
    s = s & "join Standardcategories c on d.divisionid=c.divisionid and d.linecategory=c.category" & vbCrLf
    s = s & "left join UnPostedJobCostCodes upcc on j.divisionid=upcc.divisionid and j.job_no=upcc.job and cc.costcode=upcc.jccostcode" & vbCrLf
    s = s & "left join UnPostedJobCostCategories upc on j.divisionid=upc.divisionid and j.job_no=upc.job and cc.costcode=upc.jccostcode and c.category=upc.jccategory" & vbCrLf
    s = s & "where d.PostingBatch=" & DbQuote(Num, Batch) & vbCrLf
    s = s & "order by 1,3" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    While Not rs.EOF
        ThisPO = "" & rs("PO")
        If ThisPO <> LastPO Then
            
            If LastPO <> "" Then
                If TaxNJCAmount <> 0 Then Call Intacct.AddPOTranItem(GstItemID, "tax (non-costed)", 1, GstItemUOM, TaxNJCAmount, TaxLocation, "", "", False, "", False, "", "", Form1099, Box1099, "")
                Call Intacct.ClosePOTran
            End If
            
            
            
            If ReferenceFld = "PO Number" Then
                Reference = "" & rs("PO")
            Else
                Reference = "" & rs("POJob") & IIf(PostExtra And "" & rs("LineExtra") <> "", "-" & rs("LineExtra"), "")
            End If
            If "" & rs("potype") = "Subcontract" Then
                Call Intacct.OpenPOTran(SCTranType, format("" & rs("podate"), IntacctDateFormat), format(Now(), IntacctDateFormat), "", "" & rs("vendor"), "" & rs("po"), "", Reference)
                poCSV = poCSV & "," & DbQuote(Str, SCTranType & "-" & rs("po"))
            Else
                Call Intacct.OpenPOTran(POTranType, format("" & rs("podate"), IntacctDateFormat), format(Now(), IntacctDateFormat), "", "" & rs("vendor"), "" & rs("po"), "", Reference)
                poCSV = poCSV & "," & DbQuote(Str, POTranType & "-" & rs("po"))
            End If
            
            TaxNJCAmount = 0
            TaxLocation = ""
            Form1099 = "" & rs("Form1099")
            Box1099 = "" & rs("Box1099")
            LastPO = ThisPO
                    
        End If
        
        Call Intacct.AddPOTranItem("" & rs("IntacctItemID"), "" & rs("linedesc"), Val("" & rs("linecommittedquantity")), "" & rs("intacctitemuom"), _
                                  Round((Val("" & rs("linepretax")) + Val("" & rs("linejctax"))) / Val("" & rs("linecommittedquantity")), 10), _
                                  "" & rs("GL_Prefix"), "" & rs("linejob"), IIf(PostExtra, "" & rs("LineExtra"), ""), "" & rs("IsNewCostCode") = "1", "" & rs("linecostcode"), "" & rs("IsNewCategory") = "1", "" & rs("linecategory"), "" & rs("IntacctDepartment"), _
                                  Form1099, Box1099, "" & rs("arcustomer"))
                                  
        TaxNJCAmount = TaxNJCAmount + Round(Val("" & rs("linenjctax")), 2)
        TaxLocation = "" & rs("GL_Prefix")
        
        rs.MoveNext
    Wend
    If ThisPO <> "" Then
        If TaxNJCAmount <> 0 Then Call Intacct.AddPOTranItem(GstItemID, "tax", 1, GstItemUOM, TaxNJCAmount, TaxLocation, "", "", True, "", True, "", "", Form1099, Box1099, "")
        Call Intacct.ClosePOTran
        Call Intacct.CloseMessage
        req = Intacct.xml()
        FileName = WriteLogFile("intacct.purchaseorders.req.xml", req)
        xmlCC = Intacct.xmlCC()
        Call WriteLogFile("intacct.purchaseorders.CC.req.xml", xmlCC)
        
        
        s = ""
        s = s & "update batches set" & vbCrLf
        s = s & " datafile=cast(" & DbQuote(Str, req) & " as varbinary)" & vbCrLf
        s = s & ",rejectfile=null" & vbCrLf
        s = s & "where batch=" & DbQuote(Num, Batch) & vbCrLf
        Call HFApp.SqlExec(s)
        
        s = Intacct.PostMessage(False)
        res = Intacct.lastResponse()
        Call WriteLogFile("intacct.purchaseorders.res.xml", res)
        
        s = ""
        s = s & "update batches set" & vbCrLf
        s = s & " rejectfile=cast(" & DbQuote(Str, res) & " as varbinary)" & vbCrLf
        s = s & "where batch=" & DbQuote(Num, Batch) & vbCrLf
        Call HFApp.SqlExec(s)
        
        poCSV = Mid(poCSV, 2)
                
        'put tax linenumbers on pomaster
        s = ""
        s = s & "update p set" & vbCrLf
        s = s & " IntacctTransactionType  = t.optionvalue" & vbCrLf
        s = s & ",IntacctNJCTaxLineNumber = (select max(line)+1 from jcpodetails i where p.divisionid=i.divisionid and p.ponumber=i.po)" & vbCrLf
        s = s & "from pomaster p" & vbCrLf
        s = s & "join tblpoindex x on p.divisionid=x.divisionid and p.poindex=x.poindex" & vbCrLf
        s = s & "join appoptions t on p.divisionid=t.divisionid and t.optionname=case when isnull(nullif(x.potype,'') ,'Purchase Order') ='Purchase Order' then 'IntacctPOType' else 'IntacctSubcontractType' end" & vbCrLf
        s = s & "where PostingBatch=" & DbQuote(Num, Batch) & vbCrLf
        HFApp.SqlExec s


    End If
    
    'remove failed pos from batch
    If req <> "" Then
        s = "exec Intacct_ConfirmPOPosting " & DbQuote(Num, HFApp.DivisionID) & "," & DbQuote(xml, req) & "," & DbQuote(xml, res)
        Call HFApp.SqlExec(s, dbHomefront)
    End If
    
    
Exit Sub
eh:
    Dim ejob As String
    Dim ecostcode As String
    Dim ecategory As String
    Select Case True
    Case Err.Description Like "*The task you selected is not associated with the project you selected. When you select a task dimension, that task must be associated with the project dimension that you selected at the line level. The task you selected is *. Go back and select another task.*"
        ecostcode = Parse(Parse(Err.Description, 3, "The task you selected is "), 1, ". Go back and select another task.")
        MsgBox "Intacct rejected this PO because cost code " & ecostcode & " has not been added to the job.", vbCritical, App.ProductName & " (version " & App.Major & "." & format(App.Minor, "00") & "." & format(App.Revision, "0000") & ")"

    Case Err.Description Like "*Cost type 'LAB--D60.10.30--ML390' specified is not valid.*"
        ejob = Parse(Parse(Err.Description, 3, "--"), 1, "'")
        ecostcode = Parse(Parse(Err.Description, 2, "--"), 1, "--")
        ecategory = Parse(Parse(Err.Description, 1, "--"), 2, "'")
        MsgBox "Intacct rejected this PO because cost type " & ecategory & " has not been added to cost code " & ecostcode & " for job " & ejob & ".", vbCritical, App.ProductName & " (version " & App.Major & "." & format(App.Minor, "00") & "." & format(App.Revision, "0000") & ")"
    
    Case Else
        Call errHandler(SRCFILE & "SendPOsToIntacct", FileName)
        
    End Select
    Resume Next 'this is required. do not remove
End Sub


Public Sub ReadPOItemIDs(Optional POWhereClause As String, Optional COBatch As Long)
On Error GoTo eh:
    Dim s As String
    Dim poCSV As String
    Dim rs As Recordset
    Dim Intacct As New IntacctWrapper.IntacctWrapper
    
    'when looking up POs -- i dont think this gets used...
    If POWhereClause <> "" Then
        s = ""
        s = s & "select distinct isnull(p.IntacctTransactionType,'')+'-'+p.ponumber DocID" & vbCrLf
        s = s & "  from pomaster p " & vbCrLf
        s = s & " where p.DivisionID =" & HFApp.DivisionID & " and 1=1" & vbCrLf
        s = s & POWhereClause & vbCrLf
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
            poCSV = poCSV & "," & DbQuote(Str, "" & rs("DocID"))
            rs.MoveNext
        Wend
    End If
    
    'when looking up POs in a CO batch
    If COBatch <> 0 Then
        s = ""
        s = s & "select distinct isnull(po.IntacctTransactionType,'')+'-'+po.ponumber DocID" & vbCrLf
        s = s & "from pochangeorders co " & vbCrLf
        s = s & "join pomaster po on co.divisionid=po.divisionid and co.ponumber=po.ponumber" & vbCrLf
        s = s & "where po.postingbatch<>0 and co.PostingBatch=" & DbQuote(Num, COBatch) & vbCrLf
        s = s & "and isnull(po.ExternalID,'')=''" & vbCrLf
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
            poCSV = poCSV & "," & DbQuote(Str, "" & rs("DocID"))
            rs.MoveNext
        Wend
    End If
    
    poCSV = Mid(poCSV, 2)
    If poCSV = "" Then Exit Sub
        
    Call Intacct.OpenMessage(HFApp.Options.ValueByName("IntacctCompanyID"), HFApp.Options.ValueByName("IntacctUID"), HFApp.Options.ValueByName("IntacctPWD"), HFApp.Options.ValueByName("IntacctEntity"))
    Call Intacct.GetPOTranItems(poCSV)
    Call Intacct.CloseMessage
    Call WriteLogFile("intacct.getpoitems.req.xml", Intacct.xml())
    s = Intacct.PostMessage(False)
    Call WriteLogFile("intacct.getpoitems.res.xml", Intacct.lastResponse())
    If s <> "" Then s = "exec Intacct_SetPOItemIDs " & DbQuote(Num, HFApp.DivisionID) & "," & DbQuote(Str, s)
    Call HFApp.SqlExec(s, dbHomefront)

Exit Sub
eh: Call errHandler(SRCFILE & "ReadPOItemIDs")
End Sub




Private Function WriteLogFile(FileName As String, Text As String, Optional Append As Boolean = False) As String
    Dim i As Integer
    Dim p As String
    
    If Not InIde() Then
        Text = Replace(Text, "<password>1Hyphensolutions!</password>", "<password>####</password>")
    End If
    
    i = FreeFile()
    p = PathAppend(GetFolderPath(CSIDL_LOCAL_APPDATA), "HomeFront")
    Call CreatePath("", p)
    If Append Then
        Open PathAppend(p, FileName) For Append As #i
    Else
        Open PathAppend(p, FileName) For Output As #i
    End If
    Print #i, Text
    Close #i
    
    WriteLogFile = PathAppend(p, FileName)

End Function

