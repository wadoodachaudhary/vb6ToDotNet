Attribute VB_Name = "mAccounting"
Option Explicit
Public AccountMask As String


Public Function TimberlineAccounting(Optional InvValidation As Boolean = False) As Boolean

If InvValidation Then
    TimberlineAccounting = HFApp.Options(AccountingSystem) = asTimberline
Else
    TimberlineAccounting = HFApp.Options(AccountingSystem) = asTimberline And App.Options.Value(UseTimberlineAsPOSource)
End If


End Function

Public Function AccountingDB(Optional InvValidation As Boolean = False) As Connections

    If TimberlineAccounting(InvValidation) Then
        If HFApp.ConnectionString(dbTimberlinePVdata) <> "" Then
            AccountingDB = dbTimberlinePVdata
        Else
            AccountingDB = dbAccountingDictionary
        End If
    Else
        AccountingDB = dbHomeFront
    End If
End Function


Public Function ValidateJob(ByRef Job As String) As Recordset
On Error GoTo err
    Dim s As String
    Dim rs As Recordset
    
    If TimberlineAccounting Then
        s = ""
        s = s & "SELECT job,jdesc description,jaddr1 Municipal_address,jstatus Status,0 WrapInsuranceExempt" & vbCrLf
        s = s & "  FROM master_jcm_record_1_1" & vbCrLf
        If AccountingDB = dbTimberlinePVdata Then
            s = s & " WHERE job=" & DbQuote(Str, FormatTSField(10, Job, True))
        Else
            s = s & " WHERE job=" & DbQuote(Str, HFApp.FormatJob(Job)) & vbCrLf
        End If
    Else
        
        
        s = ""
        s = s & "SELECT j.job_no job,j.description,j.Municipal_address,j.Status,c.WrapInsuranceExempt" & vbCrLf
        s = s & "FROM tbljobs j" & vbCrLf
        s = s & "LEFT JOIN tblLocality c on j.community=c.area" & vbCrLf
        s = s & "WHERE j.DivisionID = " & HFApp.DivisionID & vbCrLf
        If HFApp.Options(AccountingSystem) = asTimberline Then
            s = s & "and j.job_no=" & DbQuote(Str, StripFormating(Job)) & vbCrLf
        Else
            s = s & "and j.job_no=" & DbQuote(Str, Job) & vbCrLf
        End If
    End If
    
    Set ValidateJob = HFApp.SqlExec(s, AccountingDB)
    Exit Function
err:
    If err.Number <> 0 Then
        MsgBox err.Description & " Validate Job"
    End If
    
End Function

Public Function ValidateJobExtra(Job As String, Extra As String) As String
    Dim s As String
    If TimberlineAccounting Then
        If AccountingDB = dbTimberlinePVdata Then
           s = "select extra,xdesc description FROM master_jcm_record_2 where xjob=" & DbQuote(Str, FormatTSField(10, Job, True)) & " and extra=" & DbQuote(Str, FormatTSField(10, Extra))
        Else
           s = "select extra,xdesc description FROM master_jcm_record_2 where xjob=" & DbQuote(Str, Job) & " and extra=" & DbQuote(Str, Extra)
        End If
    Else
        s = "select extra,description FROM jobextras where DivisionID = " & HFApp.DivisionID & " and job=" & DbQuote(Str, Job) & " and extra=" & DbQuote(Str, Extra)
    End If
    ValidateJobExtra = s
End Function

Public Function ValidateJobExtraPhase(Job As String, Extra As String, Phase As String) As String
    Dim s As String
    
    
    If TimberlineAccounting Then
        If AccountingDB = dbTimberlinePVdata Then
            s = ""
            s = s & "select 1,phase costcode,pdesc description,'' DebitAccount,'' DebitAccountDesc" & vbCrLf
            s = s & "  from master_jcm_record_3 " & vbCrLf
            s = s & " where pjob=" & DbQuote(Str, FormatTSField(10, Job, True)) & vbCrLf
            s = s & "   and pextra=" & DbQuote(Str, FormatTSField(10, Extra)) & vbCrLf
            s = s & "   and phase=" & DbQuote(Str, FormatTSField(12, Phase)) & vbCrLf
            s = s & "union all" & vbCrLf
            s = s & "select 2,sphase costcode,spdesc description,'' DebitAccount,'' DebitAccountDesc" & vbCrLf
            s = s & "  from master_jcm_record_16" & vbCrLf
            s = s & " where sphase=" & DbQuote(Str, FormatTSField(12, Phase)) & vbCrLf
            s = s & "order by 1" & vbCrLf
        
        Else
            s = ""
            s = s & "select 1,phase costcode,pdesc description,'' DebitAccount,'' DebitAccountDesc" & vbCrLf
            s = s & "  from master_jcm_record_3 " & vbCrLf
            s = s & " where pjob=" & DbQuote(Str, HFApp.FormatJob(Job)) & vbCrLf
            s = s & "   and pextra=" & DbQuote(Str, Extra) & vbCrLf
            s = s & "   and phase=" & DbQuote(Str, HFApp.FormatCostCode(Phase)) & vbCrLf
            s = s & "union all" & vbCrLf
            s = s & "select 2,sphase costcode,spdesc description,'' DebitAccount,'' DebitAccountDesc" & vbCrLf
            s = s & "  from master_jcm_record_16" & vbCrLf
            s = s & " where sphase=" & DbQuote(Str, HFApp.FormatCostCode(Phase)) & vbCrLf
            s = s & "order by 1" & vbCrLf
        End If
    Else
        s = ""
        s = s & "select 1,c.costcode,c.description,a.account debitaccount,a.description debitaccountdesc" & vbCrLf
        s = s & "  from JobExtraCostCodes c" & vbCrLf
        s = s & "  left outer join standardcostcodes s on(s.DivisionID = " & HFApp.DivisionID & " and c.costcode=s.costcode)"
        s = s & "  left outer join glaccounts a on(s.DivisionID = a.DivisionID and s.debitaccount=a.account)"
        s = s & " where c.DivisionID = " & HFApp.DivisionID & " and c.job=" & DbQuote(Str, Job) & vbCrLf
        s = s & "   and c.extra=" & DbQuote(Str, Extra) & vbCrLf
        s = s & "   and c.costcode=" & DbQuote(Str, HFApp.FormatCostCode(Phase)) & vbCrLf
        s = s & "union all" & vbCrLf
        s = s & "select 2,s.costcode,s.description,a.account debitaccount,a.description debitaccountdesc" & vbCrLf
        s = s & "  from standardcostcodes s" & vbCrLf
        s = s & "  left outer join glaccounts a on(s.DivisionID = a.DivisionID and s.debitaccount=a.account)"
        s = s & " where s.DivisionID = " & HFApp.DivisionID & " and costcode=" & DbQuote(Str, HFApp.FormatCostCode(Phase)) & vbCrLf
    End If
    ValidateJobExtraPhase = s
End Function
Public Function ValidateJobExtraPhaseCategory(Job As String, Extra As String, Phase As String, Category As String) As String
    Dim s As String
    If TimberlineAccounting Then
        If AccountingDB = dbTimberlinePVdata Then
            s = ""
            s = s & "select 1,cat Category ,cdesc Description,'' DebitAccount,'' DebitAccountDesc" & vbCrLf
            s = s & "  from master_jcm_record_4 " & vbCrLf
            s = s & " where cjob=" & DbQuote(Str, FormatTSField(10, Job, True)) & vbCrLf
            s = s & "   and cextra=" & DbQuote(Str, Extra) & vbCrLf
            s = s & "   and cphase=" & DbQuote(Str, FormatTSField(12, Phase)) & vbCrLf
            s = s & "   And cat=" & DbQuote(Str, FormatTSField(3, Category)) & vbCrLf
            s = s & "union all" & vbCrLf
            s = s & "select 2,scat Category ,scdesc Description,'' DebitAccount,'' DebitAccountDesc" & vbCrLf
            s = s & "  from master_jcm_record_17" & vbCrLf
            s = s & " where scat=" & DbQuote(Str, FormatTSField(3, Category)) & vbCrLf
            s = s & "order by 1" & vbCrLf
        
        Else
            s = ""
            s = s & "select 1,cat Category ,cdesc Description,'' DebitAccount,'' DebitAccountDesc" & vbCrLf
            s = s & "  from master_jcm_record_4 " & vbCrLf
            s = s & " where cjob=" & DbQuote(Str, HFApp.FormatJob(Job)) & vbCrLf
            s = s & "   and cextra=" & DbQuote(Str, Extra) & vbCrLf
            s = s & "   and cphase=" & DbQuote(Str, HFApp.FormatCostCode(Phase)) & vbCrLf
            s = s & "   And cat=" & DbQuote(Str, Category) & vbCrLf
            s = s & "union all" & vbCrLf
            s = s & "select 2,scat Category ,scdesc Description,'' DebitAccount,'' DebitAccountDesc" & vbCrLf
            s = s & "  from master_jcm_record_17" & vbCrLf
            s = s & " where scat=" & DbQuote(Str, Category) & vbCrLf
            s = s & "order by 1" & vbCrLf
        End If
    Else
        s = ""
        s = s & "select 1,j.category,j.description,gl.account DebitAccount,gl.description DebitAccountDesc" & vbCrLf
        s = s & "from JobExtraCostCodeCategories j" & vbCrLf
        s = s & "left outer join standardcostcodes cc on(j.divisionid=cc.divisionid and j.costcode=cc.costcode)" & vbCrLf
        s = s & "left outer join standardcategories ca on(j.divisionid=ca.divisionid and j.category=ca.category)" & vbCrLf
        s = s & "left outer join glaccounts gl on(j.DivisionID=gl.DivisionID and gl.account=isnull(nullif(cc.debitaccount,''),ca.debitaccount))" & vbCrLf
        s = s & "where j.DivisionID=" & DbQuote(num, HFApp.DivisionID) & " and j.job=" & DbQuote(Str, Job) & " and j.extra=" & DbQuote(Str, Extra) & " and j.costcode=" & DbQuote(Str, Phase) & " and j.category=" & DbQuote(Str, Category) & vbCrLf
        s = s & "union all" & vbCrLf
        s = s & "select 2,ca.category,ca.description,gl.account DebitAccount,gl.description DebitAccountDesc" & vbCrLf
        s = s & "from standardcategories ca" & vbCrLf
        s = s & "left outer join standardcostcodes cc on(ca.divisionid=cc.divisionid and cc.costcode=" & DbQuote(Str, Phase) & ")" & vbCrLf
        s = s & "left outer join glaccounts gl on(ca.DivisionID=gl.DivisionID and gl.account=isnull(nullif(cc.debitaccount,''),ca.debitaccount))" & vbCrLf
        s = s & "where ca.DivisionID=" & DbQuote(num, HFApp.DivisionID) & " and ca.category=" & DbQuote(Str, Category) & vbCrLf
    End If
    ValidateJobExtraPhaseCategory = s
End Function

Public Function ValidateEquipment(Equipment As String) As String
    Dim s As String
    s = "select Equipment,Description from equipment where sold<>1 and DivisionID = " & HFApp.DivisionID & " and Equipment=" & DbQuote(Str, Equipment)
    ValidateEquipment = s
End Function
Public Function ValidateEquipmentCostCode(CostCode As String) As String
    Dim s As String
    s = "select CostCode,Description from standardeqcostcodes where DivisionID = " & HFApp.DivisionID & " and CostCode=" & DbQuote(Str, CostCode)
    ValidateEquipmentCostCode = s
End Function


Public Function ValidateGlAccount(Account As String) As String
    Dim s As String
    If TimberlineAccounting Then
        If AccountingDB = dbTimberlinePVdata Then
            s = "select AAcct account,Atitle description FROM master_GLm_record_1 where aacct=" & DbQuote(Str, FormatTSField(25, Account))
        Else
            s = "select AAcct account,Atitle description FROM master_GLm_record_1 where aacct=" & DbQuote(Str, FormatAccount(Account))
        End If
    Else
        s = "select account,description FROM glaccounts where DivisionID = " & HFApp.DivisionID & " and account=" & DbQuote(Str, FormatAccount(Account))
    End If
    ValidateGlAccount = s
End Function

Public Function SelectJob() As String
    Dim s As String
    
    If TimberlineAccounting Then
        
        s = ""
        s = s & "SELECT job " & Quote(App.Options(Caption_Job))
        s = s & "      ,Jdesc Description"
        s = s & "      ,jstatus Status"
        If App.Options(ShowJobPM) Then s = s & "      ,jprjmgr " & Quote(App.Options(Caption_JobPM))
        If App.Options(ShowJobTitle1) Then s = s & "      ,jtitl1 " & Quote(App.Options(Caption_JobTitle1))
        If App.Options(ShowJobTitle2) Then s = s & "      ,jtitl2 " & Quote(App.Options(Caption_JobTitle2))
        If App.Options(ShowJobTitle3) Then s = s & "      ,jtitl3 " & Quote(App.Options(Caption_JobTitle3))
        If App.Options(ShowJobTitle4) Then s = s & "      ,jtitl4 " & Quote(App.Options(Caption_JobTitle4))
        s = s & "      ,jaddr1 Address"
        If AccountingDB = dbTimberlinePVdata Then
            s = s & " FROM master_jcm_r1"
            s = s & " WHERE jstatus<>3"
        Else
            s = s & " FROM master_jcm_record_1_1"
            s = s & " WHERE jstatus<>'Closed'"
        End If
    Else
        s = ""
        s = s & "SELECT job_no " & Quote(App.Options(Caption_Job)) & vbCrLf
        s = s & "      ,Description" & vbCrLf
        s = s & "      ,case isnull(inactive,0) when 1 then 'Closed' else 'In progress' end Status" & vbCrLf
        If App.Options(ShowJobPM) Then s = s & "          ,pm " & Quote(App.Options(Caption_JobPM)) & vbCrLf
        If App.Options(ShowJobTitle1) Then s = s & "      ,Purchaser " & Quote(App.Options(Caption_JobTitle1)) & vbCrLf
        If App.Options(ShowJobTitle2) Then s = s & "      ,Estimator " & Quote(App.Options(Caption_JobTitle2)) & vbCrLf
        s = s & "      ,Municipal_Address Address" & vbCrLf
        s = s & "  FROM tbljobs" & vbCrLf
        s = s & " WHERE DivisionID = " & HFApp.DivisionID & " and isnull(inactive,0)=0 and isnull(isquote,0)=0 " & vbCrLf
        
        If IsIn("" & HFApp.Options(AccountingSystem), "" & asSimply, "" & asQuickBooks, "" & asQuickBooksOnline) Then
            s = s & " and isnull(ExternalJobID,'')<>''" & vbCrLf
        End If
        
    End If
    
    SelectJob = s
    
End Function

Public Function SelectPO(Optional Vendor As String = "", Optional Job As String = "") As String
    Dim s As String
    
    If TimberlineAccounting Then
        s = ""
        s = s & "select sub " & App.Options(Caption_Commitment) & vbCrLf
        s = s & "      ,sdesc Description" & vbCrLf
        s = s & "      ,vendor " & App.Options(Caption_Vendor) & vbCrLf
        s = s & "      ,vname " & App.Options(Caption_Vendor) & "Name" & vbCrLf
        s = s & "      ,sactcd Completed" & vbCrLf
        s = s & "      ,samt OriginalAmt" & vbCrLf
        s = s & "      ,sapprco Changes" & vbCrLf
        s = s & "      ,samtinv InvoicedAmt" & vbCrLf
        s = s & "      ,samt+sapprco-samtinv RemainingAmt" & vbCrLf
        s = s & "      ,sjob " & App.Options(Caption_Job) & vbCrLf
        If AccountingDB = dbTimberlinePVdata Then
            s = s & "  from master_jcm_r12" & vbCrLf
            s = s & "  inner join master_apm_r9 on master_jcm_r12.svendor=master_apm_r9.vendor" & vbCrLf
            s = s & " where sclosed=0" & vbCrLf
            If (Not App.Options(AllowCrossPayingPOs)) And Vendor <> "" Then
                If Val(Vendor) = 0 Then
                    s = s & "   and vendor=" & DbQuote(Str, Vendor) & vbCrLf
                Else
                    s = s & "   and vendor=" & DbQuote(Str, FormatTSField(10, Vendor)) & vbCrLf
                End If
            End If
            If Job <> "" Then
                s = s & "   and sjob=" & DbQuote(Str, FormatTSField(10, Job, True)) & vbCrLf
            End If
            s = s & "   and samt+sapprco-samtinv<>0" & vbCrLf

        Else
            s = s & "  from master_jcm_record_12" & vbCrLf
            s = s & "  inner join master_apm_record_9 on master_jcm_record_12.svendor=master_apm_record_9.vendor" & vbCrLf
            s = s & " where sclosed=0" & vbCrLf
            If (Not App.Options(AllowCrossPayingPOs)) And Vendor <> "" Then
                s = s & "   and vendor=" & DbQuote(Str, Vendor) & vbCrLf
            End If
            If Job <> "" Then
                s = s & "   and sjob=" & DbQuote(Str, Job) & vbCrLf
            End If
            s = s & "   and samt+sapprco-samtinv<>0" & vbCrLf
        End If

    Else
        s = ""
        s = s & "select po Commitment,podesc Description,Vendor,vendordesc VendorName,pocompleted CompletedDate,pojob Job" & vbCrLf
        s = s & "      ,sum(linepretax+linetax) OriginalAmt" & vbCrLf
        s = s & "      ,sum(lineinvoicedpretax+lineinvoicedtax) InvoicedAmt" & vbCrLf
        s = s & "      ,sum(linepretax+linetax)-sum(lineinvoicedpretax+lineinvoicedtax) RemainingAmt" & vbCrLf
        s = s & "from jcpodetails" & vbCrLf
        s = s & "where isnull(TBDVendor,0)=0 and pocancelled<>1 and pojobstatus<>'closed' and DivisionID=" & DbQuote(num, HFApp.DivisionID) & vbCrLf
        If Not App.Options(AllowCrossPayingPOs) And Vendor <> "" Then
            s = s & " and vendor=" & DbQuote(Str, Vendor) & vbCrLf
        End If
        If Job <> "" Then
            s = s & " and pojob=" & DbQuote(Str, Job) & vbCrLf
        End If
        If HFApp.Options(AccountingSystem) = asTimberline Then
            s = s & " AND postingbatch<>0" & vbCrLf
        End If
        s = s & "group by po,podesc,Vendor,vendordesc,pocompleted,pojob" & vbCrLf
        s = s & "having sum(linepretax+linetax)-sum(lineinvoicedpretax+lineinvoicedtax)<>0" & vbCrLf
    End If
    
    SelectPO = s
    
End Function

Public Function InvoiceInHomefront(Vendor As String, Invoice As String, ByRef Status As String) As Boolean
    Dim s As String
    Dim rs As Recordset
    
    
    s = ""
    s = s & "delete invoices" & vbCrLf
    s = s & " where DivisionID=" & DbQuote(num, HFApp.DivisionID) & vbCrLf
    s = s & "   and isnull(vendor,'')=''" & vbCrLf
    s = s & "   and isnull(invoice,'')=''" & vbCrLf
    Call HFApp.SqlExec(s, dbHomeFront)
    
    
    s = ""
    s = s & "select v.vendor_name,i.status " & vbCrLf
    s = s & "  from invoices i left outer join tblvendors v on i.DivisionID = v.DivisionID and i.vendor=v.vendor_id" & vbCrLf
    s = s & " where i.DivisionID=" & DbQuote(num, HFApp.DivisionID) & vbCrLf
    s = s & "   and i.vendor=" & DbQuote(Str, Vendor) & vbCrLf
    s = s & "   and i.invoice=" & DbQuote(Str, Invoice) & vbCrLf
    Set rs = HFApp.SqlExec(s, dbHomeFront)
    
    If Not rs.EOF Then
        Status = "" & rs("status")
        InvoiceInHomefront = True
    Else
        InvoiceInHomefront = False
    End If

End Function

Public Function InvoiceInAccounting(Vendor As String, Invoice As String) As Boolean
On Error Resume Next
    Dim s As String
    Dim rs As Recordset
    
    If HFApp.Options.Value(AccountingSystem) <> asTimberline Then
        InvoiceInAccounting = False
    Else
        'look in new file
        s = ""
        s = s & "select * " & vbCrLf
        s = s & "  from new_api_record_1" & vbCrLf
        If HFApp.ConnectionString(dbTimberlinePVdata) <> "" Then
            s = s & " where oivnd=" & DbQuote(Str, FormatTSField(10, Vendor, False, True)) & vbCrLf
            s = s & "   and oiinv=" & DbQuote(Str, FormatTSField(15, Invoice, False, True)) & vbCrLf
            Set rs = HFApp.SqlExec(s, dbTimberlinePVdata)
        Else
            s = s & " where oivnd=" & DbQuote(Str, Vendor) & vbCrLf
            s = s & "   and oiinv=" & DbQuote(Str, Invoice) & vbCrLf
            Set rs = HFApp.SqlExec(s, dbAccountingDictionary)
        End If
        If err.Number = 0 Then
        If Not rs.EOF Then
            InvoiceInAccounting = True
            Exit Function
        End If
        End If
        
        'look in master file
        s = ""
        s = s & "select * " & vbCrLf
        s = s & "  from master_apm_record_1" & vbCrLf
        If HFApp.ConnectionString(dbTimberlinePVdata) <> "" Then
            s = s & " where oivnd=" & DbQuote(Str, FormatTSField(10, Vendor, False, True)) & vbCrLf
            s = s & "   and oiinv=" & DbQuote(Str, FormatTSField(15, Invoice, False, True)) & vbCrLf
            Set rs = HFApp.SqlExec(s, dbTimberlinePVdata)
        Else
            s = s & " where oivnd=" & DbQuote(Str, Vendor) & vbCrLf
            s = s & "   and oiinv=" & DbQuote(Str, Invoice) & vbCrLf
            Set rs = HFApp.SqlExec(s, dbAccountingDictionary)
        End If
        InvoiceInAccounting = Not rs.EOF
    End If
End Function


Public Function ValidateCommitment(Commitment As String, Vendor As String) As String
    Dim s As String
    If TimberlineAccounting Then
        If AccountingDB = dbTimberlinePVdata Then
            s = ""
            s = s & "select sub ponumber,svendor vendor,if(svendor=" & DbQuote(Str, FormatTSField(10, Vendor, False, True)) & ",0,1) CrossPay, 0 WrapInsuranceExempt"
            s = s & "  from master_jcm_r12, master_jcm_r1"
            s = s & " where sjob=job and sclosed=0 and jstatus<>3 and sub=" & DbQuote(Str, FormatTSField(12, Commitment))
            If App.Options(AllowCrossPayingPOs) = False And Vendor <> "" Then
                s = s & " and svendor = " & DbQuote(Str, FormatTSField(10, Vendor, False, True))
            End If
        Else
            s = ""
            s = s & "select sub ponumber,svendor vendor,if(svendor=" & DbQuote(Str, Vendor) & ",0,1) CrossPay"
            s = s & "  from master_jcm_record_12, master_jcm_record_1_1"
            s = s & " where sjob=job and sclosed=0 and jstatus<>'Closed' and sub=" & DbQuote(Str, Commitment)
            If App.Options(AllowCrossPayingPOs) = False And Vendor <> "" Then
                s = s & " and svendor = " & DbQuote(Str, Vendor)
            End If
        End If
    Else
        s = ""
        s = s & "select p.ponumber,p.vendor,case p.vendor when " & DbQuote(Str, Vendor) & " then 0 else 1 end CrossPay, c.WrapInsuranceExempt" & vbCrLf
        s = s & "  from pomaster p" & vbCrLf
        s = s & "  left join tbljobs j on(p.DivisionID = j.DivisionID and p.job=j.job_no)" & vbCrLf
        s = s & "  left join tblLocality c on j.community=c.area" & vbCrLf
        s = s & " where p.DivisionID = " & HFApp.DivisionID & " and isnull(j.inactive,0)=0" & vbCrLf
        s = s & "   and ponumber=" & DbQuote(Str, Commitment)
        If App.Options(AllowCrossPayingPOs) = False And Vendor <> "" Then
            s = s & " and p.vendor = " & DbQuote(Str, Vendor)
        End If
    End If
    ValidateCommitment = s
End Function

Public Function ValidateCommitmentItem(Commitment As String, CommitmentItem As Long) As String
    Dim s As String
    If TimberlineAccounting Then
        s = ""
        s = s & "select svendor              CommitmentVendor" & vbCrLf
        s = s & "      ,isub                 Commitment" & vbCrLf
        s = s & "      ,item                 CommitmentItem" & vbCrLf
        s = s & "      ,idesc                CommitmentItemDesc" & vbCrLf
        s = s & "      ,job                  Job" & vbCrLf
        s = s & "      ,jdesc                JobDesc" & vbCrLf
        s = s & "      ,extra                Extra" & vbCrLf
        s = s & "      ,xdesc                ExtraDesc" & vbCrLf
        s = s & "      ,phase                Phase" & vbCrLf
        s = s & "      ,pdesc                PhaseDesc" & vbCrLf
        s = s & "      ,cat                  Category" & vbCrLf
        s = s & "      ,cdesc                CategoryDesc" & vbCrLf
        s = s & "      ,iunits               Quantity" & vbCrLf
        s = s & "      ,iunits-iuntinv       RemainingQuantity" & vbCrLf
        s = s & "      ,iuntcst              UnitPrice" & vbCrLf
        s = s & "      ,t.""group""            TaxGroup" & vbCrLf
        s = s & "      ,gdesc                TaxGroupDesc" & vbCrLf
        s = s & "      ,grate                TaxRate" & vbCrLf
        s = s & "      ,iamt+iappcoa-iamtinv Amount" & vbCrLf
        s = s & "      ,if(iretpct=0,sretpct,iretpct) RetainageRate" & vbCrLf
        s = s & "      ,''                   DebitAccount" & vbCrLf
        s = s & "      ,''                   DebitAccountDesc" & vbCrLf
        s = s & "  from master_jcm_record_13 " & vbCrLf
        s = s & "       left join master_jcm_record_12 on (isub = sub)" & vbCrLf
        s = s & "       left join master_jcm_record_1_1  on (ijob = job)" & vbCrLf
        s = s & "       left join master_jcm_record_2  on (ijob = xjob and iextra = extra)" & vbCrLf
        s = s & "       left join master_txm_record_2 t on (itxgrp = t.""group"")" & vbCrLf
        s = s & "       left join master_jcm_record_3  on (iphase = phase)" & vbCrLf
        s = s & "       left join master_jcm_record_4  on (icat = cat)" & vbCrLf
        If AccountingDB = dbTimberlinePVdata Then
            s = s & " where isub   = " & DbQuote(Str, FormatTSField(12, Commitment)) & vbCrLf
            s = s & "   and item   = " & DbQuote(num, CommitmentItem) & vbCrLf
        Else
            s = s & " where isub   = " & DbQuote(Str, Commitment) & vbCrLf
            s = s & "   and item   = " & DbQuote(num, CommitmentItem) & vbCrLf
        End If
    Else
        s = ""
        s = s & "select vendor                 CommitmentVendor" & vbCrLf
        s = s & "      ,po                     Commitment" & vbCrLf
        s = s & "      ,line                   CommitmentItem" & vbCrLf
        s = s & "      ,linedesc               CommitmentItemDesc" & vbCrLf
        s = s & "      ,linejob                Job" & vbCrLf
        s = s & "      ,linedesc               JobDesc" & vbCrLf
        s = s & "      ,lineextra              Extra" & vbCrLf
        s = s & "      ,lineextradesc          ExtraDesc" & vbCrLf
        s = s & "      ,linecostcode           Phase" & vbCrLf
        s = s & "      ,linecostcodedesc       PhaseDesc" & vbCrLf
        s = s & "      ,linecategory           Category" & vbCrLf
        s = s & "      ,linecategorydesc       CategoryDesc" & vbCrLf
        s = s & "      ,linecommittedquantity  Quantity" & vbCrLf
        s = s & "      ,lineremainingquantity  RemainingQuantity" & vbCrLf
        s = s & "      ,linecommittedunitprice UnitPrice" & vbCrLf
        s = s & "      ,linetaxgroup           TaxGroup" & vbCrLf
        s = s & "      ,linetaxgroupdesc       TaxGroupDesc" & vbCrLf
        s = s & "      ,linetaxrate            TaxRate" & vbCrLf
        s = s & "      ,lineremainingpretax+lineremainingtax Amount" & vbCrLf
        s = s & "      ,retainagerate          RetainageRate" & vbCrLf
        s = s & "      ,lineAccount            DebitAccount" & vbCrLf
        s = s & "      ,lineaccountDesc        DebitAccountDesc" & vbCrLf
        s = s & "  from jcpodetails" & vbCrLf
        s = s & " where DivisionID=" & DbQuote(num, HFApp.DivisionID) & vbCrLf
        s = s & "   and PO=" & DbQuote(Str, Commitment) & vbCrLf
        s = s & "   and Line=" & DbQuote(num, CommitmentItem) & vbCrLf
    End If
    ValidateCommitmentItem = s
End Function

Public Function GetBudget(Job As String, Extra As String, Phase As String, Optional Category As String = "") As String
    Dim s As String
    If TimberlineAccounting Then
        s = ""
        s = s & "SELECT SUM(cjtdc) Actual" & vbCrLf
        s = s & "      ,SUM(ctest) Budget" & vbCrLf
        s = s & "  FROM master_jcm_record_4" & vbCrLf
        If AccountingDB = dbTimberlinePVdata Then
            s = s & " WHERE cjob  =" & DbQuote(Str, FormatTSField(10, Job, True)) & vbCrLf
            s = s & "   AND cextra=" & DbQuote(Str, FormatTSField(10, Extra)) & vbCrLf
            s = s & "   AND cphase=" & DbQuote(Str, FormatTSField(12, Phase)) & vbCrLf
            If Category <> "" Then
                s = s & "   AND cat   =" & DbQuote(Str, FormatTSField(3, Category)) & vbCrLf
            End If
        
        Else
            s = s & " WHERE cjob  =" & DbQuote(Str, Job) & vbCrLf
            s = s & "   AND cextra=" & DbQuote(Str, Extra) & vbCrLf
            s = s & "   AND cphase=" & DbQuote(Str, Phase) & vbCrLf
            If Category <> "" Then
                s = s & "   AND cat   =" & DbQuote(Str, Category) & vbCrLf
            End If
        End If
    
    Else
        s = ""
        s = s & "select 0                 Actual" & vbCrLf
        s = s & "      ,sum(budgetpretax+budgetjctax) Budget" & vbCrLf
        s = s & "from estimateitems" & vbCrLf
        s = s & "where DivisionID = " & HFApp.DivisionID & " and job=" & DbQuote(Str, Job) & vbCrLf
        s = s & "and jcextra=" & DbQuote(Str, Extra) & vbCrLf
        s = s & "and jccostcode=" & DbQuote(Str, Phase) & vbCrLf
        If Category <> "" Then
            s = s & "   AND jccategory   =" & DbQuote(Str, Category) & vbCrLf
        End If
    End If
    GetBudget = s
End Function

Public Function GetCommitmentItemAmount(PONumber As String, LineNumber As Long) As String
    Dim s As String
    If TimberlineAccounting Then
        s = ""
        s = s & "select iamt+iappcoa-itxamt-iaptcoa Pretax" & vbCrLf
        s = s & "      ,iamtinv                     InvoicedTaxIn" & vbCrLf
        s = s & "      ,grate                       TaxRate" & vbCrLf
        s = s & "  from master_jcm_record_13 left outer join master_txm_record_2 t on itxgrp=t.""group""" & vbCrLf
        If AccountingDB = dbTimberlinePVdata Then
            s = s & " where isub=" & DbQuote(Str, FormatTSField(12, PONumber)) & vbCrLf
            s = s & "   and item=" & DbQuote(num, LineNumber) & vbCrLf
        Else
            s = s & " where isub=" & DbQuote(Str, PONumber) & vbCrLf
            s = s & "   and item=" & DbQuote(num, LineNumber) & vbCrLf
        End If
    Else
        s = ""
        s = s & "select p.pretax Pretax" & vbCrLf
        s = s & "      ,0 InvoicedTaxIn" & vbCrLf
        s = s & "      ,isnull(p.jctaxrate,0)+isnull(p.njctaxrate,0) TaxRate" & vbCrLf
        s = s & "from poitems p" & vbCrLf
        s = s & "     left outer join poinvoicedamounts x on(p.DivisionID = x.DivisionID and p.ponumber=x.ponumber and p.LineNumber=x.itemseq)" & vbCrLf
        s = s & "where p.ponumber=" & DbQuote(Str, PONumber) & vbCrLf
        s = s & "  and p.LineNumber=" & DbQuote(num, LineNumber) & vbCrLf
    End If
    GetCommitmentItemAmount = s
End Function



Public Function GetSelectedCommitments(ByVal WhereClause As String) As String
    Dim s As String
    If TimberlineAccounting Then
        s = ""
        s = s & "select distinct isub " & vbCrLf
        s = s & "  from master_jcm_record_13" & vbCrLf
        s = s & "  left join master_jcm_record_12 on (isub=sub)" & vbCrLf
        s = s & " where " & WhereClause & vbCrLf
    Else
        WhereClause = Replace(WhereClause, "isub=", "ponumber=")
        WhereClause = Replace(WhereClause, "item=", "LineNumber=")
    
        s = ""
        s = s & "select distinct ponumber " & vbCrLf
        s = s & "from poitems" & vbCrLf
        s = s & " where " & WhereClause & " and DivisionID =" & HFApp.DivisionID
    End If
    GetSelectedCommitments = s
End Function


Public Function GetSelectedCommitmentItems(ByVal WhereClause As String) As String
    Dim s As String
    If TimberlineAccounting Then
        s = ""
        s = s & "select svendor              CommitmentVendor" & vbCrLf
        s = s & "      ,isub                 Commitment" & vbCrLf
        s = s & "      ,item                 CommitmentItem" & vbCrLf
        s = s & "      ,idesc                CommitmentItemDesc" & vbCrLf
        s = s & "      ,job                  Job" & vbCrLf
        s = s & "      ,jaddr1               JobAddr" & vbCrLf
        s = s & "      ,jdesc                JobDesc" & vbCrLf
        s = s & "      ,extra                Extra" & vbCrLf
        s = s & "      ,xdesc                ExtraDesc" & vbCrLf
        s = s & "      ,phase                Phase" & vbCrLf
        s = s & "      ,pdesc                PhaseDesc" & vbCrLf
        s = s & "      ,scat                 Category" & vbCrLf
        s = s & "      ,scdesc               CategoryDesc" & vbCrLf
        s = s & "      ,''                   Account" & vbCrLf
        s = s & "      ,''                   AccountDesc" & vbCrLf
        s = s & "      ,ct.""group""         TaxGroup" & vbCrLf
        s = s & "      ,ct.gdesc             TaxGroupDesc" & vbCrLf
        If App.Options(RecalcPOTaxAtCurrentRate) Then
            s = s & "      ,dt.grate             TaxRate" & vbCrLf
        Else
            s = s & "      ,ct.grate             TaxRate" & vbCrLf
        End If
        s = s & "      ,iunits               CommittedQuantity" & vbCrLf
        s = s & "      ,iunits-iuntinv       RemainingQuantity" & vbCrLf
        s = s & "      ,iuntcst              CommittedUnitPrice" & vbCrLf
        s = s & "      ,iamt+iappcoa-iamtinv RemainingAmount" & vbCrLf
        s = s & "      ,if(iretpct=0,sretpct,iretpct) RetainageRate" & vbCrLf
        s = s & "      ,iamt+iappcoa OriginalAmount" & vbCrLf
        s = s & "      ,iamtinv      InvoicedAmount" & vbCrLf
        s = s & "  from master_jcm_record_13" & vbCrLf
        s = s & "       left join master_jcm_record_12 on (isub = sub)" & vbCrLf
        s = s & "       left join master_jcm_record_1_1  on (ijob = job)" & vbCrLf
        s = s & "       left join master_jcm_record_2  on (ijob = xjob and iextra = extra)" & vbCrLf
        s = s & "       left join master_txm_record_2 ct on (itxgrp = ct.""group"")" & vbCrLf
        s = s & "       left join master_jcm_record_3  on (ijob=pjob and iextra=pextra and iphase = phase)" & vbCrLf
        s = s & "       left join master_jcm_record_17 on (icat = scat)" & vbCrLf
        If App.Options(RecalcPOTaxAtCurrentRate) Then
            s = s & "      ,master_txm_record_2 dt" & vbCrLf
        End If
        s = s & " where iamt+iappcoa-iamtinv<>0" & vbCrLf
        s = s & "   and (" & WhereClause & ")"
        If App.Options(RecalcPOTaxAtCurrentRate) Then
            If AccountingDB = dbTimberlinePVdata Then
                s = s & "   and " & DbQuote(Str, FormatTSField(6, App.Options(DefaultTaxGroup))) & " = dt.""group""" & vbCrLf
            Else
                s = s & "   and " & DbQuote(Str, App.Options(DefaultTaxGroup)) & " = dt.""group""" & vbCrLf
            End If
        End If
        s = s & "order by iamt+iappcoa-iamtinv"
    Else
        
        s = ""
        s = s & "select " & vbCrLf
        s = s & " vendor                 CommitmentVendor" & vbCrLf
        s = s & ",po                     Commitment" & vbCrLf
        s = s & ",line                   CommitmentItem" & vbCrLf
        s = s & ",linedesc               CommitmentItemDesc" & vbCrLf
        s = s & ",linejob                Job" & vbCrLf
        s = s & ",linejobaddr            JobAddr" & vbCrLf
        s = s & ",linejobdesc            JobDesc" & vbCrLf
        s = s & ",lineextra              Extra" & vbCrLf
        s = s & ",lineextradesc          ExtraDesc" & vbCrLf
        s = s & ",linecostcode           Phase" & vbCrLf
        s = s & ",linecostcodedesc       PhaseDesc" & vbCrLf
        s = s & ",linecategory           Category" & vbCrLf
        s = s & ",linecategorydesc       CategoryDesc" & vbCrLf
        s = s & ",lineaccount            Account" & vbCrLf
        s = s & ",lineaccountdesc        AccountDesc" & vbCrLf
        s = s & ",linetaxgroup           TaxGroup" & vbCrLf
        s = s & ",linetaxgroupdesc       TaxGroupDesc" & vbCrLf
        s = s & ",linetaxrate            TaxRate" & vbCrLf
        s = s & ",linecommittedquantity  CommittedQuantity" & vbCrLf
        s = s & ",lineremainingquantity  RemainingQuantity" & vbCrLf
        s = s & ",linecommittedunitprice CommittedUnitPrice" & vbCrLf
        s = s & ",RetainageRate          RetainageRate" & vbCrLf
        s = s & ",linepretax+linetax     OriginalAmount" & vbCrLf
        s = s & ",lineinvoicedpretax+lineinvoicedtax InvoicedAmount" & vbCrLf
        s = s & ",lineremainingpretax+lineremainingtax RemainingAmount" & vbCrLf
        s = s & "from jcpodetails" & vbCrLf
        
        WhereClause = Replace(WhereClause, "isub=", "po=")
        WhereClause = Replace(WhereClause, "item=", "line=")
        s = s & " where " & WhereClause
            
        s = s & " order by lineremainingpretax+lineremainingtax" & vbCrLf
            
    End If
    GetSelectedCommitmentItems = s
End Function


Public Function SelectGLAccount() As String
    Dim s As String
    If TimberlineAccounting Then
        If HFApp.ConnectionString(dbTimberlinePVdata) <> "" Then
            s = "select AAcct Account,Atitle Description FROM master_glm_r1"
        Else
            s = "select AAcct Account,Atitle Description FROM master_glm_record_1"
        End If
    Else
        s = "SELECT Account ,Description FROM glaccounts where DivisionID = " & HFApp.DivisionID
    End If
    SelectGLAccount = s
End Function

Public Function SelectCommitmentItem(Commitment As String) As String
    Dim s As String
    If TimberlineAccounting Then
        s = ""
        s = s & "select Item Item" & vbCrLf
        s = s & "      ,idesc Description" & vbCrLf
        s = s & "      ,(iamt+iappcoa-iamtinv) * 100 / (grate+100) Remaining" & vbCrLf
        s = s & "      ,iamt+iappcoa-iamtinv - (iamt+iappcoa-iamtinv) * 100 / (grate+100) Tax" & vbCrLf
        s = s & "      ,ltrim(ijob) " & Quote(App.Options(Caption_Job)) & vbCrLf
        s = s & "      ,iextra " & Quote(App.Options(Caption_Extra)) & vbCrLf
        s = s & "      ,ltrim(iphase) " & Quote(App.Options(Caption_Phase)) & vbCrLf
        s = s & "      ,icat " & Quote(App.Options(Caption_Category)) & vbCrLf
        s = s & "      ,itxgrp " & Quote(App.Options(Caption_TaxGroup)) & vbCrLf
        s = s & "      ,iamtret " & Quote(App.Options(Caption_Retainage)) & vbCrLf
        If AccountingDB = dbTimberlinePVdata Then
            s = s & "  from master_jcm_r13 left outer join master_txm_r2 t on itxgrp=t.""ggroup""" & vbCrLf
            s = s & " where isub=" & DbQuote(Str, FormatTSField(12, Commitment))
        Else
            s = s & "  from master_jcm_record_13 left outer join master_txm_record_2 t on itxgrp=t.""group""" & vbCrLf
            s = s & " where isub=" & DbQuote(Str, Commitment)
        End If
    Else
        s = ""
        s = s & "select Line Item" & vbCrLf
        s = s & "      ,LineDesc Descriptoin" & vbCrLf
        s = s & "      ,LineRemainingPretax Remaining" & vbCrLf
        s = s & "      ,LineRemainingTax Tax" & vbCrLf
        s = s & "      ,LineJob " & Quote(App.Options(Caption_Job)) & vbCrLf
        s = s & "      ,LineExtra " & Quote(App.Options(Caption_Extra)) & vbCrLf
        s = s & "      ,LineCostCode " & Quote(App.Options(Caption_Phase)) & vbCrLf
        s = s & "      ,LineCategory " & Quote(App.Options(Caption_Category)) & vbCrLf
        s = s & "      ,LineTaxgroup " & Quote(App.Options(Caption_TaxGroup)) & vbCrLf
        s = s & "      ,RetainageRate " & Quote(App.Options(Caption_Retainage)) & vbCrLf
        s = s & "  from jcpodetails" & vbCrLf
        s = s & " where DivisionID=" & DbQuote(num, HFApp.DivisionID) & vbCrLf
        s = s & "   and PO=" & DbQuote(Str, Commitment) & vbCrLf
    End If
    SelectCommitmentItem = s

End Function



Public Function SelectExtra(Job As String) As String
    Dim s As String
    If TimberlineAccounting Then
        s = ""
        s = s & "SELECT LTRIM(EXTRA) " & Quote(App.Options(Caption_Extra)) & vbCrLf
        s = s & "      ,Xdesc Description" & vbCrLf
        If AccountingDB = dbTimberlinePVdata Then
            s = s & "  FROM master_jcm_r2" & vbCrLf
            s = s & " WHERE XJob=" & DbQuote(Str, FormatTSField(10, Job, True))
        Else
            s = s & "  FROM master_jcm_record_2" & vbCrLf
            s = s & " WHERE XJob=" & DbQuote(Str, Job)
        End If
    Else
        s = ""
        s = s & "SELECT Extra " & Quote(App.Options(Caption_Extra)) & vbCrLf
        s = s & "      ,Description" & vbCrLf
        s = s & "  FROM JobExtras" & vbCrLf
        s = s & " WHERE DivisionID = " & HFApp.DivisionID & " and Job=" & DbQuote(Str, Job)
    End If
    SelectExtra = s

End Function

Public Function StripFormating(s As String) As String
    Const FORMATCHRS = "+=_-)(*&^%$#@!~`[]{}\|/?'""<>;:,."
    Dim i As Long
    For i = 1 To Len(FORMATCHRS)
        s = Replace(s, Mid(FORMATCHRS, i, 1), "")
    Next
    StripFormating = s
End Function

Public Function SelectJobExtraCostCode(Optional Job As String = "", Optional Extra As String = "") As String
    Dim s As String
    If TimberlineAccounting Then
        If Job = "" Then
            'standards
            s = ""
            s = s & "SELECT DISTINCT" & vbCrLf
            s = s & "       LTRIM(sphase) " & Quote(App.Options(Caption_Phase)) & vbCrLf
            s = s & "      ,spdesc Description" & vbCrLf
            s = s & "  FROM master_jcm_record_16" & vbCrLf
            s = s & " WHERE spgphas=0"
        Else
            'job specific
            s = ""
            s = s & "SELECT DISTINCT" & vbCrLf
            s = s & "       LTRIM(phase) " & Quote(App.Options(Caption_Phase)) & vbCrLf
            s = s & "      ,pdesc Description" & vbCrLf
            s = s & "  FROM master_jcm_record_3" & vbCrLf
            s = s & " WHERE PJob=" & DbQuote(Str, Job) & vbCrLf
            s = s & "   AND pextra=" & DbQuote(Str, Extra) & vbCrLf
            s = s & "   AND pgphase=0"
        End If
    Else
        If Job = "" Then
            'standards
            s = ""
            s = s & "SELECT costcode " & Quote(App.Options(Caption_Phase)) & vbCrLf
            s = s & "      ,isnull(nullif(Description,''),costcode) Description" & vbCrLf
            s = s & "  FROM standardcostcodes where DivisionID = " & HFApp.DivisionID & vbCrLf
        Else
            'job specific
            s = ""
            s = s & "select costcode " & Quote(App.Options(Caption_Phase)) & vbCrLf
            s = s & "      ,isnull(nullif(Description,''),costcode) Description" & vbCrLf
            s = s & "  from jobextracostcodes" & vbCrLf
            s = s & " WHERE DivisionID = " & HFApp.DivisionID & " and job=" & DbQuote(Str, Job) & vbCrLf
            s = s & "   AND extra=" & DbQuote(Str, Extra) & vbCrLf
        End If
    End If
    SelectJobExtraCostCode = s
End Function
Public Function SelectJobExtraCostCodeCategory(Optional Job As String = "", Optional Extra As String = "", Optional CostCode As String) As String
    Dim s As String
    If TimberlineAccounting Then
        If Job = "" Then
            'standards
            s = ""
            s = s & "SELECT LTRIM(SCAT) " & Quote(App.Options(Caption_Category)) & vbCrLf
            s = s & "      ,SCDESC Description" & vbCrLf
            s = s & "  FROM master_jcm_record_17" & vbCrLf
        Else
            'job specific
            s = ""
            s = s & "SELECT ltrim(cat) " & Quote(App.Options(Caption_Category)) & vbCrLf
            s = s & "      ,cdesc Description" & vbCrLf
            s = s & "      ,ctest Estimated" & vbCrLf
            s = s & "      ,ctuest Units" & vbCrLf
            s = s & "      ,crvcom Committed" & vbCrLf
            s = s & "      ,cjtdc CostToDate" & vbCrLf
            s = s & "      ,cjtdu UnitsToDate" & vbCrLf
            s = s & "      ,ctest-cjtdc EstRemaining" & vbCrLf
            s = s & "      ,crvcom-cjtdc PoRemaining" & vbCrLf
            s = s & "  FROM master_jcm_record_4" & vbCrLf
            If AccountingDB = dbTimberlinePVdata Then
                s = s & " where cjob = " & DbQuote(Str, FormatTSField(10, Job, True)) & vbCrLf
                s = s & "   And cextra = " & DbQuote(Str, FormatTSField(10, Extra)) & vbCrLf
                s = s & "   And CPHASE = " & DbQuote(Str, FormatTSField(12, CostCode))
            Else
                s = s & " where cjob = " & DbQuote(Str, Job) & vbCrLf
                s = s & "   And cextra = " & DbQuote(Str, Extra) & vbCrLf
                s = s & "   And CPHASE = " & DbQuote(Str, CostCode)
            End If
        End If
    Else
        If Job = "" Then
            'standards
            s = ""
            s = s & "SELECT category " & Quote(App.Options(Caption_Category)) & vbCrLf
            s = s & "      ,Description" & vbCrLf
            s = s & "  FROM standardcategories where DivisionID = " & HFApp.DivisionID & vbCrLf
        Else
            'job specific
            s = ""
            s = s & "select distinct category " & Quote(App.Options(Caption_Category)) & vbCrLf
            s = s & "      ,Description" & vbCrLf
            s = s & "  from jobextracostcodecategories" & vbCrLf
            s = s & " WHERE DivisionID = " & HFApp.DivisionID & " and job=" & DbQuote(Str, Job) & vbCrLf
            s = s & "   AND extra=" & DbQuote(Str, Extra) & vbCrLf
            s = s & "   AND costcode=" & DbQuote(Str, CostCode) & vbCrLf
        End If
    End If
    SelectJobExtraCostCodeCategory = s
End Function



Public Sub GetFormatMasks()
    Dim rs As Recordset
    Dim s  As String
    Dim s1 As Long
    Dim s2 As Long
    Dim s3 As Long
    Dim s4 As Long
    Dim s5 As Long
    Dim sep1 As String
    Dim sep2 As String
    Dim sep3 As String
    Dim sep4 As String
    Dim m As String

    If HFApp.Options(AccountingSystem) = asTimberline And AccountingDB = dbTimberlinePVdata And HFApp.Databases(dbAccountingDictionary).State = adStateOpen Then
        s = "select section,cslen,cnpunct from ts_ctl_record_3 where section>0 and fldcode=9"
        'cant use pervasive connection ts_ctl tables are not exposed to pervasive.
        Set rs = HFApp.SqlExec(s, dbAccountingDictionary)
        If Not rs.EOF Then
            While Not rs.EOF
                Select Case Val("" & rs("section"))
                    Case 1: s1 = Val("" & rs("cslen")):    If "" & rs("cnpunct") <> "" Then sep1 = "" & rs("cnpunct")
                    Case 2: s2 = Val("" & rs("cslen")):    If "" & rs("cnpunct") <> "" Then sep2 = "" & rs("cnpunct")
                    Case 3: s3 = Val("" & rs("cslen")):    If "" & rs("cnpunct") <> "" Then sep3 = "" & rs("cnpunct")
                    Case 4: s4 = Val("" & rs("cslen")):    If "" & rs("cnpunct") <> "" Then sep4 = "" & rs("cnpunct")
                    Case 5: s5 = Val("" & rs("cslen"))   ' this is the suffix
                End Select
                rs.MoveNext
            Wend
            
            m = ""
            If s1 <> 0 Then m = m & String(s1, "&") & sep1
            If s2 <> 0 Then m = m & String(s2, "&") & sep2
            If s3 <> 0 Then m = m & String(s3, "&") & sep3
            If s4 <> 0 Then m = m & String(s4, "&")
            If s5 <> 0 Then m = m & sep4 & String(s5, "&")
            
            AccountMask = m
        End If
    End If

End Sub

Public Function FormatAccount(Account As String) As String
    If AccountMask = "" Then
        FormatAccount = Account
    Else
        FormatAccount = Format(StripFormating(Trim(Account)), AccountMask)
    End If
End Function


Public Function SelectVendor() As String
    Dim s As String
    
    If HFApp.Options(AccountingSystem) = asTimberline Then
        s = "SELECT Vendor_id Vendor, Vendor_Name Name FROM tblVendors WHERE DivisionID =" & HFApp.DivisionID & " and isnull(inactive,0)=0 order by vendor_name"
    Else
        s = "SELECT Vendor_Name Name, Vendor_id Vendor FROM tblVendors WHERE DivisionID =" & HFApp.DivisionID & " and isnull(inactive,0)=0 order by vendor_name"
    End If
    SelectVendor = s
    
End Function


Public Function WriteInvoice(InvoiceID As Long, RequireLienRelease As Boolean, _
                             Vendor, VendorType, VendorName, VendorAddr1, VendorAddr2, VendorCity, VendorProv, VendorPostal, _
                             Invoice, Job, JobDesc, Status, Approver, Comments, DeptID, Pretax, Tax, Discount, _
                             MiscDeductionAmt, MiscDeductionRate, DiscountDate, ReceivedDate, InvoiceDate, PaymentDate, AccountingDate, Description, _
                             InvoiceCode1, InvoiceCode2, Errors, IsRetainageInvoice, RetainageID) As Long
    Dim s As String
    
    s = ""
    If InvoiceID = 0 Then
        s = s & "INSERT INTO Invoices(DivisionID,Vendor,VendorType,VendorName,VendorAddr1,VendorAddr2,VendorCity,VendorProv,VendorPostal"
        s = s & "      ,Invoice,Job,JobDesc,Status,Approver,Comments,DeptID,PreTax,Tax,Discount,DiscountDate,ReceivedDate" & vbCrLf
        s = s & "      ,InvoiceDate,PaymentDate,AccountingDate,Description,Advance,MiscDeductionAmt,MiscDeductionRate" & vbCrLf
        s = s & "      ,InvoiceCode1,InvoiceCode2,Source,Errors,RetainageInvoice,RetainageID,UStmp,DStmp,TStmp)" & vbCrLf
        s = s & "VALUES(" & HFApp.DivisionID & vbCrLf
        s = s & "      ," & DbQuote(Str, Vendor) & vbCrLf
        s = s & "      ," & DbQuote(Str, VendorType) & vbCrLf
        s = s & "      ," & DbQuote(Str, VendorName) & vbCrLf
        s = s & "      ," & DbQuote(Str, VendorAddr1, , , 33) & vbCrLf
        s = s & "      ," & DbQuote(Str, VendorAddr2, , , 33) & vbCrLf
        s = s & "      ," & DbQuote(Str, VendorCity, , , 30) & vbCrLf
        s = s & "      ," & DbQuote(Str, VendorProv, , , 4) & vbCrLf
        s = s & "      ," & DbQuote(Str, VendorPostal, , , 10) & vbCrLf
        s = s & "      ," & DbQuote(Str, Invoice) & vbCrLf
        If HFApp.Options(AccountingSystem) = asTimberline Then
            s = s & "      ," & DbQuote(Str, StripFormating("" & Job)) & vbCrLf
        Else
            s = s & "      ," & DbQuote(Str, Job) & vbCrLf
        End If
        s = s & "      ," & DbQuote(Str, JobDesc) & vbCrLf
        
        If RequireLienRelease Then
            s = s & "      ,'Lien Hold'" & vbCrLf
        Else
            s = s & "      ," & DbQuote(Str, Status) & vbCrLf
        End If
        s = s & "      ," & DbQuote(Str, IIf(Status = "Approved", HFApp.LoginID, Approver)) & vbCrLf
        s = s & "      ," & DbQuote(Str, Comments) & vbCrLf
        s = s & "      ," & DbQuote(num, DeptID) & vbCrLf
        s = s & "      ," & DbQuote(num, Pretax) & vbCrLf
        s = s & "      ," & DbQuote(num, Tax) & vbCrLf
        s = s & "      ," & DbQuote(num, Discount) & vbCrLf
        s = s & "      ," & DbQuote(Date, DiscountDate) & vbCrLf
        s = s & "      ," & DbQuote(Date, ReceivedDate) & vbCrLf
        s = s & "      ," & DbQuote(Date, InvoiceDate) & vbCrLf
        s = s & "      ," & DbQuote(Date, PaymentDate) & vbCrLf
        s = s & "      ," & DbQuote(Date, AccountingDate) & vbCrLf
        s = s & "      ," & DbQuote(Str, Description) & vbCrLf
        s = s & "      ,0" & vbCrLf
        s = s & "      ," & DbQuote(num, MiscDeductionAmt) & vbCrLf
        s = s & "      ," & DbQuote(num, MiscDeductionRate) & vbCrLf
        s = s & "      ," & DbQuote(Str, InvoiceCode1) & vbCrLf
        s = s & "      ," & DbQuote(Str, InvoiceCode2) & vbCrLf
        s = s & "      ,'Desk Entry'" & vbCrLf
        s = s & "      ," & DbQuote(Str, Errors, , , 254) & vbCrLf
        s = s & "      ," & DbQuote(Bit, IsRetainageInvoice, , , 254) & vbCrLf
        s = s & "      ," & DbQuote(num, RetainageID, , , 254) & vbCrLf
        s = s & "      ," & DbQuote(Str, HFApp.LoginID) & vbCrLf
        s = s & "      ,getdate()" & vbCrLf
        s = s & "      ,getdate())" & vbCrLf
        Call HFApp.SqlExec(s)
        WriteInvoice = HFApp.SqlIdentity("Invoices")
    Else
        s = s & "UPDATE Invoices" & vbCrLf
        s = s & "SET Vendor=" & DbQuote(Str, Vendor) & vbCrLf
        s = s & "   ,VendorType=" & DbQuote(Str, VendorType) & vbCrLf
        s = s & "   ,VendorName=" & DbQuote(Str, VendorName) & vbCrLf
        s = s & "   ,VendorAddr1=" & DbQuote(Str, VendorAddr1, , , 33) & vbCrLf
        s = s & "   ,VendorAddr2=" & DbQuote(Str, VendorAddr2, , , 33) & vbCrLf
        s = s & "   ,VendorCity=" & DbQuote(Str, VendorCity, , , 30) & vbCrLf
        s = s & "   ,VendorProv=" & DbQuote(Str, VendorProv, , , 4) & vbCrLf
        s = s & "   ,VendorPostal=" & DbQuote(Str, VendorPostal, , , 10) & vbCrLf
        s = s & "   ,Invoice=" & DbQuote(Str, Invoice) & vbCrLf
        If HFApp.Options(AccountingSystem) = asTimberline Then
            s = s & "   ,Job=" & DbQuote(Str, StripFormating("" & Job)) & vbCrLf
        Else
            s = s & "   ,Job=" & DbQuote(Str, Job) & vbCrLf
        End If
        s = s & "   ,JobDesc=" & DbQuote(Str, JobDesc) & vbCrLf
        s = s & "   ,Comments=" & DbQuote(Str, Comments) & vbCrLf
        s = s & "   ,Status=" & DbQuote(Str, Status) & vbCrLf
        s = s & "   ,Approver=" & DbQuote(Str, IIf(Status = "Approved", HFApp.LoginID, Approver)) & vbCrLf
        s = s & "   ,DeptID=" & DbQuote(num, DeptID) & vbCrLf
        s = s & "   ,PreTax=" & DbQuote(num, Pretax) & vbCrLf
        s = s & "   ,Tax=" & DbQuote(num, Tax) & vbCrLf
        s = s & "   ,Discount=" & DbQuote(num, Discount) & vbCrLf
        s = s & "   ,DiscountDate=" & DbQuote(Date, DiscountDate) & vbCrLf
        s = s & "   ,ReceivedDate=" & DbQuote(Date, ReceivedDate) & vbCrLf
        s = s & "   ,InvoiceDate=" & DbQuote(Date, InvoiceDate) & vbCrLf
        s = s & "   ,PaymentDate=" & DbQuote(Date, PaymentDate) & vbCrLf
        s = s & "   ,AccountingDate=" & DbQuote(Date, AccountingDate) & vbCrLf
        s = s & "   ,Description=" & DbQuote(Str, Description) & vbCrLf
        s = s & "   ,Advance=0" & vbCrLf

        s = s & "   ,MiscDeductionAmt=" & DbQuote(num, MiscDeductionAmt) & vbCrLf
        s = s & "   ,MiscDeductionRate=" & DbQuote(num, MiscDeductionRate) & vbCrLf
        s = s & "   ,InvoiceCode1=" & DbQuote(Str, InvoiceCode1) & vbCrLf
        s = s & "   ,InvoiceCode2=" & DbQuote(Str, InvoiceCode2) & vbCrLf
        s = s & "   ,Errors=" & DbQuote(Str, Errors, , , 254) & vbCrLf
        
        'only change these if they are being added - not deleted. ie: user added retainage to existing invoice.
        If IsRetainageInvoice Then s = s & "   ,RetainageInvoice=" & DbQuote(Bit, IsRetainageInvoice, , , 254) & vbCrLf
        If RetainageID <> 0 Then s = s & "   ,RetainageID=" & DbQuote(num, RetainageID, , , 254) & vbCrLf
        
        s = s & "   ,UStmp=" & DbQuote(Str, HFApp.LoginID) & vbCrLf
        s = s & "   ,DStmp=CURDATE()" & vbCrLf
        s = s & "   ,TStmp=CURTIME()" & vbCrLf
        s = s & "WHERE InvoiceID=" & DbQuote(num, InvoiceID) & vbCrLf
        Call HFApp.SqlExec(s)
        WriteInvoice = InvoiceID
    End If
    
    
End Function

Public Function WriteInvoiceItem(InvoiceID As Long, ItemID As Long, _
                                 Vendor, Invoice, CommitmentVendor, Commitment, CommitmentDesc, _
                                 CommitmentItem, Job, JobDesc, Extra, ExtraDesc, Phase, _
                                 PhaseDesc, Category, CategoryDesc, DebitAccount, DebitAccountDesc, _
                                 Equipment, EquipmentDesc, EquipmentCostCode, EquipmentCostCodeDesc, _
                                 TaxGroup, TaxGroupDesc, TaxRate, CommittedQuantity, CommittedUnitPrice, _
                                 InvoicedQuantity, InvoicedUnitPrice, Pretax, Tax, Retainage, RetainageRate, Description, _
                                 Billable, WrapInsuranceExempt, JointPayee) As Long
        
    Dim s As String
    
    s = ""
    If ItemID = 0 Then
        s = s & "INSERT INTO InvoiceItems(DivisionID,InvoiceID,Vendor,Invoice,CommitmentVendor,Commitment,CommitmentDesc" & vbCrLf
        s = s & "                        ,CommitmentItem,Job,JobDesc,Extra,ExtraDesc,Phase" & vbCrLf
        s = s & "                        ,PhaseDesc,Category,CategoryDesc,DebitAccount,DebitAccountDesc" & vbCrLf
        s = s & "                        ,Equipment,EquipmentDesc,EquipmentCostCode,EquipmentCostCodeDesc" & vbCrLf
        s = s & "                        ,TaxGroup,TaxGroupDesc,TaxRate,CommittedQuantity,CommittedUnitPrice" & vbCrLf
        s = s & "                        ,InvoicedQuantity,InvoicedUnitPrice,PreTax,Tax,Retainage,RetainageRate,Description" & vbCrLf
        s = s & "                        ,Billable,WrapInsuranceExempt,UStmp,DStmp,TStmp,JointPayee)" & vbCrLf
        s = s & "VALUES(" & HFApp.DivisionID & vbCrLf
        s = s & "      ," & DbQuote(num, InvoiceID) & vbCrLf
        s = s & "      ," & DbQuote(Str, Vendor) & vbCrLf
        s = s & "      ," & DbQuote(Str, Invoice) & vbCrLf
        s = s & "      ," & DbQuote(Str, CommitmentVendor) & vbCrLf
        s = s & "      ," & DbQuote(Str, Commitment) & vbCrLf
        s = s & "      ," & DbQuote(Str, CommitmentDesc) & vbCrLf
        s = s & "      ," & DbQuote(num, CommitmentItem) & vbCrLf
        If HFApp.Options(AccountingSystem) = asTimberline Then
            s = s & "      ," & DbQuote(Str, StripFormating("" & Job)) & vbCrLf
        Else
            s = s & "      ," & DbQuote(Str, Job) & vbCrLf
        End If
        s = s & "      ," & DbQuote(Str, JobDesc) & vbCrLf
        s = s & "      ," & DbQuote(Str, Extra) & vbCrLf
        s = s & "      ," & DbQuote(Str, ExtraDesc) & vbCrLf
        s = s & "      ," & DbQuote(Str, Phase) & vbCrLf
        s = s & "      ," & DbQuote(Str, PhaseDesc) & vbCrLf
        s = s & "      ," & DbQuote(Str, Category) & vbCrLf
        s = s & "      ," & DbQuote(Str, CategoryDesc) & vbCrLf
        s = s & "      ," & DbQuote(Str, DebitAccount) & vbCrLf
        s = s & "      ," & DbQuote(Str, DebitAccountDesc) & vbCrLf
        s = s & "      ," & DbQuote(Str, Equipment) & vbCrLf
        s = s & "      ," & DbQuote(Str, EquipmentDesc) & vbCrLf
        s = s & "      ," & DbQuote(Str, EquipmentCostCode) & vbCrLf
        s = s & "      ," & DbQuote(Str, EquipmentCostCodeDesc) & vbCrLf
        s = s & "      ," & DbQuote(Str, TaxGroup) & vbCrLf
        s = s & "      ," & DbQuote(Str, TaxGroupDesc) & vbCrLf
        s = s & "      ," & DbQuote(num, TaxRate) & vbCrLf
        s = s & "      ," & DbQuote(num, CommittedQuantity) & vbCrLf
        s = s & "      ," & DbQuote(num, CommittedUnitPrice) & vbCrLf
        s = s & "      ," & DbQuote(num, InvoicedQuantity) & vbCrLf
        s = s & "      ," & DbQuote(num, InvoicedUnitPrice) & vbCrLf
        s = s & "      ," & DbQuote(num, Pretax) & vbCrLf
        s = s & "      ," & DbQuote(num, Tax) & vbCrLf
        s = s & "      ," & DbQuote(num, Retainage) & vbCrLf
        s = s & "      ," & DbQuote(num, RetainageRate) & vbCrLf
        s = s & "      ," & DbQuote(Str, Description) & vbCrLf
        s = s & "      ," & DbQuote(Bit, Billable) & vbCrLf
        s = s & "      ," & DbQuote(Bit, WrapInsuranceExempt) & vbCrLf
        s = s & "      ," & DbQuote(Str, HFApp.LoginID) & vbCrLf
        s = s & "      ,getdate()" & vbCrLf
        s = s & "      ,getdate()" & vbCrLf
        s = s & "      ," & DbQuote(Str, JointPayee) & vbCrLf
        s = s & ")" & vbCrLf
        Call HFApp.SqlExec(s)
        WriteInvoiceItem = HFApp.SqlIdentity("InvoiceItems")
    Else
        s = s & "UPDATE InvoiceItems" & vbCrLf
        s = s & "SET InvoiceID=" & DbQuote(num, InvoiceID) & vbCrLf
        s = s & "   ,Vendor=" & DbQuote(Str, Vendor) & vbCrLf
        s = s & "   ,Invoice=" & DbQuote(Str, Invoice) & vbCrLf
        s = s & "   ,CommitmentVendor=" & DbQuote(Str, CommitmentVendor) & vbCrLf
        s = s & "   ,Commitment=" & DbQuote(Str, Commitment) & vbCrLf
        s = s & "   ,CommitmentDesc=" & DbQuote(Str, CommitmentDesc) & vbCrLf
        s = s & "   ,CommitmentItem=" & DbQuote(num, CommitmentItem) & vbCrLf
        If HFApp.Options(AccountingSystem) = asTimberline Then
            s = s & "   ,Job=" & DbQuote(Str, StripFormating("" & Job)) & vbCrLf
        Else
            s = s & "   ,Job=" & DbQuote(Str, Job) & vbCrLf
        End If
        s = s & "   ,JobDesc=" & DbQuote(Str, JobDesc) & vbCrLf
        s = s & "   ,Extra=" & DbQuote(Str, Extra) & vbCrLf
        s = s & "   ,ExtraDesc=" & DbQuote(Str, ExtraDesc) & vbCrLf
        s = s & "   ,Phase=" & DbQuote(Str, Phase) & vbCrLf
        s = s & "   ,PhaseDesc=" & DbQuote(Str, PhaseDesc) & vbCrLf
        s = s & "   ,Category=" & DbQuote(Str, Category) & vbCrLf
        s = s & "   ,CategoryDesc=" & DbQuote(Str, CategoryDesc) & vbCrLf
        s = s & "   ,DebitAccount=" & DbQuote(Str, DebitAccount) & vbCrLf
        s = s & "   ,DebitAccountDesc=" & DbQuote(Str, DebitAccountDesc) & vbCrLf
        s = s & "   ,Equipment=" & DbQuote(Str, Equipment) & vbCrLf
        s = s & "   ,EquipmentDesc=" & DbQuote(Str, EquipmentDesc) & vbCrLf
        s = s & "   ,EquipmentCostCode=" & DbQuote(Str, EquipmentCostCode) & vbCrLf
        s = s & "   ,EquipmentCostCodeDesc=" & DbQuote(Str, EquipmentCostCodeDesc) & vbCrLf
        s = s & "   ,TaxGroup=" & DbQuote(Str, TaxGroup) & vbCrLf
        s = s & "   ,TaxGroupDesc=" & DbQuote(Str, TaxGroupDesc) & vbCrLf
        s = s & "   ,TaxRate=" & DbQuote(num, TaxRate) & vbCrLf
        s = s & "   ,CommittedQuantity=" & DbQuote(num, CommittedQuantity) & vbCrLf
        s = s & "   ,CommittedUnitPrice=" & DbQuote(num, CommittedUnitPrice) & vbCrLf
        s = s & "   ,InvoicedQuantity=" & DbQuote(num, InvoicedQuantity) & vbCrLf
        s = s & "   ,InvoicedUnitPrice=" & DbQuote(num, InvoicedUnitPrice) & vbCrLf
        s = s & "   ,PreTax=" & DbQuote(num, Pretax) & vbCrLf
        s = s & "   ,Tax=" & DbQuote(num, Tax) & vbCrLf
        s = s & "   ,Retainage=" & DbQuote(num, Retainage) & vbCrLf
        s = s & "   ,RetainageRate=" & DbQuote(num, RetainageRate) & vbCrLf
        s = s & "   ,Description=" & DbQuote(Str, Description) & vbCrLf
        s = s & "   ,Billable=" & DbQuote(Bit, Billable) & vbCrLf
        s = s & "   ,WrapInsuranceExempt=" & DbQuote(Bit, WrapInsuranceExempt) & vbCrLf
        s = s & "   ,Errors=''" & vbCrLf
        s = s & "   ,UStmp=" & DbQuote(Str, HFApp.LoginID) & vbCrLf
        s = s & "   ,DStmp=CURDATE()" & vbCrLf
        s = s & "   ,TStmp=CURTIME()" & vbCrLf
        s = s & "   ,JointPayee=" & DbQuote(Str, JointPayee) & vbCrLf
        s = s & "WHERE ItemID=" & DbQuote(num, ItemID) & vbCrLf
        Call HFApp.SqlExec(s)
        WriteInvoiceItem = ItemID
    End If
    
End Function


