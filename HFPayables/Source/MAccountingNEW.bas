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


Public Function ValidateJob(ByRef Job As String) As String
On Error GoTo err
    Dim s As String
    
    If TimberlineAccounting Then
        s = ""
        s = s & "SELECT job" & vbCrLf
        s = s & "      ,jdesc description,jdesc Municipal_address" & vbCrLf
        s = s & "  FROM master_jcm_record_1_1" & vbCrLf
        If AccountingDB = dbTimberlinePVdata Then
            s = s & " WHERE job=" & DbQuote(Str, FormatTSField(10, Job, True))
        Else
            s = s & " WHERE job=" & DbQuote(Str, HFApp.FormatJob(Job)) & vbCrLf
        End If
    Else
        s = ""
        s = s & "SELECT job_no job" & vbCrLf
        s = s & "      ,description,Municipal_address" & vbCrLf
        s = s & "  FROM tbljobs" & vbCrLf
        If HFApp.Options(AccountingSystem) = asTimberline Then
            s = s & " WHERE DivisionID = " & HFApp.DivisionID & " and job_no=" & DbQuote(Str, StripFormating(Job)) & vbCrLf
        Else
            s = s & " WHERE DivisionID = " & HFApp.DivisionID & " and job_no=" & DbQuote(Str, Job) & vbCrLf
        End If
    End If
    
    ValidateJob = s
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
            s = s & "select 1,cat Category ,cdesc Description" & vbCrLf
            s = s & "  from master_jcm_record_4 " & vbCrLf
            s = s & " where cjob=" & DbQuote(Str, FormatTSField(10, Job, True)) & vbCrLf
            s = s & "   and cextra=" & DbQuote(Str, Extra) & vbCrLf
            s = s & "   and cphase=" & DbQuote(Str, FormatTSField(12, Phase)) & vbCrLf
            s = s & "   And cat=" & DbQuote(Str, FormatTSField(3, Category)) & vbCrLf
            s = s & "union all" & vbCrLf
            s = s & "select 2,scat Category ,scdesc Description" & vbCrLf
            s = s & "  from master_jcm_record_17" & vbCrLf
            s = s & " where scat=" & DbQuote(Str, FormatTSField(3, Category)) & vbCrLf
            s = s & "order by 1" & vbCrLf
        
        Else
            s = ""
            s = s & "select 1,cat Category ,cdesc Description" & vbCrLf
            s = s & "  from master_jcm_record_4 " & vbCrLf
            s = s & " where cjob=" & DbQuote(Str, HFApp.FormatJob(Job)) & vbCrLf
            s = s & "   and cextra=" & DbQuote(Str, Extra) & vbCrLf
            s = s & "   and cphase=" & DbQuote(Str, HFApp.FormatCostCode(Phase)) & vbCrLf
            s = s & "   And cat=" & DbQuote(Str, Category) & vbCrLf
            s = s & "union all" & vbCrLf
            s = s & "select 2,scat Category ,scdesc Description" & vbCrLf
            s = s & "  from master_jcm_record_17" & vbCrLf
            s = s & " where scat=" & DbQuote(Str, Category) & vbCrLf
            s = s & "order by 1" & vbCrLf
        End If
    Else
        s = ""
        s = s & "select 1,category,description" & vbCrLf
        s = s & "  from JobExtraCostCodeCategories " & vbCrLf
        s = s & " where DivisionID =" & HFApp.DivisionID & " and job=" & DbQuote(Str, Job) & vbCrLf
        s = s & "   and extra=" & DbQuote(Str, Extra) & vbCrLf
        s = s & "   and costcode=" & DbQuote(Str, Phase) & vbCrLf
        s = s & "   and category=" & DbQuote(Str, Category) & vbCrLf
        s = s & "union all" & vbCrLf
        s = s & "select 2,category,description" & vbCrLf
        s = s & "  from standardcategories " & vbCrLf
        s = s & " where DivisionID = " & HFApp.DivisionID & " and category=" & DbQuote(Str, Category) & vbCrLf
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
        If App.Options(ShowJobTitle1) Then s = s & "      ,purchaser " & Quote(App.Options(Caption_JobTitle1)) & vbCrLf
        If App.Options(ShowJobTitle2) Then s = s & "      ,estimator " & Quote(App.Options(Caption_JobTitle2)) & vbCrLf
        s = s & "  FROM tbljobs" & vbCrLf
        s = s & " WHERE DivisionID = " & HFApp.DivisionID & " and isnull(inactive,0)=0 and isnull(isquote,0)=0 " & vbCrLf
        
        If IsIn("" & HFApp.Options(AccountingSystem), "" & asSimply, "" & asQuickBooks) Then
            s = s & " and isnull(ExternalJobID,'')<>''" & vbCrLf
        End If
        
    End If
    
    SelectJob = s
    
End Function

Public Function SelectPO(Optional vendor As String = "", Optional Job As String = "") As String
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
        s = s & "      ,samtinv AmtInvoiced" & vbCrLf
        s = s & "      ,samt+sapprco-samtinv AmtRemaining" & vbCrLf
        s = s & "      ,sjob " & App.Options(Caption_Job) & vbCrLf
        If AccountingDB = dbTimberlinePVdata Then
            s = s & "  from master_jcm_r12" & vbCrLf
            s = s & "  inner join master_apm_r9 on master_jcm_r12.svendor=master_apm_r9.vendor" & vbCrLf
            s = s & " where sclosed=0" & vbCrLf
            If (Not App.Options(AllowCrossPayingPOs)) And vendor <> "" Then
                If Val(vendor) = 0 Then
                    s = s & "   and vendor=" & DbQuote(Str, vendor) & vbCrLf
                Else
                    s = s & "   and vendor=" & DbQuote(Str, FormatTSField(10, vendor)) & vbCrLf
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
            If (Not App.Options(AllowCrossPayingPOs)) And vendor <> "" Then
                s = s & "   and vendor=" & DbQuote(Str, vendor) & vbCrLf
            End If
            If Job <> "" Then
                s = s & "   and sjob=" & DbQuote(Str, Job) & vbCrLf
            End If
            s = s & "   and samt+sapprco-samtinv<>0" & vbCrLf
        End If

    Else
        s = ""
        s = s & "select" & vbCrLf

        
        s = s & " p.PONumber        Commitment" & vbCrLf
        s = s & " ,p.Description     Description" & vbCrLf
        s = s & " ,p.Vendor          Vendor" & vbCrLf
        s = s & " ,v.vendor_name     VendorName" & vbCrLf
        s = s & " ,p.CompletedDate   Completed  " & vbCrLf
        s = s & " ,isnull(pt.PreTaxTotal+pt.JCTaxTotal+pt.NJCTaxTotal,0) OriginalAmt " & vbCrLf
        s = s & " ,isnull(pti.pretax + pti.tax,0) AmtInvoiced" & vbCrLf
        s = s & " ,isnull(pt.PreTaxTotal+pt.JCTaxTotal+pt.NJCTaxTotal,0)-isnull(pti.pretax + pti.tax,0) AmtRemaining" & vbCrLf
        s = s & " ,p.job Job" & vbCrLf
        
        s = s & "FROM" & vbCrLf
        
        s = s & "pomaster p" & vbCrLf
        s = s & " join pototal pt on(p.DivisionID = pt.DivisionID and p.ponumber=pt.ponumber)" & vbCrLf
        s = s & " left outer join pototalinvoiced pti on(p.DivisionID = pti.DivisionID and p.ponumber=pti.ponumber)" & vbCrLf
        s = s & " join tblvendors v on(p.DivisionID = v.DivisionID and p.vendor=v.vendor_id)" & vbCrLf
        s = s & "WHERE p.DivisionID = " & HFApp.DivisionID & " and p.cancelleddate is null" & vbCrLf

        s = s & " and pt.pretaxtotal - isnull(pti.pretax,0) <> 0" & vbCrLf
        

        
        If Not App.Options(AllowCrossPayingPOs) And vendor <> "" Then
            s = s & " and v.vendor_id=" & DbQuote(Str, vendor) & vbCrLf
        End If
        If Job <> "" Then
            s = s & " and p.job=" & DbQuote(Str, Job) & vbCrLf
        End If
        

        If HFApp.Options(AccountingSystem) = asTimberline Then
            s = s & " AND p.postingbatch<>0" & vbCrLf
        End If

    
    End If
    
    SelectPO = s
    
End Function

Public Function InvoiceInHomefront(vendor As String, invoice As String) As Boolean
    Dim s As String
    Dim rs As Recordset
    
    s = ""
    s = s & "select v.vendor_name,i.* " & vbCrLf
    s = s & "  from invoices i left outer join tblvendors v on i.DivisionID = v.DivisionID and i.vendor=v.vendor_id" & vbCrLf
    s = s & " where i.DivisionID=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
    s = s & "   and i.vendor=" & DbQuote(Str, vendor) & vbCrLf
    s = s & "   and i.invoice=" & DbQuote(Str, invoice) & vbCrLf
    Set rs = HFApp.SqlExec(s, dbHomeFront)
    InvoiceInHomefront = Not rs.EOF

End Function

Public Function InvoiceInAccounting(vendor As String, invoice As String) As Boolean
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
            s = s & " where oivnd=" & DbQuote(Str, FormatTSField(10, vendor, False, True)) & vbCrLf
            s = s & "   and oiinv=" & DbQuote(Str, FormatTSField(15, invoice, False, True)) & vbCrLf
            Set rs = HFApp.SqlExec(s, dbTimberlinePVdata)
        Else
            s = s & " where oivnd=" & DbQuote(Str, vendor) & vbCrLf
            s = s & "   and oiinv=" & DbQuote(Str, invoice) & vbCrLf
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
            s = s & " where oivnd=" & DbQuote(Str, FormatTSField(10, vendor, False, True)) & vbCrLf
            s = s & "   and oiinv=" & DbQuote(Str, FormatTSField(15, invoice, False, True)) & vbCrLf
            Set rs = HFApp.SqlExec(s, dbTimberlinePVdata)
        Else
            s = s & " where oivnd=" & DbQuote(Str, vendor) & vbCrLf
            s = s & "   and oiinv=" & DbQuote(Str, invoice) & vbCrLf
            Set rs = HFApp.SqlExec(s, dbAccountingDictionary)
        End If
        InvoiceInAccounting = Not rs.EOF
    End If
End Function


Public Function ValidateCommitment(Commitment As String, vendor As String) As String
    Dim s As String
    If TimberlineAccounting Then
        If AccountingDB = dbTimberlinePVdata Then
            s = ""
            s = s & "select sub ponumber,svendor vendor,if(svendor=" & DbQuote(Str, FormatTSField(10, vendor, False, True)) & ",0,1) CrossPay"
            s = s & "  from master_jcm_r12, master_jcm_r1"
            s = s & " where sjob=job and sclosed=0 and jstatus<>3 and sub=" & DbQuote(Str, FormatTSField(12, Commitment))
            If App.Options(AllowCrossPayingPOs) = False And vendor <> "" Then
                s = s & " and svendor = " & DbQuote(Str, FormatTSField(10, vendor, False, True))
            End If
        Else
            s = ""
            s = s & "select sub ponumber,svendor vendor,if(svendor=" & DbQuote(Str, vendor) & ",0,1) CrossPay"
            s = s & "  from master_jcm_record_12, master_jcm_record_1_1"
            s = s & " where sjob=job and sclosed=0 and jstatus<>'Closed' and sub=" & DbQuote(Str, Commitment)
            If App.Options(AllowCrossPayingPOs) = False And vendor <> "" Then
                s = s & " and svendor = " & DbQuote(Str, vendor)
            End If
        End If
    Else
        s = ""
        s = s & "select p.ponumber,p.vendor,case p.vendor when " & DbQuote(Str, vendor) & " then 0 else 1 end CrossPay" & vbCrLf
        s = s & "  from pomaster p" & vbCrLf
        s = s & "  left outer join tbljobs j on(p.DivisionID = j.DivisionID and p.job=j.job_no)" & vbCrLf
        s = s & " where p.DivisionID = " & HFApp.DivisionID & " and isnull(j.inactive,0)=0" & vbCrLf
        s = s & "   and ponumber=" & DbQuote(Str, Commitment)
        If App.Options(AllowCrossPayingPOs) = False And vendor <> "" Then
                s = s & " and p.vendor = " & DbQuote(Str, vendor)
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
        s = s & "      ,if(gtax1tt=1,gtax1rt,0)+if(gtax2tt=1,gtax2rt,0)+if(gtax3tt=1,gtax3rt,0)+if(gtax4tt=1,gtax4rt,0)+if(gtax5tt=1,gtax5rt,0)+if(gtax6tt=1,gtax6rt,0)+if(gtax7tt=1,gtax7rt,0)+if(gtax8tt=1,gtax8rt,0)+if(gtax9tt=1,gtax9rt,0) TaxRetainageRate" & vbCrLf
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
            s = s & "   and item   = " & DbQuote(Num, CommitmentItem) & vbCrLf
        Else
            s = s & " where isub   = " & DbQuote(Str, Commitment) & vbCrLf
            s = s & "   and item   = " & DbQuote(Num, CommitmentItem) & vbCrLf
        End If
    Else
        s = ""
        s = s & "select p.vendor             CommitmentVendor" & vbCrLf
        s = s & "      ,p.ponumber           Commitment" & vbCrLf
        s = s & "      ,i.linenumber         CommitmentItem" & vbCrLf
        s = s & "      ,i.description        CommitmentItemDesc" & vbCrLf
        s = s & "      ,i.job                Job" & vbCrLf
        s = s & "      ,j.description        JobDesc" & vbCrLf
        s = s & "      ,i.jcextra            Extra" & vbCrLf
        s = s & "      ,i.jcextra            ExtraDesc" & vbCrLf
        s = s & "      ,i.jccostcode         Phase" & vbCrLf
        s = s & "      ,c.description        PhaseDesc" & vbCrLf
        s = s & "      ,i.jccategory         Category" & vbCrLf
        s = s & "      ,t.description        CategoryDesc" & vbCrLf
        s = s & "      ,i.orderqty           Quantity" & vbCrLf
        s = s & "      ,isnull(i.orderqty,0)-isnull(x.qty,0) RemainingQuantity" & vbCrLf
        s = s & "      ,i.rate               UnitPrice" & vbCrLf
        s = s & "      ,i.taxgroup           TaxGroup" & vbCrLf
        s = s & "      ,g.description        TaxGroupDesc" & vbCrLf
        s = s & "      ,isnull(i.jctaxrate,0) + isnull(i.njctaxrate,0) TaxRate" & vbCrLf
        s = s & "      ,i.pretax+i.jctax+i.njctax-isnull(x.pretax,0)-isnull(x.tax,0) Amount" & vbCrLf
        s = s & "      ,p.retainagepercent   RetainageRate" & vbCrLf
        s = s & "      ,g.RetainageRate      TaxRetainageRate" & vbCrLf
        s = s & "      ,a.Account            DebitAccount" & vbCrLf
        s = s & "      ,a.Description        DebitAccountDesc" & vbCrLf
        s = s & "from poitems i" & vbCrLf
        s = s & "     left outer join pomaster p on(p.DivisionID = i.DivisionID and p.ponumber=i.ponumber)" & vbCrLf
        s = s & "     left outer join poinvoicedamounts x on(i.DivisionID = x.DivisionID and i.ponumber=x.ponumber and i.LineNumber=x.itemseq)" & vbCrLf
        s = s & "     left outer join tbljobs j on(i.DivisionID = j.DivisionID and i.job=j.job_no)" & vbCrLf
        s = s & "     left outer join standardcostcodes c on(i.DivisionID = c.DivisionID and i.jccostcode=c.costcode)" & vbCrLf
        s = s & "     left outer join glaccounts a on(c.DivisionID = a.DivisionID and c.debitaccount=a.account)" & vbCrLf
        s = s & "     left outer join standardcategories t on(i.DivisionID = t.DivisionID and i.jccategory=t.category)" & vbCrLf
        s = s & "     left outer join taxgroups g on(i.DivisionID = g.DivisionID and i.taxgroup=g.taxgroup)" & vbCrLf
        s = s & " where i.DivisionID = " & HFApp.DivisionID & " and i.ponumber = " & DbQuote(Str, Commitment) & vbCrLf
        s = s & "   and i.linenumber = " & DbQuote(Num, CommitmentItem) & vbCrLf
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
        s = s & "      ,sum(budgetpretax) Budget" & vbCrLf
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
            s = s & "   and item=" & DbQuote(Num, LineNumber) & vbCrLf
        Else
            s = s & " where isub=" & DbQuote(Str, PONumber) & vbCrLf
            s = s & "   and item=" & DbQuote(Num, LineNumber) & vbCrLf
        End If
    Else
        s = ""
        s = s & "select p.pretax Pretax" & vbCrLf
        s = s & "      ,0 InvoicedTaxIn" & vbCrLf
        s = s & "      ,isnull(p.jctaxrate,0)+isnull(p.njctaxrate,0) TaxRate" & vbCrLf
        s = s & "from poitems p" & vbCrLf
        s = s & "     left outer join poinvoicedamounts x on(p.DivisionID = x.DivisionID and p.ponumber=x.ponumber and p.LineNumber=x.itemseq)" & vbCrLf
        s = s & "where p.ponumber=" & DbQuote(Str, PONumber) & vbCrLf
        s = s & "  and p.LineNumber=" & DbQuote(Num, LineNumber) & vbCrLf
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
        If App.Options(DefaultTaxGroup) <> "" Then
            s = s & "      ,dt.grate             DefaultTaxRate" & vbCrLf
        Else
            s = s & "      ,ct.grate             DefaultTaxRate" & vbCrLf
        End If
        s = s & "      ,ct.grate             CommittedTaxRate" & vbCrLf
        s = s & "      ,iunits               CommittedQuantity" & vbCrLf
        s = s & "      ,iunits-iuntinv       RemainingQuantity" & vbCrLf
        s = s & "      ,iuntcst              CommittedUnitPrice" & vbCrLf
        s = s & "      ,iamt+iappcoa-iamtinv Amount" & vbCrLf
        s = s & "      ,if(iretpct=0,sretpct,iretpct) RetainageRate" & vbCrLf
        If App.Options(RecalcPOTaxAtCurrentRate) Then
            s = s & "      ,if(dt.gtax1tt=1,dt.gtax1rt,0)+if(dt.gtax2tt=1,dt.gtax2rt,0)+if(dt.gtax3tt=1,dt.gtax3rt,0)+if(dt.gtax4tt=1,dt.gtax4rt,0)+if(dt.gtax5tt=1,dt.gtax5rt,0)+if(dt.gtax6tt=1,dt.gtax6rt,0)+if(dt.gtax7tt=1,dt.gtax7rt,0)+if(dt.gtax8tt=1,dt.gtax8rt,0)+if(dt.gtax9tt=1,dt.gtax9rt,0) TaxRetainageRate" & vbCrLf
        Else
            s = s & "      ,if(ct.gtax1tt=1,ct.gtax1rt,0)+if(ct.gtax2tt=1,ct.gtax2rt,0)+if(ct.gtax3tt=1,ct.gtax3rt,0)+if(ct.gtax4tt=1,ct.gtax4rt,0)+if(ct.gtax5tt=1,ct.gtax5rt,0)+if(ct.gtax6tt=1,ct.gtax6rt,0)+if(ct.gtax7tt=1,ct.gtax7rt,0)+if(ct.gtax8tt=1,ct.gtax8rt,0)+if(ct.gtax9tt=1,ct.gtax9rt,0) TaxRetainageRate" & vbCrLf
        End If
        s = s & "      ,iamt+iappcoa IRAmount" & vbCrLf
        s = s & "      ,iamtinv      IRInvoiced" & vbCrLf
        s = s & "  from master_jcm_record_13" & vbCrLf
        s = s & "       left join master_jcm_record_12 on (isub = sub)" & vbCrLf
        s = s & "       left join master_jcm_record_1_1  on (ijob = job)" & vbCrLf
        s = s & "       left join master_jcm_record_2  on (ijob = xjob and iextra = extra)" & vbCrLf
        s = s & "       left join master_txm_record_2 ct on (itxgrp = ct.""group"")" & vbCrLf
        s = s & "       left join master_jcm_record_3  on (ijob=pjob and iextra=pextra and iphase = phase)" & vbCrLf
        s = s & "       left join master_jcm_record_17 on (icat = scat)" & vbCrLf
        If App.Options(DefaultTaxGroup) <> "" Then
            s = s & "      ,master_txm_record_2 dt" & vbCrLf
        End If
        s = s & " where iamt+iappcoa-iamtinv<>0" & vbCrLf
        s = s & "   and (" & WhereClause & ")"
        If App.Options(DefaultTaxGroup) <> "" Then
            If AccountingDB = dbTimberlinePVdata Then
                s = s & "   and " & DbQuote(Str, FormatTSField(6, App.Options(DefaultTaxGroup))) & " = dt.""group""" & vbCrLf
            Else
                s = s & "   and " & DbQuote(Str, App.Options(DefaultTaxGroup)) & " = dt.""group""" & vbCrLf
            End If
        End If
        s = s & "order by iamt+iappcoa-iamtinv"
    Else
        WhereClause = Replace(WhereClause, "isub=", "commitment=")
        WhereClause = Replace(WhereClause, "item=", "commitmentLineNumber=")
    
        s = ""
        s = s & "select * from jcpodetails" & vbCrLf
        s = s & "where 1=1 and " & WhereClause & vbCrLf
        s = s & "order by Commitment, CommitmentItem" & vbCrLf

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
        s = s & "select i.LineNumber Item" & vbCrLf
        s = s & "      ,i.Description" & vbCrLf
        s = s & "      ,isnull(i.pretax,0)-isnull(x.pretax,0) Remaining" & vbCrLf
        s = s & "      ,isnull(i.jctax,0)+isnull(i.njctax,0)-isnull(x.tax,0) Tax" & vbCrLf
        s = s & "      ,i.job " & Quote(App.Options(Caption_Job)) & vbCrLf
        s = s & "      ,i.jcextra " & Quote(App.Options(Caption_Extra)) & vbCrLf
        s = s & "      ,i.jccostcode " & Quote(App.Options(Caption_Phase)) & vbCrLf
        s = s & "      ,i.jccategory " & Quote(App.Options(Caption_Category)) & vbCrLf
        s = s & "      ,i.taxgroup " & Quote(App.Options(Caption_TaxGroup)) & vbCrLf
        s = s & "      ,p.retainagepercent " & Quote(App.Options(Caption_Retainage)) & vbCrLf
        s = s & "  from pomaster p" & vbCrLf
        s = s & "       join poitems i on(p.DivisionID = i.DivisionID and p.ponumber=i.ponumber)" & vbCrLf
        s = s & "       left outer join poinvoicedamounts x on(i.DivisionID = x.DivisionID and i.ponumber=x.ponumber and i.LineNumber=x.itemseq)" & vbCrLf
        s = s & "where p.DivisionID = " & HFApp.DivisionID & " and p.ponumber=" & DbQuote(Str, Commitment)
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
            s = s & "select category " & Quote(App.Options(Caption_Category)) & vbCrLf
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

    If HFApp.Options(AccountingSystem) = asTimberline And HFApp.Databases(dbAccountingDictionary).State = adStateOpen Then
        Set rs = HFApp.SqlExec("select section,cslen,cnpunct from ts_ctl_record_3 where section>0 and fldcode=9", dbAccountingDictionary)
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
    
    s = ""
    s = s & "select vendor_id Vendor" & vbCrLf
    s = s & "      ,vendor_name Name" & vbCrLf
    s = s & "      ,City" & vbCrLf
    s = s & "from tblvendors" & vbCrLf
    s = s & "where DivisionID = " & HFApp.DivisionID & " and isnull(inactive,0)=0"
    SelectVendor = s
    
End Function

