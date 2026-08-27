Attribute VB_Name = "MIntacct"
Option Explicit
Option Compare Text
Const SRCFILE = "MIntacct::"





Public Function Intacct_Entities(CompanyID As String, uid As String, pwd As String, entity As String) As Recordset
    Dim s As String
    Dim Intacct As New IntacctWrapper.IntacctWrapper
    Call Intacct.OpenMessage(CompanyID, uid, pwd, entity)
    Call Intacct.GetEntities
    Call Intacct.CloseMessage
    s = Intacct.PostMessage(False)
    If s <> "" Then s = "exec Intacct_Entities " & DbQuote(Str, s)
    Set Intacct_Entities = HFApp.SqlExec(s, dbHomefront)
End Function
Public Function Intacct_EntityList(CompanyID As String, uid As String, pwd As String, entity As String) As String
    Dim s As String
    Dim Intacct As New IntacctWrapper.IntacctWrapper
    Call Intacct.OpenMessage(CompanyID, uid, pwd, entity)
    Call Intacct.GetEntities
    Call Intacct.CloseMessage
    s = Intacct.PostMessage(False)
    Intacct_EntityList = "exec Intacct_Entities " & DbQuote(Str, s) & " -- no order by" '<-- no order by is required by fpicklist
End Function




Public Function Intacct_TransactionTypes(CompanyID As String, uid As String, pwd As String, entity As String) As Recordset
    Dim s As String
    Dim Intacct As New IntacctWrapper.IntacctWrapper
    Call Intacct.OpenMessage(CompanyID, uid, pwd, entity)
    Call Intacct.GetTranTypes
    Call Intacct.CloseMessage
    s = Intacct.PostMessage(False)
    If s <> "" Then s = "exec Intacct_TransactionTypes " & DbQuote(Str, s)
    Set Intacct_TransactionTypes = HFApp.SqlExec(s, dbHomefront)
End Function

Public Function Intacct_EstimateTypes(CompanyID As String, uid As String, pwd As String, entity As String) As Recordset
    Dim s As String
    Dim Intacct As New IntacctWrapper.IntacctWrapper
    Call Intacct.OpenMessage(CompanyID, uid, pwd, entity)
    Call Intacct.GetEstimateTypes
    Call Intacct.CloseMessage
    s = Intacct.PostMessage(False)
    If s <> "" Then s = "exec Intacct_EstimateTypes " & DbQuote(Str, s)
    Set Intacct_EstimateTypes = HFApp.SqlExec(s, dbHomefront)
End Function

Public Function Intacct_Items(CompanyID As String, uid As String, pwd As String, entity As String) As Recordset
    Dim s As String
    Dim Intacct As New IntacctWrapper.IntacctWrapper
    Call Intacct.OpenMessage(CompanyID, uid, pwd, entity)
    Call Intacct.GetItems
    Call Intacct.CloseMessage
    s = Intacct.PostMessage(False)
    If s <> "" Then s = "exec Intacct_Items " & DbQuote(Str, s)
    Set Intacct_Items = HFApp.SqlExec(s, dbHomefront)
End Function


Public Sub WriteJobToIntacct(JobNumber As String)
On Error GoTo eh
    Dim X As New IntacctWrapper.IntacctWrapper

    Dim s As String
    Dim i As Integer
    Dim rs As Recordset
    Dim bIsConnecting As Boolean
    
    Dim entity As String
    
    Dim customFields As String
    
    Dim PostExtra As Boolean
    PostExtra = HFApp.Options.ValueByName("PostJCExtraAsSubJob") = "True"
    
    
    On Error Resume Next
    s = "exec Intacct_CustomFieldsValues " & DbQuote(Num, HFApp.DivisionID) & "," & DbQuote(Str, JobNumber)
    customFields = HFApp.SqlExec(s, dbHomefront)(0)
    On Error GoTo eh
    
    
    s = ""
    s = s & "select" & vbCrLf
    s = s & " j.job_no Job" & vbCrLf
    s = s & ",case when s.MultiUnitsAsExtras=1 then c.unit_no else '' end JCExtra" & vbCrLf
    s = s & ",isnull(nullif(j.Description,''),j.Job_no) Name" & vbCrLf
    s = s & ",isnull(m.description,'') + char(13) + char(10)" & vbCrLf
    s = s & "      + isnull(substring(j.municipal_address,1,30),'') + ' '" & vbCrLf
    s = s & "      + isnull(substring(j.municipal_address,31,60),'') + char(13) + char(10)" & vbCrLf
    s = s & "      + isnull(j.city,'') + char(13) + char(10)" & vbCrLf
    s = s & "      + isnull(j.province,'') + char(13) + char(10)" & vbCrLf
    s = s & "      + isnull(j.zip,'') Description" & vbCrLf
    s = s & ",arcustomer customer" & vbCrLf
    s = s & ",isnull(nullif(j.gl_prefix,''),l.Prefix1) as GL_Prefix" & vbCrLf
    s = s & ",l.IntacctParentJob,l.IntacctEntity,j.IntacctDepartment" & vbCrLf
    s = s & "FROM tbljobs j" & vbCrLf
    s = s & "JOIN system_setup s on s.id=j.divisionid" & vbCrLf
    s = s & "LEFT JOIN tblCustomers c ON j.divisionid=c.divisionid and j.job_no=c.job_no " & vbCrLf
    s = s & "LEFT JOIN customer_date cd on c.customer_no=cd.customer_no and isnull(cd.amountpaid,0)<>0" & vbCrLf
    s = s & "LEFT JOIN master_date md on cd.date_field=md.date_field and md.date_type=3" & vbCrLf
    s = s & "LEFT JOIN distinctmodelsbydivision m on isnull(nullif(c.Model,''),j.model)=m.model and j.divisionid=m.divisionid" & vbCrLf
    s = s & "LEFT JOIN tbllocality l on isnull(c.community,j.community)=l.area" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "WHERE j.job_no=" & DbQuote(Str, JobNumber) & vbCrLf
    s = s & "AND j.DivisionID = " & HFApp.DivisionID & vbCrLf
    s = s & "AND (" & vbCrLf
    s = s & "      (isnull(cd.amountpaid,0)<>0 " & vbCrLf
    s = s & "           OR (c.purchased=1 AND c.cancelled=0 AND c.inactive=0 AND ((s.accounting_approve=1 AND c.approved=1) OR (s.accounting_approve=0 AND c.contract_assigned=1)))" & vbCrLf
    s = s & "      )" & vbCrLf
    s = s & "      OR (isnull(c.home_selection,'') <>'PreSale' AND isnull(c.sold_to_Customer,'')='')" & vbCrLf
    s = s & "      OR (c.customer_no is null)" & vbCrLf
    s = s & "    )" & vbCrLf
    s = s & "order by c.unit_no" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    If rs.EOF Then Exit Sub
    
    entity = "" & rs("IntacctEntity")
    If entity = "" Then entity = HFApp.Options.ValueByName("IntacctEntity")

    bIsConnecting = True
    Call X.OpenMessage(HFApp.Options.ValueByName("IntacctCompanyID"), HFApp.Options.ValueByName("IntacctUID"), HFApp.Options.ValueByName("IntacctPWD"), entity)
    bIsConnecting = False
    
    While Not rs.EOF
        Call X.WriteJob(True, "" & rs("Job"), "" & rs("IntacctParentJob"), "" & rs("Name"), IIf(PostExtra, "" & rs("JCExtra"), ""), "" & rs("Description"), "" & rs("Customer"), "" & rs("GL_Prefix"), "" & rs("IntacctDepartment"), customFields)
        rs.MoveNext
    Wend
    
    X.CloseMessage
    Call WriteLogFile("intacct.writejobs.req.xml", X.xml())
    s = ""
    s = X.PostMessage(True)
    Call WriteLogFile("intacct.writejobs.res.xml", s)
    
    'i was going to try to reduce failed create api calls by only sending creates on new jobs but it
    'gets complicated when unit's are used. New tblcustomers with different unit numbers create additional
    'jobs. the initial simplified solution isn't enough.
    's = Parse(Parse(s, 2, "<RECORDNO>"), 1, "</RECORDNO>")
    
Exit Sub
eh: Select Case True
        Case bIsConnecting And Err.Description = "Object reference not set to an instance of an object."
            MsgBox "Unable to connect to Intacct entity """ & entity & """." & vbCrLf & vbCrLf & " Check your system settings.", vbExclamation, "Unable to post job"
        Case Else
            Call errHandler(SRCFILE & "WriteJobToIntacct", s)
    End Select
End Sub



Public Sub WriteCustomerToIntacct(JobNumber As String, CustomerNumber As String)
On Error GoTo eh
    Dim X As New IntacctWrapper.IntacctWrapper

    Dim i As Long
    Dim s As String
    Dim rs As Recordset
    Dim CNUM As String
    
    'select data from hf tables
    
    s = ""
    s = s & "SELECT DISTINCT" & vbCrLf
    s = s & " c.ar_customer_deposit Customer" & vbCrLf
    s = s & ",isnull(nullif(c.Description,''),c.customer_no) Description" & vbCrLf
    s = s & ",ISNULL(c.Customer_Name,'') + ISNULL(' '+c.Customer_LName,'') Contact" & vbCrLf
    s = s & ",c.Phone,c.cellphone,c.Fax,c.Email" & vbCrLf
    s = s & ",c.Address1,c.Address2,c.City,c.Province,c.Zip" & vbCrLf
    s = s & ",c.Comments" & vbCrLf
    s = s & "FROM tbljobs j" & vbCrLf
    s = s & "JOIN system_setup s on s.id=j.divisionid" & vbCrLf
    s = s & "LEFT JOIN tblCustomers c ON j.divisionid=c.divisionid and j.job_no=c.job_no " & vbCrLf
    s = s & "LEFT JOIN customer_date cd on c.customer_no=cd.customer_no and isnull(cd.amountpaid,0)<>0" & vbCrLf
    s = s & "LEFT JOIN master_date md on cd.date_field=md.date_field and md.date_type=3" & vbCrLf
    s = s & "LEFT JOIN distinctmodelsbydivision m on isnull(nullif(c.Model,''),j.model)=m.model and j.divisionid=m.divisionid" & vbCrLf
    s = s & "LEFT JOIN tbllocality l on isnull(c.community,j.community)=l.area" & vbCrLf
    
    s = s & "WHERE j.DivisionID = " & HFApp.DivisionID & vbCrLf
    s = s & "AND isnull(c.Job_no,'')<>''" & vbCrLf
    s = s & "AND isnull(c.ar_customer_deposit,'')<>''" & vbCrLf
    
    If CustomerNumber = "" Then
        s = s & "AND c.job_no=" & DbQuote(Str, JobNumber) & vbCrLf
    Else
        s = s & "AND c.Customer_no=" & DbQuote(Str, CustomerNumber) & vbCrLf
    End If
    
    s = s & "AND (" & vbCrLf
    s = s & "      (isnull(cd.amountpaid,0)<>0 " & vbCrLf
    s = s & "           OR (c.purchased=1 AND c.cancelled=0 AND c.inactive=0 AND ((s.accounting_approve=1 AND c.approved=1) OR (s.accounting_approve=0 AND c.contract_assigned=1)))" & vbCrLf
    s = s & "      )" & vbCrLf
    s = s & "      OR (isnull(c.home_selection,'') <>'PreSale' AND isnull(c.sold_to_Customer,'')='')" & vbCrLf
    s = s & "      OR (c.customer_no is null)" & vbCrLf
    s = s & "    )" & vbCrLf
    
    
    Set rs = HFApp.SqlExec(s)
    If rs.EOF Then Exit Sub
    
    
    Call X.OpenMessage(HFApp.Options.ValueByName("IntacctCompanyID"), HFApp.Options.ValueByName("IntacctUID"), HFApp.Options.ValueByName("IntacctPWD"), HFApp.Options.ValueByName("IntacctEntity"))
    While Not rs.EOF
        CNUM = "" & rs("Customer")
        Call X.WriteCustomer(HFApp.Options(Country), "" & rs("Customer"), "" & rs("Description"), "" & rs("Contact"), "" & rs("Phone"), "" & rs("CellPhone"), _
                          "" & rs("Fax"), "" & rs("Email"), "" & rs("Address1"), "" & rs("Address2"), "" & rs("City"), "" & rs("Province"), "" & rs("Zip"), _
                          "" & rs("Comments"))
        rs.MoveNext
    Wend
    X.CloseMessage
    Call WriteLogFile("intacct.writecustomers.req.xml", X.xml())
    
    s = X.PostMessage(False)
    Call WriteLogFile("intacct.writecustomers.res.xml", X.lastResponse)
        
        
Exit Sub
eh:
Select Case True
    Case Err.Description Like "*Another * already exists*"
        Resume Next
    Case Else
        Call errHandler(SRCFILE & "WriteCustomerToIntacct(Job='" & JobNumber & "', Customer='" & CNUM & "')", s)
End Select
End Sub
