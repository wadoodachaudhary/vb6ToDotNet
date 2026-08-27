Attribute VB_Name = "MBuildPro"
Option Explicit
Const SRCFILE = "MBuildPro::"

Private Function BuildProEnabled() As Boolean
    BuildProEnabled = HFApp.Options.ValueByName("BuildProCompanyCode") <> ""
End Function

Private Function GetImpDate(Name As String) As Date
On Error Resume Next
    Dim dt As Date

    dt = "1970-01-01"
    dt = CDate(HFApp.Options.ValueByName(Name))
    If dt < DateValue("1970-01-01") Then dt = DateValue("1970-01-01")
    GetImpDate = dt

End Function

Private Sub PutImpDate(Name As String)
On Error Resume Next

    HFApp.Options.ValueByName(Name) = Now()

End Sub

Private Sub LogXML(Name As String, XmlMessage As String)
    Dim i As Integer
    Dim p As String
    i = FreeFile()
    p = PathAppend(GetFolderPath(CSIDL_LOCAL_APPDATA), "HomeFront", VB.App.EXEName)
    Call CreatePath("", p)
    Open PathAppend(p, "buildpro_" & Name & ".xml") For Output As #i
    Print #i, XmlMessage
    Close #i
End Sub


Public Sub SendBuildProPOIndexes()
On Error GoTo eh
    Dim s As String
    Dim rs As Recordset
    Dim X As New HyphenSys.BuildProWrapper
    Dim divCode As String

    If Not BuildProEnabled Then Exit Sub

    Screen.MousePointer = vbHourglass

    divCode = HFApp.SqlExec("select divisioncode from divisions where divisionid=" & DbQuote(Num, HFApp.DivisionID))(0)
    Call X.OpenMessage(HFApp.Options.ValueByName("BuildProCompanyCode"), divCode, HFApp.Options.ValueByName("BuildProURL"), HFApp.Options.ValueByName("BuildProUID"), HFApp.Options.ValueByName("BuildProPwd"))
    X.OpenPOIndexes

    s = ""
    s = s & "select * from BuildPro_POIndexes " & vbCrLf
    s = s & "where divisionid=" & DbQuote(Num, HFApp.DivisionID)
    Set rs = HFApp.SqlExec(s, dbHomefront)
    While Not rs.EOF
        Call X.POIndex("" & rs("POIndex"), "" & rs("CostType"), "" & rs("Description"))
        rs.MoveNext
    Wend
    X.ClosePOIndexes
    X.CloseMessage
    LogXML "poindexes", X.xml
    X.PostMessage
    Screen.MousePointer = vbNormal


Exit Sub
eh: Call errHandler(SRCFILE & "SendBuildProPOIndexes")
End Sub


Public Sub SendBuildProVendors()
On Error GoTo eh
    Dim s As String
    Dim rs As Recordset
    Dim X As New HyphenSys.BuildProWrapper
    Dim divCode As String

    If Not BuildProEnabled Then Exit Sub

    Screen.MousePointer = vbHourglass

    divCode = HFApp.SqlExec("select divisioncode from divisions where divisionid=" & DbQuote(Num, HFApp.DivisionID))(0)

    Call X.OpenMessage(HFApp.Options.ValueByName("BuildProCompanyCode"), divCode, HFApp.Options.ValueByName("BuildProURL"), HFApp.Options.ValueByName("BuildProUID"), HFApp.Options.ValueByName("BuildProPwd"))
    X.OpenVendors

    s = ""
    s = s & "select * from dbo.BuildPro_Vendors" & vbCrLf
    s = s & "Where DivisionID = " & DbQuote(Num, HFApp.DivisionID)
    Set rs = HFApp.SqlExec(s, dbHomefront)

    s = ""
    While Not rs.EOF
        Call X.vendor("" & rs("Vendor_ID") _
                    , "" & rs("Vendor_Name") _
                    , "" _
                    , "" & rs("Addr1") _
                    , "" & rs("Addr2") _
                    , "" & rs("City") _
                    , "" & rs("State") _
                    , "" & rs("PostalCode") _
                    , "" & rs("Phone") _
                    , "" & rs("Fax") _
                    , "" & rs("Email") _
                    , "" & rs("SchedEmail") _
                    , "" & rs("ServiceEmail") _
                    , rs("tbd") _
                    , rs("active"))

        s = s & "update tblVendors set DateSentToBuildPro=getdate() where divisionid=" & DbQuote(Num, HFApp.DivisionID) & " and Vendor_ID=" & DbQuote(Str, "" & rs("Vendor_ID"))

        rs.MoveNext
    Wend

    If s <> "" Then
        X.CloseVendors
        X.CloseMessage
        LogXML "vendors", X.xml
        If X.PostMessage Then Call HFApp.SqlExec(s, dbHomefront)
    End If
    Screen.MousePointer = vbNormal


Exit Sub
eh: Call errHandler(SRCFILE & "SendBuildProVendors")
End Sub

Public Sub SendBuildProChecks()
On Error GoTo eh
    Dim s As String
    Dim rs As Recordset
    Dim X As New HyphenSys.BuildProWrapper
    Dim divCode As String
    Dim lastDoc As String

    If Not BuildProEnabled Then Exit Sub
    Screen.MousePointer = vbHourglass

    'for payments which cross divisions BP needs to be configured with a 3rd AP only division.
    divCode = Trim(HFApp.Options.ValueByName("BuildProAPDivCode"))
    If divCode = "" Then divCode = HFApp.SqlExec("select divisioncode from divisions where divisionid=" & DbQuote(Num, HFApp.DivisionID))(0)

    s = ""
    s = s & "select * from dbo.BuildPro_RemittanceAdvice_Checks" & vbCrLf
    s = s & "Where DivisionID = " & DbQuote(Num, HFApp.DivisionID)
    s = s & "and DateSentToBuildPro is null" & vbCrLf
    s = s & "order by ChequeNumber"
    Set rs = HFApp.SqlExec(s, dbHomefront)
    While Not rs.EOF

        If lastDoc <> "" & rs("ChequeNumber") Then

            If lastDoc <> "" Then
                Call X.CloseRemittanceDoc
                Call X.CloseRemittanceAdvice
                Call X.CloseMessage
                LogXML "RemittanceChecks", X.xml
                If X.PostMessage Then
                    s = "update AccountingAPPayments set DateSentToBuildPro=getdate() where divisionid=" & DbQuote(Num, HFApp.DivisionID) & " and ChequeNumber=" & DbQuote(Str, lastDoc)
                    Call HFApp.SqlExec(s)
                End If
            End If

            lastDoc = "" & rs("ChequeNumber")
            Call X.OpenMessage(HFApp.Options.ValueByName("BuildProCompanyCode"), divCode, HFApp.Options.ValueByName("BuildProURL"), HFApp.Options.ValueByName("BuildProUID"), HFApp.Options.ValueByName("BuildProPwd"))
            Call X.OpenRemittanceAdvice
            Call X.AddRemittanceDoc_Cheque("" & rs("ActionCode"), "" & rs("IssueDate"), "" & rs("ChequeNumber"), "" & rs("ChequeAmount"), "" & rs("ChequeDate"), "" & rs("ChequeType"), "" & rs("Vendor"))

        End If

        Call X.AddRemittanceDoc_Line("" & rs("Job"), "" & rs("PONumber"), "" & rs("InvoiceNumber"), "" & rs("InvoiceDate"), "" & rs("AmountPaid"), "" & rs("POIndex"), "" & rs("Status"))

        rs.MoveNext
    Wend

    If lastDoc <> "" Then
        Call X.CloseRemittanceDoc
        Call X.CloseRemittanceAdvice
        Call X.CloseMessage
        LogXML "RemittanceChecks", X.xml
        If X.PostMessage Then
            s = "update AccountingAPPayments set DateSentToBuildPro=getdate() where divisionid=" & DbQuote(Num, HFApp.DivisionID) & " and ChequeNumber=" & DbQuote(Str, lastDoc)
            Call HFApp.SqlExec(s)
        End If
    End If

    Screen.MousePointer = vbNormal


Exit Sub
eh: Call errHandler(SRCFILE & "SendBuildProChecks")
End Sub

Public Sub SendBuildProVouchers()
On Error GoTo eh
    Dim s As String
    Dim rs As Recordset
    Dim X As New HyphenSys.BuildProWrapper
    Dim divCode As String
    Dim lastDoc As String

    If Not BuildProEnabled Then Exit Sub
    Screen.MousePointer = vbHourglass

    divCode = HFApp.SqlExec("select divisioncode from divisions where divisionid=" & DbQuote(Num, HFApp.DivisionID))(0)

    s = ""
    s = s & "select * from dbo.BuildPro_RemittanceAdvice_Vouchers" & vbCrLf
    s = s & "Where DivisionID = " & DbQuote(Num, HFApp.DivisionID) & vbCrLf
    s = s & "and DateSentToBuildPro is null" & vbCrLf
    s = s & "order by Voucher"
    Set rs = HFApp.SqlExec(s, dbHomefront)
    While Not rs.EOF

        If lastDoc <> "" & rs("Voucher") Then

            If lastDoc <> "" Then
                Call X.CloseRemittanceDoc
                Call X.CloseRemittanceAdvice
                Call X.CloseMessage
                LogXML "RemittanceVouchers", X.xml
                If X.PostMessage Then
                    s = "update LienVouchers set DateSentToBuildPro=getdate() where Voucher=" & DbQuote(Str, lastDoc)
                    Call HFApp.SqlExec(s)
                End If
            End If

            lastDoc = "" & rs("Voucher")
            Call X.OpenMessage(HFApp.Options.ValueByName("BuildProCompanyCode"), divCode, HFApp.Options.ValueByName("BuildProURL"), HFApp.Options.ValueByName("BuildProUID"), HFApp.Options.ValueByName("BuildProPwd"))
            Call X.OpenRemittanceAdvice
            Call X.AddRemittanceDoc_Voucher("" & rs("ActionCode"), "" & rs("IssueDate"), "" & rs("Voucher"), "" & rs("VoucherDate"), "" & rs("VoucherAmount"), "" & rs("Vendor"))

        End If

        Call X.AddRemittanceDoc_Line("" & rs("Job"), "" & rs("PONumber"), "" & rs("InvoiceNumber"), "" & rs("InvoiceDate"), "" & rs("AmountPaid"), "" & rs("POIndex"), "" & rs("Status"))

        rs.MoveNext
    Wend

    If lastDoc <> "" Then
        Call X.CloseRemittanceDoc
        Call X.CloseRemittanceAdvice
        Call X.CloseMessage
        LogXML "RemittanceVouchers", X.xml
        If X.PostMessage Then
            s = "update LienVouchers set DateSentToBuildPro=getdate() where Voucher=" & DbQuote(Str, lastDoc)
            Call HFApp.SqlExec(s)
        End If
    End If

    Screen.MousePointer = vbNormal

Exit Sub
eh: Call errHandler(SRCFILE & "SendBuildProVouchers")
End Sub

Public Sub SendBuildProCommunities()
On Error GoTo eh
    Dim s As String
    Dim rs As Recordset
    Dim X As New HyphenSys.BuildProWrapper
    Dim Y As New HyphenSys.BuildProWrapper
    Dim divCode As String

    If Not BuildProEnabled Then Exit Sub
    Screen.MousePointer = vbHourglass

    divCode = HFApp.SqlExec("select divisioncode from divisions where divisionid=" & DbQuote(Num, HFApp.DivisionID))(0)


    'communities
    Call X.OpenMessage(HFApp.Options.ValueByName("BuildProCompanyCode"), divCode, HFApp.Options.ValueByName("BuildProURL"), HFApp.Options.ValueByName("BuildProUID"), HFApp.Options.ValueByName("BuildProPwd"))
    X.OpenCommunities
    s = ""
    s = s & "select * from BuildPro_Communities" & vbCrLf
    s = s & "where phase='0' and divisionid = " & DbQuote(Num, HFApp.DivisionID)
    Set rs = HFApp.SqlExec(s, dbHomefront)
    s = ""
    While Not rs.EOF
        Call X.Community("" & rs("Community") _
                       , Val("" & rs("Phase")) _
                       , "" & rs("Description") _
                       , "" & rs("County") _
                       , "primary", "" & rs("Address1"), "" & rs("Address2"), "" & rs("City"), "" & rs("State"), "" & rs("PostalCode"), "" & rs("phone"), "" & rs("Fax"), "" & rs("Email") _
                       , "billing", "" & rs("Address1"), "" & rs("Address2"), "" & rs("City"), "" & rs("State"), "" & rs("PostalCode"), "" & rs("phone"), "" & rs("Fax"), "" & rs("Email") _
                       , "shipping", "" & rs("Address1"), "" & rs("Address2"), "" & rs("City"), "" & rs("State"), "" & rs("PostalCode"), "" & rs("phone"), "" & rs("Fax"), "" & rs("Email"))
        s = s & "update divisioncommunities set DateSentToBuildPro=getdate() where divisionid=" & DbQuote(Num, HFApp.DivisionID) & " and community=" & DbQuote(Str, "" & rs("Community"))
        rs.MoveNext
    Wend
    If s <> "" Then
        X.CloseCommunities
        X.CloseMessage
        LogXML "communities", X.xml
        X.PostMessage
    End If


    'phases
    Call X.OpenMessage(HFApp.Options.ValueByName("BuildProCompanyCode"), divCode, HFApp.Options.ValueByName("BuildProURL"), HFApp.Options.ValueByName("BuildProUID"), HFApp.Options.ValueByName("BuildProPwd"))
    X.OpenCommunities
    s = ""
    s = s & "select * from BuildPro_Communities" & vbCrLf
    s = s & "where phase<>'0' and divisionid = " & DbQuote(Num, HFApp.DivisionID)
    Set rs = HFApp.SqlExec(s, dbHomefront)
    s = ""
    While Not rs.EOF
        Call X.Community("" & rs("Community") _
                       , Val("" & rs("Phase")) _
                       , "" & rs("Description") _
                       , "" & rs("County") _
                       , "primary", "" & rs("Address1"), "" & rs("Address2"), "" & rs("City"), "" & rs("State"), "" & rs("PostalCode"), "" & rs("phone"), "" & rs("Fax"), "" & rs("Email") _
                       , "billing", "" & rs("Address1"), "" & rs("Address2"), "" & rs("City"), "" & rs("State"), "" & rs("PostalCode"), "" & rs("phone"), "" & rs("Fax"), "" & rs("Email") _
                       , "shipping", "" & rs("Address1"), "" & rs("Address2"), "" & rs("City"), "" & rs("State"), "" & rs("PostalCode"), "" & rs("phone"), "" & rs("Fax"), "" & rs("Email"))
        s = s & "update divisioncommunities set DateSentToBuildPro=getdate() where divisionid=" & DbQuote(Num, HFApp.DivisionID) & " and community=" & DbQuote(Str, "" & rs("Community"))
        rs.MoveNext
    Wend
    If s <> "" Then
        X.CloseCommunities
        X.CloseMessage
        LogXML "phases", X.xml
        X.PostMessage
    End If

    Screen.MousePointer = vbNormal


Exit Sub
eh: Call errHandler(SRCFILE & "SendBuildCommunities")
End Sub

Public Function SendBuildProJobs(Job As String) As Boolean
On Error GoTo eh
    Dim s As String
    Dim i As Long
    Dim rs As Recordset
    Dim X As New HyphenSys.BuildProWrapper
    Dim divCode As String

    If Not BuildProEnabled Then Exit Function

    If Job = "" Then

        s = ""
        s = s & "New Jobs" & Chr(1)
        s = s & "select HFJob Job" & IIf(HFApp.Options(MultiFamily), ",job Schedule", "") & ", Description,CommunityNumber Community,PhaseNumber Phase, LotNumber, BuyerName, isnull(AddressLine1,'') Address, City" & vbCrLf
        s = s & "from BuildPro_Jobs" & vbCrLf
        s = s & "where divisionid=" & DbQuote(Num, HFApp.DivisionID)
        s = s & "and DateSentToBuildPro is null" & vbCrLf
        s = s & Chr(0)
        s = s & "All Jobs" & Chr(1)
        s = s & "select HFJob Job" & IIf(HFApp.Options(MultiFamily), ",job Schedule", "") & ",Description,CommunityNumber Community,PhaseNumber Phase, LotNumber, BuyerName, isnull(AddressLine1,'') Address, City" & vbCrLf
        s = s & "from BuildPro_Jobs" & vbCrLf
        s = s & "where divisionid=" & DbQuote(Num, HFApp.DivisionID)

        If Not FPickList.Choose(HFApp.Databases(dbHomefront), "Job", s, , , , , , True) Then Exit Function
        For i = 1 To FPickList.SelectedItems
            Job = Job & "," & DbQuote(Str, FPickList.SelectedItem("Job", i))
        Next
        If Job = "" Then
            Exit Function
        Else
            Job = Mid(Job, 2)
        End If
    Else
        Job = DbQuote(Str, Job)
    End If

    Screen.MousePointer = vbHourglass

  '  Call SendBuildProCommunities

    divCode = HFApp.SqlExec("select divisioncode from divisions where divisionid=" & DbQuote(Num, HFApp.DivisionID))(0)
    Call X.OpenMessage(HFApp.Options.ValueByName("BuildProCompanyCode"), divCode, HFApp.Options.ValueByName("BuildProURL"), HFApp.Options.ValueByName("BuildProUID"), HFApp.Options.ValueByName("BuildProPwd"))
    X.OpenJobs

    s = "select * from BuildPro_Jobs where hfjob in(" & Job & ")"
    Set rs = HFApp.SqlExec(s, dbHomefront)
    s = ""
    While Not rs.EOF

        Call X.Job("" & rs("Job"), "" & rs("Description"), "" & rs("CommunityNumber"), "" & rs("County"), Val("" & rs("PhaseNumber")), "" & rs("LotNumber"), "" & rs("Building"), _
                   "" & rs("Unit"), "" & rs("Plan"), "" & rs("Elevation"), "" & rs("Swing"), "" & rs("ColorPackage"), "" & rs("ScheduleTemplate"), "" & rs("StartDate"), _
                   "" & rs("PromisedDeliveryDate"), "" & rs("BuyerCloseDate"), "" & rs("PermitReceivedDate"), "" & rs("PermitNumber"), "" & rs("LegalLotNumber"), "" & rs("BlockNumber"), _
                   "" & rs("TractNumber"), "" & rs("ConstructionType"), "" & rs("AddressLine1"), "" & rs("AddressLine2"), "" & rs("City"), "" & rs("State"), _
                   "" & rs("PostalCode"), "" & rs("LotComment"), "" & rs("ConstructionSequence"), "" & rs("Budget"), "" & rs("SalesPrice"), "" & rs("Customer_No"), _
                   "" & rs("BuyerFName"), "" & rs("BuyerLName"), "" & rs("BuyerAddressLine1"), "" & rs("BuyerAddressLine2"), "" & rs("BuyerCity"), "" & rs("BuyerState"), _
                   "" & rs("BuyerPostalCode"), "" & rs("BuyerHomePhone"), "" & rs("BuyerWorkPhone"), "" & rs("BuyerMobilePhone"), "" & rs("BuyerFax"), "" & rs("BuyerEmail"), "" & rs("BuyerType"), _
                   "" & rs("CoBuyerFName"), "" & rs("CoBuyerLName"), "" & rs("CoBuyerAddressLine1"), "" & rs("CoBuyerAddressLine2"), "" & rs("CoBuyerCity"), "" & rs("CoBuyerState"), _
                   "" & rs("CoBuyerPostalCode"), "" & rs("CoBuyerHomePhone"), "" & rs("CoBuyerWorkPhone"), "" & rs("CoBuyerMobilePhone"), "" & rs("CoBuyerFax"), "" & rs("CoBuyerEmail"), "" & rs("CoBuyerType"), _
                   "" & rs("TarionBuilderNumber"), "" & rs("TarionEnrollmentNumber"))

        rs.MoveNext
    Wend

    X.CloseJobs
    X.CloseMessage
    LogXML "jobs", X.xml
    If X.PostMessage() Then
        s = "update tblJobs set DateSentToBuildPro=getdate() where divisionid=" & DbQuote(Num, HFApp.DivisionID) & " and Job_no in(" & Job & ")"
        Call HFApp.SqlExec(s, dbHomefront)
        SendBuildProJobs = True
    End If
    Screen.MousePointer = vbNormal

Exit Function
eh: Call errHandler(SRCFILE & "SendBuildProJobs")
End Function

Public Function CancelBuildProPO(Job As String, PONumber As String) As Boolean
    Dim s As String
    Dim rs As Recordset

    If Not BuildProEnabled Then Exit Function

    'if not posted then do nothing.. NO ALWAYS SEND IT BUT IGNORE ANY ERRORS
'    s = "select * from pomaster where datesenttobuildpro is not null and divisionid=" & HFApp.DivisionID & " and PONumber=" & DbQuote(Str, PONumber)
'    Set rs = HFApp.SqlExec(s, dbHomefront)
'    If rs.EOF Then
'        CancelBuildProPO = True
'        Exit Function
'    End If

    'reset date so it can be sent and send it
    s = "update pomaster set datesenttobuildpro=null where divisionid=" & HFApp.DivisionID & " and PONumber=" & DbQuote(Str, PONumber)
    Call HFApp.SqlExec(s, dbHomefront)

    If SendBuildProPOs(DbQuote(Str, Job), DbQuote(Str, PONumber), True) Then
        CancelBuildProPO = True
    Else
        'reset date so you can retry the cancel.
        s = "update pomaster set datesenttobuildpro=getdate() where divisionid=" & HFApp.DivisionID & " and PONumber=" & DbQuote(Str, PONumber)
        Call HFApp.SqlExec(s, dbHomefront)
    End If


End Function

Public Function SendBuildProPOs(Job As String, Pos As String, Optional HideErrors As Boolean = False) As Boolean

On Error GoTo eh
'both parameters are already quoted..
    Dim i As Long
    Dim s As String
    Dim rs As Recordset
    Dim X As New HyphenSys.BuildProWrapper
    Dim divCode As String

    If Not BuildProEnabled Then Exit Function


    If Job = "" And Pos = "" Then
        s = ""
        s = s & "Un-sent POs" & Chr(1)
        s = s & "select distinct HFJob Job" & IIf(HFApp.Options(MultiFamily), ",job Schedule", "") & ",PONumber,POIndex" & vbCrLf
        s = s & "from BuildPro_POs" & vbCrLf
        s = s & "where divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        s = s & "and DateSentToBuildPro is null" & vbCrLf
        s = s & Chr(0)
'        s = s & "Changed POs" & Chr(1)
'        s = s & "select distinct HFJob Job" & IIf(HFApp.Options(MultiFamily), ",job Schedule", "") & ",PONumber,POIndex,ModifiedDate Changed,DateSentToBuildPro SentToBuildPro" & vbCrLf
'        s = s & "from BuildPro_POs" & vbCrLf
'        s = s & "where divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
'        s = s & "and DateSentToBuildPro is not null and DateSentToBuildPro < ModifiedDate" & vbCrLf
'        s = s & Chr(0)
        s = s & "All POs" & Chr(1)
        s = s & "select distinct HFJob Job" & IIf(HFApp.Options(MultiFamily), ",job Schedule", "") & ",PONumber,POIndex,ModifiedDate Changed,DateSentToBuildPro SentToBuildPro" & vbCrLf
        s = s & "from BuildPro_POs" & vbCrLf
        s = s & "where divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        If Not FPickList.Choose(HFApp.Databases(dbHomefront), "POs", s, , , , , , True) Then Exit Function
        For i = 1 To FPickList.SelectedItems
            Pos = Pos & "," & DbQuote(Str, FPickList.SelectedItem("PONumber", i))
        Next
        If Pos = "" Then
            Exit Function
        Else
            Pos = Mid(Pos, 2)
        End If
    End If


    Screen.MousePointer = vbHourglass

    divCode = HFApp.SqlExec("select divisioncode from divisions where divisionid=" & DbQuote(Num, HFApp.DivisionID))(0)
    Call X.OpenMessage(HFApp.Options.ValueByName("BuildProCompanyCode"), divCode, HFApp.Options.ValueByName("BuildProURL"), HFApp.Options.ValueByName("BuildProUID"), HFApp.Options.ValueByName("BuildProPwd"))
    X.OpenCommitments

'    s = ""
'    s = s & "select * from buildpro_POs" & vbCrLf
'    s = s & "where divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
'    If Pos <> "" Then
'        s = s & "and PONumber in(" & Pos & ")" & vbCrLf
'    Else
'        s = s & "and Job=" & DbQuote(Str, Job) & vbCrLf
'        s = s & "and isnull(EPOID,'')=''" & vbCrLf           'dont send epos
'        s = s & "and DateSentToBuildPro is null" & vbCrLf    'only send new things
'    End If
'    s = s & "order by ponumber,linenumber"

' -- VIEW IS 2-3 TIMES SLOWER THAN PROC? use proc where possible ---
    If Pos = "" Then
        s = "exec dbo.BuildPro_GetPOs " & DbQuote(Num, HFApp.DivisionID) & ", " & DbQuote(Str, Job)
    Else
        s = ""
        s = s & "select * from buildpro_POs" & vbCrLf
        s = s & "where divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        s = s & "and PONumber in(" & Pos & ")" & vbCrLf
        s = s & "order by ponumber,linenumber"
    End If


    Set rs = HFApp.SqlExec(s, dbHomefront)

    Pos = ""
    While Not rs.EOF

        Call X.CommitmentLine("" & rs("PONumber"), "" & rs("PODate"), "" & rs("PODocType"), "" & rs("PODocSuffix"), "" & rs("EPOid"), "" & rs("EPODocType"), "" & rs("EPODocSuffix"), _
            "" & rs("Vendor"), "" & rs("POIndex"), "" & rs("POCancelled"), "" & rs("Job"), "" & rs("Community"), Val("" & rs("CommunityPhase")), "" & rs("SKU"), _
            Val("" & rs("LineNumber")), "" & rs("ItemDescription"), Val("" & rs("Qty")), "" & rs("UOM"), Val("" & rs("UnitPrice")), Val("" & rs("Pretax")), _
            Val("" & rs("JCTax")), _
            Val("" & rs("NJCTax")), _
            "" & rs("OptionID"), "" & rs("OptionDesc"), "" & rs("OptionQty"), Val("" & rs("OptionCost")), "" & rs("OptionNotes"), _
            "" & rs("AttrColor"), "" & rs("AttrFinish"), "" & rs("AttrStyle"), "" & rs("AttrLocation"), "" & rs("AttrOther"), "" & rs("itemComments"))

        Pos = Pos & "," & DbQuote(Str, "" & rs("PONumber"))
        rs.MoveNext
    Wend
    Pos = Mid(Pos, 2)

    If Pos <> "" Then
        X.CloseCommitments
        X.CloseMessage
        LogXML "commitments", X.xml

        If HideErrors Then
            Call X.PostMessageQuietly
        Else

            If X.PostMessage Then
                SendBuildProPOs = True

                s = ""
                s = s & "update pomaster set DateSentToBuildPro=getdate(),EPOid=''" & vbCrLf
                s = s & "where divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
                s = s & "and PONumber in(" & Pos & ")"
                Call HFApp.SqlExec(s, dbHomefront)

            End If
        End If
    End If
    Screen.MousePointer = vbNormal



Exit Function
eh: Call errHandler(SRCFILE & "SendBuildProPOs")
End Function





