Attribute VB_Name = "MQuickBooksOnline"
Option Explicit
Const SRCFILE = "MQuickBooksOnline::"



Public Function SubmitQBOXml(ActionCode As String, XmlMessage As String) As String
On Error GoTo eh
    Dim s As String
    Dim Response As String
    Dim errtext As String
    Dim i As Integer
    Dim c As New BrokerWrapper.HomeFrontClient


    'url defaults to production
    'prod= "https://www.homefrontcrm.com/hfbroker"
    'uat = "http://uat.homefrontcrm.com/broker"
    
    s = HFApp.Options.ValueByName("BrokerURL")
    If s = "" Then s = "https://www.homefrontcrm.com/hfbroker"
    If InIde() Then s = "https://www.homefrontcrm.com/hfbrokeruat"
    c.BasePath = s
    

    
    'log request
    i = FreeFile()
    Open PathAppend(GetFolderPath(CSIDL_LOCAL_APPDATA), "HomeFront\QBOnlinerequest.xml") For Output As #i
    Print #i, XmlMessage
    Close #i

    'process request
   
    errtext = vbCrLf & vbCrLf & vbCrLf & XmlMessage
    Response = c.DoTransaction("QBOnline", HFApp.ClientID, HFApp.DivisionID, HFApp.LoginID, ActionCode, XmlMessage, True, True)

    
    'clean response -- come on quickbooks, get your shit together
    Response = Replace(Response, Chr(188), "1/4")
    Response = Replace(Response, Chr(189), "1/2")
    Response = Replace(Response, Chr(190), "3/4")
    Response = Replace(Response, Chr(150), "-")
    Response = Replace(Response, Chr(145), Chr(39)) 'left curly single quote
    Response = Replace(Response, Chr(146), Chr(39)) 'right curly single quote
    Response = Replace(Response, Chr(147), Chr(34)) 'left curly double quote
    Response = Replace(Response, Chr(148), Chr(34)) 'right curly double quote
    Response = Replace(Response, Chr(152), Chr(34)) 'another weird double quote
    
    On Error Resume Next

    On Error GoTo eh
    errtext = "Create Response xml file"
    
    'log response
    i = FreeFile()
    Open PathAppend(GetFolderPath(CSIDL_LOCAL_APPDATA), "HomeFront\QBOnlineresponse.xml") For Output As #i
    Print #i, Response
    errtext = "Close Response xml file"
    Close #i

    SubmitQBOXml = Response

Exit Function
eh: Select Case True
        Case Err.Number = 0
        Case Err.Description Like "*<status>401</status>*"
        Case Else
            MsgBox Err.Description & " " & errtext, vbCritical, "SubmitQBOXml Failed"
            SubmitQBOXml = ""
    End Select
End Function



Public Sub WriteCustomerToQBO(JobNumber As String, CustomerNumber As String)
On Error GoTo eh
    Dim i As Long
    Dim s As String
    Dim rs As Recordset
    Dim listid     As String
    Dim parentid   As String

    s = ""
    s = s & "select * from qbo_customers where divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
    If CustomerNumber = "" Then
        s = s & "and job_no=" & DbQuote(Str, JobNumber) & vbCrLf
    Else
        s = s & "and Customer_no=" & DbQuote(Str, CustomerNumber) & vbCrLf
    End If
    Set rs = HFApp.SqlExec(s)
    While Not rs.EOF
    
        s = "<?xml version=""" & "1.0""" & " encoding=""" & "utf-8""" & "?>" & vbCrLf
        s = s & "<Customer xmlns:xsi=""" & "http://www.w3.org/2001/XMLSchema-instance""" & " xmlns:xsd=""" & "http://www.w3.org/2001/XMLSchema""" & ">" & vbCrLf
        s = s & AddQBXml(c, 9000, "Title", "" & rs("Title"))
        If "" & rs("ExternalJobID") <> "" Then s = s & AddQBXml(c, 9000, "Id", "" & rs("ExternalJobID"))
        s = s & AddQBXml(c, 9000, "FirstName", "" & rs("FirstName"))
        s = s & AddQBXml(c, 9000, "LastName", "" & rs("LastName"))
        s = s & AddQBXml(c, 9000, "Description", "" & rs("Description"))
        s = s & AddQBXml(c, 9000, "Phone", "" & rs("Phone"))
        s = s & AddQBXml(c, 9000, "Email", "" & rs("Email"))
        s = s & AddQBXml(c, 9000, "Address1", "" & rs("Address1"))
        s = s & AddQBXml(c, 9000, "Address2", "" & rs("Address2"))
        s = s & AddQBXml(c, 9000, "City", "" & rs("City"))
        s = s & AddQBXml(c, 9000, "Country", "" & rs("Country"))
        s = s & AddQBXml(c, 9000, "State", "" & rs("State"))
        s = s & AddQBXml(c, 9000, "PostalCode", "" & rs("PostalCode"))
        s = s & AddQBXml(c, 9000, "JobHierarchy", "" & rs("JobHierarchy"))
        s = s & AddQBXml(c, 9000, "JobDescription", "" & rs("JobDescription"))
        s = s & "</Customer>" & vbCrLf
        
        s = SubmitQBOXml("PostCustomer", s)
        listid = Parse(s, 2, "<Id>")
        listid = Parse(listid, 1, "</Id>")
        parentid = Parse(s, 2, "<Parent>")
        parentid = Parse(parentid, 1, "</Parent>")
        
        
        If listid <> "" Then
            
            'set external id on customer
            s = ""
            s = s & "update tblcustomers" & vbCrLf
            s = s & "   set ar_customer_deposit=" & DbQuote(Str, listid) & vbCrLf
            s = s & " where divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
            s = s & "   and customer_no=" & DbQuote(Str, "" & rs("customer_no")) & vbCrLf
            Call HFApp.SqlExec(s, dbHomefront)
            
            'set external id on job
            s = ""
            s = s & "update tbljobs" & vbCrLf
            s = s & "   set ExternalJobID=" & DbQuote(Str, listid) & vbCrLf
            s = s & " where divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
            s = s & "   and job_no=" & DbQuote(Str, "" & rs("Job_No")) & vbCrLf
            Call HFApp.SqlExec(s, dbHomefront)
            
            
            'if style = simple then write listid to customer and to job
            'if style = heirarchy then write parentid to customer and listid to job
            If HFApp.Options.ValueByName("QuickBooksJobStyle") <> "Simple" Then listid = parentid
            
            'add customer to HF
            s = ""
            s = s & "insert into arcustomers(arcustomer,description,contact1,phone1,email1,phone2,billaddr1,billaddr2,billcity,billprovince,billpostalcode,shipaddr1,shipcity,shipprovince,shippostalcode)" & vbCrLf
            s = s & "values(" & DbQuote(Str, listid) & vbCrLf
            s = s & "      ," & DbQuote(Str, "" & rs("customer_no") & " - " & rs("description")) & vbCrLf
            s = s & "      ," & DbQuote(Str, "" & rs("FirstName") & " " & rs("FirstName")) & vbCrLf
            s = s & "      ," & DbQuote(Str, "" & rs("phone")) & vbCrLf
            s = s & "      ," & DbQuote(Str, "" & rs("email")) & vbCrLf
            s = s & "      ," & DbQuote(Str, "" & rs("address1")) & vbCrLf
            s = s & "      ," & DbQuote(Str, "" & rs("address2")) & vbCrLf
            s = s & "      ," & DbQuote(Str, "" & rs("city")) & vbCrLf
            s = s & "      ," & DbQuote(Str, "" & rs("state")) & vbCrLf
            s = s & "      ," & DbQuote(Str, "" & rs("postalcode")) & vbCrLf
            s = s & "      ," & DbQuote(Str, "" & rs("address1")) & vbCrLf
            s = s & "      ," & DbQuote(Str, "" & rs("city")) & vbCrLf
            s = s & "      ," & DbQuote(Str, "" & rs("state")) & vbCrLf
            s = s & "      ," & DbQuote(Str, "" & rs("postalcode")) & ")"
            On Error Resume Next
            Call HFApp.SqlExec(s, dbHomefront)
            On Error GoTo eh
            
            
        End If
        
        rs.MoveNext
    Wend





Exit Sub
eh:
If InStr(1, Err.Description, "already in use") > 0 Or InStr(1, Err.Description, "duplicate") > 0 Then
        Resume Next
    Else
        Call errHandler("WriteCustomerToQBO", s)
    End If
End Sub


