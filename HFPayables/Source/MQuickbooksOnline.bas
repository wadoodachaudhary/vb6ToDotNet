Attribute VB_Name = "MQuickbooksOnline"
Option Explicit
Const SRCFILE = "MQuickBooksOnline::"


Public Function QBOGetPODetails(PONumber As String) As String
    Dim q As String
    Dim s As String
    
    Dim c As New BrokerWrapper.HomeFrontClient
    
    
    'url defaults to production
    s = HFApp.Options.ValueByName("BrokerURL")
    If s = "" Then s = "https://www.homefrontcrm.com/hfbroker"
    c.BasePath = s
    'prod= "https://www.homefrontcrm.com/hfbroker"
    'uat = "http://uat.homefrontcrm.com/broker"
    
    q = "select * from purchaseorder where docnumber=" & DbQuote(Str, PONumber)
    s = c.DoTransaction("QBOnline", HFApp.ClientID, HFApp.DivisionID, HFApp.LoginID, "Query", q, True, False)
    
    QBOGetPODetails = s

End Function



Public Function SubmitBrokerXml(SystemId As String, ActionCode As String, XmlMessage As String) As String
    On Error GoTo eh
    Dim Response As String
    Dim errtext As String
    Dim i As Integer
    Dim c As New BrokerWrapper.HomeFrontClient
    Dim s As String

    'url defaults to production
    s = HFApp.Options.ValueByName("BrokerURL")
    If s = "" Then s = "https://www.homefrontcrm.com/hfbroker"
    c.BasePath = s
    'prod= "https://www.homefrontcrm.com/hfbroker"
    'uat = "http://uat.homefrontcrm.com/broker"


    
    'log request
    i = FreeFile()
    Open PathAppend(GetFolderPath(CSIDL_LOCAL_APPDATA), "HomeFront\" & SystemId & "Request.xml") For Output As #i
    Print #i, XmlMessage
    Close #i

    'process request
   
    errtext = vbCrLf & vbCrLf & vbCrLf & XmlMessage
    Response = c.DoTransaction(SystemId, HFApp.ClientID, HFApp.DivisionID, HFApp.LoginID, ActionCode, XmlMessage, True, False)

    
    'clean response -- come on quickbooks, get your shit together
    Response = Replace(Response, Chr(188), "1/4")
    Response = Replace(Response, Chr(189), "1/2")
    Response = Replace(Response, Chr(190), "3/4")
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
    Open PathAppend(GetFolderPath(CSIDL_LOCAL_APPDATA), "HomeFront\" & SystemId & "Response.xml") For Output As #i
    Print #i, Response
    errtext = "Close Response xml file"
    Close #i

    SubmitBrokerXml = Response

Exit Function
eh: Select Case True
        Case err.Number = 0
        Case err.Description Like "*<status>401</status>*"
        Case Else
            MsgBox err.Description & " " & errtext, vbCritical, "SubmitBrokerXml Failed"
            SubmitQBOXml = ""
    End Select
End Function

