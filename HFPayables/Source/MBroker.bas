Attribute VB_Name = "MBroker"
Option Explicit
Const SRCFILE = "MBroker::"


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
    s = "https://www.homefrontcrm.com/hfbroker"
    If InIde() Then s = "https://www.homefrontcrm.com/hfbrokeruat"
    c.BasePath = s


    
    'process request
    errtext = vbCrLf & vbCrLf & vbCrLf & XmlMessage
    Call Log(SystemId & ".Request.xml", XmlMessage, False)
    
    'Call c.RevokeToken(SystemId, HFApp.ClientID, HFApp.DivisionID, HFApp.LoginID)
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
    Response = Replace(Response, Chr(150), "-")
    
    'log response
    Call Log(SystemId & ".Response.xml", Response, False)


    'parse the response for errors... this is terrible.
    Select Case True
        
        Case Response Like "*Xero API error*"
            err.Raise vbObjectError + 78, SystemId, Response
        
        Case Response Like "Successfully created/updated*"
            Response = ""
    
    End Select
    
    SubmitBrokerXml = Response

Exit Function
eh:

Select Case True
        Case err.Number = 0
        Case err.Description Like "*<status>401</status>*"
        Case err.Description Like "<!DOCTYPE html>*"
            If MsgBox("An error has occured" & vbCrLf & vbCrLf & "Would you like to view the HTML error response?", vbExclamation + vbYesNo, "Error") = vbYes Then
                Call ShellFile(FMain.hWnd, Log(SystemId & ".Error.html", err.Description, False), , False)
            End If
            SubmitBrokerXml = err.Description
        Case Else
            Call Log(SystemId & ".Error.txt", err.Description, False)
            MsgBox err.Description & " " & errtext, vbCritical, "SubmitBrokerXml Failed"
            SubmitBrokerXml = err.Description
    End Select
End Function

Private Function Log(FileName As String, Data As String, append As Boolean) As String
Dim i As Integer
Dim s As String
    
    s = PathAppend(GetFolderPath(CSIDL_LOCAL_APPDATA), "HomeFront\" & FileName)
    i = FreeFile()
    If append Then
        Open s For Append As #i
    Else
        Open s For Output As #i
    End If
    Print #i, Data
    Close #i
    
    Log = s
End Function
