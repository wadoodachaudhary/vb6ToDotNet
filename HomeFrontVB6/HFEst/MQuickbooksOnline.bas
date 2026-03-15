Attribute VB_Name = "MQuickBooksOnline"
Option Explicit
Option Compare Text
Private Const SRCFILE = "MQuickBooksOnline::"

Public Function SubmitQBOXml(ActionCode As String, XmlMessage As String) As String
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
    'uat = "https://www.homefrontcrm.com/hfbrokeruat"
    
    'log request
    i = FreeFile()
    Open PathAppend(GetFolderPath(CSIDL_LOCAL_APPDATA), "HomeFront\QBOnlinerequest.xml") For Output As #i
    Print #i, XmlMessage
    Close #i

    'process request
   
    errtext = vbCrLf & vbCrLf & vbCrLf & XmlMessage
    Response = c.DoTransaction("QBOnline", HFApp.ClientID, HFApp.DivisionID, HFApp.LoginID, ActionCode, XmlMessage, True, False)

    
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

