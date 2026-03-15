Attribute VB_Name = "XmlQB"
Option Explicit

Public Function SubmitQBXml(XmlMessage As String) As String
'this must reraise errors so payables will fail out of postings

    Dim ticket  As String
    Dim Response As String
    Dim Status   As String
    Dim Reason   As String
    Dim RP As New QBXMLRP2Lib.RequestProcessor2
    Dim errtext As String
    Dim i As Integer
    
    'log request
    i = FreeFile()
    Open PathAppend(GetFolderPath(CSIDL_LOCAL_APPDATA), "HomeFront\QBrequest.xml") For Output As #i
    Print #i, XmlMessage
    Close #i

    'process request
    errtext = "OpenConnection2"
    Call RP.OpenConnection2("HomeFront", "HomeFront", localQBD)
    errtext = "BeginSession " & HFApp.Options(QuickBooksDataFile)
    ticket = RP.BeginSession(HFApp.Options(QuickBooksDataFile), qbFileOpenMultiUser)
    errtext = "ProcessRequest " & XmlMessage
    Response = RP.ProcessRequest(ticket, XmlMessage)
    
    'clean response -- comeon quickbooks, get your shit together
    Response = Replace(Response, "`", "'")
    Response = Replace(Response, "„", ",,")
    Response = Replace(Response, "…", "...")
    Response = Replace(Response, "ˆ", "^")
    Response = Replace(Response, "‰", "")
    Response = Replace(Response, "‹", "<")
    Response = Replace(Response, "‘", "'")
    Response = Replace(Response, "’", "'")
    Response = Replace(Response, "“", """")
    Response = Replace(Response, "”", """")
    Response = Replace(Response, "•", "*")
    Response = Replace(Response, "–", "-")
    Response = Replace(Response, "—", "-")
    Response = Replace(Response, "™", "(TM)")
    Response = Replace(Response, "©", "(C)")
    Response = Replace(Response, "«", "<<")
    Response = Replace(Response, "®", "(R)")
    Response = Replace(Response, "²", "2")
    Response = Replace(Response, "³", "3")
    Response = Replace(Response, "·", "")
    Response = Replace(Response, "¸", "")
    Response = Replace(Response, "¹", "1")
    Response = Replace(Response, "º", "0")
    Response = Replace(Response, "»", ">>")
    Response = Replace(Response, "¼", "1/4")
    Response = Replace(Response, "½", "1/2")
    Response = Replace(Response, "¾", "3/4")
    
    'another weird double quote that vb cant draw
    Response = Replace(Response, Chr(152), Chr(34))

    On Error Resume Next
    
    errtext = "End Session " & ticket
    Call RP.EndSession(ticket)
    errtext = "Close Connection"
    Call RP.CloseConnection
    On Error GoTo 0
    errtext = "Create Response xml file"
    'log response
    i = FreeFile()
    Open PathAppend(GetFolderPath(CSIDL_LOCAL_APPDATA), "HomeFront\QBresponse.xml") For Output As #i
    Print #i, Response
    errtext = "Close Response xml file"
    Close #i
    
    'parse and return response
    errtext = "Get Status from response"
    Status = Parse(Parse(Response, 2, "statusCode="), 2, """")
    errtext = "Get Reason from response"
    Reason = Parse(Parse(Response, 2, "statusMessage="), 2, """")
    Reason = Replace(Reason, "&quot;", """")
    
    Select Case Status
        Case "0", "1" 'no error   or   query request did not find a matching object in QuickBooks
            SubmitQBXml = Response
        Case Else
            Err.Raise Val(Status), "SubmitQBXml", Reason & vbCrLf & vbCrLf & XmlMessage
            SubmitQBXml = Reason
    End Select
End Function

Public Function StartQBXml() As String
    Dim s As String
    s = ""
    s = s & "<?xml version=""1.0"" ?>" & vbCrLf
    s = s & "<?qbxml version=""6.0""?>" & vbCrLf
    s = s & "<QBXML>" & vbCrLf
    s = s & "<QBXMLMsgsRq onError=""stopOnError"">" & vbCrLf
    StartQBXml = s
End Function

Public Function EndQBXml() As String
    Dim s As String
    s = ""
    s = s & "</QBXMLMsgsRq>" & vbCrLf
    s = s & "</QBXML>" & vbCrLf
    EndQBXml = s
End Function



Public Function AddQBXml(FieldType As XMLFieldTypes, SIZE As Double, Name As String, Value) As String
                     
    'field types are copied over from MasterBuilders implementation
                     
    Dim s As String
    s = s & "<" & Name & ">"
    If UCase(Right(Trim(Name), 3)) = "REF" Then s = s & "<ListID>"
    If SIZE = 0 Then SIZE = 99999
    Select Case FieldType
        Case d:     s = s & Format(Value, "YYYY-MM-DD")
        Case n:     s = s & Format(Value, "0" & IIf((SIZE - Int(SIZE)) * 10 > 0, "." & String((SIZE - Int(SIZE)) * 10, "0"), ""))
        Case Else:  s = s & CleanXMLText(left(Value, Int(SIZE)))
    End Select
    If UCase(Right(Trim(Name), 3)) = "REF" Then s = s & "</ListID>"
    s = s & "</" & Name & ">" & vbCrLf
    
    AddQBXml = s
        
End Function





