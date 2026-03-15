Attribute VB_Name = "XmlMB"
Option Explicit


Public Function StartMBXml(Company As String, UserName As String) As String
    Dim s As String
    
        
    If HFApp.Options.ValueByName("Sage100APILevel") <> "v18" Then
        s = "<api:MBXML xmlns:api=""http://sage100contractor.com/api"">" & vbCrLf
    Else
        s = "<api:MBXML xmlns:api=""http://sagemasterbuilder.com/api"">" & vbCrLf
    End If
    
    
    s = s & "<MBXMLSessionRq>" & vbCrLf
    s = s & AddMBXml(c, 9999, "Company", Company)
    s = s & AddMBXml(c, 9999, "User", UserName)
    s = s & "</MBXMLSessionRq>" & vbCrLf
    s = s & "<MBXMLMsgsRq messageSetID=""1"" onError=""continueOnError"">" & vbCrLf
    StartMBXml = s
End Function

Public Function EndMBXml() As String
    Dim s As String
    s = s & "</MBXMLMsgsRq>" & vbCrLf
    s = s & "</api:MBXML>" & vbCrLf
    EndMBXml = s
End Function

Public Function AddMBXml(FieldType As XMLFieldTypes, SIZE As Double, Name As String, Value) As String
                     
    Dim s As String
    s = s & "<" & Name & ">"
    If UCase(Right(Trim(Name), 3)) = "REF" Then s = s & "<ObjectID>"
    If SIZE = 0 Then SIZE = 99999
    Select Case FieldType
        Case d:     s = s & Format(Value, "YYYY-MM-DD")
        Case n:     s = s & Format(Value, "0" & IIf((SIZE - Int(SIZE)) * 10 > 0, "." & String((SIZE - Int(SIZE)) * 10, "0"), ""))
        Case Else:  s = s & CleanXMLText(left(Value, Int(SIZE)))
    End Select
    If UCase(Right(Trim(Name), 3)) = "REF" Then s = s & "</ObjectID>"
    s = s & "</" & Name & ">" & vbCrLf
    
    AddMBXml = s
        
End Function

Public Function SubmitMBXml(MBXmlMessage As String, Pswd As String) As Boolean

    Dim i As Integer
    Dim Response As String
    
    Dim MB As Object
    
    Dim ErrNum As String
    Dim ErrMsg As String
    
    'log request
    i = FreeFile()
    Open PathAppend(GetFolderPath(CSIDL_LOCAL_APPDATA), "HomeFront\MBrequest.xml") For Output As #i
    Print #i, MBXmlMessage
    Close #i
        
    'process request
    Select Case HFApp.Options.ValueByName("Sage100APILevel")
    Case "v20"
        Set MB = CreateObject("Sage100v20Wrapper.MBXML")
        Response = MB.MBSubmitXML(MBXmlMessage, HFApp.Options(MasterBuilderDataFolder), HFApp.Options(MasterBuilderPWD))
        
    Case "v19"
        Set MB = CreateObject("Sage100Wrapper.MBXML")
        Response = MB.MBSubmitXML(MBXmlMessage, LCase(left(HFApp.Options(MasterBuilderDataFolder), 1)), HFApp.Options(MasterBuilderPWD))
    
    Case Else
        Set MB = CreateObject("MBAPI.IMBXML")
        Response = MB.submitXML(MBXmlMessage, Pswd)
    
    End Select
    
    'log response
    i = FreeFile()
    Open PathAppend(GetFolderPath(CSIDL_LOCAL_APPDATA), "HomeFront\MBresponse.xml") For Output As #i
    Print #i, Response
    Close #i
    
    'process response
    ErrNum = Parse(Parse(Response, 2, "statusCode="""), 1, """")
    ErrMsg = Parse(Parse(Response, 2, "statusMessage="""), 1, """")
    Select Case True
        Case ErrNum = "0":      SubmitMBXml = True
        
        'this also uses errnum 14201 so make sure it comes first.
        Case ErrMsg Like "Failed Coded Constraint*":  Err.Raise Val(ErrNum), "SubmitMBXml", ErrMsg
        Case ErrNum = "14201":  'IGNORE -- Failed Coded Constraint Purchase Order number already exists
        
        Case ErrNum = "12000":  Err.Raise Val(ErrNum), "SubmitMBXml", "Cannot insert duplicate key"
        Case ErrNum = "12802":  Err.Raise Val(ErrNum), "SubmitMBXml", ErrMsg  'Invalid Object reference
        Case ErrNum = "12806":  Err.Raise Val(ErrNum), "SubmitMBXml", "Your Sage 100 security settings do not permit this action."
        Case ErrNum = "14050":  Err.Raise Val(ErrNum), "SubmitMBXml", ErrMsg  'invalid data, ErrMsg contains details
        Case ErrNum = "":       Err.Raise vbObjectError + 99, "SubmitMBXml", Response
        Case Else:              Err.Raise Val(ErrNum), "SubmitMBXml", Response
    End Select
            
End Function

Private Function MBParserError(xml As String) As String
    
    Dim SC As New XMLSchemaCache60
    Dim idoc As New MSXML2.DOMDocument60
    Dim sdoc As New MSXML2.DOMDocument60
    Dim res As Boolean
    Dim strTestXML As String
    sdoc.async = False
    Dim tns As String
    Dim schemafile As String
    
    
    schemafile = Trim$(left(HFApp.Options(MasterBuilderDataFolder), 2)) & "\MB7\Programs\mbxml.xsd"
    
        
    If (xml = "") Then
        MBParserError = "xml is missing"
        Exit Function
    End If
    If Not FileExists(schemafile) Then
        MBParserError = "schema file is missing" & vbCrLf & schemafile
        Exit Function
    End If
    
    
     
    'load schema
    res = sdoc.Load(schemafile)
    If sdoc Is Nothing Or res = False Then
        MBParserError = "schema is not wellformed"
        Exit Function
    Else
        If sdoc.selectSingleNode("//*/@targetNamespace") Is Nothing Then
            MBParserError = "missing namespace"
            Exit Function
        End If
        tns = sdoc.selectSingleNode("//*/@targetNamespace").Text
        SC.Add tns, sdoc
    End If
    
    
    
    'load xml
    res = idoc.loadXML(xml)
    If idoc Is Nothing Or res = False Then
        MBParserError = "xml is not wellformed"
        Exit Function
    End If
    If idoc.namespaces.Length <> 1 Then
        MBParserError = "missing namespace"
        Exit Function
    End If
    If idoc.namespaces.namespaceURI(0) <> SC.namespaceURI(0) Then
       MBParserError = "incorrect namespace"
       Exit Function
    End If
    
    
    'parse xml/schema
    Set idoc.schemas = SC
    Dim result As MSXML2.IXMLDOMParseError
    Set result = idoc.Validate
    If result.errorCode <> 0 Then
        MBParserError = result.Reason
        Exit Function
    End If
        
End Function

Public Function CleanXMLText(s) As String
'    Const SC = "&!$%'*+:;<>?^`{}|~"
    Const SC = "&<>'"""
    Dim i As Long
    For i = 1 To Len(SC)
        s = Replace(s, Mid(SC, i, 1), "&#" & Format(Asc(Mid(SC, i, 1)), "00") & ";")
    Next
    CleanXMLText = s
End Function


