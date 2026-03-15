Attribute VB_Name = "MFax"
Option Explicit


Public Sub WindowsFaxServices(Optional SenderName As String, Optional SenderCompany As String, Optional SenderFaxNumber As String, Optional SenderVoiceNumber As String, Optional SenderEmailAddress As String, _
                              Optional RecipientName As String, Optional RecipientCompany As String, Optional RecipientFaxNumber As String, Optional RecipientVoiceNumber As String, Optional RecipientEmailAddress As String, _
                              Optional Subject As String, Optional BillingCode As String, _
                              Optional Attachment1 As String, Optional Attachment2 As String, Optional Attachment3 As String, Optional Attachment4 As String, Optional Attachment5 As String, Optional Attachment6 As String, Optional Attachment7 As String, Optional Attachment8 As String, Optional Attachment9 As String, Optional Attachment10 As String)
On Error GoTo eh
    
'    Dim FaxServer As New FAXCOMEXLib.FaxServer
'    Dim FaxDocument As New FAXCOMEXLib.FaxDocument
'    Dim ServerName As String
'    Dim FileName As String
'    Dim CoverPage As String
'
'
'    ServerName = Trim(HFApp.Options.ValueByName("FaxServer"))
'    FileName = Attachment1
'
'
'    Validate File
'    If Not FileExists(FileName) Then
'        MsgBox "File not found" & vbCrLf & FileName, vbExclamation
'        Exit Sub
'    End If
'    Select Case LCase(FileExt(FileName))
'        Case "tif", "tiff", "pdf", "doc", "xls", "xlsx", "bmp", "jpg", "jpeg", "jpe", "jfif", "png", "gif"
'        Case Else
'            MsgBox "Unsupported file format", vbExclamation
'            Exit Sub
'    End Select
'
'
'    connect to server
'    On Error Resume Next
'    Call FaxServer.Connect(ServerName)
'    Select Case Err.Description
'        Case "The fax server API version does not support the requested operation."
'            MsgBox "Unable to connect to the fax server " & Chr(34) & ServerName & Chr(34) & vbCrLf & vbCrLf & Err.Source & " reports:" & vbCrLf & Err.Description & vbCrLf & "Your user account may have insufficient privileges.", vbExclamation
'            Exit Sub
'        Case Else
'            MsgBox "Unable to connect to the fax server " & Chr(34) & ServerName & Chr(34) & vbCrLf & vbCrLf & Err.Source & " reports:" & vbCrLf & Err.Description, vbExclamation
'            Exit Sub
'    End Select
'    On Error GoTo eh
'
'
'    With FaxDocument
'
'        .Sender.Company = SenderCompany
'        .Sender.BillingCode = BillingCode
'        .Sender.FaxNumber = SenderFaxNumber
'        .Sender.Name = SenderName
'        .Sender.Email = SenderEmailAddress
'
'        .Recipients.Add RecipientFaxNumber, RecipientName
'
'        .Subject = Subject
'        .Body = FileName
'
'        CoverPage = Trim(HFApp.Options.ValueByName("FaxCoverPage"))
'        If CoverPage <> "" Then
'            .CoverPageType = fcptSERVER
'            .CoverPage = CoverPage
'        End If
'
'    End With
'
'
'    On Error Resume Next
'    Call FaxDocument.ConnectedSubmit(FaxServer)
'    If Err.Number <> 0 Then
'        MsgBox "Unable to send " & vbCrLf & vbCrLf & Err.Source & " reports:" & vbCrLf & Err.Description, vbExclamation
'        Exit Sub
'    End If
'    On Error GoTo eh
'
'
'    On Error Resume Next
'    Call FaxServer.Disconnect
'
    
Exit Sub
eh: Call errHandler("WindowsFaxServices")
End Sub



Public Sub GFIFAXmaker(Optional SenderName As String, Optional SenderCompany As String, Optional SenderFaxNumber As String, Optional SenderVoiceNumber As String, Optional SenderEmailAddress As String, _
                       Optional RecipientName As String, Optional RecipientCompany As String, Optional RecipientFaxNumber As String, Optional RecipientVoiceNumber As String, Optional RecipientEmailAddress As String, _
                       Optional Subject As String, Optional BillingCode As String, _
                       Optional Attachment1 As String, Optional Attachment2 As String, Optional Attachment3 As String, Optional Attachment4 As String, Optional Attachment5 As String, Optional Attachment6 As String, Optional Attachment7 As String, Optional Attachment8 As String, Optional Attachment9 As String, Optional Attachment10 As String)

    Dim Folder As String
    Dim i As Integer
    Dim FileName  As String
    
    Folder = RegGetKey(HKEY_LOCAL_MACHINE, "Software\GFI Fax & VOICE\FAXmaker\Config", "XMLAPI")
    If Folder = "" Then
        MsgBox mProductName & " cannot read the GFI FAXmaker configuration settings from" & vbCrLf & _
               "your registry. Please ensure that GFI FAXmaker has been installed. For" & vbCrLf & _
               "assistance contact your system administrator.", vbExclamation, App.ProductName
        Exit Sub
    End If

    FileName = PathAppend(Folder, ForceExt(Replace(Replace(CreateGUID, "}", ""), "{", ""), "xml"))
    
    i = FreeFile
    Open FileName For Output As #i
    Print #i, "<?xml version=""1.0"" encoding=""utf-8""?>"
    Print #i, "<faxmakerdata>"
    Print #i, "  <fields>"
    Print #i, "    <subject>" & Subject & "</subject>"
    Print #i, "    <billingcode>" & BillingCode & "</billingcode>"
    Print #i, "    <resolution>high</resolution>"
    Print #i, "    <faxheader>" & SenderCompany & "</faxheader>"
    If Attachment1 <> "" Then Print #i, "    <attachment>" & Attachment1 & "</attachment>"
    If Attachment2 <> "" Then Print #i, "    <attachment>" & Attachment2 & "</attachment>"
    If Attachment3 <> "" Then Print #i, "    <attachment>" & Attachment3 & "</attachment>"
    If Attachment4 <> "" Then Print #i, "    <attachment>" & Attachment4 & "</attachment>"
    If Attachment5 <> "" Then Print #i, "    <attachment>" & Attachment5 & "</attachment>"
    If Attachment6 <> "" Then Print #i, "    <attachment>" & Attachment6 & "</attachment>"
    If Attachment7 <> "" Then Print #i, "    <attachment>" & Attachment7 & "</attachment>"
    If Attachment8 <> "" Then Print #i, "    <attachment>" & Attachment8 & "</attachment>"
    If Attachment9 <> "" Then Print #i, "    <attachment>" & Attachment9 & "</attachment>"
    If Attachment10 <> "" Then Print #i, "    <attachment>" & Attachment10 & "</attachment>"
    Print #i, "  </fields>"
    Print #i, "  <sender>"
    Print #i, "    <firstname>" & SenderName & "</firstname>"
    Print #i, "    <company>" & SenderCompany & "</company>"
    Print #i, "    <faxnumber>" & SenderFaxNumber & "</faxnumber>"
    Print #i, "    <voicenumber>" & SenderVoiceNumber & "</voicenumber>"
    Print #i, "    <emailaddress>" & SenderEmailAddress & "</emailaddress>"
    Print #i, "  </sender>"
    Print #i, "  <recipients>"
    Print #i, "    <fax>"
    Print #i, "      <recipient>"
    Print #i, "        <firstname>" & RecipientName & "</firstname>"
    Print #i, "        <company>" & RecipientCompany & "</company>"
    Print #i, "        <faxnumber>" & RecipientFaxNumber & "</faxnumber>"
    Print #i, "        <voicenumber>" & RecipientVoiceNumber & "</voicenumber>"
    Print #i, "        <emailaddress>" & RecipientEmailAddress & "</emailaddress>"
    Print #i, "      </recipient>"
    Print #i, "    </fax>"
    Print #i, "  </recipients>"
    Print #i, "</faxmakerdata>"
    Close i

    'GFI FAXmaker deletes the xml and the attachments when the job is processed therefore no clean up is required.

End Sub



'Public Sub RightFax(Optional SenderName As String, Optional SenderCompany As String, Optional SenderFaxNumber As String, Optional SenderVoiceNumber As String, Optional SenderEmailAddress As String, _
'                    Optional RecipientName As String, Optional RecipientCompany As String, Optional RecipientFaxNumber As String, Optional RecipientVoiceNumber As String, Optional RecipientEmailAddress As String, _
'                    Optional Subject As String, Optional BillingCode As String, _
'                    Optional Attachment1 As String, Optional Attachment2 As String, Optional Attachment3 As String, Optional Attachment4 As String, Optional Attachment5 As String, Optional Attachment6 As String, Optional Attachment7 As String, Optional Attachment8 As String, Optional Attachment9 As String, Optional Attachment10 As String)
'
'
'
'
'    Dim FaxServer As Object 'RFCOMAPILib.FaxServer
'    Dim Fax As Object 'RFCOMAPILib.Fax
'
'    Set FaxServer = CreateObject("RFComAPI.FaxServer.1")
'    FaxServer.UseNTAuthentication = True
'
'
'    Set Fax = FaxServer.CreateObject(5) 'coFax
'
'    Fax.ToName = RecipientName
'    Fax.ToFaxNumber = RecipientFaxNumber
'    Fax.ToVoiceNumber = RecipientVoiceNumber
'    Fax.ToCompany = RecipientCompany
'    If Attachment1 <> "" Then Fax.Attachments.Add Attachment1
'    If Attachment2 <> "" Then Fax.Attachments.Add Attachment2
'    If Attachment3 <> "" Then Fax.Attachments.Add Attachment3
'    If Attachment4 <> "" Then Fax.Attachments.Add Attachment4
'    If Attachment5 <> "" Then Fax.Attachments.Add Attachment5
'    If Attachment6 <> "" Then Fax.Attachments.Add Attachment6
'    If Attachment7 <> "" Then Fax.Attachments.Add Attachment7
'    If Attachment8 <> "" Then Fax.Attachments.Add Attachment8
'    If Attachment9 <> "" Then Fax.Attachments.Add Attachment9
'    If Attachment10 <> "" Then Fax.Attachments.Add Attachment10
'    Fax.send
'
'End Sub

'Public Sub WinFax(Optional SenderName As String, Optional SenderCompany As String, Optional SenderFaxNumber As String, Optional SenderVoiceNumber As String, Optional SenderEmailAddress As String, _
'                  Optional RecipientName As String, Optional RecipientCompany As String, Optional RecipientFaxNumber As String, Optional RecipientVoiceNumber As String, Optional RecipientEmailAddress As String, _
'                  Optional Subject As String, Optional BillingCode As String, _
'                  Optional Attachment1 As String, Optional Attachment2 As String, Optional Attachment3 As String, Optional Attachment4 As String, Optional Attachment5 As String, Optional Attachment6 As String, Optional Attachment7 As String, Optional Attachment8 As String, Optional Attachment9 As String, Optional Attachment10 As String)
'On Error GoTo eh
'    Dim e As String
'
'    Const YES = 1
'    Const NO = 0
'
'    Dim i As Long
'    Dim s As String
'    Dim SendObj As Object
'
'e = "Error connecting to WinFax"
'    Set SendObj = CreateObject("WinFax.SDKSend")
'
'    With SendObj
'
'        'recipient
'e = "Error setting recipient properties"
'        .SetClientID MachineName & Timer 'winfax documentation says you have to set this but i'm not sure how it is used
'        .SetTo RecipientName
'        .SetCompany RecipientCompany
'        .SetCountryCode ParsePhoneNumber(RecipientFaxNumber, "country")
'        .SetAreaCode ParsePhoneNumber(RecipientFaxNumber, "area")
'        .SetNumber ParsePhoneNumber(RecipientFaxNumber, "number")
'        .SetExtension ParsePhoneNumber(RecipientFaxNumber, "ext")
'        .AddRecipient
'
'        'coversheet
'e = "Error selecting cover page"
'        .SetUseCover YES
'        .SetQuickCover YES
'        .SetSubject Subject
'
'        'attachements
'e = "Error attaching document"
'        If Attachment1 <> "" Then .AddAttachmentFile Attachment1
'        If Attachment2 <> "" Then .AddAttachmentFile Attachment2
'        If Attachment3 <> "" Then .AddAttachmentFile Attachment3
'        If Attachment4 <> "" Then .AddAttachmentFile Attachment4
'        If Attachment5 <> "" Then .AddAttachmentFile Attachment5
'        If Attachment6 <> "" Then .AddAttachmentFile Attachment6
'        If Attachment7 <> "" Then .AddAttachmentFile Attachment7
'        If Attachment8 <> "" Then .AddAttachmentFile Attachment8
'        If Attachment9 <> "" Then .AddAttachmentFile Attachment9
'        If Attachment10 <> "" Then .AddAttachmentFile Attachment10
'
'
'        'send it
'e = "Error sending document"
'        .SetPreviewFax NO
'        .SetDeleteAfterSend YES
'        .ShowCallProgess NO
'        .ShowSendScreen NO
'        If .send(YES) <> 0 Then
'            Select Case .GetLastError
'                Case 0:  s = "NO_ERROR = 0"
'                Case 1:  s = "NOTHING_TO_SEND = 1"
'                Case 2:  s = "ADDR_EMPTY = 2"
'                Case 3:  s = "NOENTRYID = 3"
'                Case 4:  s = "CANCEL_ON_FILLER = 4"
'                Case 5:  s = "NO_ATTACHMENT_FOUND = 5"
'                Case 6:  s = "NO_COVER_PAGE_FOUND = 6"
'                Case 7:  s = "SET_CURRENT_RECIPIENT = 7"
'                Case 8:  s = "NO_RECIPIENT_SENDJOB = 8"
'                Case 9:  s = "TRANSPORT_TYPE = 9"
'                Case 10: s = "SET_GROUP_RECIPIENT = 10"
'                Case 11: s = "SET_USER_RECIPIENT = 11"
'                Case 12: s = "SET_GROUP_RECIPIENT = 12"
'                Case 13: s = "NUMBER_TOO_LONG = 16"
'                Case 14: s = "BRD_WARNING = 32"
'                Case 15: s = "BRD_INVALID = 64"
'                Case 16: s = "NO_INTERNET_FAX_PSWD = 128"
'            End Select
'        End If
'
'    End With
'
'
'
'Exit Sub
'eh: MsgBox e & vbCrLf & vbCrLf & Err.Description, vbExclamation, "WinFax"
'End Sub



Private Function ParsePhoneNumber(Number, item As String) As String
On Error Resume Next
    'expects:  011(403)341-1234ext1234
    'spaces and no dash are ok but
    '  -area code must be in parenthesis
    '  -"ext" must preceed extension number
    'country,area,ext are optional
    Dim s As String
    Number = LCase(Number)
    Select Case LCase(item)
        Case "country"
            If Len(Number) > 8 Then
                s = Mid(Number, 1, InStr(1, Number, "(") - 1)
            End If
        Case "area"
            s = Mid(Number, InStr(1, Number, "(") + 1, InStr(1, Number, ")") - InStr(1, Number, "(") - 1)
        Case "number"
            If InStr(1, Number, "ext") = 0 Then
                s = Mid(Number, InStr(1, Number, ")") + 1)
            Else
                s = Mid(Number, InStr(1, Number, ")") + 1, InStr(1, Number, "ext") - InStr(1, Number, ")") - 1)
            End If
        Case "ext"
            If InStr(1, Number, "ext") = 0 Then
            Else
                s = Mid(Number, InStr(1, Number, "ext") + 3)
            End If
    End Select
    ParsePhoneNumber = Trim(s)
End Function

