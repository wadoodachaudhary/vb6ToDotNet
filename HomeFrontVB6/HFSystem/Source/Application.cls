VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Application"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
Private Const SRCFILE = "CApp::"

Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Private Const SEM_NOGPFAULTERRORBOX = &H2

Private IsProduction As Boolean


Public Enum Connections
    dbHomefront
    dbEstimating
    dbAccounting
    dbReceivables
    dbAccountingDictionary
    dbAccountingCustom
    dbReceivablesDictionary
    dbTimberlinePVdata
End Enum
Private mDatabases(7) As adodb.Connection
Private mConnectionStrings(7) As String
Private mLicense As New CLicense

Private mSQLEdition As String

Public Enum EstimatingSystems
    esNone = 0
    esPipeline = 1
    esTimberline = 2
End Enum
Public Enum AccountingSystems
    asNone = 0
    asTimberline = 1
    asMasterBuilder = 2
    asQuickBooks = 3
    asSimply = 4
    asMYOB = 5
    asPeachtree = 6
    asXero = 7
    asSpectrum = 8
    asIntacct = 9
    asQuickbooksOnline = 10
    asD365_ABN = 11
End Enum
Private SimplyOdbcUtilities As SimplyWrapper.ODBCUtilities

Public Enum SalesSystems
    asNone = 0
    asHomeFront = 1
    asBuilder1440 = 2
    'asSalesSimplicity = 3
End Enum


Private mUserLevel As Long
Private mLoginID  As String
Private mLoginPswd  As String
Private mDivisionID As String


Public Enum XMLFieldTypes
    a 'BankAcct
    b 'RoutingNum
    c 'Char
    d 'Date
    f 'File
    g 'SocialSecurityNumber
    m 'Memo
    n 'Numeric
    o 'AssemblyPartInclusionFlag
    p 'PhoneNumber
    r 'Terms
    st 'State
    t 'Time
    Y 'YesNo
    z 'zipcode
End Enum

Public Options    As Options


'Private WithEvents mOLApp As Outlook.Application
'Private mOLSec As New OutlookSecurityManager
'Private WithEvents mOLButton As Office.CommandBarButton


Private Sub HookOutlook()
On Error GoTo eh
    Dim i As Long
    
'    Set mOLApp = New Outlook.Application
'    Call mOLSec.ConnectTo(mOLApp)
'
'    For i = 1 To mOLApp.ActiveExplorer.CommandBars("Standard").Controls.Count
'        If mOLApp.ActiveExplorer.CommandBars("Standard").Controls.item(i).Caption = "Add To Job Docs" Then
'            Exit Sub 'Call mOLApp.ActiveExplorer.CommandBars("Standard").Controls.item(i).Delete
'        End If
'    Next
'
'    Set mOLButton = mOLApp.ActiveExplorer.CommandBars("Standard").Controls.Add(MsoControlType.msoControlButton)
'    With mOLButton
'        .Parameter = "HF"
'        .Caption = "Add To Job Docs"
'        .Style = msoButtonIconAndCaption
'        .FaceId = 5
'        .tag = "HF_AddToJobDocs"
'        .Visible = True
'    End With

eh: Exit Sub
End Sub


Private Sub mOLButton_Click(ByVal Ctrl, CancelDefault As Boolean)
'Private Sub mOLButton_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'On Error GoTo eh
'    Dim s As String
'
'    Dim sTo  As String
'    Dim sCC  As String
'    Dim sBCC As String
'    Dim sJob As String
'    Dim iRFI As Long
'    Dim OLMail As Outlook.MailItem
'
'
'
'    mOLSec.DisableOOMWarnings = True
'    On Error Resume Next
'    Set OLMail = mOLApp.ActiveExplorer.Selection.item(1)
'    If Err.Number <> 0 Then
'        mOLSec.DisableOOMWarnings = False
'        Exit Sub
'    End If
'
'
'    With OLMail
'
'        On Error Resume Next
'        s = "select job from messages where id=" & DbQuote(Str, .EntryID)
'        sJob = HFApp.SqlExec(s)(0)
'        On Error GoTo eh
'
'        s = ""
'        s = s & "select job_no Job" & vbCrLf
'        s = s & "      ,Description  JobDesc" & vbCrLf
'        s = s & "      ,null RFI" & vbCrLf
'        s = s & "      ,null Subject" & vbCrLf
'        s = s & "from tbljobs" & vbCrLf
'        s = s & "where inactive=0" & vbCrLf
'        s = s & "union all " & vbCrLf
'        s = s & "select r.Job" & vbCrLf
'        s = s & "      ,j.Description  JobDesc" & vbCrLf
'        s = s & "      ,r.rfiid RFI" & vbCrLf
'        s = s & "      ,r.Subject" & vbCrLf
'        s = s & "from rfis r" & vbCrLf
'        s = s & "join tbljobs j on j.job_no=r.job" & vbCrLf
'        s = s & "where r.status<>'Closed'" & vbCrLf
'        s = s & "order by 1,3" & vbCrLf
'
'        If FPickList.Choose(Databases(dbHomeFront), "Job/RFI", s, sJob) Then
'            sJob = FPickList.SelectedItem("Job")
'            iRFI = Val(FPickList.SelectedItem("RFI"))
'        Else
'            mOLSec.DisableOOMWarnings = False
'            Exit Sub
'        End If
'
'        Call GetRecips(.Recipients, sTo, sCC, sBCC)
'
'        s = ""
'        s = s & "insert into messages(ID,messagetype) values(" & DbQuote(Str, .EntryID) & ",'Email')"
'        Call HFApp.SqlExec(s)
'
'        s = ""
'        s = s & "update messages" & vbCrLf
'        s = s & "set job=" & DbQuote(Str, sJob) & vbCrLf
'        s = s & "   ,RFIID=" & DbQuote(Num, iRFI) & vbCrLf
'        s = s & "   ,SentDate=" & DbQuote(Date, .SentOn) & vbCrLf
'        s = s & "   ,ReceivedDate=" & DbQuote(Date, .ReceivedTime) & vbCrLf
'        s = s & "   ,DueDate=" & DbQuote(Date, .FlagDueBy) & vbCrLf
'        s = s & "   ,Priority=" & DbQuote(Num, .Importance) & vbCrLf
'        s = s & "   ,Sender=" & DbQuote(Str, .SenderName) & vbCrLf
'        s = s & "   ,Recipients=" & DbQuote(Str, sTo) & vbCrLf
'        s = s & "   ,RecipientsCC=" & DbQuote(Str, sCC) & vbCrLf
'        s = s & "   ,RecipientsBCC=" & DbQuote(Str, sBCC) & vbCrLf
'        s = s & "   ,Subject=" & DbQuote(Str, .Subject) & vbCrLf
'        s = s & "   ,Body=" & DbQuote(Str, .Body) & vbCrLf
'        s = s & "   ,Attachments=" & DbQuote(Str, .Attachments.Count) & vbCrLf
'        s = s & "where id=" & DbQuote(Str, .EntryID) & vbCrLf
'        Call HFApp.SqlExec(s)
'
'
'        s = TempFile("msg")
'        Call OLMail.SaveAs(s, OlSaveAsType.olMSG)
'        Call DBPutFile(HFApp.Databases(dbHomeFront), "Messages WHERE id=" & DbQuote(Str, .EntryID), "Original", s)
'        Kill s
'
'
'    End With
'
'
'    mOLSec.DisableOOMWarnings = False
'
'Exit Sub
'eh:
'If Err.Description Like "*duplicate*" Then
'    Resume Next
'Else
'    mOLSec.DisableOOMWarnings = False
'    Call errHandler(SRCFILE & "mOLButton_Click", s)
'End If
End Sub

Friend Property Get License() As CLicense
    Set License = mLicense
End Property
Public Sub SendFAX(Optional SenderName As String, Optional SenderCompany As String, Optional SenderFaxNumber As String, Optional SenderVoiceNumber As String, Optional SenderEmailAddress As String, _
                   Optional RecipientName As String, Optional RecipientCompany As String, Optional RecipientFaxNumber As String, Optional RecipientVoiceNumber As String, Optional RecipientEmailAddress As String, _
                   Optional Subject As String, Optional BillingCode As String, _
                   Optional Attachment1 As String, Optional Attachment2 As String, Optional Attachment3 As String, Optional Attachment4 As String, Optional Attachment5 As String, Optional Attachment6 As String, Optional Attachment7 As String, Optional Attachment8 As String, Optional Attachment9 As String, Optional Attachment10 As String)

    Dim sBillingCode As String
    
    Dim i As String
    Dim s
    
    sBillingCode = BillingCode
    If sBillingCode = "" Then sBillingCode = Trim(HFApp.Options.ValueByName("FaxBillingCode"))

    Select Case HFApp.Options(FaxProvider)
        Case "GFI FAXmaker":         Call GFIFAXmaker(SenderName, SenderCompany, SenderFaxNumber, SenderVoiceNumber, SenderEmailAddress, RecipientName, RecipientCompany, RecipientFaxNumber, RecipientVoiceNumber, RecipientEmailAddress, Subject, sBillingCode, Attachment1, Attachment2, Attachment3, Attachment4, Attachment5, Attachment6, Attachment7, Attachment8, Attachment9, Attachment10)
        'Case "Captaris RightFax":    Call RightFax(SenderName, SenderCompany, SenderFaxNumber, SenderVoiceNumber, SenderEmailAddress, RecipientName, RecipientCompany, RecipientFaxNumber, RecipientVoiceNumber, RecipientEmailAddress, Subject, sBillingCode, Attachment1, Attachment2, Attachment3, Attachment4, Attachment5, Attachment6, Attachment7, Attachment8, Attachment9, Attachment10)
        Case "Windows Fax Services": Call WindowsFaxServices(SenderName, SenderCompany, SenderFaxNumber, SenderVoiceNumber, SenderEmailAddress, RecipientName, RecipientCompany, RecipientFaxNumber, RecipientVoiceNumber, RecipientEmailAddress, Subject, sBillingCode, Attachment1, Attachment2, Attachment3, Attachment4, Attachment5, Attachment6, Attachment7, Attachment8, Attachment9, Attachment10)
        'Case "Symantec WinFax":      Call WinFax(SenderName, SenderCompany, SenderFaxNumber, SenderVoiceNumber, SenderEmailAddress, RecipientName, RecipientCompany, RecipientFaxNumber, RecipientVoiceNumber, RecipientEmailAddress, Subject, sBillingCode, Attachment1, Attachment2, Attachment3, Attachment4, Attachment5, Attachment6, Attachment7, Attachment8, Attachment9, Attachment10)
        Case "", "none":             Call MsgBox("No fax delivery provider has been configured. Please check your system settings.", vbExclamation, mProductName)
        Case Else:                   Err.Raise 5, , "Unsupported Provider"
    End Select

End Sub

Public Function SendMail(Preview As Boolean, SendTo As String, CCTo As String, Subject As String, Body As String, Attachments As String, Optional FromAddress As String) As Boolean
On Error GoTo eh:
'V1------------
'    Dim IOutlook As Object  'Outlook.Application
'    On Error Resume Next
'    Set IOutlook = CreateObject("Outlook.Application")
'    If Err.Number = 0 Then
'        SendMail = SendOutlookMail(IOutlook, Preview, SendTo, CCTo, Subject, Body, Attachments)
'    Else
'        SendMail = SendMapiMail(Preview, SendTo, CCTo, Subject, Body, Attachments)
'    End If
    
'V2------------
'    Dim IOutlook As Object  'Outlook.Application
'    If HFApp.Options.ValueByName("SendPO_ViaMapi") Then
'        SendMail = SendMapiMail(Preview, SendTo, CCTo, Subject, Body, Attachments)
'    Else
'        Set IOutlook = CreateObject("Outlook.Application")
'        SendMail = SendOutlookMail(IOutlook, Preview, SendTo, CCTo, Subject, Body, Attachments)
'    End If

'V3------------
'    Dim IOutlook As Object
'    If HFApp.Options.ValueByName("SendPO_ViaOutlook") Then
'        Set IOutlook = CreateObject("Outlook.Application")
'        SendMail = SendOutlookMail(IOutlook, Preview, SendTo, CCTo, Subject, Body, Attachments)
'    ElseIf HFApp.Options.ValueByName("SendPO_ViaSMTP") Then
'        SendMail = SendSMTPMail(SendTo, CCTo, Subject, Body, Attachments)
'    Else
'        SendMail = SendMapiMail(Preview, SendTo, CCTo, Subject, Body, Attachments)
'    End If
    
'V4------------
    Dim Filename As String
    Dim stream   As New adodb.stream
    Dim s As String
    Dim Server As String
    Dim Port As String
    
    
    'smtp has no preview window so show ours.
    If Preview And (HFApp.Options.ValueByName("SendPO_ViaSMTP") = "True" Or HFApp.Options.ValueByName("SendPO_Via365SMTP")) = "True" Then
        If Not FMail.Edit(SendTo, CCTo, Subject, Body, Attachments) Then
            SendMail = False
            Exit Function
        End If
    End If
    
    
    'trailing semicolons in sendto and ccto cause errors for outlook users.
    SendTo = Trim(SendTo)
    While Right(SendTo, 1) = ";"
        SendTo = Mid(SendTo, 1, Len(SendTo) - 1)
    Wend
    CCTo = Trim(CCTo)
    While Right(CCTo, 1) = ";"
        CCTo = Mid(CCTo, 1, Len(CCTo) - 1)
    Wend
    
    
    If FromAddress = "" Then
    
        If HFApp.Options(SendPO_FromCurrentUsersEmail) = "True" Then
            FromAddress = UserEmail
        End If
        
        If FromAddress = "" Then
            FromAddress = HFApp.Options(Estimator_Email)
        End If
        
    End If
    
    
    stream.Open
    stream.Position = 0
    stream.Charset = "UTF-16"
    stream.WriteText "<?xml version='1.0' encoding=""UTF-16"" ?>" & vbCrLf
    stream.WriteText "<xml>" & vbCrLf
    stream.WriteText "  <system>" & vbCrLf
    
    Select Case True
        Case HFApp.Options.ValueByName("SendPO_ViaOutlook")
            stream.WriteText "    <type>outlook</type>" & vbCrLf
        
        Case HFApp.Options.ValueByName("SendPO_ViaSMTP")
            stream.WriteText "    <type>smtp</type>" & vbCrLf
            s = HFApp.Options.ValueByName("SendPO_SMTPServer")
            Server = Parse(s, 1, ":")
            Port = "" & Parse(s, 2, ":")
            If Port = "" Then Port = HFApp.Options.ValueByName("SmtpPort")
            stream.WriteText "    <server>" & fixxml(Server) & "</server>" & vbCrLf
            stream.WriteText "    <port>" & fixxml(Port) & "</port>" & vbCrLf
            If HFApp.Options.ValueByName("Email_Authenticate") = "True" Then
                stream.WriteText "    <user>" & fixxml(HFApp.Options.ValueByName("Email_User")) & "</user>" & vbCrLf
                stream.WriteText "    <pswd>" & fixxml(HFApp.Options.ValueByName("Email_Pswd")) & "</pswd>" & vbCrLf
            End If
            If HFApp.Options.ValueByName("Email_SSL") = "True" Then
                stream.WriteText "    <ssl>true</ssl>" & vbCrLf
            End If
            If HFApp.Options.ValueByName("Email_STARTTLS") = "True" Then
                stream.WriteText "    <starttls>true</starttls>" & vbCrLf
            End If
            stream.WriteText "    <from>" & fixxml(FromAddress) & "</from>" & vbCrLf
            stream.WriteText "    <reply>" & fixxml(FromAddress) & "</reply>" & vbCrLf
        
        Case HFApp.Options.ValueByName("SendPO_Via365SMTP")
            stream.WriteText "    <type>smtp365</type>" & vbCrLf
            stream.WriteText "    <appid>" & fixxml(HFApp.Options.ValueByName("Smtp365EmailAppID")) & "</appid>" & vbCrLf
            stream.WriteText "    <tenantid>" & fixxml(HFApp.Options.ValueByName("Smtp365EmailTenantID")) & "</tenantid>" & vbCrLf
            stream.WriteText "    <secret>" & fixxml(HFApp.Options.ValueByName("Smtp365EmailSecret")) & "</secret>" & vbCrLf
            stream.WriteText "    <user>" & fixxml(HFApp.Options.ValueByName("Smtp365EmailUser")) & "</user>" & vbCrLf
            stream.WriteText "    <pswd>" & fixxml(HFApp.Options.ValueByName("Smtp365EmailPswd")) & "</pswd>" & vbCrLf
            stream.WriteText "    <from>" & fixxml(FromAddress) & "</from>" & vbCrLf
            stream.WriteText "    <reply>" & fixxml(FromAddress) & "</reply>" & vbCrLf
                    
        Case Else
            stream.WriteText "    <type>mapi</type>" & vbCrLf
    End Select
    
    stream.WriteText "  </system>" & vbCrLf
    stream.WriteText "  <mail>" & vbCrLf
    stream.WriteText "    <to>" & fixxml(SendTo) & "</to>" & vbCrLf
    stream.WriteText "    <cc>" & fixxml(CCTo) & "</cc>" & vbCrLf
    'stream.WriteText "    <bcc>" &  & "</bcc>" & vbCrLf
    'stream.WriteText "    <priority>" &  & "</priority>" & vbCrLf
    stream.WriteText "    <subject>" & fixxml(Subject) & "</subject>" & vbCrLf
    stream.WriteText "    <body>" & fixxml(Body) & "</body>" & vbCrLf
    stream.WriteText "    <attachment>" & fixxml(Attachments) & "</attachment>" & vbCrLf
    stream.WriteText "    <display>" & IIf(Preview = True, "yes", "no") & "</display>" & vbCrLf
    stream.WriteText "  </mail>" & vbCrLf
    stream.WriteText "</xml>" & vbCrLf
    On Error Resume Next
    Filename = TempFile("xml")
    Kill Filename
    On Error GoTo eh
    
    Call stream.SaveToFile(Filename, adSaveCreateOverWrite)
    Call stream.SaveToFile(PathAppend(GetFolderPath(CSIDL_LOCAL_APPDATA), "HomeFront\LastEmailMessage.xml"), adSaveCreateOverWrite)
    
    s = PathAppend(App.Path, "hfmailer.exe") & " " & Filename
    Call ShellAndLoop(s, vbHide)
    Exit Function
eh:
    If Err.Number <> 0 Then
        MsgBox "Error sending email" & vbCrLf & Err.Description
    End If
End Function

Private Function fixxml(s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "'", "&apos;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    fixxml = s
End Function


Public Property Get DivisionID() As String
    DivisionID = Val("" & mDivisionID)
End Property


Public Property Get LoginPswd() As String
    LoginPswd = mLoginPswd
End Property
Public Property Get LoginID() As String
    LoginID = mLoginID
End Property
Friend Sub SetLoginID(uid As String, pwd As String)
    mLoginID = uid
    mLoginPswd = pwd
End Sub

Friend Sub SetUserLevel(UserLevel As Long)
    mUserLevel = UserLevel
End Sub

Public Function SetDivision(DivisionID As String)
    Call SetDivisionID(DivisionID)
End Function
Friend Sub SetDivisionID(DivisionID As String)
    mDivisionID = DivisionID
End Sub
Public Property Get LoginDSN() As String
    LoginDSN = Parse(Parse(HFApp.ConnectionString(dbHomefront), 2, "DSN="), 1, ";")
End Property
Public Property Get LoginDBName() As String
    LoginDBName = HFApp.Databases(dbHomefront).DefaultDatabase
End Property
Public Property Get LoginDBUID() As String
    LoginDBUID = Parse(Parse(HFApp.ConnectionString(dbHomefront), 2, "UID="), 1, ";")
End Property
Public Property Get LoginDBPWD() As String
    LoginDBPWD = Parse(Parse(HFApp.ConnectionString(dbHomefront), 2, "PWD="), 1, ";")
End Property




Private Sub Class_Initialize()
    Set HFApp = Me
    Set mDatabases(dbHomefront) = New adodb.Connection
    Set Options = New Options
End Sub

Public Sub Activations()
    
End Sub

Public Sub About()
    FAbout.Show vbModal
End Sub


Public Sub Logout()
On Error Resume Next
    Closing = True
    Call Options.SaveData
    
        
    
    Call SetErrorMode(SEM_NOGPFAULTERRORBOX)
    
    Call mDatabases(dbHomefront).Close
    Call mDatabases(dbEstimating).Close
    Call mDatabases(dbAccounting).Close
    
    Call mDatabases(dbReceivables).Close
    Call mDatabases(dbAccountingDictionary).Close
    Call mDatabases(dbAccountingCustom).Close
    Call mDatabases(dbReceivablesDictionary).Close
    Call mDatabases(dbTimberlinePVdata).Close
    
    
'    Call mOLButton.Delete
    
    Set mDatabases(dbHomefront) = Nothing
    Set mDatabases(dbEstimating) = Nothing
    Set mDatabases(dbAccounting) = Nothing
    Set mDatabases(dbReceivables) = Nothing
    Set mDatabases(dbAccountingDictionary) = Nothing
    Set mDatabases(dbAccountingCustom) = Nothing
    Set mDatabases(dbReceivablesDictionary) = Nothing
    Set mDatabases(dbTimberlinePVdata) = Nothing
    Set HFApp = Nothing
    
End Sub

Public Function Login(ByVal ProductCode As String, ByVal ProductName As String, ByVal ProductVersion As String, ByVal EXEPath As String) As Boolean

    Dim rs As Recordset
    Dim s As String
    
    mProductName = ProductName
    mProductCode = ProductCode
    mModuleCode = ProductCode
    mModuleName = ProductName
    mModuleVersion = ProductVersion
    'mSystemVersion = App.Major & "." & App.Minor & IIf(App.Revision = 0, "", " R" & App.Revision)
    mSystemVersion = App.Major & "." & App.Minor
    
    mExePath = EXEPath
    If Right(mExePath, 1) = "\" Then mExePath = left(mExePath, Len(mExePath) - 1)
        
    'make sure working folders exists...
    On Error Resume Next
    Call CreatePath("working folder", PathAppend(GetFolderPath(CSIDL_LOCAL_APPDATA), "Homefront", Parse(mExePath, Parse(mExePath, , "\"), "\")))
    Call CreatePath("working folder", PathAppend(GetFolderPath(CSIDL_LOCAL_APPDATA), "HomeFront Software Ltd\HFSales"))
    Call mLicense.ReadLicense(HFApp.LicenseFile)
    On Error GoTo 0


    'move old user settings file from local app data to the HF system folder path
    If Not FileExists(AppIni) Then
        Call CreatePath("", PathAppend(HFApp.SystemFolder, "User Settings"))
        s = PathAppend(GetFolderPath(CSIDL_LOCAL_APPDATA), "Homefront\" & App.EXEName, "settings.ini")
        If FileExists(s) Then Call FileCopy(s, AppIni)
    End If
    

    'warn that homefront.ini is depreciated
    If FileExists(PathAppend(mExePath, "Homefront.ini")) Then
        s = "A depreciated homefront.ini file has been detected. HomeFront uses a new configuration model. Please contact support."
        Call MsgBox(s, vbExclamation, App.ProductName)
        Login = False
        Exit Function
    End If




    If mTrustedConnection = "" Then
        If FLogin.Login() Then
            Login = True
            Call WriteUserSession
            Call HookOutlook
        End If
    Else
    
        Screen.MousePointer = vbHourglass
        
        'open databases
        On Error Resume Next
        Databases(dbHomefront).Close
        ConnectionString(dbHomefront) = mTrustedConnection
        Databases(dbHomefront).Open ConnectionString(dbHomefront)
        If Databases(dbHomefront).State <> adStateOpen Then
            Screen.MousePointer = vbDefault
            MsgBox "Unable to open the Homefront database" & IIf(Err.Number <> 0, vbCrLf & Parse(Err.Description, Parse(Err.Description, , "]"), "]"), ""), vbCritical
            Exit Function
        End If
        On Error GoTo 0
        mLoginID = IIf(mTrustedUser = "", "system", mTrustedUser)
        mLoginPswd = ""
        Options.ReadData
        
        Login = True
        
        
        Screen.MousePointer = vbDefault
        
    End If

End Function


Friend Property Get SQLEdition() As String
    If mSQLEdition = "" Then
        mSQLEdition = SqlExec("SELECT cast(SERVERPROPERTY('edition') as varchar(150))", dbHomefront)(0)
    End If
    SQLEdition = mSQLEdition
End Property

Private Sub WriteUserSession()
On Error Resume Next
    Dim s As String
    
    s = ""
    s = s & "delete from UserSessions where LoginDate <= getdate()-14 and WorkstationName=host_name()"
    Call Me.SqlExec(s)
    
    s = ""
    s = s & "insert into UserSessions(Host_ID,LoginName,WorkstationName,LoginDate,HFLoginID)" & vbCrLf
    s = s & "select Host_ID(),System_User,host_name(),GetDate()," & DbQuote(Str, mLoginID) & vbCrLf
    Call Me.SqlExec(s, dbHomefront)
    
    s = ""
    s = s & "UPDATE UserSessions" & vbCrLf
    s = s & "   SET LoginName=System_User" & vbCrLf
    s = s & "      ,LoginDate=GetDate()" & vbCrLf
    s = s & "      ,HFLoginID=" & DbQuote(Str, mLoginID) & vbCrLf
    s = s & "WHERE Host_ID=Host_ID() AND WorkstationName=host_name()" & vbCrLf
    Call Me.SqlExec(s, dbHomefront)

End Sub

Friend Property Get LicenseFile() As String
    LicenseFile = PathAppend(SystemFolder, "license.lic")
End Property


Public Property Get SystemFolder() As String
    Dim s As String
    s = RegGetKey(HKEY_LOCAL_MACHINE, "SOFTWARE\HomeFront", "HomeFrontLicensePath")
    If Right(s, 1) <> "\" Then s = s & "\"
    SystemFolder = s
End Property
Public Property Get ConnectionString(d As Connections) As String
    ConnectionString = mConnectionStrings(d)
End Property
Friend Property Let ConnectionString(d As Connections, RHS As String)
    mConnectionStrings(d) = RHS
End Property


Public Property Get Databases(d As Connections) As adodb.Connection
    Set Databases = mDatabases(d)
End Property
Friend Property Set Databases(d As Connections, RHS As adodb.Connection)
    Set mDatabases(d) = RHS
End Property



Public Function SqlExec(sql As String, Optional Database As Connections, Optional RecordsAffected As Long) As adodb.Recordset
'
'    mDatabases(Database).CommandTimeout = 9999999
'    Call ConvertSQL(sql, mDatabases(Database))
'    Set SqlExec = mDatabases(Database).Execute(sql, RecordsAffected)
'
' THIS WAS CHANGED BECAUSE ONCE YOU READ COLUMN THREE, THE DATA FOR COLUMNS ONE
' AND TWO IS NO LONGER AVAILABLE. THE SQL SERVER FIREHOSE CURSOR SUCKS THAT WAY.
' YOU MUST CHECK THE DBMS NAME, AS TIMBERLINE'S SIMBA DRIVER ONLY SUPPORTS DYNAMIC
' CURSORS ON ROW RETURNING QUERIES. IT SUPPORTS FIREHOSE ON ALL AND DOESN'T MESS
' IT UP THE WAY MICROSOFT DOES.
'
'    mDatabases(Database).CommandTimeout = 9999999
'    Call ConvertSQL(sql, mDatabases(Database))
'    If mDatabases(Database).Properties("DBMS Name") = "Microsoft SQL Server" And Not sql Like "*for xml explicit*" Then
'        Dim rs As New ADODB.Recordset
'        Call rs.Open(sql, mDatabases(Database), adOpenDynamic)
'        Set SqlExec = rs
'    Else
'        Set SqlExec = mDatabases(Database).Execute(sql, RecordsAffected)
'    End If
'
' THIS WAS CHANGED BECAUSE YOU CANT LOOP THRU A DYNAMIC CURSOR WHILE RUNNING INSERT OR
' UPDATE STATEMENTS AGAINST THE SAME SQL DATABASE. HFSync THRU AN ERROR LIKE connection
' busy with results from previous hstmt. I CHANGED adOpenDynamic TO adOpenForwardOnly
' WHICH IS WHAT THE .Execute(sql) FUNCTION IS SUPPOSED TO RETURN IN THE FIRST FRICKEN
' PLACE. SURE HOPE THIS WORKS...
'
'    mDatabases(Database).CommandTimeout = 9999999
'    Call ConvertSQL(sql, mDatabases(Database))
'    If mDatabases(Database).Properties("DBMS Name") = "Microsoft SQL Server" And Not sql Like "*for xml explicit*" Then
'        Dim rs As New ADODB.Recordset
'        rs.CursorType = adOpenForwardOnly
'        Call rs.Open(sql, mDatabases(Database), adOpenForwardOnly)
'        Set SqlExec = rs
'    Else
'        Set SqlExec = mDatabases(Database).Execute(sql, RecordsAffected)
'    End If
'
'
' THIS WAS CHANGED BECAUSE WE STILL GET THE for xml clause is not allowed in a cursor statement ERROR.
' PARTLY BECAUSE WE ARE USING for xml auto AS WELL AS for xml explicit BUT THERE IS MORE TO IT ALSO. IT
' SEEMS TO WORK IF WE CREATE A NEW CONNECTION. IT FAILS IF WE USE THE mDatabases ARRAY.
'
    Dim c As New Connection
    Dim rs As New Recordset
    'Dim sql As String
    Dim s As String
    Dim dbmsname As String
    
    mDatabases(Database).CommandTimeout = 9999999
        
    Call ConvertSQL(sql, mDatabases(Database))
    
    On Error Resume Next
    dbmsname = mDatabases(Database).Properties("DBMS Name")
    On Error GoTo 0
    
    If dbmsname = "Microsoft SQL Server" And Not sql Like "*for xml *" And RecordsAffected <> 99 Then
        If LCase(Mid(LTrim(sql), 1, 7)) = "insert " Or LCase(Mid(LTrim(sql), 1, 7)) = "update " Or LCase(Mid(LTrim(sql), 1, 7)) = "delete " Then
            s = "SET NOCOUNT ON" & vbCrLf
            s = s & "DECLARE @RetryCounter Int" & vbCrLf
            s = s & "SET @RetryCounter = 1" & vbCrLf
            s = s & "RETRY:" & vbCrLf
            s = s & "   BEGIN TRANSACTION" & vbCrLf
            s = s & "   BEGIN TRY" & vbCrLf
            s = s & "      " & sql & vbCrLf
            s = s & "      COMMIT TRANSACTION" & vbCrLf
            s = s & "   END TRY" & vbCrLf
            s = s & "   BEGIN CATCH" & vbCrLf
            s = s & "   ROLLBACK TRANSACTION" & vbCrLf
            s = s & "   DECLARE @DoRetry bit" & vbCrLf
            s = s & "   SET @DoRetry = 0" & vbCrLf
            s = s & "   DECLARE @ErrMsg nvarchar(4000), @ErrSeverity int" & vbCrLf
            s = s & "   SET @ErrMsg = ERROR_MESSAGE()" & vbCrLf
            s = s & "   SET @ErrSeverity = ERROR_SEVERITY()" & vbCrLf
            s = s & "   insert into SQLErrorLog(ErrorMessage, SQLStatement) Values(@ErrMsg," & DbQuote(Str, sql) & ")" & vbCrLf
            s = s & "   IF ERROR_NUMBER() = 1205 " & vbCrLf '-- Deadlock Error Number
            s = s & "   BEGIN" & vbCrLf
            s = s & "       SET @DoRetry = 1" & vbCrLf
            s = s & "   END" & vbCrLf
            s = s & "   IF @DoRetry = 1" & vbCrLf
            s = s & "   BEGIN" & vbCrLf
            s = s & "       SET @RetryCounter = @RetryCounter + 1 -- Increment Retry Counter By one" & vbCrLf
            s = s & "       IF (@RetryCounter > 3) -- Check whether Retry Counter reached to 3" & vbCrLf
            s = s & "       BEGIN" & vbCrLf
            s = s & "           RAISERROR('%s', 18, 1, @ErrMsg) -- Raise Error Message" & vbCrLf
            s = s & "               -- if still deadlock occurred after three retries" & vbCrLf
            s = s & "       END" & vbCrLf
            s = s & "       ELSE" & vbCrLf
            s = s & "       BEGIN" & vbCrLf
            s = s & "           WAITFOR DELAY '00:00:00.05' -- Wait for 5 ms" & vbCrLf
            s = s & "           GOTO RETRY  -- Go to Label RETRY" & vbCrLf
            s = s & "       END" & vbCrLf
            s = s & "   END" & vbCrLf
            s = s & "   ELSE" & vbCrLf
            s = s & "   BEGIN" & vbCrLf
            s = s & "       RAISERROR('%s', @ErrSeverity, 1, @ErrMsg)" & vbCrLf
            s = s & "   END" & vbCrLf
            s = s & "   END CATCH" & vbCrLf
            
            Call c.Open(HFApp.ConnectionString(Database))
            c.CommandTimeout = 9999999
            c.BeginTrans
            
            
'Call LogInfoToFile("c:\log.txt", s)
            Set rs = c.Execute(s, RecordsAffected, adOpenForwardOnly)
            c.CommitTrans
            Set SqlExec = rs
            c.Close
           
        Else
        
            Call c.Open(HFApp.ConnectionString(Database))
            c.CommandTimeout = 9999999
            'rs.CursorType = adOpenForwardOnly
'Call LogInfoToFile("c:\log.txt", sql)
            Call rs.Open(sql, c, adOpenForwardOnly)
            Set SqlExec = rs
            'doesn't close the connection or the rs will be closed as well
            'the calling proc wont be able to read the resultset
            'c.Close
            
            
        End If
    Else
        'dont open new connections. each time you open a connection to sage50 it consumes a license. they are left in memory until the app closes.
        'this eats up all the sage 50 licenses and people get angry.
        
        'Call c.Open(HFApp.ConnectionString(Database))
        'c.CommandTimeout = 99999998
        If (mProductCode = "hfpayables" Or mProductCode = "JobSimplicity") And HFApp.Options(AccountingSystem) = AccountingSystems.asTimberline Then
            If sql Like "*Select *" And Database = dbTimberlinePVdata Then
                sql = Replace(sql, "master_jcm_record_2", "master_jcm_r2")
                sql = Replace(sql, "master_jcm_record_3", "master_jcm_r3")
                sql = Replace(sql, "master_jcm_record_4", "master_jcm_r4")
                sql = Replace(sql, "master_jcm_record_12", "master_jcm_r12")
                sql = Replace(sql, "master_jcm_record_17", "master_jcm_r17")
                sql = Replace(sql, "master_jcm_record_16", "master_jcm_r16")
                sql = Replace(sql, "master_jcm_record_1_1", "master_jcm_r1")
                sql = Replace(sql, "master_jcm_record_13", "master_jcm_r13")
                sql = Replace(sql, "master_apm_record_9", "master_apm_r9")
                sql = Replace(sql, "master_apm_record_1", "master_apm_r1")
                sql = Replace(sql, "master_txm_record_2", "master_txm_r2")
                sql = Replace(sql, "master_glm_record_1", "master_glm_r1")
                sql = Replace(sql, "t.""group""", "t.""ggroup""")
            End If
        End If
        'Set rs = c.Execute(sql, RecordsAffected, adOpenForwardOnly)
'Call LogInfoToFile("c:\log.txt", sql)
        Set rs = HFApp.Databases(Database).Execute(sql, RecordsAffected, adOpenForwardOnly)
        Set SqlExec = rs
        

    End If


End Function


Private Sub ConvertSQL(sql As String, c As adodb.Connection)
On Error GoTo eh

    Dim i As Long
    Dim pieces() As String
    Dim s As String
    
    Select Case c.Properties("DBMS Name")
        
        Case "Pervasive.SQL"
            pieces = split(sql, "'")
            For i = LBound(pieces) To UBound(pieces) Step 2
                s = pieces(i)
                s = Replace(s, "dbo.", "", , , vbTextCompare)
                pieces(i) = s
            Next
            s = ""
            For i = LBound(pieces) To UBound(pieces)
                s = s & "'" & pieces(i)
            Next
            sql = Mid(s, 2)
                        
        Case "Microsoft SQL Server"
            If LCase(left(Trim(sql), 11)) <> "alter table" Then
                pieces = split(sql, "'")
                For i = LBound(pieces) To UBound(pieces) Step 2
                    s = pieces(i)
                    s = Replace(s, "IFNULL", "ISNULL", , , vbTextCompare)
                    s = Replace(s, "CURDATE", "GETDATE", , , vbTextCompare)
                    's = Replace(s, "NOW()", "GETDATE()", , , vbTextCompare)
                    s = Replace(s, "LONGVARCHAR", "VARCHAR(MAX)", , , vbTextCompare)
                    's = Replace(s, "DOUBLE", "DECIMAL", , , vbTextCompare)
                    's = Replace(s, " IDENTITY", " INTEGER IDENTITY", , , vbTextCompare)
                    's = Replace(s, " DATE", " DATETIME", , , vbTextCompare)
                    's = Replace(s, " DATETIME_", " DATE_", , , vbTextCompare)
                    s = Replace(s, " DATETIMETIME", " DATETIME", , , vbTextCompare)
                    s = Replace(s, " DATETIMEDIFF", " DATEDIFF", , , vbTextCompare)
                    s = Replace(s, " TIMESTAMP", " DATETIME", , , vbTextCompare)
                    's = Replace(s, " TIME", " DATETIME", , , vbTextCompare)
                    s = Replace(s, "UCASE(", "UPPER(", , , vbTextCompare)
                    s = Replace(s, "LCASE(", "LOWER(", , , vbTextCompare)
                    s = Replace(s, "CURTIME(", "GETDATE(", , , vbTextCompare)
                    's = Replace(s, "CONVERT(", "CAST(", , , vbTextCompare)
                    s = Replace(s, ", SQL_VARCHAR)", " AS VARCHAR)", , , vbTextCompare)
                    s = Replace(s, ",SQL_VARCHAR)", " AS VARCHAR)", , , vbTextCompare)
                    s = Replace(s, ", SQL_INTEGER)", " AS INTEGER)", , , vbTextCompare)
                    s = Replace(s, ",SQL_INTEGER)", " AS INTEGER)", , , vbTextCompare)
                    's = Replace(s, " MODIFY ", " ALTER COLUMN ", , , vbTextCompare)
                    s = Replace(s, "CAST(VARCHAR", "CONVERT(VARCHAR", , , vbTextCompare)
                    s = Replace(s, "INTEGER IDENTITY_INSERT", "IDENTITY_INSERT", , , vbTextCompare)
                    pieces(i) = s
                Next
                s = ""
                For i = LBound(pieces) To UBound(pieces)
                    s = s & "'" & pieces(i)
                Next
                sql = Mid(s, 2)
            End If
'            pieces = tokenize(sql, "'""")
'            For i = LBound(pieces) To UBound(pieces) Step 2
'
'                s = pieces(i)
'                s = Replace(s, "IFNULL", "ISNULL", , , vbTextCompare)
'                s = Replace(s, "CURDATE", "GETDATE", , , vbTextCompare)
'                s = Replace(s, "NOW()", "GETDATE()", , , vbTextCompare)
'                s = Replace(s, "LONGVARCHAR", "VARCHAR(8000)", , , vbTextCompare)
'                s = Replace(s, "DOUBLE", "DECIMAL", , , vbTextCompare)
'                s = Replace(s, " IDENTITY", " INTEGER IDENTITY", , , vbTextCompare)
'                s = Replace(s, " DATE", " DATETIME", , , vbTextCompare)
'                s = Replace(s, " DATETIME_", " DATE_", , , vbTextCompare)
'                s = Replace(s, " DATETIMETIME", " DATETIME", , , vbTextCompare)
'                s = Replace(s, " TIMESTAMP", " DATETIME", , , vbTextCompare)
'                s = Replace(s, " TIME", " DATETIME", , , vbTextCompare)
'                s = Replace(s, "UCASE(", "UPPER(", , , vbTextCompare)
'                s = Replace(s, "LCASE(", "LOWER(", , , vbTextCompare)
'                s = Replace(s, "CURTIME(", "GETDATE(", , , vbTextCompare)
'                s = Replace(s, "CONVERT(", "CAST(", , , vbTextCompare)
'                s = Replace(s, ", SQL_VARCHAR)", " AS VARCHAR)", , , vbTextCompare)
'                s = Replace(s, ",SQL_VARCHAR)", " AS VARCHAR)", , , vbTextCompare)
'                s = Replace(s, ", SQL_INTEGER)", " AS INTEGER)", , , vbTextCompare)
'                s = Replace(s, ",SQL_INTEGER)", " AS INTEGER)", , , vbTextCompare)
'                s = Replace(s, " MODIFY ", " ALTER COLUMN ", , , vbTextCompare)
'                s = Replace(s, "CAST(VARCHAR", "CONVERT(VARCHAR", , , vbTextCompare)
'                s = Replace(s, "INTEGER IDENTITY_INSERT", "IDENTITY_INSERT", , , vbTextCompare)
'                pieces(i) = s
'            Next
'            s = ""
'            For i = LBound(pieces) To UBound(pieces) - 1
'                s = s & pieces(i)
'            Next
'            sql = s & pieces(UBound(pieces))
            
    End Select
    Exit Sub
eh:
    Select Case Err.Number
    
        Case 91   'Object variable or With block variable not set
            Err.Raise vbObjectError, "App::ConvertSQL()", "Database connection is nothing"
            
        Case 3265 'Item cannot be found in the collection corresponding to the requested name or ordinal.
            Exit Sub
            
        Case Else
            Err.Raise Err.Number
            
    End Select
End Sub



Public Function SqlIdentity(TableName As String, Optional Database As Connections) As Long
On Error Resume Next
    If mDatabases(Database).Properties("DBMS Name") = "Microsoft SQL Server" Then
        SqlIdentity = Val("" & SqlExec("SELECT IDENT_CURRENT('" & TableName & "')", Database)(0))
    Else
        SqlIdentity = Val("" & SqlExec("SELECT @@IDENTITY", Database)(0))
    End If
End Function


Public Sub WriteAuditLog(task As String, Description As String, Optional XRef1 As String, Optional XRef2 As String, Optional XRef3 As String, Optional XRef4 As String, Optional XRef5 As String, Optional XRef6 As String)
    Dim s As String
    
    s = ""
    s = s & "INSERT INTO UserLog(Workstation,UserID,LogDate,LogTime,Task,Description)" & vbCrLf
    s = s & "VALUES(" & DbQuote(Str, MachineName(), , , 30) & vbCrLf
    s = s & "      ," & DbQuote(Str, LoginID, , , 10) & vbCrLf
    s = s & "      ,GETDATE()" & vbCrLf
    s = s & "      ," & DbQuote(Str, Format(Now(), "hh:mm AMPM"), , , 10) & vbCrLf
    s = s & "      ," & DbQuote(Str, task, , , 50) & vbCrLf
    s = s & "      ," & DbQuote(Str, Description, , , 200) & ")"
    Call SqlExec(s)
    
End Sub




Public Sub LockRecord(TableName As String, RecordKey)
On Error GoTo eh
    Dim s As String
    
    s = ""
    s = s & "SELECT * FROM TableLog" & vbCrLf
    s = s & " WHERE TableName=" & DbQuote(Str, TableName) & vbCrLf
    s = s & "   AND KeyRecord=" & DbQuote(Str, RecordKey) & vbCrLf
    If SqlExec(s).EOF Then
    
        s = ""
        s = s & "INSERT INTO TableLog(TableName,KeyRecord)" & vbCrLf
        s = s & "VALUES(" & DbQuote(Str, TableName) & vbCrLf
        s = s & "      ," & DbQuote(Str, RecordKey) & ")"
        Call SqlExec(s)
    
        s = ""
        s = s & "UPDATE TableLog" & vbCrLf
        s = s & "   SET RecordLocked=1" & vbCrLf
        s = s & "      ,User_ID=" & DbQuote(Str, LoginID) & vbCrLf
        s = s & "      ,WorkStation_Name=" & DbQuote(Str, MachineName) & vbCrLf
        s = s & "      ,LockDate=GETDATE()" & vbCrLf
        s = s & " WHERE TableName=" & DbQuote(Str, TableName) & vbCrLf
        s = s & "   AND KeyRecord=" & DbQuote(Str, RecordKey) & vbCrLf
        Call SqlExec(s)
        
    End If
Exit Sub
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Sub

Public Sub UnLockRecord(TableName As String, RecordKey)
    Dim s As String
    s = ""
    s = s & "DELETE FROM TableLog" & vbCrLf
    s = s & " WHERE TableName=" & DbQuote(Str, TableName) & vbCrLf
    s = s & "   AND KeyRecord=" & DbQuote(Str, RecordKey) & vbCrLf
    Call SqlExec(s)
End Sub

Public Property Get RecordLockedBy(TableName As String, RecordKey) As String
On Error Resume Next
    Dim s As String
    
    s = ""
    s = s & "SELECT User_ID" & vbCrLf
    s = s & "FROM TableLog" & vbCrLf
    s = s & "WHERE TableName=" & DbQuote(Str, TableName) & vbCrLf
    s = s & "  AND RecordLocked=1" & vbCrLf
    s = s & "  AND KeyRecord=" & DbQuote(Str, RecordKey) & vbCrLf
    s = s & "  AND User_ID<>" & DbQuote(Str, LoginID) & vbCrLf
    RecordLockedBy = "" & SqlExec(s)(0)
    
End Property

Private Property Get UserEmail() As String
On Error Resume Next
    UserEmail = "" & SqlExec("SELECT email FROM User_Manager WHERE User_ID=" & DbQuote(Str, LoginID), dbHomefront)(0)
End Property

Public Property Get UserPermission(Permission As String) As Boolean
On Error Resume Next
    Static p() As String
    
    Dim rs As Recordset
    Dim i As Long
    
    ReDim p(0) As String
    
    Set rs = SqlExec("SELECT * FROM User_Manager WHERE User_ID=" & DbQuote(Str, LoginID), dbHomefront)
    For i = 0 To rs.fields.Count - 1
        If rs.fields(i).Type = adBoolean And "" & rs(i) = "True" Then
            ReDim Preserve p(UBound(p) + 1)
            p(UBound(p)) = UCase("" & rs.fields(i).Name)
        End If
    Next


    Select Case Permission
        Case "AdminUser":           UserPermission = mUserLevel = 1  'administrator
        Case "SalesManagerUser":    UserPermission = mUserLevel = 2  'salesmanger
        Case "AccountingUser":      UserPermission = mUserLevel = 3  'accounting
        Case "SalesPersonUser":     UserPermission = mUserLevel = 4  'salesperson
        Case "DCSalesPersonUser":   UserPermission = mUserLevel = 5  'dc salesperson
        Case "ReadOnlyUser":        UserPermission = mUserLevel = 6  'readonly
        Case "EstimatingUser":      UserPermission = mUserLevel = 7  'estimating
        Case Else
            UserPermission = False
            For i = 1 To UBound(p)
                If UCase(Permission) = p(i) Then
                    UserPermission = True
                    Exit Property
                End If
            Next
    End Select
    
End Property

Public Sub POInquiry(Optional PO As String)
    Call FPOInquiry.ShowForm(PO)
End Sub

Public Sub EditJob(Optional Job As String)
    Call FJob.EditJob(Job)
End Sub
Public Sub EditVendor(Optional vendor As String)
    Call FVendor.ShowForm(vendor)
End Sub

Private Sub EditCustomer(Optional Customer As String)
    Call FCustomer.ShowForm(Customer)
End Sub

Public Sub EditPOIndexList()
    Call RunTask("EditPOIndexes")
End Sub
Public Sub EditCommunities()
    Call FCommunities.Show(vbModal)
End Sub
Public Sub EditProjectManagerList()
    Call FProjectManagers.Show(vbModal)
End Sub


Private Sub EditSecurity()
    If HFApp.UserPermission("ToolsUsrAdmin") Then
        Call FUserPermissions.Show(vbModal)
    Else
        MsgBox "Your security permissions don't allow you to do that. Sorry.", vbInformation, App.ProductName
    End If
End Sub

Public Sub EditOptions()
    If HFApp.UserPermission("ToolsSysAdmin") Then
        Call FOptions.Show(vbModal)
    Else
        MsgBox "Your security permissions don't allow you to do that. Sorry.", vbInformation, App.ProductName
    End If
End Sub

Public Function ImportData(TableName As String, KeyColumns As String, Title As String, Optional RequiredColumns As String, Optional ColumnMappings As String) As Boolean
    ImportData = FDataImport.ImportData(TableName, KeyColumns, Title, RequiredColumns, ColumnMappings)
End Function

Public Sub RunTask(TaskName As String)
'On Error Resume Next
    
    Dim s As String
    
    Dim task As String
    Dim p1 As String
    Dim p2 As String
    Dim p3 As String
    Dim p4 As String
    Dim p5 As String
    Dim p6 As String
    Dim p7 As String
    Dim p8 As String
    Dim p9 As String

    task = Parse(TaskName, 1, "|")
    p1 = Parse(TaskName, 2, "|")
    p2 = Parse(TaskName, 3, "|")
    p3 = Parse(TaskName, 4, "|")
    p4 = Parse(TaskName, 5, "|")
    p5 = Parse(TaskName, 6, "|")
    p6 = Parse(TaskName, 7, "|")
    p7 = Parse(TaskName, 8, "|")
    p8 = Parse(TaskName, 9, "|")
    p9 = Parse(TaskName, 10, "|")
    
    Select Case task
        
        Case "EditAttachments":                 Call FAttachments.ShowForm("Attachments", p1, p2, p3, p4, p5, p6, p7, p8)
        Case "AddAttachment":                   Call FAttachments.AddFile(p1, p2, p3 = "True", p4)
        Case "EditItemAttachments":             Call FAttachments.ShowForm(p1, p2, p3, p4, p5, p6, p7, p8, p9)
                
        Case "EditOptions":                     Call EditOptions
        Case "EditSecurity":                    Call EditSecurity
        Case "LockedRecords":                   FTableLocks.Show vbModal
        Case "EditJCCategories":                Call EditJCCategories
        Case "EditJCCostCodes":                 Call EditJCCostCodes
        Case "EditAttributeLists":              Call FAttributeLists.ShowForm(Val(p1))
        Case "JournalEntries":                  Call FDirectCosts.Show(vbModal)
        Case "ShowUserMonitor":                 Call FInquiry.ShowInquiry("Logged In Users", "select * from loggedinusers order by module")
        Case "RFIs":                            Call FRfis.ShowForm(p1, Val(p2))
        Case "POInquiry":                       Call POInquiry(p1)
        Case "EditCustomer":                    Call EditCustomer(p1)
        Case "EditVendor":                      Call EditVendor(p1)
        Case "EditJob":                         Call EditJob(p1)
        Case "EditPOItems":                     FPOItems.Show vbModal
        Case "EditPOIndexes":                   FPOFormats.Show vbModal
        Case "EditDefaultVendors":              FDefaultVendors.Show vbModal
        Case "EditProjectManagerList":          EditProjectManagerList
        Case "EditCommunities":                 EditCommunities
        Case "WriteJobToAccounting":            Call WriteJobToAccounting(p1)
        Case "WriteCustomerToAccounting":       Call WriteCustomerToAccounting(p1, p2)
        
        Case "TrustedLogin":                    mTrustedConnection = p1
                                                mTrustedSkipAccounting = p2 = "1"
                                                mTrustedUser = p3
        
        Case "ExportData":                      Call ExportData(p1, p2)
        Case Else
            s = "Unable to execute task " & vbQuote & TaskName & vbQuote & " - it has not been implemented in this version." & vbCrLf & vbCrLf & _
                "Supported Tasks:" & vbCrLf & _
                "    `" & vbCrLf & _
                "    WriteJobToAccounting|JobNumber" & vbCrLf & _
                "    WriteCustomerToAccounting|JobNumber|CustomerNumber"
            MsgBox s, vbCritical
    
    End Select
End Sub

Private Sub EditJCCostCodes()
    Dim s As String
            
    s = "select DivisionID,GroupDescription,CostCode, Description, DebitAccount,ABN_D04SpendCategory D04_SpendCategory,ABN_ProjectCategory ProjectCategory,ABN_ItemCode ItemCode from standardcostcodes where DivisionID = " & HFApp.DivisionID
'    Select Case HFApp.Options(AccountingSystem)
'        Case asTimberline, asIntacct, asQuickBooks, asQuickBooksOnline
'            Call FDBGrid.ShowForm("Cost Codes", s, "standardcostcodes", "divisionid,costcode", True, "DivisionID,GroupDescription")
'        Case Else
            Call FDBGrid.ShowForm("Cost Codes", _
                                  s, _
                                  "standardcostcodes", _
                                  "divisionid,costcode", _
                                  False, _
                                  "DivisionID" & IIf(HFApp.Options(AccountingSystem) = asD365_ABN, ",DebitAccount,GroupDescription", ",D04_SpendCategory,ProjectCategory,ItemCode"), _
                                  "", _
                                  True, _
                                  True)
'    End Select

End Sub

Private Sub EditJCCategories()
    Dim s As String
    
    s = "select DivisionID,Category,Description,CostType,DebitAccount,IsVariance from standardcategories where DivisionID = " & HFApp.DivisionID
    Select Case HFApp.Options(AccountingSystem)
        Case asTimberline, asIntacct, asQuickBooks, asQuickbooksOnline
            Call FDBGrid.ShowForm("Categories", s, "standardcategories", "divisionid,category", False, "DivisionID", "DivisionID,Category,Description,DebitAccount")
        Case Else
            Call FDBGrid.ShowForm("Categories", s, "standardcategories", "divisionid,category", False, "DivisionID")
    End Select
    
End Sub
Public Function Encrypt(Text As String) As String
    Dim hff As Object
    Set hff = CreateObject("ZYBFunctions.HomeFrontFunctions")
    Encrypt = hff.MyCrypt(Text, "Fazlul")
End Function



Private Sub WriteJobToSage300(JobNumber As String)
On Error GoTo eh
    Dim s As String
    Dim rs As Recordset
        
    'select data from hf tables
    s = ""
    s = s & "select Top 1" & vbCrLf
    s = s & "       s.job_mask" & vbCrLf
    s = s & "      ,j.description     Description" & vbCrLf
    s = s & "      ,j.municipal_address Address_1" & vbCrLf
    s = s & "      ,case " & vbCrLf
    s = s & "        when isnull(s.TLMunicipal_Add,0)=1 then isnull(c.customer_name,'') + isnull(' ' + c.customer_lname,'') " & vbCrLf
    s = s & "        else '' end Address_2" & vbCrLf
    s = s & "      ,c.City " & vbCrLf
    s = s & "      ,c.province State" & vbCrLf
    s = s & "      ,c.zip Zip_Code" & vbCrLf
    s = s & "      ,c.Customer_no     Customer " & vbCrLf
    s = s & "      ,case" & vbCrLf
    s = s & "        when isnull(s.TLMunicipal_Add,0)=1 and isnull(j.Description,'')<>'' then isnull(c.customer_name,'') + isnull(' ' + c.customer_lname,'') " & vbCrLf
    s = s & "        else j.Description end Customer_Name" & vbCrLf
    s = s & "      ,c.address1 customer_address_1 " & vbCrLf
    s = s & "      ,c.address2 customer_address_2" & vbCrLf
    s = s & "      ,c.city customer_city" & vbCrLf
    s = s & "      ,c.province customer_state" & vbCrLf
    s = s & "      ,c.zip customer_zip_code" & vbCrLf
    s = s & "      ,isnull(c.customer_name,'') + isnull(' ' + c.customer_lname,'') customer_contact_1" & vbCrLf
    s = s & "      ,c.phone customer_phone_1" & vbCrLf
    s = s & "      ,isnull(c.cobuyer_name,'') + isnull(' ' + c.cobuyer_lname,'') customer_contact_2" & vbCrLf
    s = s & "      ,c.c_phone customer_phone_2" & vbCrLf
    s = s & "      ,c.address1 bill_to_address_1 " & vbCrLf
    s = s & "      ,c.address2 bill_to_address_2" & vbCrLf
    s = s & "      ,c.city bill_to_city" & vbCrLf
    s = s & "      ,c.province bill_to_state" & vbCrLf
    s = s & "      ,c.zip bill_to_zip_code" & vbCrLf
    s = s & "      ,s.tljcacctformat" & vbCrLf
    s = s & "      ,s.working_Progress" & vbCrLf
    s = s & "      ,j.Gl_prefix" & vbCrLf
    s = s & "from system_setup s" & vbCrLf
    s = s & "     join tbljobs j on(1=1)" & vbCrLf
    s = s & "     left outer join tblcustomers c on(j.job_no=c.job_no and j.DivisionID = c.DivisionID AND ((c.purchased=1 AND c.cancelled=0 AND c.inactive=0 AND ((s.accounting_approve=1 AND c.approved=1) OR (s.accounting_approve=0 AND c.contract_assigned=1))) OR (c.home_selection <>'PreSale' AND c.sold_to_Customer='')))" & vbCrLf
    s = s & "where j.job_no=" & DbQuote(Str, JobNumber) & vbCrLf
    s = s & " and s.id = " & HFApp.DivisionID & vbCrLf
    s = s & " and j.DivisionID = " & HFApp.DivisionID & vbCrLf
    s = s & "order by c.customer_no desc" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    If rs.EOF Then Exit Sub
    
    JobNumber = Format(JobNumber, "" & rs("Job_Mask"))
    
    s = ""
    s = s & "insert into jcm_master__job_1(" & vbCrLf
    s = s & "  job" & vbCrLf
    s = s & " ,description" & vbCrLf
    Select Case True
        'combined
        Case Val("" & rs("tljcacctformat")) = 1 And ("" & rs("Gl_prefix") <> "") And ("" & rs("working_Progress") <> "")
            s = s & " ,cost_account" & vbCrLf
        'separate
        Case Val("" & rs("tljcacctformat")) = 2 And "" & rs("Gl_prefix") <> ""
            s = s & " ,cost_account,cost_account_prefix" & vbCrLf
    End Select
    s = s & " ,address_1, address_2, city, state, zip_code" & vbCrLf
    s = s & " ,customer, customer_name" & vbCrLf
    s = s & " ,customer_address_1,customer_address_2,customer_city,customer_state,customer_zip_code" & vbCrLf
    s = s & " ,customer_contact_1,customer_phone_1" & vbCrLf
    s = s & " ,customer_contact_2,customer_phone_2" & vbCrLf
    s = s & " ,bill_to_address_1,bill_to_address_2,bill_to_city,bill_to_state,bill_to_zip_code)" & vbCrLf & vbCrLf
    s = s & "values(" & vbCrLf
    s = s & "  " & DbQuote(Str, JobNumber) & vbCrLf
    s = s & " ," & DbQuote(Str, "" & rs("description"), , , 30) & vbCrLf
    Select Case True
        'combined
        Case Val("" & rs("tljcacctformat")) = 1 And ("" & rs("Gl_prefix") <> "") And ("" & rs("working_Progress") <> "")
            s = s & " ," & DbQuote(Str, "" & rs("Gl_prefix") & "-" & rs("working_Progress"), , , 25) & vbCrLf
        'separate
        Case Val("" & rs("tljcacctformat")) = 2 And "" & rs("Gl_prefix") <> ""
            s = s & " ," & DbQuote(Str, "" & rs("working_Progress"), , , 25)
            s = s & "," & DbQuote(Str, "" & rs("Gl_prefix"), , , 22) & vbCrLf
    End Select
    
    s = s & " ," & DbQuote(Str, "" & rs("Address_1"), , , 30)
    s = s & "," & DbQuote(Str, "" & rs("Address_2"), , , 30)
    s = s & "," & DbQuote(Str, "" & rs("city"), , , 15)
    s = s & "," & DbQuote(Str, "" & rs("state"), , , 4)
    s = s & "," & DbQuote(Str, "" & rs("zip_code"), , , 10) & vbCrLf
    s = s & " ," & DbQuote(Str, "" & rs("customer"), , , 10)
    s = s & "," & DbQuote(Str, "" & rs("customer_name"), , , 30) & vbCrLf
    s = s & " ," & DbQuote(Str, "" & rs("customer_address_1"), , , 30)
    s = s & "," & DbQuote(Str, "" & rs("customer_address_2"), , , 30)
    s = s & "," & DbQuote(Str, "" & rs("customer_city"), , , 15)
    s = s & "," & DbQuote(Str, "" & rs("customer_state"), , , 4)
    s = s & "," & DbQuote(Str, "" & rs("customer_zip_code"), , , 10) & vbCrLf
    s = s & " ," & DbQuote(Str, "" & rs("customer_contact_1"), , , 30)
    s = s & "," & DbQuote(Str, "" & rs("customer_phone_1"), , , 15) & vbCrLf
    s = s & " ," & DbQuote(Str, "" & rs("customer_contact_2"), , , 30)
    s = s & "," & DbQuote(Str, "" & rs("customer_phone_2"), , , 15) & vbCrLf
    s = s & " ," & DbQuote(Str, "" & rs("bill_to_address_1"), , , 30)
    s = s & "," & DbQuote(Str, "" & rs("bill_to_address_2"), , , 30)
    s = s & "," & DbQuote(Str, "" & rs("bill_to_city"), , , 15)
    s = s & "," & DbQuote(Str, "" & rs("bill_to_state"), , , 4)
    s = s & "," & DbQuote(Str, "" & rs("bill_to_zip_code"), , , 10)
    s = s & ")"
    HFApp.SqlExec s, dbAccounting
    

Exit Sub
eh: If InStr(1, Err.Description, "duplicate") > 0 Then
        Exit Sub
    Else
        Call errHandler(SRCFILE & "WriteJobToSage300", s)
    End If
End Sub

Private Sub WriteJobToSage100(JobNumber As String)
On Error GoTo eh
    Dim s As String
    Dim action As String
    Dim i As Integer
    Dim rs As Recordset
    Dim rs1 As Recordset
    Dim JStatus As String
    JStatus = "" & HFApp.Options.ValueByName("Sage100SoldJobStatus")
    If JStatus = "" Then JStatus = "4"
    
    Dim DBUG As Boolean
    DBUG = "" & HFApp.Options.ValueByName("UseJobAsSubAcctAUDITLOG") = "true"
    
    'select data from hf tables
    s = ""
    s = s & "select TOP 1" & vbCrLf
    s = s & "       j.job_no      JobNo" & vbCrLf
    s = s & "      ,isnull(nullif(j.Description,''),j.Job_no) Name" & vbCrLf
    s = s & "      ,j.job_no      ShortName" & vbCrLf
    s = s & "      ,substring(j.municipal_address,1,30) Addr1" & vbCrLf
    s = s & "      ,substring(j.municipal_address,31,60) Addr2" & vbCrLf
    s = s & "      ,j.city " & vbCrLf
    s = s & "      ,j.province    State" & vbCrLf
    s = s & "      ,j.zip         PostalCode" & vbCrLf
    s = s & "      ,j.SitePhone   Phone" & vbCrLf
    s = s & "      ,j.SiteFax     Fax" & vbCrLf
    s = s & "      ,m.description       Model" & vbCrLf
    s = s & "      ,c.total_sales_price" & vbCrLf
    s = s & "      ,c.square_footage" & vbCrLf
    s = s & "      ,c.lot" & vbCrLf
    s = s & "      ,case j.inactive when 1 then 6 else " & JStatus & " end JobStatus" & vbCrLf
    s = s & "      ,c.Customer_No" & vbCrLf
    s = s & "      ,s.Working_Progress WIPAccount" & vbCrLf
    s = s & "from system_setup s" & vbCrLf
    s = s & "     join tbljobs j on(s.ID = j.DivisionID)" & vbCrLf
    s = s & "     left outer join tblcustomers c on(j.job_no=c.job_no and j.DivisionId = c.DivisionID AND ((c.purchased=1 AND c.cancelled=0 AND c.inactive=0 AND ((s.accounting_approve=1 AND c.approved=1) OR (s.accounting_approve=0 AND c.contract_assigned=1))) OR (c.home_selection <>'PreSale' AND c.sold_to_Customer='')))" & vbCrLf
    s = s & "     left outer join distinctmodelsbydivision m on isnull(nullif(c.Model,''),j.model)=m.model and j.divisionid=m.divisionid" & vbCrLf
    s = s & "where j.job_no=" & DbQuote(Str, JobNumber) & vbCrLf
    s = s & " and j.DivisionID = " & HFApp.DivisionID & vbCrLf
    s = s & "order by c.customer_no desc" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    If rs.EOF Then Exit Sub
    
    'if job not in MB then create it
    Set rs1 = HFApp.SqlExec("select jobnme from actrec where recnum=" & DbQuote(Num, JobNumber), dbAccounting)
    If rs1.EOF Then
        action = "Add"
        s = StartMBXml(HFApp.Options(MasterBuilderCompany), HFApp.Options(MasterBuilderUID))
        s = s & "<Job" & action & "Rq requestID=""1"">" & vbCrLf
        s = s & AddMBXml(c, 75, "ObjectRef", "" & rs("JobNo"))
        s = s & AddMBXml(c, 30, "Name", "" & rs("Name"))
        s = s & AddMBXml(c, 30, "ShortName", "" & rs("ShortName"))
        s = s & AddMBXml(n, 10, "ClientRef", "" & rs("Customer_No"))
        s = s & AddMBXml(c, 30, "Addr1", "" & rs("Addr1"))
        s = s & AddMBXml(c, 30, "Addr2", "" & rs("Addr2"))
        s = s & AddMBXml(c, 25, "City", "" & rs("City"))
        
        If HFApp.Options.ValueByName("MasterBuilderEdition") = "CA" Then
            s = s & AddMBXml(c, 2, "Province", "" & rs("State"))
        Else
            s = s & AddMBXml(c, 2, "State", "" & rs("State"))
        End If
        
        s = s & AddMBXml(c, 10, "PostalCode", "" & rs("PostalCode"))
        
        'not supported in CA
        If HFApp.Options.ValueByName("MasterBuilderEdition") <> "CA" Then
            s = s & AddMBXml(p, 14, "Phone", "" & rs("Phone"))
            s = s & AddMBXml(p, 14, "Fax", "" & rs("Fax"))
        End If
        
        If HFApp.Options.ValueByName("Sage100ModelAsJobType") = "True" Then
            s = s & AddMBXml(n, 10, "JobTypeRef", "" & rs("Model"))
        End If
        s = s & AddMBXml(n, 2, "JobStatus", "" & rs("JobStatus"))
        
        'CA only supported in v20+
        If HFApp.Options.ValueByName("MasterBuilderEdition") = "CA" And Not IsIn(HFApp.Options.ValueByName("Sage100APILevel"), "v18", "v19") Then
            s = s & AddMBXml(n, 0, "ActualSalesPriceAmount", "" & rs("total_sales_price"))
            s = s & AddMBXml(c, 10, "LotNumber", "" & rs("lot"))
            s = s & AddMBXml(c, 15, "ModelName", "" & rs("model"))
            s = s & AddMBXml(n, 0, "SquareFeet", "" & rs("square_footage"))
        End If
        
        s = s & "</Job" & action & "Rq>" & vbCrLf
        s = s & EndMBXml()
        Call SubmitMBXml(s, HFApp.Options(MasterBuilderPWD))
    End If
    
    
    'Now create subledger
    If HFApp.Options.ValueByName("UseJobAsSubAcct") = "True" Then
        If DbQuote(Num, "" & rs("WIPAccount")) = 0 Then
            MsgBox "Missing WIP Account in System Settings. This is a required field." & vbCrLf & "Sage 100 may not import POs correctly." & vbCrLf & "Correct, and try again if an error in posting occurs", vbCritical, "Missing Information"
            Exit Sub
        End If
        i = 0
        i = HFApp.SqlExec("Select count(*) from lgrsub where recnum = " & DbQuote(Num, "" & rs("JobNo")) & " and ctract=" & "" & DbQuote(Num, "" & rs("WIPAccount")), dbAccounting)(0)
        If i = 0 Then
            'so now we do this also
            If Val(Mid(HFApp.Options.ValueByName("Sage100APILevel"), 2)) < 20 Then
                s = ""
                s = s & "insert into lgrsub(recnum, shtnme, lngnme, ctract, begbal, endbal, strbal, ntetxt)" & vbCrLf
                s = s & "Values(" & DbQuote(Num, "" & rs("JobNo")) & vbCrLf
                s = s & "      ," & DbQuote(Str, "" & rs("ShortName")) & vbCrLf
                s = s & "      ," & DbQuote(Str, "" & rs("Name")) & vbCrLf
                s = s & "      ," & DbQuote(Num, "" & rs("WIPAccount")) & vbCrLf
                s = s & "      ,0,0,0,'')"
            Else
                'v20 and up is mssql which added the _idnum column
                s = ""
                s = s & "insert into lgrsub(_idnum, recnum, shtnme, lngnme, ctract, begbal, endbal, strbal, ntetxt)" & vbCrLf
                s = s & "Values(newid()" & vbCrLf
                s = s & "      ," & DbQuote(Num, "" & rs("JobNo")) & vbCrLf
                s = s & "      ," & DbQuote(Str, "" & rs("ShortName")) & vbCrLf
                s = s & "      ," & DbQuote(Str, "" & rs("Name")) & vbCrLf
                s = s & "      ," & DbQuote(Num, "" & rs("WIPAccount")) & vbCrLf
                s = s & "      ,0,0,0,'')"
            End If
            On Error Resume Next
            Call HFApp.SqlExec(s, dbAccounting)
            'If Err.Number <> 0 And DBUG Then Call HFApp.WriteAuditLog("Create lgrsub", Err.Description)
            On Error GoTo eh
        
        End If
    End If
    
Exit Sub
eh: If InStr(1, Err.Description, "duplicate") > 0 Then
        Resume Next
    Else
        Call errHandler(SRCFILE & "WriteJobToSage100", s)
    End If
End Sub




Private Sub WriteCustomerToQuickBooks(JobNumber As String, CustomerNumber As String)
    If "" & HFApp.Options(MultiFamily) = "True" Then
        Call WriteCustomerToQuickBooksMF(JobNumber, CustomerNumber)
    Else
        Call WriteCustomerToQuickBooksSF(JobNumber, CustomerNumber)
    End If
End Sub


Private Sub WriteCustomerToQuickBooksMF(JobNumber As String, CustomerNumber As String)
On Error GoTo eh
    Dim i As Long
    Dim s As String
    Dim rs As Recordset

    Dim listid     As String

    'first write jobs
    s = ""
    s = s & "SELECT DISTINCT" & vbCrLf
    s = s & "       j.job_no" & vbCrLf
    s = s & "      ,j.description" & vbCrLf
    s = s & "  FROM tblCustomers c join tbljobs j on(c.job_no=j.job_no and c.DivisionID = j.DivisionID)" & vbCrLf
    s = s & "      ,system_setup s" & vbCrLf
    s = s & " WHERE isnull(c.Job_no,'')<>'' and((c.purchased=1  AND isnull(c.cancelled,0)=0  AND isnull(c.inactive,0)=0 AND ((s.accounting_approve=1 AND c.approved=1) OR (s.accounting_approve=0 AND c.contract_assigned=1))) OR (c.home_selection <>'PreSale' AND isnull(c.sold_to_Customer,'')=''))" & vbCrLf
    s = s & "   and isnull(j.ExternalJobID,'')=''" & vbCrLf
    If CustomerNumber = "" Then
        s = s & "   AND c.job_no=" & DbQuote(Str, JobNumber) & vbCrLf
    Else
        s = s & "   AND c.Customer_no=" & DbQuote(Str, CustomerNumber) & vbCrLf
    End If
    s = s & " and s.id = " & HFApp.DivisionID & " and c.DivisionID = " & HFApp.DivisionID
    Set rs = HFApp.SqlExec(s)
    While Not rs.EOF

        s = StartQBXml()
        s = s & "<CustomerAddRq requestID=""1"">" & vbCrLf
        s = s & "<CustomerAdd>" & vbCrLf
        s = s & AddQBXml(c, 41, "Name", "" & rs("job_no") & " - " & rs("description"))
        s = s & "</CustomerAdd>" & vbCrLf
        s = s & "</CustomerAddRq>" & vbCrLf
        s = s & EndQBXml()

        s = SubmitQBXml(s)
        listid = Parse(s, 2, "ListID>")
        listid = left(listid, Len(listid) - 2)

        s = ""
        s = s & "update tbljobs" & vbCrLf
        s = s & "   set ExternalJobID=" & DbQuote(Str, listid) & vbCrLf
        s = s & " where job_no=" & DbQuote(Str, "" & rs("Job_No")) & vbCrLf
        Call HFApp.SqlExec(s, dbHomefront)
        
        rs.MoveNext
    Wend


    'now write customers
    s = ""
    s = s & "SELECT c.customer_no" & vbCrLf
    s = s & "      ,c.description" & vbCrLf
    s = s & "      ,j.ExternalJobID" & vbCrLf
    s = s & "      ,c.Buyer_salutation Salutation" & vbCrLf
    s = s & "      ,c.Customer_Name    FirstName" & vbCrLf
    s = s & "      ,c.Customer_LName   LastName" & vbCrLf
    s = s & "      ,c.Address1         BillAddr1" & vbCrLf
    s = s & "      ,c.Address2         BillAddr2" & vbCrLf
    s = s & "      ,c.City             BillCity" & vbCrLf
    s = s & "      ,c.Province         BillState" & vbCrLf
    s = s & "      ,c.zip              BillPostalCode" & vbCrLf
    s = s & "      ,c.Municipal_Address    ShipAddr1" & vbCrLf
    s = s & "      ,c.Lot_City             ShipCity" & vbCrLf
    s = s & "      ,c.Lot_Province         ShipState" & vbCrLf
    s = s & "      ,c.Lot_zip              ShipPostalCode" & vbCrLf
    s = s & "      ,c.Phone                Phone" & vbCrLf
    s = s & "      ,c.c_phone              AltPhone" & vbCrLf
    s = s & "      ,c.fax                  Fax" & vbCrLf
    s = s & "      ,c.email                Email" & vbCrLf
    s = s & "      ,ISNULL(c.Customer_Name,'') + ISNULL(' '+c.Customer_LName,'') Contact" & vbCrLf
    s = s & "      ,ISNULL(c.CoBuyer_Name,'') + ISNULL(' '+c.CoBuyer_LName,'')   AltContact" & vbCrLf
    s = s & "      ,c.Comments             Notes" & vbCrLf
    s = s & "  FROM tblCustomers c join tbljobs j on(c.job_no=j.job_no and c.DivisionID = j.DivisionID)" & vbCrLf
    s = s & "      ,system_setup s" & vbCrLf
    s = s & " WHERE isnull(c.Job_no,'')<>'' and((c.purchased=1  AND c.cancelled=0  AND c.inactive=0 AND ((s.accounting_approve=1 AND c.approved=1) OR (s.accounting_approve=0 AND c.contract_assigned=1))) OR (c.home_selection <>'PreSale' AND c.sold_to_Customer=''))" & vbCrLf
    s = s & "   and isnull(c.ar_customer_deposit,'')=''" & vbCrLf
    If CustomerNumber = "" Then
        s = s & "   AND c.job_no=" & DbQuote(Str, JobNumber) & vbCrLf
    Else
        s = s & "   AND c.Customer_no=" & DbQuote(Str, CustomerNumber) & vbCrLf
    End If
    s = s & " and s.id = " & HFApp.DivisionID & " and c.DivisionID = " & HFApp.DivisionID
    Set rs = HFApp.SqlExec(s)
    While Not rs.EOF

        s = StartQBXml()
        s = s & "<CustomerAddRq requestID=""1"">" & vbCrLf
        s = s & "<CustomerAdd>" & vbCrLf
        s = s & AddQBXml(c, 41, "Name", "" & rs("customer_no") & " - " & rs("description"))
        s = s & "<ParentRef>" & vbCrLf
        s = s & AddQBXml(c, 41, "ListID", "" & rs("ExternalJobID"))
        s = s & "</ParentRef>" & vbCrLf
        s = s & AddQBXml(c, 15, "Salutation", "" & rs("Salutation"))
        s = s & AddQBXml(c, 25, "FirstName", "" & rs("FirstName"))
        s = s & AddQBXml(c, 25, "LastName", "" & rs("LastName"))
        s = s & "<BillAddress>" & vbCrLf
        s = s & AddQBXml(c, 41, "Addr1", "" & rs("BillAddr1"))
        s = s & AddQBXml(c, 41, "Addr2", "" & rs("BillAddr2"))
        s = s & AddQBXml(c, 31, "City", "" & rs("BillCity"))
        s = s & AddQBXml(c, 21, "State", "" & rs("BillState"))
        s = s & AddQBXml(c, 13, "PostalCode", "" & rs("BillPostalCode"))
        s = s & "</BillAddress>" & vbCrLf
        s = s & "<ShipAddress>" & vbCrLf
        s = s & AddQBXml(c, 41, "Addr1", "" & rs("ShipAddr1"))
        s = s & AddQBXml(c, 31, "City", "" & rs("ShipCity"))
        s = s & AddQBXml(c, 21, "State", "" & rs("ShipState"))
        s = s & AddQBXml(c, 13, "PostalCode", "" & rs("ShipPostalCode"))
        s = s & "</ShipAddress>" & vbCrLf
        s = s & AddQBXml(c, 21, "Phone", "" & rs("Phone"))
        s = s & AddQBXml(c, 21, "AltPhone", "" & rs("AltPhone"))
        s = s & AddQBXml(c, 21, "Fax", "" & rs("Fax"))
        s = s & AddQBXml(c, 1023, "Email", "" & rs("Email"))
        s = s & AddQBXml(c, 41, "Contact", "" & rs("Contact"))
        s = s & AddQBXml(c, 41, "AltContact", "" & rs("AltContact"))
        s = s & AddQBXml(c, 4095, "Notes", "" & rs("Notes"))
        s = s & "</CustomerAdd>" & vbCrLf
        s = s & "</CustomerAddRq>" & vbCrLf
        s = s & EndQBXml()

        s = SubmitQBXml(s)
        listid = Parse(s, 2, "ListID>")
        listid = left(listid, Len(listid) - 2)

        s = ""
        s = s & "update tblcustomers" & vbCrLf
        s = s & "   set ar_customer_deposit=" & DbQuote(Str, listid) & vbCrLf
        s = s & " where customer_no=" & DbQuote(Str, "" & rs("customer_no")) & vbCrLf
        Call HFApp.SqlExec(s, dbHomefront)

        s = ""
        s = s & "insert into arcustomers(arcustomer,description,contact1,phone1,fax1,email1,contact2,phone2,billaddr1,billaddr2,billcity,billprovince,billpostalcode,shipaddr1,shipcity,shipprovince,shippostalcode,comments)" & vbCrLf
        s = s & "values(" & DbQuote(Str, listid) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("customer_no") & " - " & rs("description")) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("contact")) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("phone")) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("fax")) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("email")) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("altcontact")) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("altphone")) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("billaddr1")) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("billaddr2")) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("billcity")) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("billstate")) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("billpostalcode")) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("shipaddr1")) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("shipcity")) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("shipstate")) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("shippostalcode")) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("notes")) & ")"
        Call HFApp.SqlExec(s, dbHomefront)
        
        rs.MoveNext
    Wend


Exit Sub
eh: If InStr(1, Err.Description, "duplicate") > 0 Then
        Resume Next
    Else
        Call errHandler("WriteCustomerToQuickBooksMF", s)
    End If
End Sub
Private Sub WriteCustomerToQuickBooksSF(JobNumber As String, CustomerNumber As String)
On Error GoTo eh
    Dim i As Long
    Dim s As String
    Dim rs As Recordset
    Dim action As String

    Dim listid     As String

    'first write customers
    s = ""
    s = s & "SELECT j.job_no" & vbCrLf
    s = s & "      ,j.description JobDesc" & vbCrLf
    s = s & "      ,c.customer_no" & vbCrLf
    s = s & "      ,c.description CustDesc" & vbCrLf
    s = s & "      ,c.ar_customer_deposit" & vbCrLf
    s = s & "      ,c.Buyer_salutation Salutation" & vbCrLf
    s = s & "      ,c.Customer_Name    FirstName" & vbCrLf
    s = s & "      ,c.Customer_LName   LastName" & vbCrLf
    s = s & "      ,c.Address1         BillAddr1" & vbCrLf
    s = s & "      ,c.Address2         BillAddr2" & vbCrLf
    s = s & "      ,c.City             BillCity" & vbCrLf
    s = s & "      ,c.Province         BillState" & vbCrLf
    s = s & "      ,c.zip              BillPostalCode" & vbCrLf
    s = s & "      ,c.Municipal_Address    ShipAddr1" & vbCrLf
    s = s & "      ,c.Lot_City             ShipCity" & vbCrLf
    s = s & "      ,c.Lot_Province         ShipState" & vbCrLf
    s = s & "      ,c.Lot_zip              ShipPostalCode" & vbCrLf
    s = s & "      ,c.Phone                Phone" & vbCrLf
    s = s & "      ,c.c_phone              AltPhone" & vbCrLf
    s = s & "      ,c.fax                  Fax" & vbCrLf
    s = s & "      ,c.email                Email" & vbCrLf
    s = s & "      ,ISNULL(c.Customer_Name,'') + ISNULL(' '+c.Customer_LName,'') Contact" & vbCrLf
    s = s & "      ,ISNULL(c.CoBuyer_Name,'') + ISNULL(' '+c.CoBuyer_LName,'')   AltContact" & vbCrLf
    s = s & "      ,j.Start_date           JobStartDate" & vbCrLf
    s = s & "      ,j.Complete_date        JobEndDate" & vbCrLf
    s = s & "      ,c.Comments             Notes" & vbCrLf
    s = s & "      ,isnull(j.ExternalJobID,'') ExternalJobID" & vbCrLf
    s = s & "  FROM tblCustomers c join tbljobs j on(c.job_no=j.job_no and c.DivisionID = j.DivisionID)" & vbCrLf
    s = s & "      ,system_setup s" & vbCrLf
    s = s & " WHERE isnull(c.Job_no,'')<>'' and((c.purchased=1  AND isnull(c.cancelled,0)=0  AND isnull(c.inactive,0)=0 AND (c.approved=1 OR c.contract_assigned=1))" & vbCrLf
    s = s & "       OR (c.home_selection <>'PreSale' AND isnull(c.sold_to_Customer,'')=''))" & vbCrLf
    If CustomerNumber = "" Then
        s = s & "   AND c.job_no=" & DbQuote(Str, JobNumber) & vbCrLf
    Else
        s = s & "   AND c.Customer_no=" & DbQuote(Str, CustomerNumber) & vbCrLf
    End If
    s = s & " and s.id = " & HFApp.DivisionID & " and c.DivisionID = " & HFApp.DivisionID
    Set rs = HFApp.SqlExec(s)
    While Not rs.EOF
    
        action = IIf("" & rs("ar_customer_deposit") = "", "Add", "Mod")

        'add customer to quickbooks
        s = StartQBXml()
        s = s & "<Customer" & action & "Rq requestID=""1"">" & vbCrLf
        s = s & "<Customer" & action & ">" & vbCrLf
        If action = "Mod" Then
            s = s & AddQBXml(c, 41, "ListID", "" & rs("ar_customer_deposit"))
            s = s & AddQBXml(c, 41, "EditSequence", getCustEditSeq("" & rs("ar_customer_deposit")))
        End If
        s = s & AddQBXml(c, 41, "Name", "" & rs("customer_no") & " - " & rs("custdesc"))
        s = s & AddQBXml(c, 15, "Salutation", "" & rs("Salutation"))
        s = s & AddQBXml(c, 25, "FirstName", "" & rs("FirstName"))
        s = s & AddQBXml(c, 25, "LastName", "" & rs("LastName"))
        s = s & "<BillAddress>" & vbCrLf
        s = s & AddQBXml(c, 41, "Addr1", "" & rs("BillAddr1"))
        s = s & AddQBXml(c, 41, "Addr2", "" & rs("BillAddr2"))
        s = s & AddQBXml(c, 31, "City", "" & rs("BillCity"))
        s = s & AddQBXml(c, 21, "State", "" & rs("BillState"))
        s = s & AddQBXml(c, 13, "PostalCode", "" & rs("BillPostalCode"))
        s = s & "</BillAddress>" & vbCrLf
        s = s & "<ShipAddress>" & vbCrLf
        s = s & AddQBXml(c, 41, "Addr1", "" & rs("ShipAddr1"))
        s = s & AddQBXml(c, 31, "City", "" & rs("ShipCity"))
        s = s & AddQBXml(c, 21, "State", "" & rs("ShipState"))
        s = s & AddQBXml(c, 13, "PostalCode", "" & rs("ShipPostalCode"))
        s = s & "</ShipAddress>" & vbCrLf
        s = s & AddQBXml(c, 21, "Phone", "" & rs("Phone"))
        s = s & AddQBXml(c, 21, "AltPhone", "" & rs("AltPhone"))
        s = s & AddQBXml(c, 21, "Fax", "" & rs("Fax"))
        s = s & AddQBXml(c, 1023, "Email", "" & rs("Email"))
        s = s & AddQBXml(c, 41, "Contact", "" & rs("Contact"))
        s = s & AddQBXml(c, 41, "AltContact", "" & rs("AltContact"))
        s = s & AddQBXml(c, 4095, "Notes", "" & rs("Notes"))
        s = s & "</Customer" & action & ">" & vbCrLf
        s = s & "</Customer" & action & "Rq>" & vbCrLf
        s = s & EndQBXml()
        If action = "Add" Then
            s = SubmitQBXml(s)
            listid = Parse(s, 2, "ListID>")
        Else
            listid = "" & rs("ar_Customer_Deposit")
        End If
        If listid <> "" Then
            If Right(listid, 2) = "</" Then listid = left(listid, Len(listid) - 2)
            
            'set external id in HF
            s = ""
            s = s & "update tblcustomers" & vbCrLf
            s = s & "   set ar_customer_deposit=" & DbQuote(Str, listid) & vbCrLf
            s = s & " where customer_no=" & DbQuote(Str, "" & rs("customer_no")) & vbCrLf
            Call HFApp.SqlExec(s, dbHomefront)
            
            'add customer to HF
            s = ""
            s = s & "insert into arcustomers(arcustomer,description,contact1,phone1,fax1,email1,contact2,phone2,billaddr1,billaddr2,billcity,billprovince,billpostalcode,shipaddr1,shipcity,shipprovince,shippostalcode,comments)" & vbCrLf
            s = s & "values(" & DbQuote(Str, listid) & vbCrLf
            s = s & "      ," & DbQuote(Str, "" & rs("customer_no") & " - " & rs("custdesc")) & vbCrLf
            s = s & "      ," & DbQuote(Str, "" & rs("contact")) & vbCrLf
            s = s & "      ," & DbQuote(Str, "" & rs("phone")) & vbCrLf
            s = s & "      ," & DbQuote(Str, "" & rs("fax")) & vbCrLf
            s = s & "      ," & DbQuote(Str, "" & rs("email")) & vbCrLf
            s = s & "      ," & DbQuote(Str, "" & rs("altcontact")) & vbCrLf
            s = s & "      ," & DbQuote(Str, "" & rs("altphone")) & vbCrLf
            s = s & "      ," & DbQuote(Str, "" & rs("billaddr1")) & vbCrLf
            s = s & "      ," & DbQuote(Str, "" & rs("billaddr2")) & vbCrLf
            s = s & "      ," & DbQuote(Str, "" & rs("billcity")) & vbCrLf
            s = s & "      ," & DbQuote(Str, "" & rs("billstate")) & vbCrLf
            s = s & "      ," & DbQuote(Str, "" & rs("billpostalcode")) & vbCrLf
            s = s & "      ," & DbQuote(Str, "" & rs("shipaddr1")) & vbCrLf
            s = s & "      ," & DbQuote(Str, "" & rs("shipcity")) & vbCrLf
            s = s & "      ," & DbQuote(Str, "" & rs("shipstate")) & vbCrLf
            s = s & "      ," & DbQuote(Str, "" & rs("shippostalcode")) & vbCrLf
            s = s & "      ," & DbQuote(Str, "" & rs("notes")) & ")"
            Call HFApp.SqlExec(s, dbHomefront)
            
            
            'add job to quickbooks
            If HFApp.Options.ValueByName("QuickBooksJobStyle") <> "Simple" Then
                If "" & rs("ExternalJobID") = "" Then
                    s = StartQBXml()
                    s = s & "<CustomerAddRq requestID=""1"">" & vbCrLf
                    s = s & "<CustomerAdd>" & vbCrLf
                    s = s & AddQBXml(c, 41, "Name", "" & rs("job_no") & " - " & rs("jobdesc"))
                    s = s & "<ParentRef>" & vbCrLf
                    s = s & AddQBXml(c, 41, "ListID", listid)
                    s = s & "</ParentRef>" & vbCrLf
                    s = s & AddQBXml(c, 15, "Salutation", "" & rs("Salutation"))
                    s = s & AddQBXml(c, 25, "FirstName", "" & rs("FirstName"))
                    s = s & AddQBXml(c, 25, "LastName", "" & rs("LastName"))
                    s = s & "<BillAddress>" & vbCrLf
                    s = s & AddQBXml(c, 41, "Addr1", "" & rs("BillAddr1"))
                    s = s & AddQBXml(c, 41, "Addr2", "" & rs("BillAddr2"))
                    s = s & AddQBXml(c, 31, "City", "" & rs("BillCity"))
                    s = s & AddQBXml(c, 21, "State", "" & rs("BillState"))
                    s = s & AddQBXml(c, 13, "PostalCode", "" & rs("BillPostalCode"))
                    s = s & "</BillAddress>" & vbCrLf
                    s = s & "<ShipAddress>" & vbCrLf
                    s = s & AddQBXml(c, 41, "Addr1", "" & rs("ShipAddr1"))
                    s = s & AddQBXml(c, 31, "City", "" & rs("ShipCity"))
                    s = s & AddQBXml(c, 21, "State", "" & rs("ShipState"))
                    s = s & AddQBXml(c, 13, "PostalCode", "" & rs("ShipPostalCode"))
                    s = s & "</ShipAddress>" & vbCrLf
                    s = s & AddQBXml(c, 21, "Phone", "" & rs("Phone"))
                    s = s & AddQBXml(c, 21, "AltPhone", "" & rs("AltPhone"))
                    s = s & AddQBXml(c, 21, "Fax", "" & rs("Fax"))
                    s = s & AddQBXml(c, 1023, "Email", "" & rs("Email"))
                    s = s & AddQBXml(c, 41, "Contact", "" & rs("Contact"))
                    s = s & AddQBXml(c, 41, "AltContact", "" & rs("AltContact"))
                    s = s & AddQBXml(d, 41, "JobStartDate", "" & rs("JobStartDate"))
                    s = s & AddQBXml(d, 41, "JobEndDate", "" & rs("JobEndDate"))
                    s = s & AddQBXml(c, 99, "JobDesc", "" & rs("jobDesc"))
                    s = s & AddQBXml(c, 4095, "Notes", "" & rs("Notes"))
                    s = s & "</CustomerAdd>" & vbCrLf
                    s = s & "</CustomerAddRq>" & vbCrLf
                    s = s & EndQBXml()
                    s = SubmitQBXml(s)
                    listid = Parse(s, 2, "ListID>")
                    If listid <> "" Then listid = left(listid, Len(listid) - 2)
                    
                    'set externalid in hf
                    If listid <> "" And "" & rs("ExternalJobID") = "" Then
                        s = ""
                        s = s & "update tbljobs" & vbCrLf
                        s = s & "   set ExternalJobID=" & DbQuote(Str, listid) & vbCrLf
                        s = s & " where job_no=" & DbQuote(Str, "" & rs("Job_No")) & vbCrLf
                        Call HFApp.SqlExec(s, dbHomefront)
                    End If
                End If
            Else
                If listid <> "" Then
                    Call HFApp.SqlExec("update tbljobs set ExternalJobID = " & DbQuote(Str, listid) & " where job_no=" & DbQuote(Str, rs("Job_no")), dbHomefront)
                End If
            End If
        End If
        rs.MoveNext
    Wend





Exit Sub
eh:
If InStr(1, Err.Description, "already in use") > 0 Or InStr(1, Err.Description, "duplicate") > 0 Then
        Resume Next
    Else

        Call errHandler("WriteCustomerToQuickBooksSF", s)
        
    End If
End Sub

Private Function getCustEditSeq(listid As String) As String
    Dim s As String
    Dim editseq As String
    
    s = StartQBXml()
    's = s & "<CustomerQuery>" & vbCrLf
    s = s & "<CustomerQueryRq requestID=""1"">" & vbCrLf
    's = s & "<CustomerQuery>" & vbCrLf
    s = s & AddQBXml(c, 41, "ListID", listid)
   '  s = s & "</CustomerQuery>" & vbCrLf
    s = s & "</CustomerQueryRq>" & vbCrLf
   ' s = s & "</CustomerQuery>" & vbCrLf
    s = s & EndQBXml()
    s = SubmitQBXml(s)
    
    editseq = Parse(s, 2, "EditSequence>")
    editseq = left(editseq, Len(editseq) - 2)
    getCustEditSeq = editseq
    
End Function

Private Sub WriteCustomerToSage100(JobNumber As String, CustomerNumber As String)
On Error GoTo eh
    Dim i As Long
    Dim s As String
    Dim Status As Long
    Dim rs As Recordset
    Dim rs1 As Recordset
    
    
    'select data from hf tables
    s = ""
    s = s & "SELECT c.customer_no" & vbCrLf
    s = s & "      ,isnull(nullif(c.Description,''),c.customer_no) Description" & vbCrLf
    s = s & "      ,c.Address1" & vbCrLf
    s = s & "      ,c.Address2" & vbCrLf
    s = s & "      ,c.City" & vbCrLf
    s = s & "      ,c.Province" & vbCrLf
    s = s & "      ,c.Zip" & vbCrLf
    s = s & "      ,ISNULL(c.Customer_Name,'') + ISNULL(' '+c.Customer_LName,'') Contact1" & vbCrLf
    s = s & "      ,c.Phone" & vbCrLf
    s = s & "      ,c.fax" & vbCrLf
    s = s & "      ,c.email" & vbCrLf
    s = s & "      ,c.cellphone" & vbCrLf
    s = s & "      ,ISNULL(c.CoBuyer_Name,'') + ISNULL(' '+c.CoBuyer_LName,'') Contact2" & vbCrLf
    s = s & "      ,c.C_Phone" & vbCrLf
    s = s & "      ,c.C_fax" & vbCrLf
    s = s & "      ,c.C_email" & vbCrLf
    s = s & "      ,c.C_cellphone" & vbCrLf
    s = s & "      ,Purchased_Date" & vbCrLf
    s = s & "      ,c.Comments" & vbCrLf
    s = s & "  FROM tblCustomers c ,system_setup s" & vbCrLf
    s = s & " WHERE isnull(c.Job_no,'')<>'' and((c.purchased=1  AND c.cancelled=0  AND c.inactive=0 AND ((s.accounting_approve=1 AND c.approved=1)" & vbCrLf
    s = s & "    OR  (s.accounting_approve=0 AND c.contract_assigned=1))) OR (c.home_selection <>'PreSale' AND c.sold_to_Customer=''))" & vbCrLf
    If CustomerNumber = "" Then
        s = s & "   AND job_no=" & DbQuote(Str, JobNumber) & vbCrLf
    Else
        s = s & "   AND Customer_no=" & DbQuote(Str, CustomerNumber) & vbCrLf
    End If
    s = s & " and s.id = " & HFApp.DivisionID & " and c.DivisionID = " & HFApp.DivisionID
    Set rs = HFApp.SqlExec(s)
    If Not rs.EOF Then
        
        'status is required but we don't have that. use first in list and only set on add
        On Error Resume Next
        Status = 1
        Status = HFApp.SqlExec("select min(recnum) from clnsts", dbAccounting)(0)
        On Error GoTo eh
        
        i = 0
        s = StartMBXml(HFApp.Options(MasterBuilderCompany), HFApp.Options(MasterBuilderUID))
        While Not rs.EOF
            'check if exists in MB
            Set rs1 = HFApp.SqlExec("select recnum from reccln where recnum=" & DbQuote(Num, CustomerNumber), dbAccounting)
            If rs1.EOF Then
                i = i + 1
                s = s & "<ClientAddRq requestID=""" & i & """>" & vbCrLf
                s = s & AddMBXml(n, 10, "ObjectRef", "" & rs("Customer_No"))
                s = s & AddMBXml(c, 15, "ShortName", "" & rs("Description"))
                s = s & AddMBXml(c, 30, "Name", "" & rs("Description"))
                s = s & AddMBXml(c, 30, "Addr1", "" & rs("Address1"))
                s = s & AddMBXml(c, 30, "Addr2", "" & rs("Address2"))
                s = s & AddMBXml(c, 25, "City", "" & rs("City"))
                If HFApp.Options.ValueByName("MasterBuilderEdition") = "CA" Then
                    s = s & AddMBXml(st, 2, "Province", "" & rs("Province"))
                Else
                    s = s & AddMBXml(st, 2, "State", "" & rs("Province"))
                End If
                s = s & AddMBXml(z, 10, "PostalCode", "" & rs("Zip"))
                If "" & rs("Purchased_Date") <> "" Then
                    s = s & AddMBXml(d, 10, "PurchasedDate", "" & rs("Purchased_Date"))
                End If
                s = s & AddMBXml(n, 2, "ClientStatusRef", Status)
                s = s & AddMBXml(m, 10, "Memo", "" & rs("Comments"))
                s = s & "<ClientContactAdd>" & vbCrLf
                    s = s & AddMBXml(c, 30, "ContactName", "" & rs("Contact1"))
                    s = s & AddMBXml(c, 20, "JobTitle", "Buyer")
                    s = s & AddMBXml(p, 14, "Phone", "" & rs("Phone"))
                    s = s & AddMBXml(c, 30, "Email", "" & rs("email"))
                    s = s & AddMBXml(p, 14, "Mobile", "" & rs("CellPhone"))
                    s = s & AddMBXml(p, 14, "Fax", "" & rs("Fax"))
                s = s & "</ClientContactAdd>" & vbCrLf
                s = s & "<ClientContactAdd>" & vbCrLf
                    s = s & AddMBXml(c, 30, "ContactName", "" & rs("Contact2"))
                    s = s & AddMBXml(c, 20, "JobTitle", "CoBuyer")
                    s = s & AddMBXml(p, 14, "Phone", "" & rs("C_Phone"))
                    s = s & AddMBXml(c, 30, "Email", "" & rs("C_email"))
                    s = s & AddMBXml(p, 14, "Mobile", "" & rs("C_CellPhone"))
                    s = s & AddMBXml(p, 14, "Fax", "" & rs("C_Fax"))
                s = s & "</ClientContactAdd>" & vbCrLf
                s = s & "</ClientAddRq>" & vbCrLf
            End If
            rs.MoveNext
        Wend
        s = s & EndMBXml()
        If i > 0 Then Call SubmitMBXml(s, HFApp.Options(MasterBuilderPWD))
    End If
Exit Sub
eh: If InStr(1, Err.Description, "duplicate") > 0 Then
        Resume Next
    Else
        Call errHandler(SRCFILE & "WriteCustomerToSage100", s)
    End If
End Sub





Public Sub ExportData(SelectStatement As String, DialogTitle As String)
    Call FDataExport.ExportData(SelectStatement, DialogTitle)
End Sub


Public Function EstFieldSize(field As String) As Long
    Dim i As Long
    Select Case field
        Case "Phase":         i = 20
        Case "PhaseDesc":     i = 150
        Case "Item":          i = 15
        Case "ItemDesc":      i = 150
        Case "ItemNotes":     i = 1520
        Case "AssemblyDesc":  i = 70
        Case "AssemblyNotes": i = 1520
        Case Else: Err.Raise 5
    End Select
    EstFieldSize = i
End Function
Public Function FormatPhase(Phase As String) As String
    If Val(HFApp.Options(EstimatingMajorVersion)) >= 13 Then
        FormatPhase = Trim(Phase)
    Else
        'prior to v13 was right padded
        FormatPhase = Format(Trim(Phase), String(EstFieldSize("Phase"), "@"))
    End If
End Function
Public Function FormatItem(item As String) As String
    Select Case True
        
        Case item = "AppVersion"
            FormatItem = App.Major & "." & App.Minor & "." & App.Revision

        Case item = "TitlebarCaption"
            FormatItem = " (" & Trim(LoginDSN) & ", Login User: " & Trim(LoginID) & ") - version " & App.Major & "." & App.Minor & " Division: " & "" & HFApp.SqlExec("Select DivisionCode from Divisions where DivisionID = " & DivisionID, dbHomefront)(0)

    
        'original intent of function was to format the sage estimating item value that was version dependent
        Case Val(HFApp.Options(EstimatingMajorVersion)) >= 13
            FormatItem = Trim(item)
        Case Else
            'prior to v13 was right padded
            FormatItem = Format(Trim(item), String(EstFieldSize("Item"), "@"))
    End Select
End Function
Public Function FormatJob(Job As String) As String
    Select Case HFApp.Options(AccountingSystem)
        Case asTimberline:      FormatJob = Format(Replace(StripFormating(Trim(Job)), ".", ""), HFApp.Options(Job_Mask))
        Case Else:              FormatJob = Job
    End Select
End Function
Public Function FormatCostCode(CostCode As String) As String
    Select Case HFApp.Options(AccountingSystem)
        Case asTimberline:      FormatCostCode = Format(Replace(StripFormating(Trim(CostCode)), ".", ""), HFApp.Options(Cost_Code_Mask))
        Case asMasterBuilder:   FormatCostCode = Format(Val(Trim(CostCode)), "0.000")
        Case Else:              FormatCostCode = CostCode
    End Select
End Function
Public Function XmlMbStart(Company As String, UserName As String) As String
    XmlMbStart = XmlMB.StartMBXml(Company, UserName)
End Function
Public Function XmlMBEnd() As String
    XmlMBEnd = XmlMB.EndMBXml
End Function
Public Function XmlMBAdd(FieldType As XMLFieldTypes, SIZE As Double, Name As String, Value) As String
     XmlMBAdd = XmlMB.AddMBXml(FieldType, SIZE, Name, Value)
End Function
Public Function XmlMbSubmit(MBXmlMessage As String, Pswd As String) As Boolean
    XmlMbSubmit = XmlMB.SubmitMBXml(MBXmlMessage, Pswd)
End Function

Public Function XmlQBSubmit(XmlMessage As String) As String
    XmlQBSubmit = XmlQB.SubmitQBXml(XmlMessage)
End Function
Public Function XmlQBStart() As String
    XmlQBStart = XmlQB.StartQBXml
End Function
Public Function XmlQBEnd() As String
    XmlQBEnd = XmlQB.EndQBXml
End Function
Public Function XmlQBAdd(FieldType As XMLFieldTypes, SIZE As Double, Name As String, Value) As String
    Value = Replace(Value, Chr(188), "1/4")
    Value = Replace(Value, Chr(189), "1/2")
    Value = Replace(Value, Chr(190), "3/4")
    Value = Replace(Value, Chr(150), "-")
    Value = Replace(Value, Chr(145), Chr(39)) 'left curly single quote
    Value = Replace(Value, Chr(146), Chr(39)) 'right curly single quote
    Value = Replace(Value, Chr(147), Chr(34)) 'left curly double quote
    Value = Replace(Value, Chr(148), Chr(34)) 'right curly double quote
    Value = Replace(Value, Chr(152), Chr(34)) 'another weird double quote
    
    Value = Replace(Value, "&lt;", "<")
    Value = Replace(Value, "&gt;", ">")
    
    
    XmlQBAdd = XmlQB.AddQBXml(FieldType, SIZE, Name, Value)
End Function

Private Function StripFormating(s As String) As String
    Const FORMATCHRS = "+=_-)(*&^%$#@!~`[]{}\|/?'""<>;:,"
    Dim i As Long
    For i = 1 To Len(FORMATCHRS)
        s = Replace(s, Mid(FORMATCHRS, i, 1), "")
        s = Replace(s, Chr(189), "1/2")
    Next
    StripFormating = s
End Function

Public Sub LogInfoToFile(Filename As String, s As String)
    Dim i As Integer
    
    On Error Resume Next
    Call CreatePath("", FilePath(Filename))
    On Error GoTo 0
    
    'log info
    i = FreeFile()
    Open Filename For Append As #i
    Print #i, String(50, "=")
    Print #i, s
    Close #i
End Sub

Friend Function AcquireLicense() As Boolean
    Dim Filename As String
    Dim s As String
    Dim i As Long
    Dim a As Long
    Dim b As Long
    
    
    
    Filename = HFApp.LicenseFile
    If Not FileExists(Filename) Then
        MsgBox "Access denied." & vbCrLf & vbCrLf & "No license could be found.", vbCritical, "HomeFront"
        Exit Function
    End If
    
    With mLicense
        
        If Not .ReadLicense(Filename) Then
            MsgBox "Access denied." & vbCrLf & vbCrLf & "Invalid license. Please contact your HomeFront solution provider for assistance.", vbCritical, "HomeFront"
            Exit Function
        End If
        
        
        If .Leased And .LeaseExpiryDate < VBA.Date Then
            MsgBox "Your software lease has expired. Please contact your HomeFront solution provider for assistance.", vbInformation, "HomeFront"
            Exit Function
        End If
        
'        If Not da.AuthorizeClient(.ClientID) Then
'            MsgBox "Access denied." & vbCrLf & vbCrLf & "Your license has expired. Please contact your HomeFront solution provider for assistance.", vbCritical, "HomeFront"
'            Exit Function
'        End If
        
        
        If JobSimplicity Then mProductCode = "JobSimplicity"
        
        If .ModuleDemo(.ModuleIndex(mProductCode)) Then
            i = .ModuleDemoExpiry(.ModuleIndex(mProductCode)) - VBA.Date()
            If i < 1 Then
                MsgBox "The evaluation period has expired. PLease contact your HomeFront solution provider for assistance.", vbInformation, "HomeFront"
                Exit Function
            End If
        End If
        
        
        mProductName = .ModuleName(.ModuleIndex(mProductCode))
        
        a = .ModuleSeats(.ModuleIndex(mProductCode))
        If mProductCode = "HFStaging" Then a = 9999
        If mProductCode = "HFSalesMgmt" Then a = 9999
        If mProductCode = "HFSynch" Then a = 9999
        b = .ModuleSeatsInUse(.ModuleIndex(mProductCode), s)

        If a > b Then
            AcquireLicense = True
        Else
            If s = "" Then
                MsgBox "All " & a & " licenses are in use." & vbCrLf & vbCrLf & "Please contact your HomeFront solution provider if you require additional seats.", vbCritical, "HomeFront"
            Else
                MsgBox "SQL Server configuration error." & vbCrLf & "Connecting user requires VIEW SERVER STATE. Please contact your system administrator.", vbCritical, "HomeFront"
            End If
        End If
    End With
    
End Function
Public Property Get LicenseNumber() As String
    LicenseNumber = mLicense.SerialNumber
End Property
Public Property Get ClientID() As String
    ClientID = mLicense.ClientID
End Property
Public Property Get LicensedOwner() As String
    LicensedOwner = mLicense.CompanyName
End Property
Public Property Get LicensesFree() As Long
    LicensesFree = mLicense.ModuleSeats(mLicense.ModuleIndex(mProductCode)) - mLicense.ModuleSeatsInUse(mLicense.ModuleIndex(mProductCode), "")
End Property
Public Property Get LicensedUsers() As Long
    LicensedUsers = mLicense.ModuleSeats(mLicense.ModuleIndex(mProductCode))
End Property
Friend Property Get LicensedSeats(ProductCode As String) As Long
    LicensedSeats = mLicense.ModuleSeats(mLicense.ModuleIndex(ProductCode))
End Property
Public Property Get LicenseLimit() As String
'    Dim d As Long
'    Dim m As Long
'    Dim Y As Long
'    Dim r As Long
'    Call LLGetExpiry(d, m, Y, r)
'    Select Case True
'        Case r <> 0: LicenseLimit = r & " runs left in the trial period."
'        Case d <> 0: LicenseLimit = (DateSerial(Y, m, d) - VBA.Date) & " days left in the trial period."
'        Case Else:   LicenseLimit = ""
'    End Select
End Property

Public Sub WriteCustomerToAccounting(JobNumber As String, CustomerNumber As String)
    Select Case Options(AccountingSystem)
        Case asTimberline:           'sales does this
        Case asIntacct:              Call WriteCustomerToIntacct(JobNumber, CustomerNumber)
        Case asMasterBuilder:        Call WriteCustomerToSage100(JobNumber, CustomerNumber)
        Case asQuickBooks:           Call WriteCustomerToQuickBooks(JobNumber, CustomerNumber)
        
        Case asQuickbooksOnline:     Call WriteCustomerToQBO(JobNumber, CustomerNumber)
        
        Case asSimply:
                                      If JobNumber = "" Then
                                          Call WriteCustomerToSage50(JobNumber, CustomerNumber, True)
                                      Else
                                          Call WriteCustomerToSage50(JobNumber, CustomerNumber)
                                      End If
        Case Else:
    End Select
End Sub

Private Sub WriteCustomerToSage50(JobNumber As String, CustomerNumber, Optional ARCustomers As Boolean)

    Dim X As SimplyWrapper.Customer
    Dim s As String
    Dim ID As Long
    Dim rs As Recordset
        
    'select data from hf tables
    If ARCustomers = True And "" & CustomerNumber <> "" Then
        s = ""
        s = "select arcustomer Customer_No,description,comments,billaddr1 Address1,billaddr2 Address2,billcity City,billprovince Province,billpostalcode Zip"
        s = s & ",'' Contact1,phone1 Phone,fax1 Fax,Email1 Email,''CellPhone,'' Contact2, '' C_Phone,'' C_Fax"
        s = s & ",'' c_email,'' C_cellphone,null Purchased_Date  from arcustomers"
        s = s & " where arcustomer = " & DbQuote(Str, CustomerNumber)
    Else
        s = ""
        s = s & "SELECT c.customer_no" & vbCrLf
        s = s & "      ,c.Description" & vbCrLf
        s = s & "      ,c.Address1" & vbCrLf
        s = s & "      ,c.Address2" & vbCrLf
        s = s & "      ,c.City" & vbCrLf
        s = s & "      ,c.Province" & vbCrLf
        s = s & "      ,c.Zip" & vbCrLf
        s = s & "      ,ISNULL(c.Customer_Name,'') + ISNULL(' '+c.Customer_LName,'') Contact1" & vbCrLf
        s = s & "      ,c.Phone" & vbCrLf
        s = s & "      ,c.fax" & vbCrLf
        s = s & "      ,c.email" & vbCrLf
        s = s & "      ,c.cellphone" & vbCrLf
        s = s & "      ,ISNULL(c.CoBuyer_Name,'') + ISNULL(' '+c.CoBuyer_LName,'') Contact2" & vbCrLf
        s = s & "      ,c.C_Phone" & vbCrLf
        s = s & "      ,c.C_fax" & vbCrLf
        s = s & "      ,c.C_email" & vbCrLf
        s = s & "      ,c.C_cellphone" & vbCrLf
        s = s & "      ,Purchased_Date" & vbCrLf
        s = s & "      ,c.Comments" & vbCrLf
        s = s & "  FROM tblCustomers c ,system_setup s" & vbCrLf
        s = s & " WHERE isnull(c.ar_customer_deposit,'')=''" & vbCrLf
        s = s & "   and isnull(c.Job_no,'')<>''" & vbCrLf
        s = s & "   and((c.purchased=1  AND c.cancelled=0  AND c.inactive=0 AND ((s.accounting_approve=1 AND c.approved=1)" & vbCrLf
        s = s & "        OR (s.accounting_approve=0 AND c.contract_assigned=1))) OR (c.home_selection <>'PreSale' AND c.sold_to_Customer=''))" & vbCrLf
        If CustomerNumber = "" Then
            s = s & "   AND job_no=" & DbQuote(Str, JobNumber) & vbCrLf
        Else
            s = s & "   AND Customer_no=" & DbQuote(Str, CustomerNumber) & vbCrLf
        End If
        s = s & " and s.id = " & HFApp.DivisionID & " and c.DivisionID = " & HFApp.DivisionID
    End If
    Set rs = HFApp.SqlExec(s)
    
    If rs.EOF Then Exit Sub
    
    Set X = New SimplyWrapper.Customer
    Call X.OpenDB(HFApp.Options(SimplyDataFile), HFApp.Options.ValueByName("SimplyUID"), HFApp.Options.ValueByName("SimplyPWD"))
    While Not rs.EOF
        Call X.CreateCustomer("" & rs("description") & "  (" & rs("customer_no") & ")", _
                              "" & rs("contact1"), _
                              "" & rs("phone"), _
                              "" & rs("cellphone"), _
                              "" & rs("email"), _
                              "" & rs("fax"), _
                              "" & rs("address1"), _
                              "" & rs("address2"), _
                              "" & rs("city"), _
                              "" & rs("province"), _
                              "" & rs("zip"), _
                              "")
        ID = X.SaveCustomer()
        If ID <> 0 Then
            s = "update tblcustomers set ar_customer_deposit=" & DbQuote(Str, ID) & " where customer_no=" & DbQuote(Str, "" & rs("customer_no"))
            Call HFApp.SqlExec(s)
            If ARCustomers = True Then
                s = "update arcustomers set ExternalID = " & DbQuote(Str, ID) & " where ARCustomer=" & DbQuote(Str, "" & rs("customer_no"))
                Call HFApp.SqlExec(s)
            End If
        End If
        Call rs.MoveNext
    Wend
    Call X.CloseDB
    
End Sub

Public Sub WriteJobToAccounting(JobNumber As String)
    Dim rs As Recordset
    Dim s As String
    
    
'JobNumber should have been passed byval. Too late to change it now so just make sure it doesnt change
Dim jn As String
jn = JobNumber
    
    
    s = "select isquote from tbljobs where job_no=" & DbQuote(Str, JobNumber)
    Set rs = HFApp.SqlExec(s)
    If rs.EOF Then Exit Sub
    If "" & rs("isquote") = "True" Then Exit Sub


    Call WriteCustomerToAccounting(JobNumber, "")
JobNumber = jn
        
    Select Case Options(AccountingSystem)
        Case asTimberline:     Call WriteJobToSage300(JobNumber)
        Case asIntacct:        Call WriteJobToIntacct(JobNumber)
        Case asMasterBuilder:  Call WriteJobToSage100(JobNumber)
        Case asSimply:         Call WriteJobToSage50(JobNumber)
        
        Case asQuickBooks, asQuickbooksOnline:
            'QuickBooks customer and job are same thing. Job and customer records are written in WriteCustomerToAccounting()
        
        Case Else:
    End Select
JobNumber = jn
    
    
End Sub

Public Sub OpenSimplyODBC()
On Error GoTo eh

    Set SimplyOdbcUtilities = New SimplyWrapper.ODBCUtilities
    Call SimplyOdbcUtilities.StartMySQL(HFApp.Options.Value(SimplyDataFile), HFApp.Options.ValueByName("SimplyUID"), HFApp.Options.ValueByName("SimplyPWD"))
    ConnectionString(dbAccounting) = SimplyOdbcUtilities.GetOdbcConnectionString
    ConnectionString(dbReceivables) = ConnectionString(dbAccounting)
    ConnectionString(dbAccountingDictionary) = ConnectionString(dbAccounting)
    ConnectionString(dbReceivablesDictionary) = ConnectionString(dbAccounting)
    Call Databases(dbAccounting).Open(ConnectionString(dbAccounting))
    Set Databases(dbReceivables) = Databases(dbAccounting)
    Set Databases(dbAccountingDictionary) = Databases(dbAccounting)
    Set Databases(dbReceivablesDictionary) = Databases(dbAccounting)

Exit Sub
eh: MsgBox "Failed to connect to Simply Accounting" & vbCrLf & Err.Description & vbCrLf & vbCrLf & HFApp.Options.Value(SimplyDataFile), vbExclamation, "HFSystem.dll"
End Sub
Public Sub CloseSimplyODBC()
On Error Resume Next
    Databases(dbAccounting).Close
    Databases(dbReceivables).Close
    Databases(dbAccountingDictionary).Close
    Databases(dbReceivablesDictionary).Close
    Call SimplyOdbcUtilities.StopMySQL
    Set SimplyOdbcUtilities = Nothing
End Sub


Private Sub WriteJobToSage50(JobNumber As String)
On Error GoTo eh
    Dim X As SimplyWrapper.Project
    Dim s As String
    Dim ID As Long
    Dim rs As Recordset
        
    'select data from hf tables
    s = ""
    s = s & "select job_no,description" & vbCrLf
    s = s & "from tbljobs" & vbCrLf
    s = s & "where job_no=" & DbQuote(Str, JobNumber) & vbCrLf
    s = s & "  and isnull(externaljobid,'')=''" & vbCrLf
    s = s & " and DivisionID = " & HFApp.DivisionID
    Set rs = HFApp.SqlExec(s)
    If rs.EOF Then Exit Sub
    
    Set X = New SimplyWrapper.Project
    
    Call X.OpenDB(HFApp.Options(SimplyDataFile), HFApp.Options.ValueByName("SimplyUID"), HFApp.Options.ValueByName("SimplyPWD"))
    Call X.CreateProject("" & rs("job_no") & " - " & rs("Description"))
    
    ID = X.SaveProject()
    If ID <> 0 Then
        s = "update tbljobs set externaljobid=" & DbQuote(Str, ID) & " where DivisionID = " & HFApp.DivisionID & " and job_no=" & DbQuote(Str, "" & rs("job_no"))
        Call HFApp.SqlExec(s)
    End If
    
    
    Call X.CloseDB
Exit Sub
eh: Call errHandler(SRCFILE & "WriteJobToSage50", s)
End Sub




Private Sub Class_Terminate()
On Error Resume Next
'    Call mOLButton.Delete
    Call SetErrorMode(SEM_NOGPFAULTERRORBOX)
End Sub

Friend Function CreateDivisionUsers() As Boolean
On Error GoTo eh
    Dim s As String
    
    s = ""
    s = s & "CREATE TABLE [dbo].[DivisionUsers](" & vbCrLf
    s = s & "   [DivisionID] [int] NOT NULL," & vbCrLf
    s = s & "   [USERID] [varchar](10) NOT NULL," & vbCrLf
    s = s & " CONSTRAINT [PK_DivisionUsers] PRIMARY KEY CLUSTERED " & vbCrLf
    s = s & "(" & vbCrLf
    s = s & "   [DivisionID] ASC," & vbCrLf
    s = s & "   [USERID] ASC" & vbCrLf
    s = s & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]" & vbCrLf
    s = s & ") ON [PRIMARY]" & vbCrLf
    HFApp.SqlExec s
        
    s = "insert into DivisionUsers(DivisionID,UserID) Select d.divisionid,u.user_id from user_manager u join divisions d on(1=1)"
    HFApp.SqlExec s

    CreateDivisionUsers = True

Exit Function
eh:
Select Case True
    Case InStr(1, Err.Description, "There is already an object named", vbTextCompare)
        CreateDivisionUsers = True
    Case Else
        'MsgBox "Unable to upgrade the Homefront database" & IIf(Err.Number <> 0, vbCrLf & Parse(Err.Description, Parse(Err.Description, , "]"), "]"), ""), vbCritical, App.ProductName
        'CreateDivisionUsers = False
        'WHO CARES! THIS SHOULD NEVER HAPPEN. EVEN IF IT DOES IT WONT BREAK ANYTHING
        CreateDivisionUsers = True
End Select
End Function
