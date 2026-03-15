VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLicense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ClientID        As String
Public CompanyName     As String
Public Leased          As Boolean
Public LeaseExpiryDate As Date
     
Private mSerialNumber As String
Private mModules(18)  As ModuleType
Private Type ModuleType
    Code       As String  'exename without .exe           ie: hfest
    Name       As String  'name as printed in license         Precision Builder
    Program    As String  'name as appears in SQLServer       HomeFront Precision Builder
    Seats      As Long    'number of concurrent uses
    Demo       As Boolean
    CountType  As String  'what field is used to determine seat usage (loginame or hostprocess)
    DemoExpiry As Date
End Type

Private Sub Class_Initialize()
    Dim i As Long
    
        
    
    'do not change the order of these 6... you will break old license files
    i = i + 1: mModules(i).CountType = "hostprocess":     mModules(i).Code = "hfsales":           mModules(i).Name = "Profit Builder    ":   mModules(i).Program = "'Sales'" 'verified
    i = i + 1: mModules(i).CountType = "hostprocess":     mModules(i).Code = "hfest":             mModules(i).Name = "Precision Builder ":   mModules(i).Program = "'Precision Builder'" 'verified
    i = i + 1: mModules(i).CountType = "hostprocess":     mModules(i).Code = "hfsched":           mModules(i).Name = "Project Builder   ":   mModules(i).Program = "'Project Editor'" 'verified
    i = i + 1: mModules(i).CountType = "hostprocess":     mModules(i).Code = "hfwarranty":        mModules(i).Name = "Quality Builder   ":   mModules(i).Program = "'Quality Builder'" 'verified
    i = i + 1: mModules(i).CountType = "hostprocess":     mModules(i).Code = "hfpayables":        mModules(i).Name = "Payables Desk     ":   mModules(i).Program = "'HomeFront Payables'" 'verified
    i = i + 1: mModules(i).CountType = "loginame":        mModules(i).Code = "JobSimplicity":     mModules(i).Name = "Job Simplicity    ":   mModules(i).Program = "'Sales','Precision Builder','Project Editor','Quality Builder','Payables Desk'"
    
    i = i + 1: mModules(i).CountType = "hostprocess":     mModules(i).Code = "hfsalesmgmt":       mModules(i).Name = "Sales Management  ":   mModules(i).Program = "'Sales Management'"
    i = i + 1: mModules(i).CountType = "hostprocess":     mModules(i).Code = "hfstaging":         mModules(i).Name = "Data Staging      ":   mModules(i).Program = "'HomeFront Data Staging'"
    i = i + 1: mModules(i).CountType = "hostprocess":     mModules(i).Code = "hfdataconnector":   mModules(i).Name = "Data Connector    ":   mModules(i).Program = "'HomeFront Data Connector'"
    i = i + 1: mModules(i).CountType = "hostprocess":     mModules(i).Code = "workflowbuilder":   mModules(i).Name = "Workflow Builder  ":   mModules(i).Program = "'AutoNotice'" 'verified
    i = i + 1: mModules(i).CountType = "hostprocess":     mModules(i).Code = "CommandCenter":     mModules(i).Name = "Command Center    ":   mModules(i).Program = "'Command Center'"
    i = i + 1: mModules(i).CountType = "hostprocess":     mModules(i).Code = "Workticket":        mModules(i).Name = "Workticket        ":   mModules(i).Program = "'Workticket'" 'verified
    
    'web apps
    i = i + 1: mModules(i).CountType = "hostprocess":     mModules(i).Code = "FieldPO":           mModules(i).Name = "Field Purchasing  ":    mModules(i).Program = "'zyberNothing'"
    i = i + 1: mModules(i).CountType = "hostprocess":     mModules(i).Code = "FieldSched":        mModules(i).Name = "Field Scheduling  ":    mModules(i).Program = "'zyberNothing'"
    i = i + 1: mModules(i).CountType = "hostprocess":     mModules(i).Code = "CustomPortal":      mModules(i).Name = "Customer Portal   ":    mModules(i).Program = "'zyberNothing'"
    i = i + 1: mModules(i).CountType = "hostprocess":     mModules(i).Code = "VendorPortal":      mModules(i).Name = "Vendor Portal     ":    mModules(i).Program = "'zyberNothing'"
    i = i + 1: mModules(i).CountType = "hostprocess":     mModules(i).Code = "CRM":               mModules(i).Name = "Homefront CRM     ":    mModules(i).Program = "'zyberNothing'"
    
    'utility apps
    i = i + 1: mModules(i).CountType = "hostprocess":     mModules(i).Code = "UNSECUREDAPP":      mModules(i).Name = "Unsecured App     ":    mModules(i).Program = "'zyberNothing'"


'to add a new module
' 1.) add to mModules array
' 2.) add to LoggedInUsers view

End Sub

Public Property Get Modules() As Long
    Modules = UBound(mModules)
End Property

Public Property Get ModuleIndex(ByVal Code As String) As Long
    Dim i As Long
    For i = 1 To Modules
        If UCase(Code) = UCase(mModules(i).Code) Then
            ModuleIndex = i
            Exit Property
        End If
    Next
End Property

Public Property Get ModuleCode(ByVal Index As Long) As String
    ModuleCode = mModules(Index).Code
End Property
Public Property Let ModuleCode(ByVal Index As Long, RHS As String)
    mModules(Index).Code = RHS
End Property
Public Property Get ModuleName(ByVal Index As Long) As String
    ModuleName = mModules(Index).Name
End Property
Public Property Let ModuleName(ByVal Index As Long, RHS As String)
    mModules(Index).Name = RHS
End Property

Public Property Get ModuleDemo(ByVal Index As Long) As Boolean
    ModuleDemo = mModules(Index).Demo
End Property
Public Property Let ModuleDemo(ByVal Index As Long, RHS As Boolean)
    mModules(Index).Demo = RHS
End Property

Public Property Get ModuleDemoExpiry(ByVal Index As Long) As Date
    ModuleDemoExpiry = mModules(Index).DemoExpiry
End Property
Public Property Let ModuleDemoExpiry(ByVal Index As Long, RHS As Date)
    mModules(Index).DemoExpiry = RHS
End Property


Public Property Get ModuleSeats(ByVal Index As Long) As Long
    ModuleSeats = mModules(Index).Seats
End Property
Public Property Let ModuleSeats(ByVal Index As Long, RHS As Long)
    mModules(Index).Seats = RHS
End Property
Public Property Get ModuleSeatsInUse(ByVal Index As Long, ByRef Errors As String) As Long
On Error Resume Next
    Dim s As String
    
    s = "select has_perms_by_name(null,null,'view server state')"
    If Val("" & HFApp.SqlExec(s, dbHomefront)(0)) <> 1 Then
        'user doesnt have view server state so the cant see sysprocesses
        Errors = "nsf"
        ModuleSeatsInUse = ModuleSeats(Index)
    Else
        Select Case mModules(Index).CountType
        
        
            'homefront license
            Case "hostprocess"
                s = ""
                s = s & "select count(distinct hostprocess)" & vbCrLf
                s = s & "  from master.dbo.sysprocesses" & vbCrLf
                s = s & " where hostprocess<>host_id()" & vbCrLf
                s = s & "   and program_name in(" & mModules(Index).Program & ")"
                
            'job simplicity license
            Case Else
                s = ""
                s = s & "select count(distinct isnull(loginame,''))" & vbCrLf
                s = s & "  from master.dbo.sysprocesses" & vbCrLf
                s = s & " where loginame<>SYSTEM_USER" & vbCrLf
                s = s & "   and program_name in(" & mModules(Index).Program & ")"
                
                
        End Select
        ModuleSeatsInUse = Val("" & HFApp.SqlExec(s, dbHomefront)(0))
    End If
    
If Err.Number <> 0 Then MsgBox Err.Description
End Property








Private Function CreateSerialNumber() As String
    Dim c As New CCrypto
    Dim s As String
    Dim i As Long
    s = ClientID & Chr(1) & _
        CompanyName & Chr(1) & _
        UsageMode & Chr(1) & _
        Format(LeaseExpiryDate, "mmm d, yyyy") & Chr(1)
    
    For i = 1 To Modules
        If mModules(i).Code <> "UNSECUREDAPP" Then
            If mModules(i).Demo Then
                s = s & "Evaluation period expires " & Format(mModules(i).DemoExpiry, "mmm d, yyyy") & vbCrLf
            Else
                If mModules(i).Seats > 0 Or i < 6 Then s = s & mModules(i).Seats & Chr(1)
            End If
        End If
    Next
        
    CreateSerialNumber = Format$(c.CreateHash(UCase(s)), "@@@@@@@@-@@@@-@@@@-@@@@-@@@@@@@@@@@@")
End Function













Public Sub WriteLicense(ByVal filename As String)
    Dim i As Integer
    i = FreeFile
    Open filename For Output As #i
    Print #i, Me.LicenseText
    Close #i
End Sub

Public Function ReadLicense(ByVal filename As String) As Boolean
On Error GoTo eh
    Dim i As Long
    Dim s As String

    ClientID = IniGet(filename, "Licensed To", "Client ID")
    CompanyName = IniGet(filename, "Licensed To", "Company")
    mSerialNumber = IniGet(filename, "Licensed To", "Serial Number")
    s = IniGet(filename, "Licensed To", "Lease Expiry")
    If s = "" Then
        Leased = False
        LeaseExpiryDate = 0
    Else
        Leased = True
        LeaseExpiryDate = DateValue(s)
    End If
    

    For i = 1 To Modules
        s = IniGet(filename, "Licensed Modules", mModules(i).Name)
        
        Select Case True
                
            Case left(s, 10) = "Evaluation"
                mModules(i).Seats = 999999999
                mModules(i).Demo = True
                mModules(i).DemoExpiry = DateValue(Mid(s, 27))
            
            Case IsNumeric(s)
                mModules(i).Seats = Val(s)
                mModules(i).Demo = False
                mModules(i).DemoExpiry = 0
                
            Case mModules(i).Code = "UNSECUREDAPP"
                mModules(i).Seats = 999999999
                mModules(i).Demo = False
                mModules(i).DemoExpiry = 0
                        
            Case Else
                mModules(i).Seats = 0
                mModules(i).Demo = False
                mModules(i).DemoExpiry = 0
        End Select
    Next

    If mSerialNumber = CreateSerialNumber Then
        ReadLicense = True
    Else
        ReadLicense = False
        ClientID = ""
        CompanyName = ""
        mSerialNumber = ""
        Leased = False
        For i = 1 To Modules
            mModules(i).Seats = 0
            mModules(i).Demo = False
            mModules(i).DemoExpiry = 0
        Next
    End If
Exit Function
eh: ReadLicense = False
End Function

Public Property Get LicenseText() As String
    Dim s As String
    Dim i As Long

    s = ""
    s = s & "[Licensed To]" & vbCrLf
    s = s & "  Client ID     = " & ClientID & vbCrLf
    s = s & "  Company       = " & CompanyName & vbCrLf
    If Leased Then
    s = s & "  Lease Expiry  = " & Format(LeaseExpiryDate, "mmm d, yyyy") & vbCrLf
    End If
    s = s & "  Serial Number = " & CreateSerialNumber() & vbCrLf
    s = s & vbCrLf
    s = s & "[Licensed Modules]" & vbCrLf
    For i = 1 To Modules
        If mModules(i).Demo Then
            s = s & "  " & mModules(i).Name & " = Evaluation period expires " & Format(mModules(i).DemoExpiry, "mmm d, yyyy") & vbCrLf
        ElseIf mModules(i).Seats > 0 Then
            s = s & "  " & mModules(i).Name & " = " & mModules(i).Seats & vbCrLf
        End If
    Next
    LicenseText = s
End Property

Public Property Get SerialNumber() As String
    SerialNumber = mSerialNumber
End Property

Private Property Get UsageMode() As String
    UsageMode = IIf(Leased, "Leased", "Licensed")
End Property


