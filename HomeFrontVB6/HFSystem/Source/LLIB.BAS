Attribute VB_Name = "MLicensingLib"

' MODULE:  LLIB.BAS      Version 3.00
'
' PURPOSE: Licensing Library definitions
'
' AUTHOR:  GPP Software
'
' DATE:    10/09/2004

Option Explicit

' License version
Global Const LLMINFILEVERSION = 103
Global Const LLFILEVERSION = 300

' General size definitions
Global Const LLMAXLICENSELEN = 56
Global Const LLMINLICENSELEN = 22
Global Const LLMAXEXTENSIONPACKLICENSELEN = 30
Global Const LLMAXUSERNAMELEN = 32
Global Const LLMAXCOMPANYNAMELEN = 32
Global Const LLMAXEXTENSIONPACKS = 16
Global Const LLMAXPRODCODELEN = 3
Global Const LLMAXPRODNAMELEN = 32
Global Const LLMAXADDRESSLEN = 32
Global Const LLMAXPHONELEN = 32
Global Const LLMAXEMAILLEN = 64
Global Const LLMAXMACHINEIDLEN = 8
Global Const LLMAXMACHINENAMELEN = 32
Global Const LLMAXFUNCTIONNAMELEN = 24
Global Const LLMAXVARIABLES = 6
Global Const LLMAXVARTEXTLEN = 32

Global Const LLMAXFILENAMELEN = 255
        
' Error codes
Global Const LLER_OK = 0
Global Const LLER_LICENSEFILEMISSING = 1
Global Const LLER_LICENSEEXPIRED = 2
Global Const LLER_LICENSEUNRECFILE = 3
Global Const LLER_LICENSEUNRECVER = 4
Global Const LLER_LICENSEUNRECLICENSENOVER = 5
Global Const LLER_LICENSEINVALIDOPTION = 6
Global Const LLER_LICENSEINVALIDPRODCODE = 7
Global Const LLER_LICENSEINVALID = 8
Global Const LLER_LICENSECREATEFAIL = 9
Global Const LLER_EXPIRESONDATE = 10
Global Const LLER_EXPIRESONRUNS = 11
Global Const LLER_LICENSENOUSERSLEFT = 12
Global Const LLER_LICENSETAMPERED = 13
Global Const LLER_LICENSEINVALIDMODULUS = 14
Global Const LLER_NOTRUNLIMITED = 15
Global Const LLER_FILEALREADYEXISTS = 16
Global Const LLER_LICENSEDISABLED = 18
Global Const LLER_LICENSEOPENFAIL = 20
Global Const LLER_LOCKCREATEFAIL = 21
Global Const LLER_LOCKOPENFAIL = 22
Global Const LLER_LICENSEFILENOTOPEN = 23
Global Const LLER_LICENSENOTSET = 28
Global Const LLER_INVALIDEXTENSIONPACKLICENSE = 29
Global Const LLER_INVALIDPACKFORPRODREG = 30
Global Const LLER_EXTENSIONPACKALREADYENTERED = 31
Global Const LLER_NOEXTENSIONPACKSPACE = 32
Global Const LLER_CANNOTSETDUETOEXTENSIONPACKS = 33
Global Const LLER_INVALIDEXTENSIONPACKINDEX = 35
Global Const LLER_MACHINEIDERR = 36
Global Const LLER_INVALIDMACHINEID = 37
Global Const LLER_INVALIDLICENSENOLEN = 38
Global Const LLER_INVALIDMACHINEIDFORMAT = 39
Global Const LLER_INVALIDENCRYPTIONKEY = 40
Global Const LLER_INVALIDEXPIRYDATE = 41
Global Const LLER_LICENSEINVALIDTYPE = 42
Global Const LLER_INVALIDUSERNAME = 43
Global Const LLER_INVALIDCOMPANY = 44
Global Const LLER_EXTENSIONPACKSNOTSUPPORTED = 45
Global Const LLER_CANTREENTERUNREGLICENSE = 46
Global Const LLER_LICENSEINVALIDKEY = 47
Global Const LLER_NOVARSPACE = 48
Global Const LLER_VARIABLENOTFOUND = 49
Global Const LLER_VARIABLESNOTSUPPORTED = 50
Global Const LLER_INVALIDFILEVERSION = 51
Global Const LLER_LLMEMALLOCFAIL = 52
Global Const LLER_INVALIDAPPVERSION = 53
Global Const LLER_HTTPNOTFOUND = 54
Global Const LLER_HTTPNORESPONSE = 55
        
' License types
Global Const LICTYPE_STANDARD = 0
Global Const LICTYPE_USERPACK = 1

' License options
Global Const OPTION_UNREGISTERED = "U"
Global Const OPTION_LIMITED = "L"
Global Const OPTION_TEMPORARYREGISTRATION = "T"
Global Const OPTION_FULLREGISTRATION = "R"
Global Const OPTION_EXTRAUSERLICENSE = "X"

' License functionality options
Global Const OPTION_RUNLIMITED = 1
Global Const OPTION_DAYLIMITED = 2
Global Const OPTION_EXPIRYDATE = 4
Global Const OPTION_MACHINEID = 8
Global Const OPTION_USERS = 16
Global Const OPTION_NAME = 32
Global Const OPTION_COMPANY = 64
Global Const OPTION_FUNCTIONFLAGS = 128
Global Const OPTION_ACCEPTSEXTENSIONPACKS = 256
Global Const OPTION_ACCEPTSVARIABLES = 512
Global Const OPTION_APPVERSION = 1024

' Function flags
Global Const LLFUNCTION_1 = 1
Global Const LLFUNCTION_2 = 2
Global Const LLFUNCTION_3 = 2 ^ 2
Global Const LLFUNCTION_4 = 2 ^ 3
Global Const LLFUNCTION_5 = 2 ^ 4
Global Const LLFUNCTION_6 = 2 ^ 5
Global Const LLFUNCTION_7 = 2 ^ 6
Global Const LLFUNCTION_8 = 2 ^ 7
Global Const LLFUNCTION_9 = 2 ^ 8
Global Const LLFUNCTION_10 = 2 ^ 9
Global Const LLFUNCTION_11 = 2 ^ 10
Global Const LLFUNCTION_12 = 2 ^ 11
Global Const LLFUNCTION_13 = 2 ^ 12
Global Const LLFUNCTION_14 = 2 ^ 13
Global Const LLFUNCTION_15 = 2 ^ 14
Global Const LLFUNCTION_16 = 2 ^ 15
Global Const LLFUNCTION_17 = 2 ^ 16
Global Const LLFUNCTION_18 = 2 ^ 17
Global Const LLFUNCTION_19 = 2 ^ 18
Global Const LLFUNCTION_20 = 2 ^ 19
Global Const LLFUNCTION_21 = 2 ^ 20
Global Const LLFUNCTION_22 = 2 ^ 21
Global Const LLFUNCTION_23 = 2 ^ 22
Global Const LLFUNCTION_24 = 2 ^ 23
Global Const LLFUNCTION_25 = 2 ^ 24
Global Const LLFUNCTION_26 = 2 ^ 25
Global Const LLFUNCTION_27 = 2 ^ 26
Global Const LLFUNCTION_28 = 2 ^ 27
Global Const LLFUNCTION_29 = 2 ^ 28
Global Const LLFUNCTION_30 = 2 ^ 29
Global Const LLFUNCTION_31 = 2 ^ 30
Global Const LLFUNCTION_32 = 2 ^ 31

' Metrics
Global Const LLMETRIC_NEXTVERSION = 1
Global Const LLMETRIC_NAMELEN = 2
Global Const LLMETRIC_COMPANYLEN = 3
Global Const LLMETRIC_MINKEY = 4
Global Const LLMETRIC_MAXKEY = 5
Global Const LLMETRIC_SUPPORTSKEYS = 6
Global Const LLMETRIC_SUPPORTSEXTPACKS = 7
Global Const LLMETRIC_SUPPORTSVARIABLES = 8
Global Const LLMETRIC_SUPPORTSMACHINEID = 9
Global Const LLMETRIC_SUPPORTSNAMEENCODING = 10


' License checking - used by sample applications to perform remote validation
Global Const HTTPCHECKURL = "http://www.gppsoftware.com/LLLIB/LLIBTest.asp"


' Function Prototypes
Declare Function LLSetLibraryLicense Lib "LLIB32D.DLL" (ByVal vsLicense As String) As Long
Declare Sub LLSetLicenseFile Lib "LLIB32D.DLL" (ByVal vsFile As String)
Declare Sub LLGetLicenseFile Lib "LLIB32D.DLL" (ByVal vsFile As String)
Declare Function LLOpenLicenseFile Lib "LLIB32D.DLL" (ByVal vbExclusive As Long) As Long
Declare Sub LLCloseLicenseFile Lib "LLIB32D.DLL" ()
Declare Function LLLockLicenseFile Lib "LLIB32D.DLL" (ByVal vbLock As Long) As Long
Declare Sub LLUnLockLicenseFile Lib "LLIB32D.DLL" ()
Declare Function LLWriteDefaultLicense Lib "LLIB32D.DLL" (ByVal vsOption As String, ByVal vlMod As Long, ByVal vlProduct As Long, ByVal vlProdVersion As Long, ByVal vlUsers As Long, ByVal vlPeriod As Long, ByVal vlFileVer As Long) As Long
Declare Function LLWriteLicense Lib "LLIB32D.DLL" (ByVal vsLicense As String, ByVal vsUser As String, ByVal vsCompany As String) As Long
Declare Function LLWriteLicenseUser Lib "LLIB32D.DLL" (ByVal vsName As String, ByVal vsCompany As String) As Long
Declare Function LLCheckLicenseNo Lib "LLIB32D.DLL" (ByVal vsOptions As String, ByVal vlProduct As Long, ByVal vlProdVersion As Long, ByVal vlKey As Long, ByVal vlMod As Long) As Long
Declare Function LLHttpCheckLicenseNo Lib "LLIB32D.DLL" (ByVal vsUrl As String, ByVal vsLicense As String, ByVal vsMachineId As String, rlRc As Long, ByVal vsMessage As String) As Long
Declare Function LLGetLicense Lib "LLIB32D.DLL" (ByVal vsLicense As String) As Long
Declare Function LLGetLicenseOption Lib "LLIB32D.DLL" (ByVal vsOption As String) As Long
Declare Function LLGetLicenseOptions Lib "LLIB32D.DLL" (rlOptions As Long) As Long
Declare Function LLGetLicenseUser Lib "LLIB32D.DLL" (ByVal vsUser As String) As Long
Declare Function LLGetLicenseCompany Lib "LLIB32D.DLL" (ByVal vsCompany As String) As Long
Declare Function LLGetExpiry Lib "LLIB32D.DLL" (rlDay As Long, rlMonth As Long, rlYear As Long, rlRuns As Long) As Long
Declare Function LLDescribeLicense Lib "LLIB32D.DLL" (ByVal vsMessage As String) As Long

Declare Function LLIsLicenseLimited Lib "LLIB32D.DLL" () As Long
Declare Function LLIsRunLimited Lib "LLIB32D.DLL" () As Long
Declare Function LLIsTimeLimited Lib "LLIB32D.DLL" () As Long
Declare Function LLIsFunctionAvailable Lib "LLIB32D.DLL" (ByVal vlFlag As Long) As Long
Declare Function LLGetMachineId Lib "LLIB32D.DLL" (ByVal vsMachineAs As String) As Long
Declare Sub LLSetMachineIdDrive Lib "LLIB32D.DLL" (ByVal vsDrive As String)
    
Declare Function LLCalcLicenseNo Lib "LLIB32D.DLL" (ByVal vlVersion As Long, ByVal bpType As Long, ByVal vsOption As String, ByVal vlMod As Long, ByVal vlKey As Long, ByVal vlOptions As Long, ByVal vlRuns As Long, ByVal vlDays As Long, ByVal vsExpiry As String, ByVal vsMachineId As String, ByVal vlLicNo As Long, ByVal vlRegNo As Long, ByVal vlUsers As Long, ByVal vsName As String, ByVal vsCompany As String, ByVal vlProduct As Long, ByVal vlPropductVersion As Long, ByVal vlFlags As Long, ByVal vsBuffer As String) As Long
Declare Function LLValidateLicenseNo Lib "LLIB32D.DLL" (ByVal vsOption As String, ByVal vlMod As Long, ByVal vlProduct As Long, ByVal vsLicense As String, ByVal vsUser As String, ByVal vsCompany As String) As Long
    
Declare Sub LLGetLicenseErrorMsg Lib "LLIB32D.DLL" (ByVal vlError As Long, ByVal vsMessage As String)
    
Declare Function LLIncreaseLicenseRuns Lib "LLIB32D.DLL" () As Long
Declare Function LLGetLicenseRuns Lib "LLIB32D.DLL" () As Long
    
Declare Function LLGetLicenseUsers Lib "LLIB32D.DLL" () As Long
Declare Function LLIncreaseLicenseUsers Lib "LLIB32D.DLL" () As Long
Declare Function LLDecreaseLicenseUsers Lib "LLIB32D.DLL" () As Long
Declare Function LLGetFreeLicenseUsers Lib "LLIB32D.DLL" () As Long
Declare Function LLGetUsedLicenseUsers Lib "LLIB32D.DLL" () As Long
Declare Function LLLockUserLicense Lib "LLIB32D.DLL" () As Long
Declare Function LLUnlockUserLicense Lib "LLIB32D.DLL" () As Long


Declare Function LLSetLicenseVariable Lib "LLIB32D.DLL" (ByVal vsVarName As String, ByVal vsValue As String) As Long
Declare Function LLGetLicenseVariable Lib "LLIB32D.DLL" (ByVal vsVarName As String, ByVal vsValue As String) As Long
Declare Function LLGetLicenseVariableName Lib "LLIB32D.DLL" (ByVal vlIndex As Long, ByVal vsVarName As String) As Long

Declare Function LLGetLicenseVersionRange Lib "LLIB32D.DLL" (riLo As Long, riHi As Long) As Long
Declare Function LLGetLicenseVersion Lib "LLIB32D.DLL" (riVersion As Long) As Long
Declare Function LLSetLicenseVersion Lib "LLIB32D.DLL" (ByVal vlVersion As Long) As Long

Declare Function LLWriteExtensionPack Lib "LLIB32D.DLL" (ByVal vsPack As String) As Long
Declare Function LLDeleteExtensionPack Lib "LLIB32D.DLL" (ByVal vlIndex As Long) As Long
Declare Function LLGetExtensionPack Lib "LLIB32D.DLL" (ByVal vlIndex As Long, ByVal vsPack As String) As Long

Declare Function LLGetMetric Lib "LLIB32D.DLL" (ByVal vlMetric As Long, ByVal vlP1 As Long) As Long

Function GetLicenseErrorMsg(ByVal vlError As Long) As String
    Dim lsString As String
    lsString = String$(255, 0)
    LLGetLicenseErrorMsg vlError, lsString
    GetLicenseErrorMsg = StrToSz(lsString)
End Function

Function GetLicenseDescription() As String
    Dim liRc As Integer
    Dim lsString As String
    lsString = String$(255, 0)
    liRc = LLDescribeLicense(lsString)
    GetLicenseDescription = StrToSz(lsString)
End Function

Function GetMachineIdentifier() As String
    Dim liRc As Integer
    Dim lsString As String
    lsString = String$(LLMAXMACHINEIDLEN + 1, 0)
    liRc = LLGetMachineId(lsString)
    GetMachineIdentifier = StrToSz(lsString)
End Function
Function FormatLicense(ByVal LicenseNumber As String) As String
    Dim i As Long
    Dim s As String
    LicenseNumber = Replace(LicenseNumber, "-", "")
    For i = 1 To Len(LicenseNumber)
        s = s & Mid(LicenseNumber, i, 1)
        If i Mod 5 = 0 Then s = s & "-"
    Next
    If Right(s, 1) = "-" Then s = Mid(s, 1, Len(s) - 1)
    FormatLicense = s
End Function


Function GetLicense() As String
    Dim liRc As Integer
    Dim lsString As String
    
    lsString = String$(LLMAXLICENSELEN + 1, 0)
    liRc = LLGetLicense(lsString)
    lsString = StrToSz(lsString)
    
    GetLicense = FormatLicense(lsString)
    
    
End Function


Function GetLicenseOption() As String
    Dim liRc As Integer
    Dim lsString As String
    lsString = String$(1, 0)
    liRc = LLGetLicenseOption(lsString)
    GetLicenseOption = StrToSz(lsString)
End Function

Function GetLicenseUser() As String
    Dim liRc As Integer
    Dim lsString As String
    lsString = String$(LLMAXUSERNAMELEN + 1, 0)
    liRc = LLGetLicenseUser(lsString)
    GetLicenseUser = StrToSz(lsString)
End Function






        
        
Function GetLicenseCompany() As String
    Dim liRc As Integer
    Dim lsString As String
    lsString = String$(LLMAXCOMPANYNAMELEN + 1, 0)
    liRc = LLGetLicenseCompany(lsString)
    GetLicenseCompany = StrToSz(lsString)
End Function

Function StrToSz(ByVal vsString As String) As String

    StrToSz = vsString

    If InStr(1, vsString, Chr(0)) <> 0 Then
        StrToSz = left$(vsString, InStr(1, vsString, Chr$(0)) - 1)
    End If

End Function



Sub ReleaseLicense()
On Error Resume Next
    Call LLUnlockUserLicense
    Call LLCloseLicenseFile
End Sub




Function AcquireLicense(FileName As String, Product As Long, Modulus As Long) As Long
    
    Dim rc    As Integer
    Dim day   As Long
    Dim month As Long
    Dim year  As Long
    Dim runs  As Long
    Dim Msg   As String
    
    Product = -1
    If Product = -1 Then
        'not licensed
        AcquireLicense = True
        Exit Function
    End If
    
    AcquireLicense = False
    
    On Error Resume Next
    Kill StripExtension(FileName) & "*.USR"
    On Error GoTo 0
    
    rc = LLSetLibraryLicense("R110050010000267")
    Call LLSetLicenseFile(FileName)
    Call LLSetMachineIdDrive(FileDrive(FileName))
    
    
    rc = LLOpenLicenseFile(False)
    Select Case rc
        Case LLER_OK
            rc = LLCheckLicenseNo("ULTR", Product, 0, 0, Modulus)
            Select Case rc
                Case LLER_OK
                    rc = LLGetExpiry(day, month, year, runs)
                    Select Case rc
                        Case LLER_OK
                            AcquireLicense = True 'LockUser(FileName, Product, Modulus)
                        
                        Case LLER_EXPIRESONDATE
                            If DateSerial(year, month, day) > Now Then
                                AcquireLicense = LockUser(FileName, Product, Modulus)
                            Else
                                If FLicense.EnterLicense(FileName, Product, Modulus, _
                                    "Your trial period has expired. To continue using this software" & vbCrLf & _
                                    "you must enter a registered activation code." & vbCrLf & vbCrLf & _
                                    "To obtain an activation code, contact the HomeFront sales" & vbCrLf & _
                                    "department at 1-800-887-0257. Enter your client id, company" & vbCrLf & _
                                    "name and activation code exactly as you received them from" & vbCrLf & _
                                    "the HomeFront activations department.") Then AcquireLicense = AcquireLicense(FileName, Product, Modulus)
                            End If
                            
                        Case LLER_EXPIRESONRUNS
                            If runs > 0 Then
                                AcquireLicense = LockUser(FileName, Product, Modulus)
                            Else
                                If FLicense.EnterLicense(FileName, Product, Modulus, _
                                    "Your trial period has expired. To continue using this software" & vbCrLf & _
                                    "you must enter a registered activation code." & vbCrLf & vbCrLf & _
                                    "To obtain an activation code, contact the HomeFront sales" & vbCrLf & _
                                    "department at 1-800-887-0257. Enter your client id, company" & vbCrLf & _
                                    "name and activation code exactly as you received them from" & vbCrLf & _
                                    "the HomeFront activations department.") Then AcquireLicense = AcquireLicense(FileName, Product, Modulus)
                            End If
                    End Select
                Case Else
                    Call VBA.MsgBox(GetLicenseErrorMsg(rc), vbCritical, "Homefront")
                
            End Select
        Case LLER_LICENSEFILEMISSING
            If FLicense.EnterLicense(FileName, Product, Modulus, _
                "Your license file could not be found. Your software has been" & vbCrLf & _
                "locked. To reactivate it please contact the Homefront Software" & vbCrLf & _
                "sales department at 1-800-887-0257. Enter your client id, company" & vbCrLf & _
                "name and activation code exactly as you received them from" & vbCrLf & _
                "the HomeFront activations department.") Then AcquireLicense = AcquireLicense(FileName, Product, Modulus)
                
        Case Else
            Call VBA.MsgBox(GetLicenseErrorMsg(rc), vbCritical, "Homefront")
            
    End Select
    
    
End Function

Private Function LockUser(FileName As String, Product As Long, Modulus As Long) As Boolean
    Dim rc As Long
    LockUser = False
    If LLGetFreeLicenseUsers() > 0 Then
        rc = LLLockUserLicense
        If rc = 0 Then
            Call LLIncreaseLicenseRuns
            LockUser = True
        Else
            Call VBA.MsgBox(GetLicenseErrorMsg(rc), vbCritical, "Homefront")
        End If
    Else
        Call FLicense.EnterLicense(FileName, Product, Modulus, _
            "All " & LLGetLicenseUsers() & " available licenses are in use." & vbCrLf & vbCrLf & _
            "To purchase additional user licenses, contact the Homefront Software" & vbCrLf & _
            "sales department at 1-800-887-0257. Enter your activation" & vbCrLf & _
            "code, company name and client id exactly as you received" & vbCrLf & _
            "them from Homefront Software.")
    End If
End Function


Function WriteLicense(FileName As String, _
                      Product As Long, _
                      Modulus As Long, _
                      LicenseNumber As String, _
                      UserName As String, _
                      CompanyName As String, _
                      CreateFile As Boolean) As Boolean
    Dim rc As Long
                  
    WriteLicense = True
                      
    Call LLSetLibraryLicense("R110050010000267")
    Call LLSetLicenseFile(FileName)
    
    Call LLSetMachineIdDrive(FileDrive(FileName))
    
    
    If CreateFile Then
        Call CreatePath("License Path", FilePath(FileName))
        rc = LLWriteDefaultLicense("U", Modulus, Product, 0, 5, 30, 300)
    End If
    If rc <> 0 And rc <> 16 Then
        WriteLicense = False
        VBA.MsgBox GetLicenseErrorMsg(rc), vbExclamation, "Homefront"
    Else
        rc = LLOpenLicenseFile(0)
        If rc <> 0 Then
            WriteLicense = False
            VBA.MsgBox GetLicenseErrorMsg(rc), vbExclamation, "Homefront"
        Else
            If Trim(LicenseNumber) <> "" Then
                rc = LLWriteLicense(Replace(LicenseNumber, "-", ""), UserName, CompanyName)
                If rc <> 0 Then
                    WriteLicense = False
                    VBA.MsgBox GetLicenseErrorMsg(rc), vbExclamation, "Homefront"
                End If
            ElseIf Trim(UserName) <> "" Or Trim(CompanyName) <> "" Then
                rc = LLWriteLicenseUser(UserName, CompanyName)
                If rc <> 0 Then
                    WriteLicense = False
                    VBA.MsgBox GetLicenseErrorMsg(rc), vbExclamation, "Homefront"
                End If
            End If
        End If
    End If
    
    Call LLCloseLicenseFile
    
End Function



