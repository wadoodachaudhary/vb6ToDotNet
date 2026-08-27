Attribute VB_Name = "MMain"
Option Explicit
Option Compare Text
Private Const SRCFILE = "MMain::"
Public Const MARKUPPRECISION = 2
Public ProjectBased As Boolean
Public ProjectPhaseBased As Boolean
Public TakeoffSystem As TakeoffSystems

Public IsMultiFamily As Boolean
Public PostPOQtyToAccounting As Boolean

Public TakeoffParameters As Collection
Public LoadedForm As String
'file menu
Public Const mcFILE_SAVE = 1
Public Const mcFILE_CHANGEDIV = 3
Public Const mcFILE_CLOSE = 4

'Global gString As String
Global UsingToolsClass As Boolean ' set=true on all tool actions. when true sub main() is not run.

Public Enum EstimatingSystems
    esNone = 0
    esPipeline = 1
    esTimberline = 2
End Enum



'view menu
Public Const mcVIEW_TASKPANEL = 0
Public Const mcVIEW_WORKFLOW = 1

'for shared menus
Public MouseForm  As Form
Public MouseCtrl  As Control
Public MouseCol   As Long
Public Const SW_SHOWNORMAL = 1
Public Const SE_ERR_FNF = 2&
Public Const SE_ERR_PNF = 3&
Public Const SE_ERR_ACCESSDENIED = 5&
Public Const SE_ERR_OOM = 8&
Public Const SE_ERR_DLLNOTFOUND = 32&
Public Const SE_ERR_SHARE = 26&
Public Const SE_ERR_ASSOCINCOMPLETE = 27&
Public Const SE_ERR_DDETIMEOUT = 28&
Public Const SE_ERR_DDEFAIL = 29&
Public Const SE_ERR_DDEBUSY = 30&
Public Const SE_ERR_NOASSOC = 31&
Public Const ERROR_BAD_FORMAT = 11&

' Used to indicate what to enumerate
Private Const PRINTER_ENUM_DEFAULT         As Long = &H1

Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Enum SortAlgorithms
    InsertSort
    MergeSort
    QuickSOrt
    SelectionSort
End Enum

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

' Structure used to obtain the data from Windows.
Private Type PRINTER_INFO_2
   pServerName As Long
   pPrinterName As Long
   pShareName As Long
   pPortName As Long
   pDriverName As Long
   pComment As Long
   pLocation As Long
   pDevMode As Long 'DEVMODE
   pSepFile As Long
   pPrintProcessor As Long
   pDatatype As Long
   pParameters As Long
   pSecurityDescriptor As Long 'SECURITY_DESCRIPTOR
   Attributes As Long
   Priority As Long
   DefaultPriority As Long
   StartTime As Long
   UntilTime As Long
   Status As Long
   cJobs As Long
   AveragePPM As Long
End Type

' Printer status flags used with PRINTER_INFORMATION_2
Private Const PRINTER_STATUS_READY              As Long = &H0
Private Const PRINTER_STATUS_PAUSED             As Long = &H1
Private Const PRINTER_STATUS_ERROR              As Long = &H2
Private Const PRINTER_STATUS_PENDING_DELETION   As Long = &H4
Private Const PRINTER_STATUS_PAPER_JAM          As Long = &H8
Private Const PRINTER_STATUS_PAPER_OUT          As Long = &H10
Private Const PRINTER_STATUS_MANUAL_FEED        As Long = &H20
Private Const PRINTER_STATUS_PAPER_PROBLEM      As Long = &H40
Private Const PRINTER_STATUS_OFFLINE            As Long = &H80
Private Const PRINTER_STATUS_IO_ACTIVE          As Long = &H100
Private Const PRINTER_STATUS_BUSY               As Long = &H200
Private Const PRINTER_STATUS_PRINTING           As Long = &H400
Private Const PRINTER_STATUS_OUTPUT_BIN_FULL    As Long = &H800
Private Const PRINTER_STATUS_NOT_AVAILABLE      As Long = &H1000
Private Const PRINTER_STATUS_WAITING            As Long = &H2000
Private Const PRINTER_STATUS_PROCESSING         As Long = &H4000
Private Const PRINTER_STATUS_INITIALIZING       As Long = &H8000
Private Const PRINTER_STATUS_WARMING_UP         As Long = &H10000
Private Const PRINTER_STATUS_TONER_LOW          As Long = &H20000
Private Const PRINTER_STATUS_NO_TONER           As Long = &H40000
Private Const PRINTER_STATUS_PAGE_PUNT          As Long = &H80000
Private Const PRINTER_STATUS_USER_INTERVENTION  As Long = &H100000
Private Const PRINTER_STATUS_OUT_OF_MEMORY      As Long = &H200000
Private Const PRINTER_STATUS_DOOR_OPEN          As Long = &H400000
Private Const PRINTER_STATUS_SERVER_UNKNOWN     As Long = &H800000
Private Const PRINTER_STATUS_POWER_SAVE         As Long = &H1000000

Public Enum PrinterStatusCodes
   psReady = PRINTER_STATUS_READY
   psPaused = PRINTER_STATUS_PAUSED
   psError = PRINTER_STATUS_ERROR
   psPendingDeletion = PRINTER_STATUS_PENDING_DELETION
   psPaperJam = PRINTER_STATUS_PAPER_JAM
   psPaperOut = PRINTER_STATUS_PAPER_OUT
   psManualFeed = PRINTER_STATUS_MANUAL_FEED
   psPaperProblem = PRINTER_STATUS_PAPER_PROBLEM
   psOffline = PRINTER_STATUS_OFFLINE
   psIoActive = PRINTER_STATUS_IO_ACTIVE
   psBusy = PRINTER_STATUS_BUSY
   psPrinting = PRINTER_STATUS_PRINTING
   psOutputBinFull = PRINTER_STATUS_OUTPUT_BIN_FULL
   psNotAvailable = PRINTER_STATUS_NOT_AVAILABLE
   psWaiting = PRINTER_STATUS_WAITING
   psProcessing = PRINTER_STATUS_PROCESSING
   psInitializing = PRINTER_STATUS_INITIALIZING
   psWarmingUp = PRINTER_STATUS_WARMING_UP
   psTonerLow = PRINTER_STATUS_TONER_LOW
   psNoToner = PRINTER_STATUS_NO_TONER
   psPagePrint = PRINTER_STATUS_PAGE_PUNT
   psUserIntervention = PRINTER_STATUS_USER_INTERVENTION
   psOutOfMemory = PRINTER_STATUS_OUT_OF_MEMORY
   psDoorOpen = PRINTER_STATUS_DOOR_OPEN
   psServerUnknown = PRINTER_STATUS_SERVER_UNKNOWN
End Enum
   
' Printer attribute flags used with PRINTER_INFORMATION_2
Private Const PRINTER_ATTRIBUTE_QUEUED            As Long = &H1
Private Const PRINTER_ATTRIBUTE_DIRECT            As Long = &H2
Private Const PRINTER_ATTRIBUTE_DEFAULT           As Long = &H4
Private Const PRINTER_ATTRIBUTE_SHARED            As Long = &H8
Private Const PRINTER_ATTRIBUTE_NETWORK           As Long = &H10
Private Const PRINTER_ATTRIBUTE_HIDDEN            As Long = &H20
Private Const PRINTER_ATTRIBUTE_LOCAL             As Long = &H40
Private Const PRINTER_ATTRIBUTE_ENABLE_DEVQ       As Long = &H80
Private Const PRINTER_ATTRIBUTE_KEEPPRINTEDJOBS   As Long = &H100
Private Const PRINTER_ATTRIBUTE_DO_COMPLETE_FIRST As Long = &H200
Private Const PRINTER_ATTRIBUTE_WORK_OFFLINE      As Long = &H400
Private Const PRINTER_ATTRIBUTE_ENABLE_BIDI       As Long = &H800
Private Const PRINTER_ATTRIBUTE_RAW_ONLY          As Long = &H1000
Private Const PRINTER_ATTRIBUTE_PUBLISHED         As Long = &H2000

Public Enum PrinterAttributeCodes
   paQueued = PRINTER_ATTRIBUTE_QUEUED
   paDirect = PRINTER_ATTRIBUTE_DIRECT
   paDefault = PRINTER_ATTRIBUTE_DEFAULT
   paShared = PRINTER_ATTRIBUTE_SHARED
   paNetwork = PRINTER_ATTRIBUTE_NETWORK
   paHidden = PRINTER_ATTRIBUTE_HIDDEN
   paLocal = PRINTER_ATTRIBUTE_LOCAL
   paEnableDevQ
End Enum


' Some calls need to know OS
Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

' Platform ID constants
Private Const VER_PLATFORM_WIN32s As Long = &H0
Private Const VER_PLATFORM_WIN32_WINDOWS As Long = &H1
Private Const VER_PLATFORM_WIN32_NT As Long = &H2

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As Any) As Long

'grid menus
Public Const mcGRID_GROUP = 0
Public Const mcGRID_EXPANDALL = 1
Public Const mcGRID_COLLAPSEALL = 2
Public Const mcGRID_HIDE = 4
Public Const mcGRID_INSERT = 5
Public Const mcGRID_RENAME = 6
Public Const mcGRID_PRINT = 8
Public Const mcGRID_SAVEAS = 9
Public Const mcGRID_ATTACHMENTS = 10

Public Const BUFFER_SIZE = 32767


Public Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Public Const SEM_FAILCRITICALERRORS = &H1
Public Const SEM_NOOPENFILEERRORBOX = &H8000&
Public Const SEM_NOGPFAULTERRORBOX = &H2
Public Const SEM_NOERRORS = SEM_FAILCRITICALERRORS Or SEM_NOOPENFILEERRORBOX Or SEM_NOGPFAULTERRORBOX

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Public Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long
Public Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function CoCreateGuid Lib "ole32.dll" (pguid As Guid) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ReleaseCapture& Lib "user32" ()

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpszOp As String, ByVal lpszFile As String, ByVal lpszParams As String, ByVal lpszDir As String, ByVal FsShowCmd As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function PathIsUNC Lib "shlwapi.dll" Alias "PathIsUNCA" (ByVal pszPath As String) As Long
Private Declare Function NetShareGetInfo Lib "NETAPI32" (ByRef ServerName As Byte, ByRef NetName As Byte, ByVal Level As Long, ByRef Buffer As Long) As Long
Private Declare Function PtrToInt Lib "kernel32" Alias "lstrcpynW" (RetVal As Any, ByVal ptr As Long, ByVal nCharCount As Long) As Long
Private Declare Function PtrToStr Lib "kernel32" Alias "lstrcpyW" (RetVal As Byte, ByVal ptr As Long) As Long
Private Declare Function StrLen Lib "kernel32" Alias "lstrlenW" (ByVal ptr As Long) As Long
Private Declare Function NetAPIBufferFree Lib "netapi32.dll" Alias "NetApiBufferFree" (bufptr As Any) As Long
Public Declare Function PathIsNetworkPath Lib "shlwapi.dll" Alias "PathIsNetworkPathA" (ByVal pszPath As String) As Long
Public Declare Function PathStripToRoot Lib "shlwapi.dll" Alias "PathStripToRootA" (ByVal pPath As String) As Long
Public Declare Function PathSkipRoot Lib "shlwapi.dll" Alias "PathSkipRootA" (ByVal pPath As String) As Long
Public Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Public Declare Function lstrcpyA Lib "kernel32" (ByVal RetVal As String, ByVal ptr As Long) As Long

Public Declare Function GetComputerName& Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long)
Public Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As Any, ByVal lpDefaultValue As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'pricing worksheet grid
'Public Const mcPRICE_COPY2COMMUNITY = 0

Public Const mcPRICE_REFRESHCOSTS = 1
Public Const mcPRICE_ADJUSTPRICES = 2
Public Const mcPRICE_REMOVEASSEMBLY = 4

Type POINTAPI
        X As Long
        Y As Long
End Type

Public Enum OptionLocations
    olModel
    olScheduleB
    olChangeOrder
    olDesignCenter
    olDeletedItems
End Enum

Public Enum AssemblyStatuses
    asOpen
    asPriceOnly
    asClosed
End Enum

Public Enum AssemblyTypes
    atModel = 0
    atElevaton = 1
    atoption = 2
    atGlobal = 3
    atDesignCenter = 4
    atCustomOption = -1
End Enum

Public Enum OptionTypes
    otModel = 0
    otOption = 1
    otGlobal = 2
    otDesignCenter = 3
End Enum

Public Enum DeliveryTypes
    dtPrint
    dtEmail
    dtFax
End Enum

Public Enum CostBasisTypes
    cbCurrent
    cbNext1
    cbNext2
    cbLast1
    cbLast2
    cbLast3
    cbForecast1
    cbForecast2
    cbForecast3
    cbForecast4
    cbForecast5
    cbForecast6
    cbForecast7
    cbForecast8
    cbForecast9
    cbForecast10
    cbForecast11
    cbForecast12
End Enum
    
    
Public Enum ZeroQtyTakeoffModes
    ztPrompt = 0
    ztIgnore = 1
    ztAccept = 2
End Enum

Public Enum MultiStateEnum
    msNone = 0
    msSome = 1
    msAll = 2
End Enum

Private Type MungeLong
    X As Long
    Dummy As Integer
End Type

Private Type MungeInt
    XLo As Integer
    XHi As Integer
    Dummy As Integer
End Type

Private Type Guid
   Data1 As Long
   Data2 As Long
   Data3 As Long
   Data4(8) As Byte
End Type

Public Const Delete = &H10000
Public Const READ_CONTROL = &H20000
Public Const WRITE_DAC = &H40000
Public Const WRITE_OWNER = &H80000
Public Const SYNCHRONIZE = &H100000
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Public Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const SPECIFIC_RIGHTS_ALL = &HFFFF
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))

Public Const ERROR_SUCCESS = 0&
Public Const ERROR_MORE_DATA = 234
Public Const ERROR_NO_MORE_DATA = 259

Public Const REG_SZ = 1                          ' Unicode nul terminated string

' Win32 API declares
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrn As Long, pDefault As Any) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrn As Long) As Long
Private Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal cbBuf As Long, pcbNeeded As Long) As Long
Private Declare Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal Command As Long) As Long
Private Declare Function EnumPrinters Lib "winspool.drv" Alias "EnumPrintersA" (ByVal Flags As Long, ByVal Name As String, ByVal Level As Long, pPrinterEnum As Any, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Private Declare Function PrinterProperties Lib "winspool.drv" (ByVal hwnd As Long, ByVal hPrinter As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long


Private Declare Function GetDefaultPrinter Lib "winspool.drv" Alias "GetDefaultPrinterA" (ByVal pszBuffer As String, pcchBuffer As Long) As Long
Private Declare Function SetDefaultPrinter Lib "winspool.drv" Alias "SetDefaultPrinterA" (ByVal pszPrinter As String) As Long


' Clipboard Manager Functions
Public Declare Function EmptyClipboard Lib "user32" () As Long
Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CloseClipboard Lib "user32" () As Long
Public Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Public Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Public Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long

'Public Declare Function RegOpenCurrentUser Lib "advapi32.dll" (ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter AS String, you must pass it By Value.
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter AS String, you must pass it By Value.
Public Declare Function RegRegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Public Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As Any) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, ByVal lpType As Long, ByVal lpData As Long, ByVal lpcbData As Long) As Long

' Reg result codes
'
Public Const REG_CREATED_NEW_KEY = &H1                       ' New Registry Key created
Public Const REG_OPENED_EXISTING_KEY = &H2                       ' Existing Key opened
'
' Reg Create Type Values...
'
Public Const REG_OPTION_RESERVED = 0            ' Parameter is reserved
Public Const REG_OPTION_NON_VOLATILE = 0        ' Key is preserved when system is rebooted
Public Const REG_OPTION_VOLATILE = 1            ' Key is not preserved when system is rebooted
Public Const REG_OPTION_CREATE_LINK = 2         ' Created key is a symbolic link
Public Const REG_OPTION_BACKUP_RESTORE = 4      ' open for backup or restore

Public HFApp  As New HFSystem.Application
Public MaxPOAmount As Double



Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SetFocusAPI Lib "user32.dll" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Sub SetTopMostWindow(hwnd As Long, Topmost As Boolean)

   If Topmost = True Then 'Make the window topmost
      Call SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
   Else
      Call SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
   End If
End Sub



Public Sub SetWindowFocus(hwnd As Long)
'Private Declare Function AttachThreadInput Lib "user32.dll" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
'Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As Long, ByRef lpdwProcessId As Long) As Long
   
 
 '   Dim lThreadID As Long
    
  '  lThreadID = GetWindowThreadProcessId(hwnd, ByVal 0&)
   ' AttachThreadInput lThreadID, App.ThreadID, True
    SetFocusAPI hwnd
    'SendKeys "{DOWN}"
    'AttachThreadInput lThreadID, App.ThreadID, False

End Sub

  
Public Sub Main()

'    If Not UsingToolsClass Then
'
'        'normal app launch
'        InitCommonControls
'        If HFApp.Login(App.Title, App.ProductName, App.Major & "." & App.Minor , App.Path) Then
'            Call ReadUserPermissions
'            FMain.Show
'            Call RollPrices(True)
'        End If
'
'    End If

'ACTIVEX EXE WONT RETURN AN OBJECT UNTIL THIS PROC COMPLETES ...
'I DONT WANT IT TO DO ANYTHING UNLESS IT IS A NORMAL APP START.
'SO I CREATED THIS STARTUP HELPER SCREEN THAT IS NOTHING BUT A TIMER
'WHICH RUNS THE NORMAL STARTUP CODE. SEEMS TO WORK.
    
    Load FStartupHelper

    'Dim X As New HFEstTools
    'Call X.CreateSinglePO(1, "homefront", "admin", "02001", "2200", "02001/032")

End Sub


Public Sub ReadUserPermissions()
On Error Resume Next
   MaxPOAmount = HFApp.SqlExec("select MaxPOAmount from user_manager where user_id=" & DbQuote(Str, HFApp.LoginID))(0)
End Sub

Public Function SaveToCSV(FileName As String, Optional Sheet As Integer = 1) As String
'    Dim s As String
'    Dim xlSheet As Object 'Excel.Worksheet
'    Set xlSheet = GetObject(FileName).Sheets(Sheet)
'    s = TempFile("csv")
'    Kill s
'    Call xlSheet.SaveAs(s, 6, , , , , False)
'    Set xlSheet = Nothing
'    SaveToCSV = s


    Dim s As String
    Dim xlApp As Object 'New Excel.Application
    Dim xlWkbk As Object 'Excel.Workbook
    Dim xlSheet As Object 'Excel.Worksheet
    
    Set xlApp = CreateObject("Excel.Application")
    
    Set xlWkbk = xlApp.Workbooks.Open(FileName)
    
    On Error Resume Next
    Set xlSheet = xlWkbk.Sheets(Sheet)
    If Err.Number = 9 Then
        On Error GoTo 0
        Err.Raise 9, "SaveToCSV", "Worksheet number " & Sheet & " not found."
    End If
    On Error GoTo 0
    
    s = TempFile("csv")
    On Error Resume Next
    Kill s
    On Error GoTo 0
    Call xlSheet.SaveAs(s, 6, , , , , False)
    
    Call xlWkbk.Close(False)
    Set xlSheet = Nothing
    Set xlWkbk = Nothing
    Set xlApp = Nothing
    
    SaveToCSV = s

End Function

Public Function LoadExcelSheet(Optional Workbook As Object, Optional SheetName As String, Optional Grid As Object, Optional CSVFile As String)
    Dim i As Long
    Dim s As String
    Dim xlSheet As Object 'Excel.Worksheet

    'had lots of problems having the grid load xls file directly.
    'goofy inconsistant errors galore. found that the safest way
    'was to save excel sheet to csv file then load that.
    
    On Error Resume Next
    
    If CSVFile = "" Then
        If SheetName = "" Then
            Set xlSheet = Workbook.Sheets(1)
        Else
            Set xlSheet = Workbook.Sheets(SheetName)
        End If
        
        On Error GoTo 0
        If xlSheet Is Nothing Then
            Err.Raise vbObjectError + 88, "LoadExcelSheet", SheetName
        End If
        
        s = TempFile("csv")
        On Error Resume Next
        Kill s
        Call xlSheet.SaveAs(s, 6, , , , , False)
        Set xlSheet = Nothing
    Else
        s = CSVFile
    End If
    
    'load csv file into grid then delete the temp file
    Call Grid.LoadGrid(s, flexFileCommaText)
    On Error Resume Next
    Kill s
    
    'write original row and column numbers into grid.headings
    With Grid
        For i = 1 To .Rows - 1
            .TextMatrix(i, 0) = i
        Next
        For i = 1 To .Cols - 1
            .TextMatrix(0, i) = FormatExcelColRef(i)
        Next
        .ColWidth(0) = 360
        .ColAlignment(0) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
    End With
    
End Function


Public Function FormatExcelColRef(Index As Long) As String
    Dim i As Long
    Dim s As String
    Const abc = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    
    If Index > 26 Then
        i = (Index - 1) \ 26
        If i > 0 Then
            s = s & Mid(abc, i, 1)
        End If
    End If
    
    i = Index Mod 26
    If i = 0 Then
        s = s & "Z"
    Else
        s = s & Mid(abc, i, 1)
    End If
    FormatExcelColRef = s
End Function

Public Sub PublishPricingWorksheet(Worksheet As Long)
On Error GoTo eh
    Dim i  As Long
    Dim s  As String
    Dim rs As Recordset
    Dim AssemblyType As AssemblyTypes
    Dim community As String
    Dim CommunityPhase As String
    

    s = ""
    s = s & "select sd.assemblytype" & vbCrLf
    s = s & "      ,sd.community" & vbCrLf
    s = s & "      ,sd.CommunityPhase" & vbCrLf
    s = s & "      ,sd.OptionID" & vbCrLf
    s = s & "      ,sd.model" & vbCrLf
    s = s & "      ,sd.JCExtra" & vbCrLf
    s = s & "      ,sd.elevation" & vbCrLf
    s = s & "      ,sd.series" & vbCrLf
    s = s & "      ,sd.description" & vbCrLf
    s = s & "      ,sd.AssemblyUOM" & vbCrLf
    s = s & "      ,sd.floorarea" & vbCrLf
    s = s & "      ,sd.bedrooms" & vbCrLf
    s = s & "      ,sd.bathrooms" & vbCrLf
    s = s & "      ,sd.style" & vbCrLf
    s = s & "      ,sd.cost" & vbCrLf
    s = s & "      ,sd.pretax" & vbCrLf
    s = s & "      ,sd.copretax" & vbCrLf
    s = s & "      ,sd.IncludeTax" & vbCrLf
    s = s & "      ,sd.IncentiveCost" & vbCrLf
    s = s & "      ,sd.IncentiveRetail" & vbCrLf
    s = s & "      ,sd.tax" & vbCrLf
    s = s & "      ,sd.cotax" & vbCrLf
    s = s & "      ,sd.optionid" & vbCrLf
    s = s & "      ,sd.constcutoff" & vbCrLf
    s = s & "      ,sd.assembly" & vbCrLf
    s = s & "      ,sd.category" & vbCrLf
    
    s = s & "      ,sd.GraphicPath,sd.SpecDocument,sd.MaxWidth,sd.MaxLength,sd.Inactive" & vbCrLf
    
    s = s & "      ,sd.color" & vbCrLf
    s = s & "      ,sd.location" & vbCrLf
    s = s & "      ,sd.qty" & vbCrLf
    s = s & "      ,sd.includedoption" & vbCrLf
    
    
    s = s & "      ,oc.group_code" & vbCrLf
    s = s & "      ,sd.notes" & vbCrLf
    s = s & "      ,sd.comments" & vbCrLf
    s = s & "      ,sd.Markup" & vbCrLf
    s = s & "      ,sd.Margin" & vbCrLf
    For i = 2 To 10
        s = s & "      ,sd.pretax" & i & vbCrLf
        s = s & "      ,sd.tax" & i & vbCrLf
        s = s & "      ,sd.Margin" & i & vbCrLf
        s = s & "      ,sd.Markup" & i & vbCrLf
    Next
    
    s = s & "      ,sd.ColorListID" & vbCrLf
    s = s & "      ,sd.StyleListID" & vbCrLf
    s = s & "      ,sd.FinishListID" & vbCrLf
    s = s & "      ,sd.OtherListID" & vbCrLf
    s = s & "      ,sd.StyleValue" & vbCrLf
    s = s & "      ,sd.FinishValue" & vbCrLf
    s = s & "      ,sd.OtherValue" & vbCrLf
    s = s & "      ,m.Worksheet" & vbCrLf
    
    s = s & "      ,sd.DesignCenterSalesOnly" & vbCrLf
    s = s & "      ,sd.SelectByRoom" & vbCrLf
    s = s & "      ,sd.DisplayTotalOnly" & vbCrLf
    s = s & "  from tblSalesSheetDetails sd JOIN tblSalesSheetMaster m on m.Worksheet=sd.Worksheet  LEFT OUTER JOIN tblCategories oc ON sd.Category=oc.Category" & vbCrLf
    s = s & " where sd.Worksheet=" & DbQuote(Num, Worksheet) & vbCrLf
    's = s & " where sd.ReadytoPublish = 1 and sd.Worksheet=" & DbQuote(Num, Worksheet) & vbCrLf
    
    
    Set rs = HFApp.SqlExec(s)
    While Not rs.EOF
        AssemblyType = Val("" & rs("AssemblyType"))
        
        Select Case AssemblyType

            Case atModel
                
                s = ""
                s = s & "INSERT INTO tblModels(DivisionID,Area,CommunityPhase,Model,Series,Elevation,Inactive)" & vbCrLf
                s = s & "VALUES(" & HFApp.DivisionID & "," & DbQuote(Str, "" & rs("Community")) & vbCrLf
                s = s & "      ," & DbQuote(Str, "" & rs("CommunityPhase")) & vbCrLf
                s = s & "      ," & DbQuote(Str, "" & rs("Model")) & vbCrLf
                s = s & "      ," & DbQuote(Str, "" & rs("Series")) & vbCrLf
                s = s & "      ," & DbQuote(Str, "" & rs("Elevation")) & ",0)"
                Call HFApp.SqlExec(s)
                
                s = ""
                s = s & "UPDATE tblModels" & vbCrLf
                s = s & "SET LastSalesWorksheet=SalesWorksheet" & vbCrLf
                s = s & "   ,SalesWorksheet=" & DbQuote(Num, Worksheet) & vbCrLf
                s = s & "   ,Assembly=" & DbQuote(Str, "" & rs("Assembly")) & vbCrLf
                s = s & "   ,Description=" & DbQuote(Str, "" & rs("Description")) & vbCrLf
                s = s & "   ,model_picture=" & DbQuote(Str, "" & rs("GraphicPath")) & vbCrLf
                s = s & "   ,Spec_Document=" & DbQuote(Str, "" & rs("SpecDocument")) & vbCrLf
                s = s & "   ,Max_Width=" & DbQuote(Num, "" & rs("MaxWidth")) & vbCrLf
                s = s & "   ,Max_Length=" & DbQuote(Num, "" & rs("MaxLength")) & vbCrLf
                s = s & "   ,Inactive=" & DbQuote(Bit, "" & rs("Inactive")) & vbCrLf
                s = s & "   ,Comments=" & DbQuote(Str, "" & rs("Comments")) & vbCrLf
                s = s & "   ,Style=" & DbQuote(Str, "" & rs("Style")) & vbCrLf
                s = s & "   ,NoOfBedrooms=" & DbQuote(Str, "" & rs("Bedrooms")) & vbCrLf
                s = s & "   ,NoOfBathrooms=" & DbQuote(Str, "" & rs("Bathrooms")) & vbCrLf
                s = s & "   ,ModelSize=" & DbQuote(Num, "" & rs("FloorArea")) & vbCrLf
                s = s & "   ,IncentiveCost=" & DbQuote(Cur, "" & rs("IncentiveCost")) & vbCrLf
                s = s & "   ,IncentiveRetail=" & DbQuote(Cur, "" & rs("IncentiveRetail")) & vbCrLf
                s = s & "   ,Cost_amount=" & DbQuote(Cur, "" & rs("Cost")) & vbCrLf
                s = s & "   ,base_house=" & DbQuote(Cur, "" & rs("Pretax")) & vbCrLf
                s = s & "   ,Margin=" & DbQuote(Num, "" & rs("Margin")) & vbCrLf
                s = s & "   ,Markup=" & DbQuote(Num, "" & rs("Markup")) & vbCrLf
                If CBool("" & rs("IncludeTax")) Then
                    s = s & "   ,UseTax=1" & vbCrLf
                    s = s & "   ,NetTax=" & DbQuote(Cur, "" & rs("Tax")) & vbCrLf
                    s = s & "   ,TotalAmount=" & DbQuote(Cur, Val("" & rs("Tax")) + Val("" & rs("Pretax"))) & vbCrLf
                Else
                    s = s & "   ,UseTax=0" & vbCrLf
                    s = s & "   ,NetTax=0" & vbCrLf
                    s = s & "   ,TotalAmount=" & DbQuote(Cur, "" & rs("Pretax")) & vbCrLf
                End If
                s = s & "WHERE ISNULL(Area,'')=" & DbQuote(Str, "" & rs("Community")) & vbCrLf
                s = s & "  AND ISNULL(CommunityPhase,'')=" & DbQuote(Str, "" & rs("CommunityPhase")) & vbCrLf
                s = s & "  AND ISNULL(Model,'')=" & DbQuote(Str, "" & rs("Model")) & vbCrLf
                s = s & "  AND ISNULL(Series,'')=" & DbQuote(Str, "" & rs("Series")) & vbCrLf
                s = s & "  AND ISNULL(Elevation,'')=" & DbQuote(Str, "" & rs("Elevation")) & vbCrLf
                s = s & "  AND DivisionID = " & HFApp.DivisionID
                Call HFApp.SqlExec(s)

            
            Case atoption
                s = ""
                s = s & "INSERT INTO tblOptions(DivisionID,Area,CommunityPhase,Model,Series,Elevation,Opt,Option_Type,Option_Type_Desc,Inactive)" & vbCrLf
                s = s & "VALUES(" & HFApp.DivisionID & "," & DbQuote(Str, "" & rs("Community")) & vbCrLf
                s = s & "      ," & DbQuote(Str, "" & rs("CommunityPhase")) & vbCrLf
                s = s & "      ," & DbQuote(Str, "" & rs("Model")) & vbCrLf
                s = s & "      ," & DbQuote(Str, "" & rs("Series")) & vbCrLf
                s = s & "      ," & DbQuote(Str, "" & rs("Elevation")) & vbCrLf
                s = s & "      ," & DbQuote(Str, "" & rs("OptionID")) & vbCrLf
                s = s & "      ,1" & vbCrLf
                s = s & "      ,'Sales Center'" & ",0)" & vbCrLf
                Call HFApp.SqlExec(s)

                s = ""
                s = s & "UPDATE tblOptions" & vbCrLf
                s = s & "SET LastSalesWorksheet=SalesWorksheet" & vbCrLf
                s = s & "   ,SalesWorksheet=" & DbQuote(Num, Worksheet) & vbCrLf
                s = s & "   ,Assembly=" & DbQuote(Str, "" & rs("Assembly")) & vbCrLf
                s = s & "   ,UOM=" & DbQuote(Str, "" & rs("AssemblyUOM")) & vbCrLf
                s = s & "   ,Description=" & DbQuote(Str, "" & rs("Description")) & vbCrLf
                s = s & "   ,Color=" & DbQuote(Str, "" & rs("color")) & vbCrLf
                s = s & "   ,ColorListID=" & DbQuote(Num, "" & rs("ColorListID")) & vbCrLf
                s = s & "   ,StyleListID=" & DbQuote(Num, "" & rs("StyleListID")) & vbCrLf
                s = s & "   ,FinishListID=" & DbQuote(Num, "" & rs("FinishListID")) & vbCrLf
                s = s & "   ,OtherListID=" & DbQuote(Num, "" & rs("OtherListID")) & vbCrLf
                s = s & "   ,Style=" & DbQuote(Str, "" & rs("StyleValue")) & vbCrLf
                s = s & "   ,Finish=" & DbQuote(Str, "" & rs("FinishValue")) & vbCrLf
                s = s & "   ,Other=" & DbQuote(Str, "" & rs("OtherValue")) & vbCrLf
                s = s & "   ,location=" & DbQuote(Str, "" & rs("location")) & vbCrLf
                s = s & "   ,qty=" & DbQuote(Num, "" & rs("qty")) & vbCrLf
                s = s & "   ,includedoption=" & DbQuote(Bit, "" & rs("IncludedOption")) & vbCrLf
                s = s & "   ,Construction_Cut_Off=" & DbQuote(Num, "" & rs("ConstCutoff")) & vbCrLf
                s = s & "   ,Graphic_Path=" & DbQuote(Str, "" & rs("GraphicPath")) & vbCrLf
                s = s & "   ,Inactive=" & DbQuote(Bit, "" & rs("Inactive")) & vbCrLf
                s = s & "   ,DesignCenterSalesOnly=" & DbQuote(Bit, "" & rs("DesignCenterSalesOnly")) & vbCrLf
                s = s & "   ,SelectByRoom=" & DbQuote(Bit, "" & rs("SelectByRoom")) & vbCrLf
                s = s & "   ,DisplayTotalOnly=" & DbQuote(Bit, "" & rs("DisplayTotalOnly")) & vbCrLf
                s = s & "   ,Cost_Amount=" & DbQuote(Num, Val("" & rs("Cost"))) & vbCrLf
                s = s & "   ,Category=" & DbQuote(Str, "" & rs("Category")) & vbCrLf
                s = s & "   ,Major_Group=" & DbQuote(Str, "" & rs("Group_Code")) & vbCrLf
                s = s & "   ,EstimatorNotes=" & DbQuote(Str, "" & rs("Notes")) & vbCrLf
                s = s & "   ,comments=" & DbQuote(Str, "" & rs("comments")) & vbCrLf
                s = s & "   ,TL_Extra=" & DbQuote(Str, "" & rs("JCExtra")) & vbCrLf
                s = s & "   ,Price=" & DbQuote(Num, Val("" & rs("Pretax"))) & vbCrLf
                s = s & "   ,CO_Price=" & DbQuote(Num, Max(Val("" & rs("Pretax")), Val("" & rs("COPretax")))) & vbCrLf
                s = s & "   ,Margin=" & DbQuote(Num, "" & rs("Margin")) & vbCrLf
                s = s & "   ,Markup=" & DbQuote(Num, "" & rs("Markup")) & vbCrLf
                If CBool("" & rs("IncludeTax")) Then
                    s = s & "   ,UseTax=1" & vbCrLf
                    s = s & "   ,NetTax=" & DbQuote(Num, "" & rs("Tax")) & vbCrLf
                    s = s & "   ,TotalAmount=" & DbQuote(Num, Val("" & rs("Tax")) + Val("" & rs("Pretax"))) & vbCrLf
                    s = s & "   ,NetCOTax=" & DbQuote(Num, "" & rs("COTax")) & vbCrLf
                    s = s & "   ,TotalCOAmount=" & DbQuote(Num, Val("" & rs("COTax")) + Val("" & rs("COPretax"))) & vbCrLf
                Else
                    s = s & "   ,UseTax=0" & vbCrLf
                    s = s & "   ,NetTax=0" & vbCrLf
                    s = s & "   ,TotalAmount=" & DbQuote(Num, "" & rs("Pretax")) & vbCrLf
                    s = s & "   ,NetCOTax=0" & vbCrLf
                    s = s & "   ,TotalCOAmount=" & DbQuote(Num, "" & rs("COPretax")) & vbCrLf
                End If
                s = s & "WHERE ISNULL(Area,'')=" & DbQuote(Str, "" & rs("Community")) & vbCrLf
                s = s & "  AND ISNULL(CommunityPhase,'')=" & DbQuote(Str, "" & rs("CommunityPhase")) & vbCrLf
                s = s & "  AND Model=" & DbQuote(Str, "" & rs("Model")) & vbCrLf
                s = s & "  AND Elevation=" & DbQuote(Str, "" & rs("Elevation")) & vbCrLf
                s = s & "  AND Series=" & DbQuote(Str, "" & rs("Series")) & vbCrLf
                s = s & "  AND Opt=" & DbQuote(Str, "" & rs("OptionID")) & vbCrLf
                s = s & "  AND DivisionID = " & HFApp.DivisionID
                Call HFApp.SqlExec(s)



            Case atDesignCenter
                s = ""
                s = s & "INSERT INTO tblDCOptions(DivisionID,Option_Type,Community,CommunityPhase,Opt,Inactive)" & vbCrLf
                s = s & "VALUES(" & HFApp.DivisionID & ",3," & DbQuote(Str, "" & rs("Community")) & vbCrLf
                s = s & "      ," & DbQuote(Str, "" & rs("CommunityPhase")) & vbCrLf
                s = s & "      ," & DbQuote(Str, "" & rs("OptionID")) & ",0)"
                Call HFApp.SqlExec(s)

                s = ""
                s = s & "UPDATE tblDCOptions" & vbCrLf
                s = s & "SET SalesWorksheet=" & DbQuote(Num, Worksheet) & vbCrLf
                s = s & "   ,Assembly=" & DbQuote(Str, "" & rs("Assembly")) & vbCrLf
                s = s & "   ,TL_Extra=" & DbQuote(Str, "" & rs("JCExtra")) & vbCrLf
                s = s & "   ,Description=" & DbQuote(Str, "" & rs("Description")) & vbCrLf
                s = s & "   ,UOM=" & DbQuote(Str, "" & rs("AssemblyUOM")) & vbCrLf
                s = s & "   ,EstimatorNotes=" & DbQuote(Str, "" & rs("Notes")) & vbCrLf
                s = s & "   ,comments=" & DbQuote(Str, "" & rs("comments")) & vbCrLf
                s = s & "   ,Item1=" & DbQuote(Num, Val("" & rs("Pretax"))) & vbCrLf
                s = s & "   ,Item1_Cost=" & DbQuote(Num, Val("" & rs("Cost"))) & vbCrLf
                s = s & "   ,Color=" & DbQuote(Str, "" & rs("color")) & vbCrLf
                s = s & "   ,ColorListID=" & DbQuote(Num, "" & rs("ColorListID")) & vbCrLf
                s = s & "   ,StyleListID=" & DbQuote(Num, "" & rs("StyleListID")) & vbCrLf
                s = s & "   ,FinishListID=" & DbQuote(Num, "" & rs("FinishListID")) & vbCrLf
                s = s & "   ,OtherListID=" & DbQuote(Num, "" & rs("OtherListID")) & vbCrLf
                s = s & "   ,DesignCenterSalesOnly=" & DbQuote(Bit, "" & rs("DesignCenterSalesOnly")) & vbCrLf
                s = s & "   ,SelectByRoom=" & DbQuote(Bit, "" & rs("SelectByRoom")) & vbCrLf
                s = s & "   ,DisplayTotalOnly=" & DbQuote(Bit, "" & rs("DisplayTotalOnly")) & vbCrLf
                s = s & "   ,Style=" & DbQuote(Str, "" & rs("StyleValue")) & vbCrLf
                s = s & "   ,Finish=" & DbQuote(Str, "" & rs("FinishValue")) & vbCrLf
                s = s & "   ,Graphic_Path=" & DbQuote(Str, "" & rs("GraphicPath")) & vbCrLf
                s = s & "   ,Inactive=" & DbQuote(Bit, "" & rs("Inactive")) & vbCrLf
                s = s & "   ,Other=" & DbQuote(Str, "" & rs("OtherValue")) & vbCrLf
                s = s & "   ,location=" & DbQuote(Str, "" & rs("location")) & vbCrLf
                s = s & "   ,qty=" & DbQuote(Num, "" & rs("qty")) & vbCrLf
                s = s & "   ,includedoption=" & DbQuote(Bit, "" & rs("IncludedOption")) & vbCrLf
                s = s & "   ,Nettax1=" & DbQuote(Num, Val("" & rs("Tax"))) & vbCrLf
                s = s & "   ,TotalAmount1=" & DbQuote(Num, Val("" & rs("Pretax")) + Val("" & rs("Tax"))) & vbCrLf
                s = s & "   ,Margin1=" & DbQuote(Num, "" & rs("Margin")) & vbCrLf
                s = s & "   ,Markup1=" & DbQuote(Num, "" & rs("Markup")) & vbCrLf
                s = s & "   ,Category=" & DbQuote(Str, "" & rs("Category")) & vbCrLf
                For i = 2 To 10
                    s = s & "   ,Margin" & i & "=" & DbQuote(Num, "" & rs("Margin" & i)) & vbCrLf
                    s = s & "   ,Markup" & i & "=" & DbQuote(Num, "" & rs("Markup" & i)) & vbCrLf
                    s = s & "   ,Item" & i & "=" & DbQuote(Num, Val("" & rs("Pretax" & i))) & vbCrLf
                    s = s & "   ,Item" & i & "_Cost=" & DbQuote(Num, Val("" & rs("Cost"))) & vbCrLf
                    s = s & "   ,Nettax" & i & "=" & DbQuote(Num, Val("" & rs("Tax" & i))) & vbCrLf
                    s = s & "   ,TotalAmount" & i & "=" & DbQuote(Num, Val("" & rs("Pretax" & i)) + Val("" & rs("Tax" & i))) & vbCrLf
                Next
                s = s & "WHERE ISNULL(Community,'')=" & DbQuote(Str, "" & rs("Community")) & vbCrLf
                s = s & "  AND ISNULL(CommunityPhase,'')=" & DbQuote(Str, "" & rs("CommunityPhase")) & vbCrLf
                s = s & "  AND Opt=" & DbQuote(Str, "" & rs("OptionID")) & vbCrLf
                s = s & "  AND DivisionID = " & HFApp.DivisionID
                Call HFApp.SqlExec(s)

            Case atGlobal
                s = ""
                s = s & "INSERT INTO tblGlobalOptions(DivisionID,Option_Type,Community,CommunityPhase,Category,Opt,Inactive)" & vbCrLf
                s = s & "VALUES(" & HFApp.DivisionID & ",2," & DbQuote(Str, "" & rs("Community")) & vbCrLf
                s = s & "      ," & DbQuote(Str, "" & rs("CommunityPhase")) & vbCrLf
                s = s & "      ," & DbQuote(Str, "" & rs("Category")) & vbCrLf
                s = s & "      ," & DbQuote(Str, "" & rs("OptionID")) & ",0)"
                Call HFApp.SqlExec(s)
                
                s = ""
                s = s & "UPDATE tblGlobalOptions" & vbCrLf
                s = s & "SET LastSalesWorksheet=SalesWorksheet" & vbCrLf
                s = s & "   ,SalesWorksheet=" & DbQuote(Num, Worksheet) & vbCrLf
                s = s & "   ,Assembly=" & DbQuote(Str, "" & rs("Assembly")) & vbCrLf
                s = s & "   ,UOM=" & DbQuote(Str, "" & rs("AssemblyUOM")) & vbCrLf
                s = s & "   ,Description=" & DbQuote(Str, "" & rs("Description")) & vbCrLf
                s = s & "   ,Construction_Cut_Off=" & DbQuote(Num, "" & rs("ConstCutoff")) & vbCrLf
                s = s & "   ,Graphic_Path=" & DbQuote(Str, "" & rs("GraphicPath")) & vbCrLf
                s = s & "   ,Inactive=" & DbQuote(Bit, "" & rs("Inactive")) & vbCrLf
                s = s & "   ,Cost_Amount=" & DbQuote(Num, Val("" & rs("Cost"))) & vbCrLf
                s = s & "   ,DesignCenterSalesOnly=" & DbQuote(Bit, "" & rs("DesignCenterSalesOnly")) & vbCrLf
                s = s & "   ,SelectByRoom=" & DbQuote(Bit, "" & rs("SelectByRoom")) & vbCrLf
                s = s & "   ,DisplayTotalOnly=" & DbQuote(Bit, "" & rs("DisplayTotalOnly")) & vbCrLf
                s = s & "   ,Category=" & DbQuote(Str, "" & rs("Category")) & vbCrLf
                s = s & "   ,TL_Extra=" & DbQuote(Str, "" & rs("JCExtra")) & vbCrLf
                s = s & "   ,Major_Group=" & DbQuote(Str, "" & rs("Group_Code")) & vbCrLf
                s = s & "   ,Color=" & DbQuote(Str, "" & rs("color")) & vbCrLf
                s = s & "   ,ColorListID=" & DbQuote(Num, "" & rs("ColorListID")) & vbCrLf
                s = s & "   ,StyleListID=" & DbQuote(Num, "" & rs("StyleListID")) & vbCrLf
                s = s & "   ,FinishListID=" & DbQuote(Num, "" & rs("FinishListID")) & vbCrLf
                s = s & "   ,OtherListID=" & DbQuote(Num, "" & rs("OtherListID")) & vbCrLf
                s = s & "   ,Style=" & DbQuote(Str, "" & rs("StyleValue")) & vbCrLf
                s = s & "   ,Finish=" & DbQuote(Str, "" & rs("FinishValue")) & vbCrLf
                s = s & "   ,Other=" & DbQuote(Str, "" & rs("OtherValue")) & vbCrLf
                s = s & "   ,location=" & DbQuote(Str, "" & rs("location")) & vbCrLf
                s = s & "   ,qty=" & DbQuote(Num, "" & rs("qty")) & vbCrLf
                s = s & "   ,includedoption=" & DbQuote(Bit, "" & rs("IncludedOption")) & vbCrLf
                s = s & "   ,EstimatorNotes=" & DbQuote(Str, "" & rs("Notes")) & vbCrLf
                s = s & "   ,comments=" & DbQuote(Str, "" & rs("comments")) & vbCrLf
                s = s & "   ,Price=" & DbQuote(Num, "" & rs("Pretax")) & vbCrLf
                s = s & "   ,CO_Price=" & DbQuote(Num, Max(Val("" & rs("Pretax")), Val("" & rs("COPretax")))) & vbCrLf
                s = s & "   ,Margin=" & DbQuote(Num, "" & rs("Margin")) & vbCrLf
                s = s & "   ,Markup=" & DbQuote(Num, "" & rs("Markup")) & vbCrLf
                If CBool("" & rs("IncludeTax")) Then
                    s = s & "   ,UseTax=1" & vbCrLf
                    s = s & "   ,NetTax=" & DbQuote(Num, "" & rs("Tax")) & vbCrLf
                    s = s & "   ,TotalAmount=" & DbQuote(Num, Val("" & rs("Tax")) + Val("" & rs("Pretax"))) & vbCrLf
                    s = s & "   ,NetCOTax=" & DbQuote(Num, "" & rs("COTax")) & vbCrLf
                    s = s & "   ,TotalCOAmount=" & DbQuote(Num, Val("" & rs("COTax")) + Val("" & rs("COPretax"))) & vbCrLf
                Else
                    s = s & "   ,UseTax=0" & vbCrLf
                    s = s & "   ,NetTax=0" & vbCrLf
                    s = s & "   ,TotalAmount=" & DbQuote(Num, "" & rs("Pretax")) & vbCrLf
                    s = s & "   ,NetCOTax=0" & vbCrLf
                    s = s & "   ,TotalCOAmount=" & DbQuote(Num, "" & rs("COPretax")) & vbCrLf
                End If
                s = s & "WHERE ISNULL(Community,'')=" & DbQuote(Str, "" & rs("Community")) & vbCrLf
                s = s & "  AND ISNULL(CommunityPhase,'')=" & DbQuote(Str, "" & rs("CommunityPhase")) & vbCrLf
                s = s & "  AND Opt=" & DbQuote(Str, "" & rs("OptionID")) & vbCrLf
                s = s & "  AND DivisionID = " & HFApp.DivisionID
                Call HFApp.SqlExec(s)

        End Select

        rs.MoveNext
    Wend
    
    s = "UPDATE tblSalesSheetMaster SET SalesEffectiveDate=GETDATE() WHERE Worksheet=" & DbQuote(Num, Worksheet)
    Call HFApp.SqlExec(s)
    
    s = "UPDATE tblSalesSheetDetails SET ReadyToPublish = 0 WHERE Worksheet=" & DbQuote(Num, Worksheet)
    Call HFApp.SqlExec(s)
    
    
    Exit Sub
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Call errHandler(SRCFILE & "PublishPricingWorksheet")
    End If
End Sub

Public Sub LoadCostTypes(Optional Combobox, Optional Grid, Optional ForecastsOnly As Boolean = False)
    Dim s  As String
    Dim rs As Recordset
    
    If Not IsMissing(Combobox) Then
        Combobox.Clear
        If Not ForecastsOnly Then
            Call Combobox.AddItem("Current")
            Call Combobox.AddItem("Next 1")
            Call Combobox.AddItem("Next 2")
            Call Combobox.AddItem("Last 1")
            Call Combobox.AddItem("Last 2")
            Call Combobox.AddItem("Last 3")
        End If
    End If
    
    
    s = ""
    s = s & "SELECT Item,MAX(Custom_Description),'Forecast 1'  FROM CustomDescriptions WHERE Item ='Forecast1' GROUP BY Item UNION ALL" & vbCrLf
    s = s & "SELECT Item,MAX(Custom_Description),'Forecast 2'  FROM CustomDescriptions WHERE Item ='Forecast2' GROUP BY Item UNION ALL" & vbCrLf
    s = s & "SELECT Item,MAX(Custom_Description),'Forecast 3'  FROM CustomDescriptions WHERE Item ='Forecast3' GROUP BY Item UNION ALL" & vbCrLf
    s = s & "SELECT Item,MAX(Custom_Description),'Forecast 4'  FROM CustomDescriptions WHERE Item ='Forecast4' GROUP BY Item UNION ALL" & vbCrLf
    s = s & "SELECT Item,MAX(Custom_Description),'Forecast 5'  FROM CustomDescriptions WHERE Item ='Forecast5' GROUP BY Item UNION ALL" & vbCrLf
    s = s & "SELECT Item,MAX(Custom_Description),'Forecast 6'  FROM CustomDescriptions WHERE Item ='Forecast6' GROUP BY Item UNION ALL" & vbCrLf
    s = s & "SELECT Item,MAX(Custom_Description),'Forecast 7'  FROM CustomDescriptions WHERE Item ='Forecast7' GROUP BY Item UNION ALL" & vbCrLf
    s = s & "SELECT Item,MAX(Custom_Description),'Forecast 8'  FROM CustomDescriptions WHERE Item ='Forecast8' GROUP BY Item UNION ALL" & vbCrLf
    s = s & "SELECT Item,MAX(Custom_Description),'Forecast 9'  FROM CustomDescriptions WHERE Item ='Forecast9' GROUP BY Item UNION ALL" & vbCrLf
    s = s & "SELECT Item,MAX(Custom_Description),'Forecast 10' FROM CustomDescriptions WHERE Item ='Forecast10' GROUP BY Item UNION ALL" & vbCrLf
    s = s & "SELECT Item,MAX(Custom_Description),'Forecast 11' FROM CustomDescriptions WHERE Item ='Forecast11' GROUP BY Item UNION ALL" & vbCrLf
    s = s & "SELECT Item,MAX(Custom_Description),'Forecast 12' FROM CustomDescriptions WHERE Item ='Forecast12' GROUP BY Item" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    While Not rs.EOF
    
        If Not IsMissing(Combobox) Then Call Combobox.AddItem(IIf("" & rs(1) = "", "" & rs(2), "" & rs(1)))
        If Not IsMissing(Grid) Then Grid.TextMatrix(0, Grid.ColIndex("" & rs(0))) = IIf("" & rs(1) = "", "" & rs(2), "" & rs(1))
        
        rs.MoveNext
    Wend
    
    
    If Not IsMissing(Combobox) Then Combobox.ListIndex = 0
 
End Sub

Public Function ValidateField(Grid As VSFlexGrid, ByRef EditText As String, WarningMsg As String, sql As String, ParamArray OtherColumns())
On Error GoTo eh
    
    Dim rs As Recordset
    Dim i As Long
    
    With Grid
    If EditText = "" Then
        ValidateField = True
        For i = 0 To UBound(OtherColumns)
            .Cell(flexcpText, .Row, .ColIndex(OtherColumns(i)), .RowSel, .ColIndex(OtherColumns(i))) = ""
        Next
    Else
        Set rs = HFApp.SqlExec(sql)
        EditText = "" & rs(0)
        For i = 0 To UBound(OtherColumns)
            .Cell(flexcpText, .Row, .ColIndex(OtherColumns(i)), .RowSel, .ColIndex(OtherColumns(i))) = "" & rs(i + 1)
        Next
    End If
    End With
    ValidateField = True
    
Exit Function
eh: If WarningMsg <> "" Then MsgBox WarningMsg, vbExclamation, App.ProductName
    ValidateField = False
End Function




Public Property Let CtrlEnabled(c As Control, RHS As Boolean)
On Error Resume Next
    c.Enabled = RHS
    c.ForeColor = IIf(c.Enabled, vbWindowText, vbGrayText)
    c.BackColor = IIf(c.Enabled, vbWindowBackground, vbButtonFace)
End Property



Public Function CancelPO(PONumber As String, Optional PromptForCancelReason As Boolean = True, Optional EmailPrompt As Long = 0, Optional CancelReason As String) As Boolean
'EmailPrompt = 0,1,2 -- ask,send,dontsend

    Dim i As Long
    Dim rs As Recordset
    Dim Batch As Long
    Dim OldVendor As String
    Dim OldVendorName As String
    Dim EmailAddr As String
    Dim PMEmailAddr As String
    Dim b As Boolean
    Dim s As String
    Dim Subject As String
    Dim body As String
    Dim sJob As String
    Dim AmtInvoiced As Double
    Dim reason As String
    
    reason = CancelReason
    
    'cant if invoiced
    'first check locally
    AmtInvoiced = 0
    s = "Select isnull(sum(PreTax),0) from dbo.POInvoicedAmounts  where DivisionID = " & HFApp.DivisionID & " and PONumber = " & DbQuote(Str, PONumber)
    Set rs = HFApp.SqlExec(s, dbHomefront)
    If Not rs.EOF Then AmtInvoiced = Val("" & rs(0))
    
    If AmtInvoiced = 0 And HFApp.Options(AccountingSystem) = asTimberline Then
        'now check timberline
        s = "select samtinv from master_jcm_record_12 where sub=" & DbQuote(Str, PONumber)
        Set rs = HFApp.SqlExec(s, dbAccountingDictionary)
        If Not rs.EOF Then AmtInvoiced = Val("" & rs(0))
    End If
    
    
    If AmtInvoiced <> 0 Then
        Call MsgBox("Unable to cancel " & PONumber & ". It has invoices applied to it.", vbOKOnly + vbExclamation, "Cancel PO")
        CancelPO = False
        Exit Function
    Else
        If PromptForCancelReason Then
            'If vbNo = MsgBox("Are you sure you want to cancel this purchase order?", vbQuestion + vbYesNo, "Cancel Purchase Order") Then
            '    Exit Function
            'End If
            If Not FCancelPO.CancelPO(reason) Then
                Exit Function
            End If
        End If
    End If
    
    
    'build mail string for vendor notification
    s = ""
    If "True" = "" & HFApp.Options.ValueByName("SendPO_CCPrjMgr") Then
        s = s & "SELECT p.Vendor,isnull(nullif(p.DeliveryAddress,''),nullif(v.purchemail,'')) PurchEmail,v.Vendor_Name,p.PostingBatch,p.Job,pm.email PMEmail " & vbCrLf
        s = s & "FROM POMaster p" & vbCrLf
        s = s & "LEFT OUTER JOIN tblVendors v ON(p.Vendor=v.Vendor_ID and p.DivisionID = v.DivisionID) " & vbCrLf
        s = s & "left outer join tbljobs j on (p.job=j.job_no and p.DivisionID = j.DivisionID)" & vbCrLf
        s = s & "left outer join tblprojectmanager pm on (j.pm=pm.pm)" & vbCrLf
    Else
        s = s & "SELECT p.Vendor,isnull(nullif(p.DeliveryAddress,''),nullif(v.purchemail,'')) PurchEmail,v.Vendor_Name,p.PostingBatch,p.Job,null PMEmail " & vbCrLf
        s = s & "FROM POMaster p" & vbCrLf
        s = s & "LEFT OUTER JOIN tblVendors v ON(p.Vendor=v.Vendor_ID and p.DivisionID = v.DivisionID) " & vbCrLf
    End If
    s = s & "WHERE p.DivisionID = " & HFApp.DivisionID & vbCrLf
    s = s & " and p.PONumber=" & DbQuote(Str, PONumber)
    Set rs = HFApp.SqlExec(s, dbHomefront)
    If rs.EOF Then Exit Function
    
    OldVendor = "" & rs(0)
    EmailAddr = "" & rs(1)
    OldVendorName = "" & rs(2)
    Batch = Val("" & rs(3))
    sJob = "" & rs(4)
    PMEmailAddr = "" & rs(5)
    If Trim(OldVendorName) = "" Then OldVendorName = OldVendor
    If Trim(EmailAddr) = "" Then EmailAddr = OldVendorName
    
    
    Select Case True
    Case Batch <> 0 And HFApp.Options(AccountingSystem) = asIntacct
        b = CancelPOInIntacct(PONumber, reason)
        If Not b Then
            b = MsgBox("Unable to remove PO from Intacct." & vbCrLf & vbCrLf & " Would you like to remove it from HomeFront? You will have to manually remove the PO from Intacct.", vbYesNo + vbQuestion, App.ProductName) = vbYes
        End If
        
    Case Batch <> 0 And HFApp.Options(AccountingSystem) = asTimberline:        b = CancelPO_TL(PONumber, reason)
    Case Batch <> 0 And HFApp.Options(AccountingSystem) = asQuickBooks:        b = CancelPO_QB(PONumber, reason)
    Case Batch <> 0 And HFApp.Options(AccountingSystem) = asMasterBuilder:     b = CancelPO_MB(PONumber, reason)
    Case Else
        'hasn't been posted yet
        Call HFApp.WriteAuditLog("Edit PO", "Cancel unposted PO """ & PONumber)
        b = True
    End Select
    
    
    
    
    If b Then
                 
        'mark po as cancelled
        s = ""
        s = s & "UPDATE POMaster" & vbCrLf
        s = s & "SET Cancelled=1" & vbCrLf
        s = s & "   ,CancelledBy=" & DbQuote(Str, HFApp.LoginID) & vbCrLf
        s = s & "   ,CancelledNotes=" & DbQuote(Str, reason) & vbCrLf
        s = s & "   ,CancelledDate=GETDATE()" & vbCrLf
        s = s & "   ,TStmp=GETDATE()" & vbCrLf
        s = s & "   ,DateSentToBuildPro=null" & vbCrLf
        s = s & "WHERE DivisionID = " & HFApp.DivisionID & vbCrLf
        s = s & "  AND PONumber=" & DbQuote(Str, PONumber) & vbCrLf
        Call HFApp.SqlExec(s, dbHomefront)
        Call HFApp.SqlExec("UPDATE EstimateItems SET POGenBatch=0,PONumber='' WHERE divisionid=" & HFApp.DivisionID & " and PONumber=" & DbQuote(Str, PONumber), dbHomefront)
        
        
        
        Dim Response As Integer
        Select Case EmailPrompt
        Case 0 'ask
            s = "Purchase order """ & PONumber & """ has been cancelled. Please" & vbCrLf & "notify """ & OldVendorName & """ of the change." & vbCrLf & vbCrLf & "Would you like to send them an email?"
            Response = MsgBox(s, vbQuestion + vbYesNo, "Purchase Order Cancelled")
        Case 1 'send
            Response = vbYes
        Case 2 'dontsend
            Response = vbNo
        End Select
        If Response = vbYes Then
            'build message subject and body
            s = ""
            s = s & "SELECT pm.PMName SenderName" & vbCrLf
            s = s & "      ,s.CompanyName SenderCompany" & vbCrLf
            s = s & "      ,CASE WHEN s.SendFaxUsingPurchasersInfo=1 THEN pm.fax   ELSE s.fax END SenderFax" & vbCrLf
            s = s & "      ,CASE WHEN s.SendFaxUsingPurchasersInfo=1 THEN pm.phone ELSE s.fax END SenderPhone" & vbCrLf
            s = s & "      ,CASE WHEN s.SendFaxUsingPurchasersInfo=1 THEN pm.email ELSE s.estimator_email END SenderEmail" & vbCrLf
            s = s & "      ,v.PurchContact RecipientName" & vbCrLf
            s = s & "      ,v.Vendor_Name RecipientCompany" & vbCrLf
            s = s & "      ,p.deliveryaddress RecipientFax" & vbCrLf
            s = s & "      ,v.PurchPhone RecipientPhone" & vbCrLf
            s = s & "      ,isnull(nullif(p.DeliveryAddress,''),nullif(v.purchemail,'')) RecipientEmail" & vbCrLf
            s = s & "      ,v.PurchCell SMSCell" & vbCrLf
            s = s & "      ,v.PurchSmsAddress SMSAddress" & vbCrLf
            s = s & "      ,p.CancelledNotes POCancellationReason" & vbCrLf
            s = s & "      ,p.Job Quote" & vbCrLf
            s = s & "      ,p.Job" & vbCrLf
            s = s & "      ,p.POIndex" & vbCrLf
            s = s & "      ,pi.description POIndexDescription" & vbCrLf
            s = s & "      ,p.Vendor" & vbCrLf
            s = s & "      ,v.Vendor_Name VendorName" & vbCrLf
            s = s & "      ,j.Model" & vbCrLf
            s = s & "      ,j.Description JobDescription" & vbCrLf
            s = s & "      ,j.Municipal_Address  MunicipalAddress " & vbCrLf
            s = s & "      ,j.Legal_Address  LegalAddress " & vbCrLf
            s = s & "      ,j.County" & vbCrLf
            s = s & "      ,j.Township" & vbCrLf
            s = s & "      ,ISNULL(NULLIF(l.Description,''),j.Community) Community" & vbCrLf
            s = s & "      ,ISNULL(NULLIF(a.Description,''),j.CommunityPhase) Phase" & vbCrLf
            s = s & "  FROM POMaster p" & vbCrLf
            s = s & "       LEFT OUTER JOIN tblpoindex pi on (p.DivisionID = pi.divisionid and p.poindex=pi.poindex)" & vbCrLf
            s = s & "       LEFT OUTER JOIN System_setup s on (s.ID = p.DivisionID)" & vbCrLf
            s = s & "       LEFT OUTER JOIN tblJobs j on(p.Job=j.Job_no and p.DivisionID = j.DivisionID)" & vbCrLf
            s = s & "       LEFT OUTER JOIN tblProjectManager pm on(j.Purchaser=pm.PM)" & vbCrLf
            s = s & "       LEFT OUTER JOIN tblVendors v on(p.vendor=v.vendor_id and p.DivisionID = v.DivisionID)" & vbCrLf
            s = s & "       LEFT OUTER JOIN tblLocality l on(j.community=l.area)" & vbCrLf
            s = s & "       LEFT OUTER JOIN CommunityPhase a on(j.Community = a.Community and j.CommunityPhase=a.CommunityPhase)" & vbCrLf
            s = s & "WHERE p.DivisionID = " & HFApp.DivisionID & " and p.PONumber=" & DbQuote(Str, PONumber)
            Set rs = HFApp.SqlExec(s)
            s = Replace(Replace(PONumber, "','", ", "), "'", "")
            Subject = Replace(HFApp.Options(Msg_CancelPO_Subject), "<%PONumber%>", s, , , vbTextCompare)
            body = Replace(HFApp.Options.value(Msg_CancelPO_Body), "<%PONumber%>", s, , , vbTextCompare)
            For i = 0 To rs.fields.Count - 1
                Subject = Replace(Subject, "<%" & rs.fields(i).Name & "%>", "" & rs(i), , , vbTextCompare)
                body = Replace(body, "<%" & rs.fields(i).Name & "%>", "" & rs(i), , , vbTextCompare)
            Next
            Call HFApp.SqlExec("UPDATE POMaster SET CancellationSent=1 WHERE DivisionID = " & HFApp.DivisionID & " and PONumber=" & DbQuote(Str, PONumber), dbHomefront)
            Call HFApp.SendMail(True, EmailAddr, PMEmailAddr, Subject, body, "")
        
        End If
        
        
    End If
    CancelPO = b

End Function



Public Function ChangePOVendor(PONumber As String, Optional NewVendor As String, Optional UpdateInvoices As Boolean) As Boolean
    Dim posted As Boolean
    Dim sMailStr As String
    Dim s As String
    Dim rs As Recordset
    Dim OldVendor As String
    Dim OldVendorName As String
    Dim NewVendorName As String
    Dim EmailAddr As String
    Dim NewEmailAddr As String
    Dim NewContactPerson As String
    Dim NewDeliveryMethod As Integer
    Dim FromAddress As String
    Dim AmtInvoiced As Double
    Dim Job As String
    
    '-----------------------------
    'are we allowed to change it?
    '-----------------------------
    Select Case HFApp.Options(AccountingSystem)
    Case asTimberline, asQuickBooks
    Case Else
        s = ""
        s = s & "select v.istbd" & vbCrLf
        s = s & "from pomaster p" & vbCrLf
        s = s & "join tblvendors v on p.vendor=v.vendor_id and p.divisionid=v.divisionid" & vbCrLf
        s = s & "where p.ponumber=" & DbQuote(Str, PONumber) & vbCrLf
        s = s & "and p.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        Set rs = HFApp.SqlExec(s)
        If "" & rs(0) = "True" Then
        Else
            MsgBox "Change vendor is not supported by your accounting system.", vbInformation, App.ProductName
            ChangePOVendor = False
            Exit Function
        End If
    End Select
    
    
    If HFApp.Options(AccountingSystem) = asTimberline Then
        Set rs = HFApp.SqlExec("SELECT Amount_Invoiced FROM JCM_Master__Commitment WHERE Commitment=" & DbQuote(Str, PONumber), dbAccounting)
        If rs.EOF Then
            posted = False
            If Not UpdateInvoices Then
                AmtInvoiced = HFApp.SqlExec("Select isnull(sum(PreTax),0) from dbo.POInvoicedAmounts  where DivisionID = " & HFApp.DivisionID & " and PONumber = " & DbQuote(Str, PONumber), dbHomefront)(0)
                If AmtInvoiced <> 0 Then
                    MsgBox "The PO has been invoiced. The vendor cannot be changed.", vbInformation, App.ProductName
                    ChangePOVendor = False
                    Exit Function
                End If
            End If
        Else
            posted = True
            If 0 <> Val("" & rs(0)) Then
                MsgBox "The commitment has an invoice recorded in Sage300. The vendor cannot be changed.", vbInformation, App.ProductName
                Exit Function
            End If
            Set rs = HFApp.SqlExec("SELECT Commitment_CO FROM JCM_Master__Commitment_CO WHERE Commitment=" & DbQuote(Str, PONumber), dbAccounting)
            If Not rs.EOF Then
                MsgBox "The commitment has a change order in Sage300. The vendor cannot be changed.", vbInformation, App.ProductName
                Exit Function
            End If
        End If
    Else
        If Not UpdateInvoices Then
            AmtInvoiced = HFApp.SqlExec("Select isnull(sum(PreTax),0) from dbo.POInvoicedAmounts  where DivisionID = " & HFApp.DivisionID & " and PONumber = " & DbQuote(Str, PONumber), dbHomefront)(0)
            If AmtInvoiced <> 0 Then
                MsgBox "The PO has been invoiced. The vendor cannot be changed.", vbInformation, App.ProductName
                ChangePOVendor = False
                Exit Function
            End If
        End If
    End If
    '-----------------------------
    'cant if sent to buildpro
    '-----------------------------
    If HFApp.Options.ValueByName("BuildProCompanyCode") <> "" Then
        
        'dont need to do anything if it hasnt been sent
        s = "select * from pomaster where datesenttobuildpro is not null and divisionid=" & HFApp.DivisionID & " and PONumber=" & DbQuote(Str, PONumber)
        Set rs = HFApp.SqlExec(s, dbHomefront)
        If Not rs.EOF Then
            MsgBox "The PO has been sent to BuildPro and cannot be changed. You must cancel and re-issue a new PO.", vbInformation, App.ProductName
            Exit Function
        End If
    End If

    '-----------------------------
    'choose new vendor if not given
    '-----------------------------
    Set rs = HFApp.SqlExec("SELECT Vendor,Email,Vendor_Name,job FROM POMaster LEFT OUTER JOIN tblVendors ON(POMaster.DivisionID = tblVendors.DivisionID and Vendor=Vendor_ID) WHERE pomaster.DivisionID = " & HFApp.DivisionID & " and PONumber=" & DbQuote(Str, PONumber), dbHomefront)
    OldVendor = "" & rs("vendor")
    EmailAddr = "" & rs("email")
    OldVendorName = "" & rs("vendor_name")
    Job = "" & rs("job")
    If NewVendor = "" Then
        
        If HFApp.Options(AccountingSystem) = asQuickBooks Then
            s = "SELECT Vendor_Name Company,Vendor_ID Vendor, City,Phone,PurchContact Contact,purchdelmethod DeliveryMethod,case purchdelmethod when 1 then case isnull(Purchemail,'') when '' then Email else purchEmail end else '' end Email FROM tblVendors where DivisionID = " & HFApp.DivisionID & " and inactive=0"
        Else
            s = "SELECT Vendor_ID Vendor, Vendor_Name Company,City,Phone,PurchContact Contact,purchdelmethod DeliveryMethod,case purchdelmethod when 1 then case isnull(Purchemail,'') when '' then Email else purchEmail end else '' end Email FROM tblVendors where DivisionID = " & HFApp.DivisionID & " and inactive=0"
        End If
        
        If FPickList.Choose(HFApp.Databases(dbHomefront), "Vendor", s, OldVendor, , , FMain.SmallIcons.ListImages("vendor").Picture, IIf(HFApp.Options(AccountingSystem) = asQuickBooks, "Vendor,", "") & "DeliveryMethod") Then
            NewVendor = FPickList.SelectedItem("Vendor")
            NewVendorName = FPickList.SelectedItem("Company")
            NewEmailAddr = FPickList.SelectedItem("Email")
            NewDeliveryMethod = FPickList.SelectedItem("DeliveryMethod")
            NewContactPerson = FPickList.SelectedItem("Contact")
        Else
            Exit Function
        End If
    End If
    If UCase(Trim(OldVendor)) = UCase(Trim(NewVendor)) Then
        ChangePOVendor = False
        Exit Function
    End If
    If HFApp.Options(AccountingSystem) = asTimberline Then
        Set rs = HFApp.SqlExec("SELECT Vendor FROM APM_Master__Vendor WHERE Vendor=" & DbQuote(Str, NewVendor), dbAccounting)
        If rs.EOF Then
            Err.Raise vbObjectError + 37, "HFEst", "Vendor not found."
            ChangePOVendor = False
            Exit Function
        End If
    End If
    '-----------------------------
    'capture original vendor
    '-----------------------------
    If Trim(OldVendorName) = "" Then OldVendorName = OldVendor
    If Trim(EmailAddr) = "" Then EmailAddr = OldVendorName
    sMailStr = "mailto:" & EmailAddr & _
               "?subject=" & Replace(Replace(Replace(HFApp.Options(Msg_CancelPO_Subject), "<%PONumber%>", PONumber), vbQuote, ""), vbCrLf, " ") & _
               "&body=" & Replace(Replace(Replace(HFApp.Options(Msg_CancelPO_Body), "<%PONumber%>", PONumber), vbQuote, ""), vbCrLf, " ")


    '-----------------------------
    'write audit log then update TL
    '-----------------------------
    Call HFApp.WriteAuditLog("Edit PO", "Change vendor on """ & PONumber & """ from """ & OldVendor & """ to """ & NewVendor & """")
    If HFApp.Options(AccountingSystem) = asTimberline Then
        If posted Then
            
            
            'remove trailing slash
            s = Trim(HFApp.Options(Timberline_Data_Path))
            While Right(s, 1) = "\"
                s = Mid(s, 1, Len(s) - 1)
            Wend
            s = AddQuotes(PathAppend(App.Path, "\S3W\Sage300Wrapper.exe")) & " " & _
                AddQuotes(s) & " " & _
                AddQuotes(HFApp.Options(Timberline_UID)) & " " & _
                AddQuotes(HFApp.Options(Timberline_PWD)) & " " & _
                AddQuotes(PONumber) & " " & _
                AddQuotes(NewVendor)
            Call ShellAndLoop(s, vbHide)
            ChangePOVendor = True
        Else
            ChangePOVendor = True
        End If
    ElseIf HFApp.Options(AccountingSystem) = asQuickBooks Then
        If ChangeQBPOVendor(PONumber, NewVendor) Then
            ChangePOVendor = True
        Else
            ChangePOVendor = False
            Exit Function
        End If
    End If
    
    


    
    
    '-----------------------------
    'now update our db
    '-----------------------------
    s = "UPDATE POMaster SET DateVendorChanged=getdate(), Vendor=" & DbQuote(Str, NewVendor) & ",DeliveryMethod=" & DbQuote(Num, NewDeliveryMethod) & ",DeliveryAddress=" & DbQuote(Str, NewEmailAddr) & ",DeliveryRecipient=" & DbQuote(Str, NewContactPerson) & " WHERE DivisionID = " & HFApp.DivisionID & " and PONumber=" & DbQuote(Str, PONumber)
    Call HFApp.SqlExec(s, dbHomefront)
    s = "UPDATE EstimateItems SET POVendor=" & DbQuote(Str, NewVendor) & " WHERE DivisionID = " & HFApp.DivisionID & " and PONumber=" & DbQuote(Str, PONumber)
    Call HFApp.SqlExec(s, dbHomefront)

    
    
    
    If UpdateInvoices Then
    
        s = ""
        s = s & "update invoices " & vbCrLf
        s = s & "set vendor=" & DbQuote(Str, NewVendor) & vbCrLf
        s = s & "where invoiceid in(" & vbCrLf
        s = s & "  select invoiceid" & vbCrLf
        s = s & "  from invoiceitems " & vbCrLf
        s = s & "  where divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        s = s & "  and commitment=" & DbQuote(Str, PONumber) & vbCrLf
        s = s & ")" & vbCrLf
        
        s = s & "update invoiceitems " & vbCrLf
        s = s & "set vendor=" & DbQuote(Str, NewVendor) & vbCrLf
        s = s & "where invoiceid in(" & vbCrLf
        s = s & "  select invoiceid" & vbCrLf
        s = s & "  from invoiceitems " & vbCrLf
        s = s & "  where divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        s = s & "  and commitment=" & DbQuote(Str, PONumber) & vbCrLf
        s = s & ")" & vbCrLf
        
        s = s & "update invoiceitems set" & vbCrLf
        s = s & " commitmentvendor=" & DbQuote(Str, NewVendor) & vbCrLf
        s = s & "where divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        s = s & "and commitment=" & DbQuote(Str, PONumber) & vbCrLf
        Call HFApp.SqlExec(s)
    End If

    Call UpdateScheduleVendor(Job, DbQuote(Str, OldVendor), DbQuote(Str, NewVendor))

    '-----------------------------
    'notify original vendor
    '-----------------------------
    Set rs = HFApp.SqlExec("SELECT 1 from tblvendors where isnull(isTBD,0)=1 and divisionid=" & DbQuote(Num, HFApp.DivisionID) & " and vendor_id=" & DbQuote(Str, OldVendor))
    If rs.EOF Then
        If vbYes = MsgBox("Purchase order """ & PONumber & """ has been re-assigned to """ & NewVendorName & """." & vbCrLf & "Please notify """ & OldVendorName & """ of the change." & vbCrLf & vbCrLf & "Would you like to send them an email now?", vbQuestion + vbYesNo, "Purchase Order Re-assigned") Then
            If HFApp.Options(SendPO_FromCurrentUsersEmail) Then
                FromAddress = "" & HFApp.SqlExec("Select email from user_manager where User_id = " & DbQuote(Str, HFApp.LoginID))(0)
                Call HFApp.SendMail(True, EmailAddr, "", Replace(Replace(Replace(HFApp.Options(Msg_CancelPO_Subject), "<%PONumber%>", PONumber), vbQuote, ""), vbCrLf, " "), Replace(Replace(Replace(HFApp.Options(Msg_CancelPO_Body), "<%PONumber%>", PONumber), vbQuote, ""), vbCrLf, " "), "", FromAddress)
            Else
                Call HFApp.SendMail(True, EmailAddr, "", Replace(Replace(Replace(HFApp.Options(Msg_CancelPO_Subject), "<%PONumber%>", PONumber), vbQuote, ""), vbCrLf, " "), Replace(Replace(Replace(HFApp.Options(Msg_CancelPO_Body), "<%PONumber%>", PONumber), vbQuote, ""), vbCrLf, " "), "")
            End If
        End If
    End If
    
    ChangePOVendor = True
End Function

Public Sub RollPrices(prompt As Boolean)
    Dim s  As String
    Dim rs As Recordset
    Dim rc As Long
    Dim AllRows   As Long
    Dim ZeroRows  As Long
    
    s = ""
    s = s & "select isnull(sum(case when next_effective1<=CONVERT(datetime, getdate()) and isnull(next_cost1,0)=0 then 1 else 0 end),0) ZeroRows1" & vbCrLf
    s = s & "      ,isnull(sum(case when next_effective2<=CONVERT(datetime, getdate()) and isnull(next_cost2,0)=0 then 1 else 0 end),0) ZeroRows2" & vbCrLf
    s = s & "      ,count(*) AllRows" & vbCrLf
    s = s & "from tblvendorcost" & vbCrLf
    s = s & "where (next_effective1<=CONVERT(datetime, getdate()) or next_effective2<=CONVERT(datetime, getdate())) " & vbCrLf
    s = s & " and (divisionid=0 or DivisionID = " & HFApp.DivisionID & ")"
    Set rs = HFApp.SqlExec(s)
    
    If rs.EOF Then Exit Sub

    ZeroRows = Max(Val("" & rs(0)), Val("" & rs(1)))
    AllRows = Val("" & rs(2))
    
    If AllRows = 0 Then
        If prompt = False Then
            'nothing to update but was called from menu so give message
            Call MsgBox("No prices have expired. Your pricing database is current.", vbInformation + vbOKOnly, "Update Pricing")
        End If
    Else
        If ZeroRows > 0 Then
            rc = MsgBox("Pricing has expired for " & AllRows & " items in your database. For " & ZeroRows & vbCrLf & _
                        "items the new price is $0.00. Do you want to set the cost price" & vbCrLf & _
                        "of these items to $0.00?", vbQuestion + vbYesNoCancel, "Update Pricing")
        Else
            If prompt Then
                rc = MsgBox("Pricing has expired for " & AllRows & " items in your database. Do you" & vbCrLf & "want to update these prices now?", vbQuestion + vbOKCancel, "Update Pricing")
            Else
                rc = vbYes
            End If
        End If
        
        Select Case rc
            Case vbCancel
                Exit Sub
            
            Case vbYes
                Screen.MousePointer = vbHourglass
                s = ""
                s = s & "update tblVendorCost " & vbCrLf
                s = s & "   set last_cost3      = isnull(last_cost2,0)" & vbCrLf
                s = s & "      ,last3_expiry    = last2_expiry" & vbCrLf
                s = s & "      ,last_cost2      = isnull(last_cost1,0)" & vbCrLf
                s = s & "      ,last2_expiry    = last1_expiry" & vbCrLf
                s = s & "      ,last_cost1      = isnull(current_cost,0)" & vbCrLf
                s = s & "      ,last1_expiry    = cast(floor(cast(getdate() as float)) as datetime)" & vbCrLf
                s = s & "      ,Current_Cost    = isnull(next_Cost1,0)" & vbCrLf
                s = s & "      ,next_cost1      = isnull(next_cost2,0)" & vbCrLf
                s = s & "      ,next_effective1 = next_effective2" & vbCrLf
                s = s & "      ,next_cost2      = 0" & vbCrLf
                s = s & "      ,next_effective2 = NULL" & vbCrLf
                s = s & "where (next_effective1< = cast(floor(cast(getdate() as float)) as datetime)" & vbCrLf
                s = s & "   or  next_effective2< = cast(floor(cast(getdate() as float)) as datetime))" & vbCrLf
                s = s & " and (divisionid=0 or DivisionID =" & HFApp.DivisionID & ")"
                Call HFApp.SqlExec(s)
                Call HFApp.SqlExec(s)
                Screen.MousePointer = vbDefault
                Call MsgBox(AllRows & " items have been updated. Your pricing database is current.", vbInformation + vbOKOnly, "Update Pricing")
                
            Case vbNo, vbOK
                Screen.MousePointer = vbHourglass
                s = ""
                s = s & "update tblVendorCost " & vbCrLf
                s = s & "   set last_cost3      = isnull(last_cost2,0)" & vbCrLf
                s = s & "      ,last3_expiry    = last2_expiry" & vbCrLf
                s = s & "      ,last_cost2      = isnull(last_cost1,0)" & vbCrLf
                s = s & "      ,last2_expiry    = last1_expiry" & vbCrLf
                s = s & "      ,last_cost1      = isnull(current_cost,0)" & vbCrLf
                s = s & "      ,last1_expiry    = cast(floor(cast(getdate() as float)) as datetime)" & vbCrLf
                s = s & "      ,Current_Cost    = isnull(next_Cost1,0)" & vbCrLf
                s = s & "      ,next_cost1      = isnull(next_cost2,0)" & vbCrLf
                s = s & "      ,next_effective1 = next_effective2" & vbCrLf
                s = s & "      ,next_cost2      = 0" & vbCrLf
                s = s & "      ,next_effective2 = NULL" & vbCrLf
                s = s & "where (next_effective1< = cast(floor(cast(getdate() as float)) as datetime)" & vbCrLf
                s = s & "    or next_effective1< = cast(floor(cast(getdate() as float)) as datetime))" & vbCrLf
                s = s & "  and Next_cost1<>0" & vbCrLf
                s = s & " and (divisionid=0 or DivisionID =" & HFApp.DivisionID & ")"
                Call HFApp.SqlExec(s)
                Call HFApp.SqlExec(s)
                Screen.MousePointer = vbDefault
                Call MsgBox(AllRows - ZeroRows & " items have been updated. Your pricing database is current.", vbInformation + vbOKOnly, "Update Pricing")
            
        End Select
    End If

End Sub



Public Sub FixGroupPhaseValue()
    Dim s As String
    
    
    s = ""
    s = s & "update pp " & vbCrLf
    s = s & "set GroupPhaseValue=gg.phase" & vbCrLf
    s = s & "from" & vbCrLf
    s = s & "( " & vbCrLf
    s = s & "  select p.divisionid,max(g.sortorder) groupsort,p.phase" & vbCrLf
    s = s & "  from tblestphases p" & vbCrLf
    s = s & "  join estgroups g on g.divisionid=p.divisionid and g.sortorder<p.sortorder" & vbCrLf
    s = s & "  where p.groupphase=0" & vbCrLf
    s = s & "  group by p.divisionid,p.sortorder,p.phase" & vbCrLf
    s = s & ") x" & vbCrLf
    s = s & "join tblestphases pp on pp.divisionid=x.divisionid and pp.phase=x.phase" & vbCrLf
    s = s & "join estgroups gg on gg.divisionid=x.divisionid and gg.sortorder=x.groupsort" & vbCrLf
    s = s & "where x.DivisionID = " & HFApp.DivisionID & vbCrLf
    
    s = s & "update tblestphases" & vbCrLf
    s = s & "set GroupPhaseValue=Phase" & vbCrLf
    s = s & "where DivisionID = " & HFApp.DivisionID & " and GroupPhase=1" & vbCrLf
    
    s = s & "update tblestphases " & vbCrLf
    s = s & "set groupphasevalue=(select min(phase) from tblestphases where divisionid=" & HFApp.DivisionID & " and groupphase=1)" & vbCrLf
    s = s & "where groupphasevalue is null" & vbCrLf
    
    Call HFApp.SqlExec(s, dbHomefront)

End Sub


Public Function MbApiIsRunning() As Boolean
On Error Resume Next
    Dim MB As Object
    
    If HFApp.Options.ValueByName("Sage100APILevel") = "v18" Then
        Set MB = CreateObject("MBAPI.IMBXML")
        MbApiIsRunning = Not MB Is Nothing
    Else
        MbApiIsRunning = True
    End If
    
End Function

Public Function StripFormating(s As String) As String
    Const FORMATCHRS = "+=_-)(*&^%$#@!~`[]{}\|/?'""<>;:,"
    Dim i As Long
    For i = 1 To Len(FORMATCHRS)
        s = Replace(s, Mid(FORMATCHRS, i, 1), "")
    Next
    StripFormating = s
End Function


Public Function CalcPreTax(AssemblyType As AssemblyTypes, ByVal AfterTaxPrice As Double, Optional WARN As Boolean = False) As Double

    Dim s As String
    Dim rs As ADODB.Recordset


    If AssemblyType = atModel Then

        s = ""
        s = s & "select dbo.CalculatePretax(" & vbCrLf
        s = s & " " & DbQuote(Num, AfterTaxPrice) & vbCrLf
        s = s & ",''" & vbCrLf
        s = s & "," & DbQuote(Str, HFApp.Options(PST_Province)) & vbCrLf
        s = s & "," & DbQuote(Num, Val("" & HFApp.Options(GST_Rate))) & vbCrLf
        If HFApp.Options(UsePst) Then
            s = s & "," & DbQuote(Num, Val("" & HFApp.Options(PST_Rate))) & vbCrLf
        Else
            s = s & ",0" & vbCrLf
        End If
        s = s & ")"
        Set rs = HFApp.SqlExec(s, dbHomefront)
        If Not rs.EOF Then
            CalcPreTax = Round(Val("" & rs(0)), 2)
        Else
            If WARN Then MsgBox "Tax tables not setup for this rate." & vbCrLf & "Rebate cannot be calculated.", vbExclamation, App.ProductName
            CalcPreTax = AfterTaxPrice / (100 + Val("" & HFApp.Options(GST_Rate)) + Val("" & HFApp.Options(PST_Rate))) * 100
        End If
    Else
        CalcPreTax = AfterTaxPrice / (100 + Val("" & HFApp.Options(GST_Rate)) + Val("" & HFApp.Options(PST_Rate))) * 100
    End If
    

    
End Function

Public Function CalcTaxRebate(ByVal PretaxPrice As Double) As Double
    Dim s As String
    Dim rs As ADODB.Recordset
    s = "select dbo.CalculateFederalRebate(" & DbQuote(Num, PretaxPrice) & ") + dbo.CalculateProvincialRebate(" & DbQuote(Num, PretaxPrice) & ")"
    Set rs = HFApp.SqlExec(s)
    If Not rs.EOF Then
    If HFApp.Options(Country) <> "AU" Then
        CalcTaxRebate = Round(Val("" & rs(0)), 2)
    Else
        CalcTaxRebate = 0
    End If
    End If
End Function

Public Function CalcTax(AssemblyType As AssemblyTypes, ByVal PretaxPrice As Double) As Double
    
    Dim ftax As Double
    Dim ptax As Double
        
    ftax = Val(HFApp.Options(GST_Rate))
    ptax = Val(HFApp.Options(PST_Rate))
    If Not HFApp.Options(UsePst) Then ptax = 0
    
    If AssemblyType = atModel Then
        CalcTax = Round(PretaxPrice * (ftax + ptax) / 100, 2) - CalcTaxRebate(PretaxPrice)
    Else
        CalcTax = Round(PretaxPrice * (ftax + ptax) / 100, 2)
    End If
    
    
End Function

Public Function ISDEBUG() As Boolean
  ISDEBUG = App.EXEName Like "*DEBUG*"
End Function

Private Function CancelPO_TL(PONumber As String, reason As String) As Boolean
On Error GoTo eh
    Dim rs As Recordset
    Dim s As String
    
    Dim tlco As Long
    Dim hfco As Long
    Dim co As Long


    Call HFApp.WriteAuditLog("Edit PO", "Cancel posted PO """ & PONumber)

    'get next co number (greatest numeric value)
    On Error Resume Next
    
    
    tlco = Val("" & HFApp.SqlExec("select max(commitment_co) from jcm_master__commitment_co where commitment_co<'A' and Commitment=" & DbQuote(Str, PONumber), dbAccounting)(0))
    hfco = Val("" & HFApp.SqlExec("select max(ChangeOrder) from pochangeorders where changeorder<'A' and ponumber=" & DbQuote(Str, PONumber), dbHomefront)(0))
    co = Max(tlco, hfco) + 1
    On Error GoTo eh
    
    'write change order
    s = ""
    s = s & "INSERT INTO POChangeOrders(PONumber,DivisionID,ChangeOrder,Description,CODate,UStmp,TStmp,PostingBatch)" & vbCrLf
    s = s & "SELECT PONumber" & vbCrLf
    s = s & "      ," & HFApp.DivisionID
    s = s & "      ," & DbQuote(Str, co) & vbCrLf
    s = s & "      ,'po cancelled'" & vbCrLf
    s = s & "      ,GETDATE()" & vbCrLf
    s = s & "      ," & DbQuote(Str, HFApp.LoginID) & vbCrLf
    s = s & "      ,GETDATE()" & vbCrLf
    s = s & "      ,0" & vbCrLf
    s = s & "  FROM POMaster" & vbCrLf
    s = s & " WHERE DivisionID = " & HFApp.DivisionID & " and PONumber=" & DbQuote(Str, PONumber) & vbCrLf
    Call HFApp.SqlExec(s, dbHomefront)
    
    'write change order items
    s = ""
    s = s & "SELECT " & vbCrLf
    s = s & " description" & vbCrLf
    s = s & ",units + approved_commitment_co_units - units_invoiced as Qty" & vbCrLf
    s = s & ",unit_description UOM" & vbCrLf
    s = s & ",unit_cost Rate" & vbCrLf
    s = s & ",amount-amount_invoiced-tax-approved_commitment_co_tax_amount Pretax" & vbCrLf
    s = s & ",tax_group TaxGroup" & vbCrLf
    s = s & ",tax+approved_commitment_co_tax_amount jctax" & vbCrLf
    s = s & ",Job" & vbCrLf
    s = s & ",extra JCExtra" & vbCrLf
    s = s & ",cost_code JCCostCode" & vbCrLf
    s = s & ",category JCCategory" & vbCrLf
    s = s & ",item_number LineNumber" & vbCrLf
    s = s & "FROM JCM_Master__Commitment_Item" & vbCrLf
    s = s & "WHERE Commitment=" & DbQuote(Str, PONumber) & vbCrLf
    s = s & "ORDER BY item_number" & vbCrLf
    Set rs = HFApp.SqlExec(s, dbAccounting)
    While Not rs.EOF
        s = ""
        s = s & "insert into pochangeorderitems(DivisionID, PONumber, ChangeOrder, Description, Qty, UOM, Rate, Pretax, TaxGroup, JCTax, Job, JCExtra, JCCostCode, JCCategory, LineNumber)" & vbCrLf
        s = s & "values(" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        s = s & "      ," & DbQuote(Str, PONumber) & vbCrLf
        s = s & "      ," & DbQuote(Str, co) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("Description")) & vbCrLf
        s = s & "      ," & DbQuote(Num, -1 * Val("" & rs("Qty"))) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("UOM")) & vbCrLf
        s = s & "      ," & DbQuote(Num, 1 * Val("" & rs("Rate"))) & vbCrLf
        s = s & "      ," & DbQuote(Num, -1 * Val("" & rs("Pretax"))) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("TaxGroup")) & vbCrLf
        s = s & "      ," & DbQuote(Num, -1 * Val("" & rs("jctax"))) & vbCrLf
        s = s & "      ," & DbQuote(Str, StripFormating("" & rs("job"))) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("jcextra")) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("jccostcode")) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("jccategory")) & vbCrLf
        s = s & "      ," & DbQuote(Num, "" & rs("linenumber")) & vbCrLf
        s = s & ")"
        Call HFApp.SqlExec(s, dbHomefront)
        rs.MoveNext
    Wend
    
    
    CancelPO_TL = True
Exit Function
eh: Call errHandler(SRCFILE & "CancelPO_TL")
End Function

Private Function CancelPO_QB(PONumber As String, reason As String) As Boolean
On Error GoTo eh
    Dim rs As Recordset
    Dim s As String
    
    Dim TxnID As String
    Dim EditSeq As String
    
    
    s = "select * from pomaster where DivisionID = " & HFApp.DivisionID & " and ponumber=" & DbQuote(Str, PONumber)
    Set rs = HFApp.SqlExec(s)
    If rs.EOF Then
        MsgBox "PO number """ & PONumber & """ not found", vbExclamation, "Error"
        Exit Function
    End If
    
    'get txnid
    s = HFApp.XmlQBStart()
    s = s & "<GeneralDetailReportQueryRq>" & vbCrLf
    s = s & "<GeneralDetailReportType>OpenPOs</GeneralDetailReportType>" & vbCrLf
    s = s & "<ReportPeriod>" & vbCrLf
    s = s & HFApp.XmlQBAdd(d, 0, "FromReportDate", "" & rs("PODate"))
    s = s & HFApp.XmlQBAdd(d, 0, "ToReportDate", "" & rs("PODate"))
    s = s & "</ReportPeriod>" & vbCrLf
    s = s & "<IncludeColumn>RefNumber</IncludeColumn>" & vbCrLf
    s = s & "<IncludeColumn>TxnID</IncludeColumn>" & vbCrLf
    s = s & "<IncludeColumn>Amount</IncludeColumn>" & vbCrLf
    s = s & "<IncludeColumn>Name</IncludeColumn>" & vbCrLf
    s = s & "<IncludeColumn>Date</IncludeColumn>" & vbCrLf
    s = s & "</GeneralDetailReportQueryRq>" & vbCrLf
    s = s & HFApp.XmlQBEnd()
    s = HFApp.XmlQBSubmit(s)
    Set rs = HFApp.SqlExec("exec qb_POtxnid " & DbQuote(Str, s) & ", " & DbQuote(Str, PONumber), dbHomefront)
    If rs.EOF Then
        MsgBox "This PO was not found in Quickbooks, it may have been changed in Quickbooks.", vbExclamation, App.ProductName
    Else
        TxnID = "" & rs(0)
        
        'get edit seq and received qty
        s = HFApp.XmlQBStart()
        s = s & "<PurchaseOrderQueryRq>" & vbCrLf
        s = s & "<TxnID>" & "</TxnID>" & vbCrLf
        s = s & HFApp.XmlQBAdd(st, 40, "TxnID", TxnID)
        s = s & "<IncludeLineItems>true</IncludeLineItems>" & vbCrLf
        s = s & "</PurchaseOrderQueryRq>" & vbCrLf
        s = s & HFApp.XmlQBEnd()
        s = HFApp.XmlQBSubmit(s)
        EditSeq = Parse(s, 2, "EditSequence>")
        EditSeq = Parse(EditSeq, 1, "<")
        
        'cancel po
        s = HFApp.XmlQBStart()
        s = s & "<PurchaseOrderModRq>" & vbCrLf
        s = s & "   <PurchaseOrderMod>" & vbCrLf
        s = s & HFApp.XmlQBAdd(st, 40, "TxnID", TxnID)
        s = s & HFApp.XmlQBAdd(st, 40, "EditSequence", EditSeq)
        s = s & "      <IsManuallyClosed> 1 </IsManuallyClosed>" & vbCrLf
        s = s & HFApp.XmlQBAdd(st, 500, "Memo", "closed by " & HFApp.LoginID & ", " & Now())
        s = s & "   </PurchaseOrderMod>" & vbCrLf
        s = s & "</PurchaseOrderModRq>" & vbCrLf
        s = s & HFApp.XmlQBEnd()
        s = HFApp.XmlQBSubmit(s)
        
    End If
    

    
    CancelPO_QB = True
    
Exit Function
eh: Call errHandler(SRCFILE & "CancelPO_QB")
End Function


Private Function CancelPO_MB(PONumber As String, reason As String) As Boolean
On Error GoTo eh
    Dim rs As Recordset
    Dim s As String
    Dim MBObjectID As Long
    Dim DocType As String
    Dim AmtPd   As Double
    
    s = ""
    s = s & "select p.PONumber,p.POIndex" & vbCrLf
    s = s & "  from POMaster p" & vbCrLf
    s = s & " Where p.DivisionID = " & HFApp.DivisionID & vbCrLf
    s = s & "   and p.ponumber = " & DbQuote(Str, PONumber)
    Set rs = HFApp.SqlExec(s)
    If rs.EOF Then
        MsgBox "What! How did you get here? PO number """ & PONumber & """ not found", vbExclamation, App.ProductName
        Exit Function
    Else
    
        s = ""
        s = s & "select recnum,'po' doctype,rcvdte amtpd " & vbCrLf
        s = s & "from pchord " & vbCrLf
        s = s & "where ordnum = " & DbQuote(Str, PONumber) & vbCrLf
        s = s & "union" & vbCrLf
        s = s & "select recnum,'contract' doctype,invttl amtpd " & vbCrLf
        s = s & "from subcon " & vbCrLf
        s = s & "where ctcnum = " & DbQuote(Str, PONumber) & vbCrLf
        Set rs = HFApp.SqlExec(s, dbAccounting)
        
        If rs.EOF Then
            MsgBox "PO number """ & PONumber & """ could not be found in the Sage 100 database.", vbExclamation, App.ProductName
        Else
        
            MBObjectID = Val("" & rs("recnum"))
            DocType = Trim("" & rs("doctype"))  '<-- trim seems unnecessary but the vfpoledb driver pads 'po' so it is the same width as 'contract'
            AmtPd = Val("" & rs("amtpd"))
            If AmtPd <> 0 Then
                MsgBox "PO " & PONumber & " has been invoiced in Sage 100. It cannot be cancelled.", vbOKOnly + vbInformation, App.ProductName
                Exit Function
            End If
            
            s = HFApp.XmlMbStart(HFApp.Options(MasterBuilderCompany), HFApp.Options(MasterBuilderUID))
            If DocType = "po" Then
                s = s & "<PurchaseOrderDelRq requestID=""1"">" & vbCrLf
                s = s & "<ObjectRef>" & vbCrLf
                s = s & "<ObjectID>" & MBObjectID & "</ObjectID>" & vbCrLf
                s = s & "</ObjectRef>" & vbCrLf
                s = s & "</PurchaseOrderDelRq>" & vbCrLf
            Else
                s = s & "<SubcontractDelRq requestID=""1"">" & vbCrLf
                s = s & "<ObjectRef>" & vbCrLf
                s = s & "<ObjectID>" & MBObjectID & "</ObjectID>" & vbCrLf
                s = s & "</ObjectRef>" & vbCrLf
                s = s & "</SubcontractDelRq>" & vbCrLf
            End If
            s = s & HFApp.XmlMBEnd()
            Call HFApp.XmlMbSubmit(s, HFApp.Options(MasterBuilderPWD))
            
        End If
    End If
            
    
    CancelPO_MB = True
    
Exit Function
eh: Select Case Err.Source
    Case "SubmitMBXml":  MsgBox "Unable to cancel PO number """ & PONumber & """" & vbCrLf & vbCrLf & Err.Description, vbCritical, App.ProductName
    Case Else:           Call errHandler(SRCFILE & "CancelPO_MB", s)
    End Select
End Function




Private Function ChangeQBPOVendor(PONumber As String, NewVendor As String) As Boolean
On Error GoTo eh
    Dim rs As Recordset
    Dim s As String
    
    Dim TxnID As String
    Dim EditSeq As String
    Dim RcvdQty As String
    
    s = "select * from pomaster where divisionid=" & DbQuote(Num, HFApp.DivisionID) & " and ponumber=" & DbQuote(Str, PONumber)
    Set rs = HFApp.SqlExec(s)
    If rs.EOF Then
        MsgBox "PO number """ & PONumber & """ not found", vbExclamation, "Error"
        Exit Function
    End If
    
    If Val("" & rs("postingbatch")) = 0 Then
        'not posted yet
        ChangeQBPOVendor = True
        Exit Function
    End If
    
    'get txnid
    s = HFApp.XmlQBStart()
    s = s & "<GeneralDetailReportQueryRq>" & vbCrLf
    s = s & "<GeneralDetailReportType>OpenPOs</GeneralDetailReportType>" & vbCrLf
    s = s & "<ReportPeriod>" & vbCrLf
    s = s & HFApp.XmlQBAdd(d, 0, "FromReportDate", "" & rs("PODate"))
    s = s & HFApp.XmlQBAdd(d, 0, "ToReportDate", "" & rs("PODate"))
    s = s & "</ReportPeriod>" & vbCrLf
    s = s & "<IncludeColumn>RefNumber</IncludeColumn>" & vbCrLf
    s = s & "<IncludeColumn>TxnID</IncludeColumn>" & vbCrLf
    s = s & "<IncludeColumn>Amount</IncludeColumn>" & vbCrLf
    s = s & "<IncludeColumn>Name</IncludeColumn>" & vbCrLf
    s = s & "<IncludeColumn>Date</IncludeColumn>" & vbCrLf
    s = s & "</GeneralDetailReportQueryRq>" & vbCrLf
    s = s & HFApp.XmlQBEnd()
    s = HFApp.XmlQBSubmit(s)
    Set rs = HFApp.SqlExec("exec qb_POtxnid " & DbQuote(Str, s) & ", " & DbQuote(Str, PONumber), dbHomefront)
    If rs.EOF Then
        MsgBox "This PO was not found in Quickbooks, it may have been posted in Quickbooks.", vbExclamation, App.ProductName
        ChangeQBPOVendor = False
        Exit Function
    Else
        TxnID = "" & rs(0)
        
        'get edit seq and received qty
        s = HFApp.XmlQBStart()
        s = s & "<PurchaseOrderQueryRq>" & vbCrLf
        s = s & "<TxnID>" & "</TxnID>" & vbCrLf
        s = s & HFApp.XmlQBAdd(st, 40, "TxnID", TxnID)
        s = s & "<IncludeLineItems>true</IncludeLineItems>" & vbCrLf
        s = s & "</PurchaseOrderQueryRq>" & vbCrLf
        s = s & HFApp.XmlQBEnd()
        s = HFApp.XmlQBSubmit(s)
        EditSeq = Parse(s, 2, "EditSequence>")
        EditSeq = Parse(EditSeq, 1, "<")
        
        RcvdQty = Parse(s, 2, "ReceivedQuantity>")
        RcvdQty = Parse(RcvdQty, 1, "<")
        If Val("" & RcvdQty) = 0 Then
            'change vendor on the po
            s = HFApp.XmlQBStart()
            s = s & "<PurchaseOrderModRq>" & vbCrLf
            s = s & "   <PurchaseOrderMod>" & vbCrLf
            s = s & HFApp.XmlQBAdd(st, 40, "TxnID", TxnID)
            s = s & HFApp.XmlQBAdd(st, 40, "EditSequence", EditSeq)
            s = s & "<VendorRef>" & HFApp.XmlQBAdd(st, 40, "ListID", "" & NewVendor) & "</VendorRef>" & vbCrLf
            s = s & "   </PurchaseOrderMod>" & vbCrLf
            s = s & "</PurchaseOrderModRq>" & vbCrLf
            s = s & HFApp.XmlQBEnd()
            s = HFApp.XmlQBSubmit(s)
        Else
            MsgBox "This PO Has had invoices applied against it already, unable to change the vendor.", vbExclamation, App.ProductName
            ChangeQBPOVendor = False
            Exit Function
        End If
        
    End If
    ChangeQBPOVendor = True

    
Exit Function
eh: Call errHandler(SRCFILE & "ChangeQBPOVendor")
End Function


Public Function WriteToFile(FileName As String, Description As String)
On Error GoTo errHandler
      
Open FileName For Append As #1
Write #1, Description
Close #1
Screen.MousePointer = vbDefault
       
Exit Function
errHandler:
If Err.Number <> 0 Then
    Resume Next
End If
End Function

Public Function MyLookUp(ByVal MyTable As String, MyOutPutField As String, MyKeyField1 As String, MyLookUpVal1 As String, Optional MyDataType1 As String, Optional MyKeyField2 As String, Optional MyLookUpVal2 As String, Optional MyDataType2 As String, Optional MyKeyField3 As String, Optional MyLookUpVal3 As String, Optional MyDataType3 As String, Optional MyKeyField4 As String, Optional MyLookUpVal4 As String, Optional MyDataType4 As String)
    On Error Resume Next
    
    Dim MyCursor As New ADODB.Connection
    Dim MyRec As New ADODB.Recordset
    Dim MyCriteria As String

    Set MyCursor = New ADODB.Connection
    If Mid(MyTable, 1, 6) = "MASTER" And MyTable <> "Master_Date" Then
       MyCursor.ConnectionString = "Provider=MSDASQL.1;DRIVER={Timberline Data};" & "DBQ=" & HFApp.Options(Timberline_Data_Path) & ";UID=" & HFApp.Options(Timberline_UID) & ";PWD=" & HFApp.Options(Timberline_PWD) & ";CODEPAGE=1252;DictionaryMode=0;StandardMode=0;MaxColSupport=1536;ShortenNames=0;;DatabaseType=1;"
    Else
       MyCursor.ConnectionString = HFApp.ConnectionString(dbHomefront)
    End If
    
    MyCursor.Open
    Set MyRec = New ADODB.Recordset
    MyRec.LockType = adLockOptimistic
    MyRec.CursorType = adOpenDynamic
    MyRec.CursorLocation = adUseClient
    Set MyRec.ActiveConnection = MyCursor
    MyCriteria = ""
    
    If MyKeyField1 <> "" Then
       MyCriteria = MyKeyField1 & "=" & IIf(MyDataType1 = "S", "'" & MyLookUpVal1 & "'", MyLookUpVal1)
    End If
    If MyKeyField2 <> "" Then
       MyCriteria = MyCriteria & " and " & MyKeyField2 & "=" & IIf(MyDataType2 = "S", "'" & MyLookUpVal2 & "'", MyLookUpVal2)
    End If
    If MyKeyField3 <> "" Then
       MyCriteria = MyCriteria & " and " & MyKeyField3 & "=" & IIf(MyDataType3 = "S", "'" & MyLookUpVal3 & "'", MyLookUpVal3)
    End If
    If MyKeyField4 <> "" Then
       MyCriteria = MyCriteria & " and " & MyKeyField4 & "=" & IIf(MyDataType4 = "S", "'" & MyLookUpVal4 & "'", MyLookUpVal4)
    End If
    
    MyRec.Source = "Select " & MyOutPutField & " from " & MyTable & " where " & MyCriteria
    MyRec.Open
    If Not MyRec.EOF Then
       MyLookUp = IIf(Not IsNull(MyRec.fields(0).value), MyRec.fields(0).value, "")
    End If
End Function

Public Function RegGetKey(ByVal hive As RegistryHiveConstants, ByVal section As String, Optional ByVal key As String, Optional ByVal Default As String = "") As String
   ' Section   Required. String expression containing the name of the section WHERE the key setting is found.
   '           If omitted, key setting is assumed to be in default subkey.
   ' Key       Required. String expression containing the name of the key setting to return.
   ' Default   Optional. Expression containing the Value to return if no Value is set in the key setting.
   '           If omitted, default is assumed to be a zero-length string ("").
   Dim nRet As Long
   Dim hKey As Long
   Dim nType As Long
   Dim nBytes As Long
   Dim Buffer As String
   
   ' Assume failure AND set return to Default
   RegGetKey = Default

   ' Open key
   nRet = RegOpenKeyEx(hive, section, 0&, KEY_READ, hKey)
   If nRet = ERROR_SUCCESS Then
      ' Set appropriate Value for default query
      If key = "*" Then key = vbNullString
      
      ' Determine how large the buffer needs to be
      nRet = RegQueryValueEx(hKey, key, 0&, nType, ByVal Buffer, nBytes)
      If nRet = ERROR_SUCCESS Then
         ' Build buffer AND get data
         If nBytes > 0 Then
            Buffer = Space(nBytes)
            nRet = RegQueryValueEx(hKey, key, 0&, nType, ByVal Buffer, Len(Buffer))
            If nRet = ERROR_SUCCESS Then
               ' Trim NULL AND return successful query!
               RegGetKey = Left(Buffer, nBytes - 1)
            End If
         End If
      Call RegCloseKey(hKey)
      End If
   End If
End Function

Public Function RegEnumKeys(ByVal hive As RegistryHiveConstants, ByVal section As String, ByRef keys() As String) As Boolean
    Dim iKeyCount As Long
    Dim lResult As Long
    Dim hKey As Long
    Dim sName As String
    Dim lNameSize As Long
    Dim sData As String
    Dim lIndex As Long
    Dim cJunk As Long
    Dim cNameMax As Long
    Dim ft As Currency
   
    iKeyCount = 0
    ReDim keys(0)
     
    lIndex = 0
    
    
    'wake up the stupid registry by setting the default value to nothing
    Call RegSaveKey(hive, section)
        
    lResult = RegOpenKeyEx(hive, section, 0, KEY_QUERY_VALUE, hKey)
    
    If (lResult = ERROR_SUCCESS) Then
        ' Log "OpenedKey:" & hive & "," & section
        lResult = RegQueryInfoKey(hKey, "", cJunk, 0, cJunk, cJunk, cJunk, iKeyCount, cNameMax, cJunk, cJunk, ft)
        Do While lResult = ERROR_SUCCESS
    
            'Set buffer space
            lNameSize = cNameMax + 1
            sName = String$(lNameSize, 0)
            If (lNameSize = 0) Then lNameSize = 1
            
            ' Log "Requesting Next Value"
          
            'Get Value name:
            lResult = RegEnumValue(hKey, lIndex, sName, lNameSize, 0&, 0&, 0&, 0&)
            ' Log "RegEnumValue returned:" & lResult
            If (lResult = ERROR_SUCCESS) Then
        
                ' Although in theory you can also retrieve the actual
                ' Value AND type here, I found it always (ultimately) resulted in
                ' a GPF, on Win95 AND NT.  Why?  Can anyone help?
        
                lIndex = lIndex + 1
                sName = Left$(sName, lNameSize)
                ' Log "Enumerated Value:" & sName
                  
                iKeyCount = iKeyCount + 1
                ReDim Preserve keys(iKeyCount) As String
                keys(iKeyCount) = sName
            End If
        Loop
    End If
    If (hKey <> 0) Then
        RegCloseKey hKey
    End If

    RegEnumKeys = True
End Function


Public Function RegSectionExists(ByVal hive As RegistryHiveConstants, ByVal section As String) As Boolean
    
    Dim hKey As Long
    Dim nRet As Long
    
    nRet = RegOpenKeyEx(hive, section, 0&, KEY_READ, hKey)
    
    If nRet = ERROR_SUCCESS Then
        RegSectionExists = True
        RegCloseKey hKey
    Else
        RegSectionExists = False
    End If
    
End Function

Public Function MachineName() As String
    
    'returns the local computer's name
    
    Const MAX_COMPUTERNAME_LENGTH = 15
    Dim s  As String
    Dim dl As Long
    
    s = String(MAX_COMPUTERNAME_LENGTH + 1, 0)
    dl = GetComputerName(s, MAX_COMPUTERNAME_LENGTH + 1)
    
    MachineName = Replace(s, Chr(0), "")
    
End Function

Public Function IniGet(FileName As String, SectionName As String, VariableName As String, Optional DefaultValue As Variant) As Variant
    
    'retrieves a Value FROM an ini file
    
    Dim VariableValue As String
    Dim lnumChar As Long
    
    VariableValue = String(128, 0)
    
    'DefaultValues to ""
    lnumChar = GetPrivateProfileString(SectionName, VariableName, "", VariableValue, 127, FileName)

    If lnumChar >= 1 Then 'Null appended to string; even null string
        Select Case TypeName(DefaultValue)
            Case "Byte"
                IniGet = CByte(Left$(VariableValue, lnumChar))
            Case "Integer", "Long", "double", "double"
                IniGet = Val(Left$(VariableValue, lnumChar))
            Case "Date"
                IniGet = CDate(Left$(VariableValue, lnumChar))
            Case "Boolean"
                IniGet = CBool(Left$(VariableValue, lnumChar))
            Case Else 'DefaultValue to string
                IniGet = Left$(VariableValue, lnumChar)
        End Select
    
    Else 'DefaultValue applicable
        If IsMissing(DefaultValue) Then
            IniGet = ""
        Else
            IniGet = DefaultValue
        End If
    End If
    
End Function

Public Function IniGetSectionNames(FileName As String) As String

    'returns a csv list of the section names in file

    Dim rc     As Long
    Dim sBuff  As String

    sBuff = Space$(BUFFER_SIZE)
    rc = GetPrivateProfileSectionNames(sBuff, BUFFER_SIZE, FileName)
    If rc > 0 Then
        sBuff = Left(sBuff, rc - 1)
        sBuff = Replace(sBuff, Chr$(0), ",")
        If Right(sBuff, 1) = "," Then
            sBuff = Left(sBuff, Len(sBuff) - 1)
        End If
        IniGetSectionNames = sBuff
    End If
End Function

Public Function IniGetVariableNames(FileName As String, SectionName As String) As String
    
    'returns a csv list of the variable names in a section
    
    Dim sBuff As String
    Dim rc As Long
    sBuff = Space$(BUFFER_SIZE)
    rc = GetPrivateProfileString(SectionName, CLng(0), "", sBuff, BUFFER_SIZE, FileName)
    If rc > 0 Then
        sBuff = Left(sBuff, rc - 1)
        If Right(sBuff, 1) = "," Then
            sBuff = Left(sBuff, Len(sBuff) - 1)
        End If
        sBuff = Replace(sBuff, Chr$(0), ",")
    Else
        sBuff = ""
    End If
    IniGetVariableNames = sBuff
    
End Function

Public Sub IniPut(FileName As String, SectionName As String, VariableName As String, VariableValue As Variant)

    Const DFORMAT = "hh:mm mmmm d yyyy"

    
    Call CreatePath("", FilePath(FileName))
    
    'writes a Value to an ini file

    Dim tmpValue  As String
    Dim numChar   As Long
    
    Select Case TypeName(VariableValue)
        Case "Byte", "Integer", "Long", "double", "double"
            tmpValue = format$(VariableValue)
            
        Case "Date"
            tmpValue = format$(VariableValue, DFORMAT)
            
        Case "Boolean"
            If VariableValue Then
                tmpValue = "True"
            Else
                tmpValue = "False"
            End If
            
        Case Else 'DefaultValue to string
            tmpValue = VariableValue
            
    End Select
    
    numChar = WritePrivateProfileString(SectionName, VariableName, tmpValue, FileName)
End Sub

Public Sub IniRemove(FileName As String, SectionName As String, Optional VariableName As String)
    
    'removes a variable or a whole section from an ini file
    
    If VariableName = "" Then
        'remove section
        Call WritePrivateProfileString(SectionName, 0&, "", FileName)
    Else
        'remove variable
        Call WritePrivateProfileString(SectionName, VariableName, 0&, FileName)
    End If

End Sub

Public Function ItemInList(Item As String, List As String, Optional ByVal Delimeter As String = ",") As Boolean
    ItemInList = InStr(1, Delimeter & List & Delimeter, Delimeter & Item & Delimeter)
End Function

Public Function ParseDistinctList(List As String, Optional ByVal Delimeter As String = ",") As String
    Dim l As Variant
    Dim s As String
    Dim i As Long
    
    l = split(List, Delimeter)
    s = Delimeter
    For i = 0 To UBound(l)
        If Not s Like "*" & Delimeter & l(i) & Delimeter & "*" Then
            s = s & l(i) & Delimeter
        End If
    Next
    s = Mid(s, 1 + Len(Delimeter), Len(s) - (2 * Len(Delimeter)))
    ParseDistinctList = s
End Function

Public Function Parse(DelimitedString, Optional Index As Long, Optional ByVal Delimeter As String = ",") As Variant
    
    'Parses items out of a delimited string
    'Returns the number of items if index is omitted
    
    Dim tmpCount As Long
    Dim POs      As Long
    Dim i As Long
    Dim fStart As Long
    Dim fend As Long
    
    If Index < 0 Then
        Parse = ""
    ElseIf Index = 0 Then
        'return count
        tmpCount = 1
        POs = 0
        While InStr(POs + 1, DelimitedString, Delimeter, vbTextCompare)
            POs = InStr(POs + 1, DelimitedString, Delimeter, vbTextCompare)
            tmpCount = tmpCount + 1
        Wend
        Parse = tmpCount
    Else
        fStart = 0
        For i = 1 To Index - 1
             fStart = InStr(fStart + 1, DelimitedString, Delimeter, vbTextCompare)
             If fStart = 0 Then     'Index specified is greater than items in DelimitedString
                Parse = ""
                Exit Function
             End If
        Next
        fend = InStr(fStart + 1, DelimitedString, Delimeter, vbTextCompare)
        If fend = 0 Then
            fend = Len(DelimitedString)       'must be last item that we are getting
        Else
            fend = fend - 1          'don't include the trailing delimiter
        End If
'        Parse = Mid$(DelimitedString, fstart + 1, fend - fstart)
        If fStart = 0 Then
            Parse = Mid$(DelimitedString, 1, fend - fStart)
        Else
            Parse = Mid$(DelimitedString, fStart + Len(Delimeter), Max(0, fend - fStart - Len(Delimeter) + 1))
        End If
    End If
End Function
Public Function IsBetween(value, Minimum, Maximum, Optional Inclusive = True) As Boolean
    If Inclusive Then
        IsBetween = value >= Minimum And value <= Maximum
    Else
        IsBetween = value > Minimum And value < Maximum
    End If
End Function
Public Function Between(Minimum, Maximum, value)
    value = IIf(value > Maximum, Maximum, value)
    value = IIf(value < Minimum, Minimum, value)
    Between = value
End Function
Public Function Min(ParamArray items())
    Dim i As Long
    Dim v As Variant
    v = items(LBound(items))
    For i = LBound(items) + 1 To UBound(items)
        v = IIf(items(i) > v, v, items(i))
    Next
    Min = v
End Function
Public Function Max(ParamArray items())
    Dim i As Long
    Dim v As Variant
    v = items(LBound(items))
    For i = LBound(items) + 1 To UBound(items)
        v = IIf(items(i) > v, items(i), v)
    Next
    Max = v
 End Function


Public Function TempPath() As String
    Dim rc As Long
    Dim bufferLen As Long
    Dim bufferStr As String
        
    bufferLen = 255
    bufferStr = Space(bufferLen)
    rc = GetTempPath(bufferLen, bufferStr)
    If rc > 0 Then
        bufferStr = Trim(bufferStr)
        bufferStr = Left(bufferStr, Len(bufferStr) - 1)
    End If
    bufferStr = PathAppend(bufferStr, "HomeFront")
    Call CreatePath("", bufferStr)
    TempPath = bufferStr
    
End Function

Public Function TempFile(Optional Extension As String) As String
'
'   THIS STOPPED WORKING AT A CLIENT SITE SO I REPLACED IT WITH A GUID.
'   NO EXPLANATION WHY IT FAILS.
'
    Dim FileName As String
    FileName = CreateGUID
    FileName = Replace(FileName, "{", "")
    FileName = Replace(FileName, "}", "")
    If Extension <> "" Then FileName = ForceExt(FileName, Extension)
    
    TempFile = PathAppend(TempPath, FileName)


End Function

Public Function vbSingleQuote() As String
    vbSingleQuote = "'"
End Function
Public Function vbQuote() As String
    vbQuote = Chr(34)
End Function
Public Function vbBullet() As String
    vbBullet = Chr(149)
End Function

Public Function ShellFile(hwnd As Long, FileName As String, Optional DlgTitle, Optional FindItYourself As Boolean = True, Optional PrintIt As Boolean = False) As String
    'returns new FileName if they found the file they wanted.
    Dim bFailed As Boolean
    Dim bOkNow  As Boolean
    Dim sNewFileName As String
    Dim rc  As Long
    Dim msg As String
    Const Caption = "Error Launching File"
    
    sNewFileName = FileName
TryAgain:
    
    
    
    
    rc = ShellExecute(0, IIf(PrintIt, "print", ""), sNewFileName, vbNullString, "C:\", SW_SHOWNORMAL)
    If rc <= 32 Then
        bFailed = True
        bOkNow = False
        'There was an error
        Select Case rc
'            Case SE_ERR_FNF, SE_ERR_PNF:
'                If FindItYourself Then
'                    Msg = "Windows could not find the file '" & sNewFileName & "'. " & vbCrLf & "Would you like to locate the file yourself?"
'                    If VBA.Msgbox(Msg, vbYesNo + vbInformation + vbDefaultButton2, "Couldn't Find File") = vbYes Then
'                        If IsMissing(DlgTitle) Then
'                            If VBGetOpenFileName(sNewFileName, , , , , , , , , , , , OFN_NOVALIDATE) Then GoTo TryAgain
'                        Else
'                            If VBGetOpenFileName(sNewFileName, , , , , , , , , "" & DlgTitle, , , OFN_NOVALIDATE) Then GoTo TryAgain
'                        End If
'                    End If
'                Else
'                    VBA.Msgbox "File not found", vbOKOnly + vbExclamation, "" & DlgTitle
'                End If
            Case SE_ERR_ACCESSDENIED:     msg = "Access denied":                                           VBA.MsgBox msg, vbCritical, Caption
            Case SE_ERR_OOM:              msg = "Out of memory":                                           VBA.MsgBox msg, vbCritical, Caption
            Case SE_ERR_DLLNOTFOUND:      msg = "A required DLL could not be found":                       VBA.MsgBox msg, vbCritical, Caption
            Case SE_ERR_SHARE:            msg = "A sharing violation has occurred":                        VBA.MsgBox msg, vbCritical, Caption
            Case SE_ERR_ASSOCINCOMPLETE:  msg = "Incomplete or invalid file association":                  VBA.MsgBox msg, vbCritical, Caption
            Case SE_ERR_DDETIMEOUT:       msg = "DDE Time out":                                            VBA.MsgBox msg, vbCritical, Caption
            Case SE_ERR_DDEFAIL:          msg = "DDE transaction failed":                                  VBA.MsgBox msg, vbCritical, Caption
            Case SE_ERR_DDEBUSY:          msg = "DDE busy":                                                VBA.MsgBox msg, vbCritical, Caption
            Case SE_ERR_NOASSOC:          msg = "No association for file extension":                       VBA.MsgBox msg, vbCritical, Caption
            Case ERROR_BAD_FORMAT:        msg = "Invalid EXE file or error in EXE image":                  VBA.MsgBox msg, vbCritical, Caption
            Case Else:                    msg = "Unknown error":                                           VBA.MsgBox msg, vbCritical, Caption
        End Select
    Else
        bOkNow = True
    End If
    
    If bFailed And bOkNow Then ShellFile = sNewFileName

End Function

Public Function RegSaveKey(ByVal hive As RegistryHiveConstants, ByVal section As String, Optional ByVal key As String, Optional ByVal value As String) As Boolean
   ' Section   Required. String expression containing the name of the section WHERE the key setting is being saved.
   ' Key       Required. String expression containing the name of the key setting being saved.
   ' Setting   Required. Expression containing the Value that key is being set to.
   Dim nRet As Long
   Dim hKey As Long
   Dim nResult As Long
   
   ' Open (or create AND open) key
   nRet = RegCreateKeyEx(hive, section, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_READ, ByVal 0&, hKey, nResult)
   If nRet = ERROR_SUCCESS Then
      ' Set appropriate Value for default query
      If key = "" Then key = vbNullString
      ' Write new Value to registry
      nRet = RegSetValueEx(hKey, key, 0&, REG_SZ, ByVal value, Len(value))
      Call RegCloseKey(hKey)
   End If
   RegSaveKey = (nRet = ERROR_SUCCESS)
End Function


Public Function FileTitle(fullpath As String) As String
    FileTitle = Parse(fullpath, Parse(fullpath, , "\"), "\")
End Function

Public Function CleanFileName(FileName As String) As String
'strip out reserved characters
    FileName = Replace(FileName, "/", "")
    FileName = Replace(FileName, "\", "")
    FileName = Replace(FileName, "*", "")
    FileName = Replace(FileName, "?", "")
    FileName = Replace(FileName, ":", "")
    FileName = Replace(FileName, "<", "")
    FileName = Replace(FileName, ">", "")
    FileName = Replace(FileName, "|", "")
    CleanFileName = FileName
End Function
Public Function FileName(fullpath As String) As String
    FileName = StripExtension(FileTitle(fullpath))
End Function
Public Function FilePath(fullpath As String) As String
    FilePath = Left(fullpath, Len(fullpath) - Len(FileTitle(fullpath)))
End Function
Public Function FileExt(fullpath As String) As String
    If InStr(fullpath, ".") Then
        FileExt = LCase(Parse(fullpath, Parse(fullpath, , "."), "."))
    Else
        FileExt = ""
    End If
End Function
Public Function ForceExt(FileName As String, Extension As String) As String
    If Extension = "" Then
        ForceExt = FileName
    Else
        ForceExt = StripExtension(FileName) & "." & Extension
    End If
End Function

Public Function StripExtension(FileName As String) As String
    Dim i As Integer
    Dim s As String
    Dim b As Boolean
    For i = Len(FileName) To 1 Step -1
        If Mid(FileName, i, 1) = "." Then b = True
        If b Then s = Mid(FileName, i, 1) & s
    Next
    If b Then
        StripExtension = Mid(s, 1, Len(s) - 1)
    Else
        StripExtension = FileName
    End If
End Function



Public Function FileExists(FileName As String) As Boolean
On Error Resume Next
    FileExists = Dir(FileName) <> ""
End Function
Public Function FileDrive(fullpath As String) As String
    fullpath = Trim(fullpath)
    If Left(fullpath, 2) = "\\" Then
        'UNC names
        FileDrive = "\\" & Parse(Mid(fullpath, 3), 1, "\")
    Else
        FileDrive = Left(fullpath, 2)
    End If
End Function
Public Function DNull(s As String) As String
    DNull = Mid(s, 1, Max(0, InStr(1, s, vbNullChar) - 1))
End Function
   

Public Function PathExists(Path As String) As Boolean
    PathExists = Dir(Path, vbDirectory) <> "" Or Dir(Path) <> ""
End Function

Public Sub CreatePath(PathDescription As String, Path As String)
On Error Resume Next
    Dim i As Long
    Dim s As String
    s = Parse(Path, 1, "\")
    For i = 2 To Parse(Path, , "\")
        s = s & "\" & Parse(Path, i, "\")
        Call MkDir(s)
    Next
End Sub



Public Function GetLocalizedPath(sPath As String) As String
    Dim s As String

    If Not IsPathNetPath(sPath) Then
        GetLocalizedPath = sPath
    Else
        If PathIsUNC(sPath) = 1 Then
            s = sPath
        Else
            s = Mapped2UNC(sPath)
        End If

        If UCase(MachineName()) = UCase(Parse(s, 3, "\")) Then
            GetLocalizedPath = UNC2Local(s)
        Else
            GetLocalizedPath = sPath
        End If


    End If

End Function
Public Function PathAppend(ParamArray Path()) As String
    Dim fullpath As String
    Dim s As String
    Dim i As Long
    s = Trim(Path(i))
    If Right(s, 1) = "\" Then s = Mid(s, 1, Len(s) - 1)
    fullpath = s
    For i = LBound(Path) + 1 To UBound(Path)
        s = Trim(Path(i))
        If Right(s, 1) = "\" Then s = Mid(s, 1, Len(s) - 1)
        If Left(s, 1) = "\" Then s = Mid(s, 2)
        fullpath = fullpath & "\" & s
    Next
    PathAppend = fullpath
End Function


Public Function UNC2Local(sUNCPath As String) As String
    Dim sTemp As String
    Dim sServer As String
    Dim sShare As String
    Dim baServer() As Byte
    Dim baShare() As Byte
    Dim result As Long
    Dim Buf As Long
    Dim TempStr As MungeInt
    Dim TempPtr As MungeLong
    Dim STRArray(0 To 255) As Byte
    Dim sBasePath As String

    If sUNCPath = "" Then Exit Function
    
    sTemp = Mid(sUNCPath, 3)
    sServer = Parse(sUNCPath, 3, "\")
    sTemp = Mid(sTemp, InStr(1, sTemp, "\") + 1)
    If InStr(1, sTemp, "\") > 0 Then
        sShare = Left(sTemp, InStr(1, sTemp, "\") - 1)
        sTemp = Mid(sTemp, InStr(1, sTemp, "\") + 1)
    Else
        sShare = sTemp
        sTemp = ""
    End If



    baServer = "\\" & sServer & Chr(0)
    baShare = UCase(sShare) & Chr(0)

    result = NetShareGetInfo(baServer(0), baShare(0), 2, Buf)

    If result = 0 Then
        result = PtrToInt(TempStr.XLo, Buf + 24, 2)
        result = PtrToInt(TempStr.XHi, Buf + 26, 2)
        LSet TempPtr = TempStr
        result = PtrToStr(STRArray(0), TempPtr.X)
        sBasePath = Left(STRArray, StrLen(TempPtr.X))

        result = NetAPIBufferFree(Buf)

        UNC2Local = sBasePath & sTemp
    End If
End Function






Public Function Mapped2UNC(sMappedPath As String) As String

   Dim sLocalRoot As String
   Dim sRemoteName As String
   Dim sRemotePath As String
   Dim cbRemoteName As Long

   sRemoteName = Space$(260)
   cbRemoteName = Len(sRemoteName)

   sLocalRoot = Left(sMappedPath, 2)

   sRemotePath = StripRootFromPath(sMappedPath)

   If IsPathNetPath(sLocalRoot) Then
      If WNetGetConnection(sLocalRoot, sRemoteName, cbRemoteName) = ERROR_SUCCESS Then
         sRemoteName = PathAppend(QualifyPath(TrimNull(sRemoteName)), sRemotePath)
         If IsUNCPathValid(sRemoteName) Then
            Mapped2UNC = sRemoteName
         End If

      End If
   End If
End Function


Public Function QualifyPath(sPath As String) As String

  'add trailing slash if required
   If Right$(sPath, 1) <> "\" Then
      QualifyPath = sPath & "\"
   Else
      QualifyPath = sPath
   End If

End Function


Public Function IsUNCPathValid(ByVal sPath As String) As Boolean

  'Determines if string is a valid UNC
   IsUNCPathValid = PathIsUNC(sPath) = 1

End Function

Public Function IsPathNetPath(ByVal sPath As String) As Boolean

  'Determines whether a path represents network resource.
   IsPathNetPath = PathIsNetworkPath(sPath) = 1

End Function





Private Function StripPathToRoot(ByVal sPath As String) As String

  'Removes all of the path except for
  'the root information (ie drive. Also
  'removes any trailing slash.
   Dim POs As Integer

   Call PathStripToRoot(sPath)

   POs = InStr(sPath, Chr$(0))
   If POs Then
      StripPathToRoot = Left$(sPath, POs - 2)
   Else
      StripPathToRoot = sPath
   End If

End Function
Private Function TrimNull(startstr As String) As String
   TrimNull = Left$(startstr, lstrlenW(StrPtr(startstr)))
End Function
Private Function StripRootFromPath(ByVal sPath As String) As String
  'Parses a path, ignoring the drive
  'letter or UNC server/share path parts
   StripRootFromPath = TrimNull(GetStrFromPtrA(PathSkipRoot(sPath)))
End Function
Private Function GetStrFromPtrA(ByVal lpszA As Long) As String

  'Given a pointer to a string, return the string
   GetStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
   Call lstrcpyA(ByVal GetStrFromPtrA, ByVal lpszA)

End Function
Public Function NetworkLoginName() As String
    Dim lpBuff As String * 25
    
    Call GetUserName(lpBuff, 25)
    NetworkLoginName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
     
End Function

Public Function CreateGUID() As String
    Dim tG As Guid
    Dim b() As Byte
    Dim lSize As Long
    Dim lR As Long
    Call CoCreateGuid(tG)
    lSize = 40
    ReDim b(0 To (lSize * 2) - 1) As Byte
    lR = StringFromGUID2(tG, VarPtr(b(0)), lSize)
    CreateGUID = Left$(b, lR - 1)
End Function

Public Sub IniPutForm(Form As Form, Optional IniFile As String, Optional InstanceKey)
    Dim c         As Control
    Dim FormName  As String
    Dim CtrlName  As String
    Dim bVisible  As Boolean
    Dim bCaptions As Boolean
    If IniFile = "" Then IniFile = AppIni
    
    FormName = Form.Name
    If Not IsMissing(InstanceKey) Then FormName = FormName & "(" & InstanceKey & ")"
    
    Call IniPut(IniFile, FormName, "State", Form.WindowState)
    If Form.WindowState = vbNormal Then
        Call IniPut(IniFile, FormName, "Top", Form.Top)
        Call IniPut(IniFile, FormName, "Left", Form.Left)
        Call IniPut(IniFile, FormName, "Height", Form.Height)
        Call IniPut(IniFile, FormName, "Width", Form.Width)
    End If
    If Form.WindowState <> vbMinimized Then
        For Each c In Form.Controls
            On Error Resume Next
            CtrlName = c.Name
            CtrlName = CtrlName & "(" & c.Index & ")"
            On Error GoTo 0
            Select Case TypeName(c)
                Case "Toolbar"
                    bVisible = c.Visible
                    On Error Resume Next
                    bCaptions = CBool(Parse(c.Tag, 2))
                    On Error GoTo 0
                    Call IniPut(IniFile, FormName, CtrlName & ".Visible", bVisible)
                    Call IniPut(IniFile, FormName, CtrlName & ".Captions", bCaptions)
                Case "StatusBar", "vbalARListBar"
                    bVisible = c.Visible
                    Call IniPut(IniFile, FormName, CtrlName & ".Visible", bVisible)
                Case "Slider", "Panel"
                    bVisible = c.Visible
                    Call IniPut(IniFile, FormName, CtrlName & ".Visible", bVisible)
                    Call IniPut(IniFile, FormName, CtrlName & ".Top", c.Top)
                    Call IniPut(IniFile, FormName, CtrlName & ".Left", c.Left)
            End Select
        Next
    End If
End Sub

Public Property Get AppWorkingFolder() As String
On Error Resume Next
    Dim s As String
    s = PathAppend(GetFolderPath(CSIDL_LOCAL_APPDATA), "Homefront\" & App.EXEName)
    Call CreatePath("", s)
    AppWorkingFolder = s
End Property
Public Property Get AppIni() As String
'    AppIni = PathAppend(AppWorkingFolder, "settings.ini")
    AppIni = PathAppend(HFApp.SystemFolder, "User Settings", NetworkLoginName & ".ini")
End Property
Public Property Get AppSQLCache() As String
    AppSQLCache = PathAppend(AppWorkingFolder, "query.sql")
End Property
Public Property Get AppErrorLog() As String
    AppErrorLog = PathAppend(AppWorkingFolder, "errors.txt")
End Property

Public Sub SetToolbarIcons(Toolbar, il)
On Error Resume Next
    Dim i As Integer
    Set Toolbar.ImageList = il
    For i = 1 To Toolbar.Buttons.Count
        Toolbar.Buttons(i).Image = Toolbar.Buttons(i).key
    Next
End Sub

Public Sub ShowToolbarCaptions(Toolbar, b)
    Dim i As Integer
    Toolbar.Visible = False
    If b Then
        For i = 1 To Toolbar.Buttons.Count
            If Toolbar.Buttons(i).Caption = "" Then
                Toolbar.Buttons(i).Caption = Toolbar.Buttons(i).Tag
                Toolbar.Buttons(i).ToolTipText = ""
            End If
        Next
    Else
        For i = 1 To Toolbar.Buttons.Count
            If Toolbar.Buttons(i).Caption <> "" Then
                Toolbar.Buttons(i).ToolTipText = Toolbar.Buttons(i).Caption
                Toolbar.Buttons(i).Tag = Toolbar.Buttons(i).Caption
                Toolbar.Buttons(i).Caption = ""
            End If
        Next
    End If
    Toolbar.Visible = True
    Toolbar.Tag = Update(Toolbar.Tag, 2, b)
    Toolbar.Refresh
On Error Resume Next
    Toolbar.Parent.Refresh
End Sub


Public Function trunc(Number As Double, precision As Integer) As String
    trunc = format(Number, "0." & String(precision, "0"))
End Function

Public Function InIde() As Boolean
On Error Resume Next
    Debug.Print (1 / 0)
    InIde = (Err.Number <> 0)
    On Error GoTo 0
    Exit Function
End Function

Public Function SimpleEncrypt(ByVal Secret As String, key As String) As String
    Dim i As Long
    For i = 1 To Len(Secret)
        Mid$(Secret, i, 1) = Chr$(Asc(Mid$(Secret, i, 1)) Xor Asc(Mid$(key, (i Mod Len(key)) - Len(key) * ((i Mod Len(key)) = 0), 1)))
    Next
    SimpleEncrypt = Secret
End Function

Public Function SetListIndex(Combobox, Optional ItemData As Long, Optional Text As String) As Boolean
    Dim i As Long
    
    If Text <> "" Then
        For i = 0 To Combobox.ListCount - 1
            If Combobox.List(i) = Text Then
                Combobox.ListIndex = i
                SetListIndex = True
                Exit Function
            End If
        Next
    Else
        If ItemData = 0 Then
            On Error Resume Next
            Combobox.Text = ""
            SetListIndex = True
            Exit Function
        Else
            For i = 0 To Combobox.ListCount - 1
                If Combobox.ItemData(i) = ItemData Then
                    Combobox.ListIndex = i
                    SetListIndex = True
                    Exit Function
                End If
            Next
        End If
    End If
    
    SetListIndex = False
    
End Function

Public Function Update(ByVal InputString As String, ByVal Col As Long, ByVal value As String, Optional Delimeter As String = ",") As String

    'updates items in a delimited string
    
    Dim POs        As Integer
    Dim workString As String
    Dim idx        As Integer

    For idx = 1 To Col - 1
        POs = InStr(POs + 1, InputString, Delimeter)
        If POs = 0 Then
            ' column not found, add one
            InputString = InputString & Delimeter
            POs = Len(InputString)
        End If
    Next idx

    workString = Left(InputString, POs)
    POs = InStr(POs + 1, InputString, Delimeter)

    If POs > 0 Then 'not last column
        workString = workString & value & Mid$(InputString, POs)
    Else
        workString = workString & value
    End If

    Update = workString

End Function

Public Sub IniGetForm(Form, Optional IniFile As String, Optional InstanceKey)
'On Error GoTo exitsub
    Dim c         As Control
    Dim WinState  As Integer
    Dim FormName  As String
    Dim CtrlName  As String
    Dim bVisible  As Boolean
    Dim bCaptions As Boolean
    Dim value     As Double
    
    If IniFile = "" Then IniFile = AppIni
        
    FormName = Form.Name
    If Not IsMissing(InstanceKey) Then FormName = FormName & "(" & InstanceKey & ")"
    
    WinState = IniGet(IniFile, FormName, "State", Form.WindowState)
    
    
    value = Val("" & IniGet(IniFile, FormName, "Top", Form.Top))
    value = IIf(value < 0, 0, value)
    value = IIf(value > Screen.Height - 600, Form.Top, value)
    Form.Top = value
    
    value = Val("" & IniGet(IniFile, FormName, "Left", Form.Left))
    value = IIf(value + Form.Width < 0, 0, value)
    value = IIf(value > Screen.Width, Form.Left, value)
    Form.Left = value
    
    On Error Resume Next
    If Form.BorderStyle = 2 Or Form.BorderStyle = 5 Then
        Form.Height = IniGet(IniFile, FormName, "Height", Form.Height)
        Form.Width = IniGet(IniFile, FormName, "Width", Form.Width)
    End If
    On Error GoTo 0
    
    
    If WinState = vbMinimized Then WinState = vbNormal
    Form.WindowState = WinState
    
    For Each c In Form.Controls
        On Error Resume Next
        CtrlName = c.Name
        CtrlName = CtrlName & "(" & c.Index & ")"
        On Error GoTo 0
        
        Select Case TypeName(c)
            
            Case "Toolbar"
                bVisible = IniGet(IniFile, FormName, CtrlName & ".Visible", True)
                bCaptions = IniGet(IniFile, FormName, CtrlName & ".Captions", True)
                c.Tag = bVisible & "," & bCaptions
                c.Visible = bVisible
                Call ShowToolbarCaptions(c, bCaptions)
            
            Case "StatusBar", "vbalARListBar"
                bVisible = IniGet(IniFile, FormName, CtrlName & ".Visible", True)
                c.Tag = "" & bVisible
                c.Visible = bVisible
                
            Case "Slider", "Panel"
                bVisible = IniGet(IniFile, FormName, CtrlName & ".Visible", True)
                c.Tag = "" & bVisible
                c.Visible = bVisible
                c.Top = IniGet(IniFile, FormName, CtrlName & ".Top", c.Top)
                c.Left = IniGet(IniFile, FormName, CtrlName & ".Left", c.Left)
                
        End Select
    Next
    
    Dim X As LabelEventTrap
    On Error GoTo ExitSub
    Set Form.EventTraps = New Collection
    For Each c In Form.Controls
        If TypeName(c) = "Label" Then
            Set X = New LabelEventTrap
            On Error Resume Next
            Set X.ctrl = c
            If Err.Number = 0 Then Form.EventTraps.Add X
        End If
    Next

ExitSub: Exit Sub
End Sub






Public Function csv(ParamArray values()) As String
    Dim i As Integer
    Dim s As String
    For i = LBound(values) To UBound(values)
        s = s & "," & values(i)
    Next
    csv = Mid(s, 2)
End Function

Public Function GridWidth(g) As Long
'returns the width of the non-hidden cols in a flexgrid
    Dim i As Integer
    Dim w As Long
    For i = 0 To g.Cols - 1
        If Not g.ColHidden(i) Then
            w = w + g.ColWidth(i)
        End If
    Next
    GridWidth = w
End Function
Public Function GridHeight(g) As Long
'returns the client height of a flexgrid
    Dim i As Integer
    Dim H As Long
    For i = 0 To g.Rows - 1
        H = H + g.RowHeight(i)
    Next
    H = H + (g.Rows - 1) * g.GridLineWidth * 15
    GridHeight = H
End Function

Public Sub GridGotFocus(g)
On Error Resume Next
    If g.Col < 0 Then g.Col = 0
    If g.ColHidden(g.Col) Then g.Col = GridNextVisibleColumn(g, g.Col)
    If g.Row = -1 Then g.Row = 1
End Sub

Public Function GridLastVisibleRow(g) As Long
    Dim i As Long
    For i = g.Rows - 1 To 0 Step -1
        If Not g.RowHidden(i) Then
            GridLastVisibleRow = i
            Exit Function
        End If
    Next
End Function

Public Function GridNextVisibleColumn(g, c) As Long
    Dim i As Long
    For i = c To g.Cols - 1
        If Not g.ColHidden(i) Then
            GridNextVisibleColumn = i
            Exit Function
        End If
    Next
End Function
Public Function GridLastVisibleColumn(g) As Long
    Dim i As Long
    For i = g.Cols - 1 To 0 Step -1
        If Not g.ColHidden(i) Then
            GridLastVisibleColumn = i
            Exit Function
        End If
    Next
End Function
Public Function GridVisibleRowIndex(g, NthRow As Long, Optional StartAtRow As Long = -1) As Long
    'gets the row index of the Nth visible row
    Dim i As Long
    Dim Count As Long
    If StartAtRow = -1 Then StartAtRow = g.FixedRows
    For i = StartAtRow To g.Rows - 1
        If Not g.RowHidden(i) Then Count = Count + 1
        If Count = NthRow Then GridVisibleRowIndex = i
    Next

End Function
Public Function GridVisibleRows(g) As Long
    Dim i As Long
    Dim c As Long
    For i = g.Rows - 1 To 0 Step -1
        If Not g.ColHidden(i) Then
            c = c + 1
        End If
    Next
    GridVisibleRows = c
End Function
Public Function GridVisibleColumns(g) As Long
    Dim i As Long
    Dim c As Long
    For i = g.Cols - 1 To 0 Step -1
        If Not g.ColHidden(i) Then
            c = c + 1
        End If
    Next
    GridVisibleColumns = c
End Function
Public Function GridNextVisibleRow(g, r) As Long
    Dim i As Long
    For i = r To g.Rows - 1
        If Not g.RowHidden(i) Then
            GridNextVisibleRow = i
            Exit Function
        End If
    Next
End Function

Public Sub GridAutoSizeRows(g, c As Long)
    Dim s As String
    Dim r As Long
    With g
        For r = .FixedRows To .Rows - 1
            s = .TextMatrix(r, c)
            .RowHeight(r) = 240 * Parse(s, , vbCrLf)
        Next
    End With
End Sub


Public Sub SetCtrlFocus(c As Control)
On Error Resume Next
    Call c.SetFocus
End Sub

Public Function IsIn(Expression, ParamArray items()) As Boolean
    Dim i As Integer
    For i = 0 To UBound(items)
        If "" & Expression = "" & items(i) Then
            IsIn = True
            Exit Function
        End If
    Next
    IsIn = False
End Function

Public Sub LoadComboBox(Combobox, Connection, sql As String)

    Dim rs  'As ADODB.Recordset
    Dim Tag As String
    
    'SELECT ItemDesc  text to be displayed in list
    '      ,KeyValue  alpha key value
    '      ,IDValue   numeric key value
    '  FROM Table
    
    'ID are stored in combobox.itemdata(x)
    'Keys are stored chr(1) separated in the controls tag property
    '     use SetComboBoxListIndex & GetComboBoxListKey, GetComboBoxListID, GetComboBoxListIndex functions
    
    Set rs = Connection.Execute(sql)
    Combobox.Clear
    Tag = ""
    While Not rs.EOF
        Combobox.AddItem "" & rs(0)
        Combobox.ItemData(Combobox.NewIndex) = Val("" & rs(2))
        Tag = Tag & Chr(1) & "" & rs(1)
        rs.MoveNext
    Wend
    Combobox.Tag = Mid(Tag, 2)
End Sub

Public Sub SelectAll(c As Control)
On Error Resume Next
    ReleaseCapture
    Select Case TypeName(c)
        Case "VSFlexGrid"
            Call c.Select(c.FixedRows, c.FixedCols, c.Rows - 1, c.Cols - 1)
        Case Else
            c.SelStart = 0
            c.SelLength = Len(c)
    End Select
End Sub

Public Function GetComboBoxListKey(Combobox) As String
If Combobox.ListIndex > -1 Then GetComboBoxListKey = Parse(Combobox.Tag, Combobox.ListIndex + 1, Chr(1))
End Function
Public Function GetComboBoxListID(Combobox) As Long

    If Combobox.ListIndex > -1 Then
        GetComboBoxListID = Combobox.ItemData(Combobox.ListIndex)
    Else
        GetComboBoxListID = -1
    End If
End Function
Public Function GetComboBoxListIndex(Combobox) As Long
    GetComboBoxListIndex = Combobox.ListIndex
End Function

Public Function IsFormLoaded(FormName As String) As Boolean
    Dim i As Long
    For i = 0 To Forms.Count - 1
        If Forms(i).Name = FormName Then
            IsFormLoaded = True
            Exit Function
        End If
    Next
End Function


Public Function ImageIndex(ImageList, key As String) As Long
On Error Resume Next
    ImageIndex = ImageList.ListImages(key).Index - 1
End Function


Public Function ShowForm(FormName As String, TagValue) As Boolean
    Dim f As Form
    For Each f In Forms
        If f.Name = FormName And f.Tag = TagValue Then
            ShowForm = True
            If f.WindowState = vbMinimized Then f.WindowState = vbNormal
            f.SetFocus
            Exit Function
        End If
    Next
End Function

Public Function GridCellSelected(g, Row As Long, Col As Long) As Boolean
    GridCellSelected = IsBetween(Row, Min(g.Row, g.RowSel), Max(g.Row, g.RowSel)) And _
                       IsBetween(Col, Min(g.Col, g.ColSel), Max(g.Col, g.ColSel))
End Function


Public Function MouseX(Optional ByVal hwnd As Long) As Long
    Dim lpPoint As POINTAPI
    GetCursorPos lpPoint
    If hwnd Then ScreenToClient hwnd, lpPoint
    MouseX = lpPoint.X
End Function

Public Function MouseY(Optional ByVal hwnd As Long) As Long
    Dim lpPoint As POINTAPI
    GetCursorPos lpPoint
    If hwnd Then ScreenToClient hwnd, lpPoint
    MouseY = lpPoint.Y
End Function

Public Sub SetComboBoxListIndex(Combobox, Optional Text As String, Optional key As String, Optional ID As Long = -1, Optional Index As Long = -1)
On Error Resume Next
    Dim s As String
    Dim i As Long
    
    If Index <> -1 Then
        Combobox.ListIndex = Index
    End If
    
    If Text <> "" Then
        For i = 0 To Combobox.ListCount - 1
            If UCase(Combobox.List(i)) = UCase(Text) Then
                Combobox.ListIndex = i
                Exit Sub
            End If
        Next
    End If
    
    
    If key <> "" Then
        s = Combobox.Tag
        For i = 1 To Parse(s, , Chr(1))
            If UCase(Parse(s, i, Chr(1))) = UCase(key) Then
                Combobox.ListIndex = i - 1
                Exit Sub
            End If
        Next
    End If
    
    If ID <> -1 Then
        For i = 0 To Combobox.ListCount - 1
            If Combobox.ItemData(i) = ID Then
                Combobox.ListIndex = i
                Exit Sub
            End If
        Next
    End If
    
    Combobox.ListIndex = -1
    
End Sub

Public Sub EnableCtrl(c As Control, Enabled As Boolean)
    
    c.Enabled = Enabled
    Select Case TypeName(c)
        Case "Image"
            c.Visible = Enabled
        Case "Label"
            c.ForeColor = IIf(Enabled, vbHighlight, vbGrayText)
        Case "TextBox"
            c.BackColor = IIf(Enabled, vbWindowBackground, vbButtonFace)
    End Select
    
End Sub

Public Function DeQuote(s As String, Optional SingleQuotes As Boolean = False) As String
    
    s = Trim(s)
    If SingleQuotes Then
        If Left(s, 1) = vbSingleQuote And Right(s, 1) = vbSingleQuote Then
            s = Mid(s, 2, Len(s) - 2)
        End If
    Else
        If Left(s, 1) = vbQuote And Right(s, 1) = vbQuote Then
            s = Mid(s, 2, Len(s) - 2)
        End If
    End If
    DeQuote = s


End Function


Public Function Quote(s As String, Optional RemoveFormatting As Boolean = True, Optional Length As Long, Optional RemoveSpecial As Boolean = True) As String
    
    If RemoveFormatting Then
        s = Replace(s, vbTab, " ")
        s = Replace(s, vbLf, " ")
        s = Replace(s, vbCr, " ")
    End If

    'replace special characters
    If RemoveSpecial Then s = RemoveSpecialChar(s)

    'replace quote with 2 apostrophes
    s = Replace(s, vbQuote, "''")

    If Length <> 0 Then
        s = Left(s, Length)
    End If
    
    Quote = vbQuote & s & vbQuote
    
End Function

Private Function RemoveSpecialChar(Text As String) As String
    Dim i As Long
    Dim j As Long
    Dim Char As String

    i = 1
    For j = 1 To Len(Text)
        Char = Mid$(Text, j, 1)
        If (AscW(Char) And &HFFFF&) <= &H7F& Then
            Mid$(Text, i, 1) = Char
            i = i + 1
        End If
    Next
    
    RemoveSpecialChar = Left$(Text, i - 1)
End Function

Private Function PointerToStringA(ByVal lpStringA As Long) As String
   Dim Buffer() As Byte
   Dim nLen As Long
   
   If lpStringA Then
      nLen = lstrlenA(ByVal lpStringA)
      If nLen Then
         ReDim Buffer(0 To (nLen - 1)) As Byte
         CopyMemory Buffer(0), ByVal lpStringA, nLen
         PointerToStringA = StrConv(Buffer, vbUnicode)
      End If
   End If
End Function
Public Function DefaultPrinterName() As String
   ' HOWTO: Retrieve and Set the Default Printer in Windows
   ' http://support.microsoft.com/support/kb/articles/q246/7/72.asp
   ' HOWTO: Get and Set the Default Printer in Windows
   ' http://support.microsoft.com/support/kb/articles/q135/3/87.asp
   Dim os As OSVERSIONINFO
   Dim Buffer() As Byte
   Dim BufSize As Long
   Dim pPrinterName As Long
   Dim Returned As Long
   Dim result As String
   
   ' Get OS version info, so we know which way to fork.
   os.dwOSVersionInfoSize = Len(os)
   Call GetVersionEx(os)
   
   If os.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then  '95/98/ME
      ' Determine how big the buffer needs to be
      Call EnumPrinters(PRINTER_ENUM_DEFAULT, vbNullString, 2, ByVal 0&, 0, BufSize, Returned)
      If BufSize > 0 Then
         ' Size buffer accordingly
         ReDim Buffer(0 To BufSize - 1) As Byte
         ' Call again to retrieve needed info
         Call EnumPrinters(PRINTER_ENUM_DEFAULT, vbNullString, 2, Buffer(0), BufSize, BufSize, Returned)
         ' A pointer to the default printer name is
         ' returned at the 5th byte in the buffer.
         Call CopyMemory(pPrinterName, Buffer(4), 4)
         result = PointerToStringA(pPrinterName)
      End If
       
   ElseIf os.dwPlatformId = VER_PLATFORM_WIN32_NT Then
      ' Create satisfactory buffer.
      BufSize = 1024
      result = Space$(BufSize)
      
      ' Use either GetDefaultPrinter (2k+) or WIN.INI (NT4).
      If os.dwMajorVersion >= 5 Then
         If GetDefaultPrinter(result, BufSize) Then
            ' Truncate at first NULL
            result = Left$(result, InStr(result, vbNullChar) - 1)
         End If
      Else 'NT4 or less
         ' The old WIN.INI [Windows] section is mapped to
         ' HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows
         ' and we can just use GetProfileString to extract! :-)
         ' Returns: "printer name,driver name,port"
         If GetProfileString("Windows", ByVal "device", "", result, BufSize) Then
            ' Truncate buffer at end of name.
            result = Left$(result, InStr(result, ",") - 1)
         End If
      End If
   End If
      
   ' Return default printer name.
   DefaultPrinterName = result
End Function


Public Sub GetPrinterInfo(DeviceName As String, Port As String, Status As String, Model As String, Location As String, Comment As String, IsDefault As Boolean)
    Dim pi2 As PRINTER_INFO_2
    Dim hPrn As Long
    Dim Buffer() As Byte
    Dim BytesNeeded As Long
    Dim BytesUsed As Long
    Dim slash As Long
    
    ' Get handle to printer.
    Call OpenPrinter(DeviceName, hPrn, ByVal 0&)
    If hPrn Then
        Call GetPrinter(hPrn, 2, ByVal 0&, 0, BytesNeeded)
        ReDim Buffer(0 To BytesNeeded - 1) As Byte
        If GetPrinter(hPrn, 2, Buffer(0), BytesNeeded, BytesUsed) Then
            
            Call CopyMemory(pi2, Buffer(0), Len(pi2))
            
            Model = PointerToStringA(pi2.pDriverName)
            Comment = PointerToStringA(pi2.pComment)
            Location = Trim(PointerToStringA(pi2.pLocation))
            If Location = "" Then Location = Trim(PointerToStringA(pi2.pPortName))
            Port = Trim(PointerToStringA(pi2.pPortName))
            IsDefault = UCase(DeviceName) = UCase(DefaultPrinterName)
            Select Case pi2.Status
                Case PRINTER_STATUS_READY:                 Status = "Ready"
                Case PRINTER_STATUS_PAUSED:                Status = "Paused"
                Case PRINTER_STATUS_ERROR:                 Status = "Error"
                Case PRINTER_STATUS_PENDING_DELETION:      Status = "Deleting..."
                Case PRINTER_STATUS_PAPER_JAM:             Status = "Paper Jam"
                Case PRINTER_STATUS_PAPER_OUT:             Status = "Paper Out"
                Case PRINTER_STATUS_MANUAL_FEED:           Status = "Manual Feed Required"
                Case PRINTER_STATUS_PAPER_PROBLEM:         Status = "Paper Problem"
                Case PRINTER_STATUS_OFFLINE:               Status = "Offline"
                Case PRINTER_STATUS_IO_ACTIVE:             Status = "Downloading Job"
                Case PRINTER_STATUS_BUSY:                  Status = "Busy"
                Case PRINTER_STATUS_PRINTING:              Status = "Printing"
                Case PRINTER_STATUS_OUTPUT_BIN_FULL:       Status = "Output Bill Full"
                Case PRINTER_STATUS_NOT_AVAILABLE:         Status = "Not Available"
                Case PRINTER_STATUS_WAITING:               Status = "Waiting"
                Case PRINTER_STATUS_PROCESSING:            Status = "Processing Job"
                Case PRINTER_STATUS_INITIALIZING:          Status = "Initializing"
                Case PRINTER_STATUS_WARMING_UP:            Status = "Warming Up"
                Case PRINTER_STATUS_TONER_LOW:             Status = "Toner Low"
                Case PRINTER_STATUS_NO_TONER:              Status = "Toner Out"
                Case PRINTER_STATUS_PAGE_PUNT:             Status = "Page too Complex"
                Case PRINTER_STATUS_USER_INTERVENTION:     Status = "User Intervention Required"
                Case PRINTER_STATUS_OUT_OF_MEMORY:         Status = "Out of Memory"
                Case PRINTER_STATUS_DOOR_OPEN:             Status = "Door Open"
                Case PRINTER_STATUS_SERVER_UNKNOWN:        Status = "Unable to connect"
                Case PRINTER_STATUS_POWER_SAVE:            Status = "Power Save Mode"
                Case Else:                                 Status = Hex$(pi2.Status)
            End Select
        End If
        Call ClosePrinter(hPrn)
    End If
    
End Sub

Public Function FormatPhone(Number As String) As String
    Dim countrycode As String 'AU,CA,US
    Dim isMobile As Boolean
    Dim isInternational As Boolean
    
    countrycode = Trim(HFApp.Options(Country))
    Number = Trim(Number)
    
    Select Case countrycode
    Case "AU"
        'mobiles are 4 and 5 numbers
        'mobile international    "+61499999999" --> +61 499 999 999
        'mobile international    "61499999999"  --> +61 499 999 999
        'mobile                  "0499999999"   --> 0499 999 999
        'mobile                  "499999999"    --> 0499 999 999
        'land line international "+61999999999" --> +61 9 9999 9999
        'land line international "61999999999"  --> +61 9 9999 9999
        'land line               "0999999999"   --> 09 9999 9999
        'land line               "999999999"    --> 09 9999 9999
        If Len(Number) = 11 Then Number = "+" & Number
        If Len(Number) = 9 Then Number = "0" & Number
        isInternational = Left(Number, 1) = "+"
        If isInternational Then
            isMobile = IsIn(Mid(Number, 4, 1), "4", "5")
        Else
            isMobile = IsIn(Left(Number, 2), "04", "05")
        End If
        Select Case True
            Case Not IsNumeric(Number):          FormatPhone = Number
            Case isMobile And isInternational:   FormatPhone = format(Number, "!@@@ @@@ @@@ @@@")
            Case isMobile:                       FormatPhone = format(Number, "!@@@@ @@@ @@@")
            Case isInternational:                FormatPhone = format(Number, "!@@@ @ @@@@ @@@@")
            Case Else:                           FormatPhone = format(Number, "!@@ @@@@ @@@@")
        End Select
    
    Case Else
        Select Case True
            Case Not IsNumeric(Number):   FormatPhone = Number
            Case Len(Number) = 10:        FormatPhone = format(Number, "!(@@@) @@@-@@@@")
            Case Len(Number) = 11:        FormatPhone = format(Number, "!@ (@@@) @@@-@@@@")
            Case Len(Number) = 7:         FormatPhone = format(Number, "!@@@-@@@@")
            Case Else:                    FormatPhone = Number
        End Select
    End Select

End Function

Public Function SpaceCase(ByVal s As String) As String
On Error Resume Next
    Dim newS  As String
    Dim cPrev As String
    Dim cThis As String
    Dim cNext As String
    Dim i As Long
    
    s = Replace(s, "_", " ")
    s = Replace(s, " ", "")
    For i = 1 To Len(s)
        cPrev = Mid(s, i - 1, 1)
        cThis = Mid(s, i + 0, 1)
        cNext = Mid(s, i + 1, 1)
        If i = Len(s) Then cNext = "Z"
        If cThis = UCase(cThis) And (cPrev = LCase(cPrev) Or cNext = LCase(cNext)) Then
            newS = newS & " " & cThis
        Else
            newS = newS & cThis
        End If
    Next
    SpaceCase = Trim(newS)
    
End Function

Public Function SelectedOption(OptionButtons) As Long
    Dim i As Long
    For i = OptionButtons.LBound To OptionButtons.UBound
        If OptionButtons(i) Then
            SelectedOption = i
            Exit Function
        End If
    Next
End Function

Public Function Dir(Optional PathName, Optional Attributes As VbFileAttribute = vbNormal, Optional Sorted As Boolean = False) As String
    Static pList()   As String
    Static pIndex    As Long
    Static pSorted As Boolean
    
    Dim s As String
    
    If IsMissing(PathName) Then
        If pSorted Then
            pIndex = pIndex + 1
            If pIndex > 0 And pIndex <= UBound(pList) Then
                Dir = pList(pIndex)
            End If
        Else
            Dir = "" & VBA.Dir
        End If
    Else
        pSorted = Sorted
        pIndex = 0
        If pSorted Then
            s = VBA.Dir(PathName, Attributes)
            ReDim pList(0)
            While s <> ""
                pIndex = pIndex + 1
                ReDim Preserve pList(pIndex)
                pList(pIndex) = s
                s = VBA.Dir
            Wend
            If UBound(pList) > 0 Then
                Call Sort(pList)
                pIndex = 1
                Dir = pList(pIndex)
            End If
        Else
            Dir = VBA.Dir(PathName, Attributes)
        End If
    End If
        
End Function


Public Sub Sort(Strings() As String, Optional ByVal Ascending As Boolean = True, Optional ByVal CaseSensitive As Boolean = False, Optional Algorithm As SortAlgorithms = QuickSOrt)
    Select Case Algorithm
        Case QuickSOrt:         Call QuickSortStringsStart(Strings, Ascending, CaseSensitive)
        Case InsertSort:        Call InsertSortStringsStart(Strings, Ascending, CaseSensitive)
        Case MergeSort:         Call MergeSortStringsStart(Strings, Ascending, CaseSensitive)
        Case SelectionSort:     Call SelectionSortStrings(Strings, Ascending, CaseSensitive)
    End Select
End Sub



'*******************************************************************************
' SelectionSortStrings (SUB)
'
' PARAMETERS:
' (In/Out) - ListArray()    - String  - Array to sort
' (In)     - bAscending     - Boolean - True to sort ascending, false descending
' (In)     - bCaseSensitive - Boolean - True for a case sensitive sort, false
'                                       for an insensitive one
'
' DESCRIPTION:
' Simple Selection Sort routine for strings, fast for small arrays (say less
' than 60 values).
'*******************************************************************************
Private Sub SelectionSortStrings(ListArray() As String, _
                                Optional ByVal bAscending As Boolean = True, _
                                Optional ByVal bCaseSensitive As Boolean = False)
    
    Dim sSmallest       As String
    Dim lSmallest       As Long
    Dim lCount1         As Long
    Dim lCount2         As Long
    Dim lMin            As Long
    Dim lMax            As Long
    Dim lCompareType    As Long
    Dim lOrder          As Long
    
    lMin = LBound(ListArray)
    lMax = UBound(ListArray)
    
    If lMin = lMax Then
        Exit Sub
    End If
    
    ' Order Ascending or Descending?
    lOrder = IIf(bAscending, -1, 1)
    
    ' Case sensitive search or not?
    lCompareType = IIf(bCaseSensitive, vbBinaryCompare, vbTextCompare)
    
    ' Loop through array swapping the smallest\largest (determined by lOrder)
    ' item with the current item
    For lCount1 = lMin To lMax - 1
        sSmallest = ListArray(lCount1)
        lSmallest = lCount1
        
        ' Find the smallest\largest item in the array
        For lCount2 = lCount1 + 1 To lMax
            If StrComp(ListArray(lCount2), sSmallest, lCompareType) = lOrder Then
                sSmallest = ListArray(lCount2)
                lSmallest = lCount2
            End If
        Next
        
        ' Just swap them, even if we are swapping it with itself,
        ' as it is generally quicker to do this than test first
        ' each time if we are already the smallest with a
        ' test like: If lSmallest <> lCount1 Then
        ListArray(lSmallest) = ListArray(lCount1)
        ListArray(lCount1) = sSmallest
    Next
End Sub


'*******************************************************************************
' InsertSortStringsStart (SUB)
'
' PARAMETERS:
' (In/Out) - ListArray()    - String  - Array to sort
' (In)     - bAscending     - Boolean - True to sort ascending, false descending
' (In)     - bCaseSensitive - Boolean - True for a case sensitive sort, false
'                                       for an insensitive one
'
' DESCRIPTION:
' User friendly entry point for InsertSortStrings
'*******************************************************************************
Private Sub InsertSortStringsStart(ListArray() As String, _
                                  Optional ByVal bAscending As Boolean = True, _
                                  Optional ByVal bCaseSensitive As Boolean = False)

    Dim lMin            As Long
    Dim lMax            As Long
    Dim lOrder          As Long
    Dim lCompareType    As Long

    lMin = LBound(ListArray)
    lMax = UBound(ListArray)
    
    If lMin = lMax Then
        Exit Sub
    End If
    
    ' Order Ascending or Descending?
    lOrder = IIf(bAscending, 1, -1)
    
    ' Case sensitive search or not?
    lCompareType = IIf(bCaseSensitive, vbBinaryCompare, vbTextCompare)
    
    InsertSortStrings ListArray, lMin, lMax, lOrder, lCompareType
End Sub


'*******************************************************************************
' InsertSortStrings (SUB)
'
' PARAMETERS:
' (In/Out) - ListArray()  - String - Array to sort
' (In)     - lMin         - Long   - Start of sorting region within array
' (In)     - lMax         - Long   - End of sorting region within array
' (In)     - lOrder       - Long   - Ascending is -1, Descending is +1, used
'                                    for comparison in StrComp
' (In)     - lCompareType - Long   - Either vbBinaryCompare or vbTextCompare,
'                                    used in StrComp function
'
' DESCRIPTION:
' Simple Insert Sort routine for strings, fast for small arrays as there is no
' recursion (say less than 60 values)
'*******************************************************************************
Private Sub InsertSortStrings(ListArray() As String, _
                              ByVal lMin As Long, _
                              ByVal lMax As Long, _
                              ByVal lOrder As Long, _
                              ByVal lCompareType As Long)
    
    Dim sValue  As String
    Dim lCount1 As Long
    Dim lCount2 As Long
    
    ' Loop through array shifting elements down to their correct place
    For lCount1 = lMin + 1 To lMax
        sValue = ListArray(lCount1)
        
        ' Find the place to put it
        For lCount2 = lCount1 - 1 To lMin Step -1
            If StrComp(ListArray(lCount2), sValue, lCompareType) <> lOrder Then
                Exit For
            End If
            ListArray(lCount2 + 1) = ListArray(lCount2)
        Next lCount2
        
        ' Insert it
        ListArray(lCount2 + 1) = sValue
    Next
End Sub


'*******************************************************************************
' QuickSortStringsStart (SUB)
'
' PARAMETERS:
' (In/Out) - ListArray()    - String  - Array to sort
' (In)     - bAscending     - Boolean - True to sort ascending, false descending
' (In)     - bCaseSensitive - Boolean - True for a case sensitive sort, false
'                                       for an insensitive one
'
' DESCRIPTION:
' User friendly entry point for QuickSortStrings
'*******************************************************************************
Private Sub QuickSortStringsStart(ListArray() As String, _
                                 Optional ByVal bAscending As Boolean = True, _
                                 Optional ByVal bCaseSensitive As Boolean = False)

    Dim lMin            As Long
    Dim lMax            As Long
    Dim lOrder          As Long
    Dim lCompareType    As Long

    lMin = LBound(ListArray)
    lMax = UBound(ListArray)
    
    If lMin = lMax Then
        Exit Sub
    End If
    
    ' Order Ascending or Descending?
    lOrder = IIf(bAscending, 1, -1)
    
    ' Case sensitive search or not?
    lCompareType = IIf(bCaseSensitive, vbBinaryCompare, vbTextCompare)
    
    QuickSortStrings ListArray, lMin, lMax, lOrder, lCompareType
End Sub


'*******************************************************************************
' QuickSortStrings (SUB)
'
' PARAMETERS:
' (In/Out) - ListArray()  - String - Array to sort
' (In)     - lLowerPoint  - Long   - Start of sorting region within array
' (In)     - lUpperPoint  - Long   - End of sorting region within array
' (In)     - lOrder       - Long   - Ascending is -1, Descending is +1, used
'                                    for comparison in StrComp
' (In)     - lCompareType - Long   - Either vbBinaryCompare or vbTextCompare,
'                                    used in StrComp function
'
' DESCRIPTION:
' Quick for large arrays, delegates to Insert Sort for small arrays and when
' partition is small
'*******************************************************************************
Private Sub QuickSortStrings(ListArray() As String, _
                             ByVal lLowerPoint As Long, _
                             ByVal lUpperPoint As Long, _
                             ByVal lOrder As Long, _
                             ByVal lCompareType As Long)
    
    Const DELEGATE_POINT As Long = 60
    
    Dim lMidPoint As Long
    
    ' Delegate to an insert sort if it is a small array (this is what makes this
    ' routine so much quicker than a standard quick sort routine).  The delegation
    ' point could be tuned if necessary.
    If (lUpperPoint - lLowerPoint) <= DELEGATE_POINT Then
        InsertSortStrings ListArray, lLowerPoint, lUpperPoint, lOrder, lCompareType
        Exit Sub
    End If

    ' Do the quick sort
    Do While lLowerPoint < lUpperPoint
        ' Find a mid point (split the array into partitions)
        lMidPoint = QuickSortStringsPartition(ListArray, lLowerPoint, lUpperPoint, lOrder, lCompareType)
        
        ' Recurively sort the smaller partition
        If (lMidPoint - lLowerPoint) <= (lUpperPoint - lMidPoint) Then
            QuickSortStrings ListArray, lLowerPoint, lMidPoint - 1, lOrder, lCompareType
            lLowerPoint = lMidPoint + 1
        Else
            QuickSortStrings ListArray, lMidPoint + 1, lUpperPoint, lOrder, lCompareType
            lUpperPoint = lMidPoint - 1
        End If
    Loop
End Sub


'*******************************************************************************
' QuickSortStringsPartition (FUNCTION)
'
' PARAMETERS:
' (In/Out) - ListArray()  - String - Array to sort
' (In)     - lLow         - Long   - Start of sorting region within array
' (In)     - lHigh        - Long   - End of sorting region within array
' (In)     - lOrder       - Long   - Ascending is -1, Descending is +1, used
'                                    for comparison in StrComp
' (In)     - lCompareType - Long   - Either vbBinaryCompare or vbTextCompare,
'                                    used in StrComp function
'
' RETURN VALUE:
' Long - New pivot point
'
' DESCRIPTION:
' Selects a pivot point and moves smaller entries to one side of it and larger
' entries to the other side of it.  Returns the position of the pivot point
' when finished.
'*******************************************************************************
Private Function QuickSortStringsPartition(ListArray() As String, _
                                           ByVal lLow As Long, _
                                           ByVal lHigh As Long, _
                                           ByVal lOrder As Long, _
                                           ByVal lCompareType As Long) As Long

    Dim lPivot      As Long
    Dim sPivot      As String
    Dim lLowCount   As Long
    Dim lHighCount  As Long
    Dim sTemp       As String

    ' Select pivot point and exchange with first element
    lPivot = lLow + (lHigh - lLow) \ 2
    sPivot = ListArray(lPivot)
    ListArray(lPivot) = ListArray(lLow)
    
    lLowCount = lLow + 1
    lHighCount = lHigh
    
    ' Continually loop moving entries smaller than pivot to one side and
    ' larger than pivot to other side
    Do
        Do While lLowCount < lHighCount
            If StrComp(sPivot, ListArray(lLowCount), lCompareType) <> lOrder Then
                Exit Do
            Else
                lLowCount = lLowCount + 1
            End If
        Loop
        
        Do While lHighCount >= lLowCount
            If StrComp(ListArray(lHighCount), sPivot, lCompareType) <> lOrder Then
                Exit Do
            Else
                lHighCount = lHighCount - 1
            End If
        Loop
        
        If lLowCount >= lHighCount Then
            Exit Do
        End If
        
        ' Swap the items
        sTemp = ListArray(lLowCount)
        ListArray(lLowCount) = ListArray(lHighCount)
        ListArray(lHighCount) = sTemp
        
        lHighCount = lHighCount - 1
        lLowCount = lLowCount + 1
    Loop
    
    ListArray(lLow) = ListArray(lHighCount)
    ListArray(lHighCount) = sPivot
    QuickSortStringsPartition = lHighCount
End Function


'*******************************************************************************
' MergeSortStringsStart (SUB)
'
' PARAMETERS:
' (In/Out) - ListArray()    - String  - Array to sort
' (In)     - bAscending     - Boolean - True to sort ascending, false descending
' (In)     - bCaseSensitive - Boolean - True for a case sensitive sort, false
'                                       for an insensitive one
'
' DESCRIPTION:
' User friendly entry point for MergeSortStrings
'*******************************************************************************
Private Sub MergeSortStringsStart(ListArray() As String, _
                                 Optional ByVal bAscending As Boolean = True, _
                                 Optional ByVal bCaseSensitive As Boolean = False)

    Dim lMin            As Long
    Dim lMax            As Long
    Dim lOrder          As Long
    Dim lCompareType    As Long

    Const DELEGATE_POINT As Long = 60
    
    lMin = LBound(ListArray)
    lMax = UBound(ListArray)
    
    If lMin = lMax Then
        Exit Sub
    End If
    
    ' Order Ascending or Descending?
    lOrder = IIf(bAscending, 1, -1)
    
    ' Case sensitive search or not?
    lCompareType = IIf(bCaseSensitive, vbBinaryCompare, vbTextCompare)
    
    ' Delegate to insert sort for very small arrays for speed
    If (lMax - lMin) > DELEGATE_POINT Then
        MergeSortStrings ListArray, lMin, lMax, lOrder, lCompareType
    Else
        InsertSortStrings ListArray, lMin, lMax, lOrder, lCompareType
    End If
End Sub


'*******************************************************************************
' MergeSortStrings (SUB)
'
' PARAMETERS:
' (In/Out) - ListArray()  - String - Array to sort
' (In)     - lLowerPoint  - Long   - Start of sorting region within array
' (In)     - lUpperPoint  - Long   - End of sorting region within array
' (In)     - lOrder       - Long   - Ascending is -1, Descending is +1, used
'                                    for comparison in StrComp
' (In)     - lCompareType - Long   - Either vbBinaryCompare or vbTextCompare,
'                                    used in StrComp function
'
' DESCRIPTION:
' Quick for large arrays
'*******************************************************************************
Private Sub MergeSortStrings(ListArray() As String, _
                             ByVal lLowerPoint As Long, _
                             ByVal lUpperPoint As Long, _
                             ByVal lOrder As Long, _
                             ByVal lCompareType As Long)
    
    Dim lMidPoint As Long
    
    If lUpperPoint > lLowerPoint Then
        ' Split the array up recursively and sort (divide and conquer)
        lMidPoint = (lUpperPoint + lLowerPoint) \ 2
        MergeSortStrings ListArray, lLowerPoint, lMidPoint, lOrder, lCompareType
        MergeSortStrings ListArray, lMidPoint + 1, lUpperPoint, lOrder, lCompareType
        
        ' Merge to sort
        MergeStrings ListArray, lLowerPoint, lMidPoint, lUpperPoint, lOrder, lCompareType
    End If
End Sub


'*******************************************************************************
' MergeStrings (SUB)
'
' PARAMETERS:
' (In/Out) - ListArray()  - String - Array to sort
' (In)     - lLowerPoint  - Long   - Start of sorting region within array
' (In)     - lMidPoint    - Long   - Mid point
' (In)     - lUpperPoint  - Long   - End of sorting region within array
' (In)     - lOrder       - Long   - Ascending is -1, Descending is +1, used
'                                    for comparison in StrComp
' (In)     - lCompareType - Long   - Either vbBinaryCompare or vbTextCompare,
'                                    used in StrComp function
'
' DESCRIPTION:
' Merge part of the MergeSort that merges the two sorted parts lLowerPoint to
' lMidPoint and lMidPoint+1 to lUpperPoint
'*******************************************************************************
Private Sub MergeStrings(ListArray() As String, _
                         ByVal lLowerPoint As Long, _
                         ByVal lMidPoint As Long, _
                         ByVal lUpperPoint As Long, _
                         ByVal lOrder As Long, _
                         ByVal lCompareType As Long)
    
    Dim TempList()      As String
    Dim lcount          As Long
    Dim lBottomPointer  As Long
    Dim lTopPointer     As Long
    Dim lCurrentPointer As Long
    
    ' Prepare temporary array
    ReDim TempArray(lLowerPoint To lUpperPoint)
    
    ' Make a temporary copy of the array
    For lcount = lLowerPoint To lUpperPoint
        TempArray(lcount) = ListArray(lcount)
    Next
    
    ' Initialise pointers that will be used to move through array
    lBottomPointer = lLowerPoint
    lTopPointer = lMidPoint + 1
    lCurrentPointer = lLowerPoint
    
    ' Loop until we have got to the end of one section
    Do While (lBottomPointer <= lMidPoint And lTopPointer <= lUpperPoint)
        If StrComp(TempArray(lBottomPointer), TempArray(lTopPointer), lCompareType) <> lOrder Then
            ListArray(lCurrentPointer) = TempArray(lBottomPointer)
            lBottomPointer = lBottomPointer + 1
        Else
            ListArray(lCurrentPointer) = TempArray(lTopPointer)
            lTopPointer = lTopPointer + 1
        End If
        lCurrentPointer = lCurrentPointer + 1
    Loop
    
    ' Copy the rest of the uncompleted section onto the end
    Do While lBottomPointer <= lMidPoint
        ListArray(lCurrentPointer) = TempArray(lBottomPointer)
        lBottomPointer = lBottomPointer + 1
        lCurrentPointer = lCurrentPointer + 1
    Loop
    Do While lTopPointer <= lUpperPoint
        ListArray(lCurrentPointer) = TempArray(lTopPointer)
        lTopPointer = lTopPointer + 1
        lCurrentPointer = lCurrentPointer + 1
    Loop
End Sub

Public Function GetCustomDesc(ItemName As String, Optional Plural As Boolean = False) As String
On Error Resume Next
    Dim s As String
    s = ItemName
    s = HFApp.SqlExec("select isnull(nullif(custom_description,''),item) from customdescriptions where item=" & DbQuote(Str, ItemName), dbHomefront)(0)
    
    If Plural Then
        Select Case Right(s, 1)
            Case "s"
            Case "y":   s = Left(s, Len(s) - 1) & "ies"
            Case Else:  s = s & "s"
        End Select
    End If
    
    GetCustomDesc = s
End Function























Public Sub IniGetGrid(Form As Form, Grid As Object, Optional IniFile As String, Optional StaticPositions As Boolean = False, Optional InstanceKey, Optional Grouped As Boolean, Optional StaticNames As Boolean = False)
    If DBGetGrid(Form, Grid, StaticPositions, InstanceKey, Grouped, StaticNames) Then Exit Sub
    
    'this is temporary.
    'if the db doesnt have any data then load it from the file...
    
    Dim i         As Long
    Dim section   As String
    Dim variables As String
    Dim variable  As String
    Dim colCSV    As String
    Dim ColIndex  As Long
    Dim splitindex As Long
    Dim colPos()  As String
    Dim Caption As String

    Dim AllColsHidden As Boolean

    If IniFile = "" Then IniFile = AppIni

    section = Form.Name & "." & Grid.Name
    On Error Resume Next
    section = section & ".(" & Grid.Index & ")"
    On Error GoTo 0
    If Not IsMissing(InstanceKey) Then section = section & "-" & InstanceKey


    variables = IniGetVariableNames(IniFile, section)
    ReDim colPos(Parse(variables))
    If variables <> "" Then

        If TypeName(Grid) = "TDBGrid" Then
            For i = 1 To Parse(variables)
                variable = Parse(variables, i)
                colCSV = IniGet(IniFile, section, variable)
                If colCSV <> "" Then
                    On Error Resume Next
                    ColIndex = -1
                    splitindex = Val("" & Parse(colCSV, 5))
                    ColIndex = Grid.Splits(splitindex).columns(variable).ColIndex
                    If ColIndex <> -1 Then
                        Grid.Splits(splitindex).columns(ColIndex).Width = Val("" & Parse(colCSV, 1))
                        Grid.Splits(splitindex).columns(ColIndex).Visible = CBool(Parse(colCSV, 2))
                        If Not StaticPositions Then
                            Grid.Splits(splitindex).columns(variable).Order = Val("" & Parse(colCSV, 3))
                        End If
                        Grid.Splits(splitindex).columns(ColIndex).Caption = Parse(colCSV, 4)
                    End If
                    On Error GoTo 0
                End If
            Next
        Else
            AllColsHidden = True
            For i = 1 To Parse(variables)
                variable = Parse(variables, i)
                colCSV = IniGet(IniFile, section, variable)
                If colCSV <> "" Then
                    On Error Resume Next
                    ColIndex = -1
                    ColIndex = Grid.ColIndex(variable)
                    If ColIndex <> -1 Then
                        Grid.ColWidth(ColIndex) = Val("" & Parse(colCSV, 1))

                        If Not CBool(Parse(colCSV, 2)) Then AllColsHidden = False

                        Grid.ColHidden(ColIndex) = CBool(Parse(colCSV, 2))
                        If Not StaticNames Then
                            Caption = Parse(colCSV, 4)
                            If Caption <> "" Then Grid.TextMatrix(0, ColIndex) = Caption
                        End If
                        
                        If Grouped Then
                            Grid.ColData(ColIndex) = IIf(Parse(colCSV, 5) = "True", "GROUPED", "")
                        End If

                        If Not StaticPositions Then
                            Grid.ColPosition(ColIndex) = Val("" & Parse(colCSV, 3))
                        End If


                    End If
                    On Error GoTo 0
                End If
            Next
            If AllColsHidden Then
                'something is wrong. show all cols that have a heading
                For i = 1 To Grid.Cols - 1
                    Grid.ColHidden(i) = Grid.TextMatrix(0, i) = ""
                Next
            End If
        End If
    End If

End Sub

Public Sub IniPutGrid(Form As Form, Grid As Object, Optional IniFile As String, Optional InstanceKey)
Call DBPutGrid(Form, Grid, InstanceKey)
Exit Sub
 
' On Error Resume Next
'    Dim i As Long
'    Dim Section As String
'    Dim split As Long
'    Dim Caption As String
'
'    If IniFile = "" Then IniFile = AppIni
'
'    Section = Form.Name & "." & Grid.Name
'    Section = Section & ".(" & Grid.Index & ")"
'    If Not IsMissing(InstanceKey) Then Section = Section & "-" & InstanceKey
'
'    Call IniRemove(IniFile, Section)
'
'    If TypeName(Grid) = "TDBGrid" Then
'        For split = 0 To Grid.Splits.Count - 1
'            For i = 0 To Grid.Splits(split).Columns.Count - 1
'                Call IniPut(IniFile, Section, Grid.Splits(split).Columns(i).DataField, csv(Grid.Splits(split).Columns(i).Width, Grid.Splits(split).Columns(i).Visible, Grid.Splits(split).Columns(i).Order, Grid.Splits(split).Columns(i).Caption, split))
'            Next
'        Next
'    Else
'        For i = 0 To Grid.cols - 1
'            Caption = Grid.TextMatrix(0, i)
'            If Grid.ColKey(i) <> "" Then Call IniPut(IniFile, Section, Grid.ColKey(i), csv(Grid.ColWidth(i), Grid.ColHidden(i), i, Caption, Grid.ColData(i) = "GROUPED"))
'        Next
'    End If
'
'    Exit Sub
    
    
eh: Call errHandler("IniPutGrid")
End Sub

Public Function DBGetGrid(Form As Form, Grid As Object, Optional StaticPositions As Boolean = False, Optional InstanceKey, Optional Grouped As Boolean, Optional StaticNames As Boolean) As Boolean
On Error Resume Next
'return false if nothing found...

    Dim gridname  As String
    Dim Col       As Long
    Dim split     As Long
    Dim s         As String
    Dim rs        As Recordset
    Dim AllColsHidden     As Boolean
    
    With Grid
    
        'gridname = FormName.GridName(Index).InstanceKey
        gridname = Form.Name & "." & .Name
        gridname = gridname & ".(" & .Index & ")"
        If Not IsMissing(InstanceKey) Then gridname = gridname & "." & InstanceKey
    
        s = ""
        s = s & "select * from AppGridLayout" & vbCrLf
        s = s & "where uid=" & DbQuote(Str, HFApp.LoginID) & vbCrLf
        s = s & "and gridname=" & DbQuote(Str, gridname) & vbCrLf
        Set rs = HFApp.SqlExec(s)

        If rs.EOF Then
            DBGetGrid = False
            Exit Function
        End If
            
    
        If TypeName(Grid) = "TDBGrid" Then
            While Not rs.EOF
            
                split = Val("" & rs("split"))
                Col = .Splits(split).columns("" & rs("colkey")).ColIndex
                
                .Splits(split).columns(Col).Caption = "" & rs("caption")
                .Splits(split).columns(Col).Width = Val("" & rs("width"))
                .Splits(split).columns(Col).Visible = "" & rs("hidden") <> "True"
                If Not StaticPositions Then .Splits(split).columns(Col).Order = Val("" & rs("colindex"))
    
                rs.MoveNext
            Wend
        
        Else
        
            AllColsHidden = True
            While Not rs.EOF
                Col = .ColIndex("" & rs("colkey"))
                If Col <> -1 Then
                    If Not StaticNames Then .TextMatrix(0, Col) = "" & rs("caption")
                    .ColWidth(Col) = Val("" & rs("width"))
                    .ColHidden(Col) = "" & rs("hidden") = "True" Or .TextMatrix(0, Col) = ""
                    .ColData(Col) = IIf("" & rs("grouped") = "True", "GROUPED", "")
                    AllColsHidden = AllColsHidden And "" & rs("hidden") = "True"
                    If Not StaticPositions Then .ColPosition(Col) = Val("" & rs("colindex"))
                End If
                rs.MoveNext
            Wend
            
            If AllColsHidden Then 'something is wrong. show everything
                For Col = 1 To .Cols - 1
                    .ColHidden(Col) = .TextMatrix(0, Col) = ""
                Next
            End If
            
        End If
    
    
        
    
    End With
    DBGetGrid = True

End Function

Public Sub DBPutGrid(Form As Form, Grid As Object, Optional InstanceKey)
On Error Resume Next
 
    Dim sql As String
    
    Dim gridname As String
    Dim split As Long
    Dim Col As Long
    
    
    sql = ""
    With Grid
    
        'gridname = FormName.GridName(Index).InstanceKey
        gridname = Form.Name & "." & .Name
        gridname = gridname & ".(" & .Index & ")"
        If Not IsMissing(InstanceKey) Then gridname = gridname & "." & InstanceKey
            
            
        If TypeName(Grid) = "TDBGrid" Then
            For split = 0 To Grid.Splits.Count - 1
                For Col = 0 To Grid.columns.Count - 1
                    'uid,gridname,split,colindex,colkey,caption,width,hidden,grouped
                    sql = sql & ",(" & DbQuote(Str, HFApp.LoginID) & _
                                "," & DbQuote(Str, gridname) & _
                                "," & DbQuote(Num, split) & _
                                "," & DbQuote(Num, Grid.Splits(split).columns(Col).Order) & _
                                "," & DbQuote(Str, Grid.Splits(split).columns(Col).DataField) & _
                                "," & DbQuote(Str, Grid.Splits(split).columns(Col).Caption) & _
                                "," & DbQuote(Num, Grid.Splits(split).columns(Col).Width) & _
                                "," & DbQuote(Bit, Grid.Splits(split).columns(Col).Visible) & _
                                "," & DbQuote(Bit, False) & vbCrLf
                Next
            Next
        Else
            For Col = 0 To .Cols - 1
                'uid,gridname,split,colindex,colkey,caption,width,hidden,grouped
                sql = sql & ",(" & DbQuote(Str, HFApp.LoginID) & _
                            "," & DbQuote(Str, gridname) & _
                            ",0" & _
                            "," & DbQuote(Num, Col) & _
                            "," & DbQuote(Str, .ColKey(Col)) & _
                            "," & DbQuote(Str, .TextMatrix(0, Col)) & _
                            "," & DbQuote(Num, .ColWidth(Col)) & _
                            "," & DbQuote(Bit, .ColHidden(Col)) & _
                            "," & DbQuote(Bit, .ColData(Col) = "GROUPED") & ")" & vbCrLf
            Next
        End If
    End With
    sql = "delete AppGridLayout" & vbCrLf & _
          "where uid=" & DbQuote(Str, HFApp.LoginID) & vbCrLf & _
          "and gridname=" & DbQuote(Str, gridname) & vbCrLf & _
          vbCrLf & _
          "insert into AppGridLayout(uid,gridname,split,colindex,colkey,caption,width,hidden,grouped) values" & vbCrLf & _
          Mid(sql, 2)
    
    Call HFApp.SqlExec(sql)
    

Exit Sub
eh: Call errHandler("IniPutGrid")
End Sub



Public Sub UpdateScheduleVendor(Job As String, OldVendorCSV As String, NewVendor As String)
On Error GoTo eh
    Dim s As String
    Dim rCount As Long
    Dim r As Long
    
    
    s = ""
    s = s & "select count(*)" & vbCrLf
    s = s & "from dbo.Schedule s" & vbCrLf
    s = s & "join dbo.ScheduleTasks t on (s.scheduleid=t.scheduleid)" & vbCrLf
    s = s & "JOIN dbo.RESERVATIONS r on (t.scheduletaskid = r.scheduletaskid)" & vbCrLf
    s = s & "where s.job_no = " & DbQuote(Str, Job) & vbCrLf
    s = s & "and resourceid in(" & OldVendorCSV & ")"
    rCount = HFApp.SqlExec(s, dbHomefront)(0)
    
    If rCount > 0 Then
         If MsgBox("Would you like to update the vendor assignment on the job schedule? The old vendor will be removed and replaced with the new vendor.", vbYesNo + vbQuestion, App.ProductName) = vbYes Then
            s = ""
            s = s & "select t.ScheduleTaskid,t.Description Task,v.Vendor_Name [Currently Assigned To]" & vbCrLf
            s = s & "from dbo.Schedule s" & vbCrLf
            s = s & "join dbo.ScheduleTasks t on (s.scheduleid=t.scheduleid)" & vbCrLf
            s = s & "join dbo.RESERVATIONS r on (t.scheduletaskid = r.scheduletaskid)" & vbCrLf
            s = s & "left outer join tblvendors v on s.divisionid=v.divisionid and r.resourceid=v.vendor_id" & vbCrLf
            s = s & "where s.divisionid= " & DbQuote(Num, HFApp.DivisionID) & vbCrLf
            s = s & "and s.job_no = " & DbQuote(Str, Job) & vbCrLf
            s = s & "and resourceid in(" & OldVendorCSV & ")" & vbCrLf
            s = s & "order by s.ScheduleID"
            If FPickList.Choose(HFApp.Databases(dbHomefront), "Schedule Task", s, "", , , , IIf(HFApp.Options(AccountingSystem) = asQuickBooks, "Vendor,ScheduleTaskid", "ScheduleTaskid"), True) Then
                For r = 1 To FPickList.SelectedItems
                
                    s = "Delete from reservations where resourceid in(" & OldVendorCSV & ") and scheduletaskid = " & FPickList.SelectedItem("ScheduleTaskID", r) & vbCrLf
                    Call HFApp.SqlExec(s)
                    
                    s = "insert reservations(resourceid,scheduletaskid) values(" & DbQuote(Str, NewVendor) & "," & FPickList.SelectedItem("ScheduleTaskID", r) & ")" & vbCrLf
                    On Error Resume Next
                    Call HFApp.SqlExec(s)
                    On Error GoTo eh
                    
                Next
            End If
        End If
    End If
    
    

Exit Sub
eh: Call errHandler(SRCFILE & "UpdateSchedule")
End Sub


Public Sub GridExpandALL(g As VSFlexGrid)
    Dim r As Long
    With g
    r = 0
    While r < .Rows - 1
        r = r + 1
        .IsCollapsed(r) = flexOutlineExpanded
    Wend
    .Row = -1
    End With
End Sub

Public Sub SendingWizard(DocType As String, Optional Job As String = "", Optional PONumbers As String = "", Optional SpecifiedPOsOnly As Boolean = False)
    Dim s As String
    s = SyncCmd("SendingWizard", DocType & "|" & Job & "|" & PONumbers & "|" & SpecifiedPOsOnly, True, True)
    If HFApp.UserPermission("SendPO") Then
        Call Shell(s)
    End If
End Sub

Public Function SyncCmd(cmd As String, Optional parameters As String, Optional SkipAccounting As Boolean = False, Optional HFSend As Boolean) As String

    If InIde() Then
        SyncCmd = "C:\Program Files (x86)\HomeFront" & IIf(HFSend, "\HFSend.exe ", "\HFSync.exe ") & _
                  HFApp.ConnectionString(dbHomefront) & Chr(1) & _
                  HFApp.DivisionID & Chr(1) & _
                  IIf(SkipAccounting, 1, 0) & Chr(1) & _
                  cmd & Chr(1) & _
                  parameters
    Else
        SyncCmd = App.Path & IIf(HFSend, "\HFSend.exe ", "\HFSync.exe ") & _
                  HFApp.ConnectionString(dbHomefront) & Chr(1) & _
                  HFApp.DivisionID & Chr(1) & _
                  IIf(SkipAccounting, 1, 0) & Chr(1) & _
                  cmd & Chr(1) & _
                  parameters
    End If
    
End Function

Public Function CleanXML(ByVal s As String) As String
    CleanXML = s
    Exit Function
    
    s = Replace(s, "<removeme>", "")
    s = Replace(s, "</removeme>", "")
    
    
    'single quotes
    s = Replace(s, Chr(145), Chr(39)) 'left curly single quote
    s = Replace(s, Chr(146), Chr(39)) 'right curly single quote
    s = Replace(s, Chr(147), Chr(34)) 'left curly double quote
    s = Replace(s, Chr(148), Chr(34)) 'right curly double quote
    s = Replace(s, Chr(152), Chr(34)) 'another weird double quote
        
    s = Replace(s, Chr(96), Chr(39))  'left accent
    s = Replace(s, Chr(180), Chr(39)) 'right accent
    s = Replace(s, "&lt;", "<")
    s = Replace(s, "&gt;", ">")
        
    s = Replace(s, Chr(188), "1/4")
    s = Replace(s, Chr(189), "1/2")
    s = Replace(s, Chr(190), "3/4")
    s = Replace(s, Chr(150), "-")
    
    CleanXML = s
    
End Function

Public Function AddQuotes(s As String) As String
    AddQuotes = vbQuote & s & vbQuote
End Function
