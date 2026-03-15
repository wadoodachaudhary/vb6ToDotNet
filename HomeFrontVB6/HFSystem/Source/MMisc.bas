Attribute VB_Name = "MStuffThatShouldBeInVB"
Option Explicit
Option Compare Text
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long



Const CB_SETDROPPEDWIDTH = &H160


Public Type MAPIMessage
    Reserved As Long
    Subject As String
    NoteText As String
    MessageType As String
    DateReceived As String
    ConversationID As String
    flags As Long
    RecipCount As Long
    FileCount As Long
End Type
Public Type MapiRecip
    Reserved As Long
    RecipClass As Long
    Name As String
    Address As String
    EIDSize As Long
    EntryID As String
End Type
Public Type MapiFile
    Reserved As Long
    flags As Long
    Position As Long
    PathName As String
    Filename As String
    FileType As String
End Type
Public Declare Function MAPILogon Lib "MAPI32.DLL" (ByVal UIParam&, ByVal User$, ByVal Password$, ByVal flags&, ByVal Reserved&, Session&) As Long
Public Declare Function MAPILogoff Lib "MAPI32.DLL" (ByVal Session&, ByVal UIParam&, ByVal flags&, ByVal Reserved&) As Long
Public Declare Function BMAPIReadMail Lib "MAPI32.DLL" (lMsg&, nRecipients&, nFiles&, ByVal Session&, ByVal UIParam&, MessageID$, ByVal Flag&, ByVal Reserved&) As Long
Public Declare Function BMAPIGetReadMail Lib "MAPI32.DLL" (ByVal lMsg&, Message As MAPIMessage, Recip() As MapiRecip, file() As MapiFile, Originator As MapiRecip) As Long
Public Declare Function MAPIFindNext Lib "MAPI32.DLL" Alias "BMAPIFindNext" (ByVal Session&, ByVal UIParam&, MsgType$, SeedMsgID$, ByVal Flag&, ByVal Reserved&, MsgId$) As Long
Public Declare Function MAPISendDocuments Lib "MAPI32.DLL" (ByVal UIParam&, ByVal DelimStr$, ByVal FilePaths$, ByVal FileNames$, ByVal Reserved&) As Long
Public Declare Function MAPIDeleteMail Lib "MAPI32.DLL" (ByVal Session&, ByVal UIParam&, ByVal MsgId$, ByVal flags&, ByVal Reserved&) As Long
Public Declare Function MAPISendMail Lib "MAPI32.DLL" Alias "BMAPISendMail" (ByVal Session&, ByVal UIParam&, Message As MAPIMessage, Recipient() As MapiRecip, file() As MapiFile, ByVal flags&, ByVal Reserved&) As Long
Public Declare Function MAPISaveMail Lib "MAPI32.DLL" Alias "BMAPISaveMail" (ByVal Session&, ByVal UIParam&, Message As MAPIMessage, Recipient() As MapiRecip, file() As MapiFile, ByVal flags&, ByVal Reserved&, MsgId$) As Long
Public Declare Function BMAPIAddress Lib "MAPI32.DLL" (lInfo&, ByVal Session&, ByVal UIParam&, Caption$, ByVal nEditFields&, Label$, nRecipients&, Recip() As MapiRecip, ByVal flags&, ByVal Reserved&) As Long
Public Declare Function BMAPIGetAddress Lib "MAPI32.DLL" (ByVal lInfo&, ByVal nRecipients&, Recipients() As MapiRecip) As Long
Public Declare Function MAPIDetails Lib "MAPI32.DLL" Alias "BMAPIDetails" (ByVal Session&, ByVal UIParam&, Recipient As MapiRecip, ByVal flags&, ByVal Reserved&) As Long
Public Declare Function MAPIResolveName Lib "MAPI32.DLL" Alias "BMAPIResolveName" (ByVal Session&, ByVal UIParam&, ByVal UserName$, ByVal flags&, ByVal Reserved&, Recipient As MapiRecip) As Long
Public Const SUCCESS_SUCCESS = 0
Public Const MAPI_USER_ABORT = 1
Public Const MAPI_E_USER_ABORT = MAPI_USER_ABORT
Public Const MAPI_E_FAILURE = 2
Public Const MAPI_E_LOGIN_FAILURE = 3
Public Const MAPI_E_LOGON_FAILURE = MAPI_E_LOGIN_FAILURE
Public Const MAPI_E_DISK_FULL = 4
Public Const MAPI_E_INSUFFICIENT_MEMORY = 5
Public Const MAPI_E_BLK_TOO_SMALL = 6
Public Const MAPI_E_TOO_MANY_SESSIONS = 8
Public Const MAPI_E_TOO_MANY_FILES = 9
Public Const MAPI_E_TOO_MANY_RECIPIENTS = 10
Public Const MAPI_E_ATTACHMENT_NOT_FOUND = 11
Public Const MAPI_E_ATTACHMENT_OPEN_FAILURE = 12
Public Const MAPI_E_ATTACHMENT_WRITE_FAILURE = 13
Public Const MAPI_E_UNKNOWN_RECIPIENT = 14
Public Const MAPI_E_BAD_RECIPTYPE = 15
Public Const MAPI_E_NO_MESSAGES = 16
Public Const MAPI_E_INVALID_MESSAGE = 17
Public Const MAPI_E_TEXT_TOO_LARGE = 18
Public Const MAPI_E_INVALID_SESSION = 19
Public Const MAPI_E_TYPE_NOT_SUPPORTED = 20
Public Const MAPI_E_AMBIGUOUS_RECIPIENT = 21
Public Const MAPI_E_AMBIG_RECIP = MAPI_E_AMBIGUOUS_RECIPIENT
Public Const MAPI_E_MESSAGE_IN_USE = 22
Public Const MAPI_E_NETWORK_FAILURE = 23
Public Const MAPI_E_INVALID_EDITFIELDS = 24
Public Const MAPI_E_INVALID_RECIPS = 25
Public Const MAPI_E_NOT_SUPPORTED = 26
Public Const MAPI_ORIG = 0
Public Const MAPI_TO = 1
Public Const MAPI_CC = 2
Public Const MAPI_BCC = 3
Public Const MAPI_LOGON_UI = &H1
Public Const MAPI_NEW_SESSION = &H2
Public Const MAPI_FORCE_DOWNLOAD = &H1000
Public Const MAPI_LOGOFF_SHARED = &H1
Public Const MAPI_LOGOFF_UI = &H2
Public Const MAPI_DIALOG = &H8
Public Const MAPI_UNREAD_ONLY = &H20
Public Const MAPI_GUARANTEE_FIFO = &H100
Public Const MAPI_ENVELOPE_ONLY = &H40
Public Const MAPI_PEEK = &H80
Public Const MAPI_BODY_AS_FILE = &H200
Public Const MAPI_SUPPRESS_ATTACH = &H800
Public Const MAPI_AB_NOMODIFY = &H400
Public Const MAPI_OLE = &H1
Public Const MAPI_OLE_STATIC = &H2
Public Const MAPI_UNREAD = &H1
Public Const MAPI_RECEIPT_REQUESTED = &H2
Public Const MAPI_SENT = &H4

Public Enum SortAlgorithms
    InsertSort
    MergeSort
    QuickSOrt
    SelectionSort
End Enum

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Const COLOR_SCROLLBAR = 0
Public Const COLOR_BACKGROUND = 1
Public Const COLOR_ACTIVECAPTION = 2
Public Const COLOR_INACTIVECAPTION = 3
Public Const COLOR_MENU = 4
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWFRAME = 6
Public Const COLOR_MENUTEXT = 7
Public Const COLOR_WINDOWTEXT = 8
Public Const COLOR_CAPTIONTEXT = 9
Public Const COLOR_ACTIVEBORDER = 10
Public Const COLOR_INACTIVEBORDER = 11
Public Const COLOR_APPWORKSPACE = 12
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_BTNFACE = 15
Public Const COLOR_BTNSHADOW = 16
Public Const COLOR_GRAYTEXT = 17
Public Const COLOR_BTNTEXT = 18
Public Const COLOR_INACTIVECAPTIONTEXT = 19
Public Const COLOR_BTNHIGHLIGHT = 20

Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long

Public Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Const CLR_INVALID = -1

Public Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function CreateEllipticRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Public Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Type RECT
        left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type

' Brush Styles
Public Const BS_SOLID = 0
Public Const BS_NULL = 1
Public Const BS_HOLLOW = BS_NULL
Public Const BS_HATCHED = 2
Public Const BS_PATTERN = 3
Public Const BS_INDEXED = 4
Public Const BS_DIBPATTERN = 5
Public Const BS_DIBPATTERNPT = 6
Public Const BS_PATTERN8X8 = 7
Public Const BS_DIBPATTERN8X8 = 8

'  Hatch Styles
Public Const HS_HORIZONTAL = 0              '  -----
Public Const HS_VERTICAL = 1                '  |||||
Public Const HS_FDIAGONAL = 2               '  \\\\\
Public Const HS_BDIAGONAL = 3               '  /////
Public Const HS_CROSS = 4                   '  +++++
Public Const HS_DIAGCROSS = 5               '  xxxxx
Public Const HS_FDIAGONAL1 = 6
Public Const HS_BDIAGONAL1 = 7
Public Const HS_SOLID = 8


' DrawText() Format Flags
Public Const DT_Top = &H0
Public Const DT_LEFT = &H0
Public Const DT_CENTER = &H1
Public Const DT_RIGHT = &H2
Public Const DT_VCENTER = &H4
Public Const DT_BOTTOM = &H8
Public Const DT_WORDBREAK = &H10
Public Const DT_SINGLELINE = &H20
Public Const DT_EXPANDTABS = &H40
Public Const DT_TABSTop = &H80
Public Const DT_NOCLIP = &H100
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_CALCRECT = &H400
Public Const DT_NOPREFIX = &H800
Public Const DT_INTERNAL = &H1000

Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Declare Function SelectObject& Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long)
Declare Function GetStockObject& Lib "gdi32" (ByVal nIndex As Long)
Public Const SYSTEM_FONT = 13



Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Declare Function ReleaseCapture& Lib "user32" ()
Public Declare Function SendMessage& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any)
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetRect& Lib "user32" (lpRect As RECT, ByVal left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)




Type POINTAPI
        X As Long
        Y As Long
End Type
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long






Public Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Public Declare Function InitCommonControlsEx Lib "COMCTL32.DLL" (iccex As tagInitCommonControlsEx) As Boolean
Public Declare Sub InitCommonControls Lib "COMCTL32.DLL" ()
Public Const ICC_USEREX_CLASSES = &H200
Public Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Public Const SEM_NOGPFAULTERRORBOX = &H2


' Commands to pass WinHelp()
Public Const HELP_CONTEXT = &H1          '  Display Topic in ulTopic
Public Const HELP_QUIT = &H2             '  Terminate help
Public Const HELP_INDEX = &H3            '  Display index
Public Const HELP_CONTENTS = &H3&
Public Const HELP_HELPONHELP = &H4       '  Display help on using help
Public Const HELP_SETINDEX = &H5         '  Set current Index for multi index help
Public Const HELP_SETCONTENTS = &H5&
Public Const HELP_CONTEXTPOPUP = &H8&
Public Const HELP_FORCEFILE = &H9&
Public Const HELP_KEY = &H101            '  Display Topic for keyword in offabData
Public Const HELP_COMMAND = &H102&
Public Const HELP_PARTIALKEY = &H105&
Public Const HELP_MULTIKEY = &H201&
Public Const HELP_SETWINPOS = &H203&
Public Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long


Public Declare Function DrawEdge& Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long)
Public Const BDR_INNER = &HC
Public Const BDR_OUTER = &H3
Public Const BDR_RAISED = &H5
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKEN = &HA
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2

Public Const EDGE_BUMP = &H9&
Public Const EDGE_ETCHED = &H6&
Public Const EDGE_RAISED = &H5&
Public Const EDGE_SUNKEN = &HA&

Public Const BF_ADJUST = &H2000
Public Const BF_BOTTOM = &H8
Public Const BF_BOTTOMLEFT = &H9
Public Const BF_BOTTOMRIGHT = &HC
Public Const BF_DIAGONAL = &H10
Public Const BF_FLAT = &H4000
Public Const BF_LEFT = &H1
Public Const BF_MIDDLE = &H800
Public Const BF_MONO = &H8000&
Public Const BF_RECT = &HF
Public Const BF_RIGHT = &H4
Public Const BF_SOFT = &H1000
Public Const BF_Top = &H2
Public Const BF_TopLEFT = &H3
Public Const BF_TopRIGHT = &H6

Public Const ILD_TRANSPARENT = &H1 'display transparent
Public Declare Function ImageList_Draw Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal i As Long, ByVal hDCDest As Long, ByVal X As Long, ByVal Y As Long, ByVal flags As Long) As Long
Public Declare Function PathIsNetworkPath Lib "shlwapi.dll" Alias "PathIsNetworkPathA" (ByVal pszPath As String) As Long
 
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long


Public Declare Function PathCompactPath Lib "shlwapi" Alias "PathCompactPathA" (ByVal hDC As Long, ByVal lpszPath As String, ByVal dx As Long) As Long


Public Const BUFFER_SIZE = 32767
Public Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As Any, ByVal lpDefaultValue As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetComputerName& Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long)
 
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
' SetWindowPos Flags
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_FRAMECHANGED = &H20         '  The frame changed: send WM_NCCALCSIZE
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOOWNERZORDER = &H200       '  Don't do owner Z ordering
Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
' SetWindowPos() hWndInsertAfter Values
Public Const HWND_Top = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TopMOST = -1
Public Const HWND_NOTopMOST = -2

'
' Win32 Registry functions
'
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



'
' public const ants for Windows 32-bit Registry API
'
Enum RegistryHiveConstants
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum
'
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
'
' Reg Key Security Options
'
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
Public Const REG_WOW64 = &H100&

Public Const ERROR_SUCCESS = 0&
Public Const ERROR_MORE_DATA = 234
Public Const ERROR_NO_MORE_DATA = 259

Public Const REG_SZ = 1                          ' Unicode nul terminated string

' Clipboard Manager Functions
Public Declare Function EmptyClipboard Lib "user32" () As Long
Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CloseClipboard Lib "user32" () As Long
Public Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Public Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Public Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long

' Other required Win32 APIs
Public Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
Public Declare Function DragQueryPoint Lib "shell32.dll" (ByVal hDrop As Long, lpPoint As POINTAPI) As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


'Form MinMax Size stuff
Public Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type
Public Const WM_GETMINMAXINFO = &H24


' Predefined Clipboard Formats
Public Const CF_TEXT = 1
Public Const CF_BITMAP = 2
Public Const CF_METAFILEPICT = 3
Public Const CF_SYLK = 4
Public Const CF_DIF = 5
Public Const CF_TIFF = 6
Public Const CF_OEMTEXT = 7
Public Const CF_DIB = 8
Public Const CF_PALETTE = 9
Public Const CF_PENDATA = 10
Public Const CF_RIFF = 11
Public Const CF_WAVE = 12
Public Const CF_UNICODETEXT = 13
Public Const CF_ENHMETAFILE = 14
Public Const CF_HDROP = 15
Public Const CF_LOCALE = 16
Public Const CF_MAX = 17

' New shell-oriented clipboard formats
Public Const CFSTR_SHELLIDLIST As String = "Shell IDList Array"
Public Const CFSTR_SHELLIDLISTOFFSET As String = "Shell Object Offsets"
Public Const CFSTR_NETRESOURCES As String = "Net Resource"
Public Const CFSTR_FILEDESCRIPTOR As String = "FileGroupDescriptor"
Public Const CFSTR_FILECONTENTS As String = "FileContents"
Public Const CFSTR_FileName As String = "FileName"
Public Const CFSTR_PRINTERGROUP As String = "PrinterFriendlyName"
Public Const CFSTR_FileNameMAP As String = "FileNameMap"

' Global Memory Flags
Public Const GMEM_FIXED = &H0
Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_NOCOMPACT = &H10
Public Const GMEM_NODISCARD = &H20
Public Const GMEM_ZEROINIT = &H40
Public Const GMEM_MODIFY = &H80
Public Const GMEM_DISCARDABLE = &H100
Public Const GMEM_NOT_BANKED = &H1000
Public Const GMEM_SHARE = &H2000
Public Const GMEM_DDESHARE = &H2000
Public Const GMEM_NOTIFY = &H4000
Public Const GMEM_LOWER = GMEM_NOT_BANKED
Public Const GMEM_VALID_FLAGS = &H7F72
Public Const GMEM_INVALID_HANDLE = &H8000
Public Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Public Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

'Structure used by CF_HDROP
'   typedef struct _DROPFILES {
'       DWORD pFiles; // offset of file list
'       POINT pt;     // drop point (coordinates depend on fNC)
'       BOOL fNC;     // see below
'       BOOL fWide;   // TRUE if file contains wide characters,
'                     // FALSE otherwise
'   } DROPFILES, FAR * LPDROPFILES;
Public Type DROPFILES
   pFiles As Long
   pt As POINTAPI
   fNC As Long
   fWide As Long
End Type

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpszOp As String, ByVal lpszFile As String, ByVal lpszParams As String, ByVal lpszDir As String, ByVal FsShowCmd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long

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


Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long


Public Const EM_CANUNDO = &HC6
Public Const EM_UNDO = &HC7


Private Type Guid
   Data1 As Long
   Data2 As Long
   Data3 As Long
   Data4(8) As Byte
End Type
Private Declare Function CoCreateGuid Lib "ole32.dll" (pguid As Guid) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long

Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long








Private Type WINDOWPLACEMENT
   Length            As Long
   flags             As Long
   showCmd           As Long
   ptMinPosition     As POINTAPI
   ptMaxPosition     As POINTAPI
   rcNormalPosition  As RECT
End Type

Private Declare Function GetWindowPlacement Lib "user32" _
   (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

Private Declare Function SetWindowPlacement Lib "user32" _
   (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long








' Win32 API declares
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrn As Long, pDefault As Any) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrn As Long) As Long
Private Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal cbBuf As Long, pcbNeeded As Long) As Long
Private Declare Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal Command As Long) As Long
Private Declare Function EnumPrinters Lib "winspool.drv" Alias "EnumPrintersA" (ByVal flags As Long, ByVal Name As String, ByVal Level As Long, pPrinterEnum As Any, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Private Declare Function PrinterProperties Lib "winspool.drv" (ByVal hwnd As Long, ByVal hPrinter As Long) As Long

Private Declare Function GetDefaultPrinter Lib "winspool.drv" Alias "GetDefaultPrinterA" (ByVal pszBuffer As String, pcchBuffer As Long) As Long
Private Declare Function SetDefaultPrinter Lib "winspool.drv" Alias "SetDefaultPrinterA" (ByVal pszPrinter As String) As Long
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function GetVersion Lib "kernel32" Alias "GetVersionA" (lpVersionInformation As Any) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As Any) As Long

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

'  SendMessageTimeout values
Private Const SMTO_NORMAL = &H0
Private Const SMTO_BLOCK = &H1
Private Const SMTO_ABORTIFHUNG = &H2

' Used with SendMessageTimeout to tell all apps of changes
Private Const HWND_BROADCAST = &HFFFF&
Private Const WM_SETTINGCHANGE = &H1A

' Need defaults to OpenPrinter in some cases
Private Type PRINTER_DEFAULTS
   pDatatype As String
   pDevMode As Long
   pDesiredAccess As Long
End Type

Private Const PRINTER_ACCESS_ADMINISTER = &H4
Private Const PRINTER_ACCESS_USE = &H8
Private Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

' Used to retrieve last API error text.
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000

' The data area passed to a system call is too small.
Private Const ERROR_INSUFFICIENT_BUFFER As Long = 122

' Used to indicate what to enumerate
Private Const PRINTER_ENUM_DEFAULT         As Long = &H1
Private Const PRINTER_ENUM_LOCAL           As Long = &H2
Private Const PRINTER_ENUM_CONNECTIONS     As Long = &H4
Private Const PRINTER_ENUM_FAVORITE        As Long = &H4
Private Const PRINTER_ENUM_NAME            As Long = &H8
Private Const PRINTER_ENUM_REMOTE          As Long = &H10
Private Const PRINTER_ENUM_SHARED          As Long = &H20
Private Const PRINTER_ENUM_NETWORK         As Long = &H40

' Printer control codes
Private Const PRINTER_CONTROL_PAUSE        As Long = 1
Private Const PRINTER_CONTROL_RESUME       As Long = 2
Private Const PRINTER_CONTROL_PURGE        As Long = 3
Private Const PRINTER_CONTROL_SET_STATUS   As Long = 4

Private Enum PrinterControlCodes
   pcPause = PRINTER_CONTROL_PAUSE
   pcResume = PRINTER_CONTROL_RESUME
   pcPurge = PRINTER_CONTROL_PURGE
End Enum

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

' VB5-friendly structure used to cache the values in this class.
Private Type PrinterInfo2
   pServerName As String
   pPrinterName As String
   pShareName As String
   pPortName As String
   pDriverName As String
   pComment As String
   pLocation As String
   pDevMode As Long 'DEVMODE
   pSepFile As String
   pPrintProcessor As String
   pDatatype As String
   pParameters As String
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

Public Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long
Public Declare Function PathIsUNC Lib "shlwapi.dll" Alias "PathIsUNCA" (ByVal pszPath As String) As Long
Public Declare Function PathStripToRoot Lib "shlwapi.dll" Alias "PathStripToRootA" (ByVal pPath As String) As Long
Public Declare Function PathSkipRoot Lib "shlwapi.dll" Alias "PathSkipRootA" (ByVal pPath As String) As Long
Public Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Public Declare Function lstrcpyA Lib "kernel32" (ByVal RetVal As String, ByVal Ptr As Long) As Long



Private Type MungeLong
    X As Long
    Dummy As Integer
End Type

Private Type MungeInt
    XLo As Integer
    XHi As Integer
    Dummy As Integer
End Type

Private Declare Function NetShareGetInfo Lib "NETAPI32" (ByRef ServerName As Byte, ByRef NetName As Byte, ByVal Level As Long, ByRef Buffer As Long) As Long
Private Declare Function NetAPIBufferFree Lib "netapi32.dll" Alias "NetApiBufferFree" (bufptr As Any) As Long
Private Declare Function PtrToInt Lib "kernel32" Alias "lstrcpynW" (RetVal As Any, ByVal Ptr As Long, ByVal nCharCount As Long) As Long
Private Declare Function PtrToStr Lib "kernel32" Alias "lstrcpyW" (RetVal As Byte, ByVal Ptr As Long) As Long
Private Declare Function StrLen Lib "kernel32" Alias "lstrlenW" (ByVal Ptr As Long) As Long


Public Function GetLocalizedPath(sPath As String) As String
'Dim d As String

    Dim s As String

    If Not IsPathNetPath(sPath) Then
'd = d & sPath & " is already a local path" & vbCrLf
'MsgBox d, vbInformation, "GetLocalizedPath"
        GetLocalizedPath = sPath
    Else
        If PathIsUNC(sPath) = 1 Then
'd = d & sPath & " is already a UNC path" & vbCrLf
            s = sPath
        Else
'd = d & "convert " & sPath & " to UNC" & vbCrLf
            s = Mapped2UNC(sPath)
        End If
        
'd = d & sPath & " converted to UNC = " & s & vbCrLf


        If UCase(MachineName()) = UCase(Parse(s, 3, "\")) Then
'd = d & "this is a local UNC path so convert to local" & vbCrLf
'MsgBox d, vbInformation, "GetLocalizedPath"
            s = UNC2Local(s)
        Else
'd = d & "this is NOT a local UNC path so return original parameter" & vbCrLf
            s = sPath
        End If
        
'd = d & "Localized path = " & s & vbCrLf
'MsgBox d, vbInformation, "GetLocalizedPath"
        GetLocalizedPath = s

    End If

End Function


Public Function UNC2Local(sUNCPath As String) As String
'Dim d As String
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
    
'd = d & "converting unc path to local path" & vbCrLf

    sTemp = Mid(sUNCPath, 3)
    sServer = Parse(sUNCPath, 3, "\")
    sTemp = Mid(sTemp, InStr(1, sTemp, "\") + 1)
    If InStr(1, sTemp, "\") > 0 Then
        sShare = left(sTemp, InStr(1, sTemp, "\") - 1)
        sTemp = Mid(sTemp, InStr(1, sTemp, "\") + 1)
    Else
        sShare = sTemp
        sTemp = ""
    End If


'd = d & "    unc path = " & sUNCPath & vbCrLf
'd = d & "    server = " & sServer & vbCrLf
'd = d & "    share = " & sShare & vbCrLf
'd = d & "    path = " & sTemp & vbCrLf


    baServer = "\\" & sServer & Chr(0)
    baShare = UCase(sShare) & Chr(0)

    result = NetShareGetInfo(baServer(0), baShare(0), 2, Buf)

    If result = 0 Then
        result = PtrToInt(TempStr.XLo, Buf + 24, 2)
        result = PtrToInt(TempStr.XHi, Buf + 26, 2)
        LSet TempPtr = TempStr
        result = PtrToStr(STRArray(0), TempPtr.X)
        sBasePath = left(STRArray, StrLen(TempPtr.X))
        result = NetAPIBufferFree(Buf)
    
'd = d & "NetShareGetInfo SUCCEEDED" & vbCrLf
'd = d & "    base path = " & sBasePath & vbCrLf
'd = d & "Succeesfully converted to local" & vbCrLf
'd = d & "    result = " & PathAppend(sBasePath, sTemp) & vbCrLf
        UNC2Local = PathAppend(sBasePath, sTemp)
    Else
'd = d & "NetShareGetInfo FAILED. result=" & result & vbCrLf
    End If
    
'MsgBox d, vbInformation, "UNC2Local"
End Function






Public Function Mapped2UNC(sMappedPath As String) As String

   Dim sLocalRoot As String
   Dim sRemoteName As String
   Dim sRemotePath As String
   Dim cbRemoteName As Long

   sRemoteName = Space$(260)
   cbRemoteName = Len(sRemoteName)

   sLocalRoot = left(sMappedPath, 2)

   sRemotePath = StripRootFromPath(sMappedPath)

   If IsPathNetPath(sLocalRoot) Then
      If WNetGetConnection(sLocalRoot, sRemoteName, cbRemoteName) = ERROR_SUCCESS Then
         sRemoteName = QualifyPath(TrimNull(sRemoteName)) & sRemotePath
         If IsUNCPathValid(sRemoteName) Then
            Mapped2UNC = sRemoteName
         Else
            
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
   Dim Pos As Integer

   Call PathStripToRoot(sPath)

   Pos = InStr(sPath, Chr$(0))
   If Pos Then
      StripPathToRoot = left$(sPath, Pos - 2)
   Else
      StripPathToRoot = sPath
   End If

End Function
Private Function TrimNull(startstr As String) As String
   TrimNull = left$(startstr, lstrlenW(StrPtr(startstr)))
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
    NetworkLoginName = left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
     
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
    CreateGUID = left$(b, lR - 1)
End Function




Public Function TextboxCanUndo(Textbox) As Boolean
    Dim rc As Long
    rc = SendMessage(Textbox.hwnd, EM_CANUNDO, 0, 0)
    TextboxCanUndo = rc <> 0
End Function
Public Sub TextboxUndo(Textbox)
    Dim rc As Long
    rc = SendMessage(Textbox.hwnd, EM_UNDO, 0, 0)
End Sub


Public Function TempPath() As String
    Dim rc As Long
    Dim bufferLen As Long
    Dim bufferStr As String
        
    bufferLen = 255
    bufferStr = Space(bufferLen)
    rc = GetTempPath(bufferLen, bufferStr)
    If rc > 0 Then
        bufferStr = Trim(bufferStr)
        bufferStr = left(bufferStr, Len(bufferStr) - 1)
    End If
    TempPath = bufferStr

End Function

Public Function TempFile(Optional Extension As String) As String
'    Dim rc As Long
'    Dim bufferStr As String
'    Dim Path As String
'    Path = TempPath
'    bufferStr = Space(255)
'    rc = GetTempFileName("", "", 0, bufferStr)
'    If rc > 0 Then
'        bufferStr = Trim(bufferStr)
'        Path = PathAppend(Path, left(bufferStr, InStr(1, bufferStr, Chr(0)) - 1))
'    End If
'    If Extension <> "" Then
'        Path = ForceExt(Path, Extension)
'    End If
'    TempFile = Path
'
'   THIS STOPPED WORKING AT A CLIENT SITE SO I REPLACED IT WITH A GUID.
'   NO EXPLANATION WHY IT FAILS.
'
    Dim Filename As String
    Filename = CreateGUID
    Filename = Replace(Filename, "{", "")
    Filename = Replace(Filename, "}", "")
    If Extension <> "" Then Filename = ForceExt(Filename, Extension)
    
    TempFile = PathAppend(TempPath, Filename)



End Function

Public Function ShellFile(hwnd As Long, Filename As String, Optional DlgTitle, Optional FindItYourself As Boolean = True, Optional PrintIt As Boolean = False) As String
    'returns new FileName if they found the file they wanted.
    Dim bFailed As Boolean
    Dim bOkNow  As Boolean
    Dim sNewFileName As String
    Dim rc  As Long
     Dim Msg As String
    Const Caption = "Error Launching File"
    
    sNewFileName = Filename
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
            Case SE_ERR_ACCESSDENIED:     Msg = "Access denied":                                           VBA.MsgBox Msg, vbCritical, Caption
            Case SE_ERR_OOM:              Msg = "Out of memory":                                           VBA.MsgBox Msg, vbCritical, Caption
            Case SE_ERR_DLLNOTFOUND:      Msg = "A required DLL could not be found":                       VBA.MsgBox Msg, vbCritical, Caption
            Case SE_ERR_SHARE:            Msg = "A sharing violation has occurred":                        VBA.MsgBox Msg, vbCritical, Caption
            Case SE_ERR_ASSOCINCOMPLETE:  Msg = "Incomplete or invalid file association":                  VBA.MsgBox Msg, vbCritical, Caption
            Case SE_ERR_DDETIMEOUT:       Msg = "DDE Time out":                                            VBA.MsgBox Msg, vbCritical, Caption
            Case SE_ERR_DDEFAIL:          Msg = "DDE transaction failed":                                  VBA.MsgBox Msg, vbCritical, Caption
            Case SE_ERR_DDEBUSY:          Msg = "DDE busy":                                                VBA.MsgBox Msg, vbCritical, Caption
            Case SE_ERR_NOASSOC:          Msg = "No association for file extension":                       VBA.MsgBox Msg, vbCritical, Caption
            Case ERROR_BAD_FORMAT:        Msg = "Invalid EXE file or error in EXE image":                  VBA.MsgBox Msg, vbCritical, Caption
            Case Else:                    Msg = "Unknown error":                                           VBA.MsgBox Msg, vbCritical, Caption
        End Select
    Else
        bOkNow = True
    End If
    
    If bFailed And bOkNow Then ShellFile = sNewFileName

End Function

Public Function ClipboardSetFiles(FilesList As String) As Boolean
   'FilesList is a Chr(0) separated list of filepaths
   
   Dim df As DROPFILES
   Dim hGlobal As Long
   Dim lpGlobal As Long
   Dim i As Long
   
   ' Open AND clear clipboard.
   If OpenClipboard(0&) Then
      Call EmptyClipboard
      
      ' Allocate AND get pointer to global memory,
      ' then copy file list to it.
      hGlobal = GlobalAlloc(GHND, Len(df) + Len(FilesList))
      If hGlobal Then
         lpGlobal = GlobalLock(hGlobal)
         
         ' Build DROPFILES structure in global memory.
         df.pFiles = Len(df)
         Call CopyMem(ByVal lpGlobal, df, Len(df))
         Call CopyMem(ByVal (lpGlobal + Len(df)), ByVal FilesList, Len(FilesList))
         Call GlobalUnlock(hGlobal)
         
         ' Copy fileslist to clipboard, AND return success.
         If SetClipboardData(CF_HDROP, hGlobal) Then
            ClipboardSetFiles = True
         End If
      End If
      
      ' Clean up
      Call CloseClipboard
   End If
   
End Function

Public Function ClipboardGetFiles() As String
'returns a Chr(0) separated list of filepaths

    Dim hDrop As Long
    Dim nFiles As Long
    Dim i As Long
    Dim Filename As String
    Dim namelen As Long
    Dim returnstr As String
    Const MAX_PATH As Long = 260
    
    ' Insure desired format is there, AND open clipboard.
    If IsClipboardFormatAvailable(CF_HDROP) Then
        If OpenClipboard(0&) Then
            
            ' Get handle to Dropped Filelist data, AND number of files.
            hDrop = GetClipboardData(CF_HDROP)
            nFiles = DragQueryFile(hDrop, -1&, "", 0)
            
            ' Retrieve each FileName in Dropped Filelist.
            returnstr = ""
            For i = 0 To nFiles - 1
                Filename = String(MAX_PATH, " ")
                namelen = DragQueryFile(hDrop, i, Filename, Len(Filename))
'                returnstr = returnstr & Chr(0) & left(FileName, namelen)
                returnstr = returnstr & ";" & left(Filename, namelen)
            Next
            returnstr = Mid(returnstr, 2)
            
            ' Clean up
            Call CloseClipboard
        End If
        
        ' Assign return Value equal to number of files dropped.
        ClipboardGetFiles = returnstr
    End If
End Function

' ********************************************
'  Public Methods
' ********************************************
Public Function RegDeleteKey(ByVal hive As RegistryHiveConstants, ByVal section As String, Optional ByVal Key As String = "") As Boolean
   ' Section   Required. String expression containing the name of the section WHERE the key setting
   '           is being deleted. If only section is provided, the specified section is deleted along
   '           with all related key settings.
   ' Key       Optional. String expression containing the name of the key setting being deleted.
   Dim nRet As Long
   Dim hKey As Long

   If Len(Key) Then
      ' Open key
      nRet = RegOpenKeyEx(hive, section, 0&, KEY_ALL_ACCESS, hKey)
      If nRet = ERROR_SUCCESS Then
         ' Set appropriate Value for default query
         If Key = "*" Then Key = vbNullString
         ' Delete the requested Value
         nRet = RegDeleteValue(hKey, Key)
         Call RegCloseKey(hKey)
      End If
   Else
      ' Open parent key
      nRet = RegOpenKeyEx(hive, section, 0&, KEY_ALL_ACCESS, hKey)
      If nRet = ERROR_SUCCESS Then
         ' Attempt to delete whole section
         nRet = RegRegDeleteKey(hKey, section)
         Call RegCloseKey(hKey)
      End If
   End If
   RegDeleteKey = (nRet = ERROR_SUCCESS)
End Function

Public Function RegGetKey(ByVal hive As RegistryHiveConstants, ByVal section As String, Optional ByVal Key As String, Optional ByVal Default As String = "") As String
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
      If Key = "*" Then Key = vbNullString
      
      ' Determine how large the buffer needs to be
      nRet = RegQueryValueEx(hKey, Key, 0&, nType, ByVal Buffer, nBytes)
      If nRet = ERROR_SUCCESS Then
         ' Build buffer AND get data
         If nBytes > 0 Then
            Buffer = Space(nBytes)
            nRet = RegQueryValueEx(hKey, Key, 0&, nType, ByVal Buffer, Len(Buffer))
            If nRet = ERROR_SUCCESS Then
               ' Trim NULL AND return successful query!
               RegGetKey = left(Buffer, nBytes - 1)
            End If
         End If
      Call RegCloseKey(hKey)
      End If
   End If
End Function

Public Function RegSaveKey(ByVal hive As RegistryHiveConstants, ByVal section As String, Optional ByVal Key As String, Optional ByVal Value As String) As Boolean
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
      If Key = "" Then Key = vbNullString
      ' Write new Value to registry
      nRet = RegSetValueEx(hKey, Key, 0&, REG_SZ, ByVal Value, Len(Value))
      Call RegCloseKey(hKey)
   End If
   RegSaveKey = (nRet = ERROR_SUCCESS)
End Function

Public Function RegEnumSections(ByVal hive As RegistryHiveConstants, ByVal section As String, SubSections() As String) As Boolean
    Dim iSectCount As Long
    Dim lResult As Long
    Dim hKey As Long
    Dim dwReserved As Long
    Dim szBuffer As String
    Dim lBuffSize As Long
    Dim lIndex As Long
    Dim lType As Long
    Dim sCompKey As String
    Dim iPos As Long

    iSectCount = 0
    ReDim SubSections(0)

    lIndex = 0

    lResult = RegOpenKeyEx(hive, section, 0&, KEY_ENUMERATE_SUB_KEYS, hKey)
    Do While lResult = ERROR_SUCCESS
        'Set buffer space
        szBuffer = String$(255, 0)
        lBuffSize = Len(szBuffer)
        
        'Get next Value
        lResult = RegEnumKey(hKey, lIndex, szBuffer, lBuffSize)
                              
        If (lResult = ERROR_SUCCESS) Then
            iSectCount = iSectCount + 1
            ReDim Preserve SubSections(iSectCount) As String
            iPos = InStr(szBuffer, Chr$(0))
            If (iPos > 0) Then
                SubSections(iSectCount) = left(szBuffer, iPos - 1)
            Else
                SubSections(iSectCount) = left(szBuffer, lBuffSize)
            End If
        End If
        
        lIndex = lIndex + 1
    Loop
    If (hKey <> 0) Then
        RegCloseKey hKey
    End If
    RegEnumSections = True
End Function


Public Function RegEnumKeys(ByVal hive As RegistryHiveConstants, ByVal section As String, ByRef keys() As String, Optional WOW64 As Boolean = False) As Boolean
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
        
    If WOW64 Then
        lResult = RegOpenKeyEx(hive, section, 0, KEY_QUERY_VALUE + REG_WOW64, hKey)
    Else
        lResult = RegOpenKeyEx(hive, section, 0, KEY_QUERY_VALUE, hKey)
    End If
    
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
                sName = left$(sName, lNameSize)
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

Public Function IniGet(Filename As String, SectionName As String, VariableName As String, Optional DefaultValue As Variant) As Variant
    
    'retrieves a Value FROM an ini file
    
    Dim VariableValue As String
    Dim lnumChar As Long
    
    VariableValue = String(128, 0)
    
    'DefaultValues to ""
    lnumChar = GetPrivateProfileString(SectionName, VariableName, "", VariableValue, 127, Filename)

    If lnumChar >= 1 Then 'Null appended to string; even null string
        Select Case TypeName(DefaultValue)
            Case "Byte"
                IniGet = CByte(left$(VariableValue, lnumChar))
            Case "Integer", "Long", "double", "double"
                IniGet = Val(left$(VariableValue, lnumChar))
            Case "Date"
                IniGet = CDate(left$(VariableValue, lnumChar))
            Case "Boolean"
                IniGet = CBool(left$(VariableValue, lnumChar))
            Case Else 'DefaultValue to string
                IniGet = left$(VariableValue, lnumChar)
        End Select
    
    Else 'DefaultValue applicable
        If IsMissing(DefaultValue) Then
            IniGet = ""
        Else
            IniGet = DefaultValue
        End If
    End If
    
End Function

Public Function IniGetSectionNames(Filename As String) As String

    'returns a csv list of the section names in file

    Dim rc     As Long
    Dim sBuff  As String

    sBuff = Space$(BUFFER_SIZE)
    rc = GetPrivateProfileSectionNames(sBuff, BUFFER_SIZE, Filename)
    If rc > 0 Then
        sBuff = left(sBuff, rc - 1)
        sBuff = Replace(sBuff, Chr$(0), ",")
        If Right(sBuff, 1) = "," Then
            sBuff = left(sBuff, Len(sBuff) - 1)
        End If
        IniGetSectionNames = sBuff
    End If
End Function

Public Function IniGetVariableNames(Filename As String, SectionName As String) As String
    
    'returns a csv list of the variable names in a section
    
    Dim sBuff As String
    Dim rc As Long
    sBuff = Space$(BUFFER_SIZE)
    rc = GetPrivateProfileString(SectionName, CLng(0), "", sBuff, BUFFER_SIZE, Filename)
    If rc > 0 Then
        sBuff = left(sBuff, rc - 1)
        If Right(sBuff, 1) = "," Then
            sBuff = left(sBuff, Len(sBuff) - 1)
        End If
        sBuff = Replace(sBuff, Chr$(0), ",")
    Else
        sBuff = ""
    End If
    IniGetVariableNames = sBuff
    
End Function

Public Sub IniPut(Filename As String, SectionName As String, VariableName As String, VariableValue As Variant)

    Const DFORMAT = "hh:mm mmmm d yyyy"

    
    Call CreatePath("", FilePath(Filename))
    
    'writes a Value to an ini file

    Dim tmpValue  As String
    Dim numChar   As Long
    
    Select Case TypeName(VariableValue)
        Case "Byte", "Integer", "Long", "double", "double"
            tmpValue = Format$(VariableValue)
            
        Case "Date"
            tmpValue = Format$(VariableValue, DFORMAT)
            
        Case "Boolean"
            If VariableValue Then
                tmpValue = "True"
            Else
                tmpValue = "False"
            End If
            
        Case Else 'DefaultValue to string
            tmpValue = VariableValue
            
    End Select
    
    numChar = WritePrivateProfileString(SectionName, VariableName, tmpValue, Filename)
End Sub

Public Sub IniRemove(Filename As String, SectionName As String, Optional VariableName As String)
    
    'removes a variable or a whole section from an ini file
    
    If VariableName = "" Then
        'remove section
        Call WritePrivateProfileString(SectionName, 0&, "", Filename)
    Else
        'remove variable
        Call WritePrivateProfileString(SectionName, VariableName, 0&, Filename)
    End If

End Sub

Public Function Update(ByVal InputString As String, ByVal Col As Long, ByVal Value As String, Optional Delimeter As String = ",") As String

    'updates items in a delimited string
    
    Dim Pos        As Integer
    Dim workString As String
    Dim idx        As Integer

    For idx = 1 To Col - 1
        Pos = InStr(Pos + 1, InputString, Delimeter)
        If Pos = 0 Then
            ' column not found, add one
            InputString = InputString & Delimeter
            Pos = Len(InputString)
        End If
    Next idx

    workString = left(InputString, Pos)
    Pos = InStr(Pos + 1, InputString, Delimeter)

    If Pos > 0 Then 'not last column
        workString = workString & Value & Mid$(InputString, Pos)
    Else
        workString = workString & Value
    End If

    Update = workString

End Function

Public Function FormatPhone(Number As String, Optional countrycode As String = "") As String
'    Dim countrycode As String 'AU,CA,US
    Dim isMobile As Boolean
    Dim isInternational As Boolean
    
    If countrycode = "" Then countrycode = Trim(HFApp.Options(Country))
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
        isInternational = left(Number, 1) = "+"
        If isInternational Then
            isMobile = IsIn(Mid(Number, 4, 1), "4", "5")
        Else
            isMobile = IsIn(left(Number, 2), "04", "05")
        End If
        Select Case True
            Case Not IsNumeric(Number):          FormatPhone = Number
            Case isMobile And isInternational:   FormatPhone = Format(Number, "!@@@ @@@ @@@ @@@")
            Case isMobile:                       FormatPhone = Format(Number, "!@@@@ @@@ @@@")
            Case isInternational:                FormatPhone = Format(Number, "!@@@ @ @@@@ @@@@")
            Case Else:                           FormatPhone = Format(Number, "!@@ @@@@ @@@@")
        End Select
    
    Case Else
        Select Case True
            Case Not IsNumeric(Number):   FormatPhone = Number
            Case Len(Number) = 10:        FormatPhone = Format(Number, "!(@@@) @@@-@@@@")
            Case Len(Number) = 11:        FormatPhone = Format(Number, "!@ (@@@) @@@-@@@@")
            Case Len(Number) = 7:         FormatPhone = Format(Number, "!@@@-@@@@")
            Case Else:                    FormatPhone = Number
        End Select
    End Select

End Function
Public Function ItemInList(item As String, list As String, Optional ByVal Delimeter As String = ",") As Boolean
    ItemInList = InStr(1, Delimeter & list & Delimeter, Delimeter & item & Delimeter)
End Function

Public Function tokenize(s As String, Delimiters As String)
    'parses a string returns an array of strings broken at each
    'ie: tokenize("ab:c;def", ":;")
    '    will return a zero based array of 5 element -- ab,:,c,;,def
    
    Dim a() As String
    Dim i As Long
    
    ReDim a(0)
    For i = 1 To Len(s)
        If InStr(1, Delimiters, Mid(s, i, 1), vbTextCompare) Then
            ReDim Preserve a(UBound(a) + 1)
            a(UBound(a)) = Mid(s, i, 1)
            ReDim Preserve a(UBound(a) + 1)
        Else
            a(UBound(a)) = a(UBound(a)) & Mid(s, i, 1)
        End If
    Next
    
    tokenize = a
    
End Function
Public Function Parse(DelimitedString, Optional Index As Long, Optional ByVal Delimeter As String = ",") As Variant
    
    'Parses items out of a delimited string
    'Returns the number of items if index is omitted
    
    Dim tmpCount As Long
    Dim Pos      As Long
    Dim i As Long
    Dim fStart As Long
    Dim fend As Long
    
    If Index < 0 Then
        Parse = ""
    ElseIf Index = 0 Then
        'return count
        tmpCount = 1
        Pos = 0
        While InStr(Pos + 1, DelimitedString, Delimeter, vbTextCompare)
            Pos = InStr(Pos + 1, DelimitedString, Delimeter, vbTextCompare)
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
Public Function IsBetween(Value, Minimum, Maximum, Optional Inclusive = True) As Boolean
    If Inclusive Then
        IsBetween = Value >= Minimum And Value <= Maximum
    Else
        IsBetween = Value > Minimum And Value < Maximum
    End If
End Function
Public Function Between(Minimum, Maximum, Value) As Boolean
    Value = IIf(Value > Maximum, Maximum, Value)
    Value = IIf(Value < Minimum, Minimum, Value)
    Between = Value
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

Public Function ToolbarWidth(tb) As Long
    Dim i As Integer
    Dim w As Long
    For i = 1 To tb.Buttons.Count
        w = w + tb.Buttons(i).Width
    Next
    ToolbarWidth = w
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
    Dim h As Long
    For i = 0 To g.Rows - 1
        h = h + g.RowHeight(i)
    Next
    h = h + (g.Rows - 1) * g.GridLineWidth * 15
    GridHeight = h
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

Public Sub GridEnterKeyPress(g, Optional Shift As Integer)
    Dim i As Long
    Dim c As Long
    Dim r As Long
    
    r = g.Row
    c = g.Col
    
    
    If Shift = vbCtrlMask Then
        'next row first col
        If r < g.Rows - 1 Then
            c = GridNextVisibleColumn(g, 0)
            r = r + 1
        End If
    Else
        'next col
        If c = GridLastVisibleColumn(g) Then
            If r < g.Rows - 1 Then
                r = r + 1
                c = -1
            End If
        End If
        For i = c + 1 To g.Cols - 1
            If i >= g.FixedCols And Not g.ColHidden(i) Then
                c = i
                Exit For
            End If
        Next
    End If
    Call g.Select(r, c)
    Call g.ShowCell(r, c)
    
End Sub

Public Function IsFormLoaded(FormName As String) As Boolean
    Dim i As Long
    For i = 0 To Forms.Count - 1
        If Forms(i).Name = FormName Then
            IsFormLoaded = True
            Exit Function
        End If
    Next
End Function


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



Public Function Decode(Expression As String, ParamArray ValueReturn())
    Dim i As Integer
    For i = 0 To UBound(ValueReturn) Step 2
        If Expression = ValueReturn(i) Then
            Decode = ValueReturn(i + 1)
            Exit Function
        End If
    Next
    If (1 + UBound(ValueReturn)) Mod 2 = 0 Then
        Decode = ""
    Else
        Decode = ValueReturn(UBound(ValueReturn))
    End If
End Function

Public Sub SetToolbarIcons(Toolbar, il)
On Error Resume Next
    Dim i As Integer
    Set Toolbar.ImageList = il
    For i = 1 To Toolbar.Buttons.Count
        Toolbar.Buttons(i).Image = Toolbar.Buttons(i).Key
    Next
End Sub

Public Sub ShowToolbarCaptions(Toolbar, b)
    Dim i As Integer
    If b Then
        For i = 1 To Toolbar.Buttons.Count
            If Toolbar.Buttons(i).Caption = "" Then
                Toolbar.Buttons(i).Caption = Toolbar.Buttons(i).tag
                Toolbar.Buttons(i).ToolTipText = ""
            End If
        Next
    Else
        For i = 1 To Toolbar.Buttons.Count
            If Toolbar.Buttons(i).Caption <> "" Then
                Toolbar.Buttons(i).ToolTipText = Toolbar.Buttons(i).Caption
                Toolbar.Buttons(i).tag = Toolbar.Buttons(i).Caption
                Toolbar.Buttons(i).Caption = ""
            End If
        Next
    End If
    Toolbar.tag = Update(Toolbar.tag, 2, b)
    Toolbar.Refresh
On Error Resume Next
    Toolbar.Parent.Refresh
End Sub


Public Function trunc(Number As Double, precision As Integer) As String
    trunc = Format(Number, "0." & String(precision, "0"))
End Function

Public Function InIde() As Boolean
On Error Resume Next
    Debug.Print (1 / 0)
    InIde = (Err.Number <> 0)
    On Error GoTo 0
    Exit Function
End Function

Public Function SimpleEncrypt(ByVal Secret As String, Key As String) As String
    Dim i As Long
    For i = 1 To Len(Secret)
        Mid$(Secret, i, 1) = Chr$(Asc(Mid$(Secret, i, 1)) Xor Asc(Mid$(Key, (i Mod Len(Key)) - Len(Key) * ((i Mod Len(Key)) = 0), 1)))
    Next
    SimpleEncrypt = Secret
End Function

Public Function SetListIndex(Combobox, Optional ItemData, Optional Text As String) As Boolean
    Dim i As Long
    
    If IsMissing(ItemData) And Text = "" Then
        Combobox.Text = ""
        SetListIndex = True
    ElseIf Text <> "" Then
        For i = 0 To Combobox.ListCount - 1
            If Combobox.list(i) = Text Then
                Combobox.ListIndex = i
                SetListIndex = True
                Exit Function
            End If
        Next
    Else
        On Error Resume Next
        Combobox.Text = ""
        SetListIndex = True
            
        For i = 0 To Combobox.ListCount - 1
            If Combobox.ItemData(i) = ItemData Then
                Combobox.ListIndex = i
                SetListIndex = True
                Exit Function
            End If
        Next
    End If
    
    SetListIndex = False
    
End Function


Public Function csv(ParamArray Values()) As String
    Dim i As Integer
    Dim s As String
    For i = LBound(Values) To UBound(Values)
        s = s & "," & Values(i)
    Next
    csv = Mid(s, 2)
End Function

Public Function DList(Delimiter As String, ParamArray Values()) As String
    Dim i As Integer
    Dim s As String
    For i = LBound(Values) To UBound(Values)
        s = s & Delimiter & Values(i)
    Next
    DList = Mid(s, Len(Delimiter) + 1)
End Function

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

Public Function ShowForm(FormName As String, TagValue) As Boolean
    Dim f As Form
    For Each f In Forms
        If f.Name = FormName And f.tag = TagValue Then
            ShowForm = True
            If f.WindowState = vbMinimized Then f.WindowState = vbNormal
            f.SetFocus
            Exit Function
        End If
    Next
End Function

Public Function ListViewSelected(lv) As String
    Dim i As Long
    Dim s As String
    For i = 1 To lv.ListItems.Count
        If lv.ListItems(i).Selected Then s = s & "," & lv.ListItems(i).Key
    Next
    ListViewSelected = Mid(s, 2)
End Function

Public Function ImgComboFindItem(cbo) As Boolean
    Dim i As Long
    For i = 1 To cbo.ComboItems.Count
        If cbo.Text = cbo.ComboItems(i).Text Then
            cbo.ComboItems(i).Selected = True
            ImgComboFindItem = False
            Exit Function
        End If
    Next
    cbo.ComboItems(1).Selected = True
    ImgComboFindItem = False
End Function

Public Sub CenterForm(f As Form, Optional p As Form)
    If p Is Nothing Then
        f.Move (Screen.Width - f.Width) / 2, (Screen.Height - f.Height) / 2
    Else
        If f.MDIChild Then
            f.Move (p.ScaleWidth - f.Width) / 2, (p.ScaleHeight - f.Height) / 2
        Else
            f.Move p.left + ((p.Width - f.Width) / 2), p.Top + ((p.Height - f.Height) / 2)
        End If
    End If
End Sub

Public Function GetHotKey(s As String) As String
    Dim p As Integer
    
    p = InStr(1, s, "&")
    GetHotKey = Mid(s, p + 1, 1)
    
End Function

Public Function Random(lower, upper)
    Random = Int((upper - lower + 1) * Rnd + lower)
End Function


Public Function FileDrive(fullpath As String) As String
    fullpath = Trim(fullpath)
    If left(fullpath, 2) = "\\" Then
        'UNC names
        FileDrive = "\\" & Parse(Mid(fullpath, 3), 1, "\")
    Else
        FileDrive = left(fullpath, 2)
    End If
End Function

Public Function FileTitle(fullpath As String) As String
    FileTitle = Parse(fullpath, Parse(fullpath, , "\"), "\")
End Function

Public Function CleanFileName(Filename As String) As String
'strip out reserved characters
    Filename = Replace(Filename, "/", "")
    Filename = Replace(Filename, "\", "")
    Filename = Replace(Filename, "*", "")
    Filename = Replace(Filename, "?", "")
    Filename = Replace(Filename, ":", "")
    Filename = Replace(Filename, "<", "")
    Filename = Replace(Filename, ">", "")
    Filename = Replace(Filename, "|", "")
    CleanFileName = Filename
End Function
Public Function Filename(fullpath As String) As String
    Filename = StripExtension(FileTitle(fullpath))
End Function
Public Function FilePath(fullpath As String) As String
    FilePath = left(fullpath, Len(fullpath) - Len(FileTitle(fullpath)))
End Function
Public Function FileExt(fullpath As String) As String
    If InStr(fullpath, ".") Then
        FileExt = LCase(Parse(fullpath, Parse(fullpath, , "."), "."))
    Else
        FileExt = ""
    End If
End Function
Public Function ForceExt(Filename As String, Extension As String) As String
    If Extension = "" Then
        ForceExt = Filename
    Else
        ForceExt = StripExtension(Filename) & "." & Extension
    End If
End Function

Public Function StripExtension(Filename As String) As String
    Dim i As Integer
    Dim s As String
    Dim b As Boolean
    For i = Len(Filename) To 1 Step -1
        If Mid(Filename, i, 1) = "." Then b = True
        If b Then s = Mid(Filename, i, 1) & s
    Next
    If b Then
        StripExtension = Mid(s, 1, Len(s) - 1)
    Else
        StripExtension = Filename
    End If
End Function

Public Function CompactedPathSh(ByVal sPath As String, ByVal lMaxWidthPixels As Long, Optional hDC As Long) As String
    Dim lR As Long
    Dim iPos As Long
    
    If hDC = 0 Then hDC = Screen.ActiveForm.hDC
    lR = PathCompactPath(hDC, sPath, lMaxWidthPixels)
    iPos = InStr(sPath, Chr$(0))
    If iPos <> 0 Then
        CompactedPathSh = left$(sPath, iPos - 1)
    Else
        CompactedPathSh = sPath
    End If
End Function








Public Function TDGMouseCol(Grid, X As Single, split As Long) As Long
    Dim i As Long
    With Grid
        split = 0
        If .Splits.Count > 1 Then
            If X > .Splits(1).Columns(.Splits(1).LeftCol).left Then split = 1
        End If
        For i = .Splits(split).Columns.Count - 1 To 0 Step -1
            If .Splits(split).Columns(i).left <> 0 And .Splits(split).Columns(i).left < X Then
                TDGMouseCol = i
                Exit Function
            End If
        Next
    End With
    TDGMouseCol = -1
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

Public Function RepeatingString(Length As Long, Pattern As String) As String
    Dim s As String
    
    If Length < 0 Or Len(Pattern) < 2 Then
        If Length < 0 Then Length = 0
        If Len(Pattern) < 1 Then Pattern = " "
        RepeatingString = String(Length, Pattern)
    Else
        While Len(s) < Length
            s = s & Pattern
        Wend
        RepeatingString = left(s, Length)
    End If
    
End Function

Public Sub AddToComboList(Combobox As Combobox)
    Dim i As Long
    If Combobox.Text = "" Then Exit Sub
    For i = 0 To Combobox.ListCount - 1
        If Combobox.list(i) = Combobox.Text Then Exit Sub
    Next
    Combobox.AddItem Combobox.Text, 0
End Sub


Public Function ImageIndex(ImageList, Key As String) As Long
On Error Resume Next
    ImageIndex = ImageList.ListImages(Key).Index - 1
End Function

Public Function DNull(s As String) As String
    DNull = Mid(s, 1, Max(0, InStr(1, s, vbNullChar) - 1))
End Function

Public Function PathIsLocalPath(ByVal sPath As String) As Boolean
   PathIsLocalPath = PathIsNetworkPath(sPath) = False Or sPath Like "\\tsclient*"
End Function

Public Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0
End Function


Public Function FileExists(Filename As String) As Boolean
On Error Resume Next
    FileExists = VBA.Dir(Filename) <> ""
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


Public Function SelectedOption(OptionButtons) As Long
    Dim i As Long
    For i = OptionButtons.LBound To OptionButtons.UBound
        If OptionButtons(i) Then
            SelectedOption = i
            Exit Function
        End If
    Next
End Function

Public Sub WindowOnTop(f As Form, OnTop As Boolean)
    Call SetWindowPos(f.hwnd, IIf(OnTop, HWND_TopMOST, HWND_NOTopMOST), 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
End Sub

    

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
Public Function DeQuote(s As String, Optional SingleQuotes As Boolean = False) As String
    
    s = Trim(s)
    If SingleQuotes Then
        If left(s, 1) = vbSingleQuote And Right(s, 1) = vbSingleQuote Then
            s = Mid(s, 2, Len(s) - 2)
        End If
    Else
        If left(s, 1) = vbQuote And Right(s, 1) = vbQuote Then
            s = Mid(s, 2, Len(s) - 2)
        End If
    End If
    DeQuote = s


End Function

Public Function Quote(s As String, Optional RemoveNewLines As Boolean, Optional Length As Long) As String
    If RemoveNewLines Then
        s = Replace(s, vbLf, " ")
        s = Replace(s, vbCr, " ")
    End If
    If Length <> 0 Then
        s = left(s, Length)
    End If
    Quote = vbQuote & Replace(s, vbQuote, vbQuote & vbQuote) & vbQuote
End Function

Public Sub NumberBoxKeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 8, 46
            '0 thru 9, backspace and decimal OK
        Case Else
            'cancel anything else
            KeyAscii = 0
    End Select
End Sub
Public Function ListItem(Combobox, Optional Text, Optional ItemData, Optional Compare As VbCompareMethod = vbTextCompare) As Long
    Dim i As Long
    If Not IsMissing(Text) Then
        For i = 0 To Combobox.ListCount - 1
            If StrComp(Text, Combobox.list(i), Compare) = 0 Then
                ListItem = i
                Exit Function
            End If
        Next
    End If
    If Not IsMissing(ItemData) Then
        For i = 0 To Combobox.ListCount - 1
            If ItemData = Combobox.ItemData(i) Then
                ListItem = i
                Exit Function
            End If
        Next
    End If
    ListItem = -1
End Function





Public Function Ole2RgbColor(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long
    ' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, Ole2RgbColor) Then
        Ole2RgbColor = CLR_INVALID
    End If
End Function


Public Function PadString(ByVal s As String, ByVal SIZE As Long, Optional ByVal Alignment As VBRUN.AlignmentConstants = vbLeftJustify) As String

    If SIZE = 0 Then Exit Function
    
    If Len(s) > SIZE Then
        PadString = Mid(s, 1, SIZE)
        Exit Function
    End If
    
    Select Case Alignment
        Case vbLeftJustify
            PadString = s & String(SIZE - Len(s), " ")
        
        Case vbRightJustify
            PadString = String(SIZE - Len(s), " ") & s
        
        Case vbCenter
            s = String((SIZE - Len(s)) \ 2, " ") & s & String((SIZE - Len(s)) \ 2, " ")
            If SIZE > Len(s) Then
                PadString = s & String(SIZE - Len(s), " ")
            Else
                PadString = Mid(s, 1, SIZE)
            End If
        
    End Select
    
        
End Function
Public Sub IniGetGrid(Form As Form, Grid As Object, Optional IniFile As String, Optional StaticPositions As Boolean = False, Optional InstanceKey, Optional Grouped As Boolean, Optional StaticHeadings As Boolean = False)
    If DBGetGrid(Form, Grid, StaticPositions, InstanceKey, Grouped, StaticHeadings) Then Exit Sub
    
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
                    ColIndex = Grid.Splits(splitindex).Columns(variable).ColIndex
                    If ColIndex <> -1 Then
                        Grid.Splits(splitindex).Columns(ColIndex).Width = Val("" & Parse(colCSV, 1))
                        Grid.Splits(splitindex).Columns(ColIndex).Visible = CBool(Parse(colCSV, 2))
                        If Not StaticPositions Then
                            Grid.Splits(splitindex).Columns(variable).Order = Val("" & Parse(colCSV, 3))
                        End If
                        Grid.Splits(splitindex).Columns(ColIndex).Caption = Parse(colCSV, 4)
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
                        Caption = Parse(colCSV, 4)
                        If Caption <> "" And Not StaticHeadings Then Grid.TextMatrix(0, ColIndex) = Caption

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

Public Function DBGetGrid(Form As Form, Grid As Object, Optional StaticPositions As Boolean = False, Optional InstanceKey, Optional Grouped As Boolean, Optional StaticHeadings As Boolean = False) As Boolean
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
                Col = .Splits(split).Columns("" & rs("colkey")).ColIndex
                
                If Not StaticHeadings Then .Splits(split).Columns(Col).Caption = "" & rs("caption")
                .Splits(split).Columns(Col).Width = Val("" & rs("width"))
                .Splits(split).Columns(Col).Visible = "" & rs("hidden") <> "True"
                If Not StaticPositions Then .Splits(split).Columns(Col).Order = Val("" & rs("colindex"))
    
                rs.MoveNext
            Wend
        
        Else
        
            AllColsHidden = True
            While Not rs.EOF
                Col = .ColIndex("" & rs("colkey"))
                If Not StaticHeadings Then .TextMatrix(0, Col) = "" & rs("caption")
                .ColWidth(Col) = Val("" & rs("width"))
                .ColHidden(Col) = "" & rs("hidden") = "True"
                .ColData(Col) = IIf("" & rs("grouped") = "True", "GROUPED", "")
                AllColsHidden = AllColsHidden And "" & rs("hidden") = "True"
                If Not StaticPositions Then .ColPosition(Col) = Val("" & rs("colindex"))
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
                For Col = 0 To Grid.Columns.Count - 1
                    'uid,gridname,split,colindex,colkey,caption,width,hidden,grouped
                    sql = sql & ",(" & DbQuote(Str, HFApp.LoginID) & _
                                "," & DbQuote(Str, gridname) & _
                                "," & DbQuote(Num, split) & _
                                "," & DbQuote(Num, Grid.Splits(split).Columns(Col).Order) & _
                                "," & DbQuote(Str, Grid.Splits(split).Columns(Col).DataField) & _
                                "," & DbQuote(Str, Grid.Splits(split).Columns(Col).Caption) & _
                                "," & DbQuote(Num, Grid.Splits(split).Columns(Col).Width) & _
                                "," & DbQuote(Bit, Grid.Splits(split).Columns(Col).Visible) & _
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
        Call IniPut(IniFile, FormName, "Left", Form.left)
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
                    bCaptions = Parse(c.tag, 2) = "True"
                    Call IniPut(IniFile, FormName, CtrlName & ".Visible", bVisible)
                    Call IniPut(IniFile, FormName, CtrlName & ".Captions", bCaptions)
                Case "StatusBar", "vbalARListBar"
                    bVisible = c.Visible
                    Call IniPut(IniFile, FormName, CtrlName & ".Visible", bVisible)
                Case "Slider", "Panel"
                    bVisible = c.Visible
                    Call IniPut(IniFile, FormName, CtrlName & ".Visible", bVisible)
                    Call IniPut(IniFile, FormName, CtrlName & ".Top", c.Top)
                    Call IniPut(IniFile, FormName, CtrlName & ".Left", c.left)
            End Select
        Next
    End If
End Sub

Public Sub IniGetForm(Form, Optional IniFile As String, Optional InstanceKey)
    Dim c         As Control
    Dim WinState  As Integer
    Dim FormName  As String
    Dim CtrlName  As String
    Dim bVisible  As Boolean
    Dim bCaptions As Boolean
    Dim Value     As Double
    
    If IniFile = "" Then IniFile = AppIni
        
    FormName = Form.Name
    If Not IsMissing(InstanceKey) Then FormName = FormName & "(" & InstanceKey & ")"
    
    WinState = IniGet(IniFile, FormName, "State", Form.WindowState)
    
    Value = Val("" & IniGet(IniFile, FormName, "Top", Form.Top))
    Value = IIf(Value < 0, 0, Value)
    Value = IIf(Value > Screen.Height - 600, Form.Top, Value)
    Form.Top = Value
    
    Value = Val("" & IniGet(IniFile, FormName, "Left", Form.left))
    Value = IIf(Value + Form.Width < 0, 0, Value)
    Value = IIf(Value > Screen.Width, Form.left, Value)
    Form.left = Value
    
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
            
'            Case "Toolbar"
'                bVisible = IniGet(IniFile, FormName, CtrlName & ".Visible", True)
'                bCaptions = IniGet(IniFile, FormName, CtrlName & ".Captions", True)
'                c.tag = bVisible & "," & bCaptions
'                c.Visible = bVisible
'                Call ShowToolbarCaptions(c, bCaptions)
            
            Case "StatusBar", "vbalARListBar"
                bVisible = IniGet(IniFile, FormName, CtrlName & ".Visible", True)
                c.tag = "" & bVisible
                c.Visible = bVisible
                
            Case "Slider", "Panel"
                bVisible = IniGet(IniFile, FormName, CtrlName & ".Visible", True)
                c.tag = "" & bVisible
                c.Visible = bVisible
                c.Top = IniGet(IniFile, FormName, CtrlName & ".Top", c.Top)
                c.left = IniGet(IniFile, FormName, CtrlName & ".Left", c.left)
                
        End Select
    Next
    
'UNDOME
'    Dim X As LabelEventTrap
'    On Error GoTo exitsub
'    Set Form.EventTraps = New Collection
'    For Each c In Form.Controls
'        If TypeName(c) = "Label" Then
'            Set X = New LabelEventTrap
'            On Error Resume Next
'            Set X.Ctrl = c
'            If Err.Number = 0 Then Form.EventTraps.Add X
'        End If
'    Next
    
exitsub: Exit Sub
End Sub


Public Property Get AppWorkingFolder() As String
    AppWorkingFolder = PathAppend(GetFolderPath(CSIDL_LOCAL_APPDATA), "Homefront\" & App.EXEName)
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

Public Sub SetCtrlFocus(c As Control)
On Error Resume Next
    Call c.SetFocus
End Sub


Public Sub LoadComboBox(Combobox, Connection As adodb.Connection, sql As String)
On Error GoTo eh
    Dim rs  As adodb.Recordset
    Dim tag As String
    Dim ra As Long
    'SELECT ItemDesc  text to be displayed in list
    '      ,KeyValue  alpha key value
    '      ,IDValue   numeric key value
    '  FROM Table
    
    'ID are stored in combobox.itemdata(x)
    'Keys are stored chr(1) separated in the controls tag property
    '     use SetComboBoxListIndex & GetComboBoxListKey, GetComboBoxListID, GetComboBoxListIndex functions
    
    Set rs = Connection.Execute(sql, ra)
    Combobox.Clear
    tag = ""
    While Not rs.EOF
        Combobox.AddItem Trim("" & rs(0))
        Combobox.ItemData(Combobox.NewIndex) = Val("" & rs(2))
        tag = tag & Chr(1) & "" & rs(1)
        rs.MoveNext
    Wend
    Combobox.tag = Mid(tag, 2)
eh: Exit Sub
End Sub
Public Function GetComboBoxListKey(Combobox) As String
On Error Resume Next
If Combobox.ListIndex > -1 Then GetComboBoxListKey = Parse(Combobox.tag, Combobox.ListIndex + 1, Chr(1))
End Function
Public Function GetComboBoxListID(Combobox) As Long
    GetComboBoxListID = Combobox.ItemData(Combobox.ListIndex)
End Function
Public Function GetComboBoxListIndex(Combobox) As Long
    GetComboBoxListIndex = Combobox.ListIndex
End Function
Public Sub SetComboBoxListIndex(Combobox, Optional Text As String, Optional Key As String, Optional ID As Long = -1, Optional Index As Long = -1)
On Error Resume Next
    Dim s As String
    Dim i As Long
    
    If Index <> -1 Then
        Combobox.ListIndex = Index
    End If
    
    
    If Key <> "" Then
        s = Combobox.tag
        For i = 1 To Parse(s, , Chr(1))
            If UCase(Parse(s, i, Chr(1))) = UCase(Key) Then
                Combobox.ListIndex = i - 1
                Exit Sub
            End If
        Next
        Combobox.ListIndex = -1
    End If
    
    If ID <> -1 Then
        For i = 0 To Combobox.ListCount - 1
            If Combobox.ItemData(i) = ID Then
                Combobox.ListIndex = i
                Exit Sub
            End If
        Next
        Combobox.ListIndex = -1
    End If

    
    For i = 0 To Combobox.ListCount - 1
        If UCase(Combobox.list(i)) = UCase(Text) Then
            Combobox.ListIndex = i
            Exit Sub
        End If
    Next
    Combobox.ListIndex = -1
    
End Sub

Public Function SelectPrinter(ByVal DeviceName As String) As Boolean
    Dim i As Integer
    SelectPrinter = False
    For i = 0 To Printers.Count - 1
        If VB.Printers(i).DeviceName = DeviceName Then
            Set VB.Printer = VB.Printers(i)
            SelectPrinter = True
            Exit Function
        End If
    Next
    
'when printing from crystal call this first
'CrxReport.SelectPrinter VB.Printer.DriverName, VB.Printer.DeviceName, VB.Printer.Port


End Function

Public Sub GetPrinterInfo(DeviceName As String, Port As String, Status As String, model As String, Location As String, Comment As String, IsDefault As Boolean)
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
            
            model = PointerToStringA(pi2.pDriverName)
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
            result = left$(result, InStr(result, vbNullChar) - 1)
         End If
      Else 'NT4 or less
         ' The old WIN.INI [Windows] section is mapped to
         ' HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows
         ' and we can just use GetProfileString to extract! :-)
         ' Returns: "printer name,driver name,port"
         If GetProfileString("Windows", ByVal "device", "", result, BufSize) Then
            ' Truncate buffer at end of name.
            result = left$(result, InStr(result, ",") - 1)
         End If
      End If
   End If
      
   ' Return default printer name.
   DefaultPrinterName = result
End Function



Public Sub ShowPropertiesDialog(DeviceName As String, Optional ByVal hWndParent As Long = 0)
   Dim hPrn As Long
   Dim pd As PRINTER_DEFAULTS
   ' HOWTO: Open the Printer Properties Dialog
   ' http://support.microsoft.com/support/kb/articles/Q198/8/60.asp
   pd.pDatatype = vbNullString
   ' Try admin access first
   pd.pDesiredAccess = PRINTER_ALL_ACCESS
   If OpenPrinter(DeviceName, hPrn, pd) = 0 Then
      ' Not an admin, try reduced privileges
      pd.pDesiredAccess = STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_USE
      Call OpenPrinter(DeviceName, hPrn, pd)
   End If
   ' Show dialog, if we have a handle to printer
   If hPrn Then
      Call PrinterProperties(hWndParent, hPrn)
      Call ClosePrinter(hPrn)
   End If
End Sub

Public Function DBUG() As Boolean
    DBUG = InIde Or UCase(Right(App.EXEName, 5)) = "DEBUG"
End Function
Public Sub DPNT(s As String)
If DBUG Then MsgBox s, vbInformation, "Debug"
End Sub

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

Public Function FindForm(FormName As String, TagValue As String) As Form
    Dim f As Form
    For Each f In Forms
        If f.Name = FormName And f.tag = TagValue Then
            Set FindForm = f
            f.SetFocus
            Exit Function
        End If
    Next
End Function


Public Function ValidateItem(ByRef EditText As String, WarningMsg As String, sql As String)
On Error Resume Next

'    If EditText = "" Then
'        ValidateItem = True
'    Else
'        EditText = App.SqlExec(sql, dbWorkticket)(0)
'        If Err.Number = 0 Then
'            ValidateItem = True
'        Else
'            If WarningMsg <> "" Then MsgBox WarningMsg, vbExclamation, App.ProductName
'            ValidateItem = False
'        End If
'    End If

End Function

Public Function FullTrim(ByVal s As String) As String
    
    While IsIn(Right(s, 1), vbCr, vbLf, vbTab, " ")
        s = Mid(s, 1, Len(s) - 1)
    Wend
    While IsIn(left(s, 1), vbCr, vbLf, vbTab, " ")
        s = Mid(s, 2, Len(s) - 1)
    Wend
    FullTrim = s
    
End Function


' Set the width of the list area of a ComboBox (in pixels)

Sub SetComboDropDownWidth(Combobox As Combobox, ByVal lWidth As Long)
    SendMessage Combobox.hwnd, CB_SETDROPPEDWIDTH, lWidth, ByVal 0&
End Sub

Function ReadFileToString(ByVal sFile As String) As String
    Dim s As String
    Dim i As Integer
    
    i = FreeFile
    Open sFile For Binary Shared As #i
    s = Space(LOF(i))
    Get #i, , s
    Close #i
    
    ReadFileToString = s

End Function


