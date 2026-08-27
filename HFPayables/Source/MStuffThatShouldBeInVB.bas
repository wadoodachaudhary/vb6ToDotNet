Attribute VB_Name = "MStuffThatShouldBeInVB"
Option Explicit


Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function SefocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

Public Const KEYEVENTF_KEYUP = &H2
Public Const VK_ENTER = &HD
Public Const VK_TAB = &H9
Public Const GW_HWNDNEXT = 2
Public Const GW_CHILD = 5

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
    FileName As String
    FileType As String
End Type
Public Declare Function MAPILogon Lib "MAPI32.DLL" (ByVal UIParam&, ByVal User$, ByVal Password$, ByVal flags&, ByVal Reserved&, Session&) As Long
Public Declare Function MAPILogoff Lib "MAPI32.DLL" (ByVal Session&, ByVal UIParam&, ByVal flags&, ByVal Reserved&) As Long
Public Declare Function BMAPIReadMail Lib "MAPI32.DLL" (lMsg&, nRecipients&, nFiles&, ByVal Session&, ByVal UIParam&, MessageID$, ByVal Flag&, ByVal Reserved&) As Long
Public Declare Function BMAPIGetReadMail Lib "MAPI32.DLL" (ByVal lMsg&, Message As MAPIMessage, Recip() As MapiRecip, File() As MapiFile, Originator As MapiRecip) As Long
Public Declare Function MAPIFindNext Lib "MAPI32.DLL" Alias "BMAPIFindNext" (ByVal Session&, ByVal UIParam&, MsgType$, SeedMsgID$, ByVal Flag&, ByVal Reserved&, MsgID$) As Long
Public Declare Function MAPISendDocuments Lib "MAPI32.DLL" (ByVal UIParam&, ByVal DelimStr$, ByVal FilePaths$, ByVal FileNames$, ByVal Reserved&) As Long
Public Declare Function MAPIDeleteMail Lib "MAPI32.DLL" (ByVal Session&, ByVal UIParam&, ByVal MsgID$, ByVal flags&, ByVal Reserved&) As Long
Public Declare Function MAPISendMail Lib "MAPI32.DLL" Alias "BMAPISendMail" (ByVal Session&, ByVal UIParam&, Message As MAPIMessage, Recipient() As MapiRecip, File() As MapiFile, ByVal flags&, ByVal Reserved&) As Long
Public Declare Function MAPISaveMail Lib "MAPI32.DLL" Alias "BMAPISaveMail" (ByVal Session&, ByVal UIParam&, Message As MAPIMessage, Recipient() As MapiRecip, File() As MapiFile, ByVal flags&, ByVal Reserved&, MsgID$) As Long
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
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long

Public Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Const CLR_INVALID = -1

Public Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function CreateEllipticRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Public Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
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

Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Declare Function SelectObject& Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long)
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

Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long






Public Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Public Declare Function InitCommonControlsEx Lib "COMCTL32.DLL" _
   (iccex As tagInitCommonControlsEx) As Boolean
Public Const ICC_USEREX_CLASSES = &H200



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


Public Declare Function DrawEdge& Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long)
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


Public Declare Function PathCompactPath Lib "shlwapi" Alias "PathCompactPathA" (ByVal hdc As Long, ByVal lpszPath As String, ByVal dx As Long) As Long


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
        TempPath = left(bufferStr, Len(bufferStr) - 1)
    End If

End Function
Public Function TempFile(Optional path As String, Optional Extension As String) As String
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
    Dim FileName As String
    FileName = CreateGUID
    FileName = Replace(FileName, "{", "")
    FileName = Replace(FileName, "}", "")
    If Extension <> "" Then FileName = ForceExt(FileName, Extension)
    
    TempFile = PathAppend(TempPath, FileName)



End Function

Public Function ShellFile(hwnd As Long, FileName As String, Optional DlgTitle, Optional FindItYourself As Boolean = True) As String
'returns new FileName if they found the file they wanted.
    Dim bFailed As Boolean
    Dim bOkNow  As Boolean
    Dim sNewFileName As String
    Dim rc  As Long
    Dim Msg As String
    Const Caption = "Error Launching File"
    
    sNewFileName = FileName
TryAgain:
    
    rc = ShellExecute(hwnd, "", sNewFileName, vbNullString, "C:\", SW_SHOWNORMAL)
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
    Dim FileName As String
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
                FileName = String(MAX_PATH, " ")
                namelen = DragQueryFile(hDrop, i, FileName, Len(FileName))
'                returnstr = returnstr & Chr(0) & left(FileName, namelen)
                returnstr = returnstr & ";" & left(FileName, namelen)
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
   nRet = RegCreateKeyEx(hive, section, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, hKey, nResult)
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

Public Function IniGetSectionNames(FileName As String) As String

    'returns a csv list of the section names in file

    Dim rc     As Long
    Dim sBuff  As String

    sBuff = Space$(BUFFER_SIZE)
    rc = GetPrivateProfileSectionNames(sBuff, BUFFER_SIZE, FileName)
    If rc > 0 Then
        sBuff = left(sBuff, rc - 1)
        sBuff = Replace(sBuff, Chr$(0), ",")
        If Right(sBuff, 1) = "," Then
            sBuff = left(sBuff, Len(sBuff) - 1)
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

Public Sub IniPut(FileName As String, SectionName As String, VariableName As String, VariableValue As Variant)

    Const DFORMAT = "hh:mm mmmm d yyyy"

    
    Call CreatePath("", FilePath(FileName))
    
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

Public Function FormatPhone(Number As String) As String
    Dim s As String
    Dim i As Long
    
    'remove all alpha
    s = ""
    For i = 1 To Len(Number)
        If IsNumeric(Mid(Number, i, 1)) Or Mid(Number, i, 1) = " " Then
            s = s & Mid(Number, i, 1)
        End If
    Next
    
    'add blank area code if not given
    If Len(s) = 7 Then s = "   " & s
    
    'format it
    s = Format(s, "!(@@@) @@@-@@@@ Ext(&&&&&&&&&&&&&&&&&&&&&&&&&&)")
    
    'remove empty extensions
    If Right(s, 6) = " Ext()" Then s = Mid(s, 1, Len(s) - 6)
    
    'remove empty area codes
    If left(s, 6) = "(   ) " Then s = Mid(s, 7)
    
    FormatPhone = s
End Function


Public Function ParsePhoneNumber(Number, Item As String) As String
On Error Resume Next
    'expects:  011(403)341-1234ext1234
    'spaces and no dash are ok but
    '  -area code must be in parenthesis
    '  -"ext" must preceed extension number
    'country,area,ext are optional
    Dim s As String
    Number = LCase(Number)
    Select Case LCase(Item)
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

Public Function ItemInList(Item As String, list As String, Optional ByVal Delimeter As String = ",") As Boolean
    ItemInList = InStr(1, Delimeter & list & Delimeter, Delimeter & Item & Delimeter)
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
Public Function Between(Minimum, Maximum, Value)
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
    If g.Row < 1 Then g.Row = 1
End Sub

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


Public Function IsIn(expression, ParamArray items()) As Boolean
    Dim i As Integer
    For i = 0 To UBound(items)
        If "" & expression = "" & items(i) Then
            IsIn = True
            Exit Function
        End If
    Next
    IsIn = False
End Function

Public Function Decode(expression As String, ParamArray ValueReturn())
    Dim i As Integer
    For i = 0 To UBound(ValueReturn) Step 2
        If expression = ValueReturn(i) Then
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
    InIde = (err.Number <> 0)
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

Public Function SetListIndex(ComboBox, Optional ItemData As Long, Optional Text As String) As Boolean
    Dim i As Long
    
    If Text <> "" Then
        For i = 0 To ComboBox.ListCount - 1
            If ComboBox.list(i) = Text Then
                ComboBox.ListIndex = i
                SetListIndex = True
                Exit Function
            End If
        Next
    Else
        If ItemData = 0 Then
            On Error Resume Next
            ComboBox.Text = ""
            SetListIndex = True
            Exit Function
        Else
            For i = 0 To ComboBox.ListCount - 1
                If ComboBox.ItemData(i) = ItemData Then
                    ComboBox.ListIndex = i
                    SetListIndex = True
                    Exit Function
                End If
            Next
        End If
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

Public Function SpaceCase(s As String) As String
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

Public Function ShowForm(FormName As String, TagValue As String) As Boolean
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
        If lv.ListItems(i).selected Then s = s & "," & lv.ListItems(i).Key
    Next
    ListViewSelected = Mid(s, 2)
End Function

Public Function ImgComboFindItem(cbo) As Boolean
    Dim i As Long
    For i = 1 To cbo.ComboItems.Count
        If cbo.Text = cbo.ComboItems(i).Text Then
            cbo.ComboItems(i).selected = True
            ImgComboFindItem = False
            Exit Function
        End If
    Next
    cbo.ComboItems(1).selected = True
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
    FilePath = left(fullpath, Len(fullpath) - Len(FileTitle(fullpath)))
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

Public Function CompactedPathSh(ByVal sPath As String, ByVal lMaxWidthPixels As Long, Optional hdc As Long) As String
    Dim lR As Long
    Dim iPos As Long
    
    If hdc = 0 Then hdc = Screen.ActiveForm.hdc
    lR = PathCompactPath(hdc, sPath, lMaxWidthPixels)
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
        If .Splitl.Count > 1 Then
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

Public Sub AddToComboList(ComboBox As ComboBox)
    Dim i As Long
    If ComboBox.Text = "" Then Exit Sub
    For i = 0 To ComboBox.ListCount - 1
        If ComboBox.list(i) = ComboBox.Text Then Exit Sub
    Next
    ComboBox.AddItem ComboBox.Text, 0
End Sub


Public Function ImageIndex(ImageList, Key As String) As Long
    ImageIndex = ImageList.ListImages(Key).Index - 1
End Function

Public Function DNull(s As String) As String
    DNull = Mid(s, 1, Max(0, InStr(1, s, vbNullChar) - 1))
End Function

Public Function PathIsLocalPath(ByVal sPath As String) As Boolean
   PathIsLocalPath = PathIsNetworkPath(sPath) = False
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
   InitCommonControlsVB = (err.Number = 0)
   On Error GoTo 0
End Function


Public Function FileExists(FileName As String) As Boolean
    FileExists = Dir(FileName) <> ""
End Function

   
   

Public Function PathExists(path As String) As Boolean
    PathExists = Dir(path, vbDirectory) <> "" Or Dir(path) <> ""
End Function

Public Sub CreatePath(PathDescription As String, path As String)
On Error Resume Next
    Dim i As Long
    Dim s As String
    s = Parse(path, 1, "\")
    For i = 2 To Parse(path, , "\")
        s = s & "\" & Parse(path, i, "\")
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
Public Function DeQuote(s As String) As String
    s = Trim(s)
    If left(s, 1) = vbQuote And Right(s, 1) = vbQuote Then
        s = Mid(s, 2, Len(s) - 2)
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
Public Function ListItem(ComboBox, Optional Text, Optional ItemData, Optional Compare As VbCompareMethod = vbTextCompare) As Long
    Dim i As Long
    If Not IsMissing(Text) Then
        For i = 0 To ComboBox.ListCount - 1
            If StrComp(Text, ComboBox.list(i), Compare) = 0 Then
                ListItem = i
                Exit Function
            End If
        Next
    End If
    If Not IsMissing(ItemData) Then
        For i = 0 To ComboBox.ListCount - 1
            If ItemData = ComboBox.ItemData(i) Then
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
Public Sub IniGetGrid(Form As Form, Grid As Object, Optional IniFile As String, Optional StaticPositions As Boolean = False, Optional InstanceKey, Optional Grouped As Boolean)
    If DBGetGrid(Form, Grid, StaticPositions, InstanceKey, Grouped) Then Exit Sub
    
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
                        If Caption <> "" Then Grid.TextMatrix(0, ColIndex) = Caption

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
    
    
eh: Call ErrHandler("IniPutGrid")
End Sub

Public Function DBGetGrid(Form As Form, Grid As Object, Optional StaticPositions As Boolean = False, Optional InstanceKey, Optional Grouped As Boolean) As Boolean
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
                
                .Splits(split).Columns(Col).Caption = "" & rs("caption")
                .Splits(split).Columns(Col).Width = Val("" & rs("width"))
                .Splits(split).Columns(Col).Visible = "" & rs("hidden") <> "True"
                If Not StaticPositions Then .Splits(split).Columns(Col).Order = Val("" & rs("colindex"))
    
                rs.MoveNext
            Wend
        
        Else
        
            AllColsHidden = True
            While Not rs.EOF
                Col = .ColIndex("" & rs("colkey"))
                .TextMatrix(0, Col) = "" & rs("caption")
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
 
    Dim Sql As String
    
    Dim gridname As String
    Dim split As Long
    Dim Col As Long
    
    
    Sql = ""
    With Grid
    
        'gridname = FormName.GridName(Index).InstanceKey
        gridname = Form.Name & "." & .Name
        gridname = gridname & ".(" & .Index & ")"
        If Not IsMissing(InstanceKey) Then gridname = gridname & "." & InstanceKey
            
            
        If TypeName(Grid) = "TDBGrid" Then
            For split = 0 To Grid.Splits.Count - 1
                For Col = 0 To Grid.Columns.Count - 1
                    'uid,gridname,split,colindex,colkey,caption,width,hidden,grouped
                    Sql = Sql & ",(" & DbQuote(Str, HFApp.LoginID) & _
                                "," & DbQuote(Str, gridname) & _
                                "," & DbQuote(num, split) & _
                                "," & DbQuote(num, Grid.Splits(split).Columns(Col).Order) & _
                                "," & DbQuote(Str, Grid.Splits(split).Columns(Col).DataField) & _
                                "," & DbQuote(Str, Grid.Splits(split).Columns(Col).Caption) & _
                                "," & DbQuote(num, Grid.Splits(split).Columns(Col).Width) & _
                                "," & DbQuote(Bit, Grid.Splits(split).Columns(Col).Visible) & _
                                "," & DbQuote(Bit, False) & vbCrLf
                Next
            Next
        Else
            For Col = 0 To .Cols - 1
                'uid,gridname,split,colindex,colkey,caption,width,hidden,grouped
                Sql = Sql & ",(" & DbQuote(Str, HFApp.LoginID) & _
                            "," & DbQuote(Str, gridname) & _
                            ",0" & _
                            "," & DbQuote(num, Col) & _
                            "," & DbQuote(Str, .ColKey(Col)) & _
                            "," & DbQuote(Str, .TextMatrix(0, Col)) & _
                            "," & DbQuote(num, .ColWidth(Col)) & _
                            "," & DbQuote(Bit, .ColHidden(Col)) & _
                            "," & DbQuote(Bit, .ColData(Col) = "GROUPED") & ")" & vbCrLf
            Next
        End If
    End With
    Sql = "delete AppGridLayout" & vbCrLf & _
          "where uid=" & DbQuote(Str, HFApp.LoginID) & vbCrLf & _
          "and gridname=" & DbQuote(Str, gridname) & vbCrLf & _
          vbCrLf & _
          "insert into AppGridLayout(uid,gridname,split,colindex,colkey,caption,width,hidden,grouped) values" & vbCrLf & _
          Mid(Sql, 2)
    
    Call HFApp.SqlExec(Sql)
    

Exit Sub
eh: Call ErrHandler("IniPutGrid")
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
                    bCaptions = CBool(Parse(c.tag, 2))
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

Public Sub IniGetForm(Form As Form, Optional IniFile As String, Optional InstanceKey)
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
            
            Case "Toolbar"
                bVisible = True 'IniGet(IniFile, FormName, CtrlName & ".Visible", True)
                bCaptions = True 'IniGet(IniFile, FormName, CtrlName & ".Captions", True)
                c.tag = bVisible & "," & bCaptions
                c.Visible = bVisible
                Call ShowToolbarCaptions(c, bCaptions)
            
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
    
End Sub


Public Property Get AppWorkingFolder() As String
    AppWorkingFolder = PathAppend(GetFolderPath(CSIDL_LOCAL_APPDATA), "Homefront\" & VB.App.ExeName)
End Property
Public Property Get AppIni() As String
    AppIni = PathAppend(AppWorkingFolder, "settings.ini")
    AppIni = PathAppend(HFApp.SystemFolder, "User Settings", NetworkLoginName & ".ini")
End Property
Public Property Get AppSQLCache() As String
    AppSQLCache = PathAppend(AppWorkingFolder, "query.sql")
End Property
Public Property Get AppErrorLog() As String
    AppErrorLog = PathAppend(AppWorkingFolder, "errors.txt")
End Property
Public Property Get AppSQLLog() As String
    AppSQLLog = PathAppend(AppWorkingFolder, "sql.txt")
End Property

Public Function FullTrim(ByVal s As String) As String
    
    While IsIn(Right(s, 1), vbCr, vbLf, vbTab, " ")
        s = Mid(s, 1, Len(s) - 1)
    Wend
    While IsIn(left(s, 1), vbCr, vbLf, vbTab, " ")
        s = Mid(s, 2, Len(s) - 1)
    Wend
    FullTrim = s
    
End Function


Public Sub SetCtrlFocus(c As Control)
On Error Resume Next
    Call c.SetFocus
End Sub

Public Sub LoadComboBox(ComboBox, Connection, Sql As String)

    Dim rs  'As ADODB.Recordset
    Dim tag As String
    
    'SELECT ItemDesc  text to be displayed in list
    '      ,KeyValue  alpha key value
    '      ,IDValue   numeric key value
    '  FROM Table
    
    'ID are stored in combobox.itemdata(x)
    'Keys are stored chr(1) separated in the controls tag property
    '     use SetComboBoxListIndex & GetComboBoxListKey, GetComboBoxListID, GetComboBoxListIndex functions
    
    Set rs = Connection.Execute(Sql)
    ComboBox.Clear
    tag = ""
    While Not rs.EOF
        ComboBox.AddItem Trim("" & rs(0))
        ComboBox.ItemData(ComboBox.NewIndex) = Val("" & rs(2))
        tag = tag & Chr(1) & "" & rs(1)
        rs.MoveNext
    Wend
    ComboBox.tag = Mid(tag, 2)
End Sub
Public Function GetComboBoxListKey(ComboBox) As String
If ComboBox.ListIndex > -1 Then GetComboBoxListKey = Parse(ComboBox.tag, ComboBox.ListIndex + 1, Chr(1))
End Function
Public Function GetComboBoxListID(ComboBox) As Long
    GetComboBoxListID = ComboBox.ItemData(ComboBox.ListIndex)
End Function
Public Function GetComboBoxListIndex(ComboBox) As Long
    GetComboBoxListIndex = ComboBox.ListIndex
End Function
Public Sub SetComboBoxListIndex(ComboBox, Optional Text As String, Optional Key As String, Optional ID As Long = -1, Optional Index As Long = -1)
On Error Resume Next
    Dim s As String
    Dim i As Long
    
    If Index <> -1 Then
        ComboBox.ListIndex = Index
    End If
    
    If Text <> "" Then
        For i = 0 To ComboBox.ListCount - 1
            If UCase(ComboBox.list(i)) = UCase(Text) Then
                ComboBox.ListIndex = i
                Exit Sub
            End If
        Next
    End If
    
    
    If Key <> "" Then
        s = ComboBox.tag
        For i = 1 To Parse(s, , Chr(1))
            If UCase(Parse(s, i, Chr(1))) = UCase(Key) Then
                ComboBox.ListIndex = i - 1
                Exit Sub
            End If
        Next
    End If
    
    If ID <> -1 Then
        For i = 0 To ComboBox.ListCount - 1
            If ComboBox.ItemData(i) = ID Then
                ComboBox.ListIndex = i
                Exit Sub
            End If
        Next
    End If
End Sub

Public Sub TabToNextCtrl(f As Form)
On Error Resume Next
'move focus to next control in tab order
'send keys is not reliable
    
    
    Dim cti As Long    'current ctrl tab index
    Dim nc As Control  'next ctrl
    Dim nti As Long    'next ctrl tab index
    Dim tc As Control  'this ctrl
    Dim taf As Boolean 'this ctrl accepts focus
    Dim tti As Long    'this ctrl tab index
    
    cti = f.ActiveControl.TabIndex
    nti = 999999

    For Each tc In f.Controls
    

        tti = -1
        tti = tc.TabIndex
        taf = True
        taf = taf And tc.Enabled
        taf = taf And tc.Visible
        taf = taf And tc.TabStop
        
        If taf And tti > cti And tti < nti Then
            Set nc = tc
            nti = nc.TabIndex
        End If
    Next
    If nc Is Not Null Then nc.SetFocus

eh:
Exit Sub
End Sub

Public Function SaveToCSV(FileName As String, Optional Sheet As Integer = 1) As String
    Dim s As String
    Dim xlSheet As Object 'Excel.Worksheet
    Set xlSheet = GetObject(FileName).Sheets(Sheet)
    s = TempFile(, "csv")
    On Error Resume Next
    Kill s
    On Error GoTo 0
    Call xlSheet.SaveAs(s, 6, , , , , False)
    Set xlSheet = Nothing
    SaveToCSV = s

' THIS IS NOT VERSION INDEPENDENT??? SEEMS LIKE IT ISN'T
'    Dim s As String
'    Dim xlApp As New Excel.Application
'    Dim xlWkbk As Excel.Workbook
'    Dim xlSheet As Excel.Worksheet
'
'    Set xlWkbk = xlApp.Workbooks.Open(Filename)
'
'    On Error Resume Next
'    Set xlSheet = xlWkbk.Sheets(Sheet)
'    If err.Number = 9 Then
'        On Error GoTo 0
'        err.Raise 9, "SaveToCSV", "Worksheet number " & Sheet & " not found."
'    End If
'    On Error GoTo 0
'
'    s = TempFile(, "csv")
'    On Error Resume Next
'    Kill s
'    On Error GoTo 0
'    Call xlSheet.SaveAs(s, Excel.xlCSV, , , , , False)
'
'    Call xlWkbk.Close(False)
'    Set xlSheet = Nothing
'    Set xlWkbk = Nothing
'    Set xlApp = Nothing
'
'    SaveToCSV = s

End Function

Public Function AddQuotes(s As String) As String
    AddQuotes = vbQuote & s & vbQuote
End Function

