Attribute VB_Name = "MCmnDlg"
Option Explicit

Public Enum EErrorCommonDialog
    eeBaseCommonDialog = 13450  ' CommonDialog
End Enum

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalCompact Lib "kernel32" (ByVal dwMinFree As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Sub CopyMemoryStr Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, ByVal lpvSource As String, ByVal cbCopy As Long)

Private Const MAX_PATH = 260
Private Const MAX_FILE = 260

Private Type OpenFileName
    lStructSize As Long          ' Filled with UDT size
    hwndOwner As Long            ' Tied to Owner
    hInstance As Long            ' Ignored (used only by templates)
    lpstrFilter As String        ' Tied to Filter
    lpstrCustomFilter As String  ' Ignored (exercise for reader)
    nMaxCustFilter As Long       ' Ignored (exercise for reader)
    nFilterIndex As Long         ' Tied to FilterIndex
    lpstrFile As String          ' Tied to FileName
    nMaxFile As Long             ' Handled internally
    lpstrFileTitle As String     ' Tied to FileTitle
    nMaxFileTitle As Long        ' Handled internally
    lpstrInitialDir As String    ' Tied to InitDir
    lpstrTitle As String         ' Tied to DlgTitle
    flags As Long                ' Tied to Flags
    nFileOffset As Integer       ' Ignored (exercise for reader)
    nFileExtension As Integer    ' Ignored (exercise for reader)
    lpstrDefExt As String        ' Tied to DefaultExt
    lCustData As Long            ' Ignored (needed for hooks)
    lpfnHook As Long             ' Ignored (good luck with hooks)
    lpTemplateName As Long       ' Ignored (good luck with templates)
End Type

Private Declare Function GetOpenFileName Lib "COMDLG32" Alias "GetOpenFileNameA" (File As OpenFileName) As Long
Private Declare Function GetSaveFileName Lib "COMDLG32" Alias "GetSaveFileNameA" (File As OpenFileName) As Long
Private Declare Function GetFileTitle Lib "COMDLG32" Alias "GetFileTitleA" (ByVal szFile As String, ByVal szTitle As String, ByVal cbBuf As Long) As Long

Public Enum EOpenFile
    OFN_READONLY = &H1
    OFN_OVERWRITEPROMPT = &H2
    OFN_HIDEREADONLY = &H4
    OFN_NOCHANGEDIR = &H8
    OFN_SHOWHELP = &H10
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_NOVALIDATE = &H100
    OFN_ALLOWMULTISELECT = &H200
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_PATHMUSTEXIST = &H800
    OFN_FILEMUSTEXIST = &H1000
    OFN_CREATEPROMPT = &H2000
    OFN_SHAREAWARE = &H4000
    OFN_NOREADONLYRETURN = &H8000
    OFN_NOTESTFILECREATE = &H10000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOLONGNAMES = &H40000
    OFN_EXPLORER = &H80000
    OFN_NODEREFERENCELINKS = &H100000
    OFN_LONGNAMES = &H200000
End Enum

Private Type TCHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
End Type

Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (color As TCHOOSECOLOR) As Long

Public Enum EChooseColor
    CC_RGBInit = &H1
    CC_FullOpen = &H2
    CC_PreventFullOpen = &H4
    CC_ColorShowHelp = &H8
' Win95 only
    CC_SolidColor = &H80
    CC_AnyColor = &H100
' End Win95 only
    CC_ENABLEHOOK = &H10
    CC_ENABLETEMPLATE = &H20
    CC_EnableTemplateHandle = &H40
End Enum
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Type TCHOOSEFONT
    lStructSize As Long         ' Filled with UDT size
    hwndOwner As Long           ' Caller's window handle
    hDC As Long                 ' Printer DC/IC or NULL
    lpLogFont As Long           ' Pointer to LOGFONT
    iPointSize As Long          ' 10 * size in points of font
    flags As Long               ' Type flags
    rgbColors As Long           ' Returned text color
    lCustData As Long           ' Data passed to hook function
    lpfnHook As Long            ' Pointer to hook function
    lpTemplateName As Long      ' Custom template name
    hInstance As Long           ' Instance handle for template
    lpszStyle As String         ' Return style field
    nFontType As Integer        ' Font type bits
    iAlign As Integer           ' Filler
    nSizeMin As Long            ' Minimum point size allowed
    nSizeMax As Long            ' Maximum point size allowed
End Type
Private Declare Function ChooseFont Lib "COMDLG32" Alias "ChooseFontA" (chfont As TCHOOSEFONT) As Long

Private Const LF_FACESIZE = 32
Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Public Enum EChooseFont
    CF_ScreenFonts = &H1
    CF_PrinterFonts = &H2
    CF_BOTH = &H3
    CF_FontShowHelp = &H4
    CF_UseStyle = &H80
    CF_EFFECTS = &H100
    CF_AnsiOnly = &H400
    CF_NoVectorFonts = &H800
    CF_NoOemFonts = CF_NoVectorFonts
    CF_NoSimulations = &H1000
    CF_LimitSize = &H2000
    CF_FixedPitchOnly = &H4000
    CF_WYSIWYG = &H8000  ' Must also have ScreenFonts AND PrinterFonts
    CF_ForceFontExist = &H10000
    CF_ScalableOnly = &H20000
    CF_TTOnly = &H40000
    CF_NoFaceSel = &H80000
    CF_NoStyleSel = &H100000
    CF_NoSizeSel = &H200000
    ' Win95 only
    CF_SelectScript = &H400000
    CF_NoScriptSel = &H800000
    CF_NoVertFonts = &H1000000

    CF_InitToLogFontStruct = &H40
    CF_Apply = &H200
    CF_EnableHook = &H8
    CF_EnableTemplate = &H10
    CF_EnableTemplateHandle = &H20
    CF_FontNotSupported = &H238
End Enum

' These are extra nFontType bits that are added to what is returned to the
' EnumFonts callback routine

Public Enum EFontType
    Simulated_FontType = &H8000
    Printer_FontType = &H4000
    Screen_FontType = &H2000
    Bold_FontType = &H100
    Italic_FontType = &H200
    Regular_FontType = &H400
End Enum

Private Type TPRINTDLG
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hDC As Long
    flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As Long
    lpSetupTemplateName As Long
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type

'  DEVMODE collation selections
Private Const DMCOLLATE_FALSE = 0
Private Const DMCOLLATE_TRUE = 1

Private Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (prtdlg As TPRINTDLG) As Integer

Public Enum EPrintDialog
    PD_ALLPAGES = &H0
    PD_SELECTION = &H1
    PD_PAGENUMS = &H2
    PD_NOSELECTION = &H4
    PD_NOPAGENUMS = &H8
    PD_COLLATE = &H10
    PD_PRINTTOFILE = &H20
    PD_PRINTSETUP = &H40
    PD_NOWARNING = &H80
    PD_RETURNDC = &H100
    PD_RETURNIC = &H200
    PD_RETURNDEFAULT = &H400
    PD_SHOWHELP = &H800
    PD_ENABLEPRINTHOOK = &H1000
    PD_ENABLESETUPHOOK = &H2000
    PD_ENABLEPRINTTEMPLATE = &H4000
    PD_ENABLESETUPTEMPLATE = &H8000
    PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
    PD_ENABLESETUPTEMPLATEHANDLE = &H20000
    PD_USEDEVMODECOPIES = &H40000
    PD_USEDEVMODECOPIESANDCOLLATE = &H40000
    PD_DISABLEPRINTTOFILE = &H80000
    PD_HIDEPRINTTOFILE = &H100000
    PD_NONETWORKBUTTON = &H200000
End Enum

Private Type DEVNAMES
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
End Type

Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Type DevMode
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

' New Win95 Page Setup dialogs are up to you
Private Type POINTL
    X As Long
    Y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Private Type TPAGESETUPDLG
    lStructSize                 As Long
    hwndOwner                   As Long
    hDevMode                    As Long
    hDevNames                   As Long
    flags                       As Long
    ptPaperSize                 As POINTL
    rtMinMargin                 As RECT
    rtMargin                    As RECT
    hInstance                   As Long
    lCustData                   As Long
    lpfnPageSetupHook           As Long
    lpfnPagePaintHook           As Long
    lpPageSetupTemplateName     As Long
    hPageSetupTemplate          As Long
End Type

' EPaperSize constants same AS vbPRPS constants
Public Enum EPaperSize
    epsLetter = 1          ' Letter, 8 1/2 x 11 in.
    epsLetterSmall         ' Letter Small, 8 1/2 x 11 in.
    epsTabloid             ' Tabloid, 11 x 17 in.
    epsLedger              ' Ledger, 17 x 11 in.
    epsLegal               ' Legal, 8 1/2 x 14 in.
    epsStatement           ' Statement, 5 1/2 x 8 1/2 in.
    epsExecutive           ' Executive, 7 1/2 x 10 1/2 in.
    epsA3                  ' A3, 297 x 420 mm
    epsA4                  ' A4, 210 x 297 mm
    epsA4Small             ' A4 Small, 210 x 297 mm
    epsA5                  ' A5, 148 x 210 mm
    epsB4                  ' B4, 250 x 354 mm
    epsB5                  ' B5, 182 x 257 mm
    epsFolio               ' Folio, 8 1/2 x 13 in.
    epsQuarto              ' Quarto, 215 x 275 mm
    eps10x14               ' 10 x 14 in.
    eps11x17               ' 11 x 17 in.
    epsNote                ' Note, 8 1/2 x 11 in.
    epsEnv9                ' Envelope #9, 3 7/8 x 8 7/8 in.
    epsEnv10               ' Envelope #10, 4 1/8 x 9 1/2 in.
    epsEnv11               ' Envelope #11, 4 1/2 x 10 3/8 in.
    epsEnv12               ' Envelope #12, 4 1/2 x 11 in.
    epsEnv14               ' Envelope #14, 5 x 11 1/2 in.
    epsCSheet              ' C size sheet
    epsDSheet              ' D size sheet
    epsESheet              ' E size sheet
    epsEnvDL               ' Envelope DL, 110 x 220 mm
    epsEnvC3               ' Envelope C3, 324 x 458 mm
    epsEnvC4               ' Envelope C4, 229 x 324 mm
    epsEnvC5               ' Envelope C5, 162 x 229 mm
    epsEnvC6               ' Envelope C6, 114 x 162 mm
    epsEnvC65              ' Envelope C65, 114 x 229 mm
    epsEnvB4               ' Envelope B4, 250 x 353 mm
    epsEnvB5               ' Envelope B5, 176 x 250 mm
    epsEnvB6               ' Envelope B6, 176 x 125 mm
    epsEnvItaly            ' Envelope, 110 x 230 mm
    epsenvmonarch          ' Envelope Monarch, 3 7/8 x 7 1/2 in.
    epsEnvPersonal         ' Envelope, 3 5/8 x 6 1/2 in.
    epsFanfoldUS           ' U.S. Standard Fanfold, 14 7/8 x 11 in.
    epsFanfoldStdGerman    ' German Standard Fanfold, 8 1/2 x 12 in.
    epsFanfoldLglGerman    ' German Legal Fanfold, 8 1/2 x 13 in.
    epsUser = 256          ' User-defined
End Enum

' EPrintQuality constants same AS vbPRPQ constants
Public Enum EPrintQuality
    epqDraft = -1
    epqLow = -2
    epqMedium = -3
    epqHigh = -4
End Enum

Public Enum EOrientation
    eoPortrait = 1
    eoLandscape
End Enum

Private Declare Function PageSetupDlg Lib "COMDLG32" Alias "PageSetupDlgA" (lppage As TPAGESETUPDLG) As Boolean

Public Enum EPageSetup
    PSD_Defaultminmargins = &H0 ' Default (printer's)
    PSD_InWinIniIntlMeasure = &H0
    PSD_MINMARGINS = &H1
    PSD_MARGINS = &H2
    PSD_INTHOUSANDTHSOFINCHES = &H4
    PSD_INHUNDREDTHSOFMILLIMETERS = &H8
    PSD_DISABLEMARGINS = &H10
    PSD_DISABLEPRINTER = &H20
    PSD_NoWarning = &H80
    PSD_DISABLEORIENTATION = &H100
    PSD_ReturnDefault = &H400
    PSD_DISABLEPAPER = &H200
    PSD_ShowHelp = &H800
    PSD_EnablePageSetupHook = &H2000
    PSD_EnablePageSetupTemplate = &H8000
    PSD_EnablePageSetupTemplateHandle = &H20000
    PSD_EnablePagePaintHook = &H40000
    PSD_DisablePagePainting = &H80000
End Enum

Public Enum EPageSetupUnits
    epsuInches
    epsuMillimeters
End Enum

' Common dialog errors

Private Declare Function CommDlgExtendedError Lib "COMDLG32" () As Long

Public Enum EDialogError
    CDERR_DIALOGFAILURE = &HFFFF

    CDERR_GENERALCODES = &H0
    CDERR_STRUCTSIZE = &H1
    CDERR_INITIALIZATION = &H2
    CDERR_NOTEMPLATE = &H3
    CDERR_NOHINSTANCE = &H4
    CDERR_LOADSTRFAILURE = &H5
    CDERR_FINDRESFAILURE = &H6
    CDERR_LOADRESFAILURE = &H7
    CDERR_LOCKRESFAILURE = &H8
    CDERR_MEMALLOCFAILURE = &H9
    CDERR_MEMLOCKFAILURE = &HA
    CDERR_NOHOOK = &HB
    CDERR_REGISTERMSGFAIL = &HC

    PDERR_PRINTERCODES = &H1000
    PDERR_SETUPFAILURE = &H1001
    PDERR_PARSEFAILURE = &H1002
    PDERR_RETDEFFAILURE = &H1003
    PDERR_LOADDRVFAILURE = &H1004
    PDERR_GETDEVMODEFAIL = &H1005
    PDERR_INITFAILURE = &H1006
    PDERR_NODEVICES = &H1007
    PDERR_NODEFAULTPRN = &H1008
    PDERR_DNDMMISMATCH = &H1009
    PDERR_CREATEICFAILURE = &H100A
    PDERR_PRINTERNOTFOUND = &H100B
    PDERR_DEFAULTDIFFERENT = &H100C

    CFERR_CHOOSEFONTCODES = &H2000
    CFERR_NOFONTS = &H2001
    CFERR_MAXLESSTHANMIN = &H2002

    FNERR_FileNameCODES = &H3000
    FNERR_SUBCLASSFAILURE = &H3001
    FNERR_INVALIDFileName = &H3002
    FNERR_BUFFERTOOSMALL = &H3003

    CCERR_CHOOSECOLORCODES = &H5000
End Enum

' Array of custom colors lasts for life of app
Private alCustom(0 To 15) As Long, fNotFirst As Boolean

Public Enum EPrintRange
    eprAll
    eprPageNumbers
    eprSelection
End Enum
Private m_lApiReturn As Long
Private m_lExtendedError As Long
Private m_dvmode As DevMode

Private Const conNoError = 0&
Private Const conMaxPath = 260
Private Const conErrorExtendedError = 1208&

' Messages sent to callback function.
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2

' Messages to browser FROM callback function.
Private Const WM_USER = &H400
Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Private Const BFFM_ENABLEOK = (WM_USER + 101)
Private Const BFFM_SETSELECTION = (WM_USER + 102)

Private Type BROWSEINFO
    hwndOwner As Long           ' Owner
    pidlRoot As Long            ' Can be null
    strDisplayName As String    ' Recieves display name of folder
    strTitle As String          ' title/instructions for user
    ulFlags As Long             ' 0 or BIF constants
    lpfn As Long                ' Address for callback
    lParam As Long              ' Passes to callback
    iImage As Long              ' index to the system image list
End Type

' Shell32 APIs
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal Folder As Long, ByRef idl As Long) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal idl As Long, ByVal Path As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" (ByRef bi As BROWSEINFO) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long


' Shell constants
Public Const BIF_RETURNONLYFSDIRS = &H1&
Public Const BIF_DONTGOBELOWDOMAIN = &H2&
Public Const BIF_STATUSTEXT = &H4&
Public Const BIF_RETURNFSANCESTORS = &H8&
Public Const BIF_BROWSEFORCOMPUTER = &H1000&
Public Const BIF_BROWSEFORPRINTER = &H2000&


Private bffInitialSelection As Variant

Public Enum CSIDL_VALUES
    CSIDL_DESKTop = &H0
    CSIDL_INTERNET = &H1
    CSIDL_PROGRAMS = &H2
    CSIDL_CONTROLS = &H3
    CSIDL_PRINTERS = &H4
    CSIDL_PERSONAL = &H5
    CSIDL_FAVORITES = &H6
    CSIDL_STARTUP = &H7
    CSIDL_RECENT = &H8
    CSIDL_SENDTO = &H9
    CSIDL_BITBUCKET = &HA
    CSIDL_STARTMENU = &HB
    CSIDL_MYDOCUMENTS = &HC
    CSIDL_MYMUSIC = &HD
    CSIDL_MYVIDEO = &HE
    CSIDL_DESKTopDIRECTORY = &H10
    CSIDL_DRIVES = &H11
    CSIDL_NETWORK = &H12
    CSIDL_NETHOOD = &H13
    CSIDL_FONTS = &H14
    CSIDL_TEMPLATES = &H15
    CSIDL_COMMON_STARTMENU = &H16
    CSIDL_COMMON_PROGRAMS = &H17
    CSIDL_COMMON_STARTUP = &H18
    CSIDL_COMMON_DESKTopDIRECTORY = &H19
    CSIDL_APPDATA = &H1A
    CSIDL_PRINTHOOD = &H1B
    CSIDL_LOCAL_APPDATA = &H1C
    CSIDL_ALTSTARTUP = &H1D
    CSIDL_COMMON_ALTSTARTUP = &H1E
    CSIDL_COMMON_FAVORITES = &H1F
    CSIDL_INTERNET_CACHE = &H20
    CSIDL_COOKIES = &H21
    CSIDL_HISTORY = &H22
    CSIDL_COMMON_APPDATA = &H23
    CSIDL_WINDOWS = &H24
    CSIDL_SYSTEM = &H25
    CSIDL_PROGRAM_FILES = &H26
    CSIDL_MYPICTURES = &H27
    CSIDL_PROFILE = &H28
    CSIDL_SYSTEMX86 = &H29
    CSIDL_PROGRAM_FILESX86 = &H2A
    CSIDL_PROGRAM_FILES_COMMON = &H2B
    CSIDL_PROGRAM_FILES_COMMONX86 = &H2C
    CSIDL_COMMON_TEMPLATES = &H2D
    CSIDL_COMMON_DOCUMENTS = &H2E
    CSIDL_COMMON_ADMINTOOLS = &H2F
    CSIDL_ADMINTOOLS = &H30
    CSIDL_CONNECTIONS = &H31
    CSIDL_COMMON_MUSIC = &H35
    CSIDL_COMMON_PICTURES = &H36
    CSIDL_COMMON_VIDEO = &H37
    CSIDL_RESOURCES = &H38
    CSIDL_RESOURCES_LOCALIZED = &H39
    CSIDL_COMMON_OEM_LINKS = &H3A
    CSIDL_CDBURN_AREA = &H3B
    CSIDL_COMPUTERSNEARME = &H3D
    CSIDL_FLAG_PER_USER_INIT = &H800
    CSIDL_FLAG_NO_ALIAS = &H1000
    CSIDL_FLAG_DONT_VERIFY = &H4000
    CSIDL_FLAG_CREATE = &H8000
    CSIDL_FLAG_MASK = &HFF00
End Enum

Private Const MAX_LENGTH = 260
Private Const S_OK = 0
Private Const S_FALSE = 1

Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function SHGetFolderPath Lib "shfolder.dll" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwReserved As Long, ByVal lpszPath As String) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long

   
Public Function GetFolderPath(csidl As CSIDL_VALUES) As String

   Dim buff As String
  
   buff = Space$(MAX_LENGTH)
   
   If SHGetFolderPath(GetDesktopWindow(), csidl, -1, &H0, buff) = S_OK Then
       GetFolderPath = Left$(buff, lstrlenW(StrPtr(buff)))
   End If
   
End Function

Public Function VBChooseFolder(ByVal lngCSIdl As Long, ByVal lngBifFlags As Long, ByRef strFolder As String, _
                               Optional ByVal strTitle As String = "Select Folder", Optional ByVal hwnd As Long = 0, _
                               Optional ByVal vInitialSelection As Variant = CSIDL_DESKTop) As Boolean
 
    Dim bi As BROWSEINFO
    Dim lngReturn As Long
    Dim lngIDL As Long
    
    ' First get the special folder that lngCSIDL AND lngID specify
    If SHGetSpecialFolderLocation(hwnd, lngCSIdl, lngIDL) = 0 Then
        bffInitialSelection = vInitialSelection
        
        ' Set up the browse structure
        With bi
            .hwndOwner = hwnd
            .pidlRoot = lngIDL
            .strDisplayName = Space(conMaxPath)
            .strTitle = strTitle
            .ulFlags = lngBifFlags '+ BIF_BROWSEFORCOMPUTER
            .lpfn = FnPtrToLong(AddressOf CallBackFunction)
        End With

        ' Open the browse dialog
        lngIDL = SHBrowseForFolder(bi)
        
        ' If the dialog succeeded
        If lngIDL <> 0 Then
            strFolder = Space(conMaxPath)
            
            ' Try to resolve the lngIDL into a real path
            If CBool(SHGetPathFromIDList(lngIDL, strFolder)) Then
                ' Trim the null characters
                strFolder = TrimNull(strFolder)
                VBChooseFolder = True
            Else
                ' No real directory path is available, so at least return the virtual selection
                strFolder = TrimNull(bi.strDisplayName)
                VBChooseFolder = True
            End If
        Else
            ' Some unknown error has occurred
            VBChooseFolder = False
        End If
        ' It's up to the caller to free the memory used by
        ' the ID list.
        Call GlobalFree(lngIDL)
        
    Else
        ' Some unknown error has occurred
        VBChooseFolder = False
    End If
    
End Function

Private Function TrimNull(ByVal strValue As String) As String
    ' Find the first vbNullChar in a string, AND return
    ' everything prior to that character. Extremely
    ' useful when combined with the Windows API function calls.
    
    ' In:
    '   strValue:
    '       Input text, possibly containing a null character
    '       (chr$(0), or vbNullChar)
    ' Out:
    '   Return Value:
    '       strValue trimmed on the right, at the location
    '       of the null character, if there was one.
    
    Dim intPos As Integer
    
    intPos = InStr(strValue, vbNullChar)
    Select Case intPos
        ' It's best to put the most likely case first.
        Case Is > 1
            ' Found in the string, so return the portion
            ' up to the null character.
            TrimNull = Left$(strValue, intPos - 1)
        Case 0
            ' Not found at all, so just
            ' return the original value.
            TrimNull = strValue
        Case 1
            ' Found at the first position, so return
            ' an empty string.
            TrimNull = ""
    End Select
End Function

Private Function FnPtrToLong(ByVal lngFnPtr As Long) As Long
    ' Given a function pointer AS a Long, return a Long.
    ' Sure looks like this function isn't doing anything,
    ' AND in reality, it's not.
    FnPtrToLong = lngFnPtr
End Function

Private Function CallBackFunction(ByVal hwnd As Long, ByVal lngMsg As Long, _
 ByVal lngParam As Long, ByVal lngData As Long) As Long
    ' This is an example callback function, simply demonstrating
    ' the kinds of things this function can do.
    
    ' The formal declaration of this function is specified by the
    ' SHBrowseForFolder API call, AND cannot be changed.
    
    Dim strFolder As String
    Dim lngIDL As Long
        
    Select Case lngMsg
        Case BFFM_INITIALIZED
            Call EnableOK(hwnd, False)
            Call SetSelection(hwnd, bffInitialSelection)
            
        Case BFFM_SELCHANGED
            strFolder = Space(conMaxPath)
            
            ' Try to resolve the lngIDL into a real path
            If CBool(SHGetPathFromIDList(lngParam, strFolder)) Then
                ' Trim the null characters
                strFolder = TrimNull(strFolder)
                Call SetStatus(hwnd, strFolder)
                Call EnableOK(hwnd, True)
            Else
                ' Not a file system path, so
                ' clear that status, AND disable
                ' the OK button.
                Call SetStatus(hwnd, vbNullString)
                Call EnableOK(hwnd, False)
            End If
    End Select
    CallBackFunction = 0
End Function

Private Sub SetStatus(hwnd As Long, strText As String)
    ' Set the status text on the Shell browser dialog.
    Call SendMessage(hwnd, BFFM_SETSTATUSTEXT, 0, strText)
End Sub

Private Sub EnableOK(hwnd As Long, fEnable As Boolean)
    ' Enable the OK button on the Shell browser dialog.
    Dim lngEnable As Long
    
    If fEnable Then
        lngEnable = 1
    Else
        lngEnable = 0
    End If
    Call SendMessage(hwnd, BFFM_ENABLEOK, 0, lngEnable)
End Sub

Private Function SetSelection(hwnd As Long, varSel As Variant) As Long
    ' Set the selection in the Shell browser dialog. If
    ' varSel is numeric, the code assumes that the value is a
    ' PDIL, AND calls SendMessage with that information. If
    ' not, the code converts the value to a string, AND
    ' tells the browser that it's sending a string value instead.
    If IsNumeric(varSel) Then
        Call SendMessage(hwnd, BFFM_SETSELECTION, 0, CLng(varSel))
    Else
        Call SendMessage(hwnd, BFFM_SETSELECTION, 1, CStr(varSel & ""))
    End If
End Function


Public Property Get APIReturn() As Long
    'return object's APIReturn property
    APIReturn = m_lApiReturn
End Property
Public Property Get ExtendedError() As Long
    'return object's ExtendedError property
    ExtendedError = m_lExtendedError
End Property

#If fComponent Then
Private Sub Class_Initialize()
    InitColors
End Sub
#End If

Function VBGetOpenFileName(FileName As String, Optional FileTitle As String, Optional FileMustExist As Boolean = True, Optional MultiSelect As Boolean = False, Optional ReadOnly As Boolean = False, Optional HideReadOnly As Boolean = False, Optional filter As String = "All (*.*)| *.*", Optional FilterIndex As Long = 1, Optional InitDir As String, Optional DlgTitle As String, Optional DefaultExt As String, Optional Owner As Long = -1, Optional flags As Long = 0) As Boolean

    Dim opfile As OpenFileName, s As String, afFlags As Long
    
    m_lApiReturn = 0
    m_lExtendedError = 0

With opfile
    .lStructSize = Len(opfile)
    
    ' Add in specific flags AND strip out non-VB flags
    
    .flags = (-FileMustExist * OFN_FILEMUSTEXIST) Or (-MultiSelect * OFN_ALLOWMULTISELECT) Or (-ReadOnly * OFN_READONLY) Or (-HideReadOnly * OFN_HIDEREADONLY) Or (flags And CLng(Not (OFN_ENABLEHOOK Or OFN_ENABLETEMPLATE)))
    
    '.flags = .flags Or OFN_NODEREFERENCELINKS
    
    ' Owner can take handle of owning window
    If Owner <> -1 Then .hwndOwner = Owner
    ' InitDir can take initial directory string
    .lpstrInitialDir = InitDir
    ' DefaultExt can take default extension
    .lpstrDefExt = DefaultExt
    ' DlgTitle can take dialog box title
    .lpstrTitle = DlgTitle
    
    ' To make Windows-style filter, replace | AND : with nulls
    Dim ch As String, i As Integer
    For i = 1 To Len(filter)
        ch = Mid$(filter, i, 1)
        If ch = "|" Or ch = ":" Then
            s = s & vbNullChar
        Else
            s = s & ch
        End If
    Next
    ' Put double null at end
    s = s & vbNullChar & vbNullChar
    .lpstrFilter = s
    .nFilterIndex = FilterIndex

    ' Pad file AND file title buffers to maximum path
    s = FileName & String$(MAX_PATH - Len(FileName), 0)
    .lpstrFile = s
    .nMaxFile = MAX_PATH
    s = FileTitle & String$(MAX_FILE - Len(FileTitle), 0)
    .lpstrFileTitle = s
    .nMaxFileTitle = MAX_FILE
    ' All other fields set to zero
    
    m_lApiReturn = GetOpenFileName(opfile)
    
    
    Select Case m_lApiReturn
    Case 1
        ' Success
        VBGetOpenFileName = True
        FileName = StrZToStr(.lpstrFile)
        FileTitle = StrZToStr(.lpstrFileTitle)
        flags = .flags
        ' Return the filter index
        FilterIndex = .nFilterIndex
        ' Look up the filter the user selected AND return that
        filter = FilterLookup(.lpstrFilter, FilterIndex)
        If (.flags And OFN_READONLY) Then ReadOnly = True
    Case 0
        ' Cancelled
        VBGetOpenFileName = False
        FileName = ""
        FileTitle = ""
        flags = 0
        FilterIndex = -1
        filter = ""
    Case Else
        ' Extended error
        m_lExtendedError = CommDlgExtendedError()
        VBGetOpenFileName = False
        FileName = ""
        FileTitle = ""
        flags = 0
        FilterIndex = -1
        filter = ""
    End Select
End With

End Function
Private Function StrZToStr(s As String) As String
    StrZToStr = Left$(s, lstrlen(s))
End Function

Function VBGetSaveFileName(FileName As String, Optional FileTitle As String, Optional OverWritePrompt As Boolean = True, Optional filter As String = "All (*.*)| *.*", Optional FilterIndex As Long = 1, Optional InitDir As String, Optional DlgTitle As String, Optional DefaultExt As String, Optional Owner As Long = -1, Optional flags As Long) As Boolean
            
    Dim opfile As OpenFileName, s As String

    m_lApiReturn = 0
    m_lExtendedError = 0

With opfile
    .lStructSize = Len(opfile)
    
    ' Add in specific flags AND strip out non-VB flags
    .flags = (-OverWritePrompt * OFN_OVERWRITEPROMPT) Or OFN_HIDEREADONLY Or (flags And CLng(Not (OFN_ENABLEHOOK Or OFN_ENABLETEMPLATE)))
    ' Owner can take handle of owning window
    If Owner <> -1 Then .hwndOwner = Owner
    ' InitDir can take initial directory string
    .lpstrInitialDir = InitDir
    ' DefaultExt can take default extension
    .lpstrDefExt = DefaultExt
    ' DlgTitle can take dialog box title
    .lpstrTitle = DlgTitle
    
    ' Make new filter with bars (|) replacing nulls AND double null at end
    Dim ch As String, i As Integer
    For i = 1 To Len(filter)
        ch = Mid$(filter, i, 1)
        If ch = "|" Or ch = ":" Then
            s = s & vbNullChar
        Else
            s = s & ch
        End If
    Next
    ' Put double null at end
    s = s & vbNullChar & vbNullChar
    .lpstrFilter = s
    .nFilterIndex = FilterIndex

    ' Pad file AND file title buffers to maximum path
    s = FileName & String$(MAX_PATH - Len(FileName), 0)
    .lpstrFile = s
    .nMaxFile = MAX_PATH
    s = FileTitle & String$(MAX_FILE - Len(FileTitle), 0)
    .lpstrFileTitle = s
    .nMaxFileTitle = MAX_FILE
    ' All other fields zero
    
    m_lApiReturn = GetSaveFileName(opfile)
    Select Case m_lApiReturn
    Case 1
        VBGetSaveFileName = True
        FileName = StrZToStr(.lpstrFile)
        FileTitle = StrZToStr(.lpstrFileTitle)
        flags = .flags
        ' Return the filter index
        FilterIndex = .nFilterIndex
        ' Look up the filter the user selected AND return that
        filter = FilterLookup(.lpstrFilter, FilterIndex)
    Case 0
        ' Cancelled:
        VBGetSaveFileName = False
        FileName = ""
        FileTitle = ""
        flags = 0
        FilterIndex = 0
        filter = ""
    Case Else
        ' Extended error:
        VBGetSaveFileName = False
        m_lExtendedError = CommDlgExtendedError()
        FileName = ""
        FileTitle = ""
        flags = 0
        FilterIndex = 0
        filter = ""
    End Select
End With
End Function

Private Function FilterLookup(ByVal sFilters As String, ByVal iCur As Long) As String
    Dim iStart As Long, iEnd As Long, s As String
    iStart = 1
    If sFilters = "" Then Exit Function
    Do
        ' Cut out both parts marked by null character
        iEnd = InStr(iStart, sFilters, vbNullChar)
        If iEnd = 0 Then Exit Function
        iEnd = InStr(iEnd + 1, sFilters, vbNullChar)
        If iEnd Then
            s = Mid$(sFilters, iStart, iEnd - iStart)
        Else
            s = Mid$(sFilters, iStart)
        End If
        iStart = iEnd + 1
        If iCur = 1 Then
            FilterLookup = s
            Exit Function
        End If
        iCur = iCur - 1
    Loop While iCur
End Function

Function VBGetFileTitle(sFile As String) As String
    Dim sFileTitle As String, cFileTitle As Integer

    cFileTitle = MAX_PATH
    sFileTitle = String$(MAX_PATH, 0)
    cFileTitle = GetFileTitle(sFile, sFileTitle, MAX_PATH)
    If cFileTitle Then
        VBGetFileTitle = ""
    Else
        VBGetFileTitle = Left$(sFileTitle, InStr(sFileTitle, vbNullChar) - 1)
    End If

End Function

' ChooseColor wrapper
Function VBChooseColor(color As Long, Optional AnyColor As Boolean = True, Optional FullOpen As Boolean = False, Optional DisableFullOpen As Boolean = False, Optional Owner As Long = -1, Optional flags As Long) As Boolean

    Dim chclr As TCHOOSECOLOR
    chclr.lStructSize = Len(chclr)
    
    ' Color must get reference variable to receive result
    ' Flags can get reference variable or constant with bit flags
    ' Owner can take handle of owning window
    If Owner <> -1 Then chclr.hwndOwner = Owner

    ' Assign color (default uninitialized Value of zero is good default)
    chclr.rgbResult = color

    ' Mask out unwanted bits
    Dim afMask As Long
    afMask = CLng(Not (CC_ENABLEHOOK Or CC_ENABLETEMPLATE))
    ' Pass in flags
    chclr.flags = afMask And (CC_RGBInit Or IIf(AnyColor, CC_AnyColor, CC_SolidColor) Or (-FullOpen * CC_FullOpen) Or (-DisableFullOpen * CC_PreventFullOpen))

    ' If first time, initialize to white
    If fNotFirst = False Then InitColors

    chclr.lpCustColors = VarPtr(alCustom(0))
    ' All other fields zero
    
    m_lApiReturn = ChooseColor(chclr)
    Select Case m_lApiReturn
    Case 1
        ' Success
        VBChooseColor = True
        color = chclr.rgbResult
    Case 0
        ' Cancelled
        VBChooseColor = False
        color = -1
    Case Else
        ' Extended error
        m_lExtendedError = CommDlgExtendedError()
        VBChooseColor = False
        color = -1
    End Select

End Function

Private Sub InitColors()
    Dim i As Integer
    ' Initialize with first 16 system interface colors
    For i = 0 To 15
        alCustom(i) = GetSysColor(i)
    Next
    fNotFirst = True
End Sub

' Property to read or modify custom colors (use to save colors in registry)
Public Property Get CustomColor(i As Integer) As Long
    ' If first time, initialize to white
    If fNotFirst = False Then InitColors
    If i >= 0 And i <= 15 Then
        CustomColor = alCustom(i)
    Else
        CustomColor = -1
    End If
End Property

Public Property Let CustomColor(i As Integer, iValue As Long)
    ' If first time, initialize to system colors
    If fNotFirst = False Then InitColors
    If i >= 0 And i <= 15 Then
        alCustom(i) = iValue
    End If
End Property

' ChooseFont wrapper
Function VBChooseFont(CurFont As Font, Optional PrinterDC As Long = -1, Optional Owner As Long = -1, Optional color As Long = vbBlack, Optional MinSize As Long = 0, Optional MaxSize As Long = 0, Optional flags As Long = 0) As Boolean

    m_lApiReturn = 0
    m_lExtendedError = 0

    ' Unwanted Flags bits
    Const CF_FontNotSupported = CF_Apply Or CF_EnableHook Or CF_EnableTemplate
    
    ' Flags can get reference variable or constant with bit flags
    ' PrinterDC can take printer DC
    If PrinterDC = -1 Then
        PrinterDC = 0
        If flags And CF_PrinterFonts Then PrinterDC = Printer.hDC
    Else
        flags = flags Or CF_PrinterFonts
    End If
    ' Must have some fonts
    If (flags And CF_PrinterFonts) = 0 Then flags = flags Or CF_ScreenFonts
    ' Color can take initial color, receive chosen color
    If color <> vbBlack Then flags = flags Or CF_EFFECTS
    ' MinSize can be minimum size accepted
    If MinSize Then flags = flags Or CF_LimitSize
    ' MaxSize can be maximum size accepted
    If MaxSize Then flags = flags Or CF_LimitSize

    ' Put in required internal flags AND remove unsupported
    flags = (flags Or CF_InitToLogFontStruct) And Not CF_FontNotSupported
    
    ' Initialize LOGFONT variable
    Dim fnt As LOGFONT
    Const PointsPerTwip = 1440 / 72
    fnt.lfHeight = -(CurFont.Size * (PointsPerTwip / Screen.TwipsPerPixelY))
    fnt.lfWeight = CurFont.Weight
    fnt.lfItalic = CurFont.Italic
    fnt.lfUnderline = CurFont.Underline
    fnt.lfStrikeOut = CurFont.Strikethrough
    ' Other fields zero
    StrToBytes fnt.lfFaceName, CurFont.Name

    ' Initialize TCHOOSEFONT variable
    Dim cf As TCHOOSEFONT
    cf.lStructSize = Len(cf)
    If Owner <> -1 Then cf.hwndOwner = Owner
    cf.hDC = PrinterDC
    cf.lpLogFont = VarPtr(fnt)
    cf.iPointSize = CurFont.Size * 10
    cf.flags = flags
    cf.rgbColors = color
    cf.nSizeMin = MinSize
    cf.nSizeMax = MaxSize
    
    ' All other fields zero
    m_lApiReturn = ChooseFont(cf)
    Select Case m_lApiReturn
    Case 1
        ' Success
        VBChooseFont = True
        flags = cf.flags
        color = cf.rgbColors
        CurFont.Bold = cf.nFontType And Bold_FontType
        'CurFont.Italic = cf.nFontType AND Italic_FontType
        CurFont.Italic = fnt.lfItalic
        CurFont.Strikethrough = fnt.lfStrikeOut
        CurFont.Underline = fnt.lfUnderline
        CurFont.Weight = fnt.lfWeight
        CurFont.Size = cf.iPointSize / 10
        CurFont.Name = BytesToStr(fnt.lfFaceName)
    Case 0
        ' Cancelled
        VBChooseFont = False
    Case Else
        ' Extended error
        m_lExtendedError = CommDlgExtendedError()
        VBChooseFont = False
    End Select
        
End Function

' PrintDlg wrapper
Function VBPrintDlg(hDC As Long, Optional PrintRange As EPrintRange = eprAll, Optional DisablePageNumbers As Boolean, Optional FromPage As Long = 1, _
                    Optional Topage As Long = &HFFFF, Optional DisableSelection As Boolean, Optional Copies As Integer, Optional ShowPrintToFile As Boolean, _
                    Optional DisablePrintToFile As Boolean = True, Optional PrintToFile As Boolean, Optional Collate As Boolean, _
                    Optional PreventWarning As Boolean, Optional OwnerHWnd As Long, Optional flags As Long, Optional PrinterName As String) As Boolean
                    
    Dim afFlags As Long, afMask As Long
    
    m_lApiReturn = 0
    m_lExtendedError = 0
    
    ' Set PRINTDLG flags
    afFlags = (-DisablePageNumbers * PD_NOPAGENUMS) Or (-DisablePrintToFile * PD_DISABLEPRINTTOFILE) Or (-DisableSelection * PD_NOSELECTION) Or (-PrintToFile * PD_PRINTTOFILE) Or (-(Not ShowPrintToFile) * PD_HIDEPRINTTOFILE) Or (-PreventWarning * PD_NOWARNING) Or (-Collate * PD_COLLATE) Or PD_USEDEVMODECOPIESANDCOLLATE Or PD_RETURNDC
    If PrintRange = eprPageNumbers Then
        afFlags = afFlags Or PD_PAGENUMS
    ElseIf PrintRange = eprSelection Then
        afFlags = afFlags Or PD_SELECTION
    End If
    ' Mask out unwanted bits
    afMask = CLng(Not (PD_ENABLEPRINTHOOK Or PD_ENABLEPRINTTEMPLATE))
    afMask = afMask And CLng(Not (PD_ENABLESETUPHOOK Or PD_ENABLESETUPTEMPLATE))
    
    ' Fill in PRINTDLG structure
    Dim pd As TPRINTDLG
    pd.lStructSize = Len(pd)
    pd.hwndOwner = OwnerHWnd
    pd.flags = afFlags And afMask
    pd.nFromPage = FromPage
    pd.nToPage = Topage
    pd.nMinPage = 1
    pd.nMaxPage = &HFFFF
    
    ' Show Print dialog
    m_lApiReturn = PrintDlg(pd)
    Select Case m_lApiReturn
    Case 1
        VBPrintDlg = True
        ' Return dialog Values in parameters
        hDC = pd.hDC
        
        If (pd.flags And PD_PAGENUMS) Then
            PrintRange = eprPageNumbers
        ElseIf (pd.flags And PD_SELECTION) Then
            PrintRange = eprSelection
        Else
            PrintRange = eprAll
        End If
        FromPage = pd.nFromPage
        Topage = pd.nToPage
        PrintToFile = (pd.flags And PD_PRINTTOFILE)
        ' Get DEVMODE structure FROM PRINTDLG
        Dim pDevMode As Long
        pDevMode = GlobalLock(pd.hDevMode)
        CopyMemory m_dvmode, ByVal pDevMode, Len(m_dvmode)
        Call GlobalUnlock(pd.hDevMode)
        ' Get Copies AND Collate settings FROM DEVMODE structure
        PrinterName = m_dvmode.dmDeviceName
        Copies = m_dvmode.dmCopies
        Collate = (m_dvmode.dmCollate = DMCOLLATE_TRUE)
                
    Case 0
        ' Cancelled
        VBPrintDlg = False
    Case Else
        ' Extended error:
        m_lExtendedError = CommDlgExtendedError()
        VBPrintDlg = False
    End Select
    
End Function
Private Property Get DevMode() As DevMode
    DevMode = m_dvmode
End Property

' PageSetupDlg wrapper
Function VBPageSetupDlg(Optional Owner As Long, Optional DisableMargins As Boolean, Optional DisableOrientation As Boolean, Optional DisablePaper As Boolean, Optional DisablePrinter As Boolean, Optional LeftMargin As Long, Optional MinLeftMargin As Long, Optional RightMargin As Long, Optional MinRightMargin As Long, Optional TopMargin As Long, Optional MinTopMargin As Long, Optional BottomMargin As Long, Optional MinBottomMargin As Long, Optional PaperSize As EPaperSize = epsLetter, Optional Orientation As EOrientation = eoPortrait, Optional PrintQuality As EPrintQuality = epqDraft, Optional Units As EPageSetupUnits = epsuInches, Optional Printer As Object, Optional flags As Long) As Boolean
    Dim afFlags As Long, afMask As Long
        
    m_lApiReturn = 0
    m_lExtendedError = 0
    ' Mask out unwanted bits
    afMask = Not (PSD_EnablePagePaintHook Or _
                  PSD_EnablePageSetupHook Or _
                  PSD_EnablePageSetupTemplate)
    ' Set TPAGESETUPDLG flags
    afFlags = (-DisableMargins * PSD_DISABLEMARGINS) Or _
              (-DisableOrientation * PSD_DISABLEORIENTATION) Or _
              (-DisablePaper * PSD_DISABLEPAPER) Or _
              (-DisablePrinter * PSD_DISABLEPRINTER) Or _
              PSD_MARGINS Or PSD_MINMARGINS And afMask
    Dim lUnits As Long
    If Units = epsuInches Then
        afFlags = afFlags Or PSD_INTHOUSANDTHSOFINCHES
        lUnits = 1000
    Else
        afFlags = afFlags Or PSD_INHUNDREDTHSOFMILLIMETERS
        lUnits = 100
    End If
    
    Dim psd As TPAGESETUPDLG
    ' Fill in PRINTDLG structure
    psd.lStructSize = Len(psd)
    psd.hwndOwner = Owner
    psd.rtMargin.Top = TopMargin * lUnits
    psd.rtMargin.Left = LeftMargin * lUnits
    psd.rtMargin.Bottom = BottomMargin * lUnits
    psd.rtMargin.Right = RightMargin * lUnits
    psd.rtMinMargin.Top = MinTopMargin * lUnits
    psd.rtMinMargin.Left = MinLeftMargin * lUnits
    psd.rtMinMargin.Bottom = MinBottomMargin * lUnits
    psd.rtMinMargin.Right = MinRightMargin * lUnits
    psd.flags = afFlags
    
    ' Show Print dialog
    If PageSetupDlg(psd) Then
        VBPageSetupDlg = True
        ' Return dialog Values in parameters
        TopMargin = psd.rtMargin.Top / lUnits
        LeftMargin = psd.rtMargin.Left / lUnits
        BottomMargin = psd.rtMargin.Bottom / lUnits
        RightMargin = psd.rtMargin.Right / lUnits
        MinTopMargin = psd.rtMinMargin.Top / lUnits
        MinLeftMargin = psd.rtMinMargin.Left / lUnits
        MinBottomMargin = psd.rtMinMargin.Bottom / lUnits
        MinRightMargin = psd.rtMinMargin.Right / lUnits
        
        ' Get DEVMODE structure FROM PRINTDLG
        Dim dvmode As DevMode, pDevMode As Long
        pDevMode = GlobalLock(psd.hDevMode)
        CopyMemory dvmode, ByVal pDevMode, Len(dvmode)
        Call GlobalUnlock(psd.hDevMode)
        PaperSize = dvmode.dmPaperSize
        Orientation = dvmode.dmOrientation
        PrintQuality = dvmode.dmPrintQuality
        ' Set default printer properties
        On Error Resume Next
        If Not (Printer Is Nothing) Then
            Printer.Copies = dvmode.dmCopies
            Printer.Orientation = dvmode.dmOrientation
            Printer.PaperSize = dvmode.dmPaperSize
            Printer.PrintQuality = dvmode.dmPrintQuality
        End If
        On Error GoTo 0
    End If

End Function

#If fComponent = 0 Then
Private Sub ErrRaise(E As Long)
    Dim sText As String, sSource As String
    If E > 1000 Then
        sSource = App.EXEName & ".CommonDialog"
        Err.Raise COMError(E), sSource, sText
    Else
        ' Raise standard Visual Basic error
        sSource = App.EXEName & ".VBError"
        Err.Raise E, sSource
    End If
End Sub
#End If


Private Sub StrToBytes(ab() As Byte, s As String)
    If IsArrayEmpty(ab) Then
        ' Assign to empty array
        ab = StrConv(s, vbFromUnicode)
    Else
        Dim cab As Long
        ' Copy to existing array, padding or truncating if necessary
        cab = UBound(ab) - LBound(ab) + 1
        If Len(s) < cab Then s = s & String$(cab - Len(s), 0)
        'If UnicodeTypeLib Then
        '    Dim st AS String
        '    st = StrConv(s, vbFromUnicode)
        '    CopyMemoryStr ab(LBound(ab)), st, cab
        'Else
            CopyMemoryStr ab(LBound(ab)), s, cab
        'End If
    End If
End Sub


Private Function BytesToStr(ab() As Byte) As String
    BytesToStr = StrConv(ab, vbUnicode)
End Function

Private Function COMError(E As Long) As Long
    COMError = E Or vbObjectError
End Function
'
Private Function IsArrayEmpty(va As Variant) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = va(LBound(va))
    IsArrayEmpty = (Err <> 0)
End Function




