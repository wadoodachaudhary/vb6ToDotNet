Attribute VB_Name = "MPrinterEnum"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2011 VBnet/Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Type PRINTER_INFO_1
   Flags As Long
   pDescription As Long
   Pane As Long
   Comment As Long
End Type

Public Type PRINTER_INFO_4
   pPrinterName As Long
   pServerName As Long
   Attributes As Long
End Type

'SIZEOFxxx are non-windows constants defined for this method
Public Const SIZEOFPRINTER_INFO_1 = 16
Public Const SIZEOFPRINTER_INFO_4 = 12

Public Const PRINTER_LEVEL1 = &H1
Public Const PRINTER_LEVEL4 = &H4

'EnumPrinters enumerates available printers,
'print servers, domains, or print providers.
Public Declare Function EnumPPrinters Lib "winspool.drv" _
   Alias "EnumPrintersA" _
  (ByVal Flags As Long, _
   ByVal Name As String, _
   ByVal Level As Long, _
   pPrinterEnum As Any, _
   ByVal cbBuffer As Long, _
   pcbNeeded As Long, _
   pcReturned As Long) As Long

'EnumPrinters Parameters:
'Flags - Specifies the types of print objects that the function should enumerate.
Public Const PRINTER_ENUM_DEFAULT = &H1     'Windows 95: The function returns
                                            'information about the default printer.
Public Const PRINTER_ENUM_LOCAL = &H2       'function ignores the Name parameter,
                                            'and enumerates the locally installed
                                            'printers. Windows 95: The function will
                                            'also enumerate network printers because
                                            'they are handled by the local print provider
Public Const PRINTER_ENUM_CONNECTIONS = &H4 'Windows NT/2000: The function enumerates the
                                            'list of printers to which the user has made
                                            'previous connections
Public Const PRINTER_ENUM_NAME = &H8        'enumerates the printer identified by Name.
                                            'This can be a server, a domain, or a print
                                            'provider. If Name is NULL, the function
                                            'enumerates available print providers
Public Const PRINTER_ENUM_REMOTE = &H10     'Windows NT/2000: The function enumerates network
                                            'printers and print servers in the computer's domain.
                                            'This value is valid only if Level is 1
Public Const PRINTER_ENUM_SHARED = &H20     'enumerates printers that have the shared attribute.
                                            'Cannot be used in isolation; use an OR operation
                                            'to combine with another PRINTER_ENUM type
Public Const PRINTER_ENUM_NETWORK = &H40    'Windows NT/2000: The function enumerates network
                                            'printers in the computer's domain. This value is
                                            'valid only if Level is 1.

'''''''''''''''''''''''
'Name:
'If Level is 1, Flags contains PRINTER_ENUM_NAME, and Name is non-NULL,
'then Name is a pointer to a null-terminated string that specifies the
'name of the object to enumerate. This string can be the name of a server,
'a domain, or a print provider.
'
'If Level is 1, Flags contains PRINTER_ENUM_NAME, and Name is NULL, then
'the function enumerates the available print providers.
'
'If Level is 1, Flags contains PRINTER_ENUM_REMOTE, and Name is NULL, then
'the function enumerates the printers in the user's domain.
'
'If Level is 2 or 5, Name is a pointer to a null-terminated string that
'specifies the name of a server whose printers are to be enumerated. If
'this string is NULL, then the function enumerates the printers installed
'on the local machine.
'
'If Level is 4, Name should be NULL. The function always queries on
'the local machine.

'When Name is NULL, it enumerates printers that are installed on the
'local machine. These printers include those that are physically attached
'to the local machine as well as remote printers to which it has a
'network connection.

'''''''''''''''''''''''
'Level:
'Specifies the type of data structures pointed to by pPrinterEnum.
'Valid values are 1, 2, 4, and 5, which correspond to the
'PRINTER_INFO_1, PRINTER_INFO_2, PRINTER_INFO_4, and PRINTER_INFO_5
'data structures.
'
'Windows 95: The value can be 1, 2, or 5.
'
'Windows NT/Windows 2000: This value can be 1, 2, 4, or 5.

'''''''''''''''''''''''
'pPrinterEnum:
'Pointer to a buffer that receives an array of PRINTER_INFO_1,
'PRINTER_INFO_2, PRINTER_INFO_4, or PRINTER_INFO_5 structures.
'Each structure contains data that describes an available print object.
'
'If Level is 1, the array contains PRINTER_INFO_1 structures.
'If Level is 2, the array contains PRINTER_INFO_2 structures.
'If Level is 4, the array contains PRINTER_INFO_4 structures.
'If Level is 5, the array contains PRINTER_INFO_5 structures.
'
'The buffer must be large enough to receive the array of data
'structures and any strings or other data to which the structure
'members point. If the buffer is too small, the pcbNeeded parameter
'returns the required buffer size.
'
'Windows 95: The buffer cannot receive PRINTER_INFO_4 structures.
'It can receive any of the other types.

'''''''''''''''''''''''
'cbBuf
'Specifies the size, in bytes, of the buffer pointed to by pPrinterEnum.
'''''''''''''''''''''''
'pcbNeeded
'Pointer to a value that receives the number of bytes copied if the
'function succeeds or the number of bytes required if cbBuffer is too small.
'''''''''''''''''''''''
'pcReturned
'Pointer to a value that receives the number of PRINTER_INFO_1,
'PRINTER_INFO_2, PRINTER_INFO_4, or PRINTER_INFO_5 structures that
'the function returns in the array to which pPrinterEnum points.


'PRINTER_INFO_4 returned Attribute values
Public Const PRINTER_ATTRIBUTE_DEFAULT = &H4
Public Const PRINTER_ATTRIBUTE_DIRECT = &H2
Public Const PRINTER_ATTRIBUTE_ENABLE_BIDI = &H800&
Public Const PRINTER_ATTRIBUTE_LOCAL = &H40
Public Const PRINTER_ATTRIBUTE_NETWORK = &H10
Public Const PRINTER_ATTRIBUTE_QUEUED = &H1
Public Const PRINTER_ATTRIBUTE_SHARED = &H8
Public Const PRINTER_ATTRIBUTE_WORK_OFFLINE = &H400

'PRINTER_INFO_1 returned Flag values
Public Const PRINTER_ENUM_CONTAINER = &H8000&
Public Const PRINTER_ENUM_EXPAND = &H4000
Public Const PRINTER_ENUM_ICON1 = &H10000
Public Const PRINTER_ENUM_ICON2 = &H20000
Public Const PRINTER_ENUM_ICON3 = &H40000
Public Const PRINTER_ENUM_ICON4 = &H80000
Public Const PRINTER_ENUM_ICON5 = &H100000
Public Const PRINTER_ENUM_ICON6 = &H200000
Public Const PRINTER_ENUM_ICON7 = &H400000
Public Const PRINTER_ENUM_ICON8 = &H800000

Public Const LB_SETTABSTOPS As Long = &H192

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
  (ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long
   
Public Declare Function lstrcpyA Lib "kernel32" _
  (ByVal RetVal As String, ByVal Ptr As Long) As Long
                        
Public Declare Function lstrlenA Lib "kernel32" _
  (ByVal Ptr As Any) As Long


Private Function GetStrFromPtrA(ByVal lpszA As Long) As String

   GetStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
   Call lstrcpyA(ByVal GetStrFromPtrA, ByVal lpszA)
   
End Function


Public Function EnumPrinters() As String
    Dim s As String
    Dim Success As Boolean
    Dim cbRequired As Long
    Dim cbBuffer As Long
    Dim Ptr() As PRINTER_INFO_4
    Dim nEntries As Long
    Dim cnt As Long
    Dim sAttr As String
   
    Call EnumPPrinters(PRINTER_ENUM_CONNECTIONS Or PRINTER_ENUM_LOCAL, vbNullString, PRINTER_LEVEL4, 0, 0, cbRequired, nEntries)
   
    ReDim Ptr((cbRequired \ SIZEOFPRINTER_INFO_4))
    cbBuffer = cbRequired
    
    If EnumPPrinters(PRINTER_ENUM_CONNECTIONS Or PRINTER_ENUM_LOCAL, vbNullString, PRINTER_LEVEL4, Ptr(0), cbBuffer, cbRequired, nEntries) Then
        For cnt = 0 To nEntries - 1
            With Ptr(cnt)
                s = s & "," & GetStrFromPtrA(.pPrinterName)
            End With
        Next
    End If
   
    EnumPrinters = Mid(s, 2)
    
End Function



