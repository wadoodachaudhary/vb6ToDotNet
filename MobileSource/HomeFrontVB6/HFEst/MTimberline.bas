Attribute VB_Name = "MTimberline"
Option Explicit

Public Enum TLNameModes
    CustomNames
    StandardNames
    DictionaryNames
End Enum

Public Enum TLDatabaseTypes
    Accounting = 1
    Estimating = 2
End Enum

Public Function TimberlineInstalled() As Boolean
'    TimberlineInstalled = "" <> RegGetKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Timberline\General", "System Directory")
End Function

Public Function TimberlineSecured() As Boolean
On Error Resume Next
    Dim c As New ADODB.Connection
    c.Open "Provider=MSDASQL.1;DRIVER={Timberline Data};"
    If Err.Number = -2147467259 Then
        '[Timberline][ODBC Driver][DLL]Invalid account name.
        TimberlineSecured = True
    End If
End Function

Public Function SelectedDBQ(c As Combobox) As String
    If c.ListIndex >= 0 Then
        SelectedDBQ = Parse(c.Tag, c.ListIndex + 1, Chr(0))
    End If
End Function

Public Sub LoadDBQs(Optional List, Optional InitialDBQ = "")
    Dim dbqList As String
    Dim dbqItem As String
    Dim Name As String
    Dim Folder As String
    Dim folderlist As String
    Dim i As Long
    
    dbqList = GetDBQs
    If TypeName(List) = "VSFlexGrid" Then
        List.Rows = 1
        For i = 1 To Parse(dbqList, , Chr(1))
            dbqItem = Parse(dbqList, i, Chr(1))
            Name = Parse(dbqItem, 1, Chr(2))
            Folder = Parse(dbqItem, 2, Chr(2))
            List.AddItem Name & vbTab & Folder
        Next
        On Error Resume Next
        Call List.AutoSize(0, 1)
        List.Col = 0
        List.Sort = 7 'flexSortStringAscending
        List.Row = List.FindRow(InitialDBQ, , 1, False, True)
        Call List.ShowCell(List.Row, 0)
    Else
        List.Clear
        For i = 1 To Parse(dbqList, , Chr(1))
            dbqItem = Parse(dbqList, i, Chr(1))
            Name = Parse(dbqItem, 1, Chr(2))
            Folder = Parse(dbqItem, 2, Chr(2))
            List.AddItem Name
            List.ItemData(List.NewIndex) = i
            folderlist = folderlist & Chr(0) & Folder
        Next
        List.Tag = Mid(folderlist, 2)
        
        On Error Resume Next
        For i = 0 To List.ListCount - 1
            If InitialDBQ = Parse(List.Tag, List.ItemData(i), Chr(0)) Then
                List.ListIndex = i
            End If
        Next
        
        If List.Index < 0 And List.ListCount > 0 Then List.ListIndex = 0
    End If

End Sub

Public Function TimberlineSysDir() As String
    TimberlineSysDir = RegGetKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Timberline\General", "System Directory")
End Function

Public Function GetDBQs() As String
On Error Resume Next

    Dim iTsDir   As Integer
    Dim sTsDir   As String
    Dim CurDrive As String
    
    Dim sFolder  As String
    Dim iTsCtl   As Integer
    Dim sTsCtl   As String
    Dim sCompany As String
    Dim sConnect As String
    
    Dim List As String
    Dim tmp  As String
    Dim i As Long
    
    'open tl.dir and read data folder paths
    sTsDir = TimberlineSysDir
    If sTsDir = "" Then Exit Function
    sTsDir = sTsDir & "\misc\tl.dir"
    CurDrive = FileDrive(sTsDir)
    
    iTsDir = FreeFile
    Open sTsDir For Input Access Read As iTsDir
    While Not EOF(iTsDir)
        Line Input #iTsDir, sFolder
        sFolder = CurDrive & sFolder
        
        'open tl.ctl read company name
        sTsCtl = sFolder & "\tl.ctl"
        iTsCtl = FreeFile
        Open sTsCtl For Binary Access Read As iTsCtl
        If Err.Number = 0 Then
            sCompany = DNull(Trim(Mid(input(50, iTsCtl), 17)))
            List = List & Chr(1) & sCompany & Chr(2) & sFolder
        End If
        Close iTsCtl
    
    Wend
    Close iTsDir
    List = Mid(List, 2)
    
    'remove any with no name
    tmp = ""
    For i = 1 To Parse(List, , Chr(1))
        If "" <> Parse(Parse(List, i, Chr(1)), 1, Chr(2)) Then
            tmp = tmp & Chr(1) & Parse(List, i, Chr(1))
        End If
    Next
    List = Mid(tmp, 2)
    
    GetDBQs = List

End Function

Public Sub DONTUSEMETsObject(DataFolder As String, FileName As String, UserID As String, Pswd As String)
    Dim TSobjPath     As String
    Dim BatchFileName As String
    Dim BatchFile     As Integer

    
    TSobjPath = RegGetKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Timberline\General", "System Directory", "c:\program files\timberline office\accounting\")
    TSobjPath = PathAppend(TSobjPath, "tsobject.exe")
    BatchFileName = AppWorkingFolder & "\" & App.EXEName & ".bat"
    BatchFile = FreeFile
    
    Open BatchFileName For Output As #BatchFile
    Print #BatchFile, FileDrive(DataFolder)
    Print #BatchFile, "cd " & vbQuote & DataFolder & vbQuote
    Print #BatchFile, vbQuote & TSobjPath & vbQuote & " " & vbQuote & FileName & vbQuote & " " & UserID & " " & Pswd
    Print #BatchFile, "cd \"
    Print #BatchFile, "c:"
    Close BatchFile
    
    
    Call Shell(BatchFileName, vbHide)
    
End Sub

Public Function FormatTSField(Length As Long, value As String, Optional RightAlignStrings As Boolean, Optional RemovePunctuation As Boolean) As String
    
    If RemovePunctuation Then
        value = Replace(value, ":", "")
        value = Replace(value, ";", "")
        value = Replace(value, "-", "")
        value = Replace(value, ".", "")
        value = Replace(value, "_", "")
    End If
    
    If TsIsNumeric(value) Then
        FormatTSField = format(value, String(Length, "@"))
    Else
        If RightAlignStrings Then
            FormatTSField = format(value, String(Length, "@"))
        Else
            FormatTSField = Trim(value)
        End If
    End If
End Function


Public Function TsIsNumeric(value As String) As String
    Dim i As Long
    Dim b As Boolean
    b = True
    For i = 1 To Len(value)
        If Mid(value, i, 1) <> "." Then
            b = b And IsNumeric(Mid(value, i, 1))
        End If
    Next
    TsIsNumeric = b
End Function



Public Function TsConnectionString(DataFolder As String, Optional User As String, Optional Pswd As String, Optional NamingMode As TLNameModes = StandardNames, Optional DatabaseType As TLDatabaseTypes = Accounting) As String
If DatabaseType = Estimating Then
    TsConnectionString = "DSN=Timberline Data Source;" & _
                         "DBQ=" & DataFolder & ";" & _
                         "UID=" & User & ";" & _
                         "PWD=" & Pswd & ";" & _
                         "CustomMode=0" & ";" & _
                         "StandardMode=1" & ";" & _
                         "DictionaryMode=0" & ";" & _
                         "KeepFilesOpen=0" & ";" & _
                         "MaxColSupport=255;" & _
                         "CODEPAGE=1252;" & _
                         "DatabaseType=2;" & _
                         "ShortenNames=0;"
Else
    TsConnectionString = "DRIVER={Timberline Data};" & _
                         "DBQ=" & DataFolder & ";" & _
                         "UID=" & User & ";" & _
                         "PWD=" & Pswd & ";" & _
                         "CustomMode=" & IIf(NamingMode = CustomNames, 1, 0) & ";" & _
                         "StandardMode=" & IIf(NamingMode = StandardNames, 1, 0) & ";" & _
                         "DictionaryMode=" & IIf(NamingMode = DictionaryNames, 1, 0) & ";" & _
                         "KeepFilesOpen=0" & ";" & _
                         "MaxColSupport=255;" & _
                         "CODEPAGE=1252;" & _
                         "DatabaseType=" & DatabaseType & ";" & _
                         "ShortenNames=0;"
End If
End Function

Public Sub PECOMErrHandler(Source As String, Optional sql As String, Optional FileName As String)
    Dim s As String
    Dim ispee As Boolean
    ispee = FileExt(FileName) = "pee"
    
    Select Case True
        Case ispee And (Not PathExists(HFApp.Options(TimberlineEstimatingPath)))
            s = "    Your Timberline Estimating database path:" & vbCrLf & _
                "    " & vbQuote & HFApp.Options(TimberlineEstimatingPath) & vbQuote & vbCrLf & _
                "    does not exist or is not accessible."
                
        Case (Not ispee) And (Not PathExists(FileName))
            s = "    Your Timberline Estimating database path:" & vbCrLf & _
                "    " & vbQuote & FileName & vbQuote & vbCrLf & _
                "    does not exist or is not accessible."
                
        Case Not PathExists(HFApp.Options(TimberlineEstimatingDictionaryPath))
            s = "    Your Timberline Estimating dictionary path:" & vbCrLf & _
                "    " & vbQuote & HFApp.Options(TimberlineEstimatingDictionaryPath) & vbQuote & vbCrLf & _
                "    does not exist or is not accessible."
                   
        Case Err.Number = 9999
            s = "    This application is licensed for __ users. Currently all licenses are being used. No additional operators can access the software at this time. Try again later or call your Timberline solution provider to allow additional operators access."
                
        Case Err.Number = -2147219497 And ispee
            s = "    The database associated with this estimate has been moved or" & vbCrLf & _
                "    deleted. For Precision Builder to access this estimate it must" & vbCrLf & _
                "    have a valid database. Open the estimate in Timberline Estimating" & vbCrLf & _
                "     andchange the database." & vbCrLf & vbCrLf & _
                "    " & FileName
                              
        Case Err.Number = -2147219497 And Not ispee
            s = "    The Timberline Estimating database is invalid or corrupt. Use" & vbCrLf & _
                "    Estimating Tools, File Menu > Rebuild File to restore the integrity" & vbCrLf & _
                "    of the database." & vbCrLf & vbCrLf & "    " & FileName
                              
        Case Err.Number = -2147467259
            s = "    This file is in use or has been marked as read-only." & vbCrLf & vbCrLf & _
                "    " & FileName
        
        Case Err.Number = -2147220488
            s = "    The Timberline Estimating database is corrupt. Use Estimating Tools" & vbCrLf & _
                "    File Menu > Rebuild File to restore the integrity of the database."
                              
        Case Err.Number = -2147219444
            s = "    This assembly contains a select at takeoff item with no default item." & vbCrLf & _
                "    It cannot be used by Precision Builder until a default value is given." & vbCrLf & _
                "    Please open Timberline Estimating and set the default value to item " & HFApp.Options(SelectAtTakeoffPhase) & "/" & HFApp.Options(SelectAtTakeoffItem)
                
        Case Else
            MsgBox "Precision Builder is not able to connect to your estimating database. Please check" & vbCrLf & _
                   "your system settings or contact your system administrator for assistance." & vbCrLf & vbCrLf & _
                   FileName & vbCrLf & vbCrLf & Err.Description, vbExclamation, App.ProductName
        
        'Case -2147211456: 'estimate file corrupt data
        'Case -2147220486: 'estimate file is in use
        
    End Select
    
    If s <> "" Then
        Screen.MousePointer = vbDefault
        MsgBox "Precision Builder encountered an error when accessing Timberline Estimating." & vbCrLf & vbCrLf & _
               "Timberline error code:" & vbCrLf & _
               s & vbCrLf & vbCrLf & _
               "Please contact your system administrator for assistance.", vbExclamation, App.ProductName
    End If
End Sub



