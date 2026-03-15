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
    TimberlineInstalled = "" <> RegGetKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Timberline\General", "System Directory")
End Function

Public Function TimberlineSecured() As Boolean
On Error Resume Next
    Dim c As New adodb.Connection
    c.Open "Provider=MSDASQL.1;DRIVER={Timberline Data};"
    If Err.Number = -2147467259 Then
        '[Timberline][ODBC Driver][DLL]Invalid account name.
        TimberlineSecured = True
    End If
End Function

Public Function SelectedDBQ(c As Combobox) As String
    If c.ListIndex >= 0 Then
        SelectedDBQ = Parse(c.tag, c.ListIndex + 1, Chr(0))
    End If
End Function

Public Sub LoadDBQs(Optional list, Optional InitialDBQ = "")
    Dim dbqList As String
    Dim dbqItem As String
    Dim Name As String
    Dim Folder As String
    Dim folderlist As String
    Dim i As Long
    
    dbqList = GetDBQs
    If TypeName(list) = "VSFlexGrid" Then
        list.Rows = 1
        For i = 1 To Parse(dbqList, , Chr(1))
            dbqItem = Parse(dbqList, i, Chr(1))
            Name = Parse(dbqItem, 1, Chr(2))
            Folder = Parse(dbqItem, 2, Chr(2))
            list.AddItem Name & vbTab & Folder
        Next
        On Error Resume Next
        Call list.AutoSize(0, 1)
        list.Col = 0
        list.Sort = 7 'flexSortStringAscending
        list.Row = list.FindRow(InitialDBQ, , 1, False, True)
        Call list.ShowCell(list.Row, 0)
    Else
        list.Clear
        For i = 1 To Parse(dbqList, , Chr(1))
            dbqItem = Parse(dbqList, i, Chr(1))
            Name = Parse(dbqItem, 1, Chr(2))
            Folder = Parse(dbqItem, 2, Chr(2))
            list.AddItem Name
            list.ItemData(list.NewIndex) = i
            folderlist = folderlist & Chr(0) & Folder
        Next
        list.tag = Mid(folderlist, 2)
        
        On Error Resume Next
        For i = 0 To list.ListCount - 1
            If InitialDBQ = Parse(list.tag, list.ItemData(i), Chr(0)) Then
                list.ListIndex = i
            End If
        Next
        
        If list.Index < 0 And list.ListCount > 0 Then list.ListIndex = 0
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
    
    Dim list As String
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
            list = list & Chr(1) & sCompany & Chr(2) & sFolder
        End If
        Close iTsCtl
    
    Wend
    Close iTsDir
    list = Mid(list, 2)
    
    'remove any with no name
    tmp = ""
    For i = 1 To Parse(list, , Chr(1))
        If "" <> Parse(Parse(list, i, Chr(1)), 1, Chr(2)) Then
            tmp = tmp & Chr(1) & Parse(list, i, Chr(1))
        End If
    Next
    list = Mid(tmp, 2)
    
    GetDBQs = list

End Function

Public Sub TsObject(DataFolder As String, Filename As String, UserID As String, Pswd As String)
    Dim BatchFileName As String
    Dim BatchFile     As Integer

    BatchFileName = AppWorkingFolder & "\" & App.EXEName & ".bat"
    BatchFile = FreeFile
    
    Open BatchFileName For Output As #BatchFile
    Print #BatchFile, FileDrive(DataFolder)
    Print #BatchFile, "cd " & vbQuote & DataFolder & vbQuote
    Print #BatchFile, "tsobject.exe " & vbQuote & Filename & vbQuote & " " & UserID & " " & Pswd
    Print #BatchFile, "cd \"
    Print #BatchFile, "c:"
    Close BatchFile
    
    Call ShellAndLoop(BatchFileName)
On Error Resume Next
    Kill BatchFileName
    
End Sub

Public Function FormatTSField(Length As Long, Value As String, Optional RightAlignStrings As Boolean, Optional RemovePunctuation As Boolean) As String
    
    If RemovePunctuation Then
        Value = Replace(Value, ":", "")
        Value = Replace(Value, ";", "")
        Value = Replace(Value, "-", "")
        Value = Replace(Value, ".", "")
        Value = Replace(Value, "_", "")
    End If
    
    If TsIsNumeric(Value) Then
        FormatTSField = Format(Value, String(Length, "@"))
    Else
        If RightAlignStrings Then
            FormatTSField = Format(Value, String(Length, "@"))
        Else
            FormatTSField = Trim(Value)
        End If
    End If
End Function


Public Function TsIsNumeric(Value As String) As String
    Dim i As Long
    Dim b As Boolean
    b = True
    For i = 1 To Len(Value)
        b = b And IsNumeric(Mid(Value, i, 1))
    Next
    TsIsNumeric = b
End Function


Public Function SQLConnectionString(Server As String, Database As String, User As String, Pswd As String) As String
    Dim s As String
'    If User = "" Then
'        s = "Provider=MSOLEDBSQL;Server=" & Server & ";Database=" & Database & ";Trusted_Connection=yes;"
'    Else
'        s = "Provider=MSOLEDBSQL;Server=" & Server & ";Database=" & Database & ";UID=" & User & ";PWD=" & Pswd & ";"
'    End If
    
    If User = "" Then
        s = "Driver={SQL Server};Server=" & Server & ";Database=" & Database & ";Trusted_Connection=yes;"
    Else
        s = "Driver={SQL Server};Server=" & Server & ";Database=" & Database & ";UID=" & User & ";PWD=" & Pswd & ";"
    End If
    
    SQLConnectionString = s
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
                         "SilentLogin=1;" & _
                         "CODEPAGE=1252;" & _
                         "DatabaseType=" & DatabaseType & ";" & _
                         "ShortenNames=0;"
End If
End Function


Public Function Sage100ConnectionString() As String
    If HFApp.Options.ValueByName("Sage100APILevel") = "v20" Then
        Sage100ConnectionString = "Provider=sqloledb" & _
                                     ";Data Source=" & HFApp.Options.Value(MasterBuilderDataFolder) & _
                                     ";initial catalog=" & HFApp.Options.Value(MasterBuilderCompany) & _
                                     ";uid=" & HFApp.Options.Value(MasterBuilderUID) & _
                                     ";pwd=" & HFApp.Options.Value(MasterBuilderPWD) & _
                                     ";trusted_connection=" & IIf(HFApp.Options.Value(MasterBuilderUID) = "", "yes", "no")

    Else
        Sage100ConnectionString = "Provider=vfpoledb.1;Data Source=" & HFApp.Options.Value(MasterBuilderDataFolder) & ";Collating Sequence=general;"
    End If
End Function

