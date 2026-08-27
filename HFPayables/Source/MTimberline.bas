Attribute VB_Name = "MTimberline"
Option Explicit

Public Enum NameModes
    CustomNames
    StandardNames
    DictionaryNames
End Enum

Public Function TimberlineInstalled() As Boolean
    TimberlineInstalled = "" <> RegGetKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Timberline\General", "System Directory")
End Function

Public Function SelectedDBQ(c As ComboBox) As String
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

Public Function TimberlineSystemPath() As String
    TimberlineSystemPath = RegGetKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Timberline\General", "System Directory")
End Function
Public Function TimberlineInstallPath() As String
    TimberlineInstallPath = RegGetKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Timberline\OfficeInstalls\Accounting", "InstallPath")
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
    
    'open ts.dir and read data folder paths
    sTsDir = TimberlineSystemPath
    If sTsDir = "" Then Exit Function
    sTsDir = sTsDir & "\misc\ts.dir"
    CurDrive = FileDrive(sTsDir)
    
    iTsDir = FreeFile
    Open sTsDir For Input Access Read As iTsDir
    While Not EOF(iTsDir)
        Line Input #iTsDir, sFolder
        sFolder = CurDrive & sFolder
        
        'open ts.ctl read company name
        sTsCtl = sFolder & "\ts.ctl"
        iTsCtl = FreeFile
        Open sTsCtl For Binary Access Read As iTsCtl
        If err.Number = 0 Then
            sCompany = DNull(Trim(Mid(Input(50, iTsCtl), 17)))
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

Public Sub DONTUSETHIS_TsObject(DataFolder As String, Filename As String, UserID As String, Pswd As String)

    Dim TSobjPath     As String
    Dim BatchFileName As String
    Dim BatchFile     As Integer
    
    TSobjPath = RegGetKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Timberline\General", "Workstation System Directory", "C:\Program Files\Timberline Office\Accounting\")
    TSobjPath = PathAppend(TSobjPath, "tsobject.exe")
    BatchFileName = AppWorkingFolder & "\" & App.ExeName & ".bat"
    BatchFile = FreeFile
    
    Open BatchFileName For Output As #BatchFile
    Print #BatchFile, FileDrive(DataFolder)
    Print #BatchFile, "cd " & vbQuote & DataFolder & vbQuote
    Print #BatchFile, vbQuote & TSobjPath & vbQuote & " " & vbQuote & Filename & vbQuote & " " & UserID & " " & Pswd
    Print #BatchFile, "cd \"
    Print #BatchFile, "c:"
    Close BatchFile
    
    Call Shell(BatchFileName, vbHide)
    
End Sub

Public Function FormatTSField(Length As Long, Value As String, Optional RightAlignStrings As Boolean, Optional dontRemovePunctuation As Boolean) As String
On Error GoTo err
    If HFApp.ConnectionString(dbTimberlinePVdata) <> "" And HFApp.Options(AccountingSystem) = AccountingSystems.asTimberline Then
        If Not dontRemovePunctuation Then
            Value = Replace(Value, ":", "")
            Value = Replace(Value, ";", "")
            Value = Replace(Value, "-", "")
            Value = Replace(Value, ".", "")
            Value = Replace(Value, "_", "")
        End If
        If TsIsNumeric(Value) = "True" Or RightAlignStrings = True Then
            FormatTSField = Format(Value, String(Length, "@"))
        Else
            FormatTSField = Trim(Value)
        End If
    Else
        
        FormatTSField = Trim(Value)

    End If
Exit Function
err:
If err.Number <> 0 Then
    MsgBox err.Description & " Format TSField"
End If

End Function


Private Function TsIsNumeric(Value As String) As String
On Error Resume Next
    Dim i As Long
    Dim b As Boolean
    b = True
    For i = 1 To Len(Value)
        b = b And IsNumeric(Mid(Value, i, 1))
    Next
    TsIsNumeric = b
End Function

Public Function TsConnectionString(DataFolder As String, Optional User As String, Optional Pswd As String, Optional NamingMode As NameModes = StandardNames) As String
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
                         "DatabaseType=1;" & _
                         "ShortenNames=0;"
End Function
