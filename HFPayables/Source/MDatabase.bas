Attribute VB_Name = "MDatabase"
Option Explicit

Public Const SQLDateFormat = "'YYYY-MM-DD'"
Public Const SQLTIMEFORMAT = "'HH:MM:SS'"
Public Const SQLDATETIMEFORMAT = "'YYYY-MM-DD HH:MM:SS'"

'errors
Public Enum SqlErrors
    NODSN = -2147467259    'Data source name not found and no default driver specified
    DUPKEY = -2147467259   'The record has a key field containing a duplicate Value(Btrieve Error 5)
    DUPKEY1 = -2147217900  '[SQL Server]Cannot insert duplicate key row in object 'Contacts' with unique index 'ContactsUK1'.
    MISCERR = -2147217900  'Pervasive seems to use this for virtually any exception
    NOTNULL = -2147217900  'Column <ColumnName> not nullable.
    TYPECONV = -2147217913 'Invalid date, time or timestamp value.
End Enum

Public Enum DbDataTypes
    Bit
    Date
    Time
    Str
    num
    DateTime
    Xml
End Enum

Public Sub DBPutFile(Connection As adodb.Connection, _
                     RowReference As String, _
                     fieldname As String, _
                     Filename As String)
'writes a file into a db field
'RowReference should include table name and where clause to limit resultset to a single row.
' ie: "Batches WHERE Batch=4"
    
'added-- mar 2016 MK
'odbc driver is limited to 400k blobs.
'create oledb connection instead.
    Dim s As String
    Dim svr As String
    Dim db As String
    Dim uid As String
    Dim pwd As String
    Dim trust As String
    Dim c As New Connection

    svr = Connection.Properties("server name")
    db = Connection.Properties("current catalog")
    uid = Parse(Parse(Connection.ConnectionString, 2, "UID="), 1, ";")
    pwd = Parse(Parse(Connection.ConnectionString, 2, "PWD="), 1, ";")
    trust = Parse(Parse(Connection.ConnectionString, 2, "trusted_Connection="), 1, ";")
    s = ""
    s = s & "provider=sqloledb;"
    s = s & "data source=" & svr & ";"
    s = s & "initial catalog=" & db & ";"
    If UCase(trust) = "YES" Then
        s = s & "trusted_connection=yes;"
    Else
        s = s & "UID=" & uid & ";"
        s = s & "pwd=" & pwd & ";"
    End If
    c.Open s
'end-- mar 2016 MK
    
    Dim rs As New adodb.Recordset
    Dim stream As New adodb.stream
    rs.CursorLocation = adUseClient
'    Call rs.Open("SELECT * FROM " & RowReference, Connection, adOpenKeyset, adLockOptimistic)
    Call rs.Open("SELECT * FROM " & RowReference, c, adOpenKeyset, adLockOptimistic)
    stream.Type = adTypeBinary
    stream.Open
    stream.LoadFromFile Filename
    rs(fieldname).Value = stream.Read
    stream.Close
    Set stream = Nothing
    rs.Update
    Set rs = Nothing
End Sub

Public Function DbQuote(DType As DbDataTypes, Value, Optional Nullable As Boolean, Optional TrimStr As Boolean, Optional Length As Long) As String
    Dim s As String
    Select Case DType
        Case Bit
            Select Case UCase(Value)
                Case "TRUE", "FALSE":          DbQuote = IIf(UCase("" & Value) = "TRUE", "1", "0")
                Case "YES", "NO":              DbQuote = IIf(UCase("" & Value) = "YES", "1", "0")
                Case "1", "0", "-1", 1, 0, -1: DbQuote = IIf("" & Value <> "0", "1", "0")
                Case "":                       DbQuote = "0"
                Case Else:                     DbQuote = IIf(CBool(Value), "1", "0")
                
            End Select
            
        Case Date
            If IsNull(Value) Or Value = "" Or Value = 0 Then
                DbQuote = "NULL"
            Else
                DbQuote = Format(Value, SQLDateFormat)
            End If

        Case DateTime
            If IsNull(Value) Or Value = "" Or Value = 0 Then
                DbQuote = "NULL"
            Else
                DbQuote = Format(Value, SQLDATETIMEFORMAT)
            End If

        Case Time
            If IsNull(Value) Or Value = "" Or Value = 0 Then
                DbQuote = "NULL"
            Else
                DbQuote = Format(Value, SQLTIMEFORMAT)
            End If


        Case Xml
            '* Replaces each quote with a pair of quotes
            '* Encloses entire string in quotes
            s = Value
            
            'strip the header out
            s = Replace(s, "<?xml version=""1.0"" encoding=""UTF-8""?>", "")
            
            If TrimStr Then s = Trim(s)
            If Length > 0 Then s = left(s, Length)
            s = "'" & Replace("" & s, "'", "''") & "'"
            If s = "''" And Nullable Then
                DbQuote = "NULL"
            Else
                DbQuote = s
            End If
            
        
        Case Str
            '* Replaces each quote with a pair of quotes
            '* Encloses entire string in quotes
            s = Value
            If TrimStr Then s = Trim(s)
            If Length > 0 Then s = left(s, Length)
            s = "'" & Replace("" & s, "'", "''") & "'"
            If s = "''" And Nullable Then
                DbQuote = "NULL"
            Else
                DbQuote = s
            End If

        Case num
            DbQuote = Val("" & Replace(Replace(Replace(Value, "%", ""), ",", ""), "$", ""))

    End Select

End Function

Public Function DbCompare(DType As DbDataTypes, Value, Optional NullStr As Boolean) As String
    Dim s As String
    
    Select Case DType
        Case Date
            If IsNull(Value) Or Value = "" Then
                DbCompare = " IS NULL"
            Else
                DbCompare = "=" & Format(Value, SQLDateFormat)
            End If
        
        Case Time
            If IsNull(Value) Or Value = "" Then
                DbCompare = " IS NULL"
            Else
                DbCompare = "=" & Format(Value, SQLTIMEFORMAT)
            End If
        
        Case Str
            '* Replaces each quote with a pair of quotes
            '* Encloses entire string in quotes
            s = "'" & Replace(Value, "'", "''") & "'"
            If s = "''" And Not NullStr Then
                DbCompare = " IS NULL"
            Else
                DbCompare = "=" & s
            End If
            
        Case num
            DbCompare = "=" & Val("" & Value)
            
              
    End Select

End Function

Public Function LoadDSNs(Optional ComboBox, Optional InitialDSN = "", Optional Driver As String) As String
    Dim i      As Integer
    Dim selected As Integer
    Dim s()    As String
    Dim list   As String
    
    Dim addit  As Boolean
    
    If RegEnumKeys(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", s) Then
        For i = 1 To UBound(s)
        If s(i) <> "" Then
            addit = True
            If Driver <> "" Then
                addit = UCase(RegGetKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", s(i))) Like "*" & UCase(Driver) & "*"
            End If
            
            If addit And s(i) <> "" Then
                list = list & "|" & s(i)
                If Not IsMissing(ComboBox) Then
                    ComboBox.AddItem s(i)
                End If
            End If
        End If
        Next
    End If
    LoadDSNs = list
    
On Error Resume Next
    For i = 0 To ComboBox.ListCount - 1
        If LCase(ComboBox.list(i)) = LCase(InitialDSN) Then
            ComboBox.ListIndex = i
            Exit Function
        End If
    Next
    ComboBox.ListIndex = 0
End Function


Public Sub ShowConnectionProperties(c, Title As String)
On Error Resume Next
    Dim s As String
    Dim a As String
    Dim i As Long
    
    
    If c.State = adStateOpen Then
        s = ""
        s = s & "Connection: " & Title & vbCrLf
        s = s & vbCrLf
        s = s & "DSN: " & c.Properties("Data Source Name").Value & IIf(c.Properties("Read-Only Data Source").Value, " (read only)", "") & "  (" & c.Properties("Current Catalog").Value & "@" & c.Properties("Server Name").Value & ")" & vbCrLf
        s = s & "DBMS: " & c.Properties("DBMS Name").Value & " version " & c.Properties("DBMS Version").Value & vbCrLf
        s = s & "Driver: " & c.Properties("Driver Name").Value & " version " & c.Properties("Driver Version").Value & vbCrLf
        s = s & "Provider: " & c.Properties("Provider Friendly Name").Value & vbCrLf
        s = s & "Version: " & c.Properties("Provider Name").Value & " version " & c.Properties("Provider Version").Value
        
        a = ""
        For i = 0 To c.Properties.Count - 1
            If c.Properties(i).Value <> "" And c.Properties(i).Value <> 0 Then
                a = a & c.Properties(i).Name & " = " & c.Properties(i).Value & vbCrLf
            End If
        Next
    Else
        s = ""
        s = s & "Connection: " & Title & vbCrLf
        s = s & vbCrLf
        s = s & "DSN: (no connection)" & vbCrLf
        s = s & "DBMS: " & vbCrLf
        s = s & "Driver: " & vbCrLf
        s = s & "Provider: " & vbCrLf
        s = s & "Version: "
        
        a = ""
        For i = 0 To c.Properties.Count - 1
            If c.Properties(i).Value <> "" And c.Properties(i).Value <> 0 Then
                a = a & c.Properties(i).Name & " = " & c.Properties(i).Value & vbCrLf
            End If
        Next
    
    End If
    MsgBox s, vbInformation, "Database Connection - " & Title ', , , , a
    
End Sub

Public Function GetDrivers() As String
    Dim i      As Integer
    Dim s()    As String
    Dim list   As String
    Dim d      As String
    
    If RegEnumKeys(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", s) Then
        For i = 1 To UBound(s)
            d = RegGetKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", s(i))
            
            If d = "Pervasive ODBC Engine Interface" Then d = "Pervasive Engine"
            If d = "Pervasive ODBC Client Interface" Then d = "Pervasive Client"
            
            If Not CBool(InStr(1, list, d)) Then
                list = list & Chr(1) & d
            End If
        Next
    End If
    
    GetDrivers = Mid(list, 2)
End Function


Public Function GetDSNs(Optional Driver As String) As String
    Dim i      As Integer
    Dim s()    As String
    Dim list   As String
    Dim d      As String
    
    
    If Driver = "Pervasive Engine" Then Driver = "Pervasive ODBC Engine Interface"
    If Driver = "Pervasive Client" Then Driver = "Pervasive ODBC Client Interface"
    
    If RegEnumKeys(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", s) Then
        For i = 1 To UBound(s)
            d = RegGetKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", s(i))
            If d = Driver Or Driver = "" Then
                list = list & Chr(1) & s(i) & Chr(2) & "dsn=" & s(i)
            End If
        Next
    End If
    
    GetDSNs = Mid(list, 2)

End Function

Public Function ConnectionString(Dsn As String, Optional User As String, Optional Pswd As String) As String
    ConnectionString = "provider=MSDASQL.1;dsn=" & Dsn & ";uid=" & User & ";pwd=" & Pswd & ";"
End Function
    

Public Function Round(Number As Double, DecimalPlaces As Integer) As Double
    Round = Format(Number, "0." & String(DecimalPlaces, "0"))
End Function
