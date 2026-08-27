Attribute VB_Name = "MDatabase"
Option Explicit

Public Const SQLTIMEFORMAT = "'HH:MM:SS'"
Public Const SQLDATEFORMAT = "'YYYY-MM-DD'"
Public Const SQLDATETIMEFORMAT = "'YYYY-MM-DD HH:MM:SS'"

'errors
Public Enum SqlErrors
    NODSN = -2147467259    'Data source name not found and no default driver specified
    DUPKEY = -2147467259   'The record has a key field containing a duplicate Value(Btrieve Error 5)
    DUPKEY1 = -2147217900  '[SQL Server]Cannot insert duplicate key row in object 'tablename' with unique index 'indexname'.
    MISCERR = -2147217900  'Pervasive seems to use this for virtually any exception
    NOTNULL = -2147217900  'Column <ColumnName> not nullable.
    TYPECONV = -2147217913 'Invalid date, time or timestamp value.
End Enum

Public Enum DbDataTypes
    Date = 1
    Time
    Str
    Num
    NumInt
    DateTime
    Cur
    Bit
    xml
End Enum
Public Enum RoundDirections
    rdNone
    rdUpTo
    rdDownTo
    rdClosest
End Enum

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

Public Property Let LastUpdate(Item As String, RHS As Date)
On Error Resume Next
    Call HFApp.SqlExec("insert into appoptions(DivisionID,uid,optionvalue,optionname) values(" & HFApp.DivisionID & "'',getdate()," & DbQuote(Str, "LastUpdate" & Item) & ")")
    Call HFApp.SqlExec("UPDATE appoptions SET optionvalue=getdate() WHERE DivisionID = " & HFApp.DivisionID & " and uid='' and optionname=" & DbQuote(Str, "LastUpdate" & Item))
End Property

Public Property Get LastUpdate(Item As String) As Date
    Dim rs As ADODB.Recordset
    Set rs = HFApp.SqlExec("select optionvalue from appoptions where DivisionID = " & HFApp.DivisionID & " and uid='' and optionname=" & DbQuote(Str, "LastUpdate" & Item))
    If rs.EOF Then
        LastUpdate = DateValue("1950-01-01")
    Else
        If Not IsDate("" & rs(0)) Then
            LastUpdate = DateValue("1950-01-01")
        Else
            LastUpdate = rs(0)
        End If
    End If
    Set rs = Nothing
End Property

Public Function DbQuote(DType As DbDataTypes, value, Optional Nullable As Boolean, Optional TrimStr As Boolean, Optional Length As Long) As String
    Dim s As String
    
    Select Case DType
        Case Bit
            Select Case UCase(value)
                Case "TRUE", "FALSE":          DbQuote = IIf(UCase("" & value) = "TRUE", "1", "0")
                Case "YES", "NO":              DbQuote = IIf(UCase("" & value) = "YES", "1", "0")
                Case "1", "0", "-1", 1, 0, -1: DbQuote = IIf("" & value <> "0", "1", "0")
                Case "":                       DbQuote = "0"
                Case Else:                     DbQuote = IIf(CBool(value), "1", "0")
                
            End Select
            
            
        Case Date
            If IsNull(value) Or value = "" Or value = 0 Then
                DbQuote = "NULL"
            Else
                DbQuote = format(value, SQLDATEFORMAT)
            End If

        Case DateTime
            If IsNull(value) Or value = "" Or value = 0 Then
                DbQuote = "NULL"
            Else
                DbQuote = format(value, SQLDATETIMEFORMAT)
            End If

        Case Time
            If IsNull(value) Or value = "" Or value = 0 Then
                DbQuote = "NULL"
            Else
                DbQuote = format(value, SQLTIMEFORMAT)
            End If


        Case Str
            '* Replaces each quote with a pair of quotes
            '* Encloses entire string in quotes
            s = value
            If TrimStr Then s = Trim(s)
            If Length > 0 Then s = Left(s, Length)
            s = "'" & Replace("" & s, "'", "''") & "'"
            If s = "''" And Nullable Then
                DbQuote = "NULL"
            Else
                DbQuote = s
            End If
            
        Case xml
            '* Replaces each quote with a pair of quotes
            '* Encloses entire string in quotes
            s = value
            
            'strip the header out
            s = Replace(s, "<?xml version=""1.0"" encoding=""UTF-8""?>", "")
            
            If TrimStr Then s = Trim(s)
            If Length > 0 Then s = Left(s, Length)
            s = "'" & Replace("" & s, "'", "''") & "'"
            If s = "''" And Nullable Then
                DbQuote = "NULL"
            Else
                DbQuote = s
            End If
            
        
        Case Num, NumInt
            If Nullable And value = "" Then
                DbQuote = "NULL"
            Else
                DbQuote = Val("" & Replace(Replace(Replace(value, "%", ""), ",", ""), "$", ""))
            End If
            
        Case Cur
            If Nullable And value = "" Then
                DbQuote = "NULL"
            Else
                DbQuote = Round(Val("" & Replace(Replace(Replace(value, "%", ""), ",", ""), "$", "")), 2)
            End If
            
    End Select

End Function

Public Function DbCompare(DType As DbDataTypes, value, Optional NullStr As Boolean) As String
    Dim s As String
    
    Select Case DType
        Case Date
            If IsNull(value) Or value = "" Then
                DbCompare = " IS NULL"
            Else
                DbCompare = "=" & format(value, SQLDATEFORMAT)
            End If
        
        Case Time
            If IsNull(value) Or value = "" Then
                DbCompare = " IS NULL"
            Else
                DbCompare = "=" & format(value, SQLTIMEFORMAT)
            End If
        
        Case Str
            '* Replaces each quote with a pair of quotes
            '* Encloses entire string in quotes
            s = "'" & Replace(value, "'", "''") & "'"
            If s = "''" And Not NullStr Then
                DbCompare = " IS NULL"
            Else
                DbCompare = "=" & s
            End If
            
        Case Num
            DbCompare = "=" & Val("" & value)
            
              
    End Select

End Function

Public Function LoadDSNs(Optional Combobox, Optional InitialDSN = "", Optional Driver As String) As String
    Dim i      As Integer
    Dim Selected As Integer
    Dim s()    As String
    Dim buff   As String
    Dim List   As String
    
    Dim addit  As Boolean
    
    If RegEnumKeys(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", s) Then
        For i = 1 To UBound(s)
            If Trim(s(i)) <> "" Then
                buff = Trim(RegGetKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", s(i)))
                If UCase(buff) Like "*" & UCase(Driver) & "*" Then
                    List = List & "|" & s(i)
                    If Not IsMissing(Combobox) Then
                        Combobox.AddItem s(i) '& vbTab & buff
                    End If
                End If
            End If
        Next
    End If
    LoadDSNs = List
    
On Error Resume Next
    For i = 0 To Combobox.ListCount - 1
        If LCase(Combobox.List(i)) = LCase(InitialDSN) Then
            Combobox.ListIndex = i
            Exit Function
        End If
    Next
    Combobox.ListIndex = 0
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
        s = s & "DSN: " & c.Properties("Data Source Name").value & IIf(c.Properties("Read-Only Data Source").value, " (read only)", "") & "  (" & c.Properties("Current Catalog").value & "@" & c.Properties("Server Name").value & ")" & vbCrLf
        s = s & "DBMS: " & c.Properties("DBMS Name").value & " version " & c.Properties("DBMS Version").value & vbCrLf
        s = s & "Driver: " & c.Properties("Driver Name").value & " version " & c.Properties("Driver Version").value & vbCrLf
        s = s & "Provider: " & c.Properties("Provider Friendly Name").value & vbCrLf
        s = s & "Version: " & c.Properties("Provider Name").value & " version " & c.Properties("Provider Version").value
        
        a = ""
        For i = 0 To c.Properties.Count - 1
            If c.Properties(i).value <> "" And c.Properties(i).value <> 0 Then
                a = a & c.Properties(i).Name & " = " & c.Properties(i).value & vbCrLf
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
            If c.Properties(i).value <> "" And c.Properties(i).value <> 0 Then
                a = a & c.Properties(i).Name & " = " & c.Properties(i).value & vbCrLf
            End If
        Next
    
    End If
    MsgBox s, vbInformation, "Database Connection - " & Title ', , , , a
    
End Sub

Public Function GetDrivers() As String
    Dim i      As Integer
    Dim s()    As String
    Dim List   As String
    Dim d      As String
    
    If RegEnumKeys(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", s) Then
        For i = 1 To UBound(s)
            d = RegGetKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", s(i))
            
            If d = "Pervasive ODBC Engine Interface" Then d = "Pervasive Engine"
            If d = "Pervasive ODBC Client Interface" Then d = "Pervasive Client"
            
            If Not CBool(InStr(1, List, d)) Then
                List = List & Chr(1) & d
            End If
        Next
    End If
    
    GetDrivers = Mid(List, 2)
End Function


Public Function GetDSNs(Optional Driver As String) As String
    Dim i      As Integer
    Dim s()    As String
    Dim List   As String
    Dim d      As String
    
    
    If Driver = "Pervasive Engine" Then Driver = "Pervasive ODBC Engine Interface"
    If Driver = "Pervasive Client" Then Driver = "Pervasive ODBC Client Interface"
    
    If RegEnumKeys(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", s) Then
        For i = 1 To UBound(s)
            d = RegGetKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", s(i))
            If d = Driver Or Driver = "" Then
                List = List & Chr(1) & s(i) & Chr(2) & "dsn=" & s(i)
            End If
        Next
    End If
    
    GetDSNs = Mid(List, 2)

End Function

Public Function ConnectionString(Dsn As String, Optional User As String, Optional Pswd As String) As String
    ConnectionString = "dsn=" & Dsn & ";uid=" & User & ";pwd=" & Pswd & ";"
End Function
    

Public Function Round(Number As Double, Optional DecimalPlaces As Integer) As Double
    Round = format(Number, "0." & String(DecimalPlaces, "0"))
End Function

Public Function DecimalPlaces(ByVal Number As Currency) As Long
    Number = Number - Int(Number)
    If Number = 0 Then
        DecimalPlaces = 0
    Else
        DecimalPlaces = Len("" & Number) - 2
    End If
End Function

Public Function RoundToPrice(ByVal Number As Double, Significance As Double) As Currency
'Round to a customer friendly sales price. Significance can be a multiple like 50 or 1, or an ending suffix like .95 or 990
'
'SUFFIXES
'roundtoprice(12365.3456, 990) --> 12990.00
'roundtoprice(12365.3456, .95) --> 12365.95
'
'MULTIPLES
'roundtoprice(12365.3456, 100) --> 13400.00
'roundtoprice(12365.3456, 50)  --> 12400.00
'roundtoprice(12365.3456, 10)  --> 12470.00
'roundtoprice(12365.3456, 5)   --> 12370.00
'roundtoprice(12365.3456, 1)   --> 12366.00
'roundtoprice(12365.3456, .5)  --> 12365.50
'roundtoprice(12365.3456, .1)  --> 12365.40
'roundtoprice(12365.3456, .05) --> 12365.35
'roundtoprice(12365.3456, .01) --> 12365.35

    Dim d As Long
    Dim s1 As String
    Dim s2 As String
    Dim value As Double
    
    
    
    If Number = 0 Or Significance = 0 Then
        RoundToPrice = Number
        Exit Function
    End If
    If IsIn(Replace(Replace(Significance, "0", ""), ".", ""), 1, 5, 25) Then
        'significance is an even magnitude of 1 or 5 or 25, round to multiple of significance
        RoundToPrice = RoundTo(Number, Significance, rdUpTo)
    Else
        'is a suffix type round
        d = Max(DecimalPlaces(CCur(Number)), DecimalPlaces(CCur(Significance)))
        s1 = format(Number, "." & String(d, "0"))
        s2 = format(Significance, "." & String(d, "0"))
        value = Left(s1, Len("" & s1) - Len("" & s2)) & s2
        If value < Number Then
            value = Val(Left(s1, Len("" & s1) - Len("" & s2))) + 1 & s2
        End If
        RoundToPrice = value
    End If
    
End Function

Public Function RoundTo(ByVal Number As Double, Significance As Double, Optional Direction As RoundDirections = rdUpTo) As Double
    Dim tmp As Integer
    Dim tmpVal As Double
    Dim TmpVal2 As Double
    
    If Significance = 0 Or Number = Significance Then
        RoundTo = Number
        Exit Function
    End If
    If Number / Significance = Int(Number / Significance) Then
        RoundTo = Number
        Exit Function
    End If
    Select Case Direction
        Case rdNone
            RoundTo = Number

        Case rdDownTo
            If Number > (Int(Number) + Significance) Then
                RoundTo = Int(Number) + Significance
            Else
                tmpVal = ((Number / Significance) + (-0.5 + IIf(Direction = rdUpTo, 1, 0)))
                TmpVal2 = tmpVal
                tmpVal = CInt((tmpVal - Fix(TmpVal2)) * 10 ^ 0)
                Number = Fix(TmpVal2) + tmpVal / 10 ^ 0
                RoundTo = Number * Significance
            End If
            
        Case rdUpTo
            tmpVal = ((Number / Significance) + (-0.5 + IIf(Direction = rdUpTo, 1, 0)))
            TmpVal2 = tmpVal
            tmpVal = CInt((tmpVal - Fix(TmpVal2)) * 10 ^ 0)
            Number = Fix(TmpVal2) + tmpVal / 10 ^ 0
            RoundTo = Number * Significance

        Case rdClosest
            tmpVal = RoundTo(Number, Significance, rdDownTo)
            TmpVal2 = RoundTo(Number, Significance, rdUpTo)
            RoundTo = IIf(Abs(Number - tmpVal) < Abs(Number - TmpVal2), tmpVal, TmpVal2)

    End Select

    

End Function

Public Sub ConvertSQL(sql As String, c As ADODB.Connection)
On Error GoTo eh

    Dim i As Long
    Dim pieces() As String
    Dim s As String
    
    Select Case c.Properties("DBMS Name")
        
        Case "Pervasive.SQL"
          'do nothing
            
        Case "Microsoft SQL Server"
            pieces = split(sql, "'")
            For i = LBound(pieces) To UBound(pieces) Step 2
                s = pieces(i)
                s = Replace(s, "IFNULL", "ISNULL", , , vbTextCompare)
                s = Replace(s, "CURDATE", "GETDATE", , , vbTextCompare)
                s = Replace(s, "NOW()", "GETDATE()", , , vbTextCompare)
                s = Replace(s, "LONGVARCHAR", "VARCHAR(8000)", , , vbTextCompare)
                s = Replace(s, "DOUBLE", "DECIMAL", , , vbTextCompare)
                s = Replace(s, " IDENTITY", " INTEGER IDENTITY", , , vbTextCompare)
' why is this here                s = Replace(s, " CASE", "", , , vbTextCompare)
                s = Replace(s, " DATE", " DATETIME", , , vbTextCompare)
                s = Replace(s, " DATETIMETIME", " DATETIME", , , vbTextCompare)
                s = Replace(s, " TIMESTAMP", " DATETIME", , , vbTextCompare)
                s = Replace(s, " TIME", " DATETIME", , , vbTextCompare)
                s = Replace(s, "UCASE(", "UPPER(", , , vbTextCompare)
                s = Replace(s, "LCASE(", "LOWER(", , , vbTextCompare)
                s = Replace(s, "CURTIME(", "GETDATE(", , , vbTextCompare)
                s = Replace(s, "CONVERT(", "CAST(", , , vbTextCompare)
                s = Replace(s, ", SQL_VARCHAR)", " AS VARCHAR)", , , vbTextCompare)
                s = Replace(s, ",SQL_VARCHAR)", " AS VARCHAR)", , , vbTextCompare)
                s = Replace(s, ", SQL_INTEGER)", " AS INTEGER)", , , vbTextCompare)
                s = Replace(s, ",SQL_INTEGER)", " AS INTEGER)", , , vbTextCompare)
                s = Replace(s, " MODIFY ", " ALTER COLUMN ", , , vbTextCompare)
                s = Replace(s, "CAST(VARCHAR", "CONVERT(VARCHAR", , , vbTextCompare)
                
                
                pieces(i) = s
            Next
            s = ""
            For i = LBound(pieces) To UBound(pieces)
                s = s & "'" & pieces(i)
            Next
            sql = Mid(s, 2)
            
    End Select
    Exit Sub
eh:
    Select Case Err.Number
    
        Case 91   'Object variable or With block variable not set
            Err.Raise vbObjectError, "App::ConvertSQL()", "Database connection is nothing"
            
        Case 3265 'Item cannot be found in the collection corresponding to the requested name or ordinal.
            Exit Sub
            
        Case Else
            Err.Raise Err.Number
            
    End Select
End Sub






Public Sub DBPutFile(Connection As ADODB.Connection, _
                     RowReference As String, _
                     fieldname As String, _
                     FileName As String)
'writes a file into a db field
'RowReference should include table name and where clause to limit resultset to a single row.
' ie: "Batches WHERE Batch=4"
    
'added-- mar 2016 MK
'odbc driver is limited to 400k blobs.
'create oledb connection instead.
    Dim s As String
    Dim svr As String
    Dim db As String
    Dim uId As String
    Dim pwd As String
    Dim trust As String
    Dim c As New Connection

    svr = Connection.Properties("server name")
    db = Connection.Properties("current catalog")
    uId = Parse(Parse(Connection.ConnectionString, 2, "UID="), 1, ";")
    pwd = Parse(Parse(Connection.ConnectionString, 2, "PWD="), 1, ";")
    trust = Parse(Parse(Connection.ConnectionString, 2, "trusted_Connection="), 1, ";")
    s = ""
    s = s & "provider=sqloledb;"
    s = s & "data source=" & svr & ";"
    s = s & "initial catalog=" & db & ";"
    If UCase(trust) = "YES" Then
        s = s & "trusted_connection=yes;"
    Else
        s = s & "UID=" & uId & ";"
        s = s & "pwd=" & pwd & ";"
    End If
    c.Open s
'end-- mar 2016 MK
    
    Dim rs As New ADODB.Recordset
    Dim stream As New ADODB.stream
    rs.CursorLocation = adUseClient
'    Call rs.Open("SELECT * FROM " & RowReference, Connection, adOpenKeyset, adLockOptimistic)
    Call rs.Open("SELECT * FROM " & RowReference, c, adOpenKeyset, adLockOptimistic)
    stream.Type = adTypeBinary
    stream.Open
    stream.LoadFromFile FileName
    rs(fieldname).value = stream.Read
    stream.Close
    Set stream = Nothing
    rs.Update
    Set rs = Nothing
End Sub


Public Sub DBGetFile(FileName As String, _
                     Optional ADOField As ADODB.Field, _
                     Optional Connection As ADODB.Connection, _
                     Optional RowReference As String, _
                     Optional fieldname As String)
'extracts a file from a db field
'pass either an ADOField or Connection,RowReference,Fieldname

'RowReference should include table name and where clause to limit resultset to a single row.
' ie: "Batches WHERE Batch=4"
                     
    Dim rs As New ADODB.Recordset
    Dim i As Long
    Dim data() As Byte
    Dim Temp As Variant
    
    If RowReference <> "" Then
        Call rs.Open("SELECT * FROM " & RowReference, Connection)
        Set ADOField = rs.fields.Item(fieldname)
    End If
    
    i = FreeFile
    Open FileName For Binary As #i
    Do
        Temp = ADOField.GetChunk(16384)
        If IsNull(Temp) Then Exit Do
        data = Temp
        Put #i, , data
    Loop While LenB(Temp) = 16384
    Close #i
    
End Sub


Public Function DbQuoteType(DataType) As Long
On Error Resume Next
    Dim i As DbDataTypes
    Select Case DataType
    
        Case adTinyInt, adSmallInt, adInteger, adBigInt, adUnsignedTinyInt, adUnsignedSmallInt, adUnsignedInt, adUnsignedBigInt
            i = NumInt
        
        Case adSingle, adDouble, adDecimal, adNumeric, adCurrency
            i = Num
        
        Case adDate, adDBDate
            i = Date
        
        Case adDBTime, adDBTimeStamp
            i = DateTime

        Case adVarChar
            i = Str
        
        Case adBoolean
            i = Bit
            
    End Select
    DbQuoteType = i
End Function


Public Function ColumnDefinitionPretty(c) As String
On Error Resume Next
    Dim s As String
    Select Case c.Type
    
        Case adTinyInt, adSmallInt, adInteger, adBigInt, adUnsignedTinyInt, adUnsignedSmallInt, adUnsignedInt, adUnsignedBigInt
            s = "Numeric (integer)"
        
        Case adSingle, adDouble, adDecimal, adNumeric
            s = "Numeric (decimal)"
            
        Case adCurrency
            s = "Currency"
        
        Case adDate, adDBDate
            s = "Date"
        
        Case adDBTime, adDBTimeStamp
            s = "Date & Time"

        Case adVarChar
            s = "Text (" & c.DefinedSize & " characters)"
        
        Case adBoolean
            s = "Yes/No"
            
        Case adUserDefined:      s = "UserDefined"
        Case adVariant:          s = "Variant"
        Case adGUID:             s = "Guid"
        Case adBSTR:             s = "BSTR"
        Case adChar:             s = "Char"
        Case adLongVarChar:      s = "LongVarChar"
        Case adWChar:            s = "WChar"
        Case adVarWChar:         s = "VarWChar"
        Case adLongVarWChar:     s = "LongVarWChar"
        Case adBinary:           s = "Binary"
        Case adVarBinary:        s = "VarBinary"
        Case adLongVarBinary:    s = "LongVarBinary"
    End Select
    ColumnDefinitionPretty = s
End Function




