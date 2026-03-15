Attribute VB_Name = "MDatabase"
Option Explicit

Public Const SQLDATEFORMAT = "'YYYY-MM-DD'"
Public Const SQLTIMEFORMAT = "'HH:MM:SS'"
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
End Enum

Public Enum RoundDirections
    rdNone
    rdupto
    rdDownTo
    rdClosest
End Enum

Public Property Let LastUpdate(item As String, RHS As Date)
On Error Resume Next
    Call HFApp.SqlExec("insert into appoptions(uid,optionvalue,optionname) values('',getdate()," & DbQuote(Str, "LastUpdate" & item) & ")")
    Call HFApp.SqlExec("UPDATE appoptions SET optionvalue=getdate() WHERE uid='' and optoinname=" & DbQuote(Str, "LastUpdate" & item))
End Property

Public Property Get LastUpdate(item As String) As Date
    Dim rs As adodb.Recordset
    Set rs = HFApp.SqlExec("select optionvalue from appoptions where uid='' and optionname=" & DbQuote(Str, "LastUpdate" & item))
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
                DbQuote = Format(Value, SQLDATEFORMAT)
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
            
        Case Num, NumInt
            DbQuote = Val("" & Replace(Replace(Replace(Value, "%", ""), ",", ""), "$", ""))
        
        Case Cur
            DbQuote = Round(Val("" & Replace(Replace(Replace(Value, "%", ""), ",", ""), "$", "")), 2)

    End Select

End Function

Public Function DbCompare(DType As DbDataTypes, Value, Optional NullStr As Boolean) As String
    Dim s As String
    
    Select Case DType
        Case Date
            If IsNull(Value) Or Value = "" Then
                DbCompare = " IS NULL"
            Else
                DbCompare = "=" & Format(Value, SQLDATEFORMAT)
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
            
        Case Num
            DbCompare = "=" & Val("" & Value)
            
              
    End Select

End Function

Public Function LoadDSNs(Optional Combobox, Optional InitialDSN = "", Optional Driver As String) As String
    Dim i      As Integer
    Dim Selected As Integer
    Dim s()    As String
    Dim buff   As String
    Dim list   As String
    
    Dim addit  As Boolean
    
    If RegEnumKeys(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", s) Then
        For i = 1 To UBound(s)
            If Trim(s(i)) <> "" Then
                buff = Trim(RegGetKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", s(i)))
                If UCase(buff) Like "*" & UCase(Driver) & "*" Then
                    list = list & "|" & s(i)
                    If Not IsMissing(Combobox) Then
                        Combobox.AddItem s(i) '& vbTab & buff
                    End If
                End If
            End If
        Next
    End If
    LoadDSNs = list
    
On Error Resume Next
    For i = 0 To Combobox.ListCount - 1
        If LCase(Combobox.list(i)) = LCase(InitialDSN) Then
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
    ConnectionString = "dsn=" & Dsn & ";uid=" & User & ";pwd=" & Pswd & ";"
End Function
    

Public Function Round(Number As Double, Optional DecimalPlaces As Integer) As Double
    Round = Format(Number, "0." & String(DecimalPlaces, "0"))
End Function

Public Function RoundTo(ByVal Number As Double, Significance As Double, Optional Direction As RoundDirections = rdupto) As Double

    

    Dim tmp As Integer

    Dim tmpVal As Double

    Dim TmpVal2 As Double

    

    If Significance = 0 Or Number = Significance Or Number / Significance = Int(Number / Significance) Then

        RoundTo = Number

        Exit Function

    End If

    

    Select Case Direction

        

        Case rdNone

            RoundTo = Number

            

        Case rdDownTo, rdupto

            'special case because i cant figure it out.

            'ie round 20 down to nearest 5.... otherwise would return 15

            If Direction = rdDownTo Then

                

                If Number > (Int(Number) + Significance) Then

                    

                    RoundTo = Int(Number) + Significance

                Else

                                'Round up to a whole integer -

                'Any decimal aluev will force a round to the next integer.

                'i.e. 0.01 = 1 or 0.8 = 1

                tmpVal = ((Number / Significance) + (-0.5 + IIf(Direction = rdupto, 1, 0)))

                TmpVal2 = tmpVal

                'Assigning fix(tmpVal) to tmp causes an overflow error for some reason.

                'if you call the function directly without assigning it to a variable it works fine.

                'tmp = Fix(TmpVal2)

                tmpVal = CInt((tmpVal - Fix(TmpVal2)) * 10 ^ 0)

                Number = Fix(TmpVal2) + tmpVal / 10 ^ 0

                'Multiply by ceiling value to set RoundtoValue

                RoundTo = Number * Significance

                End If

            Else

                'Round up to a whole integer -

                'Any decimal aluev will force a round to the next integer.

                'i.e. 0.01 = 1 or 0.8 = 1

                tmpVal = ((Number / Significance) + (-0.5 + IIf(Direction = rdupto, 1, 0)))

                TmpVal2 = tmpVal

                'Assigning fix(tmpVal) to tmp causes an overflow error for some reason.

                'if you call the function directly without assigning it to a variable it works fine.

                'tmp = Fix(TmpVal2)

                tmpVal = CInt((tmpVal - Fix(TmpVal2)) * 10 ^ 0)

                Number = Fix(TmpVal2) + tmpVal / 10 ^ 0

                'Multiply by ceiling value to set RoundtoValue

                RoundTo = Number * Significance

            End If

            

        Case rdClosest

            tmpVal = RoundTo(Number, Significance, rdDownTo)

            TmpVal2 = RoundTo(Number, Significance, rdupto)

            RoundTo = IIf(Abs(Number - tmpVal) < Abs(Number - TmpVal2), tmpVal, TmpVal2)

 

    End Select

    

End Function


Public Sub ConvertSQL(sql As String, c As adodb.Connection)
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
                s = Replace(s, "DOUBLE", "FLOAT", , , vbTextCompare)
                s = Replace(s, " IDENTITY", " INTEGER IDENTITY", , , vbTextCompare)
                s = Replace(s, "UCASE(", "UPPER(", , , vbTextCompare)
                s = Replace(s, "LCASE(", "LOWER(", , , vbTextCompare)
                s = Replace(s, "CURTIME(", "GETDATE(", , , vbTextCompare)
                s = Replace(s, "CONVERT(", "CAST(", , , vbTextCompare)
                s = Replace(s, ", SQL_VARCHAR)", " AS VARCHAR)", , , vbTextCompare)
                s = Replace(s, ",SQL_VARCHAR)", " AS VARCHAR)", , , vbTextCompare)
                s = Replace(s, ", SQL_INTEGER)", " AS INTEGER)", , , vbTextCompare)
                s = Replace(s, ",SQL_INTEGER)", " AS INTEGER)", , , vbTextCompare)
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






Public Sub DBPutFile(Connection As adodb.Connection, _
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
    stream.LoadFromFile FileName
    rs(fieldname).Value = stream.Read
    stream.Close
    Set stream = Nothing
    rs.Update
    Set rs = Nothing
End Sub

Public Sub DBGetFile(FileName As String, _
                     Optional ADOField As adodb.field, _
                     Optional Connection As adodb.Connection, _
                     Optional RowReference As String, _
                     Optional fieldname As String)
'extracts a file from a db field
'pass either an ADOField or Connection,RowReference,Fieldname

'RowReference should include table name and where clause to limit resultset to a single row.
' ie: "Batches WHERE Batch=4"
                     
    Dim rs As New adodb.Recordset
    Dim i As Long
    Dim Data() As Byte
    Dim Temp As Variant
    
    If RowReference <> "" Then
        Call rs.Open("SELECT * FROM " & RowReference, Connection)
        Set ADOField = rs.fields.item(fieldname)
    End If
    
    i = FreeFile
    Open FileName For Binary As #i
    Do
        Temp = ADOField.GetChunk(16384)
        If IsNull(Temp) Then Exit Do
        Data = Temp
        Put #i, , Data
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




