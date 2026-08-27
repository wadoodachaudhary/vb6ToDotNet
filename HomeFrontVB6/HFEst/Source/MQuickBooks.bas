Attribute VB_Name = "MQuickBooks"
'======================================================================
' This file was stolen from CommandCenter. Not all procedures have been
' updated for use in HomeFront. If something doesnt make sense its
' probably because it hasnt been updated...
'======================================================================


Option Explicit
Option Compare Text
Private Const SRCFILE = "MQuickBooks::"

Public Enum XMLFieldTypes
    a 'BankAcct
    b 'RoutingNum
    c 'Char
    d 'Date
    F 'File
    g 'SocialSecurityNumber
    m 'Memo
    n 'Numeric
    o 'AssemblyPartInclusionFlag
    p 'PhoneNumber
    r 'Terms
    st 'State
    t 'Time
    Y 'YesNo
    z 'zipcode
End Enum

Public Sub QB_ImportCategories()
On Error GoTo eh
    Dim s As String
    Dim rs As Recordset
    Dim costtype As Integer
    
    s = ""
    s = s & StartQBXml()
    s = s & "<ClassQueryRq/>" & vbCrLf
    s = s & EndQBXml()
    s = SubmitQBXml(s)
    If s <> "" Then s = "exec QB_Categories " & DbQuote(Str, s)
    Set rs = HFApp.SqlExec(s, dbHomefront)
    
    If Not rs.EOF Then Call HFApp.SqlExec("delete from stdcategories", dbHomefront)
    
    While Not rs.EOF
    
        Select Case UCase(Left("" & rs("Accumulate_As"), 3))
            Case "LAB": costtype = 1
            Case "MAT": costtype = 2
            Case "SUB": costtype = 3
            Case "OVE": costtype = 5
            Case "OTH": costtype = 6
            Case Else:  costtype = 4
        End Select
        
        s = ""
        s = s & "INSERT INTO StdCategories(Category,Description,CostType)" & vbCrLf
        s = s & "VALUES(" & DbQuote(Str, "" & rs("Category")) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("Description")) & vbCrLf
        s = s & "      ," & DbQuote(Num, costtype) & ")"
        Call HFApp.SqlExec(s, dbHomefront)
        
        s = ""
        s = s & "UPDATE StdCategories" & vbCrLf
        s = s & "SET Description=" & DbQuote(Str, "" & rs("Description")) & vbCrLf
        s = s & "   ,CostType=" & DbQuote(Num, costtype) & vbCrLf
        s = s & "WHERE Category=" & DbQuote(Str, "" & rs("Category")) & vbCrLf
        Call HFApp.SqlExec(s, dbHomefront)

        
        rs.MoveNext
    Wend


Exit Sub
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Call errHandler(SRCFILE & "QB_ImportCategories")
    End If
End Sub

Public Sub QB_ImportCostCodes()
On Error GoTo eh
    Dim i As Long
    Dim c As Long
    Dim s As String
    Dim rs As Recordset
            
    s = ""
    s = s & StartQBXml()
    s = s & "<ItemQueryRq metaData=""MetaDataAndResponseData"">" & vbCrLf
    s = s & "</ItemQueryRq>" & vbCrLf
    s = s & EndQBXml()
    s = SubmitQBXml(s)
    If s <> "" Then
        On Error Resume Next
        c = Val(Parse(Parse(s, 2, "retCount="""), 1, "="))
        i = 0
        On Error GoTo eh
        s = "exec QB_CostCodes " & DbQuote(Str, s)
    End If
    
    Set rs = HFApp.SqlExec(s, dbHomefront)
    
    If Not rs.EOF Then Call HFApp.SqlExec("delete from stdcostcodes", dbHomefront)
        
    While Not rs.EOF
        
        i = i + 1
        'FMain.lblStatus.Caption = "reading cost code " & i & " of " & c
        'FMain.lblStatus.Refresh
    
        s = ""
        s = s & "INSERT INTO StdCostCodes(CostCode,Description,GroupPhase)" & vbCrLf
        s = s & "VALUES(" & DbQuote(Str, "" & rs(0)) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs(1), , True) & ",0)"
        Call HFApp.SqlExec(s, dbHomefront)
        
        s = ""
        s = s & "UPDATE StdCostCodes" & vbCrLf
        s = s & "SET Description=" & DbQuote(Str, "" & rs(1), , True) & vbCrLf
        s = s & "WHERE CostCode=" & DbQuote(Str, "" & rs(0))
        Call HFApp.SqlExec(s, dbHomefront)
        
        rs.MoveNext
    Wend

    'FMain.lblStatus.Caption = ""
    'FMain.lblStatus.Refresh

Exit Sub
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Call errHandler(SRCFILE & "QB_ImportCostCodes")
    End If
End Sub

Public Sub QB_ImportTaxGroups()
On Error GoTo eh
    Dim s As String
    Dim rs As Recordset
    
    
    s = SubmitQBXml(StartQBXml() & "<ItemQueryRq/><SalesTaxCodeQueryRq/>" & EndQBXml())
    If s <> "" Then s = "exec QB_TaxGroups " & DbQuote(Str, s)
    Set rs = HFApp.SqlExec(s, dbHomefront)

    If Not rs.EOF Then Call HFApp.SqlExec("delete from taxgroups", dbHomefront)
    
    While Not rs.EOF
        s = ""
        s = s & "INSERT INTO TaxGroups([Group],grate,GDesc,ExternalID)" & vbCrLf
        s = s & "VALUES(" & DbQuote(Str, "" & rs("TaxGroup")) & vbCrLf
        s = s & "      ," & DbQuote(Num, "" & rs("GroupRate")) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("Description"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("ExternalID")) & ")"
        Call HFApp.SqlExec(s, dbHomefront)
        
        s = ""
        s = s & "UPDATE TaxGroups" & vbCrLf
        s = s & "SET gdesc=" & DbQuote(Str, "" & rs("Description"), , True) & vbCrLf
        s = s & "   ,grate=" & DbQuote(Num, "" & rs("GroupRate")) & vbCrLf
        s = s & "   ,ExternalID=" & DbQuote(Str, "" & rs("ExternalID")) & vbCrLf
        s = s & "WHERE [group]=" & DbQuote(Str, "" & rs("TaxGroup")) & vbCrLf
        Call HFApp.SqlExec(s, dbHomefront)
        
        rs.MoveNext
    Wend
Exit Sub
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Call errHandler(SRCFILE & "QB_ImportTaxGroups")
    End If
End Sub




Private Function SubmitQBXml(XmlMessage As String) As String
'On Error Resume Next
Err.Clear
    Dim ticket  As String
    Dim Response As String
    
    Dim Status   As String
    Dim reason   As String
    Dim i As Integer
    Dim RP As New QBXMLRP2Lib.RequestProcessor2
    
'log request
i = FreeFile()
Open PathAppend(GetFolderPath(CSIDL_LOCAL_APPDATA), "Zybertech\" & VB.App.EXEName & "\request.xml") For Output As #i
Print #i, XmlMessage
Close #i

    Call RP.OpenConnection2("", "CommandCenter", localQBD)
    ticket = RP.BeginSession(HFApp.Options(QuickBooksDataFile), qbFileOpenDoNotCare)
    Response = RP.ProcessRequest(ticket, XmlMessage)
    
'log response
i = FreeFile()
Open PathAppend(GetFolderPath(CSIDL_LOCAL_APPDATA), "Zybertech\" & VB.App.EXEName & "\response.xml") For Output As #i
Print #i, Response
Close #i
    
    
    Select Case Err.Number
        Case 0
            Status = Parse(Parse(Response, 2, "statusCode="), 2, """")
            reason = Parse(Parse(Response, 2, "statusMessage="), 2, """")
            reason = Replace(reason, "&quot;", """")
            Select Case Status
                Case "0" 'no error
                    SubmitQBXml = Response
                    
                Case "1" 'A query request did not find a matching object in QuickBooks
                    SubmitQBXml = Response
                                        
'                Case "3100" ' dupkey
'                    On Error GoTo 0
'                    Err.Raise vbObjectError + 6, "QuickBooks API", "Cannot insert duplicate key"
                    
                
                Case Else
                    MsgBox "Unable to post." & vbCrLf & Status & " - " & reason & vbCrLf & vbCrLf & XmlMessage, vbExclamation, App.ProductName
            End Select
        
        Case Else
            MsgBox "Unable to post." & vbCrLf & vbCrLf & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & XmlMessage, vbExclamation, App.ProductName
                
    End Select
    
On Error Resume Next
    Call RP.EndSession(ticket)
    Call RP.CloseConnection
End Function

Private Function StartQBXml() As String
    Dim s As String
    s = ""
    s = s & "<?xml version=""1.0"" ?>" & vbCrLf
    s = s & "<?qbxml version=""6.0""?>" & vbCrLf
    s = s & "<QBXML>" & vbCrLf
    s = s & "<QBXMLMsgsRq onError=""stopOnError"">" & vbCrLf
    StartQBXml = s
End Function

Private Function EndQBXml() As String
    Dim s As String
    s = ""
    s = s & "</QBXMLMsgsRq>" & vbCrLf
    s = s & "</QBXML>" & vbCrLf
    EndQBXml = s
End Function



Private Function AddQBXml(FieldType As XMLFieldTypes, Size As Double, Name As String, value) As String
                     
    'field types are copied over from MasterBuilders implementation
                     
    Dim s As String
    
    
    s = s & "<" & Name & ">"
    If UCase(Right(Trim(Name), 3)) = "REF" Then s = s & "<ObjectID>"
    Select Case FieldType
        'A 'BankAcct
        'B 'RoutingNum
        'C 'Char
        'D 'Date
        'F 'File
        'G 'SocialSecurityNumber
        'M 'Memo
        'N 'Numeric
        'O 'AssemblyPartInclusionFlag
        'P 'PhoneNumber
        'R 'Terms
        's 'State
        'T 'Time
        'Y 'YesNo
        'Z 'zipcode
        Case d:     s = s & format(value, "YYYY-MM-DD")
        Case n:     s = s & format(value, "0" & IIf((Size - Int(Size)) * 10 > 0, "." & String((Size - Int(Size)) * 10, "0"), ""))
        Case m:     s = s & value
        Case st:     s = s & UCase(Left(XmlQuote("" & value), Int(Size)))
        Case Else:  s = s & Left(XmlQuote("" & value), Int(Size))
    End Select
    If UCase(Right(Trim(Name), 3)) = "REF" Then s = s & "</ObjectID>"
    s = s & "</" & Name & ">" & vbCrLf
    
    AddQBXml = s
        
End Function

Private Function XmlQuote(Text As String, Optional Quote As Boolean) As String
    Dim s As String
    s = Text

    s = Replace(s, "&", Chr(1))
    s = Replace(s, ";", Chr(2))
    s = Replace(s, "#", Chr(3))
    
    s = Replace(s, Chr(1), "&#38;")
    s = Replace(s, Chr(2), "&#59;")
    s = Replace(s, Chr(3), "&#35;")


    s = Replace(s, Chr(0), "&#00;")
    s = Replace(s, vbTab, "&#09;")
    s = Replace(s, vbLf, "&#10;")
    s = Replace(s, vbCr, "&#13;")
    s = Replace(s, "!", "&#33;")
    s = Replace(s, vbQuote, "&#34;")
    s = Replace(s, "$", "&#36;")
    s = Replace(s, "%", "&#37;")
    s = Replace(s, "'", "&#39;")
    s = Replace(s, "*", "&#42;")
    s = Replace(s, "+", "&#43;")
    s = Replace(s, ":", "&#58;")
    s = Replace(s, "<", "&#60;")
    s = Replace(s, ">", "&#62;")
    s = Replace(s, "?", "&#63;")
    s = Replace(s, "^", "&#94;")
    s = Replace(s, "`", "&#96;")
    s = Replace(s, "{", "&#123;")
    s = Replace(s, "}", "&#125;")
    s = Replace(s, "|", "&#124;")
    s = Replace(s, "~", "&#126;")
    
    s = Replace(s, "`", "'")
    s = Replace(s, "„", ",,")
    s = Replace(s, "…", "...")
    s = Replace(s, "ˆ", "^")
    s = Replace(s, "‰", "")
    s = Replace(s, "‹", "<")
    s = Replace(s, "‘", "'")
    s = Replace(s, "’", "'")
    s = Replace(s, "“", """")
    s = Replace(s, "”", """")
    s = Replace(s, "•", "*")
    s = Replace(s, "–", "-")
    s = Replace(s, "—", "-")
    s = Replace(s, "™", "(TM)")
    s = Replace(s, "©", "(C)")
    s = Replace(s, "«", "<<")
    s = Replace(s, "®", "(R)")
    s = Replace(s, "²", "2")
    s = Replace(s, "³", "3")
    s = Replace(s, "·", "")
    s = Replace(s, "¸", "")
    s = Replace(s, "¹", "1")
    s = Replace(s, "º", "0")
    s = Replace(s, "»", ">>")
    s = Replace(s, "¼", "1/4")
    s = Replace(s, "½", "1/2")
    s = Replace(s, "¾", "3/4")
    
    s = Replace(s, Chr(183), "&#183;")
    
    s = Replace(s, Chr(145), Chr(39)) 'left curly single quote
    s = Replace(s, Chr(146), Chr(39)) 'right curly single quote
    s = Replace(s, Chr(147), Chr(34)) 'left curly double quote
    s = Replace(s, Chr(148), Chr(34)) 'right curly double quote
    s = Replace(s, Chr(96), Chr(39))  'left accent
    s = Replace(s, Chr(180), Chr(39)) 'right accent
    s = Replace(s, Chr(150), "-")
    
    If Quote Then s = vbQuote & s & vbQuote
    XmlQuote = s
End Function







Public Sub QB_ImportCustomers()
On Error GoTo eh
    Dim c As Long
    Dim i As Long
    Dim s As String
    Dim rs As Recordset
    
    
    s = ""
    s = s & StartQBXml()
    s = s & "<CustomerQueryRq metaData=""MetaDataAndResponseData""/>" & vbCrLf
    s = s & EndQBXml()
    s = SubmitQBXml(s)
    If s <> "" Then
        On Error Resume Next
        c = Val(Parse(Parse(s, 2, "retCount="""), 1, "="))
        i = 0
        On Error GoTo eh
        s = "exec QB_Customers " & DbQuote(Str, s)
    End If
    Set rs = HFApp.SqlExec(s, dbHomefront)
    
    If Not rs.EOF Then Call HFApp.SqlExec("delete from customers", dbHomefront)
    
    While Not rs.EOF
    
        i = i + 1
        'FMain.lblStatus.Caption = "reading customer " & i & " of " & c
        'FMain.lblStatus.Refresh
        
        s = ""
        s = s & "INSERT INTO Customers(Cust,cname,cstatus,cCont1,cPhone,cFax,cEmail,ccont2,cbadd1, cbAdd2, cbCiy, cbstate, cbzip,cadd1,cadd2,ccity,cstate,czip)" & vbCrLf
        s = s & "VALUES(" & DbQuote(Str, "" & rs("ARCustomer"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("Description"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Str, IIf("" & rs("Active") = "true", "Active", "Inactive"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("Contact1"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("Phone1"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("Fax1"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("Email1"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("Contact2"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("BillAddr1"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("BillAddr2"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("BillCity"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("BillProvince"), , True, 2) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("BillPostalCode"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("ShipAddr1"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("ShipAddr2"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("ShipCity"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("ShipProvince"), , True, 2) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("ShipPostalCode"), , True) & ")"
        Call HFApp.SqlExec(s, dbHomefront)
        
        s = ""
        s = s & "UPDATE Customers" & vbCrLf
        s = s & "SET cname=" & DbQuote(Str, "" & rs("Description"), , True) & vbCrLf
        s = s & "   ,cstatus=" & DbQuote(Str, IIf("" & rs("Active") = "true", "Active", "Inactive"), , True) & vbCrLf
        s = s & "   ,ccont1=" & DbQuote(Str, "" & rs("Contact1"), , True) & vbCrLf
        s = s & "   ,cphone=" & DbQuote(Str, "" & rs("Phone1"), , True) & vbCrLf
        s = s & "   ,cfax=" & DbQuote(Str, "" & rs("Fax1"), , True) & vbCrLf
        s = s & "   ,cemail=" & DbQuote(Str, "" & rs("Email1"), , True) & vbCrLf
        s = s & "   ,ccont2=" & DbQuote(Str, "" & rs("Contact2"), , True) & vbCrLf
        s = s & "   ,cbadd1=" & DbQuote(Str, "" & rs("BillAddr1"), , True) & vbCrLf
        s = s & "   ,cbadd2=" & DbQuote(Str, "" & rs("BillAddr2"), , True) & vbCrLf
        s = s & "   ,cbciy=" & DbQuote(Str, "" & rs("BillCity"), , True) & vbCrLf
        s = s & "   ,cbstate=" & DbQuote(Str, "" & rs("BillProvince"), , True, 2) & vbCrLf
        s = s & "   ,cbzip=" & DbQuote(Str, "" & rs("BillPostalCode"), , True) & vbCrLf
        s = s & "   ,cadd1=" & DbQuote(Str, "" & rs("ShipAddr1"), , True) & vbCrLf
        s = s & "   ,cadd2=" & DbQuote(Str, "" & rs("ShipAddr2"), , True) & vbCrLf
        s = s & "   ,ccity=" & DbQuote(Str, "" & rs("ShipCity"), , True) & vbCrLf
        s = s & "   ,cstate=" & DbQuote(Str, "" & rs("ShipProvince"), , True, 2) & vbCrLf
        s = s & "   ,czip=" & DbQuote(Str, "" & rs("ShipPostalCode"), , True) & vbCrLf
        s = s & "WHERE cust=" & DbQuote(Str, "" & rs("ARCustomer"), , True) & vbCrLf
        Call HFApp.SqlExec(s, dbHomefront)
        
        rs.MoveNext
    Wend

    'FMain.lblStatus.Caption = ""
Exit Sub
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Call errHandler(SRCFILE & "QB_ImportCustomers")
        'FMain.lblStatus.Caption = ""
    End If
End Sub

Public Sub QB_ImportJobs()
On Error GoTo eh
    Dim i As Long
    Dim c As Long
    Dim s As String
    Dim rs As Recordset
    
    s = ""
    s = s & StartQBXml()
    s = s & "<CustomerQueryRq metaData=""MetaDataAndResponseData""/>" & vbCrLf
    s = s & EndQBXml()
    s = SubmitQBXml(s)
    If s <> "" Then
        On Error Resume Next
        c = Val(Parse(Parse(s, 2, "retCount="""), 1, "="))
        i = 0
        On Error GoTo eh
        s = "exec QB_Jobs " & DbQuote(Str, s)
    End If
    Set rs = HFApp.SqlExec(s, dbHomefront)
    
    
    While Not rs.EOF
        
        i = i + 1
        'FMain.lblStatus.Caption = "reading job " & i & " of " & c
        'FMain.lblStatus.Refresh
        
        s = ""
        s = s & "INSERT INTO Jobs(externalid,job,jdesc,jarcust,jaddr1,jaddr2,jcity,jstate,jzip,jbaddr1,jbaddr2,jbcity,jbstate,jbzip,jstatus,r1nts,jestsd,jestcd,jactcd,jccont1,jphn,jscope,jbmeth)" & vbCrLf
        s = s & "VALUES(" & DbQuote(Str, "" & rs("externalid"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("job"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("description"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("customer"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("ShipAddr1"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("ShipAddr2"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("ShipCity"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("ShipProvince"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("ShipPostalCode"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("BillAddr1"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("BillAddr2"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("BillCity"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("BillProvince"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("BillPostalCode"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("Status"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("Comments"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Date, "" & rs("EstStartDate"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Date, "" & rs("EstEndDate"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Date, "" & rs("ActEndDate"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("Contact"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("Phone"), , True) & vbCrLf
        s = s & "      ," & DbQuote(Str, "" & rs("Email"), , True) & vbCrLf
        s = s & "      ,'Use Quick Bill')"
        Call HFApp.SqlExec(s, dbHomefront)
        
        s = ""
        s = s & "UPDATE jobs" & vbCrLf
        s = s & "set jdesc=" & DbQuote(Str, "" & rs("Description"), , True) & vbCrLf
        s = s & "   ,jarcust=" & DbQuote(Str, "" & rs("Customer"), , True) & vbCrLf
        s = s & "   ,jaddr1=" & DbQuote(Str, "" & rs("ShipAddr1"), , True) & vbCrLf
        s = s & "   ,jaddr2=" & DbQuote(Str, "" & rs("ShipAddr2"), , True) & vbCrLf
        s = s & "   ,jcity=" & DbQuote(Str, "" & rs("ShipCity"), , True) & vbCrLf
        s = s & "   ,jstatus=" & DbQuote(Str, "" & rs("Status"), , True) & vbCrLf
        s = s & "   ,jstate=" & DbQuote(Str, "" & rs("ShipProvince"), , True) & vbCrLf
        s = s & "   ,jzip=" & DbQuote(Str, "" & rs("ShipPostalCode"), , True) & vbCrLf
        s = s & "   ,jbaddr1=" & DbQuote(Str, "" & rs("BillAddr1"), , True) & vbCrLf
        s = s & "   ,jbaddr2=" & DbQuote(Str, "" & rs("BillAddr2"), , True) & vbCrLf
        s = s & "   ,jbcity=" & DbQuote(Str, "" & rs("BillCity"), , True) & vbCrLf
        s = s & "   ,jbstate=" & DbQuote(Str, "" & rs("BillProvince"), , True) & vbCrLf
        s = s & "   ,jbzip=" & DbQuote(Str, "" & rs("BillPostalCode"), , True) & vbCrLf
        s = s & "   ,r1nts=" & DbQuote(Str, "" & rs("Comments"), , True) & vbCrLf
        s = s & "   ,jestsd=" & DbQuote(Date, "" & rs("EstStartDate"), , True) & vbCrLf
        s = s & "   ,jestcd=" & DbQuote(Date, "" & rs("EstEndDate"), , True) & vbCrLf
        s = s & "   ,jactcd=" & DbQuote(Date, "" & rs("ActEndDate"), , True) & vbCrLf
        s = s & "   ,jccont1=" & DbQuote(Str, "" & rs("Contact"), , True) & vbCrLf
        s = s & "   ,jphn=" & DbQuote(Str, "" & rs("Phone"), , True) & vbCrLf
        s = s & "   ,jscope=" & DbQuote(Str, "" & rs("Email"), , True) & vbCrLf
        s = s & "   ,jbmeth='Use Quick Bill'"
        s = s & "WHERE externalid=" & DbQuote(Str, "" & rs("externalid"), , True) & vbCrLf
        Call HFApp.SqlExec(s, dbHomefront)
        
        rs.MoveNext
    Wend

    'FMain.lblStatus.Caption = ""
    'FMain.lblStatus.Refresh

Exit Sub
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Call errHandler(SRCFILE & "QB_ImportJobs")
    End If
End Sub


Public Sub QB_UpdateJob(Job As String)
On Error GoTo eh
    Dim i As Long
    Dim s As String
    Dim rs As Recordset
    Dim eid As String

    s = "SELECT j.*,c.cbadd1,c.cbadd2,c.cbciy,c.cbstate,c.cbzip FROM jobs j LEFT OUTER JOIN customers c ON(j.jarcust=c.cust) WHERE job=" & DbQuote(Str, Job)
    Set rs = HFApp.SqlExec(s, dbHomefront)
    
    While Not rs.EOF
        eid = "" & rs("ExternalID")

        s = StartQBXml()
        If eid = "" Then
            s = s & "<CustomerAddRq requestID=""1"">" & vbCrLf
            s = s & "<CustomerAdd>" & vbCrLf
            s = s & AddQBXml(c, 41, "Name", "" & rs("job"))
        Else
            s = s & "<CustomerModRq requestID=""1"">" & vbCrLf
            s = s & "<CustomerMod>" & vbCrLf
            s = s & AddQBXml(c, 41, "ListID", eid)
            s = s & AddQBXml(c, 16, "EditSequence", GetJobEditSeq(eid))
        End If
        s = s & "<ParentRef>" & vbCrLf
        s = s & AddQBXml(c, 41, "ListID", "" & rs("jarcust"))
        s = s & "</ParentRef>" & vbCrLf
        s = s & "<BillAddress>" & vbCrLf
        s = s & AddQBXml(c, 41, "Addr1", "" & rs("cbAdd1"))
        s = s & AddQBXml(c, 41, "Addr2", "" & rs("cbAdd2"))
        s = s & AddQBXml(c, 31, "City", "" & rs("cbCiy"))
        s = s & AddQBXml(c, 21, "State", "" & rs("cbState"))
        s = s & AddQBXml(c, 13, "PostalCode", "" & rs("cbzip"))
        s = s & "</BillAddress>" & vbCrLf
        
        s = s & "<ShipAddress>" & vbCrLf
        s = s & AddQBXml(c, 41, "Addr1", "" & rs("jAddr1"))
        s = s & AddQBXml(c, 41, "Addr2", "" & rs("jAddr2"))
        s = s & AddQBXml(c, 31, "City", "" & rs("jCity"))
        s = s & AddQBXml(c, 21, "State", "" & rs("jState"))
        s = s & AddQBXml(c, 13, "PostalCode", "" & rs("jzip"))
        s = s & "</ShipAddress>" & vbCrLf
        
        s = s & AddQBXml(c, 21, "Phone", "" & rs("jphn"))
        s = s & AddQBXml(c, 1023, "Email", "" & rs("jscope"))
        s = s & AddQBXml(c, 209, "Contact", "" & rs("jccont1"))
        s = s & AddQBXml(c, 10, "JobStatus", "" & rs("jstatus"))
        
        If "" & rs("jestsd") <> "" Then s = s & AddQBXml(d, 10, "JobStartDate", "" & rs("jestsd"))
        If "" & rs("jestcd") <> "" Then s = s & AddQBXml(d, 10, "JobProjectedEndDate", "" & rs("jestcd"))
        
        s = s & AddQBXml(c, 99, "JobDesc", "" & rs("jdesc"))
        s = s & AddQBXml(c, 4095, "Notes", "" & rs("r1nts"))
        If eid = "" Then
            s = s & "</CustomerAdd>" & vbCrLf
            s = s & "</CustomerAddRq>" & vbCrLf
        Else
            s = s & "</CustomerMod>" & vbCrLf
            s = s & "</CustomerModRq>" & vbCrLf
        End If
        s = s & EndQBXml()
        s = SubmitQBXml(s)
        
        If eid = "" Then
            eid = Parse(s, 2, "ListID>")
            eid = Left(eid, Len(eid) - 2)
        
            s = ""
            s = s & "update jobs" & vbCrLf
            s = s & "   set ExternalID=" & DbQuote(Str, eid) & vbCrLf
            s = s & " where job=" & DbQuote(Str, "" & rs("Job")) & vbCrLf
            Call HFApp.SqlExec(s, dbHomefront)
        End If
        
        rs.MoveNext
    Wend


Exit Sub
eh: If InStr(1, Err.Description, "duplicate") > 0 Then
        Resume Next
    Else
        Call errHandler("QB_UpdateJob", s)
    End If
End Sub


Private Function GetJobEditSeq(ExternalID As String) As String
    Dim s As String
    Dim eid As String
    
    s = StartQBXml()
    s = s & "<CustomerQueryRq>" & vbCrLf
    s = s & "<ListID>" & ExternalID & "</ListID>" & vbCrLf
    s = s & "<IncludeRetElement> EditSequence </IncludeRetElement>" & vbCrLf
    s = s & "</CustomerQueryRq>" & vbCrLf
    s = s & EndQBXml()
    s = SubmitQBXml(s)
    
    s = Parse(s, 2, "EditSequence>")
    s = Left(s, Len(s) - 2)
    GetJobEditSeq = s

End Function




Public Sub PostEstimatesToQB(JobCSV As String)
On Error GoTo eh
    
    Dim s As String
    Dim i As Integer
    Dim rs As Recordset
    Dim lastJob As String
    Dim ThisJob As String
    Dim IsUSVersion As Boolean
    
    IsUSVersion = HFApp.Options.ValueByName("AccountingVersion") <> "CA"
    
    s = ""
    s = s & "SELECT j.ExternalID Job" & vbCrLf
    s = s & "      ,i.CostCode" & vbCrLf
    s = s & "      ,i.Category" & vbCrLf
    s = s & "      ,i.Description" & vbCrLf
    s = s & "      ,i.Qty" & vbCrLf
    s = s & "      ,i.UOM" & vbCrLf
    s = s & "      ,(i.Qty * i.Cost) Amount" & vbCrLf
    s = s & "      ,t.ExternalID TaxGroup" & vbCrLf
    s = s & "  FROM Workorders w" & vbCrLf
    s = s & "       JOIN Jobs j ON(w.job=j.job)" & vbCrLf
    s = s & "       JOIN WorkorderItems i ON(w.WorkorderID=i.WorkorderID)" & vbCrLf
    s = s & "       JOIN TaxGroups t ON(j.DivisonID = t.DivisionID and i.taxgroup=t.[group])" & vbCrLf
    s = s & " WHERE isnull(i.Cost,0)<>0" & vbCrLf
    s = s & "   AND isnull(i.Posted,0)=0" & vbCrLf
    s = s & "   AND w.Job IN(" & JobCSV & ")"
    Set rs = HFApp.SqlExec(s, dbHomefront)
    
    s = StartQBXml()
    While Not rs.EOF
        ThisJob = "" & rs("Job")
        If ThisJob <> lastJob Then
            If lastJob <> "" Then
                s = s & "</EstimateAdd></EstimateAddRq>" & vbCrLf
            End If
            i = i + 1
            lastJob = ThisJob
            s = s & "<EstimateAddRq requestID=""1"">" & vbCrLf
            s = s & "<EstimateAdd>" & vbCrLf
            s = s & "<CustomerRef>" & vbCrLf
            s = s & AddQBXml(c, 41, "ListID", "" & rs("Job"))
            s = s & "</CustomerRef>" & vbCrLf
            s = s & AddQBXml(d, 41, "TxnDate", Now())
        End If
        s = s & "<EstimateLineAdd>" & vbCrLf
        s = s & "<ItemRef>" & AddQBXml(c, 41, "ListID", "" & rs("CostCode")) & "</ItemRef>" & vbCrLf
        s = s & AddQBXml(c, 4095, "Desc", "" & rs("Description"))
        s = s & AddQBXml(n, 8.2, "Rate", "" & rs("Amount"))
        s = s & "<ClassRef>" & AddQBXml(c, 41, "ListID", "" & rs("Category")) & "</ClassRef>" & vbCrLf
        If IsUSVersion = False And "" & rs("TaxGroup") <> "" Then
            s = s & "<SalesTaxCodeRef>" & AddQBXml(c, 41, "ListID", "" & rs("TaxGroup")) & "</SalesTaxCodeRef>" & vbCrLf
        End If
        s = s & "</EstimateLineAdd>" & vbCrLf
        
        rs.MoveNext
    Wend
    s = s & "</EstimateAdd></EstimateAddRq>" & vbCrLf
    s = s & EndQBXml()
    s = SubmitQBXml(s)

    'mark as posted
    s = ""
    s = s & "UPDATE WorkorderItems"
    s = s & "   SET Posted=1" & vbCrLf
    s = s & "      ,PostedCost=Cost" & vbCrLf
    s = s & "      ,PostedQty=Qty" & vbCrLf
    s = s & "  FROM Workorders w" & vbCrLf
    s = s & "      JOIN WorkorderItems i ON(w.WorkorderID=i.WorkorderID)" & vbCrLf
    s = s & " WHERE isnull(i.Cost,0)<>0" & vbCrLf
    s = s & "   AND isnull(i.Posted,0)=0" & vbCrLf
    s = s & "   AND w.Job IN(" & JobCSV & ")"
    Set rs = HFApp.SqlExec(s, dbHomefront)


Exit Sub
eh: Call errHandler(SRCFILE & "PostEstimatesToQB")
End Sub



Public Sub QB_ImportAccounts()
On Error GoTo eh
    Dim i As Long
    Dim c As Long
    Dim s As String
    Dim rs As Recordset
    
    s = ""
    s = s & StartQBXml()
    s = s & "<AccountQueryRq metaData=""MetaDataAndResponseData"">" & vbCrLf
    s = s & "<IncludeRetElement>AccountNumber</IncludeRetElement>" & vbCrLf
    s = s & "<IncludeRetElement>Name</IncludeRetElement>" & vbCrLf
    s = s & "</AccountQueryRq>" & vbCrLf
    s = s & EndQBXml()
    s = SubmitQBXml(s)
    If s <> "" Then
        On Error Resume Next
        c = Val(Parse(Parse(s, 2, "retCount="""), 1, "="))
        i = 0
        On Error GoTo eh
        s = "exec QB_Accounts " & DbQuote(Str, s)
    End If
    
    Set rs = HFApp.SqlExec(s, dbHomefront)
    
    If Not rs.EOF Then Call HFApp.SqlExec("delete from accounts", dbHomefront)
    While Not rs.EOF
        
        i = i + 1
        'FMain.lblStatus.Caption = "reading account " & i & " of " & c
        'FMain.lblStatus.Refresh
        
        If "" & rs("Account") <> "" Then
            s = ""
            s = s & "INSERT INTO Accounts(Account,Description)" & vbCrLf
            s = s & "VALUES(" & DbQuote(Str, "" & rs("Account"), , True) & vbCrLf
            s = s & "      ," & DbQuote(Str, "" & rs("Description"), , True) & ")"
            Call HFApp.SqlExec(s, dbHomefront)
        End If
        rs.MoveNext
    Wend

    'FMain.lblStatus.Caption = ""
    'FMain.lblStatus.Refresh
    
Exit Sub
eh: Call errHandler(SRCFILE & "QB_ImportAccounts")
End Sub

