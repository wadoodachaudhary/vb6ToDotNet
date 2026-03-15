Attribute VB_Name = "MCrmDataService"
Option Explicit
Private Loaded As Boolean
Private CrmApiKey As String
Private CrmClientID As String
Private CrmEnvironment As String

Public Sub Sales_DeactivateUser(UserID As String)

    Dim s As String
    Dim rs As Recordset
    Dim xml As String
    
    If Not Integrated Then Exit Sub
    
    s = ""
    s = s & "select " & DbQuote(Str, UserID) & vbCrLf
    s = s & "for xml path('DeactivateUser'),root('Request')" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    xml = ""
    While Not rs.EOF
        xml = xml & rs(0)
        rs.MoveNext
    Wend
        
'   <Request>
'     <DeactivateUser>
'       <UserID>
'     </DeactivateUser>
'   </Request>
    Call SalesWrapper(xml)
    
    
End Sub

Private Function Integrated() As Boolean
    
    If Not Loaded Then
        CrmApiKey = HFApp.Options.ValueByName("CrmApiKey"):
        CrmClientID = HFApp.Options.ValueByName("CrmClientID"):
        CrmEnvironment = HFApp.Options.ValueByName("CrmEnvironment"):
        Loaded = True
        If Not IsIn(CrmEnvironment, "Production", "User Acceptance Testing (UAT)", "") Then
            MsgBox "Invalid system settings. Check your Sales System settings.", vbExclamation, App.ProductName
            Exit Function
        End If
    End If
    
    If CrmApiKey <> "" And CrmClientID <> "" And CrmEnvironment <> "" Then
        Integrated = True
    Else
        Integrated = False
    End If
    
End Function

Private Function SalesWrapper(xml As String) As String
    Dim crm As New HyphenSys.SalesWrapper
    Dim results As String
    SalesWrapper = crm.PostXml(CrmClientID, CrmApiKey, CrmEnvironment = "Production", "1", xml)
End Function


