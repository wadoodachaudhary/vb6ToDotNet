Attribute VB_Name = "MCrmDataService"
Option Explicit
Private Loaded As Boolean
Private CrmApiKey As String
Private CrmClientID As String
Private CrmEnvironment As String


Public Sub Sales_SetCustomOptionQuote(Location As String, seq As Long)
    Dim s As String
    Dim rs As Recordset
    Dim xml As String
    
    If Not Integrated Then Exit Sub
    
    s = ""
    s = s & "select " & vbCrLf
    s = s & " CRMID CRMID" & vbCrLf
    s = s & ",Bldr_Declined Declined" & vbCrLf
    s = s & ",Bldr_Declined_Reason DeclineReason" & vbCrLf
    s = s & ",Category" & vbCrLf
    s = s & ",Description" & vbCrLf
    s = s & ",Comments" & vbCrLf
    s = s & ",EstimatorNotes" & vbCrLf
    s = s & ",cast(override_price as varchar(50)) ExtendedPrice" & vbCrLf
    s = s & ",cast(tax as varchar(50)) Tax" & vbCrLf
    s = s & ",cast(Cost_Amount as varchar(50)) CostPerUnit" & vbCrLf
    s = s & ",cast(Rate as varchar(50)) PricePerUnit" & vbCrLf
    s = s & ",cast(Qty as varchar(50)) Units" & vbCrLf
    s = s & ",UOM" & vbCrLf
    Select Case Location
        Case "ADDENDUM":       s = s & "from tblscheduleb" & vbCrLf
        Case "CHANGEORDER":    s = s & "from changeorderdetails" & vbCrLf
        Case "DESIGNCENTER":
            'not implemented
            Exit Sub
    End Select
    s = s & "where seq = " & DbQuote(Num, seq) & vbCrLf
    s = s & "for xml path('SetCustomOptionQuote'), root('Request')" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    xml = ""
    While Not rs.EOF
        xml = xml & rs(0)
        rs.MoveNext
    Wend


'   <Request>
'     <SetCustomOptionQuote>
'       <CRMID />
'       <Declined />
'       <DeclineReason />
'       <Category />
'       <Description />
'       <Comments />
'       <EstimatorNotes />
'       <ExtendedPrice />
'       <Tax />
'       <CostPerUnit />
'       <PricePerUnit />
'       <Units />
'       <UOM />
'     <SetCustomOptionQuote>
'   </Request>
    Call SalesWrapper(xml)

End Sub

Public Sub Sales_SetEstimateIndex(InboxBatchID As String)

    Dim s As String
    Dim rs As Recordset
    Dim xml As String
    
    If Not Integrated Then Exit Sub
    
    s = ""
    s = s & "select CRMID,EstimateIndex" & vbCrLf
    s = s & "from estimateassemblies " & vbCrLf
    s = s & "where InboxBatchID = " & DbQuote(Str, InboxBatchID)
    s = s & "for xml path('SetEstimateIndex'),root('Request')" & vbCrLf
    Set rs = HFApp.SqlExec(s)
    xml = ""
    While Not rs.EOF
        xml = xml & rs(0)
        rs.MoveNext
    Wend
        
'   <Request>
'     <SetEstimateIndex>
'       <CRMID>
'       <EstimateIndex>
'     <SetEstimateIndex>
'   </Request>
    If xml <> "" Then
        Call SalesWrapper(xml)
    End If
    
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



